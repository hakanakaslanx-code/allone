"""Command-line helper to perform headless updates on Windows."""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path
from typing import Callable, Any

from version import __version__
from updater import (
    CHECK_TIMEOUT_INTERACTIVE,
    OWNER,
    REPO,
    UpdateError,
    UpdateInfo,
    _fetch_latest_release_asset,
    _remote_version_string,
    _version_is_not_newer,
    load_current_release,
    resolve_release_executable,
    start_update_installation,
)


def _printer(message: str) -> None:
    if message is None:
        return
    print(str(message))


def _load_settings(settings_path: Path) -> dict[str, Any]:
    if not settings_path.exists():
        return {}

    try:
        with settings_path.open("r", encoding="utf-8") as stream:
            payload = json.load(stream)
    except (OSError, json.JSONDecodeError):
        return {}

    if isinstance(payload, dict):
        return payload
    return {}


def _configure_auto_update_setting(
    install_root: Path, *, enable: bool, logger: Callable[[str], None]
) -> None:
    settings_path = install_root / "settings.json"
    settings = _load_settings(settings_path)

    updates_section = settings.get("updates")
    if not isinstance(updates_section, dict):
        updates_section = {}
        settings["updates"] = updates_section

    desired_value = bool(enable)
    current_value = bool(updates_section.get("auto_update_on_startup", False))

    if current_value == desired_value:
        state = "enabled" if desired_value else "disabled"
        logger(f"Auto-update on startup already {state}.")
        return

    updates_section["auto_update_on_startup"] = desired_value

    with settings_path.open("w", encoding="utf-8") as stream:
        json.dump(settings, stream, indent=4)

    state = "enabled" if desired_value else "disabled"
    logger(f"Auto-update on startup {state}.")


def _normalize_install_root(value: Path) -> Path:
    try:
        return value.expanduser().resolve()
    except Exception:
        return Path(value)


def _describe_update(update_info: UpdateInfo, logger: Callable[[str], None]) -> None:
    logger(f"Update v{update_info.version} is available.")
    if update_info.release_notes_url:
        logger(f"Release notes: {update_info.release_notes_url}")


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Check for updates and launch the AllOne updater batch script.",
    )
    parser.add_argument(
        "--install-root",
        type=Path,
        default=Path.cwd(),
        help="Base directory that contains the application executable.",
    )
    parser.add_argument(
        "--exe-name",
        default="AllOne Tools.exe",
        help="Filename of the application executable to update.",
    )
    parser.add_argument(
        "--timeout",
        type=int,
        default=CHECK_TIMEOUT_INTERACTIVE,
        help="Timeout in seconds for GitHub API requests.",
    )
    parser.add_argument(
        "--check-only",
        action="store_true",
        help="Only check for a new release without downloading it.",
    )
    parser.add_argument(
        "--force",
        action="store_true",
        help="Install even if the local version appears current.",
    )
    toggle_group = parser.add_mutually_exclusive_group()
    toggle_group.add_argument(
        "--enable-auto-update",
        action="store_true",
        help="Enable automatic updates on application startup and exit.",
    )
    toggle_group.add_argument(
        "--disable-auto-update",
        action="store_true",
        help="Disable automatic updates on application startup and exit.",
    )
    parser.add_argument(
        "--wait",
        action="store_true",
        help="Wait for the updater batch script to finish before exiting.",
    )

    args = parser.parse_args(argv)

    logger = _printer
    install_root = _normalize_install_root(args.install_root)
    target_executable = install_root / args.exe_name
    release_state = load_current_release(install_root)
    current_version = __version__
    if release_state:
        release_version = release_state.get("version")
        if release_version:
            current_version = release_version
    release_executable = resolve_release_executable(install_root, args.exe_name)
    if release_executable is not None:
        target_executable = release_executable
        if not release_state:
            current_version = release_executable.parent.name

    if args.enable_auto_update or args.disable_auto_update:
        try:
            _configure_auto_update_setting(
                install_root,
                enable=args.enable_auto_update,
                logger=logger,
            )
        except UpdateError as exc:
            print(f"Failed to update auto-update setting: {exc}", file=sys.stderr)
            return 1
        except OSError as exc:
            print(f"Unable to write settings file: {exc}", file=sys.stderr)
            return 1
        return 0

    if not target_executable.exists():
        parser.error(f"Target executable not found: {target_executable}")

    try:
        remote_version = _remote_version_string(args.timeout)
    except Exception as exc:
        print(f"Failed to check for updates: {exc}", file=sys.stderr)
        return 1

    if not args.force and _version_is_not_newer(remote_version, current_version):
        print(f"Already up-to-date (v{current_version}).")
        return 0

    if args.check_only:
        print(f"Update available: {remote_version}")
        return 0

    try:
        update_info = _fetch_latest_release_asset(OWNER, REPO, args.timeout)
    except Exception as exc:
        print(f"Unable to retrieve update metadata: {exc}", file=sys.stderr)
        return 1

    _describe_update(update_info, logger)

    try:
        process = start_update_installation(
            update_info,
            logger,
            target_executable,
            target_pid=None,
            base_dir=install_root,
            wait_for_completion=args.wait,
        )
    except UpdateError as exc:
        print(f"Update failed: {exc}", file=sys.stderr)
        return 1
    except Exception as exc:  # pragma: no cover - defensive catch-all
        print(f"Unexpected error while launching updater: {exc}", file=sys.stderr)
        return 1

    if args.wait and process is not None:
        return_code = process.wait()
        if return_code != 0:
            print(f"Updater exited with code {return_code}", file=sys.stderr)
            return return_code
        logger("Update completed successfully.")
        return 0

    logger("Updater launched in the background. The application will restart when ready.")
    return 0


if __name__ == "__main__":  # pragma: no cover - manual invocation only
    sys.exit(main())
