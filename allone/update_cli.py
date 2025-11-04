"""Command-line helper to perform headless updates on Windows."""

from __future__ import annotations

import argparse
import sys
from pathlib import Path
from typing import Callable

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
    start_update_installation,
)


def _printer(message: str) -> None:
    if message is None:
        return
    print(str(message))


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
    parser.add_argument(
        "--wait",
        action="store_true",
        help="Wait for the updater batch script to finish before exiting.",
    )

    args = parser.parse_args(argv)

    logger = _printer
    install_root = _normalize_install_root(args.install_root)
    target_executable = install_root / args.exe_name

    if not target_executable.exists():
        parser.error(f"Target executable not found: {target_executable}")

    try:
        remote_version = _remote_version_string(args.timeout)
    except Exception as exc:
        print(f"Failed to check for updates: {exc}", file=sys.stderr)
        return 1

    if not args.force and _version_is_not_newer(remote_version, __version__):
        print(f"Already up-to-date (v{__version__}).")
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
