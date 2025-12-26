"""Stable launcher for the AllOne Windows application."""

from __future__ import annotations

import argparse
import json
import subprocess
import sys
from pathlib import Path

from updater import load_current_release, resolve_release_executable


def _default_install_root() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def _write_current_release(install_root: Path, exe_path: Path, version: str) -> None:
    current_path = install_root / "current.json"
    payload = {
        "version": version,
        "path": str(exe_path),
        "timestamp": "",
    }
    with current_path.open("w", encoding="utf-8") as handle:
        json.dump(payload, handle, indent=2)


def _resolve_target(install_root: Path, exe_name: str) -> Path:
    candidate = resolve_release_executable(install_root, exe_name)
    if candidate:
        return candidate
    fallback = install_root / exe_name
    if fallback.exists():
        return fallback
    raise FileNotFoundError(
        f"Unable to locate {exe_name}. Expected releases in {install_root}."
    )


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Launch the latest AllOne release.")
    parser.add_argument(
        "--install-root",
        type=Path,
        default=_default_install_root(),
        help="Base directory containing the releases folder.",
    )
    parser.add_argument(
        "--exe-name",
        default="AllOne Tools.exe",
        help="Executable name to launch from the releases folder.",
    )
    parser.add_argument(
        "--print-target",
        action="store_true",
        help="Print the resolved executable path and exit.",
    )
    parser.add_argument(
        "args",
        nargs=argparse.REMAINDER,
        help="Arguments to forward to the application executable.",
    )
    args = parser.parse_args(argv)

    install_root = args.install_root.expanduser().resolve()
    exe_path = _resolve_target(install_root, args.exe_name)

    if args.print_target:
        print(str(exe_path))
        return 0

    if not exe_path.exists():
        print(f"Launch target does not exist: {exe_path}", file=sys.stderr)
        return 1

    release_state = load_current_release(install_root)
    if not release_state:
        version_hint = exe_path.parent.name
        _write_current_release(install_root, exe_path, version_hint)

    process = subprocess.Popen(
        [str(exe_path)] + args.args,
        cwd=str(exe_path.parent),
    )
    return process.wait()


if __name__ == "__main__":  # pragma: no cover - manual invocation only
    sys.exit(main())
