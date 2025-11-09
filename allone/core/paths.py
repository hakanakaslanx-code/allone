"""Path helpers for RugBase application data and credentials management."""

from __future__ import annotations

import os
import platform
from pathlib import Path
from typing import Optional

APP_NAME = "RugBase"


def _expand(path: Path) -> Path:
    return Path(os.path.expanduser(str(path)))


def _default_base_dir() -> Path:
    """Return the default base directory for application data."""
    override = os.environ.get("RUGBASE_APPDATA")
    if override:
        return _expand(Path(override))

    system = platform.system()
    home = Path.home()
    if system == "Windows":
        base = os.environ.get("LOCALAPPDATA")
        if base:
            return _expand(Path(base)) / APP_NAME
        return home / "AppData" / "Local" / APP_NAME
    if system == "Darwin":
        return home / "Library" / "Application Support" / APP_NAME
    xdg = os.environ.get("XDG_DATA_HOME")
    if xdg:
        return _expand(Path(xdg)) / APP_NAME
    return home / ".local" / "share" / APP_NAME


def get_appdata_dir(create: bool = True) -> Path:
    """Return the application data directory, creating it if requested."""
    base = _default_base_dir()
    if create:
        base.mkdir(parents=True, exist_ok=True)
    return base


def _subpath(name: str, *, create: bool = True) -> Path:
    base = get_appdata_dir(create=create)
    path = base / name
    if create:
        path.parent.mkdir(parents=True, exist_ok=True)
    return path


def get_settings_path() -> Path:
    return _subpath("settings.json")


def get_credentials_dir() -> Path:
    path = get_appdata_dir()
    creds = path / "credentials"
    creds.mkdir(parents=True, exist_ok=True)
    return creds


def get_service_account_path() -> Path:
    return get_credentials_dir() / "service_account.json"


def get_token_path() -> Path:
    return _subpath("token.json")


def get_backups_dir() -> Path:
    path = get_appdata_dir()
    backups = path / "backups"
    backups.mkdir(parents=True, exist_ok=True)
    return backups


def get_lock_path() -> Path:
    return get_appdata_dir() / ".sync.lock"


def resolve_path(path: Optional[str]) -> Path:
    """Resolve a user-provided path relative to the current working directory."""
    if not path:
        return Path.cwd()
    expanded = os.path.expandvars(os.path.expanduser(path))
    return Path(expanded).resolve()
