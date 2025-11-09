"""Database backup helpers for RugBase."""

from __future__ import annotations

import hashlib
import shutil
from datetime import datetime
from pathlib import Path
from typing import Dict, Tuple

from .paths import get_backups_dir


def _timestamp_for_filename(ts: datetime | None = None) -> str:
    return (ts or datetime.utcnow()).strftime("%Y%m%d-%H%M%S")


def create_database_backup(db_path: Path) -> Tuple[Path, Dict[str, str]]:
    """Copy the SQLite database to the backups directory.

    Returns a tuple containing the backup path and metadata describing the
    backup (timestamp, file size, sha256 hash).
    """
    if not db_path.exists():
        raise FileNotFoundError(f"Database not found at {db_path}")

    backups_dir = get_backups_dir()
    backups_dir.mkdir(parents=True, exist_ok=True)

    timestamp = _timestamp_for_filename()
    backup_path = backups_dir / f"rugbase-{timestamp}.db"
    shutil.copy2(db_path, backup_path)

    size = backup_path.stat().st_size
    digest = _hash_file(backup_path)
    metadata = {
        "timestamp": timestamp,
        "path": str(backup_path),
        "size_bytes": str(size),
        "sha256": digest,
    }
    return backup_path, metadata


def _hash_file(path: Path) -> str:
    sha = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(65536), b""):
            sha.update(chunk)
    return sha.hexdigest()
