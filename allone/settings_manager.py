"""Simple JSON-backed settings loader and saver."""

from __future__ import annotations

import json
from typing import Any, Dict

from core.paths import get_settings_path


def load_settings() -> Dict[str, Any]:
    """Load persisted settings from the application data directory."""
    path = get_settings_path()
    if not path.exists():
        return {}
    try:
        with path.open("r", encoding="utf-8") as handle:
            return json.load(handle)
    except json.JSONDecodeError:
        return {}


def save_settings(settings: Dict[str, Any]) -> None:
    """Persist settings to the application data directory."""
    path = get_settings_path()
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8") as handle:
        json.dump(settings, handle, indent=4)
