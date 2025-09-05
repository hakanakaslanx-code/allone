# settings_manager.py
import json
import os

SETTINGS_FILE = "settings.json"

def load_settings() -> dict:
    """Loads settings from the JSON file."""
    if os.path.exists(SETTINGS_FILE):
        try:
            with open(SETTINGS_FILE, "r", encoding='utf-8') as f:
                return json.load(f)
        except json.JSONDecodeError:
            return {}
    return {}

def save_settings(settings: dict):
    """Saves the settings dictionary to the JSON file."""
    with open(SETTINGS_FILE, "w", encoding='utf-8') as f:
        json.dump(settings, f, indent=4)