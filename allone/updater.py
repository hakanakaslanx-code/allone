"""GitHub update checker for AllOne (Safe Check-only version)."""
from __future__ import annotations

import json
import logging
import requests
from typing import Optional, Dict, Any
from allone.version import __version__

OWNER = "hakanakaslanx-code"
REPO = "allone"
GITHUB_API_URL = f"https://api.github.com/repos/{OWNER}/{REPO}/releases/latest"

def check_for_updates(timeout: int = 5) -> Optional[Dict[str, Any]]:
    """Check GitHub for the latest release version.
    
    Returns a dictionary with 'version', 'url', and 'notes' if a newer version 
    exists, otherwise returns None.
    """
    try:
        response = requests.get(GITHUB_API_URL, timeout=timeout)
        response.raise_for_status()
        data = response.json()
        
        remote_version = data.get("tag_name", "").lstrip("v")
        if not remote_version:
            return None
            
        if _is_newer(remote_version, __version__):
            download_url = None
            for asset in data.get("assets", []):
                if asset.get("name", "").endswith(".exe"):
                    download_url = asset.get("browser_download_url")
                    break

            return {
                "version": remote_version,
                "url": data.get("html_url"),
                "notes": data.get("body", ""),
                "download_url": download_url
            }
    except Exception as e:
        logging.error(f"Failed to check for updates: {e}")
        
    return None

def _is_newer(remote: str, local: str) -> bool:
    """Simple version comparison."""
    try:
        r_parts = [int(p) for p in remote.split(".")]
        l_parts = [int(p) for p in local.split(".")]
        return r_parts > l_parts
    except (ValueError, AttributeError):
        return remote > local
