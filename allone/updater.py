"""Application self-update helpers for AllOne Tools on Windows."""

from __future__ import annotations

import contextlib
import hashlib
import json
import os
import re
import subprocess
import sys
import tempfile
import threading
import time
from dataclasses import dataclass
from pathlib import Path
from typing import Callable, Dict, Optional

import shutil

import requests
from tkinter import messagebox

OWNER = "hakanakaslanx-code"
REPO = "allone"

CHECK_TIMEOUT_SILENT = 3
CHECK_TIMEOUT_INTERACTIVE = 10
DOWNLOAD_TIMEOUT = 60
CHUNK_SIZE = 1024 * 1024

DETACHED_PROCESS = 0x00000008


@dataclass
class UpdateInfo:
    """Metadata describing the newest available update."""

    version: str
    download_url: str
    download_size: Optional[int]
    file_name: str
    sha256: str
    release_notes_url: str


class UpdateError(RuntimeError):
    """Exception raised for recoverable update failures."""


def _is_windows_executable() -> bool:
    return (
        sys.platform.startswith("win")
        and getattr(sys, "frozen", False)
        and sys.executable.lower().endswith(".exe")
    )


def _install_root() -> Path:
    if _is_windows_executable():
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def _paths() -> Dict[str, Path]:
    root = _install_root()
    return {
        "lock": root / "update.lock",
        "updater": root / "run_update.bat",
        "log": root / "last_update.log",
        "success": root / "update_success.json",
    }


def _download_text(url: str, timeout: int) -> str:
    response = requests.get(url, timeout=timeout, headers={"Accept": "text/plain"})
    response.raise_for_status()
    return response.text


def _fetch_latest_release_asset(owner: str, repo: str, timeout: int) -> UpdateInfo:
    api_url = f"https://api.github.com/repos/{owner}/{repo}/releases/latest"
    response = requests.get(
        api_url,
        headers={"Accept": "application/vnd.github+json"},
        timeout=timeout,
    )
    response.raise_for_status()
    release_info = response.json()
    assets = release_info.get("assets", [])
    if not assets:
        raise UpdateError("No assets available for the latest release.")

    asset = None
    for candidate in assets:
        name = candidate.get("name", "").lower()
        if name.endswith(".zip"):
            asset = candidate
            break
    if asset is None:
        for candidate in assets:
            name = candidate.get("name", "").lower()
            if name.endswith(".exe"):
                asset = candidate
                break
    if asset is None:
        raise UpdateError("No supported update package found.")

    asset_name = asset.get("name", "update.bin")
    sha_candidate = None
    expected_hash = None
    base_name = asset_name
    if base_name.lower().endswith(".zip"):
        base_name = base_name[:-4]
    for candidate in assets:
        name = candidate.get("name", "").lower()
        if name.startswith(base_name.lower()) and "sha256" in name:
            sha_candidate = candidate
            break
    if sha_candidate is not None:
        sha_text = _download_text(
            sha_candidate.get("browser_download_url", ""), timeout
        )
        match = re.search(r"([A-Fa-f0-9]{64})", sha_text)
        if match:
            expected_hash = match.group(1).lower()
    if not expected_hash:
        raise UpdateError("Could not locate SHA256 checksum for update package.")

    download_url = asset.get("browser_download_url")
    if not download_url:
        raise UpdateError("Download URL for the latest release is missing.")

    return UpdateInfo(
        version=release_info.get("tag_name") or release_info.get("name") or "",
        download_url=download_url,
        download_size=asset.get("size"),
        file_name=asset_name,
        sha256=expected_hash,
        release_notes_url=release_info.get("html_url", ""),
    )


def _remote_version_string(timeout: int) -> str:
    version_check_url = (
        f"https://raw.githubusercontent.com/{OWNER}/{REPO}/main/allone/app_ui.py"
    )
    content = _download_text(version_check_url, timeout)
    match = re.search(r"__version__\s*=\s*[\"'](.+?)[\"']", content)
    if not match:
        raise UpdateError("Could not determine remote version.")
    return match.group(1)


def _parse_version(value: str) -> tuple:
    parts = []
    for token in re.split(r"[.]+", value):
        if token.isdigit():
            parts.append(int(token))
        else:
            parts.append(token)
    return tuple(parts)


def _download_asset(
    url: str,
    destination: Path,
    expected_size: Optional[int],
    log_callback: Callable[[str], None],
) -> None:
    with requests.get(url, stream=True, timeout=DOWNLOAD_TIMEOUT) as response:
        response.raise_for_status()
        downloaded = 0
        last_logged_percent = -10
        with open(destination, "wb") as dest_file:
            for chunk in response.iter_content(chunk_size=CHUNK_SIZE):
                if not chunk:
                    continue
                dest_file.write(chunk)
                downloaded += len(chunk)
                if expected_size:
                    percent = int(downloaded * 100 / expected_size)
                    if percent - last_logged_percent >= 10:
                        last_logged_percent = percent
                        log_callback(f"Downloading update… {percent}%")


def _calculate_sha256(file_path: Path) -> str:
    digest = hashlib.sha256()
    with open(file_path, "rb") as fp:
        for chunk in iter(lambda: fp.read(1024 * 1024), b""):
            if not chunk:
                break
            digest.update(chunk)
    return digest.hexdigest()


def _prepare_updater_script(script_path: Path) -> None:
    paths = _paths()
    log_path = paths["log"]
    content = """@echo off
setlocal EnableExtensions EnableDelayedExpansion
set "TARGET_EXE=%~1"
set "TARGET_PID=%~2"
set "PACKAGE=%~3"
set "WORKDIR=%~4"
set "TARGET_NAME=%~5"
set "LOG_FILE=%~6"
set "SUCCESS_FILE=%~7"
set "LOCK_FILE=%~8"
set "VERSION_STR=%~9"

>>"%LOG_FILE%" echo [%%date%% %%time%%] Update script started. Waiting for PID !TARGET_PID!.
:waitloop
if "!TARGET_PID!"=="" goto continue
for /f "tokens=1" %%p in ('tasklist /FI "PID eq !TARGET_PID!" ^| find " !TARGET_PID! "') do (
    timeout /T 1 /NOBREAK >NUL
    goto waitloop
)
:continue
>>"%LOG_FILE%" echo [%%date%% %%time%%] Process exited. Preparing workspace.
if exist "%WORKDIR%" rd /S /Q "%WORKDIR%"
mkdir "%WORKDIR%" >NUL 2>&1
set "EXTRACT_DIR=%WORKDIR%\extracted"
mkdir "%EXTRACT_DIR%" >NUL 2>&1
set "NEW_PAYLOAD=%PACKAGE%"
for %%I in ("%PACKAGE%") do set "PKG_EXT=%%~xI"
if /I "!PKG_EXT!"==".zip" (
    powershell -NoProfile -NoLogo -Command "Expand-Archive -LiteralPath '%PACKAGE%' -DestinationPath '%EXTRACT_DIR%' -Force" >>"%LOG_FILE%" 2>&1
    if errorlevel 1 goto fail
    if exist "%EXTRACT_DIR%\%TARGET_NAME%" set "NEW_PAYLOAD=%EXTRACT_DIR%\%TARGET_NAME%"
)
copy /Y "!NEW_PAYLOAD!" "%TARGET_EXE%.new" >>"%LOG_FILE%" 2>&1
if errorlevel 1 goto fail
move /Y "%TARGET_EXE%.new" "%TARGET_EXE%" >>"%LOG_FILE%" 2>&1
if errorlevel 1 goto fail
if exist "%EXTRACT_DIR%" (
    robocopy "%EXTRACT_DIR%" "%~dp1" /E /NFL /NDL /NJH /NJS /NP >>"%LOG_FILE%" 2>&1
    set "ROBOCODE=!errorlevel!"
    if !ROBOCODE! GEQ 8 goto fail
)
start "" "%TARGET_EXE%" >>"%LOG_FILE%" 2>&1
if exist "%SUCCESS_FILE%" del "%SUCCESS_FILE%"
(
    echo {
    echo   "version": "!VERSION_STR!",
    echo   "timestamp": "%%date%% %%time%%"
    echo }
) >"%SUCCESS_FILE%"
set "EXIT_CODE=0"
goto cleanup

:fail
set "EXIT_CODE=1"
:cleanup
if exist "%PACKAGE%" del "%PACKAGE%" >>"%LOG_FILE%" 2>&1
if exist "%WORKDIR%" rd /S /Q "%WORKDIR%" >>"%LOG_FILE%" 2>&1
if exist "%LOCK_FILE%" del "%LOCK_FILE%" >>"%LOG_FILE%" 2>&1
>>"%LOG_FILE%" echo [%%date%% %%time%%] Update script completed with exit code !EXIT_CODE!.
exit /B !EXIT_CODE!
"""
    script_path.write_text(content, encoding="utf-8")


def _launch_updater(script_path: Path, args: list[str]) -> None:
    try:
        subprocess.Popen(
            ["cmd", "/c", str(script_path)] + args,
            creationflags=DETACHED_PROCESS,
            close_fds=True,
        )
    except FileNotFoundError as exc:
        raise UpdateError(f"Failed to launch updater: {exc}")


def _create_temp_dir() -> Path:
    base = Path(tempfile.gettempdir())
    directory = base / f"allone_update_{int(time.time())}"
    directory.mkdir(parents=True, exist_ok=True)
    return directory


def perform_update_installation(
    app_instance,
    update_info: UpdateInfo,
    log_callback: Callable[[str], None],
) -> None:
    if not _is_windows_executable():
        app_instance.after(
            0,
            lambda: messagebox.showinfo(
                "Update",
                "Automatic updates are only supported on Windows builds.",
            ),
        )
        return

    paths = _paths()
    lock_path = paths["lock"]
    log_path = paths["log"]
    updater_path = paths["updater"]

    if lock_path.exists():
        raise UpdateError("Another update appears to be in progress. Please retry later.")

    lock_path.write_text(str(time.time()), encoding="utf-8")
    temp_dir = _create_temp_dir()
    package_path = temp_dir / update_info.file_name

    try:
        log_callback("Downloading update package…")
        _download_asset(update_info.download_url, package_path, update_info.download_size, log_callback)
        log_callback("Download completed. Verifying package integrity…")
        actual_hash = _calculate_sha256(package_path)
        if actual_hash.lower() != update_info.sha256.lower():
            raise UpdateError("Package checksum mismatch. Update aborted.")

        log_callback("Preparing installer…")
        _prepare_updater_script(updater_path)

        args = [
            str(Path(sys.executable).resolve()),
            str(os.getpid()),
            str(package_path),
            str(temp_dir / "work"),
            Path(sys.executable).name,
            str(log_path),
            str(paths["success"]),
            str(lock_path),
            update_info.version,
        ]
        _launch_updater(updater_path, args)
        log_callback("Update installer launched. Application will exit to apply the update.")
    except Exception:
        if lock_path.exists():
            with contextlib.suppress(OSError):
                lock_path.unlink()
        with contextlib.suppress(Exception):
            if package_path.exists():
                package_path.unlink()
            if temp_dir.exists():
                shutil.rmtree(temp_dir)
        raise


def check_for_updates(
    app_instance,
    log_callback: Callable[[str], None],
    current_version: str,
    silent: bool = False,
) -> None:
    timeout = CHECK_TIMEOUT_SILENT if silent else CHECK_TIMEOUT_INTERACTIVE

    if silent and hasattr(app_instance, "is_auto_update_enabled"):
        if not app_instance.is_auto_update_enabled():
            log_callback("Auto-update on startup disabled. Skipping update check.")
            return

    log_callback("Checking for updates…")
    try:
        remote_version = _remote_version_string(timeout)
    except requests.Timeout:
        log_callback("Update check timed out. Possibly offline.")
        return
    except Exception as exc:
        log_callback(f"Update check failed: {exc}")
        if not silent:
            app_instance.after(
                0,
                lambda: messagebox.showerror("Update Check", f"Could not check for updates.\n\n{exc}"),
            )
        return

    if _parse_version(remote_version) <= _parse_version(current_version):
        log_callback(f"✅ Your application is up-to-date. (Version: {current_version})")
        if not silent:
            app_instance.after(
                0,
                lambda: messagebox.showinfo("Update Check", "You are running the latest version."),
            )
        return

    try:
        update_info = _fetch_latest_release_asset(OWNER, REPO, timeout)
    except Exception as exc:
        log_callback(f"Unable to retrieve update metadata: {exc}")
        if not silent:
            app_instance.after(
                0,
                lambda: messagebox.showerror(
                    "Update", f"Update information could not be retrieved.\n\n{exc}"
                ),
            )
        return

    def _prompt() -> None:
        if hasattr(app_instance, "begin_update_prompt"):
            app_instance.begin_update_prompt(update_info)
        else:
            if messagebox.askyesno(
                "Update Available",
                (
                    f"A new version ({update_info.version}) is available.\n\n"
                    "Do you want to download and install it now?"
                ),
            ):
                def _launch() -> None:
                    try:
                        perform_update_installation(app_instance, update_info, log_callback)
                    except Exception as exc2:  # pragma: no cover - defensive UI feedback
                        log_callback(f"Update failed: {exc2}")
                        messagebox.showerror("Update", f"Update failed: {exc2}")

                threading.Thread(target=_launch, daemon=True).start()

    app_instance.after(0, _prompt)


def consume_update_success_message() -> Optional[Dict[str, str]]:
    success_path = _paths()["success"]
    if not success_path.exists():
        return None
    data: Dict[str, str]
    try:
        with open(success_path, "r", encoding="utf-8") as fp:
            loaded = json.load(fp)
        if isinstance(loaded, dict):
            data = {str(key): str(value) for key, value in loaded.items() if value is not None}
        else:
            data = {}
    except Exception:
        data = {}
    try:
        success_path.unlink()
    except OSError:
        pass
    return data


def cleanup_update_artifacts() -> None:
    paths = _paths()
    for key in ("lock", "updater"):
        path = paths[key]
        if path.exists():
            try:
                path.unlink()
            except OSError:
                pass

