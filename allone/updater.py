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
from typing import Callable, Dict, List, Optional, Sequence, Tuple

import shutil

import requests
from tkinter import messagebox

OWNER = "hakanakaslanx-code"
REPO = "allone"

USER_AGENT = "AllOneUpdateClient/1.0"

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
    sha256: Optional[str]
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


def _paths(root: Optional[Path] = None) -> Dict[str, Path]:
    base = Path(root) if root is not None else _install_root()
    base = base.resolve()
    return {
        "lock": base / "update.lock",
        "updater": base / "run_update.bat",
        "log": base / "last_update.log",
        "success": base / "update_success.json",
    }


def _get_latest_release_payload(owner: str, repo: str, timeout: int) -> dict:
    api_url = f"https://api.github.com/repos/{owner}/{repo}/releases/latest"
    headers = {
        "Accept": "application/vnd.github+json",
        "User-Agent": USER_AGENT,
    }
    try:
        response = requests.get(
            api_url,
            headers=headers,
            timeout=timeout,
        )
        response.raise_for_status()
    except requests.Timeout:
        raise
    except requests.RequestException as exc:
        raise UpdateError(f"GitHub release metadata request failed: {exc}") from exc
    try:
        release_info = response.json()
    except ValueError as exc:
        raise UpdateError("Received invalid JSON from GitHub releases API.") from exc
    if not isinstance(release_info, dict):
        raise UpdateError(
            "Received unexpected JSON structure from GitHub releases API."
        )
    return release_info


def _download_text(url: str, timeout: int) -> str:
    headers = {
        "Accept": "text/plain",
        "User-Agent": USER_AGENT,
    }
    response = requests.get(url, timeout=timeout, headers=headers)
    response.raise_for_status()
    return response.text


def _fetch_latest_release_asset(owner: str, repo: str, timeout: int) -> UpdateInfo:
    release_info = _get_latest_release_payload(owner, repo, timeout)
    assets = release_info.get("assets", [])
    if not isinstance(assets, list) or not assets:
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
    expected_hash: Optional[str] = None
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
    download_url = asset.get("browser_download_url")
    if not download_url:
        raise UpdateError("Download URL for the latest release is missing.")

    version_str = release_info.get("tag_name") or release_info.get("name") or ""
    release_notes_url = release_info.get("html_url") or (
        f"https://github.com/{owner}/{repo}/releases/tag/{version_str}" if version_str else ""
    )

    return UpdateInfo(
        version=version_str,
        download_url=download_url,
        download_size=asset.get("size"),
        file_name=asset_name,
        sha256=expected_hash,
        release_notes_url=release_notes_url,
    )


def _remote_version_string(timeout: int) -> str:
    release_info = _get_latest_release_payload(OWNER, REPO, timeout)
    version = release_info.get("tag_name") or release_info.get("name")
    if not version or not isinstance(version, str):
        raise UpdateError("Version information missing in GitHub release metadata.")
    version = version.strip()
    if not version:
        raise UpdateError("Version information was empty in GitHub release metadata.")
    return version


def _tokenize_version(value: str) -> List[Tuple[str, str]]:
    """Return a token list describing the numeric and textual parts of ``value``."""

    value = str(value).strip()
    if not value:
        return []

    value = re.sub(r"^[^0-9]+", "", value)
    if not value:
        return []

    tokens: List[Tuple[str, str]] = []
    for token in re.split(r"[.\-_]+", value):
        if not token:
            continue
        for chunk in re.findall(r"[0-9]+|[A-Za-z]+", token):
            kind = "num" if chunk.isdigit() else "text"
            tokens.append((kind, chunk.lower()))
    return tokens


def _parse_version(value: str) -> Tuple[Tuple[int, object], ...]:
    """Return a comparable tuple for version strings."""

    parts: List[Tuple[int, object]] = []
    for kind, chunk in _tokenize_version(value):
        if kind == "num":
            parts.append((1, int(chunk)))
        else:
            parts.append((0, chunk))
    return tuple(parts)


def _compare_version_tokens(
    remote_tokens: Sequence[Tuple[str, str]],
    current_tokens: Sequence[Tuple[str, str]],
) -> int:
    """Compare two token sequences without raising ``TypeError``.

    Returns a negative value when ``remote_tokens`` describe an older version,
    zero when they are equivalent and a positive value when ``remote_tokens``
    represent a newer version.
    """

    for (remote_kind, remote_value), (current_kind, current_value) in zip(
        remote_tokens, current_tokens
    ):
        if remote_kind == current_kind:
            if remote_kind == "num":
                remote_num = int(remote_value)
                current_num = int(current_value)
                if remote_num != current_num:
                    return remote_num - current_num
            else:
                if remote_value != current_value:
                    return -1 if remote_value < current_value else 1
        else:
            return 1 if remote_kind == "num" else -1

    if len(remote_tokens) == len(current_tokens):
        return 0

    if len(remote_tokens) > len(current_tokens):
        for kind, value in remote_tokens[len(current_tokens) :]:
            if kind == "num":
                if int(value) != 0:
                    return 1
            else:
                return -1
        return 0

    for kind, value in current_tokens[len(remote_tokens) :]:
        if kind == "num":
            if int(value) != 0:
                return -1
        else:
            return 1
    return 0


def _version_is_not_newer(
    remote_version: str,
    current_version: str,
    log_callback: Optional[Callable[[str], None]] = None,
) -> bool:
    """Return ``True`` when ``remote_version`` is not newer than ``current_version``.

    The helper attempts a fast tuple comparison via :func:`_parse_version` and
    falls back to a defensive token based comparison when legacy data triggers a
    ``TypeError`` (for example older settings files stored tuples such as
    ``("v5", 1, 3)``).
    """

    remote_key = _parse_version(remote_version)
    current_key = _parse_version(current_version)
    try:
        return remote_key <= current_key
    except TypeError as exc:  # pragma: no cover - exercised via unit tests
        if log_callback:
            log_callback(
                "⚠️ Version comparison required a compatibility fallback. "
                f"Details: {exc}."
            )
        remote_tokens = _tokenize_version(remote_version)
        current_tokens = _tokenize_version(current_version)
        return _compare_version_tokens(remote_tokens, current_tokens) <= 0


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
if exist "%TARGET_EXE%" (
    attrib -R "%TARGET_EXE%" >>"%LOG_FILE%" 2>&1
    del /F /Q "%TARGET_EXE%" >>"%LOG_FILE%" 2>&1
    if exist "%TARGET_EXE%" (
        >>"%LOG_FILE%" echo [%%date%% %%time%%] Failed to remove existing executable.
        goto fail
    )
)
attrib -R "%TARGET_EXE%.new" >>"%LOG_FILE%" 2>&1
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


def _launch_updater(
    script_path: Path,
    args: list[str],
    *,
    detach: bool = True,
) -> Optional[subprocess.Popen]:
    creationflags = DETACHED_PROCESS if detach else 0
    try:
        process = subprocess.Popen(
            ["cmd", "/c", str(script_path)] + args,
            creationflags=creationflags,
            close_fds=True,
        )
    except FileNotFoundError as exc:
        raise UpdateError(f"Failed to launch updater: {exc}")
    return None if detach else process


@dataclass
class PreparedUpdate:
    """Temporary workspace describing a staged update installation."""

    base_dir: Path
    lock_path: Path
    log_path: Path
    updater_path: Path
    success_path: Path
    package_path: Path
    temp_dir: Path


def _cleanup_prepared_update(prepared: Optional["PreparedUpdate"]) -> None:
    if prepared is None:
        return
    with contextlib.suppress(OSError):
        if prepared.lock_path.exists():
            prepared.lock_path.unlink()
    with contextlib.suppress(Exception):
        if prepared.package_path.exists():
            prepared.package_path.unlink()
    with contextlib.suppress(Exception):
        if prepared.temp_dir.exists():
            shutil.rmtree(prepared.temp_dir, ignore_errors=True)


def _create_temp_dir() -> Path:
    base = Path(tempfile.gettempdir())
    directory = base / f"allone_update_{int(time.time())}"
    directory.mkdir(parents=True, exist_ok=True)
    return directory


def _prepare_update_environment(
    update_info: UpdateInfo,
    log_callback: Callable[[str], None],
    base_dir: Optional[Path] = None,
) -> PreparedUpdate:
    resolved_base = (
        Path(base_dir).resolve() if base_dir is not None else _install_root().resolve()
    )
    paths = _paths(resolved_base)
    lock_path = paths["lock"]
    log_path = paths["log"]
    updater_path = paths["updater"]
    success_path = paths["success"]

    if lock_path.exists():
        raise UpdateError("Another update appears to be in progress. Please retry later.")

    lock_path.write_text(str(time.time()), encoding="utf-8")
    temp_dir = _create_temp_dir()
    package_path = temp_dir / update_info.file_name

    try:
        log_callback("Downloading update package…")
        _download_asset(
            update_info.download_url,
            package_path,
            update_info.download_size,
            log_callback,
        )
        if update_info.sha256:
            log_callback("Download completed. Verifying package integrity…")
            actual_hash = _calculate_sha256(package_path)
            if actual_hash.lower() != update_info.sha256.lower():
                raise UpdateError("Package checksum mismatch. Update aborted.")
        else:
            log_callback(
                "Download completed. No checksum provided; skipping integrity verification."
            )

        log_callback("Preparing installer…")
        _prepare_updater_script(updater_path)
    except Exception:
        _cleanup_prepared_update(
            PreparedUpdate(
                base_dir=resolved_base,
                lock_path=lock_path,
                log_path=log_path,
                updater_path=updater_path,
                success_path=success_path,
                package_path=package_path,
                temp_dir=temp_dir,
            )
        )
        raise

    return PreparedUpdate(
        base_dir=resolved_base,
        lock_path=lock_path,
        log_path=log_path,
        updater_path=updater_path,
        success_path=success_path,
        package_path=package_path,
        temp_dir=temp_dir,
    )


def _launch_update_process(
    target_executable: Path,
    target_pid: Optional[int],
    update_info: UpdateInfo,
    log_callback: Callable[[str], None],
    *,
    base_dir: Optional[Path] = None,
    wait_for_completion: bool = False,
) -> Optional[subprocess.Popen]:
    prepared: Optional[PreparedUpdate] = None
    try:
        prepared = _prepare_update_environment(update_info, log_callback, base_dir)
        args = [
            str(target_executable.resolve()),
            str(target_pid or ""),
            str(prepared.package_path),
            str(prepared.temp_dir / "work"),
            target_executable.name,
            str(prepared.log_path),
            str(prepared.success_path),
            str(prepared.lock_path),
            update_info.version,
        ]
        process = _launch_updater(
            prepared.updater_path,
            args,
            detach=not wait_for_completion,
        )
    except Exception:
        _cleanup_prepared_update(prepared)
        raise

    if target_pid:
        log_callback("Update installer launched. Application will exit to apply the update.")
    else:
        log_callback("Update installer launched. Application files will be refreshed shortly.")
    return process


def start_update_installation(
    update_info: UpdateInfo,
    log_callback: Callable[[str], None],
    target_executable: Path,
    *,
    target_pid: Optional[int] = None,
    base_dir: Optional[Path] = None,
    wait_for_completion: bool = False,
) -> Optional[subprocess.Popen]:
    """Initiate the installer batch script for the provided executable."""

    target_executable = Path(target_executable).resolve()
    if not target_executable.exists():
        raise UpdateError(f"Target executable not found: {target_executable}")

    resolved_base = Path(base_dir) if base_dir is not None else target_executable.parent
    return _launch_update_process(
        target_executable,
        target_pid,
        update_info,
        log_callback,
        base_dir=resolved_base,
        wait_for_completion=wait_for_completion,
    )


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

    start_update_installation(
        update_info,
        log_callback,
        Path(sys.executable).resolve(),
        target_pid=os.getpid(),
    )


def check_for_updates(
    app_instance,
    log_callback: Callable[[str], None],
    current_version: str,
    silent: bool = False,
    status_callback: Optional[Callable[[str, Dict[str, str]], None]] = None,
) -> None:
    timeout = CHECK_TIMEOUT_SILENT if silent else CHECK_TIMEOUT_INTERACTIVE

    if silent and hasattr(app_instance, "is_auto_update_enabled"):
        if not app_instance.is_auto_update_enabled():
            log_callback("Auto-update on startup disabled. Skipping update check.")
            return

    def _emit_status(state: str, **context: Optional[str]) -> None:
        if not status_callback:
            return
        try:
            payload = {
                key: str(value)
                for key, value in context.items()
                if value is not None
            }
            status_callback(state, payload)
        except Exception:
            pass

    _emit_status("checking")
    log_callback("Checking for updates…")
    try:
        remote_version = _remote_version_string(timeout)
    except requests.Timeout:
        log_callback("Update check timed out. Possibly offline.")
        _emit_status("timeout")
        return
    except Exception as exc:
        log_callback(f"Update check failed: {exc}")
        _emit_status("error", error=str(exc))
        if not silent:
            app_instance.after(
                0,
                lambda: messagebox.showerror("Update Check", f"Could not check for updates.\n\n{exc}"),
            )
        return

    if _version_is_not_newer(remote_version, current_version, log_callback):
        log_callback(f"✅ Your application is up-to-date. (Version: {current_version})")
        _emit_status("up_to_date", version=current_version)
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
        _emit_status("error", error=str(exc))
        if not silent:
            app_instance.after(
                0,
                lambda: messagebox.showerror(
                    "Update", f"Update information could not be retrieved.\n\n{exc}"
                ),
            )
        return

    _emit_status("update_available", version=update_info.version)
    if update_info.release_notes_url:
        log_callback(f"Latest release notes: {update_info.release_notes_url}")

    auto_install = (
        silent
        and hasattr(app_instance, "is_auto_update_enabled")
        and app_instance.is_auto_update_enabled()
    )

    if auto_install and hasattr(app_instance, "_start_update_flow"):
        log_callback(
            "Update available. Auto-update is enabled; preparing installation for version "
            f"{update_info.version}."
        )
        _emit_status("auto_installing", version=update_info.version)

        def _auto_start() -> None:
            try:
                app_instance._start_update_flow(update_info)
            except Exception as exc2:  # pragma: no cover - defensive UI feedback
                log_callback(f"Update failed: {exc2}")
                messagebox.showerror("Update", f"Update failed: {exc2}")

        app_instance.after(0, _auto_start)
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

