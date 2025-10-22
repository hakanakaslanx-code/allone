# updater.py
import os
import sys
import re
import time
import requests
import webbrowser
from tkinter import messagebox


def _is_windows_executable():
    return (
        sys.platform.startswith("win")
        and getattr(sys, "frozen", False)
        and sys.executable.lower().endswith(".exe")
    )


def _escape_for_batch(path):
    return path.replace("%", "%%")


def _fetch_latest_release_asset(owner, repo):
    api_url = f"https://api.github.com/repos/{owner}/{repo}/releases/latest"
    response = requests.get(api_url, headers={"Accept": "application/vnd.github+json"}, timeout=10)
    response.raise_for_status()
    release_info = response.json()
    assets = release_info.get("assets", [])
    for asset in assets:
        name = asset.get("name", "").lower()
        if name.endswith(".exe"):
            return asset
    raise RuntimeError("No executable asset found in the latest release.")


def _download_asset(download_url, destination_path, log_callback, total_size=None):
    with requests.get(download_url, stream=True, timeout=30) as response:
        response.raise_for_status()
        downloaded = 0
        last_logged_percent = -10
        with open(destination_path, "wb") as dest_file:
            for chunk in response.iter_content(chunk_size=1024 * 1024):
                if not chunk:
                    continue
                dest_file.write(chunk)
                downloaded += len(chunk)
                if total_size:
                    percent = int(downloaded * 100 / total_size)
                    if percent - last_logged_percent >= 10:
                        last_logged_percent = percent
                        log_callback(f"Downloading update... {percent}%")


def _schedule_windows_update(executable_path, downloaded_path, log_callback):
    exe_dir = os.path.dirname(executable_path)
    timestamp = int(time.time())
    script_path = os.path.join(exe_dir, f"apply_update_{timestamp}.bat")

    escaped_executable = _escape_for_batch(executable_path)
    escaped_download = _escape_for_batch(downloaded_path)

    script_contents = (
        "@echo off\r\n"
        "setlocal\r\n"
        f"set \"TARGET={escaped_executable}\"\r\n"
        f"set \"SOURCE={escaped_download}\"\r\n"
        ":retry\r\n"
        "copy /Y \"%SOURCE%\" \"%TARGET%\" >nul\r\n"
        "if errorlevel 1 (\r\n"
        "    timeout /T 1 /NOBREAK >nul\r\n"
        "    goto retry\r\n"
        ")\r\n"
        "start \"\" \"%TARGET%\"\r\n"
        "del \"%SOURCE%\"\r\n"
        "del \"%~f0\"\r\n"
    )

    with open(script_path, "w", encoding="utf-8") as script_file:
        script_file.write(script_contents)

    log_callback(f"Update script created at {script_path}")

    try:
        os.startfile(script_path)
    except AttributeError:
        # os.startfile may not exist in some environments; fall back to cmd
        import subprocess

        subprocess.Popen(["cmd", "/c", script_path], shell=False)

    return script_path

def check_for_updates(app_instance, log_callback, current_version, silent=False):
    log_callback("Checking for updates...")

    owner = "hakanakaslanx-code"
    repo = "allone"

    version_check_url = f"https://raw.githubusercontent.com/{owner}/{repo}/main/allone/app_ui.py"

    try:
        response = requests.get(version_check_url, timeout=10)
        response.raise_for_status()
        remote_script_content = response.text

        match = re.search(r"__version__\s*=\s*[\"'](.+?)[\"']", remote_script_content)
        if not match:
            if not silent:
                messagebox.showwarning("Update Check", "Could not determine remote version.")
            return

        remote_version = match.group(1)

        if remote_version > current_version:
            log_callback(f"--> New version ({remote_version}) available.")

            if silent:
                log_callback("Update available. Use 'Check for Updates' to install.")
                return

            if _is_windows_executable():
                if not messagebox.askyesno(
                    "Update Available",
                    (
                        f"A new version ({remote_version}) is available.\n\n"
                        "Do you want to download and install it automatically now?"
                    ),
                ):
                    return

                executable_path = os.path.abspath(sys.executable)
                exe_dir, exe_name = os.path.split(executable_path)
                download_path = os.path.join(exe_dir, f"{exe_name}.download")

                if os.path.exists(download_path):
                    try:
                        os.remove(download_path)
                    except OSError:
                        log_callback(f"Unable to remove old download file: {download_path}")

                try:
                    asset = _fetch_latest_release_asset(owner, repo)
                    download_url = asset.get("browser_download_url")
                    if not download_url:
                        raise RuntimeError("Download URL for the latest release is missing.")

                    total_size = asset.get("size")
                    log_callback(f"Downloading {asset.get('name', 'update package')}...")
                    _download_asset(download_url, download_path, log_callback, total_size)
                    log_callback("Download completed. Preparing installer...")

                    _schedule_windows_update(executable_path, download_path, log_callback)

                    messagebox.showinfo(
                        "Update",
                        (
                            "The application will close to finish installing the update.\n"
                            "It will restart automatically once the update is complete."
                        ),
                    )

                    app_instance.destroy()
                    sys.exit(0)

                except Exception as exc:
                    log_callback(f"Update failed: {exc}")
                    if os.path.exists(download_path):
                        try:
                            os.remove(download_path)
                        except OSError:
                            pass
                    messagebox.showerror(
                        "Update Failed",
                        (
                            "Automatic update could not be completed.\n\n"
                            f"Reason: {exc}\n\n"
                            "The latest release page will be opened so you can download it manually."
                        ),
                    )
                    latest_release_url = f"https://github.com/{owner}/{repo}/releases/latest"
                    webbrowser.open(latest_release_url)
                    return

            else:
                latest_release_url = f"https://github.com/{owner}/{repo}/releases/latest"
                if messagebox.askyesno(
                    "Update Available",
                    (
                        f"A new version ({remote_version}) is available.\n\n"
                        "Do you want to open the download page now?"
                    ),
                ):
                    log_callback(f"Opening download page: {latest_release_url}")
                    webbrowser.open(latest_release_url)

        else:
            log_callback(f"✅ Your application is up-to-date. (Version: {current_version})")
            if not silent:
                messagebox.showinfo("Update Check", "You are running the latest version.")

    except Exception as e:
        log_callback(f"Warning: Update check failed. Reason: {e}")
        if not silent:
            messagebox.showerror("Update Check Failed", f"Could not check for updates.\n\nReason: {e}")
