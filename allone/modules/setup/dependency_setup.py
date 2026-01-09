"""Dependency setup helpers for AllOne Tools."""

from __future__ import annotations

import importlib
import os
from pathlib import Path
import subprocess
import sys
from typing import Callable, Dict, Iterable, List, Optional


LogCallback = Callable[[str], None]
ProgressCallback = Callable[[int], None]


def run_setup(
    log_callback: LogCallback,
    progress_callback: ProgressCallback,
    cancel_flag,
    prompt_callback: Optional[Callable[[List[str]], bool]] = None,
) -> Dict[str, List[str]]:
    """Run dependency checks/installations and return a summary dict."""

    summary = {
        "ok": [],
        "missing_unfixable": [],
        "failed": [],
        "frozen_missing_required": [],
    }

    def log(message: str) -> None:
        log_callback(message)

    def set_progress(value: int) -> None:
        progress_callback(max(0, min(100, value)))

    def should_cancel() -> bool:
        return bool(cancel_flag and getattr(cancel_flag, "is_set", lambda: False)())

    def mark_ok(item: str) -> None:
        summary["ok"].append(item)

    def mark_missing_unfixable(item: str) -> None:
        summary["missing_unfixable"].append(item)

    def mark_failed(item: str, reason: str) -> None:
        summary["failed"].append(f"{item}: {reason}")

    def mark_required_missing(item: str) -> None:
        summary["frozen_missing_required"].append(item)

    def module_available(module_name: str) -> bool:
        try:
            importlib.import_module(module_name)
        except ModuleNotFoundError:
            return False
        return True

    def install_pip_packages(packages: Iterable[str]) -> bool:
        command = [sys.executable, "-m", "pip", "install", *packages]
        log(f"Running: {' '.join(command)}")
        process = subprocess.Popen(
            command,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
        )
        assert process.stdout is not None
        for line in process.stdout:
            if should_cancel():
                process.terminate()
                return False
            log(line.rstrip())
        return process.wait() == 0

    def run_playwright_install(browsers_path: Path) -> bool:
        env = os.environ.copy()
        env["PLAYWRIGHT_BROWSERS_PATH"] = str(browsers_path)
        command = [sys.executable, "-m", "playwright", "install", "chromium"]
        log(f"Running: {' '.join(command)}")
        process = subprocess.Popen(
            command,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            env=env,
        )
        assert process.stdout is not None
        for line in process.stdout:
            if should_cancel():
                process.terminate()
                return False
            log(line.rstrip())
        return process.wait() == 0

    def playwright_chromium_installed() -> bool:
        from playwright.sync_api import sync_playwright

        with sync_playwright() as playwright:
            executable = Path(playwright.chromium.executable_path)
        return executable.exists()

    set_progress(0)
    log("Starting dependency setup...")

    frozen = bool(getattr(sys, "frozen", False))
    if frozen:
        log("Detected frozen build (PyInstaller). Pip installs are disabled.")
    else:
        log("Detected development environment (pip installs allowed).")

    if should_cancel():
        log("Setup cancelled.")
        return summary

    steps = 5
    step_index = 0

    required_modules = {
        "playwright": "Playwright",
        "PIL": "Pillow (ImageTk)",
        "openpyxl": "openpyxl",
        "pandas": "pandas",
    }

    missing_modules: List[str] = []
    for module_name, label in required_modules.items():
        if module_available(module_name):
            mark_ok(label)
        else:
            missing_modules.append(module_name)
            if frozen:
                mark_missing_unfixable(label)
                mark_required_missing(label)
                log(f"Missing module in frozen build: {label}")
            else:
                log(f"Missing module: {label}")

    step_index += 1
    set_progress(int(step_index / steps * 100))

    if should_cancel():
        log("Setup cancelled.")
        return summary

    if not frozen and missing_modules:
        install = False
        if prompt_callback is not None:
            install = prompt_callback(missing_modules)
        if install:
            if not install_pip_packages(missing_modules):
                for module_name in missing_modules:
                    label = required_modules.get(module_name, module_name)
                    mark_failed(label, "pip install failed")
            else:
                for module_name in missing_modules:
                    label = required_modules.get(module_name, module_name)
                    if module_available(module_name):
                        mark_ok(label)
                    else:
                        mark_failed(label, "still missing after pip install")
        else:
            for module_name in missing_modules:
                label = required_modules.get(module_name, module_name)
                mark_missing_unfixable(label)

    step_index += 1
    set_progress(int(step_index / steps * 100))

    if should_cancel():
        log("Setup cancelled.")
        return summary

    playwright_ready = module_available("playwright")
    if playwright_ready:
        appdata = os.getenv("APPDATA")
        if not appdata:
            appdata = str(Path.home() / "AppData" / "Roaming")
        browsers_path = Path(appdata) / "AllOneTool" / "pw-browsers"
        browsers_path.mkdir(parents=True, exist_ok=True)
        os.environ["PLAYWRIGHT_BROWSERS_PATH"] = str(browsers_path)
        log(f"Playwright browsers path: {browsers_path}")

        try:
            chromium_installed = playwright_chromium_installed()
        except Exception as exc:
            mark_failed("Playwright Chromium", str(exc))
            chromium_installed = False

        if chromium_installed:
            mark_ok("Playwright Chromium")
            log("Playwright Chromium already installed.")
        else:
            log("Downloading Playwright Chromium...")
            if run_playwright_install(browsers_path):
                if playwright_chromium_installed():
                    mark_ok("Playwright Chromium")
                    log("Playwright Chromium installed successfully.")
                else:
                    mark_failed("Playwright Chromium", "install completed but browser not found")
            else:
                mark_failed("Playwright Chromium", "install command failed")
    else:
        if frozen:
            log("Playwright module missing in frozen build; Chromium install skipped.")
        else:
            log("Playwright module missing; Chromium install skipped.")

    step_index += 1
    set_progress(int(step_index / steps * 100))

    if should_cancel():
        log("Setup cancelled.")
        return summary

    step_index += 1
    set_progress(int(step_index / steps * 100))

    if should_cancel():
        log("Setup cancelled.")
        return summary

    set_progress(100)

    log("Dependency setup summary:")
    if summary["ok"]:
        log(f"✅ Installed/OK: {', '.join(summary['ok'])}")
    if summary["missing_unfixable"]:
        log(f"⚠️ Missing (unfixable here): {', '.join(summary['missing_unfixable'])}")
    if summary["failed"]:
        log(f"❌ Failed: {', '.join(summary['failed'])}")
    if not (summary["ok"] or summary["missing_unfixable"] or summary["failed"]):
        log("No checks were performed.")

    return summary
