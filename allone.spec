# -*- mode: python ; coding: utf-8 -*-
from __future__ import annotations

from pathlib import Path
import os

from PyInstaller.utils.hooks import collect_all

project_root = Path(os.getcwd()).resolve()

datas, binaries, hiddenimports = collect_all("playwright")
hiddenimports += ["playwright.sync_api"]

# Add icon to datas
datas += [('icon.ico', '.')]

analysis = Analysis(
    [str(project_root / "allone" / "main.py")],
    pathex=[str(project_root)],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)
pyz = PYZ(analysis.pure)

exe = EXE(
    pyz,
    analysis.scripts,
    analysis.binaries,
    analysis.zipfiles,
    analysis.datas,
    [],
    name="AllOne Tools",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False, # Disabled to reduce AV false positives and hide terminal
    icon=str(project_root / "icon.ico"),
)
