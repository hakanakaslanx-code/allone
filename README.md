# AllOne Tools

## Google Maps Scraper (Playwright)

1. Open **Google Maps Scraper** from the sidebar.
2. Enter a **Search Term** (required) and an optional **Location**.
3. Set **Max Listings** and click **Start**.
4. When the scrape finishes, click **Export Excel** or **Export CSV**.

**Output**

- Excel: `output/google_maps_data_<query>.xlsx`
- CSV: `output/google_maps_data_<query>.csv`

Spaces in the query are converted to underscores.

**Notes**

- On first run, Playwright and Chromium are installed automatically.
- Toggle **Headless** off to watch the browser automation.

## Windows packaging (PyInstaller onedir + Inno Setup)

Playwright is bundled inside the build so end users do **not** need Python/pip.
Chromium is downloaded on first run to `%APPDATA%\\AllOneTool\\pw-browsers`.

### Build steps (Windows)

The easiest way to build the executable is using the provided batch script:

1.  Make sure you have requirements installed: `pip install -r allone/requirements.txt`
2.  Run `build.bat`

Alternatively, you can run PyInstaller manually:
```bash
python -m PyInstaller --noconfirm allone_onedir.spec
```

The onedir output is `dist/AllOne Tools/`.

### Installer (Inno Setup)

Use the provided `installer/AllOne.iss` (expects the onedir output above):

```bash
"C:\\Program Files (x86)\\Inno Setup 6\\ISCC.exe" installer\\AllOne.iss
```
