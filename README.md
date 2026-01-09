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

## PyInstaller (Windows one-file)

When packaging, Playwright must be bundled with the app. Runtime installs using the
exe are not supported.

**Option A (bundle Playwright + browsers):**

1. Install Playwright + Chromium on the build machine:
   `python -m pip install playwright`
   `python -m playwright install chromium`
2. Build with Playwright collected and the browser cache bundled as
   `playwright-browsers`:

   ```bash
   pyinstaller --onefile --collect-all playwright --hidden-import=playwright.sync_api \
     --add-data "%LOCALAPPDATA%\\ms-playwright;playwright-browsers" \
     allone/main.py
   ```

**Option B (bundle Playwright only, install Chromium via system Python):**

```bash
pyinstaller --onefile --collect-all playwright --hidden-import=playwright.sync_api allone/main.py
```
