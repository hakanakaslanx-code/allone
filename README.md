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
