from __future__ import annotations

import importlib
import importlib.util
import os
import io
import re
import sys
import time
from pathlib import Path
from typing import Callable, Dict, Set, Tuple
from urllib.parse import quote_plus
from urllib.request import Request, urlopen

from modules.maps_scraper.models import Business, BusinessList

EMAIL_PATTERN = re.compile(r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}")
FACEBOOK_PATTERN = re.compile(r"https?://(?:www\.)?facebook\.com/[^\s\"'<>]+", re.IGNORECASE)
INSTAGRAM_PATTERN = re.compile(r"https?://(?:www\.)?instagram\.com/[^\s\"'<>]+", re.IGNORECASE)
LAT_LNG_PATTERN = re.compile(r"/@(-?\d+\.\d+),(-?\d+\.\d+)")


def build_query(search_term: str, location: str) -> str:
    return " ".join(part for part in [search_term.strip(), location.strip()] if part)


def sanitize_filename(value: str) -> str:
    cleaned = re.sub(r"\s+", "_", value.strip())
    cleaned = re.sub(r"[^A-Za-z0-9._-]+", "_", cleaned)
    return cleaned.strip("_") or "query"


def build_output_filename(query: str, extension: str) -> str:
    safe_query = sanitize_filename(query)
    return f"google_maps_data_{safe_query}.{extension}"


def _playwright_installed() -> bool:
    return importlib.util.find_spec("playwright") is not None


def _chromium_installed() -> bool:
    candidates = []
    browsers_path = os.environ.get("PLAYWRIGHT_BROWSERS_PATH")
    if browsers_path:
        candidates.append(Path(browsers_path))

    home = Path.home()
    candidates.extend(
        [
            home / ".cache" / "ms-playwright",
            home / "AppData" / "Local" / "ms-playwright",
            home / "Library" / "Caches" / "ms-playwright",
        ]
    )
    chromium_globs = [
        "chromium-*/chrome-linux/chrome",
        "chromium-*/chrome.exe",
        "chromium-*/chrome-win/chrome.exe",
        "chromium-*/chrome-mac/Chromium.app/Contents/MacOS/Chromium",
    ]
    for base in candidates:
        if base.exists():
            for pattern in chromium_globs:
                if any(base.glob(pattern)):
                    return True
    return False


def _is_frozen() -> bool:
    return bool(getattr(sys, "frozen", False))


def _persistent_browsers_path() -> Path:
    appdata = os.environ.get("APPDATA")
    if appdata:
        return Path(appdata) / "AllOneTool" / "pw-browsers"
    return Path.home() / ".local" / "share" / "AllOneTool" / "pw-browsers"


class _LogWriter(io.TextIOBase):
    def __init__(self, log: Callable[[str], None]) -> None:
        self._log = log
        self._buffer = ""

    def write(self, text: str) -> int:
        self._buffer += text
        while "\n" in self._buffer:
            line, self._buffer = self._buffer.split("\n", 1)
            if line:
                self._log(line)
        return len(text)

    def flush(self) -> None:
        if self._buffer:
            self._log(self._buffer)
            self._buffer = ""


def _install_chromium(log: Callable[[str], None]) -> None:
    playwright_main = importlib.import_module("playwright.__main__").main
    log("Installing Playwright Chromium...")
    writer = _LogWriter(log)
    stdout = sys.stdout
    stderr = sys.stderr
    try:
        sys.stdout = writer
        sys.stderr = writer
        playwright_main(["install", "chromium"])
    finally:
        writer.flush()
        sys.stdout = stdout
        sys.stderr = stderr


def ensure_playwright_ready(log: Callable[[str], None]) -> None:
    browsers_path = _persistent_browsers_path()
    browsers_path.mkdir(parents=True, exist_ok=True)
    os.environ.setdefault("PLAYWRIGHT_BROWSERS_PATH", str(browsers_path))
    log(f"Playwright browsers path: {os.environ['PLAYWRIGHT_BROWSERS_PATH']}")

    if _is_frozen():
        log("Running from a frozen app. Skipping runtime pip installs.")
        if not _playwright_installed():
            log("Playwright module missing in frozen build. Rebuild with --collect-all playwright.")
            return

        if _chromium_installed():
            log("Playwright Chromium already installed.")
            return

        log("Chromium not found. Downloading...")
        _install_chromium(log)
        return

    if not _playwright_installed():
        log("Playwright not found. Install it before running the scraper.")
        return

    if _chromium_installed():
        log("Playwright Chromium already installed.")
        return

    log("Chromium not found. Downloading...")
    _install_chromium(log)


def _safe_inner_text(locator, label: str, log: Callable[[str], None]) -> str:
    if locator is None or locator.count() == 0:
        log(f"Missing {label}.")
        return ""
    try:
        return locator.first.inner_text().strip()
    except Exception as exc:
        log(f"Failed to read {label}: {exc}")
        return ""


def _safe_attribute(locator, attribute: str, label: str, log: Callable[[str], None]) -> str:
    if locator is None or locator.count() == 0:
        log(f"Missing {label}.")
        return ""
    try:
        value = locator.first.get_attribute(attribute) or ""
        return value.strip()
    except Exception as exc:
        log(f"Failed to read {label}: {exc}")
        return ""


def _extract_lat_lng(url: str) -> Tuple[str, str]:
    match = LAT_LNG_PATTERN.search(url)
    if not match:
        return "", ""
    return match.group(1), match.group(2)


def _extract_review_data(text: str) -> Tuple[str, str]:
    avg_match = re.search(r"(\d+\.\d+)", text)
    count_match = re.search(r"(\d+[\d,.]*)\s+reviews", text, re.IGNORECASE)
    avg = avg_match.group(1) if avg_match else ""
    count = ""
    if count_match:
        count = count_match.group(1).replace(",", "")
    return avg, count


def _extract_socials_and_email(html: str) -> Dict[str, str]:
    facebook = ""
    instagram = ""
    email = ""
    if html:
        fb_match = FACEBOOK_PATTERN.search(html)
        ig_match = INSTAGRAM_PATTERN.search(html)
        email_match = EMAIL_PATTERN.search(html)
        facebook = fb_match.group(0) if fb_match else ""
        instagram = ig_match.group(0) if ig_match else ""
        email = email_match.group(0) if email_match else ""
    return {"facebook": facebook, "instagram": instagram, "email": email}


def _fetch_html(url: str, log: Callable[[str], None]) -> str:
    if not url:
        return ""
    try:
        request = Request(url, headers={"User-Agent": "Mozilla/5.0"})
        with urlopen(request, timeout=20) as response:
            return response.read().decode("utf-8", errors="ignore")
    except Exception as exc:
        log(f"Website fetch failed: {exc}")
        return ""


def _retry(action: Callable[[], None], log: Callable[[str], None]) -> None:
    for attempt in range(2):
        try:
            action()
            return
        except Exception as exc:
            if attempt == 0:
                log(f"Retrying after error: {exc}")
                time.sleep(1)
            else:
                log(f"Action failed after retry: {exc}")


def _collect_place_urls(page, max_listings: int, log: Callable[[str], None], stop_event) -> Iterable[str]:
    log("Collecting listings...")
    urls: Set[str] = set()
    stall_count = 0
    while len(urls) < max_listings and stall_count < 5 and not stop_event.is_set():
        anchors = page.locator('a[href*="/place/"]')
        count = anchors.count()
        for idx in range(count):
            href = anchors.nth(idx).get_attribute("href")
            if href and "/place/" in href:
                urls.add(href)
            if len(urls) >= max_listings:
                break
        if len(urls) < max_listings:
            feed = page.locator('div[role="feed"]')
            if feed.count() > 0:
                feed.evaluate("el => el.scrollBy(0, el.scrollHeight)")
            else:
                page.mouse.wheel(0, 1200)
            time.sleep(1.2)
            new_count = anchors.count()
            if new_count == count:
                stall_count += 1
            else:
                stall_count = 0
    return list(urls)[:max_listings]


def _scrape_place(page, url: str, include_socials: bool, log: Callable[[str], None]) -> Business:
    def goto_page() -> None:
        page.goto(url, wait_until="domcontentloaded", timeout=45000)

    _retry(goto_page, log)
    try:
        page.keyboard.press("Escape")
    except Exception:
        pass

    name = _safe_inner_text(page.locator("h1"), "name", log)
    address = _safe_inner_text(page.locator("button[data-item-id='address']"), "address", log)
    if not address:
        address = _safe_inner_text(page.locator("button[data-item-id^='address']"), "address", log)
    phone = _safe_inner_text(page.locator("button[data-item-id^='phone:tel:']"), "phone", log)
    website = _safe_attribute(page.locator("a[data-item-id='authority']"), "href", "website", log)

    reviews_text = _safe_inner_text(
        page.locator("button[jsaction*='reviews']"),
        "reviews",
        log,
    )
    if not reviews_text:
        reviews_text = _safe_inner_text(page.locator("span[aria-label*='reviews']"), "reviews", log)
    rating_text = _safe_attribute(
        page.locator("div[role='img'][aria-label*='stars']"),
        "aria-label",
        "rating",
        log,
    )
    reviews_average, reviews_count = _extract_review_data(f"{rating_text} {reviews_text}")
    latitude, longitude = _extract_lat_lng(page.url)

    facebook = ""
    instagram = ""
    email = ""
    if include_socials:
        html = _fetch_html(website, log) if website else page.content()
        socials = _extract_socials_and_email(html)
        facebook = socials["facebook"]
        instagram = socials["instagram"]
        email = socials["email"]

    return Business(
        name=name,
        address=address,
        website=website,
        phone_number=phone,
        reviews_count=reviews_count,
        reviews_average=reviews_average,
        latitude=latitude,
        longitude=longitude,
        facebook=facebook,
        instagram=instagram,
        email=email,
    )


def scrape_google_maps(
    search_term: str,
    location: str,
    max_listings: int,
    headless: bool,
    include_socials: bool,
    log: Callable[[str], None],
    progress: Callable[[int, int], None],
    stop_event,
) -> BusinessList:
    ensure_playwright_ready(log)

    importlib.invalidate_caches()
    playwright_sync = importlib.import_module("playwright.sync_api")
    sync_playwright = playwright_sync.sync_playwright

    query = build_query(search_term, location)
    encoded = quote_plus(query)
    search_url = f"https://www.google.com/maps/search/{encoded}"

    business_list = BusinessList()
    seen_keys: Set[Tuple[str, str]] = set()

    with sync_playwright() as playwright:
        browser = playwright.chromium.launch(headless=headless)
        context = browser.new_context()
        page = context.new_page()

        def goto_search() -> None:
            page.goto(search_url, wait_until="domcontentloaded", timeout=45000)

        _retry(goto_search, log)
        try:
            page.keyboard.press("Escape")
        except Exception:
            pass

        if stop_event.is_set():
            browser.close()
            return business_list

        place_urls = _collect_place_urls(page, max_listings, log, stop_event)
        total = len(place_urls)
        progress(0, max(total, max_listings))

        for index, url in enumerate(place_urls, start=1):
            if stop_event.is_set():
                log("Scrape cancelled by user.")
                break
            log(f"Scraping {index}/{total}: {url}")
            business = _scrape_place(page, url, include_socials, log)
            key = (business.name.strip().lower(), business.address.strip().lower())
            if key in seen_keys or not business.name:
                log("Skipping duplicate or empty entry.")
            else:
                seen_keys.add(key)
                business_list.business_list.append(business)
            progress(index, max(total, max_listings))

        browser.close()

    log(f"Scrape finished. {len(business_list.business_list)} businesses collected.")
    return business_list
