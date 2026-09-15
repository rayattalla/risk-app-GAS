"""Shared utilities for LAUSD data scrapers.

Provides:
- HTTP session with browser-like headers (requests)
- Playwright browser context (for JS-rendered pages)
- Excel workbook writer with tab + header management
- Provenance tracking columns on every output
"""

import os
import sys
import time
import json
import hashlib
import logging
from datetime import datetime, timezone

import requests
from bs4 import BeautifulSoup

# Try pandas/openpyxl for Excel; fall back gracefully
try:
    import pandas as pd
    HAS_PANDAS = True
except ImportError:
    HAS_PANDAS = False

try:
    import openpyxl
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYX = False

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------

DEFAULT_HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8",
    "Accept-Language": "en-US,en;q=0.9",
    "Accept-Encoding": "gzip, deflate",
    "Connection": "keep-alive",
    "Upgrade-Insecure-Requests": "1",
}

# Provenance columns added to every data tab
PROVENANCE_COLS = ["scrape_date", "source_url", "scrape_note"]

# Output paths
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_DIR = os.path.dirname(SCRIPT_DIR)
OUTPUT_DIR = os.path.join(PROJECT_DIR, "output")
EXCEL_PATH = os.path.join(
    os.path.expanduser("~"),
    "Documents", "GIT", "gas-agent-platform", "RiskAI_Agents_v2.xlsx"
)

# ---------------------------------------------------------------------------
# Logging
# ---------------------------------------------------------------------------

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
log = logging.getLogger("lausd-scraper")

def info(msg):  log.info(msg)
def warn(msg):  log.warning(msg)
def error(msg): log.error(msg)

# ---------------------------------------------------------------------------
# HTTP helpers
# ---------------------------------------------------------------------------

def make_session():
    """Return a requests.Session with browser-like headers."""
    s = requests.Session()
    s.headers.update(DEFAULT_HEADERS)
    return s

def fetch_html(url, session=None, timeout=30):
    """Fetch HTML content from a URL. Returns text or None on failure."""
    s = session or make_session()
    try:
        resp = s.get(url, timeout=timeout)
        resp.raise_for_status()
        return resp.text
    except requests.RequestException as e:
        error(f"fetch_html failed for {url}: {e}")
        return None

def fetch_pdf(url, session=None, timeout=60):
    """Download a PDF file. Returns bytes or None."""
    s = session or make_session()
    try:
        resp = s.get(url, timeout=timeout)
        resp.raise_for_status()
        return resp.content
    except requests.RequestException as e:
        error(f"fetch_pdf failed for {url}: {e}")
        return None

def parse_html(html_text):
    """Parse HTML string into BeautifulSoup object."""
    return BeautifulSoup(html_text, "lxml")

def now_iso():
    """Current UTC timestamp in ISO format."""
    return datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")

def now_str():
    """Current date string for provenance."""
    return datetime.now(timezone.utc).strftime("%Y-%m-%d")

# ---------------------------------------------------------------------------
# Playwright helpers (for JS-rendered pages)
# ---------------------------------------------------------------------------

def get_browser(headless=True):
    """Launch a Playwright Chromium browser. Returns browser instance."""
    from playwright.sync_api import sync_playwright
    pw = sync_playwright().start()
    browser = pw.chromium.launch(headless=headless)
    return browser, pw

def fetch_html_playwright(url, wait_ms=5000, timeout=60000):
    """Use Playwright to fetch JS-rendered page HTML. Returns text or None."""
    browser, pw = get_browser(headless=True)
    try:
        page = browser.new_page()
        page.set_extra_http_headers({"Accept-Language": "en-US,en;q=0.9"})
        page.goto(url, timeout=timeout)
        page.wait_for_load_state("networkidle", timeout=30000)
        page.wait_for_timeout(wait_ms)
        return page.content()
    except Exception as e:
        error(f"playwright fetch failed for {url}: {e}")
        return None
    finally:
        browser.close()
        pw.stop()

# ---------------------------------------------------------------------------
# Excel helpers
# ---------------------------------------------------------------------------

def ensure_output_dir():
    """Create output directory if it doesn't exist."""
    os.makedirs(OUTPUT_DIR, exist_ok=True)

def write_tab(rows, tab_name, headers, excel_path=None, session=None):
    """Write rows to an Excel workbook tab.

    Parameters
    ----------
    rows : list[dict]  — data rows (keys match headers)
    tab_name : str     — worksheet name
    headers : list[str] — column order
    excel_path : str   — override default path
    session : requests.Session — unused, kept for API compatibility

    Returns the output file path.
    """
    path = excel_path or EXCEL_PATH
    ensure_output_dir()
    # Dedupe provenance columns against declared headers BEFORE concatenating.
    # Without this, a scraper that already lists "source_url" in its own
    # headers gets it twice -- which is exactly how the original Staff column
    # shift bug happened (a duplicate column name confuses header-driven
    # consumers that index by header name).
    declared = [str(h) for h in headers]
    provenance = [c for c in PROVENANCE_COLS if c not in declared]
    all_headers = declared + provenance

    if HAS_PANDAS:
        df = pd.DataFrame(rows, columns=all_headers)
        # If file exists, append/replace the tab
        if os.path.exists(path):
            # Read existing tabs
            existing_sheets = {}
            try:
                xl = pd.ExcelFile(path)
                for sn in xl.sheet_names:
                    existing_sheets[sn] = xl.parse(sn)
            except Exception:
                pass
            existing_sheets[tab_name] = df
            with pd.ExcelWriter(path, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
                for sn, dfx in existing_sheets.items():
                    dfx.to_excel(writer, sheet_name=sn, index=False)
        else:
            with pd.ExcelWriter(path, engine="openpyxl") as writer:
                df.to_excel(writer, sheet_name=tab_name, index=False)
    else:
        # Fallback: use openpyxl directly
        if os.path.exists(path):
            wb = openpyxl.load_workbook(path)
        else:
            wb = openpyxl.Workbook()
            # Remove default sheet
            if "Sheet" in wb.sheetnames:
                wb.remove(wb["Sheet"])
        if tab_name in wb.sheetnames:
            ws = wb[tab_name]
            wb.remove(ws)
        ws = wb.create_sheet(tab_name)
        ws.append(all_headers)
        for row in rows:
            ws.append([row.get(h, "") for h in all_headers])
        wb.save(path)

    return path

def write_json(data, filename):
    """Write data as JSON to the output directory."""
    ensure_output_dir()
    path = os.path.join(OUTPUT_DIR, filename)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)
    info(f"Wrote {len(data) if isinstance(data, list) else 'data'} to {path}")
    return path

def add_provenance(rows, source_url, note=""):
    """Add provenance columns to each row dict."""
    today = now_str()
    for row in rows:
        row["scrape_date"] = today
        row["source_url"] = source_url
        row["scrape_note"] = note
    return rows

def dedupe_rows(rows, key="job_id"):
    """Remove duplicate rows by key, keeping first occurrence."""
    seen = set()
    out = []
    for r in rows:
        k = r.get(key, "")
        if k and k in seen:
            continue
        if k:
            seen.add(k)
        out.append(r)
    return out
