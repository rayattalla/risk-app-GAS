"""Salary Schedule scraper — personnel.lausd.org

Downloads the LAUSD Classified Employee Salary Schedule PDF and extracts
class code, class title, unit, rate type, and step salaries.

Source: https://personnel.lausd.org/apps/pages/index.jsp?uREC_ID=4423577&type=d&pREC_ID=2660958

Output: scrapers/output/salary_schedule.json + RiskAI_Agents_v2.xlsx (tab: "Classifications")
"""

import re
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _utils import make_session, fetch_pdf, add_provenance, write_json, now_str, info, warn, error

# The salary schedule PDF URL (found on the personnel.lausd.org page)
SALARY_PDF_URL = (
    "https://media.edlio.net/cfce5b28/ebcec7af/f87b7244/"
    "477dc75f59534ec4a4d493faa0225cbf?_=July%201%202026%20Salary%20Schedule%20082826.pdf"
)
SALARY_PAGE_URL = (
    "https://personnel.lausd.org/apps/pages/index.jsp?uREC_ID=4423577&type=d&pREC_ID=2660958"
)

# Columns for the Excel "Classifications" worksheet (maps to district-data tab)
CLASSIFICATIONS_HEADERS = [
    "class_code",
    "class_title",
    "unit",
    "rate_type",
    "step_1",
    "step_2",
    "step_3",
    "step_4",
    "step_5",
    "step_6",
    "step_7",
    "step_8",
    "step_9",
    "step_10",
    "hourly_rate",
    "schedule_year",
]


def find_salary_pdf_url(session=None):
    """Scan the personnel.lausd.org salary schedule page for the latest PDF link.

    Falls back to the known SALARY_PDF_URL if the page can't be fetched.
    Filters for links containing "alary" (Salary Schedule) in URL or text.
    """
    s = session or make_session()

    # Try fetching the Edlio page with requests
    try:
        resp = s.get(SALARY_PAGE_URL, timeout=30)
        if resp.status_code == 200:
            text = resp.text
            # Find all PDF links with surrounding text
            # Pattern: <a href="URL">TEXT</a>
            pattern = r'<a[^>]+href=["\']([^"\']*\.pdf[^"\']*)["\'][^>]*>([^<]*)</a>'
            matches = re.findall(pattern, text, re.IGNORECASE)
            for href, link_text in matches:
                combined = href + " " + link_text
                if "alary" in combined.lower() or "schedule" in combined.lower():
                    info(f"Found salary PDF: {link_text.strip()} -> {href[:80]}")
                    return href
    except Exception as e:
        warn(f"Direct fetch of salary page failed ({e})")

    return SALARY_PDF_URL


def clean_title(title):
    """Remove PDF artifacts from class titles.

    The PDF includes footnotes like 'Classifications italicized in blue font
    are currently FROZEN 1 8/28/2026' that get appended to titles during
    text extraction.
    """
    title = title.strip()
    # Remove frozen note suffixes
    title = re.split(r'\s*Classifications italicized', title)[0]
    title = re.split(r'\s*FROZEN', title, flags=re.IGNORECASE)[0]
    # Remove trailing page numbers/footnotes
    title = re.split(r'\s+\d\s+\d+/\d+/\d+\s*$', title)[0]
    title = re.split(r'\s+Page \d+$', title)[0]
    return title.strip()


def parse_salary_page_text(text):
    """Parse raw text extracted from a salary schedule PDF page.

    The PDF has a repeating header on each page:
      Los Angeles Unified School District - Personnel Commission
      July 1, 2026 Classified Salary Schedule
      Class Code | Class Title | Unit | Hourly/Monthly Rate | Step 1 ... Step 10 | Hourly

    Each data row looks like:
      2375 Absence Coordinator SS Hourly $62.65279 $66.04373 ... $73.29779 Hourly
    or:
      1890 Area Facilities Services Director JJ Monthly $12,576.11 $13,278.04 ...
    """
    # Find schedule year from header
    year_match = re.search(r'(\d{4})\s+(?:Classified\s+)?Salary\s+Schedule', text, re.IGNORECASE)
    schedule_year = f"July 1, {year_match.group(1)}" if year_match else ""

    rows = []
    lines = text.split("\n")

    current_row = None
    for line in lines:
        line = line.strip()
        if not line:
            continue

        # Skip header/footer lines (don't start with a 4-digit code)
        if not re.match(r'^\d{4}\s', line):
            # Skip known header/footer text
            if any(h in line for h in ["Los Angeles", "Salary Schedule", "Class Code",
                                        "Step 1", "Step 2", "Step 3", "Unit", "Rate",
                                        "Page ", "Classifications italicized"]):
                continue
            # Continuation of previous row's title
            if current_row is not None and "$" not in line and not re.match(r'^\d{4}\s', line):
                cleaned = clean_title(line)
                if cleaned:
                    current_row["class_title"] += " " + cleaned
                continue
            continue

        # Match data row: 4-digit code + title + unit + rate_type + step values
        match = re.match(
            r'^(\d{4})\s+(.+?)\s+(SS|JJ)\s+(Hourly|Monthly)\s+(.+)$',
            line
        )
        if match:
            # Save previous row
            if current_row:
                rows.append(current_row)

            code = match.group(1)
            title_raw = match.group(2)
            unit = match.group(3)
            rate_type = match.group(4)
            remainder = match.group(5)

            # Clean title
            title = clean_title(title_raw)

            # Extract all numeric values from remainder
            amounts = re.findall(r'[\d,]+(?:\.\d+)?', remainder)
            step_vals = []
            for amt in amounts:
                clean = amt.replace(",", "")
                try:
                    val = float(clean)
                    step_vals.append(val)
                except ValueError:
                    continue

            # For Hourly positions, the last value may be a standalone hourly rate column
            hourly_rate = ""
            if rate_type == "Hourly" and len(step_vals) > 10:
                hourly_rate = format_step_value(step_vals[-1], rate_type)
                step_vals = step_vals[:-1]

            # Map steps 1-10
            step_names = [f"step_{i}" for i in range(1, 11)]
            steps = {}
            for i, val in enumerate(step_vals[:10]):
                steps[step_names[i]] = format_step_value(val, rate_type)

            current_row = {
                "class_code": code,
                "class_title": title,
                "unit": unit,
                "rate_type": rate_type,
                "schedule_year": schedule_year,
            }
            for i in range(1, 11):
                current_row[f"step_{i}"] = steps.get(f"step_{i}", "")
            current_row["hourly_rate"] = hourly_rate

    # Don't forget the last row
    if current_row:
        rows.append(current_row)

    return rows


def parse_salary_pdf(pdf_bytes):
    """Parse a salary schedule PDF byte stream into structured rows.

    pypdf's extract_text() flattens ruled tables into a single line per
    page, which is why 189 class codes came back with zero step values --
    the table structure was lost in flattening. pdfplumber's extract_table()
    preserves the ruled-cell grid, which is what this format needs.

    Falls back to pypdf line-parsing if pdfplumber is unavailable, so the
    scraper still runs in an environment without it installed.
    """
    import io

    rows = []

    # --- Primary path: pdfplumber table extraction ---
    try:
        import pdfplumber
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            info(f"PDF has {len(pdf.pages)} pages")
            for i, page in enumerate(pdf.pages):
                tables = page.extract_tables()
                page_rows = 0
                for table in tables:
                    parsed = _parse_salary_table_(table, i + 1)
                    if parsed:
                        rows.extend(parsed)
                        page_rows += len(parsed)
                info(f"  Page {i+1}: {page_rows} classifications from {len(tables)} table(s)")
        if rows:
            rows = _dedupe_by_class_code_(rows)
            info(f"Total unique classifications: {len(rows)}")
            return rows
        warn("pdfplumber found no table rows; falling back to pypdf line parsing")
    except ImportError:
        warn("pdfplumber not installed; falling back to pypdf line parsing")
    except Exception as e:
        warn(f"pdfplumber extraction failed ({e}); falling back to pypdf line parsing")

    # --- Fallback path: pypdf text extraction ---
    from pypdf import PdfReader
    reader = PdfReader(io.BytesIO(pdf_bytes))
    info(f"PDF has {len(reader.pages)} pages (pypdf fallback)")

    for i, page in enumerate(reader.pages):
        text = page.extract_text() or ""
        parsed = parse_salary_page_text(text)
        info(f"  Page {i+1}: extracted {len(parsed)} classifications")
        rows.extend(parsed)

    rows = _dedupe_by_class_code_(rows)
    info(f"Total unique classifications: {len(rows)}")
    return rows


def _parse_salary_table_(table, page_num):
    """Parse a pdfplumber-extracted table grid into classification rows.

    pdfplumber returns list[list[str | None]], one inner list per row, cell
    values as strings. The header row identifies columns; data rows follow.
    """
    if not table or len(table) < 2:
        return []

    # Find the header row: the first row whose cells include "Class Code"
    header_idx = None
    for i, row in enumerate(table[:6]):
        joined = " ".join(str(c or "") for c in row).lower()
        if "class code" in joined or "class_code" in joined:
            header_idx = i
            break
    if header_idx is None:
        return []

    header_cells = [str(c or "").strip().lower() for c in table[header_idx]]
    col = {name: idx for idx, name in enumerate(header_cells) if name}

    # Map step columns by position: step_1..step_10 are the numeric cells
    # between the rate-type column and any trailing hourly column.
    step_idxs = []
    for idx, name in header_cells:
        m = re.match(r"step\s*0?(\d)", name)
        if m:
            step_idxs.append((int(m.group(1)), idx))
    if not step_idxs:
        # Fallback: take the 10 consecutive numeric columns after rate_type
        rate_idx = col.get("rate_type", col.get("unit", 2))
        for j in range(rate_idx + 1, min(rate_idx + 11, len(header_cells))):
            step_idxs.append((len(step_idxs) + 1, j))

    rows = []
    for row in table[header_idx + 1:]:
        if not row or all(c is None or str(c).strip() == "" for c in row):
            continue
        code = str(row[col.get("class_code", 0)] or "").strip()
        if not re.match(r"^\d{4}$", code):
            continue

        title = clean_title(str(row[col.get("class_title", 1)] or ""))
        unit = str(row[col.get("unit", 2)] or "").strip()
        rate_type = str(row[col.get("rate_type", 3)] or "").strip() or "Monthly"

        entry = {
            "class_code": code,
            "class_title": title,
            "unit": unit,
            "rate_type": rate_type,
            "schedule_year": "",
        }
        for n, idx in step_idxs:
            entry[f"step_{n}"] = _format_step_value_(_to_float(row[idx]), rate_type)
        entry["hourly_rate"] = ""
        rows.append(entry)

    return rows


def _to_float(cell):
    """Parse a cell that may be a number, a formatted string, or None."""
    if cell is None:
        return None
    if isinstance(cell, (int, float)):
        return float(cell)
    s = str(cell).replace(",", "").replace("$", "").strip()
    if not s:
        return None
    try:
        return float(s)
    except ValueError:
        return None


def _format_step_value_(val, rate_type):
    """Format a step value appropriately based on rate type."""
    if val is None:
        return ""
    if val < 100:
        return f"{val:.5f}".rstrip("0").rstrip(".")
    elif val < 10000:
        return f"{val:.2f}"
    else:
        if val == int(val):
            return f"{int(val):,}"
        return f"{val:,.2f}"


def _dedupe_by_class_code_(rows):
    """Keep the first row for each class_code."""
    seen = set()
    out = []
    for r in rows:
        code = r.get("class_code", "")
        if code and code in seen:
            continue
        if code:
            seen.add(code)
        out.append(r)
    return out


def scrape_salary_schedule():
    """Main entry point: download and parse the LAUSD salary schedule PDF.

    Returns
    -------
    list[dict] — one dict per classification
    """
    session = make_session()

    pdf_url = find_salary_pdf_url(session)
    info(f"Downloading salary PDF from: {pdf_url[:80]}...")

    pdf_data = fetch_pdf(pdf_url, session=session, timeout=60)
    if not pdf_data:
        error("Failed to download salary schedule PDF")
        return []

    rows = parse_salary_pdf(pdf_data)
    rows = add_provenance(rows, SALARY_PAGE_URL, f"Salary Schedule PDF from {pdf_url[:80]}")

    return rows


def main():
    import argparse
    parser = argparse.ArgumentParser(description="Scrape LAUSD salary schedule PDF")
    parser.add_argument("--json", action="store_true", help="Also write JSON output")
    args = parser.parse_args()

    rows = scrape_salary_schedule()

    if args.json:
        path = write_json(rows, "salary_schedule.json")

    from _utils import write_tab
    path = write_tab(rows, "Classifications", CLASSIFICATIONS_HEADERS)
    info(f"Wrote {len(rows)} classifications to {path}")

    return rows


if __name__ == "__main__":
    main()
