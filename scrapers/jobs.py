"""Jobs scraper — LAUSD careers.lausd.org

Scrapes all job postings from the LAUSD careers sitemap and individual job pages.
Uses requests + BeautifulSoup (no Playwright needed — job pages are server-rendered).

Job URL format:
  Classified:  https://careers.lausd.org/pc/job/{slug}/{id}/
  Certificated: https://careers.lausd.org/hr/job/{slug}/{id}/

Sitemap: https://careers.lausd.org/sitemap.xml

Output: scrapers/output/jobs.json + RiskAI_Agents_v2.xlsx (tab: "Jobs")
"""

import re
import sys
import os

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _utils import (
    make_session, fetch_html, parse_html, add_provenance,
    dedupe_rows, write_json, now_str, info, warn, error
)

SITEMAP_URL = "https://careers.lausd.org/sitemap.xml"

# Career-area listing pages on SAP SuccessFactors RMK (company losangel01).
# Each page lists currently-OPEN postings for that area. Discovering from
# these pages (rather than reusing stored posting URLs) is what fixes the
# 520-row historical-accumulation problem: a closed posting's body is
# replaced with "This position has closed or is no longer available." and
# SALARY DETAILS / APPLICATION FILING DATES are removed entirely, so a
# sitemap-driven scrape was silently writing 520 rows of dead pages with
# empty salary and no explanation.
CAREER_AREA_PAGES = [
    "https://careers.lausd.org/pc/go/Information-Technology-Careers/9597200/",
    "https://careers.lausd.org/pc/go/Administrative-Careers/9597199/",
    "https://careers.lausd.org/pc/go/School-Based-Careers/9597201/",
    "https://careers.lausd.org/pc/go/Classified-Careers/9597202/",
    "https://careers.lausd.org/pc/go/Certificated-Careers/9597203/",
    "https://careers.lausd.org/pc/go/Student-Services-Careers/9597204/",
    "https://careers.lausd.org/pc/go/Facilities-Careers/9597205/",
    "https://careers.lausd.org/pc/go/Transportation-Careers/9597206/",
    "https://careers.lausd.org/pc/go/Food-Services-Careers/9597207/",
    "https://careers.lausd.org/pc/go/Safety-Security-Careers/9597208/",
    "https://careers.lausd.org/pc/go/Finance-Careers/9597209/",
    "https://careers.lausd.org/pc/go/Health-Services-Careers/9597210/",
]

# Body text that SuccessFactors swaps in when a posting closes. Any page
# containing this is skipped (or written with status=closed, never written
# with empty salary and no explanation).
CLOSED_POSTING_MARKERS = [
    "This position has closed or is no longer available",
    "no longer available",
    "position has closed",
    "posting is closed",
    "this job is no longer posted",
]

# Tabs for the Excel "Jobs" worksheet
JOBS_HEADERS = [
    "job_id",
    "title",
    "job_type",
    "department",
    "salary_min",
    "salary_max",
    "salary_schedule",
    "filing_date_start",
    "filing_date_end",
    "date_posted",
    "valid_through",
    "location",
    "position_summary",
    "duties",
    "requirements",
    "education",
    "experience",
    "work_year",
    "status",
    "url",
]


def get_job_urls_from_sitemap(session):
    """DEPRECATED — kept for reference only.

    The sitemap accumulates every posting ever published, including closed
    ones. Scraping it produced 520 rows of dead pages with empty salary and
    no explanation. Use discover_open_postings_() instead.
    """
    html = fetch_html(SITEMAP_URL, session=session)
    if not html:
        return []
    urls = re.findall(r'<loc>(https://careers\.lausd\.org/(?:pc|hr)/job/[^<]+)</loc>', html)
    info(f"Found {len(urls)} job URLs in sitemap")
    return urls


def discover_open_postings_(session):
    """Discover currently-OPEN postings from the 12 career-area listing pages.

    Each listing page is server-rendered HTML containing links to job
    postings. We collect every /pc/job/ or /hr/job/ link, then return the
    unique set. Closed postings are NOT filtered out here -- they are
    detected per-page in parse_job_page() so they can be written with
    status=closed rather than silently dropped.

    Returns list of job URLs (dozens, not hundreds).
    """
    seen = set()
    for area_url in CAREER_AREA_PAGES:
        info(f"Scanning career area: {area_url}")
        html = fetch_html(area_url, session=session)
        if not html:
            warn(f"  Career area fetch failed: {area_url}")
            continue
        soup = parse_html(html)
        before = len(seen)
        for a in soup.find_all("a", href=True):
            href = a["href"]
            if "/pc/job/" in href or "/hr/job/" in href:
                if not href.startswith("http"):
                    href = "https://careers.lausd.org" + href.split("careers.lausd.org")[-1]
                    if not href.startswith("http"):
                        href = "https://careers.lausd.org" + href
                seen.add(href.rstrip("/"))
        info(f"  +{len(seen) - before} new URLs (total {len(seen)})")
    return list(seen)


def is_closed_posting_(text):
    """Return True if the page body indicates the posting is closed.

    SuccessFactors swaps the body for a closure notice and removes SALARY
    DETAILS and APPLICATION FILING DATES entirely. Detecting this lets us
    either skip the row or write it with status=closed -- never write a row
    with empty salary and no explanation.
    """
    if not text:
        return False
    lower = text.lower()
    return any(m.lower() in lower for m in CLOSED_POSTING_MARKERS)


def extract_job_id(url):
    """Extract the numeric job ID from a job URL.

    Format: https://careers.lausd.org/pc/job/{slug}/{id}/
    Returns the numeric ID as a string.
    """
    parts = url.rstrip("/").split("/")
    if len(parts) >= 2:
        return parts[-1]
    return ""


def get_job_type(url):
    """Determine if a job is Classified or Certificated based on URL path."""
    if "/pc/job/" in url:
        return "Classified"
    elif "/hr/job/" in url:
        return "Certificated"
    return "Unknown"


def extract_section_data(soup, section_title):
    """Extract text content for an H2 section on a SAP SuccessFactors job page.

    HTML structure per section:
      <div style="padding:10px...">
        <div>                           <!-- title container -->
          <h2><u><b>SECTION_TITLE</b></u></h2>
        </div>
        <div>                           <!-- content container -->
          <p>actual content text</p>
        </div>
      </div>

    We find the <h2>, navigate up to the padding div, then grab the
    last direct-child <div> (the content container).
    """
    headings = [h for h in soup.find_all(lambda t: t.name and t.name.lower() == "h2")
                if section_title.lower() in h.get_text(strip=True).lower()]
    if not headings:
        return ""

    h = headings[0]
    # Navigate: h2 -> parent (title div) -> parent (padding div)
    title_div = h.parent
    if title_div and title_div.name != "div":
        title_div = h.find_parent("div")
    padding_div = title_div.find_parent("div") if title_div else None
    if not padding_div:
        return ""

    # The content div is the last direct-child div in the padding div
    child_divs = padding_div.find_all("div", recursive=False)
    if len(child_divs) < 2:
        return ""

    content_div = child_divs[-1]
    parts = []
    for tag in content_div.find_all(True):
        text = tag.get_text(strip=True)
        if text and len(text) > 1:
            parts.append(text)

    return " ".join(parts)


def extract_salary(text):
    """Parse salary text into min, max, and schedule.

    Example input: "Updated Salary: $145,146 - $180,269 Annually"
    Returns (min, max, schedule) as strings.
    """
    if not text:
        return "", "", ""

    min_val = ""
    max_val = ""
    schedule = "Annually"

    # Find all dollar amounts
    amounts = re.findall(r'\$[\d,]+', text)
    if len(amounts) >= 2:
        min_val = amounts[0].replace("$", "").replace(",", "")
        max_val = amounts[1].replace("$", "").replace(",", "")
    elif len(amounts) == 1:
        min_val = amounts[0].replace("$", "").replace(",", "")
        max_val = min_val

    # Determine schedule (Annually/Hourly/Monthly)
    if "hourly" in text.lower() or "hour" in text.lower():
        schedule = "Hourly"
    elif "monthly" in text.lower():
        schedule = "Monthly"
    elif "annually" in text.lower() or "year" in text.lower():
        schedule = "Annually"
    elif "range" in text.lower():
        schedule = "Range"

    return min_val, max_val, schedule


def extract_filing_dates(text):
    """Parse filing dates from text.

    Example input: "Application Date: July 1, 2026 - August 15, 2026"
    Returns (start, end) as strings.
    """
    if not text:
        return "", ""

    # Look for date ranges (Month DD, YYYY - Month DD, YYYY)
    dates = re.findall(
        r'([A-Z][a-z]+ \d{1,2}, \d{4})\s*[-–]\s*([A-Z][a-z]+ \d{1,2}, \d{4})',
        text
    )
    if dates:
        return dates[0][0], dates[0][1]

    # Single date (apply by)
    single = re.findall(r'([A-Z][a-z]+ \d{1,2}, \d{4})', text)
    if single:
        return single[0], ""

    return "", ""


def extract_work_year(text):
    """Extract work year (full-time/part-time, F/T, P/T) from text."""
    if not text:
        return ""
    t = text.lower()
    if "full time" in t or "full-time" in t or "f/t" in t:
        return "Full-Time"
    elif "part time" in t or "part-time" in t or "p/t" in t:
        return "Part-Time"
    return ""


def extract_meta_field(soup, itemprop):
    """Extract a meta tag's content attribute by itemprop value.

    SAP SuccessFactors job pages use Schema.org/JobPosting meta tags:
      <meta itemprop="datePosted" content="Sun Sep 13 07:00:00 UTC 2026">
    """
    meta = soup.find("meta", attrs={"itemprop": itemprop})
    if meta and meta.get("content"):
        return meta["content"].strip()
    return ""


def parse_job_page(html, url):
    """Parse a single job page HTML and extract structured data.

    SAP SuccessFactors job pages are server-rendered HTML with:
    - <h1><span itemprop="title"> for the job title
    - <h2><u><b>SECTION_TITLE</b></u> for section headings
      (SALARY DETAILS, APPLICATION FILING DATES, DEPARTMENT, etc.)
    - <meta itemprop="..."> for schema.org fields (datePosted, validThrough, etc.)
    """
    soup = parse_html(html)

    # Closed-posting detection. SuccessFactors replaces the body with a
    # closure notice and strips SALARY DETAILS / APPLICATION FILING DATES,
    # so a naive parse returns empty salary with no explanation. We write
    # the row with status=closed instead of dropping it, so the tab stays
    # an accurate record of what was scanned.
    page_text = soup.get_text(separator=" ", strip=True)
    if is_closed_posting_(page_text):
        warn(f"  Posting is CLOSED: {url}")
        return {
            "job_id": extract_job_id(url),
            "title": "",
            "job_type": get_job_type(url),
            "department": "",
            "salary_min": "",
            "salary_max": "",
            "salary_schedule": "",
            "filing_date_start": "",
            "filing_date_end": "",
            "date_posted": "",
            "valid_through": "",
            "location": "",
            "position_summary": "",
            "duties": "",
            "requirements": "",
            "education": "",
            "experience": "",
            "work_year": "",
            "status": "closed",
            "url": url,
        }

    job_id = extract_job_id(url)
    job_type = get_job_type(url)

    # Title — inside <h1><span itemprop="title"> or just <h1>
    title_el = soup.find("span", attrs={"itemprop": "title"})
    if title_el:
        title = title_el.get_text(strip=True)
    else:
        h1 = soup.find("h1")
        title = h1.get_text(strip=True) if h1 else ""

    # Schema.org meta fields
    date_posted = extract_meta_field(soup, "datePosted") or ""
    valid_through = extract_meta_field(soup, "validThrough") or ""

    # Department
    dept_text = extract_section_data(soup, "DEPARTMENT")
    department = dept_text.strip() if dept_text else ""

    # Salary details
    salary_text = extract_section_data(soup, "SALARY DETAILS")
    salary_min, salary_max, salary_schedule = extract_salary(salary_text)

    # Filing dates — try section first, fall back to schema dates
    dates_text = extract_section_data(soup, "APPLICATION FILING DATES")
    filing_start, filing_end = extract_filing_dates(dates_text)
    if not filing_start and date_posted:
        filing_start = date_posted
    if not filing_end and valid_through:
        filing_end = valid_through

    # Work year / employment type
    work_year = extract_work_year(soup.get_text())

    # Location — look for job-location paragraph or schema
    location = ""
    loc_el = soup.find("p", class_="job-location")
    if loc_el:
        loc_text = loc_el.get_text(strip=True)
        if loc_text and loc_text != "US":
            location = loc_text
    if not location:
        loc_meta = soup.find("meta", attrs={"itemprop": "jobLocation"})
        if loc_meta and loc_meta.get("content"):
            location = loc_meta["content"]

    # Position summary
    position_text = extract_section_data(soup, "THE POSITION")
    position_summary = position_text.strip() if position_text else ""

    # Duties
    duties = extract_section_data(soup, "JOB DUTIES/RESPONSIBILITIES")

    # Requirements — Minimum Requirements section
    req_text = extract_section_data(soup, "MINIMUM REQUIREMENTS")

    # Try to split requirements into education/experience subsections
    education = ""
    experience = ""
    if req_text:
        edu_match = re.search(r'EDUCATION\s*[:\s]*(.+?)(?:EXPERIENCE\s*[:\s]|Special:|$)', req_text, re.IGNORECASE | re.DOTALL)
        exp_match = re.search(r'EXPERIENCE\s*[:\s]*(.+?)(?:EDUCATION\s*[:\s]|Special:|$)', req_text, re.IGNORECASE | re.DOTALL)
        if edu_match:
            education = edu_match.group(1).strip()[:500]
        if exp_match:
            experience = exp_match.group(1).strip()[:500]
        if not education and not experience:
            education = req_text[:500]

    return {
        "job_id": job_id,
        "title": title,
        "job_type": job_type,
        "department": department[:200] if department else "",
        "salary_min": salary_min,
        "salary_max": salary_max,
        "salary_schedule": salary_schedule,
        "filing_date_start": filing_start,
        "filing_date_end": filing_end,
        "date_posted": date_posted,
        "valid_through": valid_through,
        "location": location[:200] if location else "",
        "position_summary": position_summary[:1000] if position_summary else "",
        "duties": duties[:2000] if duties else "",
        "requirements": req_text[:2000] if req_text else "",
        "education": education,
        "experience": experience,
        "work_year": work_year,
        "status": "open",
        "url": url,
    }


def scrape_jobs(max_jobs=None):
    """Main entry point: discover and scrape currently-OPEN job postings.

    Discovers postings from the 12 career-area listing pages rather than the
    sitemap, so closed postings (whose bodies are replaced with a closure
    notice and which have no SALARY DETAILS / APPLICATION FILING DATES) are
    detected and written with status=closed instead of accumulating as
    hundreds of empty-salary rows.

    Parameters
    ----------
    max_jobs : int or None — limit for testing; None = all

    Returns
    -------
    list[dict] — one dict per job/closed posting
    """
    session = make_session()
    urls = discover_open_postings_(session)
    if not urls:
        error("No job URLs found on career-area pages")
        return []

    if max_jobs:
        urls = urls[:max_jobs]

    jobs = []
    for i, url in enumerate(urls):
        info(f"[{i+1}/{len(urls)}] Scraping: {url}")
        html = fetch_html(url, session=session, timeout=30)
        if not html:
            warn(f"  Skipping (fetch failed)")
            continue

        job = parse_job_page(html, url)
        if not job["title"] and job["status"] != "closed":
            warn(f"  Skipping (no title found)")
            continue

        jobs.append(job)
        info(f"  -> {job.get('title') or '(closed)'} ({job['job_id']}) [{job['job_type']}] status={job['status']}")

    # Deduplicate by job_id
    jobs = dedupe_rows(jobs, key="job_id")
    open_count = sum(1 for j in jobs if j.get("status") == "open")
    closed_count = sum(1 for j in jobs if j.get("status") == "closed")
    info(f"Total jobs scraped: {len(jobs)} ({open_count} open, {closed_count} closed)")

    # Add provenance
    jobs = add_provenance(jobs, CAREER_AREA_PAGES[0], "careers.lausd.org career-area pages + individual job pages")

    return jobs


def main():
    import argparse
    parser = argparse.ArgumentParser(description="Scrape LAUSD careers jobs")
    parser.add_argument("--max", type=int, default=None, help="Max jobs to scrape (for testing)")
    parser.add_argument("--json", action="store_true", help="Also write JSON output")
    args = parser.parse_args()

    jobs = scrape_jobs(max_jobs=args.max)

    if args.json:
        path = write_json(jobs, "jobs.json")

    # Write to Excel
    from _utils import write_tab, ensure_output_dir
    path = write_tab(jobs, "Jobs", JOBS_HEADERS)
    info(f"Wrote {len(jobs)} jobs to {path}")

    return jobs


if __name__ == "__main__":
    main()
