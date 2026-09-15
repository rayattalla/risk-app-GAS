"""Schools scraper — California Department of Education, Public Schools and Districts

Source: https://www.cde.ca.gov/ds/si/ds/pubschls.asp
  -> the "Public Schools and Districts Data Files" downloadable ZIP contains
     a flat file (typically SchoolDirectory.txt or Schools.txt) with one row
     per California public school.

This one fetch covers TWO tabs at once:
  Schools     — CDS code, county/district/school codes, name, address, phone,
                grade span, county/district/school type, latitude/longitude
  Principals  — school name + administrator name and title ONLY.

PRIVACY: the CDE file also carries administrator EMAIL fields. Those are NOT
stored. Names and titles only — the district-data skill has no use for an
email address, and storing one would be an unnecessary privacy exposure.

The file is a fixed-width or pipe-delimited flat file, not HTML, so there is
no BeautifulSoup step and no Playwright. Plain requests + a delimiter sniff.
"""
import os
import sys
import re
import zipfile
import io
import csv

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _utils import (
    make_session, fetch_html, add_provenance, write_json,
    write_tab, info, warn, error, now_str
)

CDE_DIR_URL = "https://www.cde.ca.gov/ds/si/ds/pubschls.asp"

# Column definitions for the Schools tab (subset of the CDE directory fields
# the district-data skill actually queries on). Order matters: it is the
# header order written to the Excel tab.
SCHOOLS_HEADERS = [
    "cds_code",
    "county_code",
    "district_code",
    "school_code",
    "school_name",
    "district_name",
    "school_type",
    "county_name",
    "address",
    "city",
    "state",
    "zip",
    "phone",
    "grades_span",
    "county_district_school_type",
    "latitude",
    "longitude",
]

# Principals tab: school name + administrator name and title ONLY. No email.
PRINCIPALS_HEADERS = [
    "cds_code",
    "school_name",
    "principal_name",
    "principal_title",
]

# CDE flat files are pipe-delimited in recent years; older ones use a fixed
# width. Sniff the first non-empty line and pick a reader.
DELIMITERS = ["|", "\t", ",", ";"]

# CDS codes are 14 characters: 2 (county) + 4 (district) + 2 (school) + 6
# (subtype/reserved). The first 6 chars identify the district; the first 8
# identify the school. Used to join the Schools and Principals tabs.
CDS_LEN = 14


def find_zip_url(session):
    """Locate the latest Public Schools and Districts downloadable file.

    The pubschls.asp page links each year's data set. We prefer the most
    recent ZIP that contains a directory flat file; fall back to the most
    recent direct file link if no ZIP is present.
    """
    html = fetch_html(CDE_DIR_URL, session=session)
    if not html:
        error("Could not fetch CDE directory page")
        return None, ""

    # Look for .zip links first (the downloadable data file set)
    zip_matches = re.findall('href=["\\\']([^"\\\']+\\.zip[^"\\\']*)["\\\']', html, re.IGNORECASE)
    if zip_matches:
        url = zip_matches[0]
        if not url.startswith("http"):
            url = "https://www.cde.ca.gov" + url
        info("Found CDE ZIP: " + url[:90])
        return url, "zip"

    # Fallback: any file link
    file_matches = re.findall('href=["\\\']([^"\\\']+\\.(?:txt|csv|xls[xm]?|zip)[^"\\\']*)["\\\']', html, re.IGNORECASE)
    if file_matches:
        url = file_matches[0]
        if not url.startswith("http"):
            url = "https://www.cde.ca.gov" + url
        info("Found CDE file: " + url[:90])
        return url, "file"

    warn("No downloadable file found on CDE directory page")
    return None, ""


def sniff_delimiter(sample):
    """Pick the delimiter with the highest count in the sample line."""
    best = ","
    best_count = 0
    for d in DELIMITERS:
        c = sample.count(d)
        if c > best_count:
            best_count = c
            best = d
    return best


def parse_flat_text(text):
    """Parse a CDE directory flat file into a list of row dicts.

    The first non-empty line is the header. Subsequent lines are data.
    """
    lines = [ln for ln in text.splitlines() if ln.strip()]
    if not lines:
        return [], []

    header_line = lines[0]
    delim = sniff_delimiter(header_line)
    headers = [h.strip().lower().replace(" ", "_") for h in header_line.split(delim)]

    rows = []
    for ln in lines[1:]:
        cells = ln.split(delim)
        row = {}
        for i, h in enumerate(headers):
            row[h] = cells[i].strip() if i < len(cells) else ""
        rows.append(row)
    return headers, rows


def _get_cell(row, *names):
    """Return the first non-empty value among the given header names.

    CDE column names vary between years and between the directory file and
    the personnel file, so lookups are name-based rather than positional.
    """
    for n in names:
        v = row.get(n)
        if v is not None and str(v).strip():
            return str(v).strip()
    return ""


def _cds_to_codes(cds):
    """Split a 14-char CDS code into county/district/school components."""
    cds = (cds or "").strip()
    return {
        "cds_code": cds,
        "county_code": cds[:2],
        "district_code": cds[2:6],
        "school_code": cds[6:8],
    }


def _build_schools_rows(rows):
    """Map raw CDE directory rows onto SCHOOLS_HEADERS."""
    out = []
    for r in rows:
        cds = _get_cell(r, "cds", "cds_code", "cdscode", "county_district_school")
        codes = _cds_to_codes(cds)
        out.append({
            "cds_code": codes["cds_code"],
            "county_code": codes["county_code"],
            "district_code": codes["district_code"],
            "school_code": codes["school_code"],
            "school_name": _get_cell(r, "school_name", "school", "school_name_and"),
            "district_name": _get_cell(r, "district_name", "district"),
            "school_type": _get_cell(r, "school_type", "type", "school_type_code"),
            "county_name": _get_cell(r, "county_name", "county"),
            "address": _get_cell(r, "address", "street", "street_address"),
            "city": _get_cell(r, "city"),
            "state": _get_cell(r, "state"),
            "zip": _get_cell(r, "zip", "zip_code", "postal_code"),
            "phone": _get_cell(r, "phone", "telephone", "phone_number"),
            "grades_span": _get_cell(r, "grades_span", "grades", "grade_span",
                                     "lowest_grade", "highest_grade"),
            "county_district_school_type": _get_cell(r, "county_district_school_type",
                                                     "county_district_school",
                                                     "cds_type"),
            "latitude": _get_cell(r, "latitude", "lat"),
            "longitude": _get_cell(r, "longitude", "lon", "lng"),
        })
    return out


def _build_principals_rows(rows):
    """Map raw CDE directory rows onto PRINCIPALS_HEADERS.

    ONLY the name and title are kept. The CDE file also carries
    administrator EMAIL fields; those are deliberately ignored here.
    """
    out = []
    for r in rows:
        cds = _get_cell(r, "cds", "cds_code", "cdscode", "county_district_school")
        name = _get_cell(r, "principal_name", "administrator_name",
                         "school_administrator_name", "principal")
        title = _get_cell(r, "principal_title", "administrator_title",
                          "school_administrator_title", "position_title",
                          "job_title")
        school = _get_cell(r, "school_name", "school")
        if not name:
            continue
        out.append({
            "cds_code": cds,
            "school_name": school,
            "principal_name": name,
            "principal_title": title,
        })
    return out


def _fetch_flat_file(url, kind, session):
    """Download the CDE data file and return its parsed rows."""
    if kind == "zip":
        resp = session.get(url, timeout=120)
        resp.raise_for_status()
        zf = zipfile.ZipFile(io.BytesIO(resp.content))
        names = [n for n in zf.namelist()
                 if not n.endswith("/") and re.search(r"(school|directory|schools)", n, re.I)]
        if not names:
            names = [n for n in zf.namelist() if not n.endswith("/")]
        if not names:
            error("CDE ZIP contains no readable file")
            return [], []
        target = names[0]
        info("Reading " + target + " from CDE ZIP")
        text = zf.read(target).decode("utf-8", errors="replace")
        return parse_flat_text(text)
    else:
        resp = session.get(url, timeout=120)
        resp.raise_for_status()
        return parse_flat_text(resp.text)


def scrape_schools(max_rows=None):
    """Main entry point: download the CDE directory file and write two tabs.

    Returns (schools_rows, principals_rows).
    """
    session = make_session()
    url, kind = find_zip_url(session)
    if not url:
        error("Could not locate CDE directory file")
        return [], []

    headers, rows = _fetch_flat_file(url, kind, session)
    info("Parsed " + str(len(rows)) + " rows from CDE directory file "
         "(" + str(len(headers)) + " columns)")

    schools = _build_schools_rows(rows)
    principals = _build_principals_rows(rows)

    if max_rows:
        schools = schools[:max_rows]
        principals = principals[:max_rows]

    schools = add_provenance(schools, CDE_DIR_URL,
                             "CDE Public Schools and Districts (" + kind + ")")
    principals = add_provenance(principals, CDE_DIR_URL,
                                "CDE Public Schools and Districts (" + kind +
                                ") — names and titles only, emails excluded")

    info("Built " + str(len(schools)) + " school rows, " +
         str(len(principals)) + " principal rows")
    return schools, principals


def main():
    import argparse
    parser = argparse.ArgumentParser(description="Scrape CDE school directory")
    parser.add_argument("--max", type=int, default=None,
                        help="Max rows per tab (for testing)")
    parser.add_argument("--json", action="store_true",
                        help="Also write JSON output")
    args = parser.parse_args()

    schools, principals = scrape_schools(max_rows=args.max)

    if args.json:
        write_json(schools, "schools.json")
        write_json(principals, "principals.json")

    path = write_tab(schools, "Schools", SCHOOLS_HEADERS)
    info("Wrote " + str(len(schools)) + " schools to " + path)
    path = write_tab(principals, "Principals", PRINCIPALS_HEADERS)
    info("Wrote " + str(len(principals)) + " principals to " + path)

    return schools, principals


if __name__ == "__main__":
    main()
