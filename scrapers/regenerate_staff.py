"""regenerate_staff.py — STUB. DO NOT RUN.

The live Sheet's Staff tab has 994 rows. This repo cannot reproduce that
number: the only staff scraper here (scrapers/staff.py) pulls three summary
counts from the LAUSD Fingertip Facts page and writes 3 rows.

That is a liability. A tab that nobody can regenerate means:
  - a data-corruption incident is unrecoverable,
  - nobody can tell whether a figure is current or stale,
  - the district-data agent will confidently cite numbers with no source.

Until a reproducible source exists, the district-data agent MUST NOT answer
questions against the Staff tab. See README.md "Unregeneratable tabs".

If you have the internal source (MiSiS, HR/Payroll, or an exported CSV),
paste it here and point STAFF_SOURCE at it. Replace scrape_staff() below.
"""
import os
import sys
import json

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _utils import info, warn, error

# Set this to a local CSV/JSON/Excel path once a reproducible source exists.
STAFF_SOURCE = os.environ.get("STAFF_SOURCE_PATH", "")

# Schema the district-data skill expects on the Staff tab.
# NOTE: source_url appears once here. _utils.write_tab() dedupes its
# PROVENANCE_COLS against declared headers, so it will not double it --
# that dedupe is what fixed the original duplicate-column shift bug.
EXPECTED_HEADERS = [
    "staff_id",
    "name",
    "confidence",
    "source_url",
    "scrape_date",
    "scrape_note",
]


def load_source():
    if not STAFF_SOURCE:
        error("STAFF_SOURCE_PATH is not set. See the module docstring.")
        return []
    if not os.path.exists(STAFF_SOURCE):
        error(f"STAFF_SOURCE_PATH does not exist: {STAFF_SOURCE}")
        return []
    if STAFF_SOURCE.endswith(".json"):
        with open(STAFF_SOURCE, encoding="utf-8") as f:
            return json.load(f)
    raise NotImplementedError("Only JSON sources are supported until a CSV loader is added.")


def main():
    rows = load_source()
    info(f"Loaded {len(rows)} staff rows from {STAFF_SOURCE or '(unset)'}")
    # TODO: validate against EXPECTED_HEADERS, then write_tab(rows, "Staff", EXPECTED_HEADERS)


if __name__ == "__main__":
    main()