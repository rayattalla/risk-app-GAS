# RiskAI — clasp project root (authoritative)

This directory is the **authoritative** source for the live Apps Script project
`1lk1SNuDFZoDTRfJZB0jrT1R22HnOW59uCNuJaTecnvuyT_fyAsNMsWQn`.

## Why this directory, and why this matters

The repo root contains a sibling `lib/` folder whose `.clasp.json` points at the
**same script ID**. Two `.clasp.json` files targeting one script ID is how
"which code is live" confusion happens — it has already cost this project a
week. This file is the permanent record of the resolution.

### Which one is live

**`template/` is live.** Evidence (filesystem + git, 2026-09-15):

- `git log` shows every district-data / skill / chat change since the initial
  commit touched `template/`, never `lib/`.
- `lib/` was committed once, at `1ef7b5e` ("Initial commit: GAS Agent Platform"),
  and has **never been modified since** — it is the frozen initial snapshot.
- Commit `ff91fde` ("redeploy HEAD to live web app") modified `template/20_chat.js`
  and `template/shell.js` and its message says explicitly: "redeployed as @21".
- `lib/20_chat.js` has no `district-data`, `cve-lookup`, `skill_refs`, or `diagram`
  logic. `template/20_chat.js` has all four. The live script has district-data,
  so `template/` is what is deployed.

### What `lib/` is

`lib/` is the **retired initial snapshot**. It is kept in git for history only
and must never be pushed again. It predates the skills migration, the persona
scope-lock, the plain-text formatting rule, and every district-data change.

## Permanent rules

1. **Only one `.clasp.json` may target this script ID.** If a second one ever
   appears, stop and resolve it the same way this was resolved.
2. **`template/` is what `clasp push` reads.** Never `cd lib && clasp push`.
3. **`lib/` is retired.** Do not edit it. If a change is ever needed that is
   only in `lib/`, it is by definition stale — port the change into `template/`
   instead.
4. **To confirm the live code at any time**, open the Apps Script editor for
   `1lk1SNu…` as the LAUSD account and look at the file list. If
   `17_district_data.js` is there, `template/` is live. If it is not, `lib/` is
   live and district-data has never run in production.

## Version

`template/config.js` holds `SHELL_VERSION`. Bump it on every change and say so
in the commit message.

## Unregeneratable tabs — resolved for Schools

The live Sheet's **Schools** tab is now reproducible: `scrapers/schools.py`
downloads the CDE Public Schools and Districts file from
`https://www.cde.ca.gov/ds/si/ds/pubschls.asp` and writes the **Schools** tab
(CDS code, county/district/school codes, name, address, phone, grade span,
county/district/school type, latitude/longitude).

The same fetch also answers the principal-names question: the CDE directory
file includes school administrator name fields, publicly published by the
state. Those are written to a **Principals** tab (cds_code, school_name,
principal_name, principal_title).

**Privacy:** the CDE file also carries administrator EMAIL fields. They are
deliberately not stored — `_build_principals_rows()` only reads name and
title. Names and titles only.

## Unregeneratable tabs — still a liability

The live Sheet's **Staff** tab has 994 rows. This repo cannot reproduce that
number: the only staff scraper (`scrapers/staff.py`) pulls three summary
counts from the LAUSD Fingertip Facts page and writes 3 rows. The 994-row
version has no committed source and no committed scraper.

A tab nobody can regenerate is a liability:
- a corruption incident is unrecoverable,
- nobody can tell whether a figure is current or stale,
- the district-data agent will confidently cite numbers with no source.

**Until a reproducible source exists, the district-data agent MUST NOT answer
questions against the Staff tab.** `scrapers/regenerate_staff.py` is a stub
that documents this and will load a source once you supply one (set
`STAFF_SOURCE_PATH` to a local JSON/CSV/Excel path).

## Scrapers

| Tab | Scraper | Status |
| --- | --- | --- |
| Schools + Principals | `scrapers/schools.py` | New. One CDE fetch covers both. Names/titles only — emails excluded. |
| Jobs | `scrapers/jobs.py` | Fixed: discovers open postings from 12 career-area pages, detects closed-posting body, writes `status=open|closed`. No Playwright. |
| Classifications | `scrapers/salary_schedule.py` | Fixed: pdfplumber `extract_table()` preserves the ruled grid that pypdf flattens. Falls back to pypdf if pdfplumber is missing. |
| Staff | `scrapers/regenerate_staff.py` | **Stub** — no reproducible source. See above. |
| Enrollment / Budget | none in repo | CDE public downloads; no scraper committed. |

`scrapers/_utils.py` `write_tab()` now dedupes `PROVENANCE_COLS` against the
declared headers before concatenating. Without this, a scraper that already
lists `source_url` in its own headers gets it twice — the same class of bug
as the original Staff column-shift bug.
