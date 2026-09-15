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
