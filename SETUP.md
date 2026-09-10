Always run the security checklist in `SECURITY.md` before the first clasp push.

**After setup:** Run `PROG → Setup & Config → Ship Checklist` and register this project (with its Script ID) in the App Registry.

# LAUSD GAS Starter — Setup Guide

Two GAS projects per new app:
1. **LAUSDLib** — the shared library (create once, reuse forever)
2. **Shell** — a copy of `template/` bound to your spreadsheet (one per app)

---

## Step 0 — Clone this repo

```
git clone <this repo> my-new-project
cd my-new-project
```

---

## Step 1 — Create the LAUSDLib project (first time only)

If you already deployed LAUSDLib for a previous project, skip to Step 2 and reuse the same script ID.

1. Go to [script.google.com](https://script.google.com) → **New project** → name it `LAUSDLib`
2. Copy all `lib/*.js` and `lib/appsscript.json` into that project via CLASP or the editor
3. In the project settings, copy the **Script ID**
4. Fill your script ID into `lib/.clasp.json`:
   ```json
   { "scriptId": "1ABC...your-lib-id...", "rootDir": "." }
   ```
5. Push with CLASP:
   ```
   cd lib
   clasp push
   ```
6. In the GAS editor → **Deploy** → **New deployment** → type = **Library** → deploy and copy the version number (or use `developmentMode: true` to always use HEAD during dev)

---

## Step 2 — Fill in lib/01_config.js

Open `lib/01_config.js` and replace every `YOUR_*` placeholder:

| Placeholder | What to put |
|---|---|
| `YOUR_SPREADSHEET_ID` | The ID from your backend Google Sheet URL |
| `YOUR_PROGRAM_NAME` | e.g. `Architecture Review Board` |
| `PROG` | Short code for menus, e.g. `ARB` |
| `your.name@lausd.net` | Your own email (always gets Admin role) |
| `your-team@lausd.net` | Team inbox CC'd on every notification |

Then push to the lib project:
```
cd lib && clasp push
```

---

## Step 3 — Create the Shell project

1. Create a **new** Google Sheet for this project's data
2. From that sheet → **Extensions → Apps Script** (this creates the bound script)
3. In the bound script's project settings, copy the **Script ID**
4. Fill your script ID into `template/.clasp.json`:
   ```json
   { "scriptId": "1DEF...your-shell-id...", "rootDir": "." }
   ```

---

## Step 4 — Wire up LAUSDLib in the shell

1. Open `template/appsscript.json`
2. Replace `YOUR_LAUSDLIB_SCRIPT_ID` with the script ID from Step 1
3. Open `template/config.js` and replace every `YOUR_*` placeholder (same values as Step 2)

Push the shell:
```
cd template && clasp push
```

---

## Step 5 — Run Setup

1. Open the bound spreadsheet
2. Menu → **PROG** → **Setup & Config** → **Run Setup**

This creates all standard sheet tabs (Master, Access_Control, Audit_Log, Config).

---

## Step 6 — Deploy as Web App

1. In the shell GAS editor → **Deploy** → **New deployment**
2. Type = **Web app**
3. Execute as = **User accessing the web app**
4. Who has access = **Anyone in LAUSD** (or your domain)
5. Copy the `/exec` URL
6. Back in the spreadsheet → Menu → **PROG** → **Setup & Config** → **Set Web App URL** → paste the URL

---

## Step 7 — Enable test mode (recommended during dev)

In the shell GAS editor, run:
```javascript
enableTestMode()   // or use the menu: PROG → Test Mode → Toggle Test Mode
```

All outgoing emails will redirect to `CONFIG.ADMIN_EMAIL` until you run `disableTestMode()`.

---

## Multi-environment deployment (dev vs prod)

You have two LAUSD tenants: `lausd.net` (staging/dev) and `lausd.io` (production). Keep separate `.clasp.json` files for each:

```
template/
  .clasp-dev.json    ← lausd.net spreadsheet + bound script ID
  .clasp-prod.json   ← lausd.io spreadsheet + bound script ID
  .clasp.json        ← gitignored; symlink or copy from above before pushing
```

Push to dev:
```powershell
Copy-Item .clasp-dev.json .clasp.json; clasp push
```

Push to prod:
```powershell
Copy-Item .clasp-prod.json .clasp.json; clasp push
```

Do the same for `lib/` if the library has separate dev/prod deployments.

**In `lib/01_config.js`:** keep `BACKEND_SHEET_ID` and `ADMIN_EMAIL` for the current environment. Swap values when switching. Or use the Config sheet (§6 in PATTERNS.md) to override at runtime so the library code itself doesn't need to change between environments.

---

## Pulling library updates into existing projects

When you improve LAUSDLib (bug fix, new function, new module), existing projects pick up the change automatically if they use `developmentMode: true` in their `appsscript.json`. No action needed.

If you locked a project to a specific library version number (for production stability), update it manually:
1. In GAS editor → Project Settings → Libraries → change version number to latest
2. Or update `"version"` in the project's `appsscript.json` and `clasp push`

**The `// keep-in-sync` constants in `template/config.js`** are the only thing that requires manual action — they don't auto-update. When you change identity values in `lib/01_config.js`, search all your shell projects for `keep-in-sync` and update those files too.

---

## Starting an AI session on a new LAUSD GAS project

Paste these two files at the start of the session so the AI has full context:
1. `PATTERNS.md` — the 13 design decisions and why they exist
2. `SETUP.md` — the file map and workflow

Without these, the AI will re-derive the patterns from scratch (or worse, violate them silently).

---

## Adding a new project (after first-time setup)

You don't need to recreate the library. Just:

1. Copy the `template/` folder
2. Create a new bound script (new spreadsheet → Extensions → Apps Script)
3. Fill in `config.js` and `appsscript.json` with the new project's IDs
4. Push and run Setup

---

## File map

```
lausd-gas-starter/
├── PATTERNS.md              ← 13 design patterns + rationale (paste into AI sessions)
├── SETUP.md                 ← this file
├── lib/                     ← LAUSDLib (one shared library project)
│   ├── 01_config.js         ← CONFIG object + test mode toggles
│   ├── 02_helpers.js        ← Sheet I/O, ID lookup, ID gen, audit log, Drive, Config sheet,
│   │                           Extra Fields JSON, column protection, changelog, getDashboard()
│   ├── 03_access.js         ← RBAC: roles, capabilities, Access_Control sheet
│   ├── 04_email.js          ← LAUSD-branded email layout + send wrapper
│   ├── 05_cache.js          ← CacheService chunking + LockService + logo caching
│   ├── 06_notify.js         ← Teams webhook, SMS via email-to-carrier, token templating,
│   │                           external webhook (SIEM/SOAR), Notification_Templates sheet
│   ├── 07_reports.js        ← Google Docs template-first generation, heading navigation,
│   │                           token substitution, update history, simple code fallback
│   ├── 08_ai.js             ← Multi-provider AI (OpenRouter + Gemini + HuggingFace),
│   │                           circuit breaker, local fallback, callAI() / getAiStatus()
│   ├── 09_scaffolder.js     ← Project scaffolding: Drive folders, central registry,
│   │                           slug generation, scaffoldProject(), Notion stub
│   └── appsscript.json      ← Library manifest
└── template/                ← Shell (copy this for each new project)
    ├── config.js            ← Shell-side constants (keep-in-sync with lib/01_config.js)
    ├── shell.js             ← doGet (inlined), onOpen, thin wrappers → LAUSDLib
    ├── index.html           ← LAUSD-branded SPA starting point
    └── appsscript.json      ← Shell manifest with LAUSDLib dependency
```

---

## Key patterns and why they exist

### Dual-script (shell + library)
The shell is bound to the spreadsheet and owns the web app URL. The library holds all logic. This means:
- Library updates deploy instantly (no re-deploy of the web app needed)
- Multiple projects can share the same library

### `doGet` is inlined in the shell
Deployed web apps snapshot the library at deploy time. Functions called from `doGet` before the page renders (`getMyAccess`, config data) use potentially stale library code. Inlining them in the shell avoids stale data.

### `keep-in-sync` comments
Constants in `template/config.js` that mirror `lib/01_config.js` are tagged `// keep-in-sync`. Update both places when changing identity values.

### Header-driven writes (`_appendRowByHeaders_`, `_updateRowByHeaders_`)
Writes are keyed by column header name, not position. Adding or reordering columns in the sheet never breaks data writes.

### Two-pass ID lookup
`findRowById_` reads display values first (what the user sees), then falls back to raw values. This avoids the Sheets auto-coercion bug where numeric IDs like `6216` appear as `6216.0` in raw values.

### Extra Fields JSON column (`parseExtraFields_`, `setExtraField_`)
Spreadsheet schemas are hard to migrate. Add an `Extra_Fields` column (type: plain text) to any sheet and store arbitrary key-value pairs there as JSON. New fields never require adding columns. Pattern from IRS v3.4.

### Column protection (`protectColumn_`)
Lock a column to admin-only editing using the Sheets Protection API. Idempotent — removes any existing protection first. Useful for Status, Risk Tier, or any field only admins should change.

### Changelog (`logChangelog_`, `_ensureChangelogSheet_`)
Install an `onEdit(e)` trigger in your shell and call `logChangelog_(recordId, field, oldValue, newValue)` to maintain a timestamped change history. Skips no-ops and ignores the `LastUpdated` field to prevent loops.

### Teams webhook + SMS gateway + external webhook (`06_notify.js`)
Three channels beyond email: post cards to Microsoft Teams, send SMS via carrier email gateways (no external API), and POST JSON to SIEM/SOAR/ticketing systems. All respect `isTestMode()`. Token templating (`renderTemplate`) and a sheet-driven template system (`getNotificationTemplate`) let admins customize messages without code changes.

### Google Docs report generation (`07_reports.js`)
Two patterns: (1) Template-first — copy a Drive Doc template, fill `{TOKEN}` placeholders, save to project folder. Falls back to a code-generated Doc if the template ID is missing. (2) Heading-based updates — locate a named heading in an existing Doc and update the paragraph, table cell, or history section after it without rewriting the whole document.

### CacheService chunking
CacheService has a 100 KB per-key limit. `cacheGetOrFetch` in `05_cache.js` splits large JSON payloads into 90 KB chunks and reassembles them on read. `LockService` prevents thundering herd on cache miss.
