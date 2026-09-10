# LAUSDLib — Design Patterns Reference

This is the living "§N" patterns document referenced in code comments throughout the library and consuming projects. Paste it at the start of any AI session working on a LAUSD GAS project so the assistant understands the non-obvious decisions without re-deriving them.

---

## §1 — Dual-Script Architecture (Shell + Library)

**Rule:** Every GAS web app has exactly two projects: a *library* that holds all business logic, and a *bound shell* that owns `onOpen`, `doGet`, and thin wrappers.

**Why:** The shell is bound to the spreadsheet, which means it owns the stable web app URL and the custom menu. The library is a separate project that can be updated independently. When you fix a bug in the library, consuming projects pick it up immediately (in `developmentMode`) without re-deploying their web app.

**How to apply:** Put ALL business logic in the library. The shell contains only:
- `onOpen()` — must run before library loads, cannot be in library
- `doGet()` — must be inlined (see §3)
- `// keep-in-sync` mirror constants (see §2)
- One-liner wrapper functions that forward to `LAUSDLib.*`

---

## §2 — Keep-in-Sync Mirror Constants

**Rule:** Any constant used by an inlined function in the shell (`doGet`, `_denyPage_`, `onOpen`) must be declared at the top of `config.js` in the shell, tagged `// keep-in-sync`.

**Why:** Inlined functions cannot reference the library's `CONFIG` object (the library isn't loaded yet when these run). If you forget to update the mirror constant, the live app serves stale data silently — no error is thrown.

**How to apply:** When you change `PROGRAM_NAME`, `ADMIN_EMAIL`, `NOTIFY_EMAIL`, or `TIMEZONE` in `lib/01_config.js`, immediately update the same values in `template/config.js`. The `// keep-in-sync` comment is the flag.

---

## §3 — Inline Hot Paths to Avoid Stale Library Snapshots

**Rule:** `doGet()`, `getMyAccess()`, and `getConfigData()` must be inlined in the shell, not delegated to the library.

**Why:** When you deploy a GAS web app, it takes a *snapshot* of all library code at that moment. Even if you update the library later, the deployed web app continues running the old snapshot. Functions that run on the initial page load — before the user sees anything — must always reflect the current library state. Inlining them in the shell (which runs "live" at each request) bypasses the snapshot.

**Symptom of violation:** New roles you added don't take effect. A user who should be blocked gets through. Dropdown options are outdated. These bugs only appear in the *deployed* web app, not in test runs.

**How to apply:** If a function must work correctly the moment a user hits the web app URL, inline it in the shell. Add a comment pointing to the library version: `// Inlined from LAUSDLib (see lib/03_access.js:getMyAccess) — do not delegate`.

---

## §4 — Two-Pass ID Lookup (Display Values First)

**Rule:** When searching for a record by ID in a sheet, always try `getDisplayValues()` first, then fall back to `getValues()` with a `Date` normalization step.

**Why:** Google Sheets auto-coerces values. A numeric ID typed as `6216` gets stored as the number `6216`, which `getValues()` returns as `6216.0` in some contexts. But `getDisplayValues()` returns what the user sees in the cell — `"6216"` — which matches the string the user passed in. The `.0` suffix is invisible to users but breaks exact string matching.

**Secondary issue:** Legacy Quickbase IDs were sometimes stored as dates (Sheets interpreted them as serial numbers). The second pass handles that by formatting the raw value as a date string and comparing again.

**How to apply:** Always use `findRowById_()` from `02_helpers.js`. Never write your own `indexOf` against raw sheet values for ID columns.

---

## §5 — Header-Driven Writes (Never Write by Column Index)

**Rule:** When appending or updating a row in a sheet, always build the row by mapping header names to values — never by column position.

**Why:** Users add columns, move columns, or the form schema changes. Any code that writes `row[3] = value` breaks the moment column 3 is no longer the right field. Header-driven writes read the live header row and align data to it, so they survive schema evolution silently.

**How to apply:** Use `_appendRowByHeaders_(sheet, {ColumnName: value})` and `_updateRowByHeaders_(sheet, rowIdx, {ColumnName: value})` from `02_helpers.js`. Never do `sheet.getRange(row, 5).setValue(x)` unless the column is guaranteed immutable (e.g., a system ID column).

---

## §6 — Config Sheet as Runtime Key-Value Store

**Rule:** Any value that needs to change without a code redeploy goes in the `Config` sheet (Key | Value | Notes). Code reads it via `_getConfigValue_(key)`.

**Why:** Users need to change things like the web app URL, report template IDs, or the root Drive folder ID after initial setup. Putting these in code requires a code push. The Config sheet is readable and editable by admins without touching Apps Script.

**Common Config sheet keys:**
- `WEB_APP_URL` — set after deployment so internal links work
- `ROOT_FOLDER_ID` — auto-set by `setup()`, override to point to a specific folder
- Template Doc IDs (e.g., `CSIRT_TEMPLATE_DOC_ID`, `REPORT_TEMPLATE_DOC_ID`)
- `TEAMS_WEBHOOK_URL` — set by admin without code change
- `SIEM_WEBHOOK_URL` — same

**How to apply:** Anything that might differ between dev and prod, or that an admin might need to change, belongs in Config sheet, not hardcoded in `01_config.js`. Reserve `CONFIG` constants for values that only change with a library update (org name, role definitions, risk tiers).

---

## §7 — Test Mode: PropertiesService Flag

**Rule:** Before any email, SMS, Teams message, or external webhook fires, check `isTestMode()`. In test mode, redirect to the admin email or log instead of sending.

**Why:** During development you run submissions, form triggers, and approval workflows repeatedly. Without test mode, every run sends real emails to real recipients (and real SIEM events to real SOAR systems). You'll anger your stakeholders.

**How to apply:** Call `enableTestMode()` once from the Apps Script editor at the start of every dev session. All `sendMail`, `sendSms`, `sendTeamsWebhook`, and `sendExternalWebhook` functions check `isTestMode()` internally — you don't need to add guards yourself. Call `disableTestMode()` before handing off to users.

---

## §8 — Two-Layer Access Control

**Rule:** Authorization has exactly two layers: (1) domain gate set in the GAS manifest (`access: DOMAIN`) and (2) sheet-driven role rows in `Access_Control` tab. Never gate on a single check.

**Why:** The manifest gate rejects anyone not in `lausd.net` before the script even runs (saves quota). The sheet layer gives fine-grained control within the domain. This means `@lausd.net` users default to `Editor` (can submit, can edit their own records) but Approver and Admin must be explicitly granted.

**Do NOT put security gates only on the frontend.** `google.script.run` calls bypass the HTML entirely — any user can call any server function directly from the browser console. Every write function in the library must call `_hasRole_(capability)` and return `_denied_()` if the check fails.

**Admin email lockout protection:** `CONFIG.ADMIN_EMAIL` always resolves to `Admin` regardless of what's in the Access_Control sheet. This prevents you from accidentally revoking your own access and locking yourself out.

---

## §9 — Extra Fields JSON Column for Forward Compatibility

**Rule:** Add an `Extra_Fields` column (plain text) to any sheet that might need new fields in the future. Store extensible data there as JSON blobs.

**Why:** Adding a column to a Google Sheet is non-destructive, but it requires updating the header-driven write code, running a migration if existing rows need the new field, and potentially re-deploying. For fields that are optional, sparse, or experimental, the `Extra_Fields` JSON blob lets you add them with zero schema work. Future code can read them with `parseExtraFields_(record, 'fieldName')`.

**When to use a real column instead:** When the field is indexed, filtered on, or shown in a dashboard table. JSON blobs can't be queried by `getValues()` — you have to read every row and parse. For any field that appears in search results or KPI counts, use a real column.

---

## §10 — CacheService Chunking with Lock

**Rule:** Never store a payload larger than 90 KB in a single CacheService key. Always use `cacheGetOrFetch()` from `05_cache.js`, which handles chunking and concurrency automatically.

**Why:** CacheService has a hard 100 KB per-key limit. Dashboard JSON payloads from large sheets routinely exceed this. The chunking pattern splits the payload into 90 KB pieces, stores a chunk-count key, and reassembles on read.

The `LockService.getScriptLock()` in `cacheGetOrFetch()` prevents the *thundering herd* problem: without a lock, a cache miss under concurrent load causes every simultaneous request to hit the spreadsheet at the same time, which (a) exhausts the 6-second execution limit fast and (b) returns stale results to some callers. The lock ensures only one request builds the cache while others wait 25 seconds and then get the warm result.

**How to apply:** Use `cacheGetOrFetch(prefix, ttl, fetchFn)`. Use a versioned prefix (e.g., `myapp_v2_`) — bumping the version instantly invalidates all existing cache without needing to call `cacheClear()`.

---

## §11 — Report Generation: Template-First with Code Fallback

**Rule:** When generating Google Docs reports, always try a Drive Doc template first. Fall back to a code-generated Doc only if the template ID is missing or the copy fails.

**Why:** Template Docs let admins customize formatting, branding, and layout without touching code. But if an admin accidentally deletes the template, or you're setting up a new environment before the template exists, the fallback ensures the system continues generating reports rather than throwing errors at the user.

**How to apply:** Store template Doc IDs in the Config sheet (see §6). Call `createReport({ templateDocId: _getConfigValue_('MY_TEMPLATE_ID'), fallbackFn: myFallback })`. Write `myFallback` using `createSimpleReport()` or inline Doc construction.

---

## §12 — Notification Templates: Sheet-Driven Editable Messages

**Rule:** Notification text (email bodies, Teams cards, SMS messages) should be editable by admins without a code push. Store templates in a `Notification_Templates` sheet with `{TOKEN}` placeholders.

**Why:** The content of notifications changes frequently (legal reviews plain-English requirements, stakeholders ask for different wording, new fields get added). If the text is hardcoded, every change requires a library update and redeploy. The sheet-driven template system lets the admin open the sheet, edit the text, and the next notification picks up the change immediately.

**How to apply:** Call `setupNotificationTemplates()` during `setup()` to create the sheet. Populate rows with `Event` key and channel columns (`Email`, `Teams`, `SMS`). At send time: `const tmpl = getNotificationTemplate('MY_EVENT', 'Email'); const body = renderTemplate(tmpl, { ID: record.id, STATUS: record.status });`

---

## §13 — doGet Variable Injection Pattern

**Rule:** Pass server-side data to the frontend by injecting `<script>var KEY = value;</script>` into the HTML before `</head>`, not via `google.script.run` on page load.

**Why:** `google.script.run` adds a round-trip after the page renders — users see a loading state while critical bootstrap data (access level, config, deep-link params) loads. By injecting the data as global JS variables in the HTML response itself, the page has what it needs the moment the DOM parses. Use `google.script.run` only for *interactive* data fetches after the initial render.

**Security note:** Always sanitize URL parameters before injecting them. In `doGet(e)`, strip everything except safe characters before embedding: `String(params.id || '').replace(/[^a-zA-Z0-9\-_]/g, '')`.

**How to apply:** In `doGet()` (shell), build a `<script>` block with the variables, then `.replace('</head>', inject + '</head>')` on the HTML content. The frontend reads `window.PROGRAM_NAME`, `window.WEBAPP_URL`, `window.DEEP_LINK_ID`, etc. directly — no async needed.

---

## §14 — Batch Sheet I/O (Read All → Process → Write Once)

**Rule:** Never call `setValue()` or `appendRow()` inside a loop. Read all data into a JavaScript array first, process it, then write once with `setValues()` or `batchUpdateRows_()`.

**Why:** Every `setValue()` call is a separate Sheets API round-trip, roughly 100ms each. Updating 20 cells in a loop takes ~2 seconds. A single `setValues()` call with a 20-cell matrix takes ~100ms total — a 20× improvement. GAS scripts have a 6-minute hard execution limit; loops that hit the API per iteration will time out on large datasets.

**How to apply:**
- To update multiple cells in existing rows: use `batchUpdateRows_(sheet, [{rowIndex, obj}, ...])` from `02_helpers.js`
- To append multiple new rows: use `batchAppendRows_(sheet, [obj, ...])` from `02_helpers.js`
- For reads: call `sheet.getRange(...).getValues()` once and iterate the JavaScript array — never call `getRange()` inside a loop

**Exception:** Single-row updates (one `_updateRowByHeaders_` call) are fine with `setValue()` per column — the overhead is bounded. Only batch when you're in a loop over multiple rows.

---

## §15 — Fetch Retry with Exponential Backoff

**Rule:** Any `UrlFetchApp.fetch()` call to an external endpoint (Teams webhook, SIEM, external API) must use `fetchWithRetry_()` from `06_notify.js`, not a bare `fetch()`.

**Why:** External endpoints fail transiently — 5xx errors, network blips, load spikes. A bare `fetch()` that fails once silently drops the notification. Exponential backoff (1s → 2s → 4s) handles transient failures without overwhelming the endpoint.

**Do NOT retry on 4xx** — a 400 or 403 means the payload or credentials are wrong; retrying won't help and wastes quota.

**How to apply:** Replace any bare `UrlFetchApp.fetch(url, options)` in your project code with `fetchWithRetry_(url, options)`. The signature is identical. In `06_notify.js`, `sendTeamsWebhook` and `sendExternalWebhook` already use it internally.

---

## §16 — AI Integration: Multi-Provider with Circuit Breaker

**Rule:** All AI calls go through `callAI(prompt, opts)` — never call a provider directly. Store API keys in PropertiesService via `setAiApiKey()`, never in sheets or code.

**Why:** Any single AI provider can be unavailable (quota, outage, key expiry). `callAI()` tries providers in order (OpenRouter → Gemini → HuggingFace), skips providers whose circuit breaker is tripped, and falls back to `localFallbackAnalysis()` so the system always produces output. Without the circuit breaker, a down provider wastes 3× retry time on every call — killing performance for all users hitting the same trigger.

**Agent types:** `'summary'`, `'severity'`, `'tags'`, `'recommendation'`, `'general'`. Each has a built-in system prompt; override via `opts.systemPrompt` for project-specific tone.

**Local fallback:** Keyword-based severity classification and tag extraction that works with zero API calls. Quality is lower but the system stays up. A severity call against a text mentioning "ransomware" returns `Critical` without any network request.

**Config sheet keys:** `AI_PROVIDER` (preferred provider), `AI_MODEL` (model override), `AI_MAX_TOKENS`.

**How to apply:**
```javascript
// One-time setup (run from editor):
setAiApiKey('openrouter', 'sk-or-...');

// In any trigger or form handler:
const summary = callAI(submissionText, { agentType: 'summary' });
const sev     = callAI(submissionText, { agentType: 'severity' });
const tags    = callAI(submissionText, { agentType: 'tags' });

// Admin diagnostics (safe to show in dashboard):
const status = getAiStatus();  // { openrouter: { hasKey, failures, circuitOpen }, ... }
```

---

## §17 — Project Scaffolding: Folder + Registry on Demand

**Rule:** Use `scaffoldProject(opts)` when code needs to create a new project programmatically. It handles Drive folder hierarchy, central registry append, and audit log in one call.

**Why:** Manually creating folders and updating registries is error-prone and non-repeatable. The scaffolder ensures every project has a consistent folder structure (`root/category/projectName/`) and a registry row with the same fields every time.

**Registry vs. backend sheet:** The registry spreadsheet is a single org-wide "all projects" tracker (not the project's own `BACKEND_SHEET_ID`). Pass `registrySheetId` to `scaffoldProject()` if you maintain one; omit it if each project is standalone.

**Notion:** A `_notionUpdatePage_()` stub is in `09_scaffolder.js` — uncomment and fill in a `NOTION_TOKEN` PropertiesService key when you need Notion integration.

**How to apply:**
```javascript
const proj = scaffoldProject({
  name:            'Browser Risk Dashboard',
  category:        'Security',
  owner:           'remon.attalla@lausd.net',
  registrySheetId: 'YOUR_REGISTRY_SHEET_ID'   // optional
});
Logger.log(proj.id);       // PROJ-2026-001
Logger.log(proj.folderId); // Drive folder ID
Logger.log(proj.slug);     // 'browser-risk-dashboard'
```

---

## Quick Reference: What Goes Where

| Concern | Where it lives |
|---|---|
| Business logic, data reads/writes | Library (02–07) |
| Role definitions, capabilities | Library (03_access.js) |
| Brand colors, email layout | Library (04_email.js) |
| Notification channel dispatch | Library (06_notify.js) |
| Report generation | Library (07_reports.js) |
| Runtime-variable config (URLs, template IDs) | Config sheet + `_getConfigValue_()` |
| Identity constants (org name, emails, ID prefix) | lib/01_config.js |
| Mirror of identity constants | template/config.js (keep-in-sync) |
| `onOpen()`, `doGet()`, `getMyAccess()` | Shell (inlined — §3) |
| Project-specific sheets, domain logic | Shell + additional files per project |
| Text content of notifications | Notification_Templates sheet (§12) |
| Template Doc IDs, folder IDs | Config sheet (§6) |
| Dashboard summary + recent records | `getDashboard()` in `02_helpers.js` — wrap with `cacheGetOrFetch()` in the shell |
| AI analysis (summary, severity, tags) | `callAI(prompt, { agentType })` in `08_ai.js` — auto-provider + local fallback |
| AI API keys | PropertiesService via `setAiApiKey(provider, key)` — never in sheets or code |
| AI circuit breaker state | `getAiStatus()` / `resetAiCircuit(provider)` in `08_ai.js` |
| New project scaffold (Drive + registry) | `scaffoldProject(opts)` in `09_scaffolder.js` |
| URL-safe slug | `generateSlug(name)` in `09_scaffolder.js` |

---

## §13 — App Registry & Ship Checklist

**Rule:** Every new project must be registered in the central App Registry right after first setup. Record the Script ID (obtain via `ScriptApp.getScriptId()` or the `.clasp.json`), status, and security review date.

**Why:** Avoids orphaned projects, makes Script IDs discoverable for updates/audits/hand-offs, and creates a single source of truth for all LAUSD GAS work.

**How to apply:**
- After running `setup()`, use the shell menu: `PROG → Setup & Config → Ship Checklist`
- Copy the printed Script ID into the registry
- Update status as the project moves through lifecycle (Dev → Production → Maintenance → Archived)
- Re-run the checklist whenever you change the bound script or library linkage

The `menuShipChecklist()` helper lives in `template/shell.js` and is wired into the admin menu.
