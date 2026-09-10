/**
 * RiskAI-GAS — Helpers (02_helpers.js)
 *
 * Sheet I/O, ID lookup, ID generation, audit logging, Drive folders,
 * Config sheet key/value store, and general utilities.
 *
 * All private functions use _underscorePadding_ naming convention.
 * Public API functions have no underscores.
 */

// ==========================================================================
// SPREADSHEET ACCESS
// ==========================================================================

function _ss_() {
  return SpreadsheetApp.openById(CONFIG.BACKEND_SHEET_ID);
}

function _getSheet_(tabName) {
  return _ss_().getSheetByName(tabName);
}

function _getOrCreateSheet_(tabName) {
  const ss = _ss_();
  let sheet = ss.getSheetByName(tabName);
  if (!sheet) sheet = ss.insertSheet(tabName);
  return sheet;
}

// ==========================================================================
// HEADER MAP — {COLUMN_NAME: zero-based-index} from row 1
// ==========================================================================

function _headerMap_(sheet) {
  if (!sheet || sheet.getLastColumn() < 1) return {};
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const map = {};
  headers.forEach((h, i) => {
    if (h) {
      const key = String(h).trim().toLowerCase();
      map[key] = i;
    }
  });
  return map;
}

// ==========================================================================
// HEADER-DRIVEN WRITES — immune to column reordering
// ==========================================================================

/**
 * Append a row to `sheet` using {columnName: value} obj.
 * Columns not present in obj are left blank; unknown keys in obj are ignored.
 */
function _appendRowByHeaders_(sheet, obj) {
  if (!sheet || !obj) return;
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const row = headers.map(h => {
    const v = obj[String(h).trim()];
    return v != null ? v : '';
  });
  sheet.appendRow(row);
  return sheet.getLastRow();
}

/**
 * Update an existing row (1-based rowIndex) using {columnName: value} obj.
 * Only the columns present in obj are written; others are untouched.
 */
function _updateRowByHeaders_(sheet, rowIndex, obj) {
  if (!sheet || !obj || rowIndex < 2) return;
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  headers.forEach((h, i) => {
    const key = String(h).trim();
    if (key && obj.hasOwnProperty(key) && obj[key] != null) {
      sheet.getRange(rowIndex, i + 1).setValue(obj[key]);
    }
  });
}

// ==========================================================================
// ID LOOKUP — display values first (avoids Sheets auto-coercion bug)
// WHY: Sheets coerces numeric IDs like "6216" → 6216 (float), so raw values
// can appear as "6216.0". getDisplayValues() returns what the user sees.
// ==========================================================================

function _normalizeId_(id) {
  if (id == null) return '';
  let s = String(id).trim().toLowerCase();
  // Pure numeric: strip leading zeros + .0 suffix (Sheets coercion)
  if (/^\d+\.?\d*$/.test(s)) return String(parseInt(s, 10));
  if (s.endsWith('.0')) s = s.slice(0, -2);
  return s;
}

function _idEq_(a, b) {
  const na = _normalizeId_(a);
  const nb = _normalizeId_(b);
  if (na === nb) return true;
  const naN = parseInt(na, 10);
  const nbN = parseInt(nb, 10);
  if (!isNaN(naN) && !isNaN(nbN) && naN === nbN) return true;
  return false;
}

/**
 * Find 1-based row index of a record by its ID (col 1 of tabName).
 * Returns -1 if not found.
 */
function findRowById_(tabName, id) {
  const sheet = _getSheet_(tabName);
  if (!sheet || sheet.getLastRow() < 2) return -1;
  const target = String(id || '').trim();
  if (!target) return -1;

  const range = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1);

  // Pass 1: display values (avoids auto-coercion)
  const display = range.getDisplayValues();
  for (let i = 0; i < display.length; i++) {
    if (_idEq_(display[i][0], target)) return i + 2;
  }
  // Pass 2: raw values with Date fallback
  const raw = range.getValues();
  for (let j = 0; j < raw.length; j++) {
    const v = raw[j][0];
    const s = (v instanceof Date)
      ? Utilities.formatDate(v, CONFIG.TIMEZONE, 'yyyy-MM-dd HH:mm')
      : v;
    if (_idEq_(s, target)) return j + 2;
  }
  return -1;
}

/**
 * Fetch a record by ID as {header: value}. Returns null if not found.
 * Includes _row (1-based) for use with _updateRowByHeaders_.
 */
function findRecordById_(tabName, id) {
  const rowIdx = findRowById_(tabName, id);
  if (rowIdx < 0) return null;
  const sheet = _getSheet_(tabName);
  const width = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, width).getValues()[0];
  const vals    = sheet.getRange(rowIdx, 1, 1, width).getValues()[0];
  const out = { _row: rowIdx };
  for (let c = 0; c < headers.length; c++) {
    const h = String(headers[c] || '').trim();
    if (!h) continue;
    let v = vals[c];
    if (v instanceof Date) v = Utilities.formatDate(v, CONFIG.TIMEZONE, 'yyyy-MM-dd HH:mm');
    out[h] = v;
  }
  return out;
}

/**
 * Fetch all data rows from tabName as an array of {header: value} objects.
 */
function getAllRecords_(tabName) {
  const sheet = _getSheet_(tabName);
  if (!sheet || sheet.getLastRow() < 2) return [];
  const width = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, width).getValues()[0];
  const rows    = sheet.getRange(2, 1, sheet.getLastRow() - 1, width).getValues();
  return rows.map(row => {
    const obj = {};
    for (let c = 0; c < headers.length; c++) {
      const h = String(headers[c] || '').trim();
      if (!h) continue;
      let v = row[c];
      if (v instanceof Date) v = Utilities.formatDate(v, CONFIG.TIMEZONE, 'yyyy-MM-dd HH:mm');
      obj[h] = v;
    }
    return obj;
  });
}

// ==========================================================================
// ID GENERATION — PREFIX-YYYY-NNN style (e.g., ARB-2026-001)
// ==========================================================================

function _nextSequentialId_(prefix, tabName) {
  const year = new Date().getFullYear();
  const sheet = _getSheet_(tabName);
  if (!sheet || sheet.getLastRow() < 2) return prefix + '-' + year + '-001';
  const ids = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getDisplayValues();
  let max = 0;
  const pat = new RegExp('^' + prefix + '-' + year + '-(\\d+)$', 'i');
  ids.forEach(r => {
    const m = String(r[0] || '').trim().match(pat);
    if (m) { const n = parseInt(m[1], 10); if (n > max) max = n; }
  });
  return prefix + '-' + year + '-' + String(max + 1).padStart(3, '0');
}

/** Generate next ID — disabled for Agent Platform MVP (no Master records). */
function nextRecordId() {
  return 'AGT-' + new Date().getFullYear() + '-001'; // stub
}

// ==========================================================================
// CONFIG SHEET — runtime key/value store (Config tab)
// ==========================================================================

function _getConfigValue_(key) {
  const sheet = _getSheet_(CONFIG.SHEET_TABS.CONFIG);
  if (!sheet || sheet.getLastRow() < 2) return '';
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][0]).trim() === key) return String(data[i][1] || '').trim();
  }
  return '';
}

function _setConfigValue_(key, value) {
  const sheet = _getOrCreateSheet_(CONFIG.SHEET_TABS.CONFIG);
  if (sheet.getLastRow() < 1) {
    sheet.getRange(1, 1, 1, 3).setValues([['Key', 'Value', 'Notes']]).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  if (sheet.getLastRow() >= 2) {
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
    for (let i = 0; i < data.length; i++) {
      if (String(data[i][0]).trim() === key) {
        sheet.getRange(i + 2, 2).setValue(value);
        return;
      }
    }
  }
  sheet.appendRow([key, value, '']);
}

// getWebAppUrl / setWebAppUrl removed to avoid dup with shell.js inlined version (uses ScriptApp directly for pilot).

// ==========================================================================
// AUDIT LOG — timestamped action log (Audit_Log tab)
// ==========================================================================

function auditLog_(recordId, action, details) {
  try {
    const sheet = _getOrCreateSheet_(CONFIG.SHEET_TABS.AUDIT_LOG);
    // Create header row if empty
    if (sheet.getLastRow() < 1) {
      sheet.appendRow(['Timestamp', 'Record_ID', 'Action', 'Actor', 'Details']);
      sheet.getRange(1, 1, 1, 5).setFontWeight('bold');
      sheet.setFrozenRows(1);
    }
    let actor = 'system';
    try { actor = Session.getActiveUser().getEmail() || 'system'; } catch (e) {}
    const ts = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyy-MM-dd HH:mm:ss');
    sheet.appendRow([ts, String(recordId || ''), String(action || ''), actor, String(details || '')]);
  } catch (e) {
    Logger.log('auditLog_ failed (non-fatal): ' + e.message);
  }
}

// ==========================================================================
// DRIVE — root folder + project subfolders
// ==========================================================================

function _getOrCreateRootFolder_() {
  // 1. Hardcoded in config
  if (CONFIG.ROOT_FOLDER_ID) {
    try { return DriveApp.getFolderById(CONFIG.ROOT_FOLDER_ID); }
    catch (e) { Logger.log('CONFIG.ROOT_FOLDER_ID invalid: ' + e.message); }
  }
  // 2. Config sheet override
  const stored = _getConfigValue_('ROOT_FOLDER_ID');
  if (stored) {
    try { return DriveApp.getFolderById(stored); }
    catch (e) { Logger.log('Config ROOT_FOLDER_ID invalid: ' + e.message); }
  }
  // 3. Auto-create by name (idempotent)
  const folderName = CONFIG.PROGRAM_NAME + ' — Files';
  const it = DriveApp.getFoldersByName(folderName);
  if (it.hasNext()) {
    const f = it.next();
    _setConfigValue_('ROOT_FOLDER_ID', f.getId());
    return f;
  }
  const f = DriveApp.createFolder(folderName);
  _setConfigValue_('ROOT_FOLDER_ID', f.getId());
  return f;
}

function _getOrCreateSubfolder_(parent, name) {
  const it = parent.getFoldersByName(name);
  if (it.hasNext()) return it.next();
  return parent.createFolder(name);
}

/** Create per-record folder under root/records/recordId. Idempotent. */
function getOrCreateRecordFolder_(recordId) {
  const root    = _getOrCreateRootFolder_();
  const records = _getOrCreateSubfolder_(root, 'records');
  return _getOrCreateSubfolder_(records, recordId);
}

// ==========================================================================
// JSON HELPERS — safe parse/stringify for sheet columns storing JSON
// ==========================================================================

function _parseJson_(raw) {
  if (!raw) return {};
  if (typeof raw === 'object') return raw;
  try { return JSON.parse(String(raw)); } catch (e) { return {}; }
}

function _stringifyJson_(obj) {
  try { return JSON.stringify(obj || {}); } catch (e) { return '{}'; }
}

// ==========================================================================
// UTILITY — user, date, HTML escape
// ==========================================================================

function _getUserEmail_() {
  try { return String(Session.getActiveUser().getEmail() || '').trim().toLowerCase(); }
  catch (e) { return ''; }
}

function _now_() {
  return Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyy-MM-dd HH:mm:ss');
}

function _fmtDate_(d, fmt) {
  if (!d) return '';
  if (!(d instanceof Date)) {
    const p = new Date(d);
    if (isNaN(p.getTime())) return String(d);
    d = p;
  }
  return Utilities.formatDate(d, CONFIG.TIMEZONE, fmt || 'yyyy-MM-dd HH:mm');
}

/** HTML-escape a value for safe embedding in email/HTML output. */
function _esc_(s) {
  return s == null ? '' : String(s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

// ==========================================================================
// BATCH SHEET WRITES — setValues() once, not setValue() per cell
// WHY: Every call to setValue() is a separate API round-trip (~100ms each).
// Updating 20 cells one at a time takes ~2 seconds; batchUpdateRows_() does it
// in one call. Grok + Google best practices both flag this as the #1 perf win.
// ==========================================================================

/**
 * Apply multiple row updates to a sheet in a single Sheets API call.
 *
 * sheet:   Sheet object
 * updates: Array of { rowIndex: Number (1-based), obj: {ColumnName: value} }
 *
 * Reads headers once, builds a value matrix, writes with one setValues() call.
 *
 * Example:
 *   batchUpdateRows_(sheet, [
 *     { rowIndex: 3, obj: { Status: 'Approved', Updated_Date: new Date() } },
 *     { rowIndex: 7, obj: { Status: 'Rejected', Updated_Date: new Date() } }
 *   ]);
 */
function batchUpdateRows_(sheet, updates) {
  if (!sheet || !updates || !updates.length) return;
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const colCount = headers.length;

  // Group updates by row so we can write each row in one range call
  const byRow = {};
  updates.forEach(function(u) {
    if (!u || !u.rowIndex || !u.obj) return;
    if (!byRow[u.rowIndex]) byRow[u.rowIndex] = {};
    Object.assign(byRow[u.rowIndex], u.obj);
  });

  // For each affected row, build the full-width value array and write once
  Object.keys(byRow).forEach(function(rowIdx) {
    const row = byRow[rowIdx];
    // Read existing row values so we only overwrite specified columns
    const existing = sheet.getRange(Number(rowIdx), 1, 1, colCount).getValues()[0];
    const newRow   = headers.map(function(h, i) {
      const key = String(h).trim();
      return row.hasOwnProperty(key) ? row[key] : existing[i];
    });
    sheet.getRange(Number(rowIdx), 1, 1, colCount).setValues([newRow]);
  });
}

/**
 * Append multiple rows to a sheet in a single setValues() call.
 * Much faster than calling appendRow() in a loop.
 *
 * sheet:   Sheet object
 * objs:    Array of {ColumnName: value} objects
 *
 * Example:
 *   batchAppendRows_(sheet, [
 *     { ID: 'REC-001', Title: 'First',  Status: 'Open' },
 *     { ID: 'REC-002', Title: 'Second', Status: 'Open' }
 *   ]);
 */
function batchAppendRows_(sheet, objs) {
  if (!sheet || !objs || !objs.length) return;
  const headers  = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const colCount = headers.length;
  const rows     = objs.map(function(obj) {
    return headers.map(function(h) {
      const v = obj[String(h).trim()];
      return v != null ? v : '';
    });
  });
  const startRow = sheet.getLastRow() + 1;
  sheet.getRange(startRow, 1, rows.length, colCount).setValues(rows);
}

// ==========================================================================
// EXTRA FIELDS JSON — forward-compatible extensibility column
// Store arbitrary key-value data in a single "Extra_Fields" JSON column so
// new fields never require a schema migration. Pattern from IRS v3.4.
// ==========================================================================

/**
 * Parse the Extra_Fields JSON column from a record object.
 * record: result of findRecordById_() or any {header: value} object
 * fieldName: optional — if provided, returns just that field value
 *
 * Example:
 *   const fields = parseExtraFields_(record);           // returns {}
 *   const team   = parseExtraFields_(record, 'irTeam'); // returns 'SOC'
 */
function parseExtraFields_(record, fieldName) {
  const raw = record && (record['Extra_Fields'] || record['Extra Fields'] || '');
  const obj = _parseJson_(raw);
  return fieldName !== undefined ? (obj[fieldName] != null ? obj[fieldName] : '') : obj;
}

/**
 * Update a single field inside the Extra_Fields JSON blob in a sheet row.
 * Does a read-modify-write on the cell — does not touch other fields.
 *
 * sheet:     Sheet object
 * rowIndex:  1-based row number
 * fieldName: key to set inside the JSON blob
 * value:     value to store
 */
function setExtraField_(sheet, rowIndex, fieldName, value) {
  const hdr    = _headerMap_(sheet);
  const colIdx = hdr['Extra_Fields'] != null ? hdr['Extra_Fields']
               : hdr['Extra Fields']  != null ? hdr['Extra Fields'] : -1;
  if (colIdx < 0) {
    Logger.log('setExtraField_: no Extra_Fields column in sheet ' + sheet.getName());
    return;
  }
  const cell     = sheet.getRange(rowIndex, colIdx + 1);
  const existing = _parseJson_(cell.getValue());
  existing[fieldName] = value;
  cell.setValue(_stringifyJson_(existing));
}

// ==========================================================================
// COLUMN PROTECTION — admin-only edit protection
// Pattern from gas-sop-lausd-standalone (menuLockStatusColumn).
// ==========================================================================

/**
 * Protect a sheet column so only the admin email can edit it.
 * Idempotent — removes any existing protection on the column first.
 *
 * sheet:           Sheet object
 * colNameOrIndex:  column header name (string) OR 1-based column number
 * adminEmail:      email that retains edit rights (defaults to CONFIG.ADMIN_EMAIL)
 *
 * Example:
 *   protectColumn_(sheet, 'Status', CONFIG.ADMIN_EMAIL);
 *   protectColumn_(sheet, 2);  // protect column B by index
 */
function protectColumn_(sheet, colNameOrIndex, adminEmail) {
  adminEmail = adminEmail || CONFIG.ADMIN_EMAIL;
  let colIndex;
  if (typeof colNameOrIndex === 'string') {
    const hdr = _headerMap_(sheet);
    if (hdr[colNameOrIndex] == null) {
      Logger.log('protectColumn_: column "' + colNameOrIndex + '" not found');
      return;
    }
    colIndex = hdr[colNameOrIndex] + 1; // 1-based
  } else {
    colIndex = colNameOrIndex;
  }
  // Remove any existing protection on this column
  sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE)
    .filter(p => p.getRange().getColumn() === colIndex && p.getRange().getNumColumns() === 1)
    .forEach(p => p.remove());

  const range = sheet.getRange(1, colIndex, sheet.getMaxRows(), 1);
  const prot  = range.protect().setDescription('Admin-only: col ' + colIndex);
  prot.removeEditors(prot.getEditors());
  if (adminEmail) prot.addEditor(adminEmail);
  prot.setWarningOnly(false);
}

// ==========================================================================
// CHANGELOG — automatic audit trail on sheet edits
// Pattern from gas-sop-lausd (onEditTrigger + _logChangelog_).
//
// Usage:
//   1. Call _ensureChangelogSheet_() once during setup().
//   2. In your onEdit(e) trigger, call logChangelog_(recordId, field, old, new).
// ==========================================================================

function _ensureChangelogSheet_() {
  const sheet = _getOrCreateSheet_(CONFIG.SHEET_TABS.CHANGELOG);
  if (sheet.getLastRow() < 1) {
    sheet.appendRow(['Timestamp', 'Record_ID', 'Field', 'Old_Value', 'New_Value', 'Changed_By']);
    sheet.getRange(1, 1, 1, 6).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/**
 * Append a row to the Changelog sheet recording a field change.
 * Safe to call from onEdit() triggers — failures are non-fatal.
 *
 * recordId: ID of the changed record (or '' for sheet-level changes)
 * field:    column/field name that changed
 * oldValue: previous value
 * newValue: new value
 */
function logChangelog_(recordId, field, oldValue, newValue) {
  if (String(oldValue) === String(newValue)) return; // no-op if unchanged
  // Never log changes to LastUpdated itself (avoid recursive loops)
  if (field === 'LastUpdated' || field === 'Last_Updated') return;
  try {
    const sheet = _ensureChangelogSheet_();
    let actor = 'system';
    try { actor = Session.getActiveUser().getEmail() || 'system'; } catch (e) {}
    sheet.appendRow([_now_(), String(recordId || ''), field, String(oldValue || ''), String(newValue || ''), actor]);
  } catch (e) {
    Logger.log('logChangelog_ failed (non-fatal): ' + e.message);
  }
}

// ==========================================================================
// SETUP — see _runBasicSetup in template/10_admin_ingest.js (inlined for pilot)
// ==========================================================================

// setup() is defined in shell.js for RPC; the full logic is _runBasicSetup (called by setup).
// Removed duplicate 'setup' from here to prevent global redeclaration.

// ==========================================================================
// DASHBOARD DATA — summary stats + recent records for the frontend
// Returns a JSON-serialisable object; wrap with cacheGetOrFetch() in the shell
// when the Master tab is large to avoid reading the sheet on every page load.
// ==========================================================================

/**
 * Read the Master sheet and return a lightweight dashboard snapshot.
 *
 * Returns:
 *   {
 *     total:      Number   — total record count
 *     openCount:  Number   — rows where Status is 'Open' or 'In Progress'
 *     recent:     Array    — last N records (newest first), capped at recentLimit
 *     asOf:       String   — timestamp of this read (for staleness display)
 *   }
 *
 * recentLimit defaults to 10. Pass a higher number if the dashboard table shows more.
 *
 * Usage in shell (inlined doGet path — see PATTERNS §3 / §13):
 *   const dash = getDashboard();  // or: cacheGetOrFetch('myapp_dash_', 300, getDashboard)
 *   const inject = '<script>var DASHBOARD=' + JSON.stringify(dash) + ';</script>';
 */
function getDashboard(recentLimit) {
  // Agent Platform MVP does not use Master/dashboard. Return empty.
  return {
    total: 0,
    openCount: 0,
    recent: [],
    asOf: _now_(),
    note: 'Dashboard disabled for gas-agent-platform MVP (use Agents tab)'
  };
}

// ==========================================================================
// AGENT PLATFORM (MVP) — sheet-driven agents + KB + chat
// NOTE: chatWithAgent_, getAgent_, getKbSnippets_ etc inlined in template/20_chat.js + shell.js
// (to avoid library and duplicate global function defs)
// ==========================================================================

/**
 * Return list of active agents (status=on) as [{slug, name, ...}]
 */
function getActiveAgents_() {
  try {
    // hardcode against the known sheet to ensure
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName('Agents');
    if (!sheet) return [];
    const map = _headerMap_(sheet);
    const out = [];
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (String(row[map.status] || '').toLowerCase() === 'on') {
        out.push({
          slug: row[map.slug] || '',
          name: row[map.name] || row[map.slug] || '',
          org: row[map.org] || 'LAUSD'
        });
      }
    }
    return out;
  } catch (e) {
    // let caller handle
    throw e;
  }
}
