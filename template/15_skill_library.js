/**
 * 15_skill_library.js
 *
 * A "Claude Skills"-style library: reusable, named instruction packages that
 * can be attached to any agent by key. This is distinct from the Agents.skills
 * column (v1.2.0) -- that column toggles built-in CAPABILITIES (kb, webpage,
 * search, memory, email, diagram). This is a library of CONTENT -- procedures,
 * checklists, methodologies -- imported from anywhere (starting with GitHub)
 * and attached per-agent via a new Agents.skill_refs column.
 *
 * SkillLibrary tab: key | title | instructions | source_url | added_at | outputs
 *   outputs = comma list of any subset of {sheet, doc, html}. Used by
 *   agent_type=skill (26_skill_agent.js). Blank is fine for library rows
 *   attached to persona agents via skill_refs.
 * Agents.skill_refs: comma list of SkillLibrary keys this agent should load.
 */

var SKILL_LIBRARY_HEADERS = ['key', 'title', 'instructions', 'source_url', 'added_at', 'outputs'];
var SKILL_OUTPUT_KINDS = ['sheet', 'doc', 'html'];
var MAX_SKILL_LIBRARY_CHARS = 6000; // hard ceiling on injected skill-library context, mirrors MAX_KB_CHARS
var MAX_SKILL_IMPORT_CHARS = 20000; // hard ceiling on any single imported skill stored in the cell

function _getSkillLibrarySheet_() {
  const sheet = _getOrCreateTab_('SkillLibrary', SKILL_LIBRARY_HEADERS);
  _ensureSkillLibrarySchema_();
  return sheet;
}

// Append-only / non-destructive: adds missing SkillLibrary columns (currently
// `outputs`) without rewriting existing cells. Blank outputs stays blank.
function _ensureSkillLibrarySchema_() {
  const sheet = _getOrCreateTab_('SkillLibrary', SKILL_LIBRARY_HEADERS);
  const lastCol = Math.max(sheet.getLastColumn(), 1);
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function (h) {
    return String(h || '').trim().toLowerCase();
  });
  const toAdd = [];
  SKILL_LIBRARY_HEADERS.forEach(function (h) {
    if (headers.indexOf(h) < 0) toAdd.push(h);
  });
  if (!toAdd.length) return [];
  sheet.getRange(1, lastCol + 1, 1, toAdd.length).setValues([toAdd]);
  sheet.getRange(1, lastCol + 1, 1, toAdd.length).setFontWeight('bold');
  return toAdd;
}

function migrateSkillLibraryAddOutputs_() {
  if (!_isAdminUser_()) { _toast_('Admin only.'); return; }
  const added = _ensureSkillLibrarySchema_();
  if (!added.length) {
    _toast_('SkillLibrary already has an outputs column.');
    return;
  }
  _toast_('Added ' + added.join(', ') + ' to SkillLibrary. For skill agents, set a comma list of sheet, doc, and/or html.');
}

function _skillLibraryHeaderMap_(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const map = {};
  headers.forEach(function (h, i) {
    const key = String(h || '').trim().toLowerCase();
    if (key) map[key] = i;
  });
  return { headers: headers, map: map };
}

// Header-driven upsert so extra columns (outputs, future fields) survive.
// `outputs` is optional: omitted on update keeps the existing cell.
function _upsertSkillLibraryRow_(key, title, instructions, sourceUrl, outputs) {
  const sheet = _getSkillLibrarySheet_();
  const meta = _skillLibraryHeaderMap_(sheet);
  const map = meta.map;
  if (map.key === undefined) throw new Error('SkillLibrary is missing a key column.');

  const data = sheet.getDataRange().getValues();
  let rowIdx = -1;
  const want = String(key || '').trim().toLowerCase();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][map.key] || '').trim().toLowerCase() === want) {
      rowIdx = i + 1;
      break;
    }
  }

  const width = Math.max(sheet.getLastColumn(), SKILL_LIBRARY_HEADERS.length);
  const row = rowIdx > 0 ? data[rowIdx - 1].slice() : [];
  while (row.length < width) row.push('');

  row[map.key] = want;
  if (map.title !== undefined) row[map.title] = title || want;
  if (map.instructions !== undefined) row[map.instructions] = instructions || '';
  if (map.source_url !== undefined) row[map.source_url] = sourceUrl || '';
  if (map.added_at !== undefined) row[map.added_at] = new Date();
  if (map.outputs !== undefined && outputs !== undefined && outputs !== null) {
    row[map.outputs] = _formatSkillOutputs_(outputs);
  }

  if (rowIdx > 0) {
    sheet.getRange(rowIdx, 1, 1, row.length).setValues([row]);
    return 'updated';
  }
  sheet.appendRow(row);
  return 'inserted';
}

function _normalizeSkillOutputs_(raw) {
  const parts = String(raw || '')
    .replace(/[\[\]"'`]/g, '')
    .split(/[,\s]+/)
    .map(function (s) { return s.trim().toLowerCase(); })
    .filter(function (s) { return SKILL_OUTPUT_KINDS.indexOf(s) >= 0; });
  const seen = {};
  const out = [];
  parts.forEach(function (p) {
    if (!seen[p]) {
      seen[p] = true;
      out.push(p);
    }
  });
  return out;
}

function _formatSkillOutputs_(raw) {
  const joined = Array.isArray(raw) ? raw.join(',') : String(raw || '');
  return _normalizeSkillOutputs_(joined).join(',');
}

function _getSkillLibraryRow_(key) {
  const want = String(key || '').trim().toLowerCase();
  if (!want) return null;
  const sheet = _getSs_().getSheetByName('SkillLibrary');
  if (!sheet || sheet.getLastRow() < 2) return null;
  _ensureSkillLibrarySchema_();
  const meta = _skillLibraryHeaderMap_(sheet);
  const map = meta.map;
  if (map.key === undefined) return null;
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][map.key] || '').trim().toLowerCase() !== want) continue;
    return {
      key: want,
      title: map.title !== undefined ? String(data[i][map.title] || '') : want,
      instructions: map.instructions !== undefined ? String(data[i][map.instructions] || '') : '',
      source_url: map.source_url !== undefined ? String(data[i][map.source_url] || '') : '',
      outputs: map.outputs !== undefined ? String(data[i][map.outputs] || '') : ''
    };
  }
  return null;
}

function _listSkillLibraryKeys_() {
  const sheet = _getSs_().getSheetByName('SkillLibrary');
  if (!sheet || sheet.getLastRow() < 2) return [];
  const meta = _skillLibraryHeaderMap_(sheet);
  if (meta.map.key === undefined) return [];
  const data = sheet.getDataRange().getValues();
  const keys = [];
  for (let i = 1; i < data.length; i++) {
    const k = String(data[i][meta.map.key] || '').trim().toLowerCase();
    if (k) keys.push(k);
  }
  return keys;
}

function _slugifySkillKey_(s) {
  return String(s || '').trim().toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/^-+|-+$/g, '');
}

// Minimal YAML frontmatter parser for Claude-style SKILL.md files.
// Supports single-line `key: value`, `description: >` / `|` blocks, and
// simple `- item` lists (used for outputs). Everything after the closing
// --- is the editable instruction body.
function _parseSkillFrontmatter_(text) {
  text = String(text || '').replace(/^\uFEFF/, '');
  const m = text.match(/^---[ \t]*\r?\n([\s\S]*?)\r?\n---[ \t]*\r?\n?([\s\S]*)$/);
  if (!m) return { meta: {}, body: text };
  const meta = {};
  const lines = m[1].split(/\r?\n/);
  let i = 0;
  while (i < lines.length) {
    const km = lines[i].match(/^([A-Za-z0-9_-]+)\s*:\s*(.*)$/);
    if (!km) { i++; continue; }
    const key = km[1].toLowerCase();
    let val = km[2].trim();
    if (val === '>' || val === '|' || val === '>-' || val === '|-') {
      const buf = [];
      i++;
      while (i < lines.length && (/^[ \t]/.test(lines[i]) || lines[i] === '')) {
        buf.push(lines[i].replace(/^[ \t]+/, ''));
        i++;
      }
      meta[key] = buf.join(' ').trim();
      continue;
    }
    if (val === '' && i + 1 < lines.length && /^\s+-\s+/.test(lines[i + 1])) {
      const items = [];
      i++;
      while (i < lines.length && /^\s+-\s+/.test(lines[i])) {
        items.push(lines[i].replace(/^\s+-\s+/, '').trim());
        i++;
      }
      meta[key] = items.join(', ');
      continue;
    }
    if ((val.charAt(0) === '"' && val.charAt(val.length - 1) === '"') ||
        (val.charAt(0) === "'" && val.charAt(val.length - 1) === "'")) {
      val = val.slice(1, -1);
    }
    meta[key] = val;
    i++;
  }
  return { meta: meta, body: m[2] };
}

function _pickMarkdownFromUnzipped_(blobs) {
  let fallback = null;
  for (let i = 0; i < blobs.length; i++) {
    const name = String(blobs[i].getName() || '').replace(/\\/g, '/');
    const base = name.split('/').pop();
    if (!/\.(md|markdown)$/i.test(base)) continue;
    if (/^skill\.md$/i.test(base)) return blobs[i].getDataAsString();
    if (!fallback) fallback = blobs[i];
  }
  if (fallback) return fallback.getDataAsString();
  throw new Error('Zip contained no markdown file (looked for SKILL.md or *.md).');
}

function _extractSkillMarkdownFromBlob_(blob, hintName) {
  const name = String(hintName || (blob && blob.getName && blob.getName()) || '').toLowerCase();
  const mime = String((blob && blob.getContentType && blob.getContentType()) || '').toLowerCase();
  const looksZip = /\.zip$/.test(name) || /zip/.test(mime);
  if (looksZip) {
    return _pickMarkdownFromUnzipped_(Utilities.unzip(blob));
  }
  try {
    const unzipped = Utilities.unzip(blob);
    if (unzipped && unzipped.length) return _pickMarkdownFromUnzipped_(unzipped);
  } catch (e) { /* not a zip -- treat as plain text */ }
  return blob.getDataAsString();
}

function _writeIngestedSkill_(rawText, sourceUrl, optKey, optTitle) {
  const parsed = _parseSkillFrontmatter_(rawText);
  const meta = parsed.meta || {};
  const key = _slugifySkillKey_(optKey || meta.name || meta.key || '');
  if (!key) throw new Error('Could not determine a skill key. Pass one, or include `name:` in the YAML frontmatter.');
  const title = String(optTitle || meta.description || meta.title || meta.name || key).trim();
  const outputs = _formatSkillOutputs_(meta.outputs || meta.output || '');
  const instructions = String(parsed.body || rawText || '').trim().slice(0, MAX_SKILL_IMPORT_CHARS);
  if (!instructions) throw new Error('Skill file had no instruction text.');
  const result = _upsertSkillLibraryRow_(key, title, instructions, sourceUrl || '', outputs);
  return { result: result, key: key, title: title, outputs: outputs };
}

// Reads GitHub content two ways: a raw.githubusercontent.com URL (or any
// plain-text URL) is fetched as-is; a normal github.com/.../blob/... URL is
// rewritten to its raw equivalent first, since fetching the blob HTML page
// would ingest GitHub's page chrome instead of the actual file content.
function _normalizeGithubRawUrl_(url) {
  url = String(url || '').trim();
  const m = url.match(/^https:\/\/github\.com\/([^\/]+)\/([^\/]+)\/blob\/(.+)$/);
  if (m) {
    return 'https://raw.githubusercontent.com/' + m[1] + '/' + m[2] + '/' + m[3];
  }
  return url;
}

// ingestSkillFromUrl_ — UrlFetchApp + Utilities.unzip for a Claude-style
// skill zip (markdown + YAML header) from a GitHub raw URL or any
// accessible URL; also supports a raw .md. Writes instruction text into
// SkillLibrary as plain text (frontmatter stripped, so the cell is editable).
function ingestSkillFromUrl_(url, optKey, optTitle) {
  if (!_isAdminUser_()) { _toast_('Admin only.'); return; }
  if (!url) { _toast_('URL is required.'); return; }

  const rawUrl = _normalizeGithubRawUrl_(url);
  let resp;
  try {
    resp = UrlFetchApp.fetch(rawUrl, { muteHttpExceptions: true, followRedirects: true });
  } catch (e) {
    _toast_('Fetch failed: ' + (e.message || e));
    return;
  }
  if (resp.getResponseCode() !== 200) {
    _toast_('Fetch failed: HTTP ' + resp.getResponseCode() + ' for ' + rawUrl);
    return;
  }

  const hint = rawUrl.split('?')[0].split('/').pop() || '';
  let markdown;
  try {
    markdown = _extractSkillMarkdownFromBlob_(resp.getBlob().setName(hint || 'skill.bin'), hint);
  } catch (e) {
    _toast_('Could not read skill from URL: ' + (e.message || e));
    return;
  }

  try {
    const written = _writeIngestedSkill_(markdown, rawUrl, optKey, optTitle);
    _toast_('Skill "' + written.key + '" ' + written.result +
      (written.outputs ? ' (outputs: ' + written.outputs + ')' : '') +
      ' from ' + rawUrl + '.');
    return written;
  } catch (e) {
    _toast_('Ingest failed: ' + (e.message || e));
  }
}

// ingestSkillFromDrive_ — zip (or md) from a Drive Skills folder / file.
function ingestSkillFromDrive_(fileId, optKey, optTitle) {
  if (!_isAdminUser_()) { _toast_('Admin only.'); return; }
  const id = _extractDriveFileId_(fileId);
  if (!id) { _toast_('Drive file ID is required.'); return; }

  let file;
  try {
    file = DriveApp.getFileById(id);
  } catch (e) {
    _toast_('Could not open Drive file: ' + (e.message || e));
    return;
  }

  let markdown;
  try {
    markdown = _extractSkillMarkdownFromBlob_(file.getBlob(), file.getName());
  } catch (e) {
    _toast_('Could not read skill from Drive: ' + (e.message || e));
    return;
  }

  const source = 'drive:' + id + ' (' + file.getName() + ')';
  try {
    const written = _writeIngestedSkill_(markdown, source, optKey, optTitle);
    _toast_('Skill "' + written.key + '" ' + written.result +
      (written.outputs ? ' (outputs: ' + written.outputs + ')' : '') +
      ' from Drive file ' + file.getName() + '.');
    return written;
  } catch (e) {
    _toast_('Ingest failed: ' + (e.message || e));
  }
}

function _extractDriveFileId_(urlOrId) {
  const s = String(urlOrId || '').trim();
  const m = s.match(/\/file\/d\/([a-zA-Z0-9_-]+)/) || s.match(/[?&]id=([a-zA-Z0-9_-]+)/);
  if (m) return m[1];
  return s;
}

function menuIngestSkillFromUrl_() {
  if (!_isAdminUser_()) { _toast_('Admin only.'); return; }
  const ui = SpreadsheetApp.getUi();
  const urlResp = ui.prompt(
    'Import Skill from URL (1/2)',
    'Paste a GitHub raw / blob URL, a direct .md URL, or a .zip URL for a Claude-style skill.',
    ui.ButtonSet.OK_CANCEL
  );
  if (urlResp.getSelectedButton() !== ui.Button.OK) return;
  const url = urlResp.getResponseText().trim();
  if (!url) { _toast_('No URL given.'); return; }

  const keyResp = ui.prompt(
    'Import Skill from URL (2/2)',
    'Short key (lowercase-with-dashes). Blank = use the YAML `name:` field:',
    ui.ButtonSet.OK_CANCEL
  );
  if (keyResp.getSelectedButton() !== ui.Button.OK) return;
  ingestSkillFromUrl_(url, keyResp.getResponseText().trim());
}

function menuIngestSkillFromDrive_() {
  if (!_isAdminUser_()) { _toast_('Admin only.'); return; }
  const ui = SpreadsheetApp.getUi();
  const idResp = ui.prompt(
    'Import Skill from Drive (1/2)',
    'Paste a Drive file URL or ID for a .zip or .md skill file.',
    ui.ButtonSet.OK_CANCEL
  );
  if (idResp.getSelectedButton() !== ui.Button.OK) return;
  const fileId = idResp.getResponseText().trim();
  if (!fileId) { _toast_('No file given.'); return; }

  const keyResp = ui.prompt(
    'Import Skill from Drive (2/2)',
    'Short key (lowercase-with-dashes). Blank = use the YAML `name:` field:',
    ui.ButtonSet.OK_CANCEL
  );
  if (keyResp.getSelectedButton() !== ui.Button.OK) return;
  ingestSkillFromDrive_(fileId, keyResp.getResponseText().trim());
}

// Imports one skill from a URL (typically a GitHub raw SKILL.md or similar
// instructions file). key is how agents reference it in skill_refs; title is
// just for display. Safe to re-run -- upserts by key.
function importSkillFromGithub_(url, key, title) {
  if (!_isAdminUser_()) { _toast_('Admin only.'); return; }
  if (!url || !key) { _toast_('URL and key are required.'); return; }

  const rawUrl = _normalizeGithubRawUrl_(url);
  let resp;
  try {
    resp = UrlFetchApp.fetch(rawUrl, { muteHttpExceptions: true });
  } catch (e) {
    _toast_('Fetch failed: ' + (e.message || e));
    return;
  }
  if (resp.getResponseCode() !== 200) {
    _toast_('Fetch failed: HTTP ' + resp.getResponseCode() + ' for ' + rawUrl);
    return;
  }

  const instructions = resp.getContentText().slice(0, 20000); // hard ceiling on any single imported skill
  const result = _upsertSkillLibraryRow_(key.trim().toLowerCase(), title || key, instructions, rawUrl);
  _toast_('Skill "' + key + '" ' + result + ' from ' + rawUrl + '.');
}

function menuImportSkillFromGithub_() {
  if (!_isAdminUser_()) { _toast_('Admin only.'); return; }
  const ui = SpreadsheetApp.getUi();

  const urlResp = ui.prompt(
    'Import Skill from GitHub (1/3)',
    'Paste a GitHub URL to a SKILL.md (or any plain-text instructions file).\n' +
    'A normal github.com/.../blob/... link works -- it gets converted to raw automatically.',
    ui.ButtonSet.OK_CANCEL
  );
  if (urlResp.getSelectedButton() !== ui.Button.OK) return;
  const url = urlResp.getResponseText().trim();
  if (!url) { _toast_('No URL given.'); return; }

  const keyResp = ui.prompt(
    'Import Skill from GitHub (2/3)',
    'Short key for this skill (lowercase, no spaces -- this is what you\'ll put in an agent\'s skill_refs column), e.g. "stride-threat-model":',
    ui.ButtonSet.OK_CANCEL
  );
  if (keyResp.getSelectedButton() !== ui.Button.OK) return;
  const key = keyResp.getResponseText().trim();
  if (!key) { _toast_('No key given.'); return; }

  const titleResp = ui.prompt('Import Skill from GitHub (3/3)', 'Display title (blank = use the key):', ui.ButtonSet.OK_CANCEL);
  if (titleResp.getSelectedButton() !== ui.Button.OK) return;
  const title = titleResp.getResponseText().trim();

  importSkillFromGithub_(url, key, title);
}

function menuListSkillLibrary_() {
  if (!_isAdminUser_()) { _toast_('Admin only.'); return; }
  const sheet = _getSkillLibrarySheet_();
  if (sheet.getLastRow() < 2) { _toast_('SkillLibrary is empty. Import one via Skill Library -> Import Skill from GitHub / URL / Drive.'); return; }
  const meta = _skillLibraryHeaderMap_(sheet);
  const data = sheet.getDataRange().getValues();
  const lines = [];
  for (let i = 1; i < data.length; i++) {
    const key = meta.map.key !== undefined ? String(data[i][meta.map.key] || '') : '';
    if (!key) continue;
    const title = meta.map.title !== undefined ? String(data[i][meta.map.title] || '') : '';
    const outputs = meta.map.outputs !== undefined ? String(data[i][meta.map.outputs] || '') : '';
    lines.push(key + ' -- ' + title + (outputs ? ' [' + outputs + ']' : ''));
  }
  SpreadsheetApp.getUi().alert(
    'Skill Library (' + lines.length + ')',
    lines.join('\n') +
      '\n\nPersona agents: put keys in skill_refs.\nSkill agents: set Agents.agent_type=skill and Agents.default_skill to a key (or have the user name it).',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

// Reads and joins the instructions for the given comma-separated skill_refs
// string. Mirrors _getKbForSlug_'s shape/cap so it's a familiar pattern.
function _getSkillLibraryText_(skillRefsField) {
  const keys = String(skillRefsField || '').split(',').map(function (s) { return s.trim().toLowerCase(); }).filter(Boolean);
  if (!keys.length) return '';

  const sheet = _getSs_().getSheetByName('SkillLibrary');
  if (!sheet) return '';
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return '';

  const headers = data[0].map(function (h) { return String(h).trim().toLowerCase(); });
  const iKey = headers.indexOf('key');
  const iTitle = headers.indexOf('title');
  const iInstr = headers.indexOf('instructions');
  if (iKey < 0 || iInstr < 0) return '';

  const out = [];
  let used = 0;
  for (let i = 1; i < data.length; i++) {
    const rowKey = String(data[i][iKey] || '').trim().toLowerCase();
    if (keys.indexOf(rowKey) < 0) continue;
    const title = iTitle >= 0 ? String(data[i][iTitle] || '') : rowKey;
    const body = String(data[i][iInstr] || '');
    if (!body) continue;
    const chunk = '--- ' + title + ' ---\n' + body;
    if (used + chunk.length > MAX_SKILL_LIBRARY_CHARS) break;
    out.push(chunk);
    used += chunk.length;
  }
  return out.join('\n\n');
}

// One-time: adds Agents.skill_refs and Agents.kb_source if missing. Mirrors
// migrateAgentsAddSkillsTone_'s pattern -- non-destructive, existing rows
// just get blank values until an admin sets them.
function migrateAgentsAddSkillRefs_() {
  if (!_isAdminUser_()) { _toast_('Admin only.'); return; }
  const sheet = _getSs_().getSheetByName('Agents');
  if (!sheet) { _toast_('No Agents tab found.'); return; }
  const lastCol = Math.max(sheet.getLastColumn(), 1);
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function (h) { return String(h || '').trim().toLowerCase(); });

  const toAdd = [];
  if (headers.indexOf('skill_refs') < 0) toAdd.push('skill_refs');
  if (headers.indexOf('kb_source') < 0) toAdd.push('kb_source');

  if (!toAdd.length) { _toast_('Agents already has skill_refs and kb_source.'); return; }

  sheet.getRange(1, lastCol + 1, 1, toAdd.length).setValues([toAdd]);
  sheet.getRange(1, lastCol + 1, 1, toAdd.length).setFontWeight('bold');
  _toast_('Added ' + toAdd.join(', ') + ' to Agents. skill_refs = SkillLibrary keys to attach; kb_source = free-text label of where that agent\'s KB content came from.');
}

// Seeds one real, working example so the library isn't empty on first look --
// the existing STRIDE builtin pack (_getBuiltinPack in 10_admin_ingest.js),
// reformatted as a SkillLibrary entry. Safe to re-run (upserts by key).
function seedExampleSkill_() {
  if (!_isAdminUser_()) { _toast_('Admin only.'); return; }
  const strideText = _getBuiltinPack('stride');
  const result = _upsertSkillLibraryRow_('stride-threat-model', 'STRIDE Threat Modeling', strideText, '(built-in, not imported)');
  _toast_('Example skill "stride-threat-model" ' + result + '. Attach it to an agent via that agent\'s skill_refs column.');
}
