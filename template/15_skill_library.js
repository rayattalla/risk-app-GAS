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
 * SkillLibrary tab: key | title | instructions | source_url | added_at
 * Agents.skill_refs: comma list of SkillLibrary keys this agent should load.
 */

var SKILL_LIBRARY_HEADERS = ['key', 'title', 'instructions', 'source_url', 'added_at'];
var MAX_SKILL_LIBRARY_CHARS = 6000; // hard ceiling on injected skill-library context, mirrors MAX_KB_CHARS

function _getSkillLibrarySheet_() {
  return _getOrCreateTab_('SkillLibrary', SKILL_LIBRARY_HEADERS);
}

function _upsertSkillLibraryRow_(key, title, instructions, sourceUrl) {
  const sheet = _getSkillLibrarySheet_();
  const data = sheet.getDataRange().getValues();
  const row = [key, title, instructions, sourceUrl || '', new Date()];
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === String(key).trim().toLowerCase()) {
      sheet.getRange(i + 1, 1, 1, row.length).setValues([row]);
      return 'updated';
    }
  }
  sheet.appendRow(row);
  return 'inserted';
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
  if (sheet.getLastRow() < 2) { _toast_('SkillLibrary is empty. Import one via Admin + Ingest -> Import Skill from GitHub.'); return; }
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues(); // key, title
  const lines = data.map(function (r) { return r[0] + ' -- ' + r[1]; });
  SpreadsheetApp.getUi().alert(
    'Skill Library (' + lines.length + ')',
    lines.join('\n') + '\n\nAttach to an agent by putting one or more keys (comma-separated) in that agent\'s skill_refs column.',
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
