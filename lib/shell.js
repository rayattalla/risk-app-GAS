/**
 * Template — Shell (shell.js)
 *
 * Inlined architecture (no external library):
 *   - This script owns: onOpen, doGet web app, menus, list/get/send RPCs,
 *     and all business logic inlined from previous library.
 *   - No external library dependency.
 *
 * doGet and hot RPCs are inlined here.
 * All constants in config.js.
 */

// ==========================================================================
// onOpen — auto-fires when spreadsheet opens
// ==========================================================================

function onOpen() {
  const ui   = SpreadsheetApp.getUi();
  const menu = ui.createMenu(PROGRAM_SHORT);

  menu.addItem('Open Web App', 'openWebApp');

  let isAdmin = false;
  try { isAdmin = _shellIsAdmin_(); } catch (e) {}

  if (isAdmin) {
    menu.addSeparator();
    menu.addSubMenu(
      ui.createMenu('Setup & Config')
        .addItem('Run Setup',       'menuSetup')
        .addItem('Set Web App URL', 'menuSetWebAppUrl')
        .addItem('Rebuild Menu (debug)', 'forceRebuildMenu')
        .addItem('Ship Checklist',  'menuShipChecklist')
    );
    menu.addSubMenu(
      ui.createMenu('Admin + Ingest')
        .addItem('Seed pilot agents + KB', 'menuSeedPilotAgentsKB')
        .addItem('Seed all 15 Risk agents', 'seedAllRiskAgents_')
        .addItem('Ingest URL / paste / Drive file', 'menuIngest')
        .addItem('Ingest built-in ITS packs', 'menuIngestBuiltin')
        .addItem('Migrate KB tabs to single KB', 'migrateKbToSingleSheet_')
        .addItem('Cleanup old KB tabs (shipping)', 'cleanupOldKbTabs_')
        .addItem('Seed Full Reference KB (8 core docs)', 'seedReferenceKB_')
        .addItem('Backfill Agent Tone (from personas)', 'backfillAgentTone_')
        .addItem('Recommend Agent Skills (webpage/search/email)', 'upgradeAgentSkills_')
        .addItem('Open Admin', 'openAdmin')
        .addItem('Set OpenRouter LLM Key', 'menuSetOpenRouterKey')
        .addItem('Set Web Search API Key (Serper)', 'menuSetSearchKey')
        .addItem('Migrate Agents (add skills/tone columns)', 'migrateAgentsAddSkillsTone_')
    );
    menu.addSubMenu(
      ui.createMenu('Automations')
        .addItem('Add Automation', 'menuAddAutomation')
        .addItem('List Automations', 'menuListAutomations')
        .addItem('Install Hourly Trigger', 'installAutomationsTrigger_')
        .addItem('Run Due Automations Now', 'runDueAutomations_')
    );
    menu.addSubMenu(
      ui.createMenu('Memory')
        .addItem('Install Hourly Memory Trigger', 'installMemoryTrigger_')
        .addItem('Run Memory Distillation Now', 'runMemoryDistillation_')
        .addItem('Clear ALL Memory (compliance)', 'menuClearAllMemory_')
    );
    menu.addSubMenu(
      ui.createMenu('Access Control')
        .addItem('List Users', 'menuListAccess')
    );
    menu.addSubMenu(
      ui.createMenu('Test Mode')
        .addItem('Toggle Test Mode', 'menuToggleTestMode')
        .addItem('Check Modes',      'menuCheckModes')
    );
    menu.addSubMenu(
      ui.createMenu('Diagnostics')
        .addItem('Run System Tests', 'menuRunTests')
    );
  }

  menu.addToUi();
}

function forceRebuildMenu() {
  if (!_shellIsAdmin_()) { _toast_('Admin only.'); return; }
  onOpen();
  _toast_('Menu rebuilt. Reload the spreadsheet tab if menu still missing.');
}

function _shellIsAdmin_() {
  try {
    const me = String(Session.getActiveUser().getEmail() || '').toLowerCase();
    const admin = (typeof ADMIN_EMAIL !== 'undefined' ? ADMIN_EMAIL : '').toLowerCase();
    if (me === admin) return true;
    return (typeof isCurrentUserAdmin === 'function') ? isCurrentUserAdmin() : false;
  } catch (e) { return false; }
}

function _toast_(msg) {
  SpreadsheetApp.getActiveSpreadsheet().toast(msg || 'Done.', PROGRAM_SHORT, 8);
}

// ==========================================================================
// MENU ACTIONS — delegate to library
// ==========================================================================

function menuSetup() {
  if (!_shellIsAdmin_()) { _toast_('Admin only.'); return; }
  _toast_('Running setup…');
  try { _toast_(String(setup())); }
  catch (e) { _toast_('Setup failed: ' + e.message); }
}

function menuSetWebAppUrl() {
  if (!_shellIsAdmin_()) { _toast_('Admin only.'); return; }
  const ui = SpreadsheetApp.getUi();
  let current = '';
  try { current = getWebAppUrl(); } catch (e) {}
  const resp = ui.prompt(
    'Set Web App URL',
    'Paste the /exec URL from Deploy → Manage deployments.\n\nCurrent: ' + (current || '(not set)'),
    ui.ButtonSet.OK_CANCEL
  );
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  const url = resp.getResponseText().trim();
  (typeof setWebAppUrl === 'function') ? setWebAppUrl(url) : null;
  _toast_(url ? 'Web app URL saved.' : 'URL cleared.');
}

function menuListAccess() {
  if (!_shellIsAdmin_()) { _toast_('Admin only.'); return; }
  const res = (typeof listAccess === 'function') ? listAccess() : {success: false};
  if (!res || !res.success) { _toast_('Failed: ' + (res && res.error)); return; }
  const lines = (res.users || []).map(u => u.email + ' → ' + u.role);
  SpreadsheetApp.getUi().alert(
    'Access Control (' + lines.length + ' users)',
    lines.length ? lines.join('\n') : '(no explicit rows — domain users default to Editor)',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function menuToggleTestMode() {
  if (!_shellIsAdmin_()) { _toast_('Admin only.'); return; }
  if ((typeof isTestMode === 'function') ? isTestMode() : false) {
    (typeof disableTestMode === 'function') ? disableTestMode() : null;
    _toast_('Test mode OFF — real emails will be sent.');
  } else {
    (typeof enableTestMode === 'function') ? enableTestMode() : null;
    _toast_('Test mode ON — emails redirect to ' + ADMIN_EMAIL);
  }
}

function menuCheckModes() {
  (typeof checkModes === 'function') ? checkModes() : null;
  _toast_('Modes logged — check Executions log.');
}

function menuRunTests() {
  if (!_shellIsAdmin_()) { _toast_('Admin only.'); return; }
  _toast_('Running tests… (may take 10-30s)');
  // Add your own test runner call here:
  // const res = 'tests removed for pilot';
  _toast_('No tests registered yet. Add a runAllTests() to the library.');
}

function openWebApp() {
  let url = '';
  try { url = getWebAppUrl(); } catch (e) {}
  if (!url) { try { url = ScriptApp.getService().getUrl(); } catch (e) {} }
  if (!url) {
    SpreadsheetApp.getUi().alert(
      'No web app URL',
      'Deploy the web app first (Deploy → Manage deployments → New deployment → Web app), then run ' + PROGRAM_SHORT + ' → Setup & Config → Set Web App URL.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }
  const safe = String(url).replace(/"/g, '%22');
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput('<script>window.open("' + safe + '","_blank");google.script.host.close();</script>')
      .setWidth(300).setHeight(50),
    'Opening ' + PROGRAM_NAME + '…'
  );
}

// ==========================================================================
// doGet — INLINED (per patterns §3 — do not move to library)
// ==========================================================================

function doGet(e) {
  e = e || {};
  const params = e.parameter || {};
  const page = String(params.page || '').toLowerCase().trim();
  const agentSlug = String(params.agent || '').trim().replace(/[^a-z0-9-]/gi, '').toLowerCase();

  // Access gate for pilot: domain-based only (no getMyAccess; inlined, no library).
  // Non-admins @lausd.net get Editor role via fallback; admins get Admin.
  const me = (Session.getActiveUser().getEmail() || '').toLowerCase().trim();
  const adminEmail = (typeof ADMIN_EMAIL !== 'undefined' ? ADMIN_EMAIL : '').toLowerCase().trim();
  let role = null;
  let isAdmin = false;
  if (me === adminEmail) {
    role = 'Admin';
    isAdmin = true;
  } else if (me.endsWith('@lausd.net')) {
    role = 'Editor';
  }
  if (!role) return _denyPage_();

  if (page === 'admin') {
    if (me !== adminEmail) return _admin403Page_();
    // serve admin.html as full page (no heavy inject needed)
    const adminHtml = HtmlService.createHtmlOutputFromFile('admin')
      .setTitle('Admin — ' + PROGRAM_SHORT)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
    return adminHtml;
  }

  // normal chat UI
  // Web app URL
  let webAppUrl = '';
  try { webAppUrl = getWebAppUrl(); } catch (e2) {}
  if (!webAppUrl) {
    try { webAppUrl = ScriptApp.getService().getUrl() || ''; } catch (e3) {}
  }

  // Use template as required
  const template = HtmlService.createTemplateFromFile('index');
  template.SHELL_VERSION = SHELL_VERSION;
  template.WEBAPP_URL = webAppUrl;
  template.ORG_NAME = ORG_NAME;
  template.ORG_SHORT = ORG_SHORT;
  template.PROGRAM_NAME = PROGRAM_NAME;
  template.PROGRAM_SHORT = PROGRAM_SHORT;
  template.AGENT_SLUG = agentSlug;
  template.ADMIN_EMAIL = ADMIN_EMAIL;

  const title = agentSlug ? (PROGRAM_SHORT + ' — ' + agentSlug) : PROGRAM_NAME;
  return template.evaluate()
    .setTitle(title + ' — ' + ORG_SHORT)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function _denyPage_() {
  let who = '';
  try { who = Session.getActiveUser().getEmail() || ''; } catch (e) {}
  const safe = String(who).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  const grad = 'linear-gradient(90deg,#F47B20,#ED1C24,#0072CE,#00A3E0,#5B6CB0)';
  const html = '<!doctype html><html><head><meta charset="utf-8"><title>Access Restricted</title>'
    + '<style>body{font-family:-apple-system,Segoe UI,Roboto,sans-serif;background:#F5F7FA;margin:0;padding:0}'
    + '.box{max-width:520px;margin:10vh auto;background:#fff;border-radius:12px;box-shadow:0 4px 16px rgba(0,0,0,.08);padding:2rem;text-align:center}'
    + '.bar{height:6px;background:' + grad + ';border-radius:12px 12px 0 0;margin:-2rem -2rem 1.5rem}'
    + 'h1{font-size:1.25rem;color:#1A237E;margin-bottom:.4rem}'
    + 'p{color:#5F6368;font-size:.92rem;line-height:1.6;margin:.4rem 0}'
    + '.em{font-family:monospace;background:#F0F4F8;padding:.1rem .4rem;border-radius:4px}'
    + 'a{color:#0072CE}</style></head><body>'
    + '<div class="box"><div class="bar"></div>'
    + '<h1>&#128274; Access Restricted</h1>'
    + '<p>' + PROGRAM_NAME + ' is limited to authorized personnel.</p>'
    + '<p>Signed in as <span class="em">' + safe + '</span>.</p>'
    + '<p>To request access, contact <a href="mailto:' + NOTIFY_EMAIL + '">' + NOTIFY_EMAIL + '</a>.</p>'
    + '</div></body></html>';
  return HtmlService.createHtmlOutput(html)
    .setTitle(PROGRAM_NAME + ' — Access Restricted')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function _admin403Page_() {
  let who = '';
  try { who = Session.getActiveUser().getEmail() || ''; } catch (e) {}
  const safe = String(who).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  const grad = 'linear-gradient(90deg,#F47B20,#ED1C24,#0072CE,#00A3E0,#5B6CB0)';
  const html = '<!doctype html><html><head><meta charset="utf-8"><title>Admin Access Required</title>'
    + '<style>body{font-family:-apple-system,Segoe UI,Roboto,sans-serif;background:#F5F7FA;margin:0;padding:0}'
    + '.box{max-width:520px;margin:10vh auto;background:#fff;border-radius:12px;box-shadow:0 4px 16px rgba(0,0,0,.08);padding:2rem;text-align:center}'
    + '.bar{height:6px;background:' + grad + ';border-radius:12px 12px 0 0;margin:-2rem -2rem 1.5rem}'
    + 'h1{font-size:1.25rem;color:#1A237E;margin-bottom:.4rem}'
    + 'p{color:#5F6368;font-size:.92rem;line-height:1.6;margin:.4rem 0}'
    + '.em{font-family:monospace;background:#F0F4F8;padding:.1rem .4rem;border-radius:4px}'
    + 'a{color:#0072CE}</style></head><body>'
    + '<div class="box"><div class="bar"></div>'
    + '<h1>&#128274; Admin Only</h1>'
    + '<p>This page is restricted to administrators.</p>'
    + '<p>Signed in as <span class="em">' + safe + '</span>.</p>'
    + '<p>Contact the platform admin if you need admin access.</p>'
    + '</div></body></html>';
  return HtmlService.createHtmlOutput(html)
    .setTitle(PROGRAM_NAME + ' — Admin Only')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ==========================================================================
// WRAPPERS — read path (inlined; no library)
// ==========================================================================

function getConfigData() {
  return {
    version: SHELL_VERSION,
    orgName: ORG_NAME,
    orgShort: ORG_SHORT,
    programName: PROGRAM_NAME,
    programShort: PROGRAM_SHORT,
    adminEmail: ADMIN_EMAIL
  };
}

function getWebAppUrl() {
  try { return ScriptApp.getService().getUrl() || ''; } catch (e) { return ''; }
}

function isCurrentUserAdmin() { 
  const me = (Session.getActiveUser().getEmail() || '').toLowerCase().trim();
  const admin = (typeof ADMIN_EMAIL !== 'undefined' ? ADMIN_EMAIL : '').toLowerCase().trim();
  return me === admin;
}

// listAccess, grantAccess, revokeAccess, setWebAppUrl deleted (unused in pilot; were stubs after library removal)

// setup implemented below (real, not recursive)

// Agent platform thin wrappers (MVP) — see sendMessage etc. below

function setup() {
  if (!_shellIsAdmin_()) return { success: false, error: 'Admin only.' };
  // _runBasicSetup is defined in 10_admin_ingest.js (inlined)
  if (typeof _runBasicSetup === 'function') {
    _runBasicSetup();
    return { success: true };
  }
  return { success: false, error: 'setup helper not found' };
}

// Agent platform thin wrappers (MVP)
function sendMessage(slug, history, text) {
  var started = Date.now();
  var email = '';
  try { email = Session.getActiveUser().getEmail() || ''; } catch (e) {}

  if (!slug) return { ok: false, error: 'Missing agent slug' };

  try {
    var reply = chatWithAgent_(slug, history, text, email);
    _logChat_({
      email: email, slug: slug, model: '', prompt: text, response: reply,
      status: 'ok', durationMs: Date.now() - started
    });
    return { ok: true, reply: reply };
  } catch (e) {
    _logChat_({
      email: email, slug: slug, model: '', prompt: text, response: '',
      status: 'error', error: e.message || String(e),
      durationMs: Date.now() - started
    });
    return { ok: false, error: e.message || String(e) };
  }
}

// Email skill (v1.2.0) — sends the current chat transcript to the requesting
// user's own address. No arbitrary recipient from the client, to avoid the
// web app being used to relay mail to third parties.
function emailTranscript(slug, history) {
  var me = '';
  try { me = Session.getActiveUser().getEmail() || ''; } catch (e) {}
  if (!me) return { ok: false, error: 'Could not determine your email.' };
  try {
    _sendTranscriptEmail_(me, slug, history || []);
    return { ok: true };
  } catch (e) {
    return { ok: false, error: e.message || String(e) };
  }
}

// Self-service memory clear — a user can only clear their own memory note
// for one agent, never anyone else's (mirrors emailTranscript's self-scoping).
function clearMyMemory(slug) {
  var me = '';
  try { me = Session.getActiveUser().getEmail() || ''; } catch (e) {}
  if (!me) return { ok: false, error: 'Could not determine your email.' };
  try {
    _clearMemoryNote_(me, slug);
    return { ok: true };
  } catch (e) {
    return { ok: false, error: e.message || String(e) };
  }
}

function listAgents() {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName('Agents');
    if (!sheet) return { ok: true, agents: [] };
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const map = {};
    headers.forEach((h, i) => { if (h) map[String(h).trim().toLowerCase()] = i; });
    const data = sheet.getDataRange().getValues();
    const out = [];
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
    return { ok: true, agents: out };
  } catch (e) {
    return { ok: false, error: e.message || String(e) };
  }
}

function getAgent(slug) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName('Agents');
    if (!sheet) return { ok: false, error: 'No Agents sheet' };
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const map = {};
    headers.forEach((h, i) => { if (h) map[String(h).trim().toLowerCase()] = i; });
    const data = sheet.getDataRange().getValues();
    const target = String(slug || '').trim().toLowerCase();
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (String(row[map.slug] || '').trim().toLowerCase() === target) {
        return {
          ok: true,
          agent: {
            slug: row[map.slug] || '',
            name: row[map.name] || '',
            status: String(row[map.status] || '').toLowerCase(),
            org: row[map.org] || 'LAUSD',
            model: row[map.model] || '',
            personality: row[map.personality] || '',
            kb_tags: row[map.kb_tags] || '',
            notes: row[map.notes] || '',
            // New (v1.2.0): per-agent tone + skills allowlist. Both optional —
            // absent 'skills' column/value means legacy default 'kb' only,
            // so existing agent rows keep working unmigrated.
            tone: map.tone !== undefined ? (row[map.tone] || '') : '',
            skills: map.skills !== undefined ? (row[map.skills] || '') : ''
          }
        };
      }
    }
    return { ok: false, error: 'Not found' };
  } catch (e) {
    return { ok: false, error: e.message || String(e) };
  }
}

// Add your project-specific write wrappers here:
// function createRecord(payload)   { /* inlined or removed for pilot */ }
// function updateRecord(id, patch) { /* inlined or removed for pilot */ }

// ==========================================================================
// DIAGNOSTICS — safe read-only visibility into sheet + RPC state
// ==========================================================================

function runDiagnostics() {
  var out = {
    ok: true,
    sheetId: SHEET_ID,
    version: SHELL_VERSION,
    agents: { found: false, headers: [], rowCount: 0, sample: [], slugColumn: null },
    kb: { found: false, headers: [], rowCount: 0, sample: [] },
    agentLookup: { slug: '', match: null, error: '' },
    errors: []
  };
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var tabs = ss.getSheets().map(function(s) { return s.getName(); });
    out.tabs = tabs;

    var aSheet = ss.getSheetByName('Agents');
    if (aSheet) {
      out.agents.found = true;
      var aData = aSheet.getDataRange().getValues();
      out.agents.rowCount = Math.max(0, aData.length - 1);
      if (aData.length > 0) {
        out.agents.headers = (aData[0] || []).map(function(h) { return String(h || '').trim(); });
        out.agents.slugColumn = out.agents.headers.findIndex(function(h) { return h.toLowerCase() === 'slug'; });
        out.agents.sample = aData.slice(1, 4);
      }
    }

    var kSheet = ss.getSheetByName('KB');
    if (kSheet) {
      out.kb.found = true;
      var kData = kSheet.getDataRange().getValues();
      out.kb.rowCount = Math.max(0, kData.length - 1);
      if (kData.length > 0) {
        out.kb.headers = (kData[0] || []).map(function(h) { return String(h || '').trim(); });
        out.kb.sample = kData.slice(1, 4);
      }
    }

    if (out.agents.found && out.agents.slugColumn >= 0 && out.agents.sample.length) {
      var probe = out.agents.sample[0][out.agents.slugColumn] || '';
      out.agentLookup.slug = String(probe).trim();
      var g = getAgent(out.agentLookup.slug);
      out.agentLookup.match = g.ok ? g.agent : null;
      out.agentLookup.error = g.ok ? '' : (g.error || '');
    }
  } catch (e) {
    out.ok = false;
    out.errors.push(e.message || String(e));
  }
  return out;
}

// ==========================================================================
// SHIP CHECKLIST
// ==========================================================================

function menuShipChecklist() {
  if (!_shellIsAdmin_()) { _toast_('Admin only.'); return; }
  const scriptId = ScriptApp.getScriptId();
  const msg =
    `🚀 SHIP CHECKLIST — ${PROGRAM_NAME}\n\n` +
    `Script ID: ${scriptId}\n\n` +
    `1. Register in App Registry (Notion)\n` +
    `   Paste this Script ID, set Status + Security Review\n\n` +
    `2. Complete SECURITY.md checklist before any push\n\n` +
    `3. Verify .clasp.json has the correct scriptId\n\n` +
    `4. Pin real library version (turn off developmentMode)\n\n` +
    `5. Disable Test Mode for production\n\n` +
    `6. Run 'Cleanup old KB tabs (shipping)' to remove legacy KB_* tabs\n\n` +
    `7. Run 'Migrate Agents (add skills/tone columns)' if not already done\n\n` +
    `8. Set a Web Search API key if any agent uses the 'search' skill\n\n` +
    `9. Run 'Install Hourly Trigger' if any Automations rows exist\n\n` +
    `10. Run 'Install Hourly Memory Trigger' if any agent has 'memory' skill\n\n` +
    `Registry: https://app.notion.com/p/f9eff5970a6d4281a259c22185ca7c1d`;
  SpreadsheetApp.getUi().alert('Ship Checklist', msg, SpreadsheetApp.getUi().ButtonSet.OK);
}
