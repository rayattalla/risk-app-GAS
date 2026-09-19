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
    // Day-to-day operations first. "When do I use this" for each submenu:
    // Agents      -> you added/changed an agent
    // Knowledge Base -> you have new content to feed an agent
    // Skill Library  -> you want to attach a reusable procedure to an agent
    // Automations / Memory -> ongoing scheduled/behavioral features
    // Diagnostics -> something looks wrong
    // Setup (one-time) -> initial bring-up only, safe to ignore day to day
    menu.addSubMenu(
      ui.createMenu('Agents')
        .addItem('Add New Agent…', 'menuAddNewAgent_')
        .addItem('Backfill Agent Tone (from personas)', 'backfillAgentTone_')
        .addItem('Recommend Agent Skills (webpage/search/email)', 'upgradeAgentSkills_')
        .addItem('Run Schema Migrations (safe to re-run)', 'runAllAgentMigrations_')
    );
    menu.addSubMenu(
      ui.createMenu('Ingest')
        .addItem('Bulletins (scrape LAUSD bulletin pages)', 'menuIngestBulletins')
        .addItem('Procurements (scrape RFP repository)', 'menuIngestProcurements')
        .addSeparator()
        .addItem('Ingest URL / paste / Drive file', 'menuIngest')
        .addItem('Ingest Drive Folder (PDF/txt/md/Docs)', 'menuIngestDriveFolder_')
        .addItem('Ingest built-in ITS packs', 'menuIngestBuiltin')
        .addSeparator()
        .addItem('Deduplicate KB rows (by id)', 'menuDedupKb_')
    );
    menu.addSubMenu(
      ui.createMenu('Skill Library')
        .addItem('Import Skill from GitHub', 'menuImportSkillFromGithub_')
        .addItem('Import Skill from URL (zip or md)', 'menuIngestSkillFromUrl_')
        .addItem('Import Skill from Drive (zip or md)', 'menuIngestSkillFromDrive_')
        .addItem('List Skill Library', 'menuListSkillLibrary_')
        .addItem('Seed skill agent', 'menuSeedSkillAgent_')
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
      ui.createMenu('Diagnostics')
        .addItem('Run System Tests', 'menuRunTests')
    );
    // RiskAI v2.0 upgrade pack (2026-09-18): separate from day-to-day
    // operations on purpose. Step 1 (setup) only ever creates tabs and
    // never touches Agents. Step 2 (promote) mutates the live Agents tab --
    // always backs itself up first, but review the Agents_v2_staging diff
    // before running it, per riskai-deploy/CLAUDE.md.
    menu.addSubMenu(
      ui.createMenu('RiskAI v2.0 Upgrade')
        .addItem('1. Run Setup Upgrade Pack (creates new tabs only)', 'setupUpgradePack')
        .addItem('2. Promote Agents to v2 (review staging first!)', 'promoteAgentsV2')
    );
    // Initial bring-up only -- run once when standing this project up, or
    // when deliberately re-seeding canned content. Not part of day-to-day
    // agent/KB operations, which is why it's last and separate.
    menu.addSubMenu(
      ui.createMenu('Setup (one-time)')
        .addItem('Run Setup', 'menuSetup')
        .addItem('Rebuild Menu (debug)', 'forceRebuildMenu')
        .addItem('Ship Checklist', 'menuShipChecklist')
        .addItem('Set OpenRouter LLM Key', 'menuSetOpenRouterKey')
        .addItem('Set Web Search API Key (Serper)', 'menuSetSearchKey')
        .addItem('Seed pilot agents + KB', 'menuSeedPilotAgentsKB')
        .addItem('Seed all Risk + Policy/HR agents', 'seedAllRiskAgents_')
        .addItem('Seed Full Reference KB (10 core docs)', 'seedReferenceKB_')
        .addItem('Seed Skills Catalog (implemented + proposed)', 'seedSkillsCatalog_')
        .addItem('Seed Example Skill (STRIDE)', 'seedExampleSkill_')
        .addItem('Seed skill agent', 'menuSeedSkillAgent_')
        .addItem('Migrate KB tabs to single KB', 'migrateKbToSingleSheet_')
        .addItem('Cleanup old KB tabs (shipping)', 'cleanupOldKbTabs_')
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
// AGENT MANAGEMENT — the two things you actually do routinely: add an
// agent, and make sure the Agents sheet has every column current code
// expects (skills/tone, skill_refs/kb_source, allow_upload/upload_folder_id).
// ==========================================================================

// Runs every Agents-column migration in sequence. Each one is independently
// idempotent (checks for its own columns before adding anything), so this
// is always safe to re-run -- the answer to "I added an agent by hand and
// something's blank" is usually just "run this."
function runAllAgentMigrations_() {
  if (!_isAdminUser_()) { _toast_('Admin only.'); return; }
  migrateAgentsAddSkillsTone_();
  migrateAgentsAddSkillRefs_();
  migrateAgentsAddUploadConfig_();
  migrateAgentsAddAgentType_();
  migrateSkillLibraryAddOutputs_();
  _toast_('Schema check complete.');
}

// Prompt wizard for creating a new agent row, mirroring menuAddAutomation's
// step-by-step pattern. Refuses to touch an existing slug -- edit that row
// directly in the Agents tab instead, this is for brand-new agents only.
function menuAddNewAgent_() {
  if (!_isAdminUser_()) { _toast_('Admin only.'); return; }
  const ui = SpreadsheetApp.getUi();

  const slugResp = ui.prompt('Add New Agent (1/6)', 'Slug (lowercase-with-dashes, e.g. "cloud-security"):', ui.ButtonSet.OK_CANCEL);
  if (slugResp.getSelectedButton() !== ui.Button.OK) return;
  const slug = slugResp.getResponseText().trim().toLowerCase().replace(/[^a-z0-9-]/g, '-');
  if (!slug) { _toast_('Slug required.'); return; }
  if (getAgent(slug).ok) { _toast_('Slug "' + slug + '" already exists — edit its row in the Agents tab directly instead of using this wizard.'); return; }

  const nameResp = ui.prompt('Add New Agent (2/6)', 'Display name (e.g. "Cloud Security"):', ui.ButtonSet.OK_CANCEL);
  if (nameResp.getSelectedButton() !== ui.Button.OK) return;
  const name = nameResp.getResponseText().trim() || slug;

  const personaResp = ui.prompt('Add New Agent (3/6)', 'Personality / system prompt — paste the full persona text (role, rules, tone-setting instructions):', ui.ButtonSet.OK_CANCEL);
  if (personaResp.getSelectedButton() !== ui.Button.OK) return;
  const personality = personaResp.getResponseText().trim();
  if (!personality) { _toast_('Personality is required — it\'s what makes this agent behave like this agent, not a generic assistant.'); return; }

  const toneResp = ui.prompt('Add New Agent (4/6)', 'Tone (optional, e.g. "Calm and procedural"):', ui.ButtonSet.OK_CANCEL);
  if (toneResp.getSelectedButton() !== ui.Button.OK) return;
  const tone = toneResp.getResponseText().trim();

  const skillsResp = ui.prompt('Add New Agent (5/6)', 'Skills, comma-separated (blank = "kb"). Options: kb, webpage, search, memory, email, diagram. See the SkillsCatalog tab for details.', ui.ButtonSet.OK_CANCEL);
  if (skillsResp.getSelectedButton() !== ui.Button.OK) return;
  const skills = skillsResp.getResponseText().trim() || 'kb';

  const modelResp = ui.prompt('Add New Agent (6/6)', 'Model override (blank = default):', ui.ButtonSet.OK_CANCEL);
  if (modelResp.getSelectedButton() !== ui.Button.OK) return;
  const model = modelResp.getResponseText().trim();

  runAllAgentMigrations_(); // make sure every column this write might target actually exists

  const sheet = _getSs_().getSheetByName('Agents');
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(function (h) { return String(h || '').trim().toLowerCase(); });
  const map = {};
  headers.forEach(function (h, i) { if (h) map[h] = i; });

  const row = new Array(headers.length).fill('');
  row[map.slug] = slug;
  row[map.name] = name;
  row[map.status] = 'on';
  row[map.org] = 'LAUSD';
  row[map.model] = model;
  row[map.personality] = personality;
  if (map.tone !== undefined) row[map.tone] = tone;
  if (map.skills !== undefined) row[map.skills] = skills;
  sheet.appendRow(row);

  _toast_('Added agent "' + slug + '".');
  SpreadsheetApp.getUi().alert(
    'Agent Added',
    'Created "' + name + '" (' + slug + ').\n\nOpen the bare web app URL as admin to get its shareable link, then hand that link to whoever should use it.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
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

function menuRunTests() {
  if (!_shellIsAdmin_()) { _toast_('Admin only.'); return; }
  _toast_('Running full diagnostic test… writing to Debug_Test tab');
  runFullDiagnosticTest_();
  _toast_('Done. See the Debug_Test tab.');
}

// Writes a full, plain-text diagnostic dump to a 'Debug_Test' tab so results
// can be copy/pasted out without needing to read Stackdriver logs. Exercises
// the exact same code path the web app uses (getSheetByName, header mapping,
// listAgents/getAgent), so whatever it reports is the real live behavior.
function runFullDiagnosticTest_() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let dbg = ss.getSheetByName('Debug_Test');
  if (dbg) { dbg.clear(); } else { dbg = ss.insertSheet('Debug_Test'); }
  dbg.getRange(1, 1, 1, 3).setValues([['Check', 'Detail', 'Result']]);
  dbg.setFrozenRows(1);

  const rows = [];
  function log(check, detail, result) {
    rows.push([check, detail, typeof result === 'string' ? result : JSON.stringify(result)]);
  }

  // Reveal invisible characters (non-breaking space, zero-width, etc.) that
  // .trim() would NOT catch, by showing raw char codes for a string.
  function codes(s) {
    s = String(s || '');
    const cc = [];
    for (let i = 0; i < s.length; i++) cc.push(s.charCodeAt(i));
    return s + '  [len=' + s.length + ' codes=' + cc.join(',') + ']';
  }

  try {
    log('Sheet ID', SHEET_ID, 'opening...');
    const tabs = ss.getSheets().map(s => s.getName());
    log('All tabs', '', tabs);

    const aSheet = ss.getSheetByName('Agents');
    log('Agents tab found', '', !!aSheet);

    if (aSheet) {
      const aData = aSheet.getDataRange().getValues();
      const headersRaw = aData[0] || [];
      log('Agents headers (raw)', '', headersRaw.map(h => codes(h)));

      const map = {};
      headersRaw.forEach((h, i) => { if (h) map[String(h).trim().toLowerCase()] = i; });
      log('Agents header map', '', map);
      log('Agents row count', '', Math.max(0, aData.length - 1));

      // Every row's slug/status, with hidden-char inspection, exactly as
      // getAgent()/listAgents() would read them.
      for (let i = 1; i < aData.length; i++) {
        const row = aData[i];
        const rawSlug = row[map.slug];
        const rawStatus = row[map.status];
        log('Row ' + i + ' slug (raw)', '', codes(rawSlug));
        log('Row ' + i + ' status (raw)', '', codes(rawStatus));
      }
    }

    log('--- listAgents() ---', '', '');
    const listRes = listAgents();
    log('listAgents() result', '', listRes);

    if (listRes && listRes.ok && listRes.agents) {
      listRes.agents.forEach(a => {
        log('--- getAgent(' + a.slug + ') ---', 'input slug (raw)', codes(a.slug));
        const g = getAgent(a.slug);
        log('getAgent(' + a.slug + ') result', '', g);
      });
    }

    const kSheet = ss.getSheetByName('KB');
    log('KB tab found', '', !!kSheet);
    if (kSheet) {
      const kData = kSheet.getDataRange().getValues();
      log('KB headers', '', (kData[0] || []).map(h => String(h || '').trim()));
      log('KB row count', '', Math.max(0, kData.length - 1));
    }

    const clSheet = ss.getSheetByName('ChatLog');
    log('ChatLog tab found', '', !!clSheet);
    if (clSheet) {
      log('ChatLog headers', '', clSheet.getRange(1, 1, 1, clSheet.getLastColumn()).getValues()[0]);
      log('ChatLog row count', '', Math.max(0, clSheet.getLastRow() - 1));
    }

    // District-data tabs (Schools/Enrollment/Jobs/Classifications/Budget/Staff)
    // -- read by the 'district-data' skill (17_district_data.js), one-time
    // snapshots with no sync. Report presence/headers/row count for each so
    // a missing/renamed tab shows up here instead of as a silent no-match.
    (typeof DISTRICT_DATA_TABS !== 'undefined' ? DISTRICT_DATA_TABS : []).forEach(function (tabName) {
      const dSheet = ss.getSheetByName(tabName);
      log(tabName + ' tab found', '', !!dSheet);
      if (dSheet && dSheet.getLastRow() > 0) {
        log(tabName + ' headers', '', dSheet.getRange(1, 1, 1, dSheet.getLastColumn()).getValues()[0]);
        log(tabName + ' row count', '', Math.max(0, dSheet.getLastRow() - 1));
      }
    });

    log('AI_KEY_OPENROUTER set?', '', !!PropertiesService.getScriptProperties().getProperty('AI_KEY_OPENROUTER'));
    log('SEARCH_API_KEY_SERPER set?', '', !!PropertiesService.getScriptProperties().getProperty('SEARCH_API_KEY_SERPER'));

    let me = '';
    try { me = Session.getActiveUser().getEmail() || ''; } catch (e) {}
    log('Session.getActiveUser()', '', me);
    log('ADMIN_EMAIL', '', ADMIN_EMAIL);
    log('isCurrentUserAdmin()', '', isCurrentUserAdmin());

    let webUrl = '';
    try { webUrl = ScriptApp.getService().getUrl() || ''; } catch (e) {}
    log('ScriptApp.getService().getUrl()', '', webUrl);

    // Simulate an actual ?agent=security-helpdesk web request and inspect the
    // RAW rendered HTML doGet() would send to a browser -- this is the only
    // way to see whether the <?= JSON.stringify(AGENT_SLUG) ?> template
    // scriptlet is embedding a clean JS string or HTML-entity-mangled garbage.
    try {
      const simulated = doGet({ parameter: { agent: 'security-helpdesk' } });
      const html = simulated.getContent();
      const idx = html.indexOf('AGENT_SLUG');
      log('Simulated doGet(?agent=security-helpdesk) html length', '', html.length);
      log('Raw HTML around AGENT_SLUG', '', idx >= 0 ? html.slice(Math.max(0, idx - 60), idx + 80) : '(AGENT_SLUG not found in output)');
      const headEnd = html.indexOf('</script>');
      log('Full head <script> block (first script tag)', '', html.slice(0, headEnd > 0 ? headEnd + 9 : 600));
    } catch (e) {
      log('Simulated doGet() FAILED', '', e.message || String(e));
    }

  } catch (e) {
    log('FATAL ERROR', '', e.message || String(e));
    log('FATAL STACK', '', e.stack || '');
  }

  if (rows.length) {
    dbg.getRange(2, 1, rows.length, 3).setValues(rows);
  }
  dbg.autoResizeColumns(1, 3);
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

  // Skill-agent HTML artifacts (printable one-pagers). Domain-gated above.
  if (String(params.page || '').toLowerCase() === 'skill-html' && params.id) {
    if (typeof serveSkillHtmlPage_ === 'function') {
      return serveSkillHtmlPage_(params.id);
    }
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

// Client-side injection filter (index.html) calls this when it flags a
// message as an attempt to change the agent's instructions. Logged to ChatLog
// with status='blocked' so the false-positive rate is visible in real use.
// Deliberately permissive on purpose: this is a UX nicety, not a security
// boundary. The persona and the server-side prompt construction are the real
// defense -- anything that slips past the client filter is still handled by
// the scope lock in _buildSystemPrompt_ (20_chat.js).
function logBlockedAttempt(slug, prompt) {
  try {
    var me = '';
    try { me = Session.getActiveUser().getEmail() || ''; } catch (e) {}
    _logChat_({
      email: me,
      slug: slug || '',
      model: '',
      skills: '',
      prompt: prompt || '',
      response: '',
      status: 'blocked',
      error: 'client-side injection filter',
      durationMs: 0
    });
  } catch (e) {
    // Logging must never break the user-facing call.
    Logger.log('logBlockedAttempt error: ' + (e && e.message));
  }
  return { ok: true };
}

// Agent platform thin wrappers (MVP)
function sendMessage(slug, history, text) {
  var started = Date.now();
  var email = '';
  try { email = Session.getActiveUser().getEmail() || ''; } catch (e) {}

  if (!slug) return { ok: false, error: 'Missing agent slug' };

  // Look up model + skills once, purely for ChatLog visibility (chatWithAgent_
  // re-reads the agent row itself; this extra read is cheap and never blocks a reply).
  var agentModel = '';
  var agentSkills = '';
  try {
    var aRes = getAgent(slug);
    if (aRes && aRes.ok) {
      agentModel = aRes.agent.model || DEFAULT_MODEL;
      agentSkills = aRes.agent.skills || DEFAULT_SKILLS.join(',');
    }
  } catch (e) {}

  // RiskAI v2.0 guardrails (25_riskai_guardrails.js): redact secrets/PII from
  // ChatLog only -- the reply returned to the client below is never touched,
  // so the user still sees their real answer. No-op (identity function) until
  // that file is pushed.
  var _redact = (typeof redactForLog_ === 'function') ? redactForLog_ : function (s) { return s; };

  try {
    var reply = chatWithAgent_(slug, history, text, email);
    _logChat_({
      email: email, slug: slug, model: agentModel, skills: agentSkills,
      prompt: _redact(text), response: _redact(reply),
      status: 'ok', durationMs: Date.now() - started
    });
    return { ok: true, reply: reply };
  } catch (e) {
    _logChat_({
      email: email, slug: slug, model: agentModel, skills: agentSkills,
      prompt: _redact(text), response: '',
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
    // Only the admin can browse the full roster. Everyone else reaches an
    // agent through the specific link they were given (?agent=slug) --
    // getAgent(slug) stays open to any @lausd.net user, this just stops
    // one shared link from letting someone discover every other agent.
    // Client-side hiding alone would not be a real boundary (same lesson as
    // runDiagnostics()), so this is enforced here too, not just in the UI.
    if (!isCurrentUserAdmin()) {
      return {
        ok: false,
        restricted: true,
        error: 'No agent specified. Use the link you were given for your assistant, or contact ' + ADMIN_EMAIL + ' for access.'
      };
    }
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
          org: row[map.org] || 'LAUSD',
          skills: map.skills !== undefined ? (row[map.skills] || '') : '',
          tone: map.tone !== undefined ? (row[map.tone] || '') : '',
          agent_type: map.agent_type !== undefined ? String(row[map.agent_type] || '').trim().toLowerCase() : ''
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
            skills: map.skills !== undefined ? (row[map.skills] || '') : '',
            // New (v1.3.0): kb_source is a free-text label (folder name/URL,
            // "SharePoint - IDM policies", etc.) purely for admins to see
            // where an agent's knowledge came from -- not read by the chat
            // pipeline. skill_refs is a comma list of SkillLibrary keys to
            // inject into this agent's system prompt (see 15_skill_library.js).
            kb_source: map.kb_source !== undefined ? (row[map.kb_source] || '') : '',
            skill_refs: map.skill_refs !== undefined ? (row[map.skill_refs] || '') : '',
            // New (v1.3.0): per-agent file upload for analysis (e.g. Ultimate
            // DAST Analyzer taking a scan report). See 16_file_upload.js.
            allow_upload: map.allow_upload !== undefined ? (row[map.allow_upload] || '') : '',
            upload_folder_id: map.upload_folder_id !== undefined ? (row[map.upload_folder_id] || '') : '',
            // Skill agents (26_skill_agent.js): blank agent_type is treated as
            // "agent" so existing persona rows are unchanged. default_skill is
            // the SkillLibrary key used when the user does not name one.
            agent_type: map.agent_type !== undefined ? String(row[map.agent_type] || '').trim().toLowerCase() : '',
            default_skill: map.default_skill !== undefined ? String(row[map.default_skill] || '').trim().toLowerCase() : ''
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
  // Client-side hiding of the Diagnostics panel is not a security boundary --
  // google.script.run is callable directly from devtools by any signed-in
  // user regardless of what the UI shows. Gate it here too.
  if (!isCurrentUserAdmin()) return { ok: false, error: 'Admin only.' };
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
