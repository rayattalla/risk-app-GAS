/**
 * 10_admin_ingest.js
 * Admin UI and ingest for Agent Platform (Sheet-backed)
 * Drop in same project as shell.js
 */

function _getSs_() {
  return SpreadsheetApp.openById(SHEET_ID);
}

function _getOrCreateTab_(name, headers) {
  const ss = _getSs_();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  if (sheet.getLastRow() < 1 && headers) {
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function _isAdminUser_() {
  const me = (Session.getActiveUser().getEmail() || '').toLowerCase();
  return me === (ADMIN_EMAIL || '').toLowerCase();
}

// ==========================================================================
// MENU ACTIONS
// ==========================================================================

function menuRunSetup() {
  if (!_isAdminUser_()) { _toast_('Admin only'); return; }
  // setup is in shell.js (inlined); fall back to local
  try {
    if (typeof setup === 'function') setup(); else _runBasicSetup();
  } catch (e) {
    _runBasicSetup();
  }
  _toast_('Setup complete');
}

function _runBasicSetup() {
  const headersAgents = ['slug', 'name', 'status', 'org', 'model', 'personality', 'kb_tags', 'notes', 'skills', 'tone'];
  _getOrCreateTab_('Agents', headersAgents);
  _getOrCreateTab_('KB', ['id', 'slug', 'title', 'body', 'tags']);
  _getOrCreateTab_('ChatLog', ['Timestamp','Email','Slug','Model','Prompt','Response','Status','Error','DurationMs']);
  _getOrCreateTab_('Access_Control', ['Email', 'Role', 'Added_Date', 'Notes']);
  _getOrCreateTab_('Automations', AUTOMATIONS_HEADERS);
  _getOrCreateTab_('Memory', MEMORY_HEADERS);
}

function menuSeedPilotAgentsKB() {
  if (!_isAdminUser_()) { _toast_('Admin only'); return; }
  _seedPilot();
  _toast_('Seeded 3 pilot agents + sample KB');
}

function _seedPilot() {
  const ss = _getSs_();

  // Agents
  let agents = ss.getSheetByName('Agents') || ss.insertSheet('Agents');
  if (agents.getLastRow() < 2) {
    agents.appendRow(['slug','name','status','org','model','personality','kb_tags','notes']);
  }

  const pilots = [
    {
      slug: 'security-helpdesk',
      name: 'Security Helpdesk',
      status: 'on',
      org: 'LAUSD',
      model: '',
      personality: 'You are the LAUSD Security Helpdesk Agent. Helpful, patient. Guide through account recovery, MFA, access issues using approved channels only. Use MyLogin, 5200-2, BUL-999.16. Never give passwords. Verify identity first. Be empathetic. End with offer for more help.',
      kb_tags: '',
      notes: 'pilot'
    },
    {
      slug: 'grc',
      name: 'GRC',
      status: 'on',
      org: 'LAUSD',
      model: '',
      personality: 'You are the LAUSD GRC Agent. Help with policy, vendor risk, compliance. Cite BULs, RUPs, NIST/CIS. For vendors use checklist. Be precise, conservative. Say when info missing. No legal advice.',
      kb_tags: '',
      notes: 'pilot'
    },
    {
      slug: 'ctu',
      name: 'CTU',
      status: 'on',
      org: 'LAUSD',
      model: '',
      personality: 'You are the LAUSD CTU Agent. Threat intel, IOCs, vulns, defensive actions. Reference CISA, MSRC, NVD. Actionable steps. Factual. Escalate high severity. No exploits.',
      kb_tags: '',
      notes: 'pilot'
    }
  ];

  pilots.forEach(p => {
    // check if exists
    const data = agents.getDataRange().getValues();
    let exists = false;
    for (let i=1; i<data.length; i++) {
      if (data[i][0] == p.slug) { exists = true; break; }
    }
    if (!exists) {
      agents.appendRow([p.slug, p.name, p.status, p.org, p.model, p.personality, p.kb_tags, p.notes]);
    }
  });

  // Sample KB (unified)
  const kb = ss.getSheetByName('KB') || ss.insertSheet('KB');
  if (kb.getLastRow() < 2) {
    kb.appendRow(['id','slug','title','body','tags']);
  }
  const kbSamples = [
    ['rule-00', 'shared', 'Core Rules', 'Follow LAUSD policies. Verify identity before actions. Log decisions. No secrets in chat.', 'rules,policy'],
    ['rule-01', 'shared', 'Escalation', 'For breaches or student data, escalate via official channels.', 'escalation'],
    ['hd-01', 'security-helpdesk', 'MyLogin', 'Use https://mylogin.lausd.net for password resets. For lockouts call (213) 241-5200 option 2.', 'mylogin,reset'],
    ['hd-02', 'security-helpdesk', '5200-2 Form', 'For account issues use form 5200-2. Never reset without verification.', 'form,access'],
    ['grc-01', 'grc', 'BUL-999.16', 'Data classification and protection policy. All systems must comply.', 'policy,data'],
    ['grc-02', 'grc', 'Vendor Checklist', 'Use 04-vendor-arb-checklist.md. Require classification, contract, CISO signoff for high risk.', 'vendor,arb'],
    ['ctu-01', 'ctu', 'Phish Flags', 'Urgency + off-domain link + request for creds = phish. Report to security.', 'phish,report'],
    ['ctu-02', 'ctu', 'Sources', 'Always cross ref CISA, MSRC, NVD. Provide patch/block/monitor steps.', 'intel,sources']
  ];
  kbSamples.forEach(row => kb.appendRow(row));
}

function menuIngest() {
  if (!_isAdminUser_()) { _toast_('Admin only'); return; }
  const ui = SpreadsheetApp.getUi();
  const res = ui.prompt('Ingest', 'Enter: URL (https://...)\n or "text:" + content\n or "drive:FILEID"\n or "sheet:Tab!A1:B10"\n or "builtin:helpdesk"', ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  const input = res.getResponseText().trim();
  try {
    _doIngest(input);
    _toast_('Ingest complete');
  } catch (e) {
    _toast_('Ingest error: ' + e.message);
  }
}

function menuIngestBuiltin() {
  if (!_isAdminUser_()) { _toast_('Admin only'); return; }
  const ui = SpreadsheetApp.getUi();
  const res = ui.prompt('Builtin pack', 'helpdesk | rup | cybersafety | stride', ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  const pack = res.getResponseText().trim().toLowerCase();
  _doIngest('builtin:' + pack);
  _toast_('Ingested builtin ' + pack);
}

function _doIngest(input) {
  const ss = _getSs_();
  let targetSlug = 'shared';
  let content = '';
  let title = 'Ingested';

  if (input.startsWith('http')) {
    content = UrlFetchApp.fetch(input, {muteHttpExceptions: true}).getContentText();
    title = 'URL: ' + input.substring(0,50);
  } else if (input.startsWith('text:')) {
    content = input.substring(5);
    title = 'Pasted text';
  } else if (input.startsWith('drive:')) {
    const id = input.substring(6).trim();
    content = DriveApp.getFileById(id).getBlob().getDataAsString();
    title = 'Drive: ' + id;
  } else if (input.startsWith('sheet:')) {
    const range = input.substring(6);
    const parts = range.split('!');
    const tab = parts[0];
    const rng = parts[1] || 'A1:Z100';
    const sh = ss.getSheetByName(tab);
    if (sh) content = sh.getRange(rng).getValues().map(r => r.join(' ')).join('\n');
    title = 'Sheet: ' + range;
  } else if (input.startsWith('builtin:')) {
    const p = input.substring(8);
    content = _getBuiltinPack(p);
    title = 'Builtin: ' + p;
    if (p === 'helpdesk') targetSlug = 'security-helpdesk';
    if (p === 'rup') targetSlug = 'grc';
    if (p === 'cybersafety') targetSlug = 'ctu';
    if (p === 'stride') targetSlug = 'ctu';
  } else {
    content = input;
  }

  content = content.replace(/<script[^>]*>[\s\S]*?<\/script>/gi, '')
                   .replace(/<style[^>]*>[\s\S]*?<\/style>/gi, '')
                   .replace(/<[^>]+>/g, ' ')
                   .replace(/\s+/g, ' ')
                   .trim();

  const sheet = _getOrCreateTab_('KB', ['id','slug','title','body','tags']);

  const chunks = [];
  const size = 2400;
  for (let i=0; i<content.length; i+=size) {
    chunks.push(content.substring(i, i+size));
  }

  chunks.forEach((c, idx) => {
    const id = Utilities.getUuid().slice(0,8);
    sheet.appendRow([id, targetSlug, title + (chunks.length>1 ? ' pt'+(idx+1) : ''), c, 'ingested,auto']);
  });
}

function _getBuiltinPack(pack) {
  if (pack === 'helpdesk') return 'MyLogin portal for resets. Call 213-241-5200 option 2 for lockouts. Use 5200-2 form. Verify identity always. Do not give passwords over chat.';
  if (pack === 'rup') return 'Reference 09-rup-obligations. All new systems need RUP review. Data classification required.';
  if (pack === 'cybersafety') return 'Report phishing to security@lausd.net. Suspicious email: do not click, forward as attachment. Use MFA everywhere.';
  if (pack === 'stride') return 'STRIDE threat modeling methodology — use when analyzing system security, threat modeling sessions, or writing security documentation.\n' +
    'S — Spoofing: an attacker impersonates a user, device, or process. Mitigate with strong auth, MFA, mutual TLS, signed tokens.\n' +
    'T — Tampering: unauthorized modification of data or code, in transit or at rest. Mitigate with integrity checks, digital signatures, checksums, access controls, audit logging.\n' +
    'R — Repudiation: a party denies performing an action and there is no proof otherwise. Mitigate with signed audit logs, timestamps, non-repudiation controls.\n' +
    'I — Information Disclosure: exposure of data to unauthorized parties. Mitigate with encryption at rest/in transit, least-privilege access, data classification (see BUL-999.16).\n' +
    'D — Denial of Service: degrading or blocking availability. Mitigate with rate limiting, redundancy, capacity planning, DDoS protection.\n' +
    'E — Elevation of Privilege: gaining higher permissions than authorized. Mitigate with least privilege, input validation, sandboxing, regular privilege reviews.\n' +
    'Workflow: (1) diagram the system and trust boundaries, (2) walk each element and data flow against all 6 STRIDE categories, (3) rate likelihood/impact, (4) map each finding to a mitigation and an owner, (5) track remediation to closure.';
  return 'Sample pack for ' + pack;
}

// ==========================================================================
// Agent CRUD is sheet-driven only (no admin web UI) — edit the Agents tab
// directly. listAgents()/getAgent() in shell.js already read it live.
// ==========================================================================

// Set OpenRouter key via menu (prompts for key, stores in Properties)
function menuSetOpenRouterKey() {
  if (!_isAdminUser_()) { _toast_('Admin only'); return; }
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    'Set OpenRouter API Key',
    'Paste your sk-or-... key (it will be stored in Script Properties as AI_KEY_OPENROUTER).',
    ui.ButtonSet.OK_CANCEL
  );
  if (response.getSelectedButton() !== ui.Button.OK) return;
  const key = response.getResponseText().trim();
  if (!key) { _toast_('No key provided'); return; }
  if (typeof setAiApiKey === 'function') {
    setAiApiKey('openrouter', key);
    _toast_('OpenRouter key set successfully. Refresh /exec to use.');
  } else {
    PropertiesService.getScriptProperties().setProperty('AI_KEY_OPENROUTER', key);
    _toast_('Key stored as AI_KEY_OPENROUTER. (setAiApiKey not found, direct set used)');
  }
}

// Set Serper.dev web-search key via menu (used by the 'search' skill).
function menuSetSearchKey() {
  if (!_isAdminUser_()) { _toast_('Admin only'); return; }
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    'Set Web Search API Key',
    'Paste your serper.dev API key (stored in Script Properties as SEARCH_API_KEY_SERPER).\n' +
    'Get one free at https://serper.dev — used only by agents with the "search" skill enabled.',
    ui.ButtonSet.OK_CANCEL
  );
  if (response.getSelectedButton() !== ui.Button.OK) return;
  const key = response.getResponseText().trim();
  if (!key) { _toast_('No key provided'); return; }
  PropertiesService.getScriptProperties().setProperty('SEARCH_API_KEY_SERPER', key);
  _toast_('Search key stored. Agents with "search" in their skills column can now use it.');
}

// ==========================================================================
// MIGRATE — one-time: add 'skills' + 'tone' columns to Agents (v1.2.0)
// Non-destructive: only appends missing columns, existing rows untouched
// (blank skills/tone = legacy default behavior, see 20_chat.js DEFAULT_SKILLS).
// ==========================================================================
function migrateAgentsAddSkillsTone_() {
  if (!_isAdminUser_()) { _toast_('Admin only.'); return; }
  const sheet = _getOrCreateTab_('Agents', ['slug', 'name', 'status', 'org', 'model', 'personality', 'kb_tags', 'notes']);
  const lastCol = Math.max(sheet.getLastColumn(), 1);
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h || '').trim().toLowerCase());

  const toAdd = [];
  if (headers.indexOf('skills') < 0) toAdd.push('skills');
  if (headers.indexOf('tone') < 0) toAdd.push('tone');

  if (!toAdd.length) {
    _toast_('Agents already has skills + tone columns.');
    return;
  }

  const startCol = lastCol + 1;
  sheet.getRange(1, startCol, 1, toAdd.length).setValues([toAdd]);
  sheet.getRange(1, startCol, 1, toAdd.length).setFontWeight('bold');

  // Backfill existing agent rows with sensible defaults so behavior stays
  // identical until an admin opts an agent into more skills.
  const numRows = sheet.getLastRow() - 1;
  if (numRows > 0) {
    const skillsColIdx = toAdd.indexOf('skills');
    if (skillsColIdx >= 0) {
      const col = startCol + skillsColIdx;
      const defaults = Array(numRows).fill(['kb']);
      sheet.getRange(2, col, numRows, 1).setValues(defaults);
    }
  }

  _toast_('Migrated Agents: added ' + toAdd.join(', ') + '. Existing agents defaulted to skills="kb".');
  SpreadsheetApp.getUi().alert(
    'Migration Complete',
    'Added column(s): ' + toAdd.join(', ') + '\n\n' +
    'Existing agents were backfilled with skills="kb" (unchanged behavior).\n' +
    'Edit the Agents sheet to add "webpage", "search", and/or "email" per agent, and set a tone string.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

// ==========================================================================
// MIGRATE — one-time: collapse KB_* + KB_shared into single KB tab
// Then run cleanupOldKbTabs_() to delete old tabs for shipping.
// Uses openById only.
// ==========================================================================
function migrateKbToSingleSheet_() {
  if (!_isAdminUser_()) { _toast_('Admin only'); return; }
  const ss = _getSs_();
  const kb = ss.getSheetByName('KB') || ss.insertSheet('KB');
  if (kb.getLastRow() < 1) {
    kb.appendRow(['id', 'slug', 'title', 'body', 'tags']);
    kb.getRange(1,1,1,5).setFontWeight('bold');
    kb.setFrozenRows(1);
  }
  const map = _headerMapForMigrate_(kb); // helper below

  let copied = 0;
  const allSheets = ss.getSheets();
  allSheets.forEach(sh => {
    const name = sh.getName();
    if (!/^KB_/.test(name) && name !== 'KB_shared') return;
    let slug = name.replace(/^KB_/, '').toLowerCase();
    if (name === 'KB_shared') slug = 'shared';
    const data = sh.getDataRange().getValues();
    for (let i=1; i<data.length; i++) {  // skip header
      const r = data[i];
      const id = String(r[0] || '').trim() || Utilities.getUuid().slice(0,8);
      const title = String(r[1] || '').trim();
      const body = String(r[2] || '').trim();
      const tags = String(r[3] || '').trim();
      // simple append; duplicates possible on re-run
      kb.appendRow([id, slug, title, body, tags]);
      copied++;
    }
  });
  const msg = 'Migrated ' + copied + ' rows from KB_* tabs into KB. Old tabs preserved.';
  _toast_(msg);
  SpreadsheetApp.getUi().alert('Migrate KB', msg + '\n\nOpen the KB tab to inspect. Do not delete old tabs yet.', SpreadsheetApp.getUi().ButtonSet.OK);
}

// ==========================================================================
// CLEANUP for shipping: delete old KB_* tabs (after migration confirmed)
// Also fixes Agents header if it still has legacy 'knowledge_tab' column name.
// Run from menu or editor: cleanupOldKbTabs_()
// ==========================================================================
function cleanupOldKbTabs_() {
  if (!_isAdminUser_()) { _toast_('Admin only'); return; }
  const ss = _getSs_();
  const toDelete = [];
  const deleted = [];
  ss.getSheets().forEach(sh => {
    const name = sh.getName();
    if (name === 'KB_shared' || /^KB_/.test(name)) {
      toDelete.push(sh);
      deleted.push(name);
    }
  });
  toDelete.forEach(sh => {
    try { ss.deleteSheet(sh); } catch (e) { /* ignore */ }
  });

  // Fix Agents header column name if legacy
  const agents = ss.getSheetByName('Agents');
  if (agents && agents.getLastColumn() >= 7) {
    const h = String(agents.getRange(1, 7).getValue() || '').trim().toLowerCase();
    if (h === 'knowledge_tab') {
      agents.getRange(1, 7).setValue('kb_tags');
      deleted.push('fixed Agents header: knowledge_tab → kb_tags');
    }
  }

  const msg = deleted.length ? 'Cleaned: ' + deleted.join(', ') : 'No old KB tabs found.';
  _toast_(msg);
  SpreadsheetApp.getUi().alert('Cleanup Complete', msg + '\n\nProject ready for shipping prep.', SpreadsheetApp.getUi().ButtonSet.OK);
}

function _headerMapForMigrate_(sheet) {
  if (!sheet || sheet.getLastColumn() < 1) return {};
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const map = {};
  headers.forEach((h, i) => { if (h) map[String(h).trim().toLowerCase()] = i; });
  return map;
}
