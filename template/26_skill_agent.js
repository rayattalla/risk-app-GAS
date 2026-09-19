/**
 * 26_skill_agent.js
 *
 * Skill-agent path: no persona, no KB, no capability skills. The model only
 * structures pasted data into JSON; GAS assembles sheet / doc / html
 * artifacts (zero tokens). Looker Studio is out of scope — there is no
 * Apps Script create-report API.
 *
 * Agents.agent_type = "skill" (blank/other = existing persona path).
 * Agents.default_skill = SkillLibrary key to use if the user does not name one.
 * SkillLibrary.outputs = comma list of any subset of {sheet, doc, html}.
 */

var SKILL_AGENT_TYPE = 'skill';
var SKILL_AGENT_DEFAULT_OUTPUTS = ['doc', 'html'];
var SKILL_ARTIFACT_HEADERS = ['id', 'title', 'html', 'created_at', 'created_by', 'slug'];
var MAX_SKILL_AGENT_INSTR_CHARS = 8000;
var MAX_SKILL_AGENT_DATA_CHARS = 12000;
var MAX_SKILL_HTML_CHARS = 45000;
var SKILL_NAVY = '#1A237E';

var SKILL_AGENT_SYSTEM_PROMPT =
  'Organize the pasted data into JSON only. No markdown fences, no commentary, no extra keys.\n' +
  'Schema: {"title":string,"summary":string,"sections":[{"heading":string,"body":string}],"table":{"headers":[string],"rows":[[string]]}}\n' +
  'Follow the skill instructions for titles and sections. Use only facts present in the user data. Omit missing fields. Never invent names, dates, or numbers.';

// ==========================================================================
// Schema — Agents.agent_type + Agents.default_skill (append-only)
// ==========================================================================

function _ensureAgentsSkillAgentColumns_() {
  const sheet = _getSs_().getSheetByName('Agents');
  if (!sheet) return [];
  const lastCol = Math.max(sheet.getLastColumn(), 1);
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function (h) {
    return String(h || '').trim().toLowerCase();
  });
  const toAdd = [];
  if (headers.indexOf('agent_type') < 0) toAdd.push('agent_type');
  if (headers.indexOf('default_skill') < 0) toAdd.push('default_skill');
  if (!toAdd.length) return [];
  sheet.getRange(1, lastCol + 1, 1, toAdd.length).setValues([toAdd]);
  sheet.getRange(1, lastCol + 1, 1, toAdd.length).setFontWeight('bold');
  return toAdd;
}

function migrateAgentsAddAgentType_() {
  if (!_isAdminUser_()) { _toast_('Admin only.'); return; }
  const added = _ensureAgentsSkillAgentColumns_();
  if (!added.length) {
    _toast_('Agents already has agent_type and default_skill.');
    return;
  }
  _toast_(
    'Added ' + added.join(', ') + ' to Agents. Existing rows stay blank ' +
    '(treated as agent_type=agent). New skill agents use agent_type=skill.'
  );
}

function _isSkillAgent_(agent) {
  return String(agent && agent.agent_type || '').trim().toLowerCase() === SKILL_AGENT_TYPE;
}

// ==========================================================================
// Chat entry — called from chatWithAgent_ when agent_type == skill
// ==========================================================================

function chatWithSkillAgent_(agent, history, text, email) {
  const pasted = _skillUserData_(history, text);
  const resolved = _resolveSkillKey_(agent, pasted);
  if (!resolved.key) {
    const keys = _listSkillLibraryKeys_();
    return 'This is a skill agent — it has no persona or knowledge base.\n\n' +
      'Name a skill (first line or `skill: key`) and paste the data to organize' +
      (keys.length ? '.\n\nAvailable skills:\n- ' + keys.join('\n- ') : '.') +
      (agent.default_skill ? '' : '\n\nAn admin can set Agents.default_skill to skip naming it each time.');
  }
  if (!resolved.skill || !resolved.skill.instructions) {
    return 'Skill "' + resolved.key + '" was not found in SkillLibrary, or its instructions cell is empty.';
  }

  const outputs = _normalizeSkillOutputs_(resolved.skill.outputs);
  const kinds = outputs.length ? outputs : SKILL_AGENT_DEFAULT_OUTPUTS.slice();

  const skillText = String(resolved.skill.instructions || '').slice(0, MAX_SKILL_AGENT_INSTR_CHARS);
  const dataText = resolved.data.slice(0, MAX_SKILL_AGENT_DATA_CHARS);
  const userContent =
    '=== Skill instructions ===\n' + skillText +
    '\n\n=== User data ===\n' + (dataText || '(no data pasted)');

  const raw = _callModel_(agent.model || DEFAULT_MODEL, [
    { role: 'system', content: SKILL_AGENT_SYSTEM_PROMPT },
    { role: 'user', content: userContent }
  ]);

  let payload;
  try {
    payload = _parseSkillModelJson_(raw);
  } catch (e) {
    return 'Could not parse structured output from the model (' + (e.message || e) +
      '). No artifacts were built. Raw reply:\n\n' + raw;
  }

  const built = _assembleSkillArtifacts_(payload, kinds, {
    slug: agent.slug,
    email: email,
    skillKey: resolved.key,
    skillTitle: resolved.skill.title || resolved.key
  });

  return _formatSkillAgentReply_(payload, built, resolved, resolved.skill.title || resolved.key);
}

function _skillUserData_(history, text) {
  const current = String(text || '').trim();
  if (current) return current;
  const h = history || [];
  for (let i = h.length - 1; i >= 0; i--) {
    if (h[i] && h[i].role === 'user' && String(h[i].content || '').trim()) {
      return String(h[i].content).trim();
    }
  }
  return '';
}

function _resolveSkillKey_(agent, text) {
  const def = _slugifySkillKey_(agent && agent.default_skill);
  let named = '';
  let data = String(text || '');

  const prefix = data.match(/^\s*(?:skill|use)\s*[:\-]\s*([A-Za-z0-9][A-Za-z0-9_-]*)\s*(?:\n+([\s\S]*))?$/i);
  if (prefix) {
    named = _slugifySkillKey_(prefix[1]);
    data = String(prefix[2] || '').trim();
  } else {
    const lines = data.split(/\r?\n/);
    const first = String(lines[0] || '').trim();
    if (/^[A-Za-z0-9][A-Za-z0-9_-]{0,40}$/.test(first) && _getSkillLibraryRow_(first.toLowerCase())) {
      named = first.toLowerCase();
      data = lines.slice(1).join('\n').trim();
    }
  }

  const key = named || def;
  return {
    key: key,
    skill: key ? _getSkillLibraryRow_(key) : null,
    data: data
  };
}

function _parseSkillModelJson_(raw) {
  let s = String(raw || '').trim();
  const fence = s.match(/```(?:json)?\s*([\s\S]*?)```/i);
  if (fence) s = fence[1].trim();
  const start = s.indexOf('{');
  const end = s.lastIndexOf('}');
  if (start < 0 || end <= start) throw new Error('no JSON object in model reply');
  const parsed = JSON.parse(s.slice(start, end + 1));
  if (!parsed || typeof parsed !== 'object') throw new Error('JSON was not an object');
  return {
    title: String(parsed.title || '').trim(),
    summary: String(parsed.summary || '').trim(),
    sections: Array.isArray(parsed.sections) ? parsed.sections.map(function (sec) {
      return {
        heading: String(sec && sec.heading || '').trim(),
        body: String(sec && sec.body || '').trim()
      };
    }).filter(function (sec) { return sec.heading || sec.body; }) : [],
    table: _normalizeSkillTable_(parsed.table)
  };
}

function _normalizeSkillTable_(table) {
  if (!table || typeof table !== 'object') return null;
  const headers = Array.isArray(table.headers)
    ? table.headers.map(function (h) { return String(h == null ? '' : h); })
    : [];
  const rows = Array.isArray(table.rows)
    ? table.rows.map(function (r) {
        return (Array.isArray(r) ? r : [r]).map(function (c) { return String(c == null ? '' : c); });
      })
    : [];
  if (!headers.length && !rows.length) return null;
  return { headers: headers, rows: rows };
}

function _assembleSkillArtifacts_(payload, kinds, ctx) {
  const out = { links: [], errors: [] };
  kinds.forEach(function (kind) {
    try {
      if (kind === 'sheet') out.links.push(_buildSkillSheet_(payload, ctx));
      else if (kind === 'doc') {
        const doc = _buildSkillDoc_(payload, ctx);
        out.links.push(doc);
        try {
          out.links.push(_exportSkillPdf_(doc.id, payload.title || doc.title, ctx.email));
        } catch (pdfErr) {
          out.errors.push('PDF from doc: ' + (pdfErr.message || pdfErr));
        }
      } else if (kind === 'html') {
        out.links.push(_buildSkillHtml_(payload, ctx));
      }
    } catch (e) {
      out.errors.push(kind + ': ' + (e.message || e));
    }
  });
  return out;
}

function _formatSkillAgentReply_(payload, built, resolved, skillTitle) {
  const lines = [];
  lines.push((payload.title || skillTitle || resolved.key) + ' — artifacts ready.');
  if (payload.summary) lines.push('', payload.summary);
  lines.push('');
  if (built.links.length) {
    built.links.forEach(function (item) {
      lines.push(item.label + ': ' + item.url);
    });
  } else {
    lines.push('No artifacts were created.');
  }
  if (built.errors.length) {
    lines.push('', 'Some formats failed:');
    built.errors.forEach(function (err) { lines.push('- ' + err); });
  }
  lines.push('', 'Skill: ' + resolved.key + ' (persona/KB skipped).');
  return lines.join('\n');
}

// ==========================================================================
// Artifact builders (pure GAS — no extra model calls)
// ==========================================================================

function _skillArtifactTitle_(payload, fallback) {
  const t = String(payload && payload.title || fallback || 'Skill output').replace(/[\r\n]+/g, ' ').trim();
  return t.slice(0, 80) || 'Skill output';
}

function _shareSkillFile_(fileId, email) {
  if (!fileId || !email) return;
  try { DriveApp.getFileById(fileId).addViewer(email); } catch (e) { /* already owned / policy */ }
}

function _buildSkillSheet_(payload, ctx) {
  const title = _skillArtifactTitle_(payload, (ctx.skillTitle || 'Skill') + ' sheet');
  const ss = SpreadsheetApp.create(title);
  const sh = ss.getActiveSheet();
  sh.setName('Output');

  let headers;
  let bodyRows;
  if (payload.table && (payload.table.headers.length || payload.table.rows.length)) {
    headers = payload.table.headers.length ? payload.table.headers : _defaultTableHeaders_(payload.table.rows[0]);
    bodyRows = payload.table.rows.map(function (r) {
      const copy = r.slice();
      while (copy.length < headers.length) copy.push('');
      return copy.slice(0, headers.length);
    });
  } else {
    headers = ['Item', 'Detail'];
    bodyRows = [];
    if (payload.title) bodyRows.push(['Title', payload.title]);
    if (payload.summary) bodyRows.push(['Summary', payload.summary]);
    (payload.sections || []).forEach(function (sec) {
      bodyRows.push([sec.heading || 'Section', sec.body || '']);
    });
    if (!bodyRows.length) bodyRows.push(['Note', 'No structured rows were produced.']);
  }

  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  sh.getRange(1, 1, 1, headers.length)
    .setBackground(SKILL_NAVY)
    .setFontColor('#ffffff')
    .setFontWeight('bold');
  if (bodyRows.length) {
    sh.getRange(2, 1, bodyRows.length, headers.length).setValues(bodyRows);
    sh.getRange(2, 1, bodyRows.length, headers.length).setWrap(true).setVerticalAlignment('top');
  }
  sh.setFrozenRows(1);
  sh.setRowHeight(1, 28);
  for (let c = 1; c <= headers.length; c++) {
    sh.autoResizeColumn(c);
    if (sh.getColumnWidth(c) < 140) sh.setColumnWidth(c, 140);
    if (sh.getColumnWidth(c) > 420) sh.setColumnWidth(c, 420);
  }
  _shareSkillFile_(ss.getId(), ctx.email);
  return { kind: 'sheet', label: 'Sheet', id: ss.getId(), url: ss.getUrl(), title: title };
}

function _defaultTableHeaders_(row) {
  const n = (row && row.length) || 2;
  const headers = [];
  for (let i = 0; i < n; i++) headers.push('Col ' + (i + 1));
  return headers;
}

function _buildSkillDoc_(payload, ctx) {
  const title = _skillArtifactTitle_(payload, (ctx.skillTitle || 'Skill') + ' doc');
  const doc = DocumentApp.create(title);
  const body = doc.getBody();
  body.clear();

  const h = body.appendParagraph(payload.title || title);
  h.setHeading(DocumentApp.ParagraphHeading.TITLE);
  h.setForegroundColor(SKILL_NAVY);

  if (payload.summary) {
    const sum = body.appendParagraph(payload.summary);
    sum.setItalic(true);
  }

  (payload.sections || []).forEach(function (sec) {
    if (sec.heading) {
      const ph = body.appendParagraph(sec.heading);
      ph.setHeading(DocumentApp.ParagraphHeading.HEADING2);
      ph.setForegroundColor(SKILL_NAVY);
    }
    if (sec.body) body.appendParagraph(sec.body);
  });

  if (payload.table && payload.table.headers.length && payload.table.rows.length) {
    body.appendParagraph('Key facts').setHeading(DocumentApp.ParagraphHeading.HEADING2).setForegroundColor(SKILL_NAVY);
    const cells = [payload.table.headers].concat(payload.table.rows.map(function (r) {
      const copy = r.slice();
      while (copy.length < payload.table.headers.length) copy.push('');
      return copy.slice(0, payload.table.headers.length);
    }));
    body.appendTable(cells);
  }

  doc.saveAndClose();
  _shareSkillFile_(doc.getId(), ctx.email);
  return { kind: 'doc', label: 'Doc', id: doc.getId(), url: doc.getUrl(), title: title };
}

function _exportSkillPdf_(fileId, title, email) {
  const file = DriveApp.getFileById(fileId);
  const pdf = file.getBlob().getAs('application/pdf');
  pdf.setName((title || file.getName()) + '.pdf');
  const saved = DriveApp.createFile(pdf);
  _shareSkillFile_(saved.getId(), email);
  return { kind: 'pdf', label: 'PDF', id: saved.getId(), url: saved.getUrl(), title: saved.getName() };
}

function _buildSkillHtml_(payload, ctx) {
  const title = _skillArtifactTitle_(payload, ctx.skillTitle || 'Skill output');
  const sectionsHtml = (payload.sections || []).map(function (sec) {
    return '<section><h2>' + _htmlEscape_(sec.heading) + '</h2><p>' +
      _htmlEscape_(sec.body).replace(/\n/g, '<br>') + '</p></section>';
  }).join('');

  let tableHtml = '';
  if (payload.table && payload.table.headers.length) {
    const head = '<tr>' + payload.table.headers.map(function (h) {
      return '<th>' + _htmlEscape_(h) + '</th>';
    }).join('') + '</tr>';
    const rows = (payload.table.rows || []).map(function (r) {
      return '<tr>' + payload.table.headers.map(function (_, i) {
        return '<td>' + _htmlEscape_(r[i] || '') + '</td>';
      }).join('') + '</tr>';
    }).join('');
    tableHtml = '<table><thead>' + head + '</thead><tbody>' + rows + '</tbody></table>';
  }

  const html =
    '<!doctype html><html lang="en"><head><meta charset="utf-8">' +
    '<meta name="viewport" content="width=device-width, initial-scale=1">' +
    '<title>' + _htmlEscape_(title) + '</title>' +
    '<style>' +
    'body{font-family:Georgia,Times,serif;max-width:800px;margin:2rem auto;padding:0 1.25rem 3rem;color:#1A1A2E;line-height:1.45}' +
    'h1{color:' + SKILL_NAVY + ';font-size:1.8rem;margin-bottom:.4rem}' +
    'h2{color:' + SKILL_NAVY + ';font-size:1.15rem;margin:1.4rem 0 .4rem}' +
    '.summary{font-style:italic;color:#333}' +
    'table{border-collapse:collapse;width:100%;margin-top:1rem;font-size:.95rem}' +
    'th{background:' + SKILL_NAVY + ';color:#fff;text-align:left;padding:.45rem .6rem}' +
    'td{border:1px solid #E8EAED;padding:.45rem .6rem;vertical-align:top}' +
    '.bar{height:6px;background:' + SKILL_NAVY + ';margin:-2rem -1.25rem 1.5rem}' +
    '.noprint{margin:0 0 1rem}' +
    '.noprint button{background:' + SKILL_NAVY + ';color:#fff;border:0;border-radius:6px;padding:.45rem .9rem;font-weight:600;cursor:pointer}' +
    '@media print{.noprint{display:none}body{margin:0}}' +
    '</style></head><body>' +
    '<div class="bar"></div>' +
    '<p class="noprint"><button type="button" onclick="window.print()">Print / Save as PDF</button></p>' +
    '<h1>' + _htmlEscape_(payload.title || title) + '</h1>' +
    (payload.summary ? '<p class="summary">' + _htmlEscape_(payload.summary) + '</p>' : '') +
    sectionsHtml + tableHtml +
    '</body></html>';

  const id = Utilities.getUuid().slice(0, 12);
  _storeSkillHtml_(id, title, html.slice(0, MAX_SKILL_HTML_CHARS), ctx);
  const base = _skillWebAppUrl_();
  const url = base
    ? (base + (base.indexOf('?') >= 0 ? '&' : '?') + 'page=skill-html&id=' + encodeURIComponent(id))
    : '(deploy the web app, then reopen this chat to get the HTML link)';
  return { kind: 'html', label: 'HTML (print to PDF in browser)', id: id, url: url, title: title };
}

function _htmlEscape_(s) {
  return String(s == null ? '' : s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function _skillWebAppUrl_() {
  try {
    if (typeof getWebAppUrl === 'function') {
      const u = getWebAppUrl();
      if (u) return u;
    }
  } catch (e) {}
  try { return ScriptApp.getService().getUrl() || ''; } catch (e2) { return ''; }
}

function _getSkillArtifactsSheet_() {
  return _getOrCreateTab_('SkillArtifacts', SKILL_ARTIFACT_HEADERS);
}

function _storeSkillHtml_(id, title, html, ctx) {
  const sheet = _getSkillArtifactsSheet_();
  sheet.appendRow([id, title, html, new Date(), ctx.email || '', ctx.slug || '']);
}

function _getSkillArtifactHtml_(id) {
  const want = String(id || '').trim();
  if (!want) return null;
  const sheet = _getSs_().getSheetByName('SkillArtifacts');
  if (!sheet || sheet.getLastRow() < 2) return null;
  const data = sheet.getDataRange().getValues();
  const headers = data[0].map(function (h) { return String(h || '').trim().toLowerCase(); });
  const iId = headers.indexOf('id');
  const iTitle = headers.indexOf('title');
  const iHtml = headers.indexOf('html');
  if (iId < 0 || iHtml < 0) return null;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][iId] || '').trim() === want) {
      return {
        id: want,
        title: iTitle >= 0 ? String(data[i][iTitle] || '') : '',
        html: String(data[i][iHtml] || '')
      };
    }
  }
  return null;
}

function serveSkillHtmlPage_(id) {
  const rec = _getSkillArtifactHtml_(id);
  if (!rec || !rec.html) {
    return HtmlService.createHtmlOutput(
      '<p style="font-family:sans-serif;padding:2rem">Skill HTML artifact not found.</p>'
    ).setTitle('Not found');
  }
  return HtmlService.createHtmlOutput(rec.html)
    .setTitle(rec.title || 'Skill artifact')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ==========================================================================
// Seed — OnePager skill + skill agent row
// ==========================================================================

var ONEPAGER_SKILL_KEY = 'onepager';
var ONEPAGER_SKILL_INSTRUCTIONS =
  'Turn the user\'s pasted notes into a single-page briefing.\n\n' +
  'Extract:\n' +
  '- A short title (8 words or fewer)\n' +
  '- A 2-3 sentence summary\n' +
  '- 3-6 sections with a heading and body (facts only from the paste)\n' +
  '- A table of key facts when the paste has comparable items (2-4 columns, or Item/Detail rows)\n\n' +
  'Do not invent facts, dates, names, or numbers that are not in the pasted data. ' +
  'If something is missing, omit it. This is an organize-only skill — no advice beyond structuring what was pasted.';

function seedSkillAgent_() {
  if (!_isAdminUser_()) { _toast_('Admin only.'); return; }
  _ensureAgentsSkillAgentColumns_();
  _ensureSkillLibrarySchema_();

  const skillResult = _upsertSkillLibraryRow_(
    ONEPAGER_SKILL_KEY,
    'One-Pager',
    ONEPAGER_SKILL_INSTRUCTIONS,
    '(built-in sample)',
    'sheet,doc,html'
  );

  const agentResult = _upsertAgentFields_('onepager', {
    slug: 'onepager',
    name: 'One-Pager',
    status: 'on',
    org: 'LAUSD',
    model: '',
    personality: '',
    kb_tags: '',
    notes: 'Skill agent — instructions live in SkillLibrary (onepager), not in personality.',
    skills: '',
    tone: '',
    agent_type: SKILL_AGENT_TYPE,
    default_skill: ONEPAGER_SKILL_KEY
  });

  _getSkillArtifactsSheet_();

  const msg = 'OnePager skill ' + skillResult + '; agent "onepager" ' + agentResult +
    '. Open ?agent=onepager, paste notes, get sheet/doc/html links.';
  _toast_(msg);
  SpreadsheetApp.getUi().alert('Seed skill agent', msg, SpreadsheetApp.getUi().ButtonSet.OK);
}

function menuSeedSkillAgent_() {
  seedSkillAgent_();
}

function _upsertAgentFields_(slug, fields) {
  const sheet = _getSs_().getSheetByName('Agents');
  if (!sheet) throw new Error('No Agents tab found. Run Setup first.');
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const map = {};
  headers.forEach(function (h, i) {
    const key = String(h || '').trim().toLowerCase();
    if (key) map[key] = i;
  });
  if (map.slug === undefined) throw new Error('Agents is missing a slug column.');

  const want = String(slug || '').trim().toLowerCase();
  const data = sheet.getDataRange().getValues();
  let rowIdx = -1;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][map.slug] || '').trim().toLowerCase() === want) {
      rowIdx = i + 1;
      break;
    }
  }

  const width = sheet.getLastColumn();
  const row = rowIdx > 0 ? data[rowIdx - 1].slice() : [];
  while (row.length < width) row.push('');

  Object.keys(fields).forEach(function (k) {
    if (map[k] !== undefined) row[map[k]] = fields[k];
  });

  if (rowIdx > 0) {
    sheet.getRange(rowIdx, 1, 1, row.length).setValues([row]);
    return 'updated';
  }
  sheet.appendRow(row);
  return 'inserted';
}
