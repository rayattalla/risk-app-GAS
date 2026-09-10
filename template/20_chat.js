/**
 * 20_chat.js — agent chat execution
 * Stateless: reads persona + KB from the Sheet, calls the model, returns text.
 * No conversation is persisted server-side beyond the audit log.
 */

var DEFAULT_MODEL = 'openai/gpt-4o-mini';   // overridable per-agent via Agents.model; confirm with Ray
var MAX_KB_CHARS  = 12000;                  // hard ceiling on injected context
var MAX_WEB_CHARS = 4000;                   // hard ceiling on webpage/search injected context

// Skills (v1.2.0) — modeled on Khoj's per-agent input_tools allowlist.
// Agents.skills is a comma list; unset/blank means legacy default 'kb' only,
// so agent rows created before this migration keep their old behavior.
var DEFAULT_SKILLS = ['kb'];

function _parseSkills_(skillsField) {
  var raw = String(skillsField || '').trim();
  if (!raw) return DEFAULT_SKILLS.slice();
  return raw.split(',').map(function (s) { return s.trim().toLowerCase(); }).filter(Boolean);
}

function chatWithAgent_(slug, history, text, email) {
  if (!slug) throw new Error('Missing agent slug');
  if (!text || !String(text).trim()) throw new Error('Empty message');

  var res = getAgent(slug);
  if (!res.ok) throw new Error('Unknown agent: ' + slug);
  var agent = res.agent;
  if (agent.status !== 'on') throw new Error('Agent is disabled: ' + slug);

  var skills = _parseSkills_(agent.skills);
  var systemPrompt = _buildSystemPrompt_(agent, skills, email);

  var webContext = '';
  if (skills.indexOf('webpage') >= 0) {
    var url = _extractFirstUrl_(text);
    if (url) webContext += _fetchWebpageSkill_(url);
  }
  if (skills.indexOf('search') >= 0 && /^search:/i.test(String(text).trim())) {
    webContext += _webSearchSkill_(String(text).trim().replace(/^search:/i, '').trim());
  }
  if (webContext) {
    systemPrompt += '\n\n=== Live Web Context (from webpage/search skill) ===\n' + webContext;
  }

  var messages = [{ role: 'system', content: systemPrompt }];

  // history arrives from the client as [{role, content}, ...]
  // client may have already appended the current user turn, so avoid double
  var h = (history || []).slice(-10);
  if (h.length && h[h.length-1] && h[h.length-1].role === 'user' &&
      String(h[h.length-1].content || '') === String(text)) {
    // already present, use as-is
  } else {
    h = h.concat([{ role: 'user', content: String(text) }]);
  }
  h.forEach(function (m) {
    if (m && m.role && m.content) {
      messages.push({ role: m.role, content: String(m.content) });
    }
  });

  return _callModel_(agent.model || DEFAULT_MODEL, messages);
}

function _buildSystemPrompt_(agent, skills, email) {
  skills = skills || _parseSkills_(agent && agent.skills);
  var ctx = '';
  if (agent && agent.personality) {
    ctx += 'You are ' + (agent.name || agent.slug) + ' for ' + (agent.org || 'LAUSD') + '.\n' + agent.personality + '\n\n';
  }
  if (agent && agent.tone) {
    ctx += 'Tone: ' + agent.tone + '\n\n';
  }
  ctx += 'Follow all LAUSD policies and the shared rules below. Be concise, professional, and cite sources or policy references when possible. If asked for something outside your role, politely redirect.\n\n';

  if (skills.indexOf('memory') >= 0 && email) {
    var mem = _getMemoryNote_(email, agent.slug);
    if (mem) {
      ctx += '=== What you remember about this user (from past sessions) ===\n' + mem + '\n' +
             'Use this only if relevant; if it conflicts with what the user says now, trust the current message.\n\n';
    }
  }

  if (skills.indexOf('kb') >= 0) {
    var kb = _getKbForSlug_(agent.slug);
    if (kb) {
      ctx += '=== Knowledge Base (shared + ' + (agent && agent.slug ? agent.slug : 'agent') + ') ===\n';
      ctx += kb + '\n\n';
      ctx += 'Use the reference knowledge above when relevant. ' +
             'If it does not cover the question, say so rather than guessing.\n\n';
    }
  }

  if (skills.indexOf('webpage') >= 0) {
    ctx += 'If the user pastes a URL, you may be given fetched page content under "Live Web Context" — use it, and say if it was unavailable.\n';
  }
  if (skills.indexOf('search') >= 0) {
    ctx += 'The user can prefix a message with "search: <query>" to have you search the live web; use any "Live Web Context" provided.\n';
  }
  if (skills.indexOf('memory') >= 0) {
    ctx += 'You retain durable facts about returning users across sessions via a memory note (see above if any exists).\n';
  }
  return ctx;
}

// ==========================================================================
// SKILLS (v1.2.0) — each gated per-agent via Agents.skills
// ==========================================================================

function _extractFirstUrl_(text) {
  var m = String(text || '').match(/https?:\/\/[^\s)]+/);
  return m ? m[0] : '';
}

function _stripHtml_(html) {
  return String(html || '')
    .replace(/<script[^>]*>[\s\S]*?<\/script>/gi, '')
    .replace(/<style[^>]*>[\s\S]*?<\/style>/gi, '')
    .replace(/<[^>]+>/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

// 'webpage' skill — fetch a URL the user pasted and inject its text.
function _fetchWebpageSkill_(url) {
  try {
    var resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true, followRedirects: true });
    if (resp.getResponseCode() !== 200) {
      return '\n[webpage] Could not fetch ' + url + ' (HTTP ' + resp.getResponseCode() + ')\n';
    }
    var text = _stripHtml_(resp.getContentText()).slice(0, MAX_WEB_CHARS);
    return '\n[webpage: ' + url + ']\n' + text + '\n';
  } catch (e) {
    return '\n[webpage] Fetch failed for ' + url + ': ' + (e.message || e) + '\n';
  }
}

// 'search' skill — requires a Serper.dev API key set via
// menu Setup & Config -> Set Web Search API Key. Fails soft (tells the model
// search wasn't available) rather than silently pretending to search.
function _webSearchSkill_(query) {
  if (!query) return '';
  var key = PropertiesService.getScriptProperties().getProperty('SEARCH_API_KEY_SERPER');
  if (!key) {
    return '\n[search] Web search skill is enabled but no SEARCH_API_KEY_SERPER is configured. ' +
           'Tell the user web search is not yet set up; do not fabricate results.\n';
  }
  try {
    var resp = UrlFetchApp.fetch('https://google.serper.dev/search', {
      method: 'post',
      contentType: 'application/json',
      headers: { 'X-API-KEY': key },
      payload: JSON.stringify({ q: query }),
      muteHttpExceptions: true
    });
    if (resp.getResponseCode() !== 200) {
      return '\n[search] Search API returned ' + resp.getResponseCode() + '. Tell the user search failed.\n';
    }
    var data = JSON.parse(resp.getContentText());
    var lines = [];
    (data.organic || []).slice(0, 5).forEach(function (r) {
      lines.push('- ' + (r.title || '') + ': ' + (r.snippet || '') + ' (' + (r.link || '') + ')');
    });
    var out = '\n[search: ' + query + ']\n' + lines.join('\n') + '\n';
    return out.slice(0, MAX_WEB_CHARS);
  } catch (e) {
    return '\n[search] Search failed: ' + (e.message || e) + '\n';
  }
}

// 'email' skill — emails the transcript to the requesting user's own address.
// Called from shell.js#emailTranscript(); requires gmail.send scope (already
// declared in appsscript.json).
function _sendTranscriptEmail_(toEmail, slug, history) {
  var lines = (history || []).map(function (m) {
    return (m.role === 'user' ? 'You: ' : 'Agent: ') + m.content;
  });
  var body = 'Transcript with ' + slug + ' — ' + PROGRAM_NAME + '\n\n' + lines.join('\n\n');
  GmailApp.sendEmail(toEmail, '[' + PROGRAM_SHORT + '] Chat transcript — ' + slug, body);
}

function _getKbForSlug_(slug) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName('KB');
  if (!sheet) return '';

  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return '';

  var headers = data[0].map(function (h) { return String(h).trim().toLowerCase(); });
  var iSlug  = headers.indexOf('slug');
  var iTitle = headers.indexOf('title');
  var iBody  = headers.indexOf('body');
  if (iSlug < 0 || iBody < 0) return '';

  var target = String(slug).trim().toLowerCase();
  var out = [];
  var used = 0;

  for (var i = 1; i < data.length; i++) {
    var rowSlug = String(data[i][iSlug] || '').trim().toLowerCase();
    if (rowSlug === 'slug') continue;            // duplicate header row guard
    if (rowSlug !== 'shared' && rowSlug !== target) continue;

    var title = iTitle >= 0 ? String(data[i][iTitle] || '') : '';
    var body  = String(data[i][iBody] || '');
    if (!body) continue;

    var chunk = (title ? title + ': ' : '') + body;
    if (used + chunk.length > MAX_KB_CHARS) break;
    out.push(chunk);
    used += chunk.length;
  }
  return out.join('\n');
}

function _callModel_(model, messages) {
  var key = PropertiesService.getScriptProperties().getProperty('AI_KEY_OPENROUTER');
  if (!key) throw new Error('AI_KEY_OPENROUTER is not set in Script Properties.');

  var resp = UrlFetchApp.fetch('https://openrouter.ai/api/v1/chat/completions', {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + key },
    payload: JSON.stringify({ model: model, messages: messages }),
    muteHttpExceptions: true
  });

  var code = resp.getResponseCode();
  var raw  = resp.getContentText();

  if (code !== 200) {
    throw new Error('Model API returned ' + code + ': ' + raw.slice(0, 300));
  }

  var parsed;
  try { parsed = JSON.parse(raw); }
  catch (e) { throw new Error('Model API returned unparseable response.'); }

  var reply = parsed &&
              parsed.choices &&
              parsed.choices[0] &&
              parsed.choices[0].message &&
              parsed.choices[0].message.content;

  if (!reply) throw new Error('Model API returned no content.');
  return String(reply);
}

function _logChat_(entry) {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName('ChatLog');
    if (!sheet) {
      sheet = ss.insertSheet('ChatLog');
      sheet.appendRow(['Timestamp','Email','Slug','Model','Prompt',
                       'Response','Status','Error','DurationMs']);
      sheet.setFrozenRows(1);
    }
    sheet.appendRow([
      new Date(),
      entry.email || '',
      entry.slug || '',
      entry.model || '',
      entry.prompt || '',
      entry.response || '',
      entry.status || '',
      entry.error || '',
      entry.durationMs || ''
    ]);
  } catch (e) {
    // Logging must never break the user-facing call.
    Logger.log('ChatLog write failed: ' + (e && e.message));
  }
}

/**
 * _purgeOldChatLog_ — dormant; Ray will set CHATLOG_RETENTION_DAYS Script Property
 * then we can add a time-driven trigger. Deletes rows older than N days.
 * Call manually or later via trigger for data protection.
 */
function _purgeOldChatLog_() {
  var days = parseInt(PropertiesService.getScriptProperties().getProperty('CHATLOG_RETENTION_DAYS') || '0', 10);
  if (!days || days <= 0) return;
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName('ChatLog');
    if (!sheet || sheet.getLastRow() < 2) return;
    var data = sheet.getDataRange().getValues();
    var cutoff = new Date();
    cutoff.setDate(cutoff.getDate() - days);
    var toDelete = [];
    for (var i = 1; i < data.length; i++) {  // skip header
      var ts = data[i][0];
      if (ts instanceof Date && ts < cutoff) {
        toDelete.push(i + 1);
      }
    }
    for (var j = toDelete.length - 1; j >= 0; j--) {
      sheet.deleteRow(toDelete[j]);
    }
  } catch (e) {
    Logger.log('purge failed: ' + e.message);
  }
}
