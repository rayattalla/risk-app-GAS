/**
 * 40_memory.js — long-term memory (v1.2.0)
 *
 * Modeled on Khoj's UserMemory, but ported without a vector DB — same idea
 * as the "total-recall" pattern found while scraping GitHub for skill
 * ideas: no embeddings, no per-message writes. An hourly sweep looks at
 * what changed in ChatLog in the last window, asks the model to fold that
 * into a short running note per (user, agent), and _buildSystemPrompt_
 * injects that note back in for agents with the 'memory' skill enabled.
 *
 * Why a sweep instead of updating memory on every message: keeps cost and
 * latency off the hot chat path entirely, and avoids the note thrashing on
 * every single turn — it only needs to be "pretty current", not real-time.
 */

var MEMORY_HEADERS = ['email', 'agent_slug', 'note', 'updated_at'];
var MEMORY_TRIGGER_FN = 'runMemoryDistillation_';
var MEMORY_SWEEP_WINDOW_MIN = 70; // >60 to cover trigger jitter, dedup is by hour bucket below
var MAX_MEMORY_CHARS = 1500;      // hard ceiling on a single memory note

function _getMemorySheet_() {
  return _getOrCreateTab_('Memory', MEMORY_HEADERS);
}

function _findMemoryRow_(sheet, email, slug) {
  const last = sheet.getLastRow();
  if (last < 2) return -1;
  const data = sheet.getRange(2, 1, last - 1, 2).getValues();
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][0]).toLowerCase() === String(email).toLowerCase() && data[i][1] === slug) return i + 2;
  }
  return -1;
}

function _getMemoryNote_(email, slug) {
  if (!email || !slug) return '';
  const sheet = _getMemorySheet_();
  const row = _findMemoryRow_(sheet, email, slug);
  if (row < 0) return '';
  return String(sheet.getRange(row, 3).getValue() || '');
}

function _updateMemoryNote_(email, slug, note) {
  const sheet = _getMemorySheet_();
  const trimmed = String(note || '').slice(0, MAX_MEMORY_CHARS);
  const row = _findMemoryRow_(sheet, email, slug);
  if (row > 0) {
    sheet.getRange(row, 3, 1, 2).setValues([[trimmed, new Date()]]);
  } else {
    sheet.appendRow([email, slug, trimmed, new Date()]);
  }
}

function _clearMemoryNote_(email, slug) {
  const sheet = _getMemorySheet_();
  const row = _findMemoryRow_(sheet, email, slug);
  if (row > 0) sheet.deleteRow(row);
}

// Asks the model to fold the latest exchanges into an updated note. Uses the
// same _callModel_ as regular chat (20_chat.js) — no new API dependency.
function _distillMemory_(existingNote, newExchangesText) {
  const prompt = [
    { role: 'system', content:
        'You maintain a short running memory note about one user, for one AI agent, across sessions. ' +
        'Given the CURRENT NOTE and the LATEST EXCHANGES, produce an UPDATED note.\n' +
        'Rules: keep only durable facts/preferences relevant to future conversations (role, recurring topics, ' +
        'stated preferences, open items they said they would follow up on). Do NOT store one-off small talk, ' +
        'secrets, passwords, or anything sensitive beyond what is operationally useful. Bullet points. ' +
        'Under 150 words. Output ONLY the updated note text, nothing else.'
    },
    { role: 'user', content: 'CURRENT NOTE:\n' + (existingNote || '(none yet)') + '\n\nLATEST EXCHANGES:\n' + newExchangesText }
  ];
  return _callModel_(DEFAULT_MODEL, prompt);
}

// Hourly sweep: group ChatLog rows from the last window by (email, slug),
// skip agents that don't have the 'memory' skill enabled, distill, upsert.
function runMemoryDistillation_() {
  const ss = _getSs_();
  const chatLog = ss.getSheetByName('ChatLog');
  if (!chatLog || chatLog.getLastRow() < 2) return;

  const cutoff = new Date(Date.now() - MEMORY_SWEEP_WINDOW_MIN * 60 * 1000);
  const data = chatLog.getRange(2, 1, chatLog.getLastRow() - 1, 9).getValues();
  // columns: Timestamp, Email, Slug, Model, Prompt, Response, Status, Error, DurationMs

  const groups = {}; // "email||slug" -> [{prompt, response}]
  data.forEach(function (row) {
    const ts = row[0];
    if (!(ts instanceof Date) || ts < cutoff) return;
    if (String(row[6] || '').toLowerCase() !== 'ok') return;
    const email = row[1], slug = row[2];
    if (!email || !slug) return;
    const key = email + '||' + slug;
    if (!groups[key]) groups[key] = [];
    groups[key].push({ prompt: row[4], response: row[5] });
  });

  Object.keys(groups).forEach(function (key) {
    const parts = key.split('||');
    const email = parts[0], slug = parts[1];

    const agentRes = getAgent(slug);
    if (!agentRes.ok) return;
    const skills = _parseSkills_(agentRes.agent.skills);
    if (skills.indexOf('memory') < 0) return; // not opted in — skip, no wasted model call

    const exchangesText = groups[key].map(function (e) {
      return 'User: ' + e.prompt + '\nAgent: ' + e.response;
    }).join('\n\n');

    try {
      const existing = _getMemoryNote_(email, slug);
      const updated = _distillMemory_(existing, exchangesText);
      _updateMemoryNote_(email, slug, updated);
    } catch (e) {
      Logger.log('runMemoryDistillation_ failed for ' + key + ': ' + (e.message || e));
    }
  });
}

// Idempotent, same pattern as installAutomationsTrigger_.
function installMemoryTrigger_() {
  if (!_isAdminUser_()) { _toast_('Admin only.'); return; }
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === MEMORY_TRIGGER_FN) ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger(MEMORY_TRIGGER_FN)
    .timeBased()
    .everyHours(1)
    .create();
  _toast_('Hourly memory distillation trigger installed. Only agents with "memory" in their skills column are processed.');
}

// Admin-only bulk wipe, for compliance/privacy requests — mirrors
// _purgeOldChatLog_'s existence in 20_chat.js as a data-retention control.
function menuClearAllMemory_() {
  if (!_isAdminUser_()) { _toast_('Admin only.'); return; }
  const sheet = _getMemorySheet_();
  const last = sheet.getLastRow();
  if (last > 1) sheet.deleteRows(2, last - 1);
  _toast_('All memory notes cleared.');
}
