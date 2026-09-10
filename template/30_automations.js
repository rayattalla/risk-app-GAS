/**
 * 30_automations.js — Khoj-style scheduled automations (v1.2.0)
 *
 * A saved natural-language query against one agent, run on a schedule, with
 * the reply emailed to whoever created it. Modeled on Khoj's
 * routers/api_automation.py (saved query + cron -> APScheduler -> email),
 * ported to Apps Script: one hourly time-driven trigger scans the
 * Automations tab for due rows instead of a real cron daemon.
 *
 * Granularity is hour-of-day + optional day-of-week list, not full cron —
 * Apps Script triggers can't fire more often than once/hour reliably anyway.
 */

var AUTOMATIONS_HEADERS = ['id', 'agent_slug', 'query', 'hour', 'days', 'notify_email', 'last_run_date', 'status', 'created_by', 'created_at'];
var AUTOMATIONS_TRIGGER_FN = 'runDueAutomations_';

function _getAutomationsSheet_() {
  return _getOrCreateTab_('Automations', AUTOMATIONS_HEADERS);
}

function menuAddAutomation() {
  if (!_isAdminUser_()) { _toast_('Admin only.'); return; }
  const ui = SpreadsheetApp.getUi();

  const slugResp = ui.prompt('Add Automation (1/5)', 'Agent slug to run (e.g. security-helpdesk):', ui.ButtonSet.OK_CANCEL);
  if (slugResp.getSelectedButton() !== ui.Button.OK) return;
  const slug = slugResp.getResponseText().trim();
  if (!slug || !getAgent(slug).ok) { _toast_('Unknown agent slug: ' + slug); return; }

  const qResp = ui.prompt('Add Automation (2/5)', 'Query to run (as if typed in chat):', ui.ButtonSet.OK_CANCEL);
  if (qResp.getSelectedButton() !== ui.Button.OK) return;
  const query = qResp.getResponseText().trim();
  if (!query) { _toast_('Query required.'); return; }

  const hourResp = ui.prompt('Add Automation (3/5)', 'Hour of day to run, 0-23 (America/Los_Angeles):', ui.ButtonSet.OK_CANCEL);
  if (hourResp.getSelectedButton() !== ui.Button.OK) return;
  const hour = parseInt(hourResp.getResponseText().trim(), 10);
  if (isNaN(hour) || hour < 0 || hour > 23) { _toast_('Hour must be 0-23.'); return; }

  const daysResp = ui.prompt('Add Automation (4/5)', 'Days: "daily" or comma list (mon,tue,wed,thu,fri,sat,sun):', ui.ButtonSet.OK_CANCEL);
  if (daysResp.getSelectedButton() !== ui.Button.OK) return;
  const days = (daysResp.getResponseText().trim() || 'daily').toLowerCase();

  const meEmail = Session.getActiveUser().getEmail() || ADMIN_EMAIL;
  const emailResp = ui.prompt('Add Automation (5/5)', 'Notify email (blank = ' + meEmail + '):', ui.ButtonSet.OK_CANCEL);
  if (emailResp.getSelectedButton() !== ui.Button.OK) return;
  const notifyEmail = emailResp.getResponseText().trim() || meEmail;

  const sheet = _getAutomationsSheet_();
  const id = Utilities.getUuid().slice(0, 8);
  sheet.appendRow([id, slug, query, hour, days, notifyEmail, '', 'on', meEmail, new Date()]);

  _toast_('Automation added (' + id + '). Run "Install Hourly Trigger" once if you haven\'t already.');
}

function menuListAutomations() {
  if (!_isAdminUser_()) { _toast_('Admin only.'); return; }
  const sheet = _getAutomationsSheet_();
  if (sheet.getLastRow() < 2) { _toast_('No automations yet.'); return; }
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, AUTOMATIONS_HEADERS.length).getValues();
  const lines = data.map(r => r[0] + ': [' + r[7] + '] ' + r[1] + ' @ hour ' + r[3] + ' (' + r[4] + ') -> ' + r[5]);
  SpreadsheetApp.getUi().alert('Automations (' + lines.length + ')', lines.join('\n'), SpreadsheetApp.getUi().ButtonSet.OK);
}

// Idempotent: removes any prior trigger for this function before adding one,
// so re-running "Install Hourly Trigger" never stacks duplicate triggers.
function installAutomationsTrigger_() {
  if (!_isAdminUser_()) { _toast_('Admin only.'); return; }
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === AUTOMATIONS_TRIGGER_FN) ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger(AUTOMATIONS_TRIGGER_FN)
    .timeBased()
    .everyHours(1)
    .create();
  _toast_('Hourly automations trigger installed.');
}

// Fired by the hourly trigger (or manually via menu). Scans Automations for
// rows due this hour that haven't already run today, executes them through
// the normal chat pipeline, and emails the reply.
function runDueAutomations_() {
  const sheet = _getAutomationsSheet_();
  if (sheet.getLastRow() < 2) return;

  const tz = TIMEZONE || 'America/Los_Angeles';
  const now = new Date();
  const currentHour = parseInt(Utilities.formatDate(now, tz, 'H'), 10);
  const todayStr = Utilities.formatDate(now, tz, 'yyyy-MM-dd');
  const currentDayName = Utilities.formatDate(now, tz, 'EEE').toLowerCase().slice(0, 3);

  const numRows = sheet.getLastRow() - 1;
  const data = sheet.getRange(2, 1, numRows, AUTOMATIONS_HEADERS.length).getValues();
  const idx = {};
  AUTOMATIONS_HEADERS.forEach(function (h, i) { idx[h] = i; });

  let ran = 0;
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    if (String(row[idx.status] || '').toLowerCase() !== 'on') continue;
    if (parseInt(row[idx.hour], 10) !== currentHour) continue;
    if (String(row[idx.last_run_date] || '') === todayStr) continue; // already ran this hour-window today

    const days = String(row[idx.days] || 'daily').toLowerCase();
    if (days !== 'daily') {
      const wanted = days.split(',').map(function (d) { return d.trim(); });
      if (wanted.indexOf(currentDayName) < 0) continue;
    }

    const slug = row[idx.agent_slug];
    const query = row[idx.query];
    const notifyEmail = row[idx.notify_email];

    try {
      const reply = chatWithAgent_(slug, [], query);
      GmailApp.sendEmail(
        notifyEmail,
        '[' + PROGRAM_SHORT + '] Automation: ' + slug,
        'Query: ' + query + '\n\n' + reply
      );
      _logChat_({ email: notifyEmail, slug: slug, model: '', prompt: '[automation] ' + query, response: reply, status: 'ok', durationMs: 0 });
    } catch (e) {
      try {
        GmailApp.sendEmail(notifyEmail, '[' + PROGRAM_SHORT + '] Automation FAILED: ' + slug, 'Query: ' + query + '\n\nError: ' + (e.message || e));
      } catch (e2) { /* swallow — never let one bad row break the sweep */ }
    }

    sheet.getRange(2 + i, idx.last_run_date + 1).setValue(todayStr);
    ran++;
  }
  Logger.log('runDueAutomations_: ' + ran + ' automation(s) run for hour ' + currentHour);
}
