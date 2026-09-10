/**
 * RiskAI-GAS — Notifications (06_notify.js)
 *
 * Multi-channel outbound notifications beyond basic email:
 *   - Microsoft Teams webhook (MessageCard format)
 *   - SMS via email-to-carrier gateways (no external service needed)
 *   - External webhook / SIEM / SOAR payload dispatch
 *   - Sheet-driven token-template system (editable without code changes)
 *
 * Extracted from: incident-reporting-nist-aligned (IRS v3.4)
 *
 * Public API:
 *   renderTemplate(template, tokens)              — replace {TOKEN} in any string
 *   getNotificationTemplate(eventKey, channel)    — read from Notification_Templates sheet
 *   sendTeamsWebhook(webhookUrl, title, body, color) — post a card to Teams
 *   sendSms(phoneNumber, carrier, message)        — email-to-SMS gateway
 *   sendExternalWebhook(url, eventType, payload)  — POST JSON to SIEM/SOAR/ticketing
 *
 * All channels respect isTestMode() from 01_config.js — in test mode they log
 * instead of sending (except Teams webhook which still fires but prefixes [TEST]).
 */

// ==========================================================================
// CARRIER SMS GATEWAYS — email-to-SMS gateway domains
// Usage: send MailApp.sendEmail(digits + '@' + CARRIER_SMS_GATEWAYS[carrier], '', msg)
// ==========================================================================

const CARRIER_SMS_GATEWAYS = {
  'AT&T':        'txt.att.net',
  'T-Mobile':    'tmomail.net',
  'Verizon':     'vtext.com',
  'Sprint':      'messaging.sprintpcs.com',
  'Metro':       'mymetropcs.com',
  'Boost':       'sms.myboostmobile.com',
  'Cricket':     'sms.cricketwireless.net',
  'US Cellular': 'email.uscc.net',
  'Google Fi':   'msg.fi.google.com'
};

// ==========================================================================
// TOKEN TEMPLATING — sheet-driven editable notifications
// ==========================================================================

/**
 * Replace {TOKEN} placeholders in a template string.
 *
 * Example:
 *   renderTemplate('Incident {ID} is now {STATUS}.', { ID: 'INC-001', STATUS: 'Resolved' })
 *   → 'Incident INC-001 is now Resolved.'
 */
function renderTemplate(template, tokens) {
  let result = String(template || '');
  Object.keys(tokens || {}).forEach(key => {
    result = result.replace(new RegExp('\\{' + key + '\\}', 'g'), String(tokens[key] != null ? tokens[key] : ''));
  });
  return result;
}

/**
 * Look up a notification template from the Notification_Templates sheet.
 *
 * Sheet layout (create via setupNotificationTemplates()):
 *   Event | Email | Teams | SMS
 *   INCIDENT_NEW | Dear {SUBMITTER}, ... | New incident {ID} ... | Inc {ID} {STATUS}
 *
 * Returns the raw template string (with {TOKENS}), or '' if not found.
 * Caller combines with renderTemplate() to fill tokens.
 *
 * eventKey: string matching the "Event" column (e.g., 'INCIDENT_NEW')
 * channel:  column name to read ('Email', 'Teams', 'SMS') — defaults to 'Email'
 */
function getNotificationTemplate(eventKey, channel) {
  channel = channel || 'Email';
  const sheet = _getSheet_(CONFIG.SHEET_TABS.NOTIFICATION_TEMPLATES);
  if (!sheet || sheet.getLastRow() < 2) return '';
  const hdr  = _headerMap_(sheet);
  const keyCol  = hdr['Event'] != null ? hdr['Event'] : 0;
  const chanCol = hdr[channel] != null ? hdr[channel] : 1;
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][keyCol] || '').trim() === eventKey) {
      return String(data[i][chanCol] || '').trim();
    }
  }
  return '';
}

/**
 * Create the Notification_Templates sheet with standard columns.
 * Call once from setup(). Safe to re-run.
 */
function setupNotificationTemplates() {
  const sheet = _getOrCreateSheet_(CONFIG.SHEET_TABS.NOTIFICATION_TEMPLATES);
  if (sheet.getLastRow() < 1) {
    sheet.appendRow(['Event', 'Email', 'Teams', 'SMS', 'Notes']);
    sheet.getRange(1, 1, 1, 5).setFontWeight('bold');
    sheet.setFrozenRows(1);
    // Example row so the admin knows the format
    sheet.appendRow([
      'EXAMPLE_EVENT',
      'Hello {RECIPIENT}, your record {ID} was updated to {STATUS}.',
      '{ID} updated to {STATUS} by {ACTOR}.',
      '{ID}: {STATUS}',
      'Delete this row; add your own event keys.'
    ]);
    sheet.setColumnWidth(2, 400);
    sheet.setColumnWidth(3, 300);
  }
}

// ==========================================================================
// FETCH WITH RETRY — exponential backoff for external HTTP calls
// WHY: Teams webhooks, SIEM endpoints, and external APIs fail transiently.
// Without retry, a 5xx or timeout silently drops the notification.
// ==========================================================================

/**
 * UrlFetchApp.fetch() with exponential backoff.
 *
 * url:     endpoint URL
 * options: same options object as UrlFetchApp.fetch() (method, contentType, payload, etc.)
 * maxRetries: number of attempts after the first (default 3 → up to 4 total)
 *
 * Returns the HTTPResponse on success.
 * Throws on final failure (caller decides whether to swallow or re-throw).
 *
 * Backoff schedule (default): 1s → 2s → 4s between attempts.
 */
function fetchWithRetry_(url, options, maxRetries) {
  maxRetries = maxRetries != null ? maxRetries : 3;
  options = options || {};
  // Always mute exceptions so we can inspect the response code ourselves
  options.muteHttpExceptions = true;

  let lastError;
  for (let attempt = 0; attempt <= maxRetries; attempt++) {
    if (attempt > 0) {
      Utilities.sleep(Math.pow(2, attempt - 1) * 1000); // 1s, 2s, 4s
    }
    try {
      const resp = UrlFetchApp.fetch(url, options);
      const code = resp.getResponseCode();
      if (code >= 200 && code < 300) return resp;
      // 4xx = don't retry (client error); 5xx = retry
      if (code >= 400 && code < 500) {
        Logger.log('fetchWithRetry_ client error ' + code + ' (no retry): ' + url);
        throw new Error('HTTP ' + code + ': ' + resp.getContentText().slice(0, 200));
      }
      lastError = new Error('HTTP ' + code + ' on attempt ' + (attempt + 1));
      Logger.log('fetchWithRetry_ server error ' + code + ', attempt ' + (attempt + 1) + '/' + (maxRetries + 1));
    } catch (e) {
      if (e.message && e.message.startsWith('HTTP 4')) throw e; // don't retry 4xx
      lastError = e;
      Logger.log('fetchWithRetry_ exception on attempt ' + (attempt + 1) + ': ' + e.message);
    }
  }
  throw lastError || new Error('fetchWithRetry_ failed after ' + (maxRetries + 1) + ' attempts');
}

// ==========================================================================
// MICROSOFT TEAMS WEBHOOK
// ==========================================================================

/**
 * Post a MessageCard to a Microsoft Teams channel via incoming webhook.
 *
 * webhookUrl: the /IncomingWebhook URL from the Teams channel connector
 * title:      card title (shown in bold)
 * body:       card body (plain text or simple HTML-like Teams markdown)
 * color:      hex color for the card left border (default: LAUSD blue)
 *
 * In test mode: prefixes title with [TEST] but still sends (Teams has no sandbox).
 */
function sendTeamsWebhook(webhookUrl, title, body, color) {
  if (!webhookUrl) { Logger.log('sendTeamsWebhook: no URL provided'); return; }
  color = (color || '#0072CE').replace('#', '');
  if (isTestMode()) title = '[TEST] ' + title;

  const card = {
    '@type':    'MessageCard',
    '@context': 'http://schema.org/extensions',
    themeColor: color,
    summary:    title,
    sections: [{ activityTitle: title, activityText: body }]
  };

  try {
    fetchWithRetry_(webhookUrl, {
      method:      'post',
      contentType: 'application/json',
      payload:     JSON.stringify(card)
    });
  } catch (e) {
    Logger.log('sendTeamsWebhook failed after retries: ' + e.message);
  }
}

// ==========================================================================
// SMS VIA EMAIL-TO-CARRIER GATEWAY
// ==========================================================================

/**
 * Send an SMS via the carrier's email-to-SMS gateway.
 * No external API or billing needed — uses MailApp quota.
 *
 * phoneNumber: 10-digit US number (digits only, or formatted — we strip non-digits)
 * carrier:     key from CARRIER_SMS_GATEWAYS (e.g., 'AT&T', 'Verizon')
 * message:     body text (truncated to 160 chars)
 *
 * In test mode: logs the message but does not send.
 */
function sendSms(phoneNumber, carrier, message) {
  const digits = _digitsOnly_(phoneNumber);
  if (!digits) { Logger.log('sendSms: no digits in phone number'); return; }
  const gateway = CARRIER_SMS_GATEWAYS[carrier];
  if (!gateway) { Logger.log('sendSms: unknown carrier "' + carrier + '"'); return; }

  const smsAddress = digits + '@' + gateway;
  const truncated  = String(message || '').slice(0, 160);

  if (isTestMode()) {
    Logger.log('sendSms [TEST — not sent] → ' + smsAddress + ': ' + truncated);
    return;
  }

  try {
    MailApp.sendEmail(smsAddress, '', truncated);
  } catch (e) {
    Logger.log('sendSms failed to ' + smsAddress + ': ' + e.message);
  }
}

function _digitsOnly_(s) {
  return String(s || '').replace(/\D/g, '');
}

// ==========================================================================
// EXTERNAL WEBHOOK — SIEM / SOAR / Ticketing
// ==========================================================================

/**
 * POST a JSON payload to an external webhook URL.
 * Wraps the payload with schema version, event type, timestamp, and test flag
 * so downstream systems can route/filter reliably.
 *
 * url:       webhook URL (leave blank to skip)
 * eventType: string like 'RECORD_CREATED', 'STATUS_CHANGED', 'REPORT_READY'
 * payload:   arbitrary object — the event-specific data
 *
 * In test mode: prefixes eventType with 'TEST_' so downstream systems can
 * distinguish test traffic from real events.
 */
function sendExternalWebhook(url, eventType, payload) {
  if (!url) return;
  const body = {
    schemaVersion: '1.0',
    eventType:     isTestMode() ? 'TEST_' + eventType : eventType,
    source:        CONFIG.PROGRAM_SHORT,
    timestamp:     _now_(),
    testing:       isTestMode(),
    data:          payload || {}
  };
  try {
    fetchWithRetry_(url, {
      method:      'post',
      contentType: 'application/json',
      payload:     JSON.stringify(body)
    });
    Logger.log('sendExternalWebhook sent: ' + eventType);
  } catch (e) {
    Logger.log('sendExternalWebhook failed after retries (' + eventType + '): ' + e.message);
  }
}
