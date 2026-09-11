/**
 * 16_file_upload.js
 *
 * Per-agent file upload + analysis (e.g. Ultimate DAST Analyzer taking a scan
 * report). Config lives on the Agents row, same pattern as skills/tone:
 *   allow_upload      "yes" to show the upload control for this agent
 *   upload_folder_id  Drive folder ID the original file is archived into
 *                      before any analysis happens (required if allow_upload
 *                      is yes -- there is deliberately no silent fallback
 *                      location for what can be sensitive scan data).
 *
 * The uploaded file is never stored in the Sheet itself -- only its
 * extracted text goes into the model prompt (capped, like KB/webpage
 * content), and only a short synthetic description goes into ChatLog. The
 * original bytes live in exactly one place: the configured Drive folder.
 */

var MAX_UPLOAD_BYTES = 8 * 1024 * 1024;      // 8 MB -- mirrors the client-side cap in index.html
var MAX_UPLOAD_TEXT_CHARS = 20000;           // generous vs. MAX_KB_CHARS -- the file IS the point of the request

function _isUploadAllowed_(agent) {
  return /^(yes|true|on)$/i.test(String(agent && agent.allow_upload || '').trim());
}

// Returns extracted text, or null if the file type isn't supported.
function _extractTextFromUploadedBlob_(blob, filename, mimeType) {
  var lower = String(filename || '').toLowerCase();
  var mime = String(mimeType || '').toLowerCase();
  try {
    if (mime === 'application/pdf' || /\.pdf$/.test(lower)) {
      return _extractPdfTextFromBlob_(blob, filename);
    }
    if (mime === 'text/html' || /\.html?$/.test(lower)) {
      return _stripHtml_(blob.getDataAsString());
    }
    var textLikeMimes = [
      'text/plain', 'text/markdown', 'text/csv', 'text/xml', 'application/xml',
      'application/json', 'application/javascript', 'application/x-yaml', 'text/yaml'
    ];
    // DAST/vuln-scan exports are almost always one of these: ZAP (xml/json/html),
    // Burp (xml), Nessus/Qualys (csv/xml), SARIF (json).
    if (textLikeMimes.indexOf(mime) >= 0 || /\.(txt|md|markdown|csv|xml|json|log|yaml|yml|sarif)$/.test(lower)) {
      return blob.getDataAsString();
    }
  } catch (e) {
    return '[extraction failed: ' + (e.message || e) + ']';
  }
  return null;
}

function _extractPdfTextFromBlob_(blob, filename) {
  if (typeof Drive === 'undefined') {
    throw new Error('PDF analysis needs the "Drive" Advanced Service (same one used by Drive-folder KB ingestion) -- see appsscript.json.');
  }
  var tempFile = DriveApp.createFile(blob).setName('UPLOAD_TEMP_' + filename);
  var ocrDoc = null;
  try {
    var resource = { title: 'OCR_TEMP_' + filename, mimeType: MimeType.GOOGLE_DOCS };
    ocrDoc = Drive.Files.insert(resource, tempFile.getBlob(), { ocr: true, ocrLanguage: 'en' });
    return DocumentApp.openById(ocrDoc.id).getBody().getText();
  } finally {
    try { tempFile.setTrashed(true); } catch (e1) { /* ignore cleanup failure */ }
    if (ocrDoc) { try { Drive.Files.remove(ocrDoc.id); } catch (e2) { /* ignore cleanup failure */ } }
  }
}

// Main RPC, called from index.html's analyzeFile(). Archives the original
// file first (so there's a record even if extraction/the model call fails),
// then extracts text, then runs it through the normal chat pipeline as a
// synthetic user turn -- so persona, tone, scope-lock, and skills all still
// apply exactly as they do to a typed message.
function analyzeUploadedFile(slug, base64Data, filename, mimeType, history) {
  var started = Date.now();
  var email = '';
  try { email = Session.getActiveUser().getEmail() || ''; } catch (e) {}

  if (!slug) return { ok: false, error: 'Missing agent slug' };

  var agentRes = getAgent(slug);
  if (!agentRes.ok) return { ok: false, error: 'Unknown agent: ' + slug };
  var agent = agentRes.agent;

  if (!_isUploadAllowed_(agent)) {
    return { ok: false, error: 'File upload is not enabled for this agent.' };
  }
  if (!agent.upload_folder_id) {
    return { ok: false, error: 'No upload folder configured for this agent (Agents.upload_folder_id is blank) -- an admin needs to set one before uploads can be enabled.' };
  }

  var bytes;
  try {
    bytes = Utilities.base64Decode(base64Data);
  } catch (e) {
    return { ok: false, error: 'Could not decode the uploaded file.' };
  }
  if (bytes.length > MAX_UPLOAD_BYTES) {
    return { ok: false, error: 'File too large (' + Math.round(bytes.length / 1024 / 1024) + ' MB) -- 8 MB max.' };
  }

  var blob = Utilities.newBlob(bytes, mimeType || 'application/octet-stream', filename || 'upload');

  var savedUrl = '';
  try {
    var folder = DriveApp.getFolderById(agent.upload_folder_id);
    var stamp = Utilities.formatDate(new Date(), (typeof TIMEZONE !== 'undefined' ? TIMEZONE : 'America/Los_Angeles'), 'yyyyMMdd_HHmmss');
    var savedFile = folder.createFile(blob).setName(stamp + '_' + (email || 'unknown') + '_' + filename);
    savedUrl = savedFile.getUrl();
  } catch (e) {
    return { ok: false, error: 'Could not save to the configured upload folder: ' + (e.message || e) };
  }

  var text = _extractTextFromUploadedBlob_(blob, filename, mimeType);
  if (text === null) {
    _logChat_({
      email: email, slug: slug, model: agent.model || '', skills: agent.skills || '',
      prompt: '[uploaded file: ' + filename + ', unsupported type ' + mimeType + ']', response: '',
      status: 'error', error: 'Unsupported file type', durationMs: Date.now() - started
    });
    return { ok: false, error: 'Unsupported file type for analysis (' + mimeType + '). Saved to the archive but not analyzed: ' + savedUrl };
  }

  var truncated = text.length > MAX_UPLOAD_TEXT_CHARS;
  var textForModel = text.slice(0, MAX_UPLOAD_TEXT_CHARS);
  var prompt = 'The user uploaded a file named "' + filename + '" for analysis.' +
    (truncated ? ' Only the first ' + MAX_UPLOAD_TEXT_CHARS + ' characters are included below (the file is longer) -- note that limit if it affects your analysis.' : '') +
    '\n\n=== File content ===\n' + textForModel;

  try {
    var reply = chatWithAgent_(slug, history || [], prompt, email);
    // Short synthetic prompt in ChatLog, not the full file content -- the
    // original is already archived in Drive; this keeps the log readable
    // and avoids duplicating potentially sensitive scan data a second time.
    _logChat_({
      email: email, slug: slug, model: agent.model || '', skills: agent.skills || '',
      prompt: '[uploaded file: ' + filename + ', ' + Math.round(bytes.length / 1024) + ' KB]', response: reply,
      status: 'ok', durationMs: Date.now() - started
    });
    return { ok: true, reply: reply, savedUrl: savedUrl };
  } catch (e) {
    _logChat_({
      email: email, slug: slug, model: agent.model || '', skills: agent.skills || '',
      prompt: '[uploaded file: ' + filename + ']', response: '',
      status: 'error', error: e.message || String(e), durationMs: Date.now() - started
    });
    return { ok: false, error: e.message || String(e) };
  }
}

// One-time: adds Agents.allow_upload and Agents.upload_folder_id if missing.
// Non-destructive, same pattern as the other Agents migrations.
function migrateAgentsAddUploadConfig_() {
  if (!_isAdminUser_()) { _toast_('Admin only.'); return; }
  var sheet = _getSs_().getSheetByName('Agents');
  if (!sheet) { _toast_('No Agents tab found.'); return; }
  var lastCol = Math.max(sheet.getLastColumn(), 1);
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function (h) { return String(h || '').trim().toLowerCase(); });

  var toAdd = [];
  if (headers.indexOf('allow_upload') < 0) toAdd.push('allow_upload');
  if (headers.indexOf('upload_folder_id') < 0) toAdd.push('upload_folder_id');

  if (!toAdd.length) { _toast_('Agents already has allow_upload and upload_folder_id.'); return; }

  sheet.getRange(1, lastCol + 1, 1, toAdd.length).setValues([toAdd]);
  sheet.getRange(1, lastCol + 1, 1, toAdd.length).setFontWeight('bold');
  _toast_(
    'Added ' + toAdd.join(', ') + ' to Agents. To enable file upload for an agent (e.g. Ultimate DAST Analyzer): ' +
    'set allow_upload to "yes" and upload_folder_id to a Drive folder ID that agent\'s uploads should be archived into.'
  );
}
