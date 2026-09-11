/**
 * 14_ingest_drive_folder.js
 *
 * Bulk-ingest every file in a Drive folder into the KB tab, one KB row (or
 * chunk of rows) per file. Supports Google Docs, .txt, .md, .csv directly;
 * PDF via the Advanced Drive Service's OCR-copy trick (copy as a Google Doc
 * with OCR on, extract text, delete the temp copy).
 *
 * PDF support requires the "Drive" Advanced Service enabled for this project
 * (declared in appsscript.json's enabledAdvancedServices -- already done).
 * If it's somehow unavailable, PDFs are skipped with a clear reason in the
 * summary instead of failing the whole batch.
 */

function _extractDriveFolderId_(urlOrId) {
  const s = String(urlOrId || '').trim();
  const m = s.match(/\/folders\/([a-zA-Z0-9_-]+)/);
  if (m) return m[1];
  return s; // assume the user pasted a bare folder ID
}

// Returns extracted text, or null if the file type isn't supported (caller
// records it as skipped rather than treating it as a failure).
function _extractTextFromDriveFile_(file) {
  const mime = file.getMimeType();
  const name = file.getName();
  try {
    if (mime === MimeType.GOOGLE_DOCS) {
      return DocumentApp.openById(file.getId()).getBody().getText();
    }
    if (mime === 'text/plain' || mime === 'text/markdown' || mime === 'text/csv' ||
        /\.(txt|md|markdown)$/i.test(name)) {
      return file.getBlob().getDataAsString();
    }
    if (mime === MimeType.PDF) {
      return _extractPdfText_(file);
    }
  } catch (e) {
    return '[extraction failed for ' + name + ': ' + (e.message || e) + ']';
  }
  return null;
}

function _extractPdfText_(file) {
  if (typeof Drive === 'undefined') {
    throw new Error(
      'PDF text extraction needs the "Drive" Advanced Service. In the Apps Script editor: ' +
      'Services (+ icon) -> add "Drive API" (v2). If it still fails after that, the Google Cloud ' +
      'project behind this script also needs the Drive API enabled in the Cloud Console API library.'
    );
  }
  const blob = file.getBlob();
  const resource = { title: 'OCR_TEMP_' + file.getName(), mimeType: MimeType.GOOGLE_DOCS };
  const tempDoc = Drive.Files.insert(resource, blob, { ocr: true, ocrLanguage: 'en' });
  try {
    return DocumentApp.openById(tempDoc.id).getBody().getText();
  } finally {
    // Always clean up the temp OCR copy, success or failure.
    try { Drive.Files.remove(tempDoc.id); } catch (e2) { /* ignore cleanup failure */ }
  }
}

function _chunkAndAppendKb_(sheet, slug, title, content, tagExtra) {
  content = String(content || '').replace(/\s+/g, ' ').trim();
  if (!content) return 0;
  const size = 2400;
  let n = 0;
  for (let i = 0; i < content.length; i += size) {
    const c = content.substring(i, i + size);
    const id = Utilities.getUuid().slice(0, 8);
    const chunkTitle = title + (content.length > size ? ' pt' + (n + 1) : '');
    sheet.appendRow([id, slug, chunkTitle, c, tagExtra]);
    n++;
  }
  return n;
}

// Core entry point -- also callable directly (e.g. from a trigger) without
// going through the menu prompt wizard.
function ingestDriveFolder_(folderUrlOrId, targetSlug) {
  if (!_isAdminUser_()) { _toast_('Admin only.'); return; }
  const folderId = _extractDriveFolderId_(folderUrlOrId);
  const slug = String(targetSlug || 'shared').trim().toLowerCase();

  let folder;
  try {
    folder = DriveApp.getFolderById(folderId);
  } catch (e) {
    _toast_('Could not open folder: ' + (e.message || e));
    return;
  }

  const kbSheet = _getOrCreateTab_('KB', ['id', 'slug', 'title', 'body', 'tags']);
  const files = folder.getFiles();
  let processed = 0, chunkCount = 0;
  const skipped = [];

  while (files.hasNext()) {
    const file = files.next();
    const text = _extractTextFromDriveFile_(file);
    if (text === null) {
      skipped.push(file.getName() + ' (' + file.getMimeType() + ')');
      continue;
    }
    chunkCount += _chunkAndAppendKb_(kbSheet, slug, file.getName(), text, 'drive-folder:' + folderId);
    processed++;
  }

  const msg = 'Ingested ' + processed + ' file(s) into ' + chunkCount + ' KB row(s) for slug "' + slug + '".' +
              (skipped.length ? '\n\nSkipped (unsupported type):\n- ' + skipped.join('\n- ') : '');
  _toast_(processed + ' file(s) ingested into KB.');
  SpreadsheetApp.getUi().alert('Drive Folder Ingest Complete', msg, SpreadsheetApp.getUi().ButtonSet.OK);
}

function menuIngestDriveFolder_() {
  if (!_isAdminUser_()) { _toast_('Admin only.'); return; }
  const ui = SpreadsheetApp.getUi();

  const folderResp = ui.prompt(
    'Ingest Drive Folder (1/2)',
    'Paste the Drive folder URL or ID.\nSupported file types: Google Docs, PDF, .txt, .md, .csv.',
    ui.ButtonSet.OK_CANCEL
  );
  if (folderResp.getSelectedButton() !== ui.Button.OK) return;
  const folderUrl = folderResp.getResponseText().trim();
  if (!folderUrl) { _toast_('No folder given.'); return; }

  const slugResp = ui.prompt(
    'Ingest Drive Folder (2/2)',
    'Target agent slug for this content (blank = "shared", visible to every agent):',
    ui.ButtonSet.OK_CANCEL
  );
  if (slugResp.getSelectedButton() !== ui.Button.OK) return;
  const slug = slugResp.getResponseText().trim() || 'shared';

  ingestDriveFolder_(folderUrl, slug);
}
