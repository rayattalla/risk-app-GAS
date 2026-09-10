/**
 * RiskAI-GAS — Reports (07_reports.js)
 *
 * Google Docs report generation with two patterns:
 *
 * PATTERN A — Template-first with code fallback
 *   Store a Doc template ID in the Config sheet. The system copies it, fills
 *   {TOKEN} placeholders throughout, and saves to the project folder.
 *   If the template ID is missing or the copy fails, a fallback function
 *   generates the Doc from code. Admins can customize templates without
 *   touching code; code never silently fails.
 *
 * PATTERN B — Heading-based section updates
 *   Find a named heading in an existing Doc, then update the paragraph,
 *   table cell, or list item immediately after it. Used for incremental
 *   report updates (status changes, timeline entries, etc.) without
 *   rewriting the whole document.
 *
 * Extracted from: incident-reporting-nist-aligned (IRS v3.4 reportGenerator.js
 *   + reportUpdater.js) and gas-arb-new (generatePhaseReport).
 *
 * Public API:
 *   createReport(options)                         — template-first + fallback
 *   populateDocTokens(doc, tokens)               — replace {TOKEN} everywhere in a Doc
 *   updateDocSection(doc, headingText, text)     — replace paragraph after heading
 *   updateDocTableCell(doc, heading, row, col, v) — update table cell after heading
 *   appendToDocSection(doc, headingText, text)   — append a line after heading
 *   appendUpdateHistory(doc, message)            — add timestamped entry to history
 */

// ==========================================================================
// TEMPLATE-FIRST REPORT CREATION
// ==========================================================================

/**
 * Create a new Google Doc report.
 *
 * Options:
 *   docName      (required) — name for the new Doc
 *   folderId     (optional) — Drive folder ID; defaults to root folder
 *   templateDocId (optional) — source Doc ID to copy; if blank, uses fallbackFn
 *   tokens       (optional) — {TOKEN: value} map for {TOKEN} substitution
 *   fallbackFn   (optional) — function(folder, docName, tokens) → {docId, docUrl}
 *                             Called when templateDocId is missing or copy fails
 *
 * Returns: { success, docId, docUrl, method: 'template'|'fallback', error? }
 *
 * Example:
 *   const result = createReport({
 *     docName:      'INC-2026-001 — CSIRT Report',
 *     folderId:     CONFIG.ROOT_FOLDER_ID,
 *     templateDocId: _getConfigValue_('CSIRT_TEMPLATE_DOC_ID'),
 *     tokens:       { INCIDENT_ID: 'INC-2026-001', STATUS: 'Closed' },
 *     fallbackFn:   (folder, name, tokens) => _buildCsirtReportFallback_(folder, name, tokens)
 *   });
 */
function createReport(options) {
  options = options || {};
  if (!options.docName) return { success: false, error: 'docName is required' };

  let folder;
  try {
    folder = options.folderId
      ? DriveApp.getFolderById(options.folderId)
      : _getOrCreateRootFolder_();
  } catch (e) {
    folder = _getOrCreateRootFolder_();
  }

  // Attempt 1: template copy
  if (options.templateDocId) {
    try {
      const copy = DriveApp.getFileById(options.templateDocId).makeCopy(options.docName, folder);
      const doc  = DocumentApp.openById(copy.getId());
      if (options.tokens) populateDocTokens(doc, options.tokens);
      doc.saveAndClose();
      auditLog_(null, 'REPORT_CREATED', options.docName + ' (template)');
      return { success: true, docId: copy.getId(), docUrl: copy.getUrl(), method: 'template' };
    } catch (e) {
      Logger.log('createReport template failed: ' + e.message + ' — trying fallback');
    }
  }

  // Attempt 2: fallback generator
  if (typeof options.fallbackFn === 'function') {
    try {
      const result = options.fallbackFn(folder, options.docName, options.tokens || {});
      auditLog_(null, 'REPORT_CREATED', options.docName + ' (fallback)');
      return { success: true, method: 'fallback', docId: result.docId, docUrl: result.docUrl };
    } catch (e) {
      Logger.log('createReport fallback failed: ' + e.message);
      return { success: false, error: 'Template and fallback both failed: ' + e.message };
    }
  }

  return { success: false, error: 'No templateDocId and no fallbackFn provided' };
}

// ==========================================================================
// TOKEN SUBSTITUTION — fill {TOKEN} placeholders throughout a Doc
// ==========================================================================

/**
 * Replace all {TOKEN} placeholders in a Google Doc body.
 * Operates on the entire body text — headers, paragraphs, tables, list items.
 *
 * doc:    DocumentApp document object (already open)
 * tokens: { TOKEN_NAME: 'value', ... }
 *
 * Example:
 *   populateDocTokens(doc, { INCIDENT_ID: 'INC-001', STATUS: 'Open', DATE: '2026-06-26' });
 */
function populateDocTokens(doc, tokens) {
  const body = doc.getBody();
  Object.keys(tokens || {}).forEach(key => {
    const val = String(tokens[key] != null ? tokens[key] : '');
    body.replaceText('\\{' + key + '\\}', val);
  });
}

// ==========================================================================
// HEADING-BASED SECTION NAVIGATION
// ==========================================================================

/**
 * Replace the paragraph immediately following a heading in a Doc.
 * Finds the first paragraph whose text matches `headingText` (case-sensitive).
 * If the next element is a paragraph, replaces its text.
 * If nothing follows the heading, inserts a new paragraph.
 *
 * doc:         DocumentApp document object
 * headingText: exact text of the heading to locate
 * text:        replacement text for the paragraph after the heading
 *
 * Returns true if the heading was found, false otherwise.
 */
function updateDocSection(doc, headingText, text) {
  const body        = doc.getBody();
  const searchResult = body.findText(headingText);
  if (!searchResult) return false;

  const headingEl  = searchResult.getElement().getParent();
  const headingIdx = body.getChildIndex(headingEl);

  if (headingIdx + 1 < body.getNumChildren()) {
    const next = body.getChild(headingIdx + 1);
    if (next.getType() === DocumentApp.ElementType.PARAGRAPH) {
      next.asParagraph().setText(text);
      return true;
    }
  }
  // Nothing follows the heading — insert a paragraph
  body.insertParagraph(headingIdx + 1, text);
  return true;
}

/**
 * Append a line of text after a heading (does not replace existing content).
 * Inserts a new paragraph immediately after the heading paragraph.
 */
function appendToDocSection(doc, headingText, text) {
  const body         = doc.getBody();
  const searchResult = body.findText(headingText);
  if (!searchResult) return false;

  const headingEl  = searchResult.getElement().getParent();
  const headingIdx = body.getChildIndex(headingEl);
  body.insertParagraph(headingIdx + 1, text);
  return true;
}

/**
 * Find the first table after a heading and update a specific cell.
 *
 * doc:         DocumentApp document object
 * headingText: exact heading text to locate (searches forward from it)
 * tableRow:    0-based row index in the table
 * tableCol:    0-based column index in the table
 * value:       text to write into that cell
 *
 * Returns true if the cell was found and updated, false otherwise.
 */
function updateDocTableCell(doc, headingText, tableRow, tableCol, value) {
  const body         = doc.getBody();
  const searchResult = body.findText(headingText);
  if (!searchResult) return false;

  const headingEl  = searchResult.getElement().getParent();
  let idx          = body.getChildIndex(headingEl);

  while (++idx < body.getNumChildren()) {
    const child = body.getChild(idx);
    if (child.getType() === DocumentApp.ElementType.TABLE) {
      const table = child.asTable();
      if (tableRow < table.getNumRows()) {
        const row = table.getRow(tableRow);
        if (tableCol < row.getNumCells()) {
          row.getCell(tableCol).setText(String(value != null ? value : ''));
          return true;
        }
      }
      return false; // found a table but indexes out of range
    }
    // Stop searching at the next heading
    if (child.getType() === DocumentApp.ElementType.PARAGRAPH) {
      const heading = child.asParagraph().getHeading();
      if (heading !== DocumentApp.ParagraphHeading.NORMAL) break;
    }
  }
  return false;
}

// ==========================================================================
// UPDATE HISTORY — append timestamped entries to a named section
// ==========================================================================

/**
 * Append a timestamped entry to an "Update History" section in a Doc.
 * Inserts a new paragraph after the "Update History" heading.
 *
 * doc:     DocumentApp document object
 * message: description of the change
 * actor:   who made the change (defaults to active user email)
 *
 * Example output appended to doc:
 *   "2026-06-26 10:30:00 — remon.attalla@lausd.net: Status changed to Resolved"
 */
function appendUpdateHistory(doc, message, actor) {
  actor = actor || _getUserEmail_() || 'system';
  const entry = _now_() + ' — ' + actor + ': ' + String(message || '');
  appendToDocSection(doc, 'Update History', entry);
}

// ==========================================================================
// SIMPLE CODE-GENERATED REPORT — no template needed
// ==========================================================================

/**
 * Create a minimal Google Doc report from key-value data.
 * Use this as the `fallbackFn` in createReport(), or stand-alone.
 *
 * folder:  Drive folder object
 * docName: name for the new Doc
 * sections: [{ heading: 'Title', rows: [['Label', 'Value'], ...] }, ...]
 *
 * Returns { docId, docUrl }
 *
 * Example:
 *   const result = createSimpleReport(folder, 'INC-001 Summary', [
 *     { heading: 'Incident Details',
 *       rows: [['ID', 'INC-001'], ['Status', 'Open'], ['Severity', 'RED']] },
 *     { heading: 'Update History', rows: [] }
 *   ]);
 */
function createSimpleReport(folder, docName, sections) {
  const doc  = DocumentApp.create(docName);
  const body = doc.getBody();
  body.clear();

  (sections || []).forEach(function(section, i) {
    if (i > 0) body.appendParagraph('');
    const heading = body.appendParagraph(section.heading || '');
    heading.setHeading(DocumentApp.ParagraphHeading.HEADING2);

    if (section.rows && section.rows.length) {
      const table = body.appendTable();
      section.rows.forEach(function(row) {
        const tr = table.appendTableRow();
        row.forEach(function(cell) { tr.appendTableCell(String(cell != null ? cell : '')); });
      });
    }
  });

  doc.saveAndClose();

  // Move to folder
  const file = DriveApp.getFileById(doc.getId());
  folder.addFile(file);
  try { DriveApp.getRootFolder().removeFile(file); } catch (e) {}

  return { docId: doc.getId(), docUrl: doc.getUrl() };
}
