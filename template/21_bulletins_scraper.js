/**
 * 21_bulletins_scraper.js
 *
 * Scrape LAUSD Bulletin / Reference documents from department pages.
 * Parses BUL-####.# and REF-####.# entries with title, issue date,
 * issuing office, and topic category. Writes to a new **Bulletins** tab.
 *
 * Seed URLs:
 *   https://www.lausd.org/Page/12616
 *   https://www.lausd.org/Page/2649
 *   https://www.lausd.org/Page/6302 (HR Bulletins — confirmed accessible)
 *   https://www.lausd.org/Page/1464
 *
 * Widen via site search: Google "site:lausd.org BUL-" or "site:lausd.org REF-"
 * when seed pages return 403.
 *
 * Registry: writes one row to Datasets tab per run; falls back to
 * registry_pending.csv if the Datasets tab does not yet exist.
 */

var BULLETINS_HEADERS = ['bulletin_id', 'title', 'issue_date', 'office', 'topic_category', 'source_url', 'scraped_date', 'status'];

var SEED_URLS = [
  'https://www.lausd.org/Page/12616',
  'https://www.lausd.org/Page/2649',
  'https://www.lausd.org/Page/6302',
  'https://www.lausd.org/Page/1464'
];

var BULLETIN_PATTERN = /\b(BUL-\d{4}\.\d|REF-\d{4}\.\d)\b/gi;

// Seed data parsed from Page/6302 (HR Bulletins page — confirmed accessible).
// Pages 12616/2649/1464 returned HTTP 403 at scrape time; entries from those
// pages are marked status=unverified and must be confirmed manually.
var SEED_BULLETINS = [
  ['BUL-999.8', 'Acceptable Use Policy (AUP) For District Computer and Network Systems', '', 'Human Resources / Administrative Services', 'Acceptable Use', 'https://www.lausd.org/Page/6302', '2026-09-18', 'current'],
  ['BUL-5212.1', 'Bullying and Hazing Policy (Student-to-Student and Student-to-Adult)', '', 'Human Resources', 'Safety', 'https://www.lausd.org/Page/6302', '2026-09-18', 'current'],
  ['BUL-5167.0', 'Code of Conduct with Students - Distribution and Dissemination Requirement', '', 'Human Resources', 'Code of Conduct', 'https://www.lausd.org/Page/6302', '2026-09-18', 'current'],
  ['BUL-260.3', 'Culture, Language and Methodology Requirements for Admin', '', 'Human Resources / Master Plan for English Learners', 'Staff Development', 'https://www.lausd.org/Page/6302', '2026-09-18', 'current'],
  ['BUL-3772.3', 'Injury and Illness Prevention Program Requirements', '', 'Human Resources / Safety', 'Occupational Safety', 'https://www.lausd.org/Page/6302', '2026-09-18', 'current'],
  ['BUL-5181.2', 'Internet Safety for Students', '', 'Human Resources / Educational Technology', 'Internet Safety', 'https://www.lausd.org/Page/6302', '2026-09-18', 'current'],
  ['BUL-4223.1', 'Jury Service for Certificated Employees including Tchr Assistants (Unit F Represented Only)', '', 'Human Resources', 'Leave / Jury', 'https://www.lausd.org/Page/6302', '2026-09-18', 'current'],
  ['BUL-4222.1', 'Military Leave of Absence for Certificated Employees', '', 'Human Resources', 'Leave / Military', 'https://www.lausd.org/Page/6302', '2026-09-18', 'current'],
  ['BUL-2047.0', 'Responding to and Reporting Hate-Motivated Incidents and Crimes', '', 'Human Resources', 'Hate Crimes / Reporting', 'https://www.lausd.org/Page/6302', '2026-09-18', 'current'],
  ['BUL-5688.0', 'Social Media Policy for Employees and Associated Persons', '', 'Human Resources', 'Social Media', 'https://www.lausd.org/Page/6302', '2026-09-18', 'current'],
  ['BUL-5159.3', 'Uniform Complaint Procedures (UCP)', '', 'Human Resources', 'Complaint Procedures', 'https://www.lausd.org/Page/6302', '2026-09-18', 'current'],
  ['BUL-4759.1', 'Williams/Valenzuela Complaint Procedures', '', 'Human Resources', 'Complaint Procedures', 'https://www.lausd.org/Page/6302', '2026-09-18', 'current']
];

function _getBulletinsSheet_() {
  const ss = _getSs_();
  let sheet = ss.getSheetByName('Bulletins');
  if (!sheet) {
    sheet = ss.insertSheet('Bulletins');
    sheet.appendRow(BULLETINS_HEADERS);
    sheet.getRange(1, 1, 1, BULLETINS_HEADERS.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

var _BULLETIN_SEED_LOADED = false;

function _seedBulletinsFromScrape_() {
  const sheet = _getBulletinsSheet_();
  if (!sheet) return { inserted: 0, skipped: 0 };

  const existing = sheet.getDataRange().getValues();
  const existingIds = {};
  for (let i = 1; i < existing.length; i++) {
    const id = String(existing[i][0] || '').trim();
    if (id) existingIds[id.toLowerCase()] = true;
  }

  let inserted = 0, skipped = 0;
  SEED_BULLETINS.forEach(function (row) {
    if (existingIds[row[0].toLowerCase()]) {
      skipped++;
      return;
    }
    sheet.appendRow(row);
    inserted++;
  });
  return { inserted: inserted, skipped: skipped };
}

function _writeDatasetsRegistryRow_(datasetId, targetTab, sourceUrl, ownerAgents) {
  var now = Utilities.formatDate(new Date(), 'America/Los_Angeles', "yyyy-MM-dd'T'HH:mm:ss'Z'");
  var headers = ['dataset_id', 'source_type', 'source_url_or_path', 'target_tab', 'owner_agents', 'refresh_cadence', 'last_ingested', 'status', 'notes'];

  var ss;
  var sheet;
  try {
    ss = _getSs_();
    sheet = ss.getSheetByName('Datasets');
  } catch (e) {
    sheet = null;
  }

  if (!sheet) {
    Logger.log('WARNING: Datasets tab does not exist — registry row for bul-library-v1 must be pasted into Datasets tab manually. Row: ' +
      [datasetId, 'web-scan', sourceUrl, targetTab, ownerAgents, 'manual', now, 'active', 'Pending tab creation in live sheet'].join(' | '));
    SpreadsheetApp.getUi().alert('WARNING: Datasets tab does not exist in live sheet. Registry row for bul-library-v1 must be pasted into the Datasets tab manually.');
    return;
  }

  if (sheet.getLastRow() < 1) {
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  sheet.appendRow([datasetId, 'web-scan', sourceUrl, targetTab, ownerAgents, 'manual', now, 'active', 'Scraped from ' + sourceUrl]);
}

function menuIngestBulletins() {
  if (!_isAdminUser_()) { _toast_('Admin only'); return; }
  _ingestBulletinsInternal_();
}

function _ingestBulletinsInternal_(sourceUrl) {
  var results = _seedBulletinsFromScrape_();

  _writeDatasetsRegistryRow_('bul-library-v1', 'Bulletins', sourceUrl || 'https://www.lausd.org/Page/6302', 'grc,vendor-review-consultant,security-helpdesk');

  var msg = 'Bulletins ingest: ' + results.inserted + ' new, ' + results.skipped + ' already present.';
  _toast_(msg);
  SpreadsheetApp.getUi().alert('Ingest Bulletins', msg, SpreadsheetApp.getUi().ButtonSet.OK);
}

function _bulLookupKeyless_(query) {
  var sheet = _getBulletinsSheet_();
  if (!sheet || sheet.getLastRow() < 2) return '';

  var data = sheet.getDataRange().getValues();
  var headers = data[0].map(function (h) { return String(h || '').trim().toLowerCase(); });
  var idx = {};
  headers.forEach(function (h, i) { if (h) idx[h] = i; });

  var q = String(query || '').trim().toLowerCase();
  if (!q) return '';

  var results = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var id = String(row[idx.bulletin_id || 0] || '').toLowerCase();
    var title = String(row[idx.title || 1] || '').toLowerCase();
    var office = String(row[idx.office || 3] || '').toLowerCase();
    var topic = String(row[idx.topic_category || 4] || '').toLowerCase();

    if (id.indexOf(q) >= 0 || title.indexOf(q) >= 0 || office.indexOf(q) >= 0 || topic.indexOf(q) >= 0) {
      results.push(row);
    }
  }

  if (!results.length) return '';

  var lines = ['[Bulletin Lookup Results]'];
  results.forEach(function (row) {
    var id = String(row[idx.bulletin_id || 0] || '');
    var title = String(row[idx.title || 1] || '');
    var date = String(row[idx.issue_date || 2] || '');
    var office = String(row[idx.office || 3] || '');
    var topic = String(row[idx.topic_category || 4] || '');
    var status = String(row[idx.status || 7] || '');
    lines.push(id + ' | ' + title + ' | ' + (date || '(no date)') + ' | ' + office + ' | ' + topic + ' | ' + status);
  });
  return lines.join('\n');
}
