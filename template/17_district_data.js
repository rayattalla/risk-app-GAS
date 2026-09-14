/**
 * 17_district_data.js
 *
 * Keyword search over the six manually-refreshed public LAUSD reference
 * tabs (Schools, Enrollment, Jobs, Classifications, Budget, Staff), for the
 * 'district-data' skill. These are one-time snapshots, refreshed by hand as
 * needed -- there is deliberately no sync/trigger here, just a read.
 *
 * Schema-agnostic on purpose: every column of every matching row is read
 * fresh from the header row each call rather than hardcoded, so this keeps
 * working whatever the actual columns turn out to be (or become, if a tab
 * is later replaced with a refreshed export that has slightly different
 * columns). It is a keyword substring match, not a real query engine --
 * enough for "does any row mention this school/CDS code/class code," not
 * built for aggregate math across rows.
 */

var DISTRICT_DATA_TABS = ['Schools', 'Enrollment', 'Jobs', 'Classifications', 'Budget', 'Staff'];
var MAX_DATA_SCAN_ROWS = 5000;     // safety cap on rows read per tab, regardless of match count
var MAX_DATA_ROWS_PER_TAB = 15;    // cap on MATCHING rows actually injected per tab
var MAX_DISTRICT_DATA_CHARS = 8000; // overall cap on injected structured-data context

function _searchDistrictData_(query) {
  var tokens = String(query || '')
    .toLowerCase()
    .replace(/[^a-z0-9\s]/g, ' ')
    .split(/\s+/)
    .filter(function (t) { return t.length >= 3; }); // skip tiny/common words
  if (!tokens.length) return '';

  var ss = _getSs_();
  var blocks = [];
  var used = 0;

  DISTRICT_DATA_TABS.forEach(function (tabName) {
    if (used >= MAX_DISTRICT_DATA_CHARS) return;
    var sheet = ss.getSheetByName(tabName);
    if (!sheet || sheet.getLastRow() < 2) return;

    var lastCol = sheet.getLastColumn();
    var lastRow = Math.min(sheet.getLastRow(), MAX_DATA_SCAN_ROWS + 1);
    var data = sheet.getRange(1, 1, lastRow, lastCol).getValues();
    var headers = data[0].map(function (h) { return String(h || '').trim(); });

    var matches = [];
    for (var i = 1; i < data.length && matches.length < MAX_DATA_ROWS_PER_TAB; i++) {
      var row = data[i];
      var rowText = row.join(' ').toLowerCase();
      var isMatch = tokens.some(function (t) { return rowText.indexOf(t) >= 0; });
      if (isMatch) matches.push(row);
    }
    if (!matches.length) return;

    var lines = [tabName + ' (' + matches.length + ' matching row(s) of ' + (lastRow - 1) + ' scanned; columns: ' + headers.join(' | ') + '):'];
    matches.forEach(function (row) {
      lines.push(headers.map(function (h, idx) {
        var v = row[idx];
        return h + '=' + (v === '' || v === null || v === undefined ? '(blank)' : v);
      }).join(', '));
    });
    var block = lines.join('\n');
    if (used + block.length > MAX_DISTRICT_DATA_CHARS) return;
    blocks.push(block);
    used += block.length;
  });

  return blocks.join('\n\n');
}
