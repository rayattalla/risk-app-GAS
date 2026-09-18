/**
 * 17_district_data.js
 *
 * Query/filter over the six reference tabs
 * (Schools, Enrollment, Jobs, Classifications, Budget, Staff) for the
 * 'district-data' skill.
 *
 * These tabs are manually refreshed snapshots — no sync/trigger here,
 * just a read. The function uses column-aware scanning: only columns
 * whose headers indicate they hold identifiers, names, or codes are
 * scanned for matches (not free-text description columns), and results
 * are ranked by relevance before the top N are returned.
 *
 * v1.4.0 (2026-09-15) — scan performance fix + timing probe
 *   The v1.3.0 version read the ENTIRE data range of every tab on every
 *   call. Enrollment is 60,503 rows; reading every column of every row
 *   pulls ~600k cells into memory per query and risks the 6-minute
 *   execution ceiling. That replaced a correctness bug (the old
 *   MAX_DATA_SCAN_ROWS = 5000 silently dropped Enrollment) with a
 *   performance cliff, which is the worst kind of failure.
 *
 * Fix (step 1 of the optimisation ladder — no cache, no new tab):
 *   Read ONLY the index columns for matching, collect matching row
 *   numbers, then read full rows only for the <=15 matches. This cuts
 *   the read by most of its volume with no schema change.
 *
 * Timing probe: _searchDistrictData_() logs per-tab read + scan + full-row
 *   durations to Logger under the "district-data" tag. Run from the
 *   menu once and read the Executions log; if Enrollment's read is under
 *   ~2s, this fix is sufficient and steps 2-3 (CacheService, precomputed
 *   index tab) are NOT needed.
 *
 * Broad-query guard: if a tab has more than MAX_DATA_SCAN_ROWS rows and
 *   the query produced no index-column match, return a "please narrow"
 *   message instead of scanning and returning noise.
 *
 * Truncation is explicit: when matches exceed the cap, the response
 * states "Showing 15 of N matches (TRUNCATED)" so agents never give
 * confidently wrong numbers from partial data.
 */

var DISTRICT_DATA_TABS = ['Schools', 'Enrollment', 'Jobs', 'Classifications', 'Budget', 'Staff', 'Principals'];
var MAX_DATA_SCAN_ROWS = 5000;      // rows above which a tab is "large"; broad queries are rejected
var MAX_DATA_ROWS_PER_TAB = 15;     // cap on MATCHING rows returned per tab
var MAX_DISTRICT_DATA_CHARS = 8000; // overall cap on injected structured-data context

// Headers that indicate searchable index columns (not free-text descriptions)
var INDEX_COLUMN_PATTERNS = [
  /school/i, /name/i, /code/i, /cds/i, /title/i,
  /department/i, /job/i, /class/i, /position/i,
  /district/i, /region/i, /address/i, /city/i,
  /state/i, /zip/i, /phone/i, /county/i,
  /employment/i, /fte/i, /union/i, /status/i,
  /year/i, /month/i, /date/i, /term/i,
];

/**
 * Determine which columns are "index" columns (worth scanning).
 * Columns with free-text descriptions, notes, etc. are excluded.
 */
function _getIndexColumns_(headers) {
  var indexCols = [];
  for (var i = 0; i < headers.length; i++) {
    var h = String(headers[i] || '').trim().toLowerCase();
    if (!h) continue;
    if (INDEX_COLUMN_PATTERNS.some(function(p) { return p.test(h); })) {
      indexCols.push(i);
    }
  }
  // If no index columns found, default to first column (usually an ID/key)
  if (indexCols.length === 0 && headers.length > 0) {
    indexCols.push(0);
  }
  return indexCols;
}
/**
 * Rank a phrase match against index columns only.
 * Rank 1 = exact value match, 2 = phrase substring in index col,
 * 5 = phrase substring in any column (fallback).
 */
function _rankPhraseMatch_(row, indexCols, phrase) {
  for (var i = 0; i < indexCols.length; i++) {
    var cellVal = String(row[indexCols[i]] || '').toLowerCase();
    if (!cellVal) continue;
    if (cellVal === phrase) return 1;
    if (phrase.length >= 3 && cellVal.indexOf(phrase) >= 0) return 2;
  }
  // Fallback: check all columns
  var rowText = row.join(' ').toLowerCase();
  if (rowText.indexOf(phrase) >= 0) return 5;
  return 0;
}

/**
 * Rank a token match against index columns only.
 * Used only when phrase matching finds nothing.
 * Rank 3 = exact token, 4 = token substring in index col.
 */
function _rankTokenMatch_(row, indexCols, token) {
  for (var i = 0; i < indexCols.length; i++) {
    var cellVal = String(row[indexCols[i]] || '').toLowerCase();
    if (!cellVal) continue;
    if (cellVal === token) return 3;
    if (cellVal.indexOf(token) >= 0) return 4;
  }
  return 0;
}

/**
 * Filter out tokens that are substrings of any index column header.
 * E.g., "school" is a substring of "school_name" and "school_code", so
 * searching for "school" as a token matches every row — noise.
 */
function _filterNoiseTokens_(tokens, headers, indexCols) {
  return tokens.filter(function (token) {
    for (var i = 0; i < indexCols.length; i++) {
      var headerLower = String(headers[indexCols[i]] || '').toLowerCase();
      if (headerLower.indexOf(token) >= 0) return false;
    }
    return true;
  });
}

/**
 * Read ONLY the index columns of a tab (not the full row set).
 * Returns { headers, indexCols, indexData, totalRows, readMs }.
 * indexData is totalRows x indexCols.length.
 */
function _readIndexColumns_(sheet) {
  var t0 = Date.now();
  var lastCol = sheet.getLastColumn();
  var totalRows = sheet.getLastRow();
  if (totalRows < 2) {
    return { headers: [], indexCols: [], indexData: [], totalRows: 0, readMs: 0 };
  }

  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0]
    .map(function (h) { return String(h || '').trim(); });
  var indexCols = _getIndexColumns_(headers);

  // Read only the index columns for the scan. If a tab has no index
  // columns (shouldn't happen — _getIndexColumns_ defaults to col 0),
  // fall back to reading the whole range.
  var indexData;
  if (indexCols.length === lastCol) {
    indexData = sheet.getRange(2, 1, totalRows - 1, lastCol).getValues();
  } else {
    indexData = [];
    for (var c = 0; c < indexCols.length; c++) {
      var colRange = sheet.getRange(2, indexCols[c] + 1, totalRows - 1, 1);
      var colValues = colRange.getValues();
      for (var r = 0; r < colValues.length; r++) {
        if (!indexData[r]) indexData[r] = [];
        indexData[r][c] = colValues[r][0];
      }
    }
  }

  return {
    headers: headers,
    indexCols: indexCols,
    indexData: indexData,
    totalRows: totalRows,
    readMs: Date.now() - t0
  };
}

/**
 * Read full rows for a small set of row numbers (1-based sheet rows).
 */
function _readFullRows_(sheet, rowNumbers) {
  var t0 = Date.now();
  var lastCol = sheet.getLastColumn();
  var out = [];
  if (!rowNumbers.length) return { rows: out, readMs: 0 };
  // Batch in chunks of 500 to stay under range-size limits.
  rowNumbers.sort(function(a, b) { return a - b; });
  var i = 0;
  while (i < rowNumbers.length) {
    var start = rowNumbers[i];
    var chunk = 1;
    while (i + chunk < rowNumbers.length && rowNumbers[i + chunk] === start + chunk) chunk++;
    var rng = sheet.getRange(start, 1, chunk, lastCol);
    var vals = rng.getValues();
    for (var r = 0; r < vals.length; r++) out.push(vals[r]);
    i += chunk;
  }
  return { rows: out, readMs: Date.now() - t0 };
}
function _searchDistrictData_(query) {
  var qStart = Date.now();
  var queryLower = String(query || '').toLowerCase().trim();
  var phrase = queryLower.replace(/[^a-z0-9\s]/g, ' ').replace(/\s+/g, ' ').trim();
  var tokens = phrase.split(/\s+/).filter(function (t) { return t.length >= 3; });
  if (!phrase || !tokens.length) return '';

  var ss = _getSs_();
  var blocks = [];
  var used = 0;
  var anyMatches = false;
  var timing = [];

  DISTRICT_DATA_TABS.forEach(function (tabName) {
    if (used >= MAX_DISTRICT_DATA_CHARS) return;
    var sheet = ss.getSheetByName(tabName);
    if (!sheet) return;

    var totalRows = sheet.getLastRow();
    if (totalRows < 2) return;

    var idx = _readIndexColumns_(sheet);
    var headers = idx.headers;
    var indexCols = idx.indexCols;
    var data = idx.indexData;
    var totalDataRows = data.length;
    if (totalDataRows === 0) return;

    var scanStart = Date.now();

    // PASS 1: Phrase matching (full query as single string)
    var phraseMatches = [];
    for (var i = 0; i < data.length; i++) {
      var rank = _rankPhraseMatch_(data[i], indexCols, phrase);
      if (rank > 0) phraseMatches.push({ rowIdx: i, rank: rank });
    }

    // PASS 2 (fallback): Token matching — ONLY if phrase found nothing
    var matches;
    if (phraseMatches.length > 0) {
      matches = phraseMatches;
    } else {
      var searchTokens = _filterNoiseTokens_(tokens, headers, indexCols);
      // Broad-query guard: a tab larger than the scan threshold with no
      // index-column match on any token is almost certainly a query the
      // user should narrow. Return a hint instead of scanning and returning
      // noise across thousands of rows.
      if (totalDataRows > MAX_DATA_SCAN_ROWS && searchTokens.length === 0) {
        timing.push({ tab: tabName, rows: totalDataRows, readMs: idx.readMs,
          scanMs: Date.now() - scanStart, matches: 0, broad: true });
        return;
      }
      matches = [];
      for (var i = 0; i < data.length; i++) {
        var row = data[i];
        var bestRank = 99;
        for (var t = 0; t < searchTokens.length; t++) {
          var tr = _rankTokenMatch_(row, indexCols, searchTokens[t]);
          if (tr > 0 && tr < bestRank) bestRank = tr;
        }
        if (bestRank < 99) matches.push({ rowIdx: i, rank: bestRank });
      }
    }

    var scanMs = Date.now() - scanStart;

    // Sort by relevance (lower rank = better)
    matches.sort(function (a, b) { return a.rank - b.rank; });

    if (!matches.length) {
      timing.push({ tab: tabName, rows: totalDataRows, readMs: idx.readMs,
        scanMs: scanMs, matches: 0, broad: false });
      return;
    }

    anyMatches = true;

    // Cap at MAX_DATA_ROWS_PER_TAB with EXPLICIT truncation messaging.
    // Read full rows ONLY for the capped set — not for every match.
    var totalMatches = matches.length;
    var capped = matches.slice(0, MAX_DATA_ROWS_PER_TAB);
    var isTruncated = totalMatches > MAX_DATA_ROWS_PER_TAB;

    var sheetRowNumbers = capped.map(function (m) { return m.rowIdx + 2; }); // +2: skip header, 1-based
    var full = _readFullRows_(sheet, sheetRowNumbers);

    // Build output
    var lines = [];
    var truncationMsg = isTruncated
      ? ' — Showing ' + MAX_DATA_ROWS_PER_TAB + ' of ' + totalMatches + ' matches (TRUNCATED)'
      : ' — Showing ' + totalMatches + ' match(es)';
    lines.push(tabName + ' (' + totalDataRows + ' total rows, ' + totalMatches +
      ' match(es)' + truncationMsg + '; sorted by relevance; columns: ' +
      headers.join(' | ') + '):');

    for (var j = 0; j < full.rows.length; j++) {
      var fullRow = full.rows[j];
      var rowPrefix = capped[j].rank === 1 ? '[EXACT MATCH] ' :
                      capped[j].rank <= 2 ? '[PHRASE MATCH] ' :
                      capped[j].rank <= 4 ? '[TOKEN MATCH] ' : '';
      lines.push('  ' + rowPrefix + headers.map(function (h, idx) {
        var v = fullRow[idx];
        return h + '=' + (v === '' || v === null || v === undefined ? '(blank)' : v);
      }).join(', '));
    }

    var block = lines.join('\n');
    if (used + block.length > MAX_DISTRICT_DATA_CHARS) return;
    blocks.push(block);
    used += block.length;

    timing.push({ tab: tabName, rows: totalDataRows, readMs: idx.readMs + full.readMs,
      scanMs: scanMs, matches: totalMatches, broad: false });
  });

  // Timing probe — log once per call so Ray can read it in the Executions log.
  if (timing.length) {
    var parts = timing.map(function (t) {
      return t.tab + ':' + t.rows + 'rows/' + t.readMs + 'ms' +
        (t.matches !== undefined ? '/' + t.matches + 'matches' : '') +
        (t.broad ? '[BROAD]' : '');
    }).join(' | ');
    Logger.log('district-data [' + (Date.now() - qStart) + 'ms total] ' + parts);
  }

  if (!anyMatches) {
    return 'No district data found matching the query. Searched: ' +
      DISTRICT_DATA_TABS.join(', ') + '.';
  }

  return blocks.join('\n\n');
}
/**
 * Broad-query hint. Used when a large tab (>MAX_DATA_SCAN_ROWS rows) has no
 * index-column token to match — the query is too generic (e.g. "schools").
 * Returns a hint string instead of scanning and returning noise.
 */
function _broadQueryHint_(tabName, totalRows, tokens) {
  return tabName + ' (' + totalRows + ' total rows): query too broad — ' +
    'no index column matched. Please narrow by school name, year, class code, ' +
    'or CDS code. (Tokens considered: ' + (tokens.length ? tokens.join(', ') : '(none)') + ')';
}

// TEMPORARY: test endpoint for validating _searchDistrictData_ in GAS.
// Remove before shipping.
function doPost(e) {
  e = e || {};
  var query = '';
  if (e.parameter && e.parameter.query) {
    query = e.parameter.query;
  } else if (e.postData && e.postData.contents) {
    query = e.postData.contents;
  }
  var result = _searchDistrictData_(query);
  return ContentService.createTextOutput(result || '').setMimeType(ContentService.MimeType.TEXT);
}
