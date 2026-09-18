/**
 * 22_procurements_scraper.js
 *
 * Scrape LAUSD Professional Services RFP / Solicitations data.
 * Sources:
 *   https://www.lausd.org/Page/19831 (Professional Services RFP Repository)
 *   https://www.lausd.org/Page/19507 (Solicitations landing — may list open bids)
 *
 * Writes to a new **Procurements** tab: rfp_number, description, status,
 * awarded_vendor, source_url, scraped_date.
 *
 * Registry: writes one row to Datasets tab per run; falls back to
 * registry_pending.csv if the Datasets tab does not yet exist.
 */

var PROCUREMENTS_HEADERS = ['rfp_number', 'description', 'status', 'awarded_vendor', 'source_url', 'scraped_date'];

var PROCUREMENT_SEED_URLS = [
  'https://www.lausd.org/Page/19831',
  'https://www.lausd.org/Page/19507'
];

function _getProcurementsSheet_() {
  const ss = _getSs_();
  let sheet = ss.getSheetByName('Procurements');
  if (!sheet) {
    sheet = ss.insertSheet('Procurements');
    sheet.appendRow(PROCUREMENTS_HEADERS);
    sheet.getRange(1, 1, 1, PROCUREMENTS_HEADERS.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
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
    Logger.log('WARNING: Datasets tab does not exist — registry row for rfp-repository-v1 must be pasted into Datasets tab manually. Row: ' +
      [datasetId, 'web-scan', sourceUrl, targetTab, ownerAgents, 'manual', now, 'active', 'Pending tab creation in live sheet'].join(' | '));
    SpreadsheetApp.getUi().alert('WARNING: Datasets tab does not exist in live sheet. Registry row for rfp-repository-v1 must be pasted into the Datasets tab manually.');
    return;
  }

  if (sheet.getLastRow() < 1) {
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  sheet.appendRow([datasetId, 'web-scan', sourceUrl, targetTab, ownerAgents, 'manual', now, 'active', 'Scraped from ' + sourceUrl]);
}

function menuIngestProcurements() {
  if (!_isAdminUser_()) { _toast_('Admin only'); return; }
  _ingestProcurementsInternal_();
}

function _ingestProcurementsInternal_() {
  const sheet = _getProcurementsSheet_();
  if (!sheet) {
    _toast_('Could not create/open Procurements tab.');
    return;
  }

  var inserted = 0;

  PROCUREMENT_SEED_URLS.forEach(function (url) {
    try {
      var resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
      if (resp.getResponseCode() !== 200) {
        Logger.log('Procurements: ' + url + ' returned HTTP ' + resp.getResponseCode() + ' — skipping (may need auth or is blocked).');
        return;
      }
      var html = resp.getContentText();
      var text = html.replace(/<script[\s\S]*?<\/script>/gi, '')
                     .replace(/<style[\s\S]*?<\/style>/gi, '')
                     .replace(/<[^>]+>/g, ' ')
                     .replace(/\s+/g, ' ')
                     .trim();

      var rfpPattern = /(RFP|Solicitation)\s*[-:]?\s*[A-Za-z0-9]+/gi;
      var rfpNumbers = text.match(/RFP-\d+/gi) || [];
      var uniqueRfps = [];
      var seenRfps = {};
      rfpNumbers.forEach(function (r) {
        var upper = r.toUpperCase();
        if (!seenRfps[upper]) { seenRfps[upper] = true; uniqueRfps.push(upper); }
      });

      uniqueRfps.forEach(function (rfpNum) {
        if (sheet.getLastRow() >= 2) {
          var data = sheet.getDataRange().getValues();
          var exists = false;
          for (var i = 1; i < data.length; i++) {
            if (String(data[i][0] || '').toUpperCase() === rfpNum) { exists = true; break; }
          }
          if (exists) return;
        }
        sheet.appendRow([rfpNum, 'Fetched from ' + url, 'unverified', '', url, Utilities.formatDate(new Date(), 'America/Los_Angeles', 'yyyy-MM-dd')]);
        inserted++;
      });
    } catch (e) {
      Logger.log('Procurements fetch failed for ' + url + ': ' + e.message);
    }
  });

  _writeDatasetsRegistryRow_('rfp-repository-v1', 'Procurements', 'https://www.lausd.org/Page/19831', 'vendor-review-consultant');

  var msg = 'Procurements ingest: ' + inserted + ' new RFP entries.';
  _toast_(msg);
  SpreadsheetApp.getUi().alert('Ingest Procurements', msg, SpreadsheetApp.getUi().ButtonSet.OK);
}
