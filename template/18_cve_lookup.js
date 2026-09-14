/**
 * 18_cve_lookup.js
 *
 * Direct CVE lookups against real vulnerability sources -- no API key for
 * either one. Auto-triggers whenever a CVE ID appears anywhere in the
 * user's message (unlike the generic 'search' skill, no "search:" prefix
 * needed -- a user typing a bare CVE number is exactly the case this
 * exists for).
 *
 * Sources:
 *  - NVD (nvd.nist.gov) REST API -- description, CVSS score/severity.
 *  - CISA Known Exploited Vulnerabilities catalog, straight from the public
 *    JSON feed -- whether it's actively/known exploited, CISA's required
 *    action and remediation due date.
 *
 * The model's training data is stale and CVE numbers/details are exactly
 * the kind of specific-sounding thing it will confidently invent -- the
 * system prompt instruction (see 20_chat.js) tells it to use ONLY what's
 * returned here for any CVE mentioned, never pre-training memory.
 */

var CVE_ID_PATTERN = /CVE-\d{4}-\d{4,7}/gi;
var CISA_KEV_URL = 'https://www.cisa.gov/sites/default/files/feeds/known_exploited_vulnerabilities.json';
var NVD_API_URL = 'https://services.nvd.nist.gov/rest/json/cves/2.0?cveId=';
var MAX_CVES_PER_MESSAGE = 3; // cap external calls per message

function _extractCveIds_(text) {
  var matches = String(text || '').match(CVE_ID_PATTERN) || [];
  var seen = {};
  var out = [];
  matches.forEach(function (m) {
    var id = m.toUpperCase();
    if (!seen[id]) { seen[id] = true; out.push(id); }
  });
  return out.slice(0, MAX_CVES_PER_MESSAGE);
}

function _lookupNvd_(cveId) {
  try {
    var resp = UrlFetchApp.fetch(NVD_API_URL + encodeURIComponent(cveId), { muteHttpExceptions: true });
    if (resp.getResponseCode() !== 200) return null;
    var data = JSON.parse(resp.getContentText());
    var vuln = data.vulnerabilities && data.vulnerabilities[0] && data.vulnerabilities[0].cve;
    if (!vuln) return null;

    var desc = '';
    (vuln.descriptions || []).some(function (d) {
      if (d.lang === 'en') { desc = d.value; return true; }
      return false;
    });

    var cvss = '';
    var metrics = vuln.metrics || {};
    var chosen = (metrics.cvssMetricV40 && metrics.cvssMetricV40[0]) ||
                 (metrics.cvssMetricV31 && metrics.cvssMetricV31[0]) ||
                 (metrics.cvssMetricV2 && metrics.cvssMetricV2[0]);
    if (chosen && chosen.cvssData) {
      cvss = chosen.cvssData.baseScore + ' (' + (chosen.cvssData.baseSeverity || chosen.baseSeverity || '') + ')';
    }

    return { id: cveId, published: vuln.published || '', description: desc, cvss: cvss };
  } catch (e) {
    return null;
  }
}

function _checkCisaKev_(cveId) {
  try {
    var resp = UrlFetchApp.fetch(CISA_KEV_URL, { muteHttpExceptions: true });
    if (resp.getResponseCode() !== 200) return null;
    var data = JSON.parse(resp.getContentText());
    var entry = (data.vulnerabilities || []).filter(function (v) { return v.cveID === cveId; })[0];
    if (!entry) return { listed: false };
    return {
      listed: true,
      dateAdded: entry.dateAdded || '',
      dueDate: entry.dueDate || '',
      requiredAction: entry.requiredAction || '',
      knownRansomwareUse: entry.knownRansomwareCampaignUse || ''
    };
  } catch (e) {
    return null;
  }
}

// Returns a formatted context block for every CVE ID found in the message,
// or '' if none were mentioned. Called from chatWithAgent_ regardless of
// which other skills are enabled, once 'cve-lookup' is in Agents.skills.
function _cveLookupSkill_(text) {
  var ids = _extractCveIds_(text);
  if (!ids.length) return '';

  var blocks = [];
  ids.forEach(function (id) {
    var nvd = _lookupNvd_(id);
    var kev = _checkCisaKev_(id);
    var lines = ['[' + id + ']'];

    if (nvd) {
      lines.push('NVD: ' + (nvd.description || '(no description in NVD record)'));
      if (nvd.cvss) lines.push('CVSS: ' + nvd.cvss);
      if (nvd.published) lines.push('Published: ' + nvd.published);
    } else {
      lines.push('NVD: not found (may not be a registered CVE, or the lookup failed -- do not assume it is invalid).');
    }

    if (kev && kev.listed) {
      lines.push(
        'CISA KEV: LISTED -- actively/known exploited. Added ' + kev.dateAdded +
        ', remediation due ' + kev.dueDate + '. Required action: ' + kev.requiredAction
      );
      if (kev.knownRansomwareUse && kev.knownRansomwareUse !== 'Unknown') {
        lines.push('Known ransomware campaign use: ' + kev.knownRansomwareUse);
      }
    } else if (kev) {
      lines.push('CISA KEV: not listed (not currently flagged by CISA as actively/known exploited).');
    } else {
      lines.push('CISA KEV: lookup failed -- treat exploitation status as unknown, not "not exploited".');
    }

    blocks.push(lines.join('\n'));
  });

  return blocks.join('\n\n');
}
