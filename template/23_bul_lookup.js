/**
 * 23_bul_lookup.js
 *
 * Keyless lookup skill against the live Bulletins tab — mirrors the
 * cve-lookup pattern (18_cve_lookup.js). Queries by bulletin number
 * (BUL-####.#) or keyword, returns live rows from the Bulletins tab,
 * never recites from memory.
 *
 * Auto-triggers when a BUL-/REF- pattern appears in the user's message
 * (no "search:" prefix needed — a user typing a bare bulletin number
 * is exactly the case this exists for).
 *
 * This replaces the old static-KB bul-lookup which recited from
 * hardcoded persona text and KB rows (including the incorrect
 * BUL-999.16 reference). Those hardcoded references have been removed.
 */

var BULLETIN_ID_PATTERN = /\b(BUL-\d{4}\.\d|REF-\d{4}\.\d)\b/gi;

function _extractBulletinIds_(text) {
  var matches = String(text || '').match(BULLETIN_ID_PATTERN) || [];
  var seen = {};
  var out = [];
  matches.forEach(function (m) {
    var id = m.toUpperCase().replace(/\s+/g, '');
    if (!seen[id]) { seen[id] = true; out.push(id); }
  });
  return out.slice(0, 5);
}

function _bulLookupSkill_(text) {
  if (!text) return '';

  var ids = _extractBulletinIds_(text);
  var keywordQuery = '';

  if (ids.length) {
    keywordQuery = ids[0];
  } else {
    var maybeBulletin = String(text || '').match(/(BUL|REF)-\d{4}/gi);
    if (maybeBulletin) {
      keywordQuery = maybeBulletin[0].toUpperCase();
    }
  }

  if (!keywordQuery) return '';

  var lookupResult = _bulLookupKeyless_(keywordQuery);
  if (!lookupResult) return '';

  var lines = ['[Bulletin Lookup — live from Bulletins tab]'];
  lines.push(lookupResult);
  lines.push('');
  lines.push('Use ONLY the data above for these bulletins — never recite bulletin text from training-data memory.');

  return lines.join('\n');
}
