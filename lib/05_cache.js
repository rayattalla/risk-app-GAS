/**
 * RiskAI-GAS — Cache (05_cache.js)
 *
 * CacheService wrapper with chunking for large JSON payloads (>90 KB)
 * and LockService to prevent thundering herd on cache miss.
 *
 * WHY chunking: CacheService has a 100 KB per-key limit. Large dashboards or
 * bulk reads easily exceed this. We split the serialised JSON string into 90 KB
 * chunks, store a chunk-count key, then reassemble on read.
 *
 * WHY LockService: Without a lock, a cache miss under concurrent load causes
 * every simultaneous request to hit the spreadsheet at once. The lock ensures
 * only one request builds the cache while others wait and then get the warm cache.
 *
 * Public API:
 *   cacheGetOrFetch(prefix, ttl, fetchFn)  — main entry point
 *   cacheRead(prefix)                      — read only (returns null on miss)
 *   cacheWrite(prefix, str, ttl)           — write (handles chunking)
 *   cacheClear(prefix)                     — delete all chunks for a prefix
 *   getLausdLogo(logoFileId)               — cached base64 Drive logo (1 hr)
 *   installWarmTrigger(handlerFn, minutes) — install a time-based warm trigger
 */

const CACHE_CHUNK_BYTES = 90000;   // 90 KB — safely under the 100 KB key limit
const CACHE_BATCH_KEYS  = 10;      // putAll sends up to 10 chunks at once (≈900 KB batch)

// ==========================================================================
// MAIN ENTRY POINT
// ==========================================================================

/**
 * Return cached string for `prefix`, or call `fetchFn()` to build it.
 * `fetchFn` must return a string (usually JSON.stringify'd + < escaped).
 * Uses LockService to serialize concurrent misses (25s timeout).
 *
 * Example:
 *   const data = cacheGetOrFetch('myapp_v1_', 300, () => JSON.stringify(buildData()));
 */
function cacheGetOrFetch(prefix, ttl, fetchFn) {
  const cache = CacheService.getScriptCache();

  // Fast path — no lock needed for reads
  const hit = cacheRead(prefix, cache);
  if (hit !== null) { Logger.log('Cache HIT: ' + prefix); return hit; }

  // Slow path — serialize misses so only ONE hits the source
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(25000);
    // Re-check after acquiring lock (another request may have just written it)
    const hit2 = cacheRead(prefix, cache);
    if (hit2 !== null) { Logger.log('Cache HIT (post-lock): ' + prefix); return hit2; }
    const str = fetchFn();
    cacheWrite(prefix, str, ttl, cache);
    return str;
  } catch (e) {
    Logger.log('cacheGetOrFetch lock/build error (' + prefix + '): ' + e.message);
    // Fallback: build without caching so the request still succeeds
    try { return fetchFn(); } catch (e2) { return null; }
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}

// ==========================================================================
// READ / WRITE / CLEAR
// ==========================================================================

/**
 * Read chunks from cache and reassemble. Returns null on miss.
 * Accepts an optional pre-obtained cache instance to avoid redundant calls.
 */
function cacheRead(prefix, cache) {
  cache = cache || CacheService.getScriptCache();
  try {
    const nStr = cache.get(prefix + 'n');
    if (!nStr) return null;
    const n    = parseInt(nStr, 10);
    const keys = Array.from({ length: n }, (_, i) => prefix + i);
    const map  = cache.getAll(keys);
    const chunks = keys.map(k => map[k]);
    if (chunks.every(Boolean)) return chunks.join('');
  } catch (e) {
    Logger.log('cacheRead error (' + prefix + '): ' + e.message);
  }
  return null;
}

/**
 * Split `str` into 90 KB chunks and write to cache with `ttl` seconds.
 * Accepts an optional pre-obtained cache instance.
 */
function cacheWrite(prefix, str, ttl, cache) {
  cache = cache || CacheService.getScriptCache();
  ttl   = ttl   || 300;
  if (!str) return;
  try {
    const chunks = [];
    for (let i = 0; i < str.length; i += CACHE_CHUNK_BYTES) {
      chunks.push(str.slice(i, i + CACHE_CHUNK_BYTES));
    }
    // Write in batches (putAll has a ≈1 MB total limit)
    for (let b = 0; b < chunks.length; b += CACHE_BATCH_KEYS) {
      const entries = {};
      chunks.slice(b, b + CACHE_BATCH_KEYS).forEach((c, j) => {
        entries[prefix + (b + j)] = c;
      });
      cache.putAll(entries, ttl);
    }
    cache.put(prefix + 'n', String(chunks.length), ttl);
    Logger.log('cacheWrite: ' + chunks.length + ' chunks (' + Math.round(str.length / 1024) + ' KB) → ' + prefix);
  } catch (e) {
    Logger.log('cacheWrite failed (' + prefix + '): ' + e.message);
  }
}

/**
 * Delete all chunks for a prefix (forces a fresh fetch on next read).
 */
function cacheClear(prefix) {
  const cache = CacheService.getScriptCache();
  const nStr  = cache.get(prefix + 'n');
  const keys  = [prefix + 'n'];
  if (nStr) {
    const n = parseInt(nStr, 10);
    for (let i = 0; i < n; i++) keys.push(prefix + i);
  }
  cache.removeAll(keys);
  Logger.log('cacheClear: removed ' + keys.length + ' keys for ' + prefix);
}

// ==========================================================================
// LOGO CACHING — Drive image → base64 data URI, cached 1 hour
// ==========================================================================

/**
 * Return the LAUSD logo as a base64 data URI string, cached for 1 hour.
 * `logoFileId` is the Google Drive file ID of the logo image.
 *
 * Returns '' on failure (img src="" hides the element via onerror).
 *
 * Example:
 *   const logo = getLausdLogo('1VCYv42yGT2YlutlHzLLPSkawsTrutyPr');
 *   html = html.replace('__LAUSD_LOGO__', logo);
 */
function getLausdLogo(logoFileId) {
  if (!logoFileId) return '';
  const cache = CacheService.getScriptCache();
  const key   = 'lausd_logo_b64_' + logoFileId;
  const hit   = cache.get(key);
  if (hit) return hit;
  try {
    const blob = DriveApp.getFileById(logoFileId).getBlob();
    const src  = 'data:' + blob.getContentType() + ';base64,' + Utilities.base64Encode(blob.getBytes());
    cache.put(key, src, 3600); // 1 hour
    return src;
  } catch (e) {
    Logger.log('getLausdLogo failed (' + logoFileId + '): ' + e.message);
    return '';
  }
}

// ==========================================================================
// WARM TRIGGER — pre-warm cache on a schedule
// ==========================================================================

/**
 * Install a time-based trigger to call `handlerFn` every `minutes` minutes.
 * Run this once from the Apps Script editor (or from setup()).
 * Idempotent — deletes any existing trigger for the same handler first.
 *
 * Example:
 *   installWarmTrigger('warmCache', 5);
 *
 * Then define a top-level warmCache() function in your shell that calls
 * cacheGetOrFetch(...) to pre-warm the cache before users request the page.
 */
function installWarmTrigger(handlerFn, minutes) {
  minutes = minutes || 5;
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === handlerFn)
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger(handlerFn).timeBased().everyMinutes(minutes).create();
  Logger.log('installWarmTrigger: ' + handlerFn + ' every ' + minutes + ' min');
}
