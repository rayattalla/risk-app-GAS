/**
 * RiskAI-GAS — AI Integration (08_ai.js)
 *
 * Unified multi-provider AI wrapper. Providers tried in order:
 *   OpenRouter (primary) → Google Gemini → HuggingFace → local fallback
 *
 * API keys: PropertiesService only — never in sheets or code.
 * Non-sensitive settings: Config sheet (AI_PROVIDER, AI_MODEL, AI_MAX_TOKENS).
 *
 * Public API:
 *   callAI(prompt, opts)                    — send a prompt, get text back
 *   localFallbackAnalysis(text, agentType)  — keyword/heuristic analysis (no API)
 *   setAiApiKey(provider, key)              — store API key in PropertiesService
 *   getAiStatus()                           — circuit breaker state for all providers
 *   resetAiCircuit(provider)                — manually re-open a tripped circuit
 *
 * Config sheet keys:
 *   AI_PROVIDER    — 'openrouter' | 'gemini' | 'huggingface' (default: openrouter)
 *   AI_MODEL       — model ID override (provider default used if blank)
 *   AI_MAX_TOKENS  — integer (default 512)
 *
 * PropertiesService keys (managed by setAiApiKey):
 *   AI_KEY_OPENROUTER, AI_KEY_GEMINI, AI_KEY_HUGGINGFACE
 *   AI_CIRCUIT_*   — internal circuit breaker state (do not set manually)
 *
 * Merged from: RayAI.Library + RayWatchLibrary (rayattalla@gmail.com)
 */

// ==========================================================================
// CONSTANTS
// ==========================================================================

var AI_CIRCUIT_FAIL_THRESHOLD_ = 3;
var AI_CIRCUIT_COOLDOWN_MS_    = 5 * 60 * 1000; // 5 minutes

var AI_PROVIDER_ORDER_ = ['openrouter', 'gemini', 'huggingface'];

var AI_DEFAULT_MODELS_ = {
  openrouter:  'mistralai/mistral-7b-instruct:free',
  gemini:      'gemini-1.5-flash',
  huggingface: 'mistralai/Mistral-7B-Instruct-v0.2'
};

// System prompts injected automatically per agentType — override via opts.systemPrompt
var AI_SYSTEM_PROMPTS_ = {
  summary:        'You are a concise summarizer. Return a 2–3 sentence summary of the input.',
  severity:       'You are a risk analyst. Classify severity as Critical, High, Medium, or Low with one sentence of reasoning.',
  tags:           'Extract 3–7 keyword tags from the input. Return only a comma-separated list, nothing else.',
  recommendation: 'You are an IT security advisor at LAUSD. Provide 3 concrete, numbered action items.',
  general:        'You are a helpful assistant. Be concise and direct.'
};

// ==========================================================================
// API KEY MANAGEMENT — PropertiesService only
// ==========================================================================

/**
 * Store an AI provider API key in script PropertiesService.
 * Run ONCE from the Apps Script editor. Keys persist across executions.
 *
 * provider: 'openrouter' | 'gemini' | 'huggingface'
 * key:      API key string
 *
 * Example: setAiApiKey('openrouter', 'sk-or-...')
 */
function setAiApiKey(provider, key) {
  if (!provider || !key) throw new Error('setAiApiKey: provider and key are required');
  PropertiesService.getScriptProperties().setProperty('AI_KEY_' + provider.toUpperCase(), key);
  Logger.log('setAiApiKey: stored key for provider "' + provider + '"');
}

function getAiApiKey_(provider) {
  return PropertiesService.getScriptProperties()
    .getProperty('AI_KEY_' + String(provider).toUpperCase()) || '';
}

// ==========================================================================
// CIRCUIT BREAKER — per-provider failure tracking
// After AI_CIRCUIT_FAIL_THRESHOLD_ consecutive failures, the provider is
// skipped for AI_CIRCUIT_COOLDOWN_MS_ before being retried automatically.
// WHY: Without this, a bad API key or provider outage causes every callAI()
// to waste 3 × retry attempts before falling through — killing performance.
// ==========================================================================

function _circuitIsOpen_(provider) {
  const props     = PropertiesService.getScriptProperties();
  const openUntil = Number(props.getProperty('AI_CIRCUIT_' + provider + '_openUntil') || 0);
  if (!openUntil) return false;
  if (Date.now() < openUntil) return true;
  // Cooldown expired — reset automatically
  props.deleteProperty('AI_CIRCUIT_' + provider + '_openUntil');
  props.deleteProperty('AI_CIRCUIT_' + provider + '_failures');
  return false;
}

function _circuitOnSuccess_(provider) {
  const props = PropertiesService.getScriptProperties();
  props.deleteProperty('AI_CIRCUIT_' + provider + '_failures');
  props.deleteProperty('AI_CIRCUIT_' + provider + '_openUntil');
}

function _circuitOnFailure_(provider) {
  const props    = PropertiesService.getScriptProperties();
  const failKey  = 'AI_CIRCUIT_' + provider + '_failures';
  const failures = Number(props.getProperty(failKey) || 0) + 1;
  props.setProperty(failKey, String(failures));
  if (failures >= AI_CIRCUIT_FAIL_THRESHOLD_) {
    const until = Date.now() + AI_CIRCUIT_COOLDOWN_MS_;
    props.setProperty('AI_CIRCUIT_' + provider + '_openUntil', String(until));
    Logger.log('AI circuit OPEN for "' + provider + '" until ' + new Date(until).toISOString());
  }
}

// ==========================================================================
// PROVIDER IMPLEMENTATIONS
// ==========================================================================

function _callOpenRouter_(prompt, opts) {
  const key = getAiApiKey_('openrouter');
  if (!key) throw new Error('openrouter: no API key — run setAiApiKey("openrouter", key)');

  const model     = opts.model || _getConfigValue_('AI_MODEL') || AI_DEFAULT_MODELS_.openrouter;
  const maxTokens = opts.maxTokens || Number(_getConfigValue_('AI_MAX_TOKENS') || 512);
  const sysPrompt = opts.systemPrompt || AI_SYSTEM_PROMPTS_[opts.agentType] || AI_SYSTEM_PROMPTS_.general;

  const messages = [
    { role: 'system', content: sysPrompt },
    { role: 'user',   content: prompt }
  ];

  const resp = fetchWithRetry_('https://openrouter.ai/api/v1/chat/completions', {
    method:      'post',
    contentType: 'application/json',
    headers:     { Authorization: 'Bearer ' + key, 'HTTP-Referer': 'https://script.google.com' },
    payload:     JSON.stringify({ model: model, messages: messages, max_tokens: maxTokens })
  }, 2);

  const body = JSON.parse(resp.getContentText());
  if (!body.choices || !body.choices[0] || !body.choices[0].message) {
    throw new Error('openrouter: unexpected response — ' + resp.getContentText().slice(0, 200));
  }
  return String(body.choices[0].message.content || '').trim();
}

function _callGemini_(prompt, opts) {
  const key = getAiApiKey_('gemini');
  if (!key) throw new Error('gemini: no API key — run setAiApiKey("gemini", key)');

  const model     = opts.model || AI_DEFAULT_MODELS_.gemini;
  const maxTokens = opts.maxTokens || Number(_getConfigValue_('AI_MAX_TOKENS') || 512);
  const sysPrompt = opts.systemPrompt || AI_SYSTEM_PROMPTS_[opts.agentType] || '';
  const fullPrompt = sysPrompt ? sysPrompt + '\n\n' + prompt : prompt;

  const url  = 'https://generativelanguage.googleapis.com/v1beta/models/' + model + ':generateContent?key=' + key;
  const resp = fetchWithRetry_(url, {
    method:      'post',
    contentType: 'application/json',
    payload:     JSON.stringify({
      contents:         [{ parts: [{ text: fullPrompt }] }],
      generationConfig: { maxOutputTokens: maxTokens }
    })
  }, 2);

  const body = JSON.parse(resp.getContentText());
  const text = body && body.candidates && body.candidates[0] &&
               body.candidates[0].content && body.candidates[0].content.parts &&
               body.candidates[0].content.parts[0] && body.candidates[0].content.parts[0].text;
  if (!text) throw new Error('gemini: unexpected response — ' + resp.getContentText().slice(0, 200));
  return String(text).trim();
}

function _callHuggingFace_(prompt, opts) {
  const key = getAiApiKey_('huggingface');
  if (!key) throw new Error('huggingface: no API key — run setAiApiKey("huggingface", key)');

  const model     = opts.model || AI_DEFAULT_MODELS_.huggingface;
  const maxTokens = opts.maxTokens || Number(_getConfigValue_('AI_MAX_TOKENS') || 256);
  const sysPrompt = opts.systemPrompt || AI_SYSTEM_PROMPTS_[opts.agentType] || '';
  // Mistral instruction format
  const fullPrompt = sysPrompt
    ? '[INST] ' + sysPrompt + '\n\n' + prompt + ' [/INST]'
    : '[INST] ' + prompt + ' [/INST]';

  const resp = fetchWithRetry_('https://api-inference.huggingface.co/models/' + model, {
    method:      'post',
    contentType: 'application/json',
    headers:     { Authorization: 'Bearer ' + key },
    payload:     JSON.stringify({
      inputs:     fullPrompt,
      parameters: { max_new_tokens: maxTokens, return_full_text: false }
    })
  }, 2);

  const body = JSON.parse(resp.getContentText());
  if (body.error) throw new Error('huggingface: ' + body.error);
  const text = Array.isArray(body)
    ? (body[0] && body[0].generated_text)
    : body.generated_text;
  if (!text) throw new Error('huggingface: no generated_text in response');
  return String(text).trim();
}

// ==========================================================================
// callAI — main public entry point
// ==========================================================================

/**
 * Send a prompt to the best available AI provider and return the response.
 * Providers are tried in order (openrouter → gemini → huggingface) with
 * circuit breaker skip, then falls back to localFallbackAnalysis().
 *
 * prompt:  string — the user prompt
 * opts:
 *   provider:     'openrouter' | 'gemini' | 'huggingface' — force a specific provider
 *   agentType:    'summary' | 'severity' | 'tags' | 'recommendation' | 'general'
 *   maxTokens:    number — override max output tokens
 *   systemPrompt: string — override the built-in system prompt for agentType
 *   noFallback:   true — throw instead of falling back to local analysis
 *
 * Returns: string. Never throws unless opts.noFallback is true.
 *
 * Example:
 *   const summary = callAI(incidentDescription, { agentType: 'summary' });
 *   const sev     = callAI(incidentDescription, { agentType: 'severity', provider: 'gemini' });
 */
function callAI(prompt, opts) {
  opts = opts || {};
  const agentType = opts.agentType || 'general';
  const callers   = {
    openrouter:  _callOpenRouter_,
    gemini:      _callGemini_,
    huggingface: _callHuggingFace_
  };

  // Build ordered provider list
  let providers;
  if (opts.provider) {
    providers = [opts.provider];
  } else {
    const pref = (_getConfigValue_('AI_PROVIDER') || '').toLowerCase().trim();
    if (pref && AI_PROVIDER_ORDER_.indexOf(pref) >= 0) {
      providers = [pref].concat(AI_PROVIDER_ORDER_.filter(p => p !== pref));
    } else {
      providers = AI_PROVIDER_ORDER_.slice();
    }
  }

  const errors = [];
  for (let i = 0; i < providers.length; i++) {
    const provider = providers[i];
    if (_circuitIsOpen_(provider)) {
      Logger.log('callAI: skipping "' + provider + '" (circuit open)');
      continue;
    }
    if (!getAiApiKey_(provider)) {
      Logger.log('callAI: skipping "' + provider + '" (no API key)');
      continue;
    }
    try {
      const result = callers[provider](prompt, opts);
      _circuitOnSuccess_(provider);
      Logger.log('callAI: success via ' + provider + ' [' + agentType + ']');
      return result;
    } catch (e) {
      _circuitOnFailure_(provider);
      errors.push(provider + ': ' + e.message);
      Logger.log('callAI: "' + provider + '" failed — ' + e.message);
    }
  }

  if (opts.noFallback) {
    throw new Error('callAI: all providers failed. ' + errors.join(' | '));
  }

  Logger.log('callAI: all providers failed, using local fallback [' + agentType + ']');
  return localFallbackAnalysis(prompt, agentType);
}

// ==========================================================================
// LOCAL FALLBACK — heuristic analysis, no API, always available
// WHY: A GAS trigger or form submission cannot wait for a human to restore
// the API key. Local fallback keeps the system producing output at lower quality
// rather than throwing errors at submitters.
// ==========================================================================

var AI_SEVERITY_KEYWORDS_ = {
  critical: ['breach', 'ransomware', 'exfiltration', 'compromised', 'attack', 'exploit',
             'zero-day', 'critical', 'emergency', 'data loss', 'unauthorized access', 'lateral movement'],
  high:     ['malware', 'phishing', 'vulnerability', 'high', 'outage', 'failure',
             'intrusion', 'suspicious', 'unauthorized', 'escalation', 'credential'],
  medium:   ['warning', 'medium', 'degraded', 'anomaly', 'unusual', 'misconfiguration',
             'outdated', 'patch', 'unpatched', 'weak'],
  low:      ['info', 'low', 'minor', 'routine', 'informational', 'scheduled', 'review']
};

var AI_STOPWORDS_ = [
  'the','a','an','and','or','but','in','on','at','to','for','of','with','by','from',
  'is','are','was','were','be','been','being','have','has','had','do','does','did',
  'will','would','could','should','may','might','that','this','these','those','it',
  'its','we','our','you','your','they','their','he','she','his','her','not','no',
  'so','as','if','then','than','when','where','which','who'
];

/**
 * Heuristic local analysis — no API calls, always available.
 * agentType: 'summary' | 'severity' | 'tags' | 'recommendation' | 'general'
 */
function localFallbackAnalysis(text, agentType) {
  const t = String(text || '').trim();
  agentType = agentType || 'general';
  if (agentType === 'severity')       return _localSeverity_(t);
  if (agentType === 'tags')           return _localTags_(t);
  if (agentType === 'summary')        return _localSummary_(t);
  if (agentType === 'recommendation') return _localRecommendation_(t);
  return _localSummary_(t) + '\n\nKeywords: ' + _localTags_(t);
}

function _localSeverity_(text) {
  const lower = text.toLowerCase();
  const tiers = ['critical', 'high', 'medium', 'low'];
  for (let i = 0; i < tiers.length; i++) {
    const kws = AI_SEVERITY_KEYWORDS_[tiers[i]];
    for (let j = 0; j < kws.length; j++) {
      if (lower.indexOf(kws[j]) >= 0) {
        const label = tiers[i].charAt(0).toUpperCase() + tiers[i].slice(1);
        return label + ' — matched keyword "' + kws[j] + '" (local fallback)';
      }
    }
  }
  return 'Medium — no severity keywords matched (local fallback)';
}

function _localTags_(text) {
  const stopSet = {};
  AI_STOPWORDS_.forEach(w => { stopSet[w] = true; });
  const words = text.toLowerCase().replace(/[^a-z0-9\s]/g, ' ').split(/\s+/);
  const freq  = {};
  words.forEach(w => {
    if (w.length > 3 && !stopSet[w]) freq[w] = (freq[w] || 0) + 1;
  });
  const tags = Object.keys(freq).sort((a, b) => freq[b] - freq[a]).slice(0, 7);
  return tags.length ? tags.join(', ') : 'general';
}

function _localSummary_(text) {
  const sentences = text.match(/[^.!?]+[.!?]+/g) || [];
  if (sentences.length <= 2) {
    return (text.slice(0, 300) + (text.length > 300 ? '...' : '')) + ' [local fallback]';
  }
  return sentences.slice(0, 2).join(' ').trim() + ' [local fallback]';
}

function _localRecommendation_(text) {
  const sevLine = _localSeverity_(text);
  const sev     = sevLine.split(' — ')[0].toLowerCase();
  const recs = {
    critical: '1. Activate incident response plan immediately.\n2. Isolate affected systems and revoke compromised credentials.\n3. Notify CISO and initiate stakeholder communications.',
    high:     '1. Escalate to security team lead within 1 hour.\n2. Collect and preserve logs and forensic evidence.\n3. Apply available patches or mitigations.',
    medium:   '1. Assign to security analyst for investigation within 24 hours.\n2. Review relevant system logs and access records.\n3. Document findings and schedule remediation.',
    low:      '1. Log and track in the next review cycle.\n2. Verify with asset owner whether action is required.\n3. Update risk register if applicable.'
  };
  return (recs[sev] || recs.medium) + '\n\n[local fallback — no AI provider available]';
}

// ==========================================================================
// STATUS AND DIAGNOSTICS
// ==========================================================================

/**
 * Return the current state of all providers (API key present, circuit status).
 * Safe to expose in admin panel / diagnostics.
 */
function getAiStatus() {
  const props  = PropertiesService.getScriptProperties();
  const status = {};
  AI_PROVIDER_ORDER_.forEach(function(provider) {
    const openUntil = Number(props.getProperty('AI_CIRCUIT_' + provider + '_openUntil') || 0);
    const failures  = Number(props.getProperty('AI_CIRCUIT_' + provider + '_failures')  || 0);
    status[provider] = {
      hasKey:      !!getAiApiKey_(provider),
      failures:    failures,
      circuitOpen: openUntil > 0 && Date.now() < openUntil,
      openUntil:   openUntil ? new Date(openUntil).toISOString() : null
    };
  });
  return status;
}

/**
 * Manually reset a provider's circuit breaker (e.g., after fixing an API key).
 * provider: 'openrouter' | 'gemini' | 'huggingface' | 'all'
 */
function resetAiCircuit(provider) {
  const props   = PropertiesService.getScriptProperties();
  const targets = provider === 'all' ? AI_PROVIDER_ORDER_ : [provider];
  targets.forEach(function(p) {
    props.deleteProperty('AI_CIRCUIT_' + p + '_failures');
    props.deleteProperty('AI_CIRCUIT_' + p + '_openUntil');
    Logger.log('resetAiCircuit: reset "' + p + '"');
  });
  Logger.log('resetAiCircuit: done — ' + targets.join(', '));
}
