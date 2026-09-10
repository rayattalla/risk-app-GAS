/**
 * RiskAI-GAS — Config (01_config.js)
 *
 * Fill in all YOUR_* placeholders before deploying this library.
 * Everything else is ready to use as-is.
 *
 * SHEET_TABS defines the standard tab set. Add your own tabs here;
 * the setup() function in 02_helpers.js will create them automatically.
 */

const CONFIG = {

  VERSION: '1.1.0',  // shipping prep + unified KB + cleanup

  // ==========================================================================
  // BACKEND — fill in before deploying
  // ==========================================================================

  /** The spreadsheet this project reads/writes. Required. Import seed xlsx first. */
  BACKEND_SHEET_ID: '16srghMJcWSWM5azEokij4010vH1sOk_Hl97o6PChJ4Q',

  /** Root Drive folder for generated files. Leave blank to auto-create on first setup(). */
  ROOT_FOLDER_ID: '',

  // ==========================================================================
  // IDENTITY — fill in before deploying
  // ==========================================================================

  ORG_NAME:      'Los Angeles Unified School District',
  ORG_SHORT:     'LAUSD',
  PROGRAM_NAME:  'LAUSD Agent Platform',   // Agent chat platform (GAS-native)
  PROGRAM_SHORT: 'AGENTS',                // used in menus + email subjects
  TIMEZONE:      'America/Los_Angeles',

  // ==========================================================================
  // EMAIL — fill in before deploying
  // ==========================================================================

  /** Owner/admin email — always has Admin role; cannot be locked out. */
  ADMIN_EMAIL: 'remon.attalla@lausd.net',

  /** Team inbox — CC'd on every workflow notification. */
  NOTIFY_EMAIL: 'remon.attalla@lausd.net',

  /** Display name in From: field. */
  EMAIL_SENDER_NAME: 'LAUSD ITS',

  // ==========================================================================
  // SHEET TABS — standard set; extend with your project-specific tabs
  // ==========================================================================

  SHEET_TABS: {
    AGENTS:                  'Agents',                 // slug|name|status|org|model|personality|kb_tags|notes
    KB:                      'KB',                     // id|slug|title|body|tags ; slug=shared or agent slug
    CHAT_LOG:                'ChatLog',                // Timestamp|Email|Slug|Model|Prompt|Response|Status|Error|DurationMs
    ACCESS_CONTROL:          'Access_Control',          // email → role rows (reuse)
    CONFIG:                  'Config',                  // runtime key/value store
    AUDIT_LOG:               'Audit_Log'               // timestamped action log
  },

  // ==========================================================================
  // RECORD ID FORMAT — e.g., 'REC' → REC-2026-001
  // ==========================================================================

  ID_PREFIX: 'REC',   // change to match your domain (ARB, INC, PROJ, etc.)

  // ==========================================================================
  // RISK TIERS — Mandiant TPRM-aligned (remove if not using risk scoring)
  // ==========================================================================

  RISK_TIERS: {
    CRITICAL: { min: 18, max: 999, label: 'Critical', color: '#B71C1C' },
    HIGH:     { min: 13, max: 17,  label: 'High',     color: '#E65100' },
    MEDIUM:   { min: 7,  max: 12,  label: 'Medium',   color: '#F9A825' },
    LOW:      { min: 0,  max: 6,   label: 'Low',      color: '#2E7D32' }
  },

  getRiskTier(score) {
    const n = Number(score) || 0;
    if (n >= this.RISK_TIERS.CRITICAL.min) return this.RISK_TIERS.CRITICAL.label;
    if (n >= this.RISK_TIERS.HIGH.min)     return this.RISK_TIERS.HIGH.label;
    if (n >= this.RISK_TIERS.MEDIUM.min)   return this.RISK_TIERS.MEDIUM.label;
    return this.RISK_TIERS.LOW.label;
  },

  getRiskColor(tier) {
    const k = String(tier || '').toUpperCase();
    return (this.RISK_TIERS[k] && this.RISK_TIERS[k].color) || '#5F6368';
  }
};

// ==========================================================================
// TEST MODE — redirect all outgoing emails to a safe address during dev/testing.
// Run enableTestMode() from the Apps Script editor, not from triggers.
// ==========================================================================

function enableTestMode() {
  PropertiesService.getScriptProperties().setProperties({
    'TEST_MODE':  'true',
    'TEST_EMAIL': CONFIG.ADMIN_EMAIL
  });
  Logger.log('TEST MODE ON — emails → ' + CONFIG.ADMIN_EMAIL);
}

function disableTestMode() {
  const p = PropertiesService.getScriptProperties();
  p.deleteProperty('TEST_MODE');
  p.deleteProperty('TEST_EMAIL');
  Logger.log('TEST MODE OFF — real emails will be sent');
}

function isTestMode() {
  return PropertiesService.getScriptProperties().getProperty('TEST_MODE') === 'true';
}

function getTestEmail() {
  return PropertiesService.getScriptProperties().getProperty('TEST_EMAIL') || CONFIG.ADMIN_EMAIL;
}

function checkModes() {
  Logger.log('Test mode: ' + (isTestMode() ? 'ON → ' + getTestEmail() : 'OFF'));
}

// ==========================================================================
// CONFIG VALIDATION — catch unfilled placeholders before they cause silent bugs
// Call from setup() or run manually from the editor after filling in your values.
// ==========================================================================

/**
 * Throw an error if any required CONFIG field still contains a placeholder.
 * Run this once after filling in your values to verify the library is ready.
 *
 * Example: call validateConfig_() as the first line of setup().
 */
function validateConfig_() {
  const required = {
    BACKEND_SHEET_ID: CONFIG.BACKEND_SHEET_ID,
    PROGRAM_NAME:     CONFIG.PROGRAM_NAME,
    ADMIN_EMAIL:      CONFIG.ADMIN_EMAIL,
    NOTIFY_EMAIL:     CONFIG.NOTIFY_EMAIL
  };
  const errors = [];
  Object.keys(required).forEach(key => {
    const v = String(required[key] || '');
    if (!v || v.startsWith('YOUR_') || v.startsWith('your.') || v.startsWith('your-')) {
      errors.push('CONFIG.' + key + ' is still a placeholder: "' + v + '"');
    }
  });
  if (errors.length) {
    throw new Error('Config not filled in:\n' + errors.join('\n')
      + '\n\nFill in lib/01_config.js before deploying.');
  }
  Logger.log('validateConfig_: all required fields set ✓');
}

// ==========================================================================
// getConfigData — safe identity + runtime snapshot for the frontend
// Called by the shell's doGet injection and by google.script.run on page load.
// Returns only what the browser needs — never sensitive keys or IDs.
// ==========================================================================

// getConfigData removed (duplicate) — shell.js provides the inlined version used by the web app.
