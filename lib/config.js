/**
 * Template — Shell Config (config.js)
 *
 * Project-specific constants for the bound shell script.
 * These mirror key values from the library's 01_config.js.
 * Tag: keep-in-sync — update both places when changing identity values.
 *
 * WHY these are duplicated here (and not only in the library):
 *   The shell's doGet() and onOpen() run BEFORE the library loads (or may use
 *   a stale library snapshot). Inlining critical constants here ensures menus,
 *   access gates, and the web app entry page always have current values.
 */

// keep-in-sync with lib/01_config.js
var SHELL_VERSION    = '1.2.1';  // + memory skill (total-recall style, no vector DB)
var ADMIN_EMAIL      = 'remon.attalla@lausd.net';      // must match CONFIG.ADMIN_EMAIL in lib
var SHEET_ID         = '16srghMJcWSWM5azEokij4010vH1sOk_Hl97o6PChJ4Q';      // must match CONFIG.BACKEND_SHEET_ID in lib (from imported seed xlsx)
var ORG_NAME         = 'Los Angeles Unified School District';
var ORG_SHORT        = 'LAUSD';
var PROGRAM_NAME     = 'LAUSD Agent Platform';         // Agent chat platform (GAS-native)
var PROGRAM_SHORT    = 'AGENTS';                      // used in menu title
var NOTIFY_EMAIL     = 'remon.attalla@lausd.net';       // must match CONFIG.NOTIFY_EMAIL in lib
var TIMEZONE         = 'America/Los_Angeles';
