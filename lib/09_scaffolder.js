/**
 * RiskAI-GAS — Project Scaffolder (09_scaffolder.js)
 *
 * Utilities for creating new projects programmatically:
 *   - URL-safe slug generation
 *   - Drive folder creation under category roots
 *   - Master project registry sheet append
 *   - Top-level coordinator (extend for Notion, extra sheets, emails, etc.)
 *
 * This module is intentionally minimal — it wires the Drive + registry
 * patterns without assuming Notion or any external project management tool.
 * When you need Notion integration, add a _notionUpdatePage_() helper here
 * and call it from scaffoldProject().
 *
 * Public API:
 *   generateSlug(name)                                  — URL-safe kebab-case slug
 *   createProjectFolder(name, category, parentFolderId) — Drive folder, idempotent
 *   appendToRegistry(registrySheetId, projectData)      — append row to central registry
 *   scaffoldProject(opts)                               — full coordinator
 *
 * Merged from: ProjectScaffolder (rayattalla@gmail.com)
 */

// ==========================================================================
// SLUG GENERATION
// ==========================================================================

/**
 * Convert a display name to a URL-safe kebab-case slug.
 * Example: 'Architecture Review Board 2026' → 'architecture-review-board-2026'
 */
function generateSlug(name) {
  return String(name || '')
    .toLowerCase()
    .trim()
    .replace(/[^a-z0-9\s-]/g, '')
    .replace(/\s+/g, '-')
    .replace(/-+/g, '-')
    .replace(/^-|-$/g, '')
    .slice(0, 60);
}

// ==========================================================================
// DRIVE FOLDER CREATION
// ==========================================================================

/**
 * Create (or find) a project folder under an optional category root.
 * All three parameters are optional — omitting all creates under CONFIG root.
 *
 * projectName:    display name for the project folder
 * categoryName:   subfolder inside parent (e.g., 'Security', 'Infrastructure')
 * parentFolderId: Drive folder ID of the root (CONFIG root folder if blank)
 *
 * Returns: DriveApp Folder object for the project folder.
 *
 * Example:
 *   const folder = createProjectFolder('New ARB Project', 'Security');
 */
function createProjectFolder(projectName, categoryName, parentFolderId) {
  let root;
  if (parentFolderId) {
    try { root = DriveApp.getFolderById(parentFolderId); }
    catch (e) { throw new Error('createProjectFolder: invalid parentFolderId — ' + e.message); }
  } else {
    root = _getOrCreateRootFolder_();
  }
  const parent = categoryName ? _getOrCreateSubfolder_(root, categoryName) : root;
  return _getOrCreateSubfolder_(parent, String(projectName || 'Unnamed Project'));
}

// ==========================================================================
// MASTER REGISTRY — central project tracking spreadsheet
// This is a DIFFERENT spreadsheet from the project's own backend sheet.
// It's the org-level "all projects" registry.
// ==========================================================================

/**
 * Append a project row to a central registry spreadsheet.
 * Creates the 'Projects' sheet with headers if it doesn't exist.
 *
 * registrySheetId: Spreadsheet ID of the central registry (required)
 * projectData: {
 *   id:          string — project ID (e.g., 'PROJ-2026-001')
 *   name:        string — display name
 *   slug:        string — URL-safe identifier
 *   category:    string — project category
 *   owner:       string — owner email
 *   status:      string — 'Planning' | 'Active' | 'Completed' | ...
 *   folderId:    string — Drive folder ID
 *   webAppUrl:   string — deployed web app URL (if applicable)
 *   createdDate: string — ISO datetime string
 *   notes:       string
 * }
 */
function appendToRegistry(registrySheetId, projectData) {
  if (!registrySheetId) throw new Error('appendToRegistry: registrySheetId is required');
  const ss    = SpreadsheetApp.openById(registrySheetId);
  let   sheet = ss.getSheetByName('Projects');
  if (!sheet) {
    sheet = ss.insertSheet('Projects');
    sheet.appendRow(['ID', 'Name', 'Slug', 'Category', 'Owner', 'Status',
                     'Folder_ID', 'Web_App_URL', 'Created_Date', 'Notes']);
    sheet.getRange(1, 1, 1, 10).setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.setColumnWidths(1, 10, 160);
  }
  const d = projectData || {};
  sheet.appendRow([
    d.id          || '',
    d.name        || '',
    d.slug        || generateSlug(d.name || ''),
    d.category    || '',
    d.owner       || '',
    d.status      || 'Planning',
    d.folderId    || '',
    d.webAppUrl   || '',
    d.createdDate || _now_(),
    d.notes       || ''
  ]);
  Logger.log('appendToRegistry: added "' + d.name + '" (' + d.id + ') to registry ' + registrySheetId);
}

// ==========================================================================
// TOP-LEVEL SCAFFOLDER
// ==========================================================================

/**
 * Create a new project scaffold: Drive folder hierarchy + registry entry + audit log.
 *
 * Extend this function to add:
 *   - Notion page creation (_notionUpdatePage_ stub below)
 *   - Auto-provisioning the project's backend spreadsheet
 *   - Sending a "project created" email via sendMail()
 *
 * opts: {
 *   name:             string  — project display name (required)
 *   category:         string  — category folder name (e.g., 'Security')
 *   owner:            string  — owner email (defaults to CONFIG.ADMIN_EMAIL)
 *   parentFolderId:   string  — root Drive folder override
 *   registrySheetId:  string  — central registry spreadsheet ID (skip if blank)
 *   idPrefix:         string  — e.g., 'PROJ' (defaults to CONFIG.ID_PREFIX)
 *   notes:            string  — freeform notes for registry row
 * }
 *
 * Returns: { id, name, slug, folderId }
 *
 * Example:
 *   const proj = scaffoldProject({ name: 'Browser Risk Dashboard', category: 'Security' });
 *   Logger.log(proj.folderId);  // Drive folder ID
 */
function scaffoldProject(opts) {
  opts = opts || {};
  if (!opts.name) throw new Error('scaffoldProject: opts.name is required');

  const prefix = opts.idPrefix || CONFIG.ID_PREFIX;
  const id     = _nextSequentialId_(prefix, CONFIG.SHEET_TABS.MASTER);
  const slug   = generateSlug(opts.name);
  const folder = createProjectFolder(opts.name, opts.category || '', opts.parentFolderId || '');

  const projectData = {
    id:          id,
    name:        opts.name,
    slug:        slug,
    category:    opts.category    || '',
    owner:       opts.owner       || CONFIG.ADMIN_EMAIL,
    status:      'Planning',
    folderId:    folder.getId(),
    webAppUrl:   '',
    createdDate: _now_(),
    notes:       opts.notes || ''
  };

  if (opts.registrySheetId) {
    appendToRegistry(opts.registrySheetId, projectData);
  }

  auditLog_(id, 'PROJECT_SCAFFOLDED',
    'name=' + opts.name + ' slug=' + slug + ' category=' + (opts.category || '') + ' folder=' + folder.getId());

  Logger.log('scaffoldProject: created "' + opts.name + '" (ID: ' + id + ', folder: ' + folder.getId() + ')');
  return { id: id, name: opts.name, slug: slug, folderId: folder.getId() };
}

// ==========================================================================
// NOTION STUB — extend when you need Notion integration
// Uncomment and fill in when you have a Notion integration token.
// ==========================================================================

/*
function _notionUpdatePage_(pageId, properties, status) {
  const token = PropertiesService.getScriptProperties().getProperty('NOTION_TOKEN');
  if (!token) throw new Error('NOTION_TOKEN not set in PropertiesService');
  fetchWithRetry_('https://api.notion.com/v1/pages/' + pageId, {
    method:      'patch',
    contentType: 'application/json',
    headers:     { 'Authorization': 'Bearer ' + token, 'Notion-Version': '2022-06-28' },
    payload:     JSON.stringify({ properties: properties })
  });
}
*/
