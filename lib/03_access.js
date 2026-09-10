/**
 * RiskAI-GAS — Access Control (03_access.js)
 *
 * Two-layer authorization:
 *   Layer 1 — Google Workspace domain gate (set on the bound shell manifest).
 *   Layer 2 — Sheet-driven role list (Access_Control tab).
 *
 * Roles (low → high):
 *   Viewer   — read-only
 *   Editor   — can create and edit their own records
 *   Reviewer — Editor + can read all records and add comments
 *   Approver — Reviewer + can approve/reject
 *   Admin    — full access: setup, manage access, override anything
 *
 * Capabilities are mapped to minimum required roles in ROLE_REQUIRED.
 * Add your own capabilities there; the gating functions do the rest.
 *
 * IMPORTANT: Frontend hiding (CSS show/hide) is polish, not security.
 * Every server-side write entry point MUST call _hasRole_(capability).
 */

const ACCESS_ROLES = ['Viewer', 'Editor', 'Reviewer', 'Approver', 'Admin'];

const ROLE_LEVEL = {
  Viewer:   1,
  Editor:   2,
  Reviewer: 3,
  Approver: 4,
  Admin:    5
};

/** Capability → minimum required role. Add project-specific capabilities here. */
const ROLE_REQUIRED = {
  // Read
  view:             'Viewer',

  // Write (own records)
  createRecord:     'Editor',
  editRecord:       'Editor',
  uploadFile:       'Editor',
  withdrawRecord:   'Editor',

  // Review (all records)
  addComment:       'Reviewer',
  viewAllRecords:   'Reviewer',

  // Approve
  decideApproval:   'Approver',
  overrideStatus:   'Approver',

  // Admin
  setup:            'Admin',
  manageAccess:     'Admin',
  regenerate:       'Admin',
  deleteRecord:     'Admin'
};

// ==========================================================================
// ROLE RESOLUTION
// ==========================================================================

/**
 * Resolve the role of the active user.
 * Order of precedence:
 *   1. CONFIG.ADMIN_EMAIL → always Admin (lockout protection)
 *   2. Explicit row in Access_Control tab
 *   3. Default for @lausd.net domain users → Editor
 *   4. Non-domain users → null (deny page)
 */
function _resolveUserRole_() {
  const email = _getUserEmail_();
  if (!email) return null;
  if (email === String(CONFIG.ADMIN_EMAIL || '').toLowerCase().trim()) return 'Admin';

  const sheet = _getSheet_(CONFIG.SHEET_TABS.ACCESS_CONTROL);
  if (sheet && sheet.getLastRow() >= 2) {
    const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
    for (let i = 0; i < rows.length; i++) {
      const e = String(rows[i][0] || '').trim().toLowerCase();
      const r = String(rows[i][1] || '').trim();
      if (e === email && ROLE_LEVEL[r]) return r;
    }
  }

  if (email.endsWith('@lausd.net')) return 'Editor';
  return null;
}

/** True iff the active user has at least the role required by `capability`. */
function _hasRole_(capability) {
  const required = ROLE_REQUIRED[capability] || capability;
  const minLevel = ROLE_LEVEL[required];
  if (!minLevel) return false;
  const role = _resolveUserRole_();
  return !!(role && ROLE_LEVEL[role] >= minLevel);
}

/** Build a standardized denial envelope and audit-log the attempt. */
function _denied_(capability) {
  const required = ROLE_REQUIRED[capability] || capability;
  const role     = _resolveUserRole_() || 'no role';
  const email    = _getUserEmail_();
  auditLog_(null, 'ACCESS_DENIED', capability + ' attempted by ' + email + ' (role: ' + role + ')');
  return {
    success: false,
    error: 'Access denied — ' + required + ' role required (your role: ' + role + ').'
  };
}

// ==========================================================================
// PUBLIC API — called by the shell (via google.script.run)
// ==========================================================================

/**
 * Return the active user's permission envelope.
 * Called by the frontend on load to drive UI hiding (NOT security).
 */
function getMyAccess() {
  const email = _getUserEmail_();
  const role  = _resolveUserRole_();
  const level = role ? ROLE_LEVEL[role] : 0;
  return {
    email:   email,
    role:    role || 'No Access',
    level:   level,
    canView:     level >= ROLE_LEVEL.Viewer,
    canEdit:     level >= ROLE_LEVEL.Editor,
    canReview:   level >= ROLE_LEVEL.Reviewer,
    canApprove:  level >= ROLE_LEVEL.Approver,
    canAdmin:    level >= ROLE_LEVEL.Admin,
    capabilities: Object.keys(ROLE_REQUIRED).reduce((acc, cap) => {
      const req = ROLE_REQUIRED[cap];
      acc[cap] = level >= (ROLE_LEVEL[req] || 999);
      return acc;
    }, {}),
    version: CONFIG.VERSION
  };
}

// isCurrentUserAdmin removed (dup) — shell.js provides email-based impl for pilot.

// ==========================================================================
// ACCESS_CONTROL SHEET MANAGEMENT (Admin only)
// ==========================================================================

/** Grant or update a user's role. Admin only. */
function grantAccess(email, role, notes) {
  if (!_hasRole_('manageAccess')) return _denied_('manageAccess');
  email = String(email || '').trim().toLowerCase();
  role  = String(role  || '').trim();
  if (!email || !ROLE_LEVEL[role]) {
    return { success: false, error: 'Invalid email or role. Valid roles: ' + ACCESS_ROLES.join(', ') };
  }
  const sheet = _getOrCreateSheet_(CONFIG.SHEET_TABS.ACCESS_CONTROL);
  if (sheet.getLastRow() >= 2) {
    const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
    for (let i = 0; i < rows.length; i++) {
      if (String(rows[i][0] || '').trim().toLowerCase() === email) {
        sheet.getRange(i + 2, 2).setValue(role);
        if (notes != null) sheet.getRange(i + 2, 4).setValue(String(notes));
        auditLog_(null, 'ACCESS_UPDATED', email + ' → ' + role);
        return { success: true, message: 'Updated ' + email + ' → ' + role };
      }
    }
  }
  sheet.appendRow([email, role, _now_(), String(notes || '')]);
  auditLog_(null, 'ACCESS_GRANTED', email + ' → ' + role);
  return { success: true, message: 'Granted ' + email + ' the ' + role + ' role' };
}

/** Revoke a user's role entirely. Admin only. */
function revokeAccess(email) {
  if (!_hasRole_('manageAccess')) return _denied_('manageAccess');
  email = String(email || '').trim().toLowerCase();
  if (!email) return { success: false, error: 'Email required' };
  const sheet = _getSheet_(CONFIG.SHEET_TABS.ACCESS_CONTROL);
  if (!sheet || sheet.getLastRow() < 2) return { success: false, error: 'No access records' };
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
  for (let i = 0; i < rows.length; i++) {
    if (String(rows[i][0] || '').trim().toLowerCase() === email) {
      sheet.deleteRow(i + 2);
      auditLog_(null, 'ACCESS_REVOKED', email);
      return { success: true, message: 'Revoked access for ' + email };
    }
  }
  return { success: false, error: 'No record for ' + email };
}

/** List all users with roles. Admin only. */
function listAccess() {
  if (!_hasRole_('manageAccess')) return _denied_('manageAccess');
  const sheet = _getSheet_(CONFIG.SHEET_TABS.ACCESS_CONTROL);
  if (!sheet || sheet.getLastRow() < 2) return { success: true, users: [] };
  const hdr  = _headerMap_(sheet);
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  const users = data.map(row => ({
    email:     row[hdr['Email']      || 0],
    role:      row[hdr['Role']       || 1],
    addedDate: _fmtDate_(row[hdr['Added_Date'] || 2]),
    notes:     row[hdr['Notes']      || 3] || ''
  }));
  return { success: true, users };
}
