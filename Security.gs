// ─────────────────────────────────────────────────────────────────────────────
// SECURITY.GS — Authentication, Sessions, RBAC
// MFN Tikinti Materialları CRM — Phase 2
// Version: 2.0.0 | 2026-05-31
//
// Architecture:
//   • Google OAuth identity  — no passwords, no USERS tab
//   • CacheService sessions  — O(1), TTL enforced by Google, 8h
//   • CacheService rate limit — no Sheets writes per request
//   • MANAGERS tab           — single source of truth for users & roles
//
// DO NOT CREATE: USERS tab, SESSIONS tab, RATE_LIMIT tab, Password_Hash col
// ─────────────────────────────────────────────────────────────────────────────

'use strict';

// ── ROLE PERMISSION MAP ──────────────────────────────────────────────────────
// Every doPost action → array of roles that may perform it.
// Backend enforces these. Frontend hides UI elements only (defence in depth).
const ROLE_PERMISSIONS = {
  createOrder:            ['Manager', 'Senior Manager', 'Admin', 'Owner'],
  updateOrder:            ['Manager', 'Senior Manager', 'Admin', 'Owner'],
  changeOrderStatus:      ['Senior Manager', 'Admin', 'Owner'],
  confirmPayment:         ['Admin', 'Owner'],
  createPayment:          ['Admin', 'Owner'],
  approveDiscount:        ['Senior Manager', 'Admin', 'Owner'],
  rejectDiscount:         ['Senior Manager', 'Admin', 'Owner'],
  createProductionRow:    ['Admin', 'Owner'],
  updateProductionStatus: ['Production', 'Admin', 'Owner'],
  updateDeliveryStatus:   ['Delivery', 'Admin', 'Owner'],
  getOrders:              ['Manager', 'Senior Manager', 'Production', 'Delivery', 'Admin', 'Owner'],
  getPayments:            ['Admin', 'Owner'],
  getDashboard:           ['Admin', 'Owner'],
  getAuditLog:            ['Admin', 'Owner'],
  getPriceTable:          ['Manager', 'Senior Manager', 'Admin', 'Owner'],
  getOrder:               ['Manager', 'Senior Manager', 'Production', 'Delivery', 'Admin', 'Owner'],
  manualBackup:           ['Owner'],
};

// ── SESSION CONSTANTS ────────────────────────────────────────────────────────
const SESSION_TTL_SECONDS  = 8 * 3600;   // 8 hours
const SESSION_KEY_PREFIX   = 'mfn_session_';

// ── RATE LIMIT CONSTANTS ─────────────────────────────────────────────────────
const RATE_LIMIT_MAX    = 30;   // max requests per window
const RATE_LIMIT_WINDOW = 60;   // window in seconds
const RATE_LIMIT_PREFIX = 'mfn_rate_';

// ─────────────────────────────────────────────────────────────────────────────
// SESSION MANAGEMENT
// No SESSIONS sheet. No linear scans. TTL enforced by Google.
// Token = UUID in browser sessionStorage. Server stores JSON in CacheService.
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Creates a new CacheService session for the authenticated email.
 * Returns { token, role, name, expires_in } or null if user not found/inactive.
 */
function createSession(email) {
  const user = getUserByEmail(email);
  if (!user) return null;
  if (!user.is_active) return null;

  const token       = Utilities.getUuid();
  const sessionData = {
    email:   email,
    role:    user.role,
    user_id: user.manager_id,
    name:    user.name,
    created: Date.now(),
    expires: Date.now() + SESSION_TTL_SECONDS * 1000
  };

  CacheService.getScriptCache().put(
    SESSION_KEY_PREFIX + token,
    JSON.stringify(sessionData),
    SESSION_TTL_SECONDS
  );

  // Log successful login (calls batchAuditLog defined in Code.gs — same project)
  batchAuditLog(token, [{
    entity_type: 'AUTH',
    entity_id:   email,
    action:      'LOGIN_SUCCESS',
    field_name:  null,
    old_value:   null,
    new_value:   user.role,
    user_email:  email,
    user_role:   user.role,
    note:        'Session created'
  }]);

  return { token, role: user.role, name: user.name, expires_in: SESSION_TTL_SECONDS };
}

/**
 * Validates a session token from CacheService.
 * Auto-refreshes TTL if less than 4 hours remaining.
 * Returns session object { email, role, user_id, name, created, expires } or null.
 */
function validateSession(token) {
  if (!token || String(token).trim() === '') return null;

  const cached = CacheService.getScriptCache().get(SESSION_KEY_PREFIX + token);
  if (!cached) return null; // expired or never existed

  let session;
  try {
    session = JSON.parse(cached);
  } catch (e) {
    return null;
  }

  // Hard expiry double-check
  if (Date.now() > session.expires) {
    CacheService.getScriptCache().remove(SESSION_KEY_PREFIX + token);
    return null;
  }

  // Auto-refresh: if less than 4h remaining, extend TTL
  const remaining = Math.floor((session.expires - Date.now()) / 1000);
  if (remaining < SESSION_TTL_SECONDS * 0.5) {
    session.expires = Date.now() + SESSION_TTL_SECONDS * 1000;
    CacheService.getScriptCache().put(
      SESSION_KEY_PREFIX + token,
      JSON.stringify(session),
      SESSION_TTL_SECONDS
    );
  }

  return session; // { email, role, user_id, name, created, expires }
}

/**
 * Invalidates (deletes) a session token from CacheService.
 * Silently ignores missing tokens.
 */
function invalidateSession(token) {
  if (!token) return;
  const session = validateSession(token); // get before removing
  CacheService.getScriptCache().remove(SESSION_KEY_PREFIX + token);
  if (session) {
    batchAuditLog(token, [{
      entity_type: 'AUTH',
      entity_id:   session.email,
      action:      'LOGOUT',
      field_name:  null,
      old_value:   null,
      new_value:   null,
      user_email:  session.email,
      user_role:   session.role,
      note:        'Session invalidated'
    }]);
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// RATE LIMITING via CacheService
// No RATE_LIMIT tab. No Sheets writes. O(1) per request.
// Owner is exempt from rate limiting.
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Checks rate limit for the given email.
 * Returns { allowed: true } or { allowed: false, code, message }.
 */
function checkRateLimit(email) {
  // Owner is always exempt
  const user = getUserByEmail(email);
  if (user && user.role === 'Owner') return { allowed: true };

  const key   = RATE_LIMIT_PREFIX + email;
  const cache = CacheService.getScriptCache();
  const count = parseInt(cache.get(key) || '0', 10);

  if (count >= RATE_LIMIT_MAX) {
    // Log the rate limit event (best-effort — don't let this fail the request handling)
    try {
      batchAuditLog(Utilities.getUuid(), [{
        entity_type: 'SECURITY',
        entity_id:   email,
        action:      'RATE_LIMITED',
        field_name:  null,
        old_value:   null,
        new_value:   String(count),
        user_email:  email,
        user_role:   user ? user.role : '',
        note:        `Exceeded ${RATE_LIMIT_MAX} req/${RATE_LIMIT_WINDOW}s`
      }]);
    } catch (e) { /* ignore */ }

    return {
      allowed: false,
      code: 'RATE_LIMITED',
      message: 'Çox sayda sorğu. 1 dəqiqə sonra yenidən cəhd edin.'
    };
  }

  // Increment counter. First request sets RATE_LIMIT_WINDOW TTL.
  cache.put(key, String(count + 1), RATE_LIMIT_WINDOW);
  return { allowed: true };
}

// ─────────────────────────────────────────────────────────────────────────────
// PERMISSION CHECK
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Returns true if the role is allowed to perform the action.
 * Unknown actions → deny by default.
 */
function checkPermission(action, role) {
  const allowed = ROLE_PERMISSIONS[action];
  if (!allowed) return false;
  return allowed.indexOf(role) !== -1;
}

// ─────────────────────────────────────────────────────────────────────────────
// USER LOOKUP
// Reads MANAGERS tab. No caching (sheet changes must be reflected immediately).
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Returns user object from MANAGERS tab or null.
 * MANAGERS columns (0-based): A=0 manager_id, B=1 name, C=2 email,
 *   D=3 role, E=4 max_discount_pct, F=5 is_active
 */
function getUserByEmail(email) {
  if (!email) return null;

  const sheet = SpreadsheetApp.getActive().getSheetByName('MANAGERS');
  if (!sheet || sheet.getLastRow() < 2) return null;

  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][2]).trim().toLowerCase() === email.toLowerCase()) {
      return {
        manager_id: data[i][0],        // col A
        name:       data[i][1],        // col B
        email:      String(data[i][2]).trim(), // col C
        role:       data[i][3],        // col D
        is_active:  data[i][5]         // col F
      };
    }
  }
  return null;
}

// ─────────────────────────────────────────────────────────────────────────────
// AUTHENTICATE REQUEST — 5-step gate
// Call at the top of every doPost/doGet handler (except login/logout).
// Returns { ok: true, session } or { ok: false, resp: <ContentService output> }.
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Runs all 5 security checks. Returns { ok, session } or { ok: false, resp }.
 * response parameter = ContentService.createTextOutput() already set to JSON mime.
 */
function authenticateRequest(payload, action, response) {
  const token = payload.session_token;

  // ── 1. Token must be present ─────────────────────────────────────────────
  if (!token || String(token).trim() === '') {
    logSecurityEvent('ANONYMOUS_REQUEST', payload.email || payload.user_email || 'unknown', action, '');
    return {
      ok: false,
      resp: response.setContent(JSON.stringify({
        status: 'error', code: 'UNAUTHORIZED',
        message: 'Sessiya tapılmadı. Yenidən daxil olun.'
      }))
    };
  }

  // ── 2. Session must be valid ──────────────────────────────────────────────
  const session = validateSession(token);
  if (!session) {
    logSecurityEvent('EXPIRED_SESSION', token.slice(0, 8) + '...', action, '');
    return {
      ok: false,
      resp: response.setContent(JSON.stringify({
        status: 'error', code: 'SESSION_EXPIRED',
        message: 'Sessiyanız bitib. Yenidən daxil olun.'
      }))
    };
  }

  // ── 3. Rate limit check ───────────────────────────────────────────────────
  const rateCheck = checkRateLimit(session.email);
  if (!rateCheck.allowed) {
    return {
      ok: false,
      resp: response.setContent(JSON.stringify({
        status: 'error', code: 'RATE_LIMITED',
        message: rateCheck.message
      }))
    };
  }

  // ── 4. Role permission check ──────────────────────────────────────────────
  if (action && !checkPermission(action, session.role)) {
    logSecurityEvent('PERMISSION_DENIED', session.email, action, session.role);
    return {
      ok: false,
      resp: response.setContent(JSON.stringify({
        status: 'error', code: 'FORBIDDEN',
        message: 'Bu əməliyyat üçün icazəniz yoxdur. (' + action + ')'
      }))
    };
  }

  // ── 5. Timestamp freshness (write actions only — 5-minute window) ─────────
  const WRITE_ACTIONS = [
    'createOrder', 'updateOrder', 'confirmPayment',
    'changeOrderStatus', 'approveDiscount', 'createPayment'
  ];
  if (WRITE_ACTIONS.indexOf(action) !== -1 && payload.timestamp) {
    const age = Date.now() - parseInt(payload.timestamp, 10);
    if (age > 5 * 60 * 1000) {
      logSecurityEvent('STALE_TIMESTAMP', session.email, action, session.role);
      return {
        ok: false,
        resp: response.setContent(JSON.stringify({
          status: 'error', code: 'STALE_REQUEST',
          message: 'Sorğu vaxtı keçib. Formu yeniləyin.'
        }))
      };
    }
  }

  return { ok: true, session };
}

// ─────────────────────────────────────────────────────────────────────────────
// SECURITY EVENT LOGGING
// Calls batchAuditLog (in Code.gs — same GAS project).
// Fails silently — logging must never block business logic.
// ─────────────────────────────────────────────────────────────────────────────

function logSecurityEvent(eventType, identifier, action, role) {
  try {
    batchAuditLog(Utilities.getUuid(), [{
      entity_type: 'SECURITY',
      entity_id:   identifier || 'unknown',
      action:      eventType,
      field_name:  action  || null,
      old_value:   role    || null,
      new_value:   null,
      user_email:  identifier || 'unknown',
      user_role:   role    || '',
      note:        'Security event: ' + eventType + ' on action ' + (action || '—')
    }]);
  } catch (e) {
    Logger.log('Security log failed: ' + e.toString());
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// LOGIN HANDLER
// Called from doPost when action === 'login'.
// Verifies the submitted email against the Google OAuth identity (when available).
// With "Anyone" deployment, Session.getActiveUser().getEmail() returns "" for
// external users — the identity check is skipped in that case (conditional guard).
// ─────────────────────────────────────────────────────────────────────────────

function handleLogin(payload, response) {
  const email = String(payload.email || '').trim();

  if (!email) {
    return response.setContent(JSON.stringify({
      status: 'error', code: 'EMAIL_REQUIRED',
      message: 'Email tələb olunur.'
    }));
  }

  // Rate-limit on login attempts (before session exists)
  const rateCheck = checkRateLimit(email);
  if (!rateCheck.allowed) {
    return response.setContent(JSON.stringify({
      status: 'error', code: 'RATE_LIMITED',
      message: rateCheck.message
    }));
  }

  // Identity cross-check via Google OAuth (only when GAS can resolve the caller).
  // Returns "" for external/public deployments — skip check in that case.
  const activeEmail = Session.getActiveUser().getEmail();
  if (activeEmail && activeEmail !== '' && activeEmail !== email) {
    logSecurityEvent('EMAIL_MISMATCH', email, 'login', '');
    return response.setContent(JSON.stringify({
      status: 'error', code: 'IDENTITY_MISMATCH',
      message: 'Email uyğun gəlmir. Google hesabınızla daxil olun.'
    }));
  }

  // Look up user in MANAGERS
  const user = getUserByEmail(email);
  if (!user) {
    logSecurityEvent('UNKNOWN_USER', email, 'login', '');
    return response.setContent(JSON.stringify({
      status: 'error', code: 'USER_NOT_FOUND',
      message: 'Bu email sistemdə qeydiyyatda deyil. İnzibatçıya müraciət edin.'
    }));
  }

  if (!user.is_active) {
    logSecurityEvent('INACTIVE_USER', email, 'login', user.role);
    return response.setContent(JSON.stringify({
      status: 'error', code: 'ACCOUNT_INACTIVE',
      message: 'Hesabınız deaktivdir. İnzibatçıya müraciət edin.'
    }));
  }

  // Create session
  const result = createSession(email);
  if (!result) {
    return response.setContent(JSON.stringify({
      status: 'error', code: 'SESSION_FAILED',
      message: 'Sessiya yaradıla bilmədi.'
    }));
  }

  return response.setContent(JSON.stringify({
    status:        'success',
    session_token: result.token,
    role:          result.role,
    name:          result.name,
    manager_id:    user.manager_id,
    expires_in:    result.expires_in
  }));
}

// ─────────────────────────────────────────────────────────────────────────────
// LOGOUT HANDLER
// ─────────────────────────────────────────────────────────────────────────────

function handleLogout(payload, response) {
  invalidateSession(payload.session_token);
  return response.setContent(JSON.stringify({
    status: 'success', message: 'Çıxış edildi.'
  }));
}
