// ─────────────────────────────────────────────────────────────────────────────
// CODE.GS — MFN Tikinti Materialları CRM Backend
// Phase 2 — Full Apps Script Implementation
// Version: 2.0.0 | 2026-05-31
//
// Depends on: Security.gs (same Apps Script project)
// Sheet ID:   1pdAWEFvl-FjSQ0Sdm0g0gZF1vI8hAfCsX39UI6KSAuw
//
// Deployment: Execute as Me (owner), Who has access: Anyone
// ─────────────────────────────────────────────────────────────────────────────

'use strict';

// ── GLOBAL CONSTANTS ─────────────────────────────────────────────────────────

/** Telegram notification routing map */
const NOTIFICATION_ROUTING = {
  N1_NEW_ORDER:          'OPS',
  N2_DISCOUNT_ALERT:     'CRITICAL',
  N3_PRODUCTION_UNLOCK:  'OPS',
  N4_ORDER_READY:        'OPS',
  N5_RETURN:             'CRITICAL',
  N6_OVERDUE_PAYMENT:    'CRITICAL',
  N7_OVERDUE_PRODUCTION: 'OPS',
  N8_PROBLEM_ESCALATION: 'CRITICAL',
  N9_UNMATCHED_PAYMENT:  'CRITICAL',
  DAILY_REPORT:          'CRITICAL'
};

/**
 * Allowed action strings for batchAuditLog.
 * Do NOT add passive events like FORM_LOAD — audit log is for mutations only.
 */
const LOGGABLE_ACTIONS = [
  'CREATE', 'UPDATE', 'STATUS_CHANGE', 'PAYMENT_CONFIRM', 'DISCOUNT_REQUEST',
  'DISCOUNT_APPROVE', 'DISCOUNT_REJECT', 'CANCEL', 'LOCK_BYPASS_ATTEMPT',
  'PRODUCTION_BLOCK', 'PAYMENT_STATUS_REVERT', 'ADMIN_UNLOCK', 'REOPEN',
  'VALIDATION_ERROR', 'LOGIN_SUCCESS', 'LOGOUT', 'RATE_LIMITED',
  'PERMISSION_DENIED', 'ANONYMOUS_REQUEST', 'EXPIRED_SESSION',
  'INACTIVE_USER', 'STALE_TIMESTAMP'
];

/**
 * Columns that handleUpdate can NEVER overwrite.
 * Formula columns, system columns, and write-once columns.
 */
const PROTECTED_COLUMNS = [
  'Order_ID', 'Request_ID', 'Created_At', 'Created_By',
  'Client_Name', 'Client_Type',                        // FORMULA — VLOOKUP from CLIENTS
  'Discount_Amount', 'Net_Amount', 'VAT_Pct',          // FORMULA
  'VAT_Amount', 'Grand_Total', 'Avans_Required',       // FORMULA
  'Total_Paid', 'Debt_Amount',                         // FORMULA
  'Production_Lock', 'Problem_Flag',                   // FORMULA
  'Overdue_Flag', 'Manager_ID',                        // FORMULA
  'Price_Snapshot_JSON',                               // write-once
  'Payment_Status', 'Discount_Approval_Status',        // protected status fields
  'Edit_Count'                                         // script-managed counter
];

// ─────────────────────────────────────────────────────────────────────────────
// FUNCTION 1: generateOrderId
// Uses LockService to guarantee unique, sequential IDs under concurrent requests.
// ─────────────────────────────────────────────────────────────────────────────

function generateOrderId() {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(5000);
    const counterRange = SpreadsheetApp.getActive().getRangeByName('ORDER_COUNTER_CELL');
    const current = Number(counterRange.getValue()) || 0;
    const next    = current + 1;
    counterRange.setValue(next);
    const year = new Date().getFullYear();
    return 'INV-' + year + '-' + String(next).padStart(4, '0');
  } finally {
    lock.releaseLock();
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// FUNCTION 2: doPost — main form handler / action router
// ─────────────────────────────────────────────────────────────────────────────

function doPost(e) {
  const response = ContentService.createTextOutput();
  response.setMimeType(ContentService.MimeType.JSON);

  try {
    const payload = JSON.parse(e.postData.contents);
    const action  = payload.action
      || (payload.order_id && String(payload.order_id).trim() !== ''
          ? 'updateOrder'
          : 'createOrder');

    // ── Login / Logout bypass the auth gate ─────────────────────────────────
    if (action === 'login')  return handleLogin(payload, response);   // Security.gs
    if (action === 'logout') return handleLogout(payload, response);  // Security.gs

    // ── Security gate ────────────────────────────────────────────────────────
    const authResult = authenticateRequest(payload, action, response); // Security.gs
    if (!authResult.ok) return authResult.resp;
    const session = authResult.session;

    // ── Idempotency (CREATE only) ─────────────────────────────────────────────
    if (action === 'createOrder') {
      const idem = checkIdempotency(payload.request_id);
      if (idem) return response.setContent(JSON.stringify(idem));
    }

    // ── Server-side field validation (create / update only) ───────────────────
    if (action === 'createOrder' || action === 'updateOrder') {
      const errs = serverValidate(payload, action);
      if (errs.length > 0) {
        return response.setContent(JSON.stringify({
          status: 'error', code: 'VALIDATION_FAILED', errors: errs
        }));
      }
    }

    // ── Action router ────────────────────────────────────────────────────────
    if (action === 'createOrder')            return handleCreate(payload, response, session);
    if (action === 'updateOrder')            return handleUpdate(payload, response, session);
    if (action === 'confirmPayment')         return handleConfirmPayment(payload, response, session);
    if (action === 'changeOrderStatus')      return handleStatusChange(payload, response, session);
    if (action === 'approveDiscount')        return handleDiscountApproval(payload, response, session);
    if (action === 'createPayment')          return handleCreatePayment(payload, response, session);
    if (action === 'updateProductionStatus') return handleProductionUpdate(payload, response, session);
    if (action === 'updateDeliveryStatus')   return handleDeliveryUpdate(payload, response, session);
    if (action === 'manualBackup') {
      weeklyBackup();
      return response.setContent(JSON.stringify({ status: 'ok', message: 'Backup started.' }));
    }

    return response.setContent(JSON.stringify({
      status: 'error', code: 'UNKNOWN_ACTION',
      message: 'Naməlum əməliyyat: ' + action
    }));

  } catch (err) {
    Logger.log('doPost error: ' + err.toString() + '\n' + (err.stack || ''));
    return response.setContent(JSON.stringify({
      status: 'error', code: 'SERVER_ERROR', message: err.message
    }));
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// doGet — read-only endpoints
// ─────────────────────────────────────────────────────────────────────────────

function doGet(e) {
  const response = ContentService.createTextOutput();
  response.setMimeType(ContentService.MimeType.JSON);

  const action = e.parameter.action;

  // Warmup ping — no auth required
  if (action === 'ping') {
    return response.setContent(JSON.stringify({ status: 'ok', ts: Date.now() }));
  }

  // getPriceTable — semi-public: price data is not sensitive, needed before login
  if (action === 'getPriceTable') return servePriceTable(response);

  // All other GET actions require auth
  const token      = e.parameter.session_token;
  const authResult = authenticateRequest({ session_token: token }, action, response);
  if (!authResult.ok) return authResult.resp;
  const session = authResult.session;

  if (action === 'getOrder')     return serveOrder(e.parameter.order_id, session, response);
  if (action === 'getDashboard') return serveDashboard(session, response);

  return response.setContent(JSON.stringify({
    status: 'error', code: 'UNKNOWN_ACTION',
    message: 'Unknown GET action: ' + action
  }));
}

// ─────────────────────────────────────────────────────────────────────────────
// FUNCTION 3: checkIdempotency
// Scans ORDERS col B (Request_ID) for a match.
// Returns existing order response object, or null if not found.
// ─────────────────────────────────────────────────────────────────────────────

function checkIdempotency(requestId) {
  if (!requestId) return null;

  const sheet   = SpreadsheetApp.getActive().getSheetByName('ORDERS');
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;

  const requestIds = sheet.getRange(2, 2, lastRow - 1, 1).getValues().flat();
  const matchIndex = requestIds.indexOf(requestId);
  if (matchIndex === -1) return null;

  const orderId = sheet.getRange(matchIndex + 2, 1).getValue();
  return {
    status:     'success',
    order_id:   orderId,
    idempotent: true,
    message:    'Bu sessiya üçün sifariş artıq mövcuddur'
  };
}

// ─────────────────────────────────────────────────────────────────────────────
// FUNCTION 4: serverValidate
// Returns array of error objects. Empty array = valid.
// ─────────────────────────────────────────────────────────────────────────────

function serverValidate(payload, action) {
  const errors   = [];
  const isCreate = !action || action === 'createOrder';

  // client_name
  if (isCreate || payload.hasOwnProperty('client_name')) {
    if (!payload.client_name || String(payload.client_name).trim().length < 2)
      errors.push({ field: 'client_name', code: 'REQUIRED', message: 'Müştəri adı tələb olunur' });
  }

  // client_phone
  if (isCreate || payload.hasOwnProperty('client_phone')) {
    if (!payload.client_phone || !isValidAzPhone(String(payload.client_phone)))
      errors.push({ field: 'client_phone', code: 'INVALID_PHONE', message: 'Düzgün AZ telefon nömrəsi daxil edin' });
  }

  // manager_id
  if (isCreate || payload.hasOwnProperty('manager_id')) {
    if (!payload.manager_id || String(payload.manager_id).trim() === '')
      errors.push({ field: 'manager_id', code: 'REQUIRED', message: 'Menecer seçilməlidir' });
  }

  // payment_terms
  if (isCreate || payload.hasOwnProperty('payment_terms')) {
    if (!payload.payment_terms || payload.payment_terms === '— Seçin —')
      errors.push({ field: 'payment_terms', code: 'REQUIRED', message: 'Ödəniş şərtləri seçilməlidir' });
  }

  // line_items
  if (isCreate || payload.hasOwnProperty('line_items')) {
    if (!payload.line_items || payload.line_items.length === 0) {
      errors.push({ field: 'line_items', code: 'NO_ITEMS', message: 'Ən azı bir material əlavə edin' });
    } else {
      payload.line_items.forEach(function (item, i) {
        if (!item.sku)
          errors.push({ field: 'line_' + i + '_sku', code: 'REQUIRED', message: 'Sətir ' + (i + 1) + ': SKU seçilməlidir' });
        if (!item.qty || item.qty <= 0)
          errors.push({ field: 'line_' + i + '_qty', code: 'MUST_BE_POSITIVE', message: 'Sətir ' + (i + 1) + ': Miqdar > 0 olmalıdır' });
        if (!item.unit_price || item.unit_price <= 0)
          errors.push({ field: 'line_' + i + '_price', code: 'MUST_BE_POSITIVE', message: 'Sətir ' + (i + 1) + ': Qiymət > 0 olmalıdır' });
        // Price floor (below Topdan minimum)
        var floor = getPriceFloor(item.sku);
        if (floor && item.unit_price < floor) {
          errors.push({
            field:     'line_' + i + '_price',
            code:      'BELOW_MINIMUM',
            message:   'Sətir ' + (i + 1) + ': ' + item.sku + ' minimum qiymət ' + floor + ' AZN (topdan)',
            min:       floor,
            submitted: item.unit_price
          });
        }
      });
    }
  }

  // grand_total (required on create)
  if (isCreate && (!payload.grand_total || payload.grand_total <= 0))
    errors.push({ field: 'grand_total', code: 'ZERO_TOTAL', message: 'Ümumi məbləğ sıfır ola bilməz' });

  return errors;
}

/** Validates Azerbaijani mobile phone numbers. */
function isValidAzPhone(phone) {
  const digits = phone.replace(/\D/g, '');
  const norm   = digits.startsWith('994') ? digits.slice(3)
               : digits.startsWith('0')   ? digits.slice(1)
               : digits;
  const validPrefixes = ['50', '51', '55', '60', '70', '77'];
  return norm.length === 9 && validPrefixes.indexOf(norm.slice(0, 2)) !== -1;
}

/** Returns the Topdan (floor) price for a SKU from PRICE_TABLE, or null if not found. */
function getPriceFloor(sku) {
  const sheet = SpreadsheetApp.getActive().getSheetByName('PRICE_TABLE');
  if (!sheet || sheet.getLastRow() < 2) return null;
  const data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    // col A (0) = SKU, col K (10) = Is_Active, col F (5) = Price_Topdan
    if (data[i][0] === sku && data[i][10] === true) return data[i][5];
  }
  return null;
}

// ─────────────────────────────────────────────────────────────────────────────
// FUNCTION 5: handleCreate
// Writes new ORDERS row, per-row formulas, discount log, audit, Telegram.
// CRITICAL: uses session.email — NOT Session.getActiveUser().getEmail()
//   (getActiveUser() returns "" for external users with "Anyone" deployment)
// ─────────────────────────────────────────────────────────────────────────────

function handleCreate(payload, response, session) {
  const orderId   = generateOrderId();
  const now       = new Date();
  const sessionId = Utilities.getUuid();
  const userEmail = session.email; // ← correct: from validated CacheService session

  // Determine discount escalation BEFORE writing (may affect Q column)
  const discountStatus = getDiscountApprovalStatus(payload.price_type, payload.discount_pct);

  // Append raw data row (formula columns get '' — filled by writeOrderFormulas below)
  const sheet  = SpreadsheetApp.getActive().getSheetByName('ORDERS');
  const newRow = buildOrderRow(payload, orderId, userEmail, now, discountStatus);
  sheet.appendRow(newRow);

  // Write per-row formulas for all ORDERS formula columns
  const appendedRow = sheet.getLastRow();
  writeOrderFormulas(sheet, appendedRow);

  // Discount escalation handling
  if (discountStatus !== 'N/A') {
    createDiscountLogRow(orderId, payload, userEmail, discountStatus);
    sendDiscountAlert(orderId, payload, userEmail);
  }

  // Audit log
  batchAuditLog(sessionId, [{
    entity_type: 'ORDER',
    entity_id:   orderId,
    action:      'CREATE',
    field_name:  null,
    old_value:   null,
    new_value:   'Yeni',
    user_email:  userEmail,
    user_role:   session.role,
    note:        'Order created via form'
  }]);

  // Telegram N1 → OPS group
  sendTelegram('N1_NEW_ORDER',
    '📋 Yeni sifariş: <b>' + orderId + '</b>\n' +
    '👤 ' + payload.client_name + '\n' +
    '💰 ' + (payload.grand_total || 0) + ' AZN\n' +
    '👨‍💼 ' + userEmail
  );

  return response.setContent(JSON.stringify({
    status:     'success',
    order_id:   orderId,
    session_id: sessionId
  }));
}

/**
 * Builds the 44-element array for ORDERS.appendRow().
 * Formula columns (J,K,U,V,W,X,AA,AC,AD,AE,AH,AI,AJ,AL) receive '' here.
 * writeOrderFormulas() fills them with per-row formulas after appendRow.
 */
function buildOrderRow(payload, orderId, userEmail, now, discountStatus) {
  return [
    orderId,                                    // A  Order_ID
    payload.request_id       || '',             // B  Request_ID      (idempotency key)
    now,                                        // C  Created_At
    userEmail,                                  // D  Created_By
    now,                                        // E  Updated_At
    userEmail,                                  // F  Updated_By
    1,                                          // G  Edit_Count
    '',                                         // H  Invoice_Link     (manager editable)
    payload.client_id        || '',             // I  Client_ID        (manager enters)
    '',                                         // J  Client_Name      ← FORMULA
    '',                                         // K  Client_Type      ← FORMULA
    payload.price_type       || 'Perakende',    // L  Price_Type
    'Yeni',                                     // M  Order_Status
    'Ödənilmədi',                               // N  Payment_Status
    'Gözlənilir',                               // O  Production_Status
    'Planlanmadı',                              // P  Delivery_Status
    discountStatus,                             // Q  Discount_Approval_Status
    payload.payment_terms    || '',             // R  Payment_Terms
    payload.subtotal         || 0,             // S  Subtotal_AZN
    payload.discount_pct     || 0,             // T  Discount_Pct
    '',                                         // U  Discount_Amount  ← FORMULA
    '',                                         // V  Net_Amount       ← FORMULA
    '',                                         // W  VAT_Pct          ← FORMULA
    '',                                         // X  VAT_Amount       ← FORMULA
    payload.delivery_zone    || '',             // Y  Delivery_Zone
    payload.delivery_fee     || 0,             // Z  Delivery_Fee
    '',                                         // AA Grand_Total      ← FORMULA
    JSON.stringify(payload.line_items || []),   // AB Price_Snapshot_JSON (write-once)
    '',                                         // AC Avans_Required   ← FORMULA
    '',                                         // AD Total_Paid       ← FORMULA
    '',                                         // AE Debt_Amount      ← FORMULA
    payload.payment_due_date || '',             // AF Payment_Due_Date
    '',                                         // AG Receipt_Link
    '',                                         // AH Production_Lock  ← FORMULA
    '',                                         // AI Problem_Flag     ← FORMULA
    '',                                         // AJ Overdue_Flag     ← FORMULA
    false,                                      // AK Price_Below_Min_Flag (script sets)
    '',                                         // AL Manager_ID       ← FORMULA
    payload.notes            || '',             // AM Notes
    payload.delivery_address || '',             // AN Delivery_Address
    payload.delivery_contact || '',             // AO Delivery_Contact
    payload.delivery_phone   || '',             // AP Delivery_Phone
    payload.planned_delivery_date || '',        // AQ Planned_Delivery_Date
    ''                                          // AR Actual_Delivery_Date
  ];
}

/**
 * Writes per-row formulas for all ORDERS formula columns.
 * Must be called immediately after appendRow with the row number.
 * Uses INDEX/MATCH for Manager_ID (VLOOKUP cannot search right-to-left).
 */
function writeOrderFormulas(sheet, row) {
  var r = row;
  sheet.getRange(r, 10).setFormula('=IFERROR(VLOOKUP(I' + r + ',CLIENTS!A:B,2,0),"")');     // J Client_Name
  sheet.getRange(r, 11).setFormula('=IFERROR(VLOOKUP(I' + r + ',CLIENTS!A:C,3,0),"")');     // K Client_Type
  sheet.getRange(r, 21).setFormula('=S' + r + '*T' + r + '/100');                           // U Discount_Amount
  sheet.getRange(r, 22).setFormula('=S' + r + '-U' + r);                                    // V Net_Amount
  sheet.getRange(r, 23).setFormula('=VAT_RATE');                                             // W VAT_Pct (named range)
  sheet.getRange(r, 24).setFormula('=V' + r + '*W' + r + '/100');                           // X VAT_Amount
  sheet.getRange(r, 27).setFormula('=V' + r + '+X' + r + '+Z' + r);                        // AA Grand_Total
  sheet.getRange(r, 29).setFormula(
    '=IF(R' + r + '="50/50",AA' + r + '*0.5,' +
      'IF(R' + r + '="Tam əvvəl",AA' + r + ',0))'
  );                                                                                         // AC Avans_Required
  sheet.getRange(r, 30).setFormula(
    '=IFERROR(SUMIF(PAYMENTS!B:B,A' + r + ',PAYMENTS!E:E),0)'
  );                                                                                         // AD Total_Paid
  sheet.getRange(r, 31).setFormula('=MAX(0,AA' + r + '-AD' + r + ')');                      // AE Debt_Amount
  sheet.getRange(r, 34).setFormula(
    '=IF(AND(N' + r + '="CONFIRMED",AD' + r + '>=AC' + r + '),FALSE,TRUE)'
  );                                                                                         // AH Production_Lock
  sheet.getRange(r, 35).setFormula(
    '=IF(OR(AE' + r + '>0,' +
      'Q' + r + '="PENDING",' +
      'Q' + r + '="PENDING_DIRECTOR",' +
      'AND(O' + r + '="İstehsalda",N' + r + '<>"CONFIRMED"),' +
      'AJ' + r + '=TRUE,' +
      'AK' + r + '=TRUE),TRUE,FALSE)'
  );                                                                                         // AI Problem_Flag
  sheet.getRange(r, 36).setFormula(
    '=IF(AND(AE' + r + '>0,AF' + r + '<>"",AF' + r + '<TODAY()),TRUE,FALSE)'
  );                                                                                         // AJ Overdue_Flag
  sheet.getRange(r, 38).setFormula(
    '=IFERROR(INDEX(MANAGERS!A:A,MATCH(D' + r + ',MANAGERS!C:C,0)),"")'
  );                                                                                         // AL Manager_ID
}

// ─────────────────────────────────────────────────────────────────────────────
// FUNCTION 6: handleUpdate — PATCH semantics
// Only writes fields present in payload that are NOT protected.
// ─────────────────────────────────────────────────────────────────────────────

function handleUpdate(payload, response, session) {
  var userEmail = session.email; // ← CRITICAL: use session.email
  var sessionId = payload.session_id || Utilities.getUuid();

  // Server-side production/price lock enforcement
  var lockResult = enforceServerLock(payload.order_id, payload, userEmail);
  if (lockResult.blocked) {
    batchAuditLog(sessionId, [{
      entity_type: 'ORDER',
      entity_id:   payload.order_id,
      action:      'LOCK_BYPASS_ATTEMPT',
      field_name:  lockResult.violations.map(function (v) { return v.field; }).join(','),
      old_value:   null,
      new_value:   null,
      user_email:  userEmail,
      user_role:   session.role,
      note:        'Blocked: ' + JSON.stringify(lockResult.violations)
    }]);
    return response.setContent(JSON.stringify({
      status: 'error', code: 'FIELD_LOCKED', violations: lockResult.violations
    }));
  }

  var sheet    = SpreadsheetApp.getActive().getSheetByName('ORDERS');
  var colMap   = getOrderColumnMap();
  var orderRow = findOrderRow(payload.order_id);

  if (!orderRow) {
    return response.setContent(JSON.stringify({ status: 'error', code: 'ORDER_NOT_FOUND' }));
  }

  // Internal/meta fields to skip
  var META_FIELDS = ['action', 'session_token', 'timestamp', 'request_id',
                     'order_id', 'session_id', 'line_items'];

  var changes = [];
  Object.keys(payload).forEach(function (field) {
    if (PROTECTED_COLUMNS.indexOf(field) !== -1) return;
    if (META_FIELDS.indexOf(field)   !== -1) return;
    var col = colMap[field];
    if (!col) return;
    var oldValue = sheet.getRange(orderRow, col).getValue();
    if (String(oldValue) !== String(payload[field])) {
      sheet.getRange(orderRow, col).setValue(payload[field]);
      changes.push({ field_name: field, old_value: oldValue, new_value: payload[field] });
    }
  });

  // line_items: update Price_Snapshot_JSON only if order is still in draft
  if (payload.line_items) {
    var currentStatus = sheet.getRange(orderRow, colMap['Order_Status']).getValue();
    if (currentStatus === 'Yeni') {
      sheet.getRange(orderRow, colMap['Price_Snapshot_JSON'])
           .setValue(JSON.stringify(payload.line_items));
      changes.push({ field_name: 'Price_Snapshot_JSON', old_value: '(prev)', new_value: '(updated)' });
    }
  }

  // Metadata always updated
  var now = new Date();
  sheet.getRange(orderRow, colMap['Updated_At']).setValue(now);
  sheet.getRange(orderRow, colMap['Updated_By']).setValue(userEmail);
  var editCount = Number(sheet.getRange(orderRow, colMap['Edit_Count']).getValue() || 0);
  sheet.getRange(orderRow, colMap['Edit_Count']).setValue(editCount + 1);

  // Audit log
  if (changes.length > 0) {
    batchAuditLog(sessionId, changes.map(function (c) {
      return {
        entity_type: 'ORDER',
        entity_id:   payload.order_id,
        action:      'UPDATE',
        field_name:  c.field_name,
        old_value:   String(c.old_value),
        new_value:   String(c.new_value),
        user_email:  userEmail,
        user_role:   session.role,
        note:        null
      };
    }));
  }

  return response.setContent(JSON.stringify({
    status:        'success',
    order_id:      payload.order_id,
    changes_count: changes.length
  }));
}

// ─────────────────────────────────────────────────────────────────────────────
// Additional action handlers
// ─────────────────────────────────────────────────────────────────────────────

/** confirmPayment — Admin/Owner only. Verifies PAYMENTS row exists first. */
function handleConfirmPayment(payload, response, session) {
  var orderId  = payload.order_id;
  var orderRow = findOrderRow(orderId);
  if (!orderRow)
    return response.setContent(JSON.stringify({ status: 'error', code: 'ORDER_NOT_FOUND' }));

  var sheet = SpreadsheetApp.getActive().getSheetByName('ORDERS');
  var COL   = getOrderColumnMap();

  // GUARD: cannot confirm without a payment record
  var payments = getPaymentsForOrder(orderId);
  if (payments.length === 0) {
    return response.setContent(JSON.stringify({
      status: 'error', code: 'NO_PAYMENT_RECORDS',
      message: 'Ödəniş qeydi olmadan CONFIRMED qoya bilməzsiniz.'
    }));
  }

  var oldStatus = sheet.getRange(orderRow, COL['Payment_Status']).getValue();
  sheet.getRange(orderRow, COL['Payment_Status']).setValue('CONFIRMED');
  sheet.getRange(orderRow, COL['Updated_At']).setValue(new Date());
  sheet.getRange(orderRow, COL['Updated_By']).setValue(session.email);

  checkAndUnlockProduction(orderId, orderRow);

  batchAuditLog(payload.session_id || Utilities.getUuid(), [{
    entity_type: 'ORDER',   entity_id: orderId,
    action:      'PAYMENT_CONFIRM',
    field_name:  'Payment_Status',
    old_value:   oldStatus, new_value: 'CONFIRMED',
    user_email:  session.email, user_role: session.role,
    note:        null
  }]);

  return response.setContent(JSON.stringify({ status: 'success', order_id: orderId }));
}

/** changeOrderStatus — Senior Manager / Admin / Owner only. */
function handleStatusChange(payload, response, session) {
  var orderId   = payload.order_id;
  var newStatus = payload.order_status;
  var orderRow  = findOrderRow(orderId);
  if (!orderRow)
    return response.setContent(JSON.stringify({ status: 'error', code: 'ORDER_NOT_FOUND' }));

  var sheet     = SpreadsheetApp.getActive().getSheetByName('ORDERS');
  var COL       = getOrderColumnMap();
  var oldStatus = sheet.getRange(orderRow, COL['Order_Status']).getValue();

  sheet.getRange(orderRow, COL['Order_Status']).setValue(newStatus);
  sheet.getRange(orderRow, COL['Updated_At']).setValue(new Date());
  sheet.getRange(orderRow, COL['Updated_By']).setValue(session.email);

  if (newStatus === 'Ləğv edildi') {
    sendTelegram('N8_PROBLEM_ESCALATION',
      '🚫 Sifariş ləğv edildi: <b>' + orderId + '</b>\n' +
      '👤 ' + oldStatus + ' → Ləğv edildi\n' +
      '👨‍💼 ' + session.email
    );
  }

  batchAuditLog(payload.session_id || Utilities.getUuid(), [{
    entity_type: 'ORDER',   entity_id: orderId,
    action:      'STATUS_CHANGE',
    field_name:  'Order_Status',
    old_value:   oldStatus, new_value: newStatus,
    user_email:  session.email, user_role: session.role, note: null
  }]);

  return response.setContent(JSON.stringify({ status: 'success', order_id: orderId }));
}

/** approveDiscount / rejectDiscount — Senior Manager / Admin / Owner only. */
function handleDiscountApproval(payload, response, session) {
  var orderId         = payload.order_id;
  var approved        = payload.approved === true || payload.approved === 'true';
  var rejectionReason = payload.rejection_reason || '';

  var orderRow = findOrderRow(orderId);
  if (!orderRow)
    return response.setContent(JSON.stringify({ status: 'error', code: 'ORDER_NOT_FOUND' }));

  var sheet     = SpreadsheetApp.getActive().getSheetByName('ORDERS');
  var COL       = getOrderColumnMap();
  var newStatus = approved ? 'APPROVED' : 'REJECTED';

  sheet.getRange(orderRow, COL['Discount_Approval_Status']).setValue(newStatus);
  sheet.getRange(orderRow, COL['Updated_At']).setValue(new Date());
  sheet.getRange(orderRow, COL['Updated_By']).setValue(session.email);

  // Update matching PENDING row in DISCOUNT_LOG
  var discLog  = SpreadsheetApp.getActive().getSheetByName('DISCOUNT_LOG');
  var discData = discLog.getDataRange().getValues();
  var PENDING_STATUSES = ['PENDING', 'PENDING_DIRECTOR', 'PENDING_SR_MGR'];
  for (var i = 1; i < discData.length; i++) {
    if (discData[i][1] === orderId && PENDING_STATUSES.indexOf(discData[i][9]) !== -1) {
      discLog.getRange(i + 1, 10).setValue(newStatus);       // J Approval_Status
      discLog.getRange(i + 1, 11).setValue(session.email);   // K Approved_By
      discLog.getRange(i + 1, 12).setValue(new Date());      // L Approved_At
      if (!approved && rejectionReason)
        discLog.getRange(i + 1, 13).setValue(rejectionReason); // M Rejection_Reason
      break;
    }
  }

  var action = approved ? 'DISCOUNT_APPROVE' : 'DISCOUNT_REJECT';
  batchAuditLog(payload.session_id || Utilities.getUuid(), [{
    entity_type: 'ORDER',   entity_id: orderId,
    action:      action,
    field_name:  'Discount_Approval_Status',
    old_value:   'PENDING', new_value: newStatus,
    user_email:  session.email, user_role: session.role,
    note:        rejectionReason || null
  }]);

  return response.setContent(JSON.stringify({ status: 'success', order_id: orderId, decision: newStatus }));
}

/** createPayment — Admin / Owner only. Appends row + per-row formulas. */
function handleCreatePayment(payload, response, session) {
  var orderId = payload.order_id;
  if (!orderId)
    return response.setContent(JSON.stringify({ status: 'error', code: 'ORDER_ID_REQUIRED' }));

  // Bank payment requires receipt_link
  if (payload.payment_method === 'Bank köçürməsi' && !payload.receipt_link) {
    return response.setContent(JSON.stringify({
      status: 'error', code: 'RECEIPT_REQUIRED',
      message: 'Bank köçürməsi üçün qəbz linki tələb olunur.'
    }));
  }

  var amount = parseFloat(payload.amount_azn);
  if (!amount || isNaN(amount))
    return response.setContent(JSON.stringify({ status: 'error', code: 'INVALID_AMOUNT', message: 'Məbləğ 0 ola bilməz.' }));

  // Cash legal limit alert
  var cashLimit = parseFloat(getSettingValue('CASH_PAYMENT_LEGAL_LIMIT') || 30000);
  if (payload.payment_method === 'Nağd' && Math.abs(amount) >= cashLimit) {
    sendTelegram('N2_DISCOUNT_ALERT',
      '⚠️ Nağd ödəniş limiti: <b>' + orderId + '</b>\n' +
      '💰 ' + amount + ' AZN — hüquqi hədd (' + cashLimit + ' AZN) aşıldı.\n' +
      '👨‍💼 ' + session.email
    );
  }

  var paySheet = SpreadsheetApp.getActive().getSheetByName('PAYMENTS');
  var payId    = 'PAY-' + String(paySheet.getLastRow()).padStart(5, '0');
  var now      = new Date();

  // Append 14-column row (C, M, N = formulas — written below)
  paySheet.appendRow([
    payId,                                                        // A  Payment_ID
    orderId,                                                      // B  Order_ID
    '',                                                           // C  Client_ID     ← FORMULA
    payload.payment_date ? new Date(payload.payment_date) : now, // D  Payment_Date
    amount,                                                       // E  Amount_AZN
    payload.payment_method || 'Bank köçürməsi',                  // F  Payment_Method
    payload.payment_type   || 'Yekun ödəniş',                   // G  Payment_Type
    session.email,                                                // H  Confirmed_By  (auto: session only)
    now,                                                          // I  Confirmed_At  (auto)
    payload.receipt_link   || '',                                 // J  Receipt_Link
    payload.bank_reference || '',                                 // K  Bank_Reference
    payload.note           || '',                                 // L  Note
    '',                                                           // M  Order_Grand_Total ← FORMULA
    ''                                                            // N  Running_Balance   ← FORMULA
  ]);

  // Write per-row formulas for PAYMENTS formula columns
  var payRow = paySheet.getLastRow();
  paySheet.getRange(payRow, 3).setFormula(
    '=IFERROR(VLOOKUP(B' + payRow + ',ORDERS!A:I,9,0),"")'
  );                                                                          // C Client_ID
  paySheet.getRange(payRow, 13).setFormula(
    '=IFERROR(VLOOKUP(B' + payRow + ',ORDERS!A:AA,27,0),"")'
  );                                                                          // M Order_Grand_Total
  paySheet.getRange(payRow, 14).setFormula(
    '=M' + payRow + '-SUMIF(B$2:B' + payRow + ',B' + payRow + ',E$2:E' + payRow + ')'
  );                                                                          // N Running_Balance

  batchAuditLog(payload.session_id || Utilities.getUuid(), [{
    entity_type: 'PAYMENT', entity_id: payId,
    action:      'PAYMENT_CONFIRM',
    field_name:  'Amount_AZN',
    old_value:   null, new_value: String(amount),
    user_email:  session.email, user_role: session.role,
    note:        'Order: ' + orderId
  }]);

  return response.setContent(JSON.stringify({ status: 'success', payment_id: payId, order_id: orderId }));
}

/** updateProductionStatus — Production / Admin / Owner only. */
function handleProductionUpdate(payload, response, session) {
  var prodSheet = SpreadsheetApp.getActive().getSheetByName('PRODUCTION');
  var data      = prodSheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][1] === payload.order_id) { // col B = Order_ID
      var newStatus = payload.prod_status;
      var oldStatus = data[i][5];          // col F = Prod_Status
      prodSheet.getRange(i + 1, 6).setValue(newStatus);
      if (newStatus === 'Hazırdır') {
        prodSheet.getRange(i + 1, 9).setValue(new Date()); // I Actual_Ready
      }

      // Mirror to ORDERS Production_Status
      var orderRow = findOrderRow(payload.order_id);
      if (orderRow) {
        var ordSheet = SpreadsheetApp.getActive().getSheetByName('ORDERS');
        var COL      = getOrderColumnMap();
        ordSheet.getRange(orderRow, COL['Production_Status']).setValue(newStatus);
        if (newStatus === 'Hazırdır') {
          createDeliveryRow(payload.order_id, orderRow);
          var clientName = ordSheet.getRange(orderRow, COL['Client_Name']).getValue();
          sendTelegram('N4_ORDER_READY',
            '✅ Hazırdır: <b>' + payload.order_id + '</b>\n' +
            '👤 ' + clientName + '\nÇatdırılma planlanmalıdır.'
          );
        }
      }

      batchAuditLog(payload.session_id || Utilities.getUuid(), [{
        entity_type: 'PRODUCTION', entity_id: payload.order_id,
        action: 'UPDATE', field_name: 'Prod_Status',
        old_value: String(oldStatus), new_value: newStatus,
        user_email: session.email, user_role: session.role, note: null
      }]);
      return response.setContent(JSON.stringify({ status: 'success', order_id: payload.order_id }));
    }
  }
  return response.setContent(JSON.stringify({ status: 'error', code: 'PRODUCTION_ROW_NOT_FOUND' }));
}

/** updateDeliveryStatus — Delivery / Admin / Owner only. */
function handleDeliveryUpdate(payload, response, session) {
  var delSheet = SpreadsheetApp.getActive().getSheetByName('DELIVERY');
  var data     = delSheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][1] === payload.order_id) {
      var newStatus = payload.delivery_status;
      var oldStatus = data[i][6]; // col G = Delivery_Status
      delSheet.getRange(i + 1, 7).setValue(newStatus);

      var orderRow = findOrderRow(payload.order_id);
      if (newStatus === 'Çatdırıldı') {
        delSheet.getRange(i + 1, 9).setValue(new Date()); // I Actual_Date
        if (orderRow) {
          var ordSheet = SpreadsheetApp.getActive().getSheetByName('ORDERS');
          var COL      = getOrderColumnMap();
          ordSheet.getRange(orderRow, COL['Delivery_Status']).setValue('Çatdırıldı');
          ordSheet.getRange(orderRow, COL['Actual_Delivery_Date']).setValue(new Date());
          ordSheet.getRange(orderRow, COL['Order_Status']).setValue('Çatdırıldı');
        }
      }
      if (newStatus === 'Qaytarıldı') {
        if (payload.return_reason)
          delSheet.getRange(i + 1, 11).setValue(payload.return_reason); // K Return_Reason
        if (orderRow) {
          SpreadsheetApp.getActive().getSheetByName('ORDERS')
            .getRange(orderRow, getOrderColumnMap()['Order_Status'])
            .setValue('Problem/Mübahisə');
        }
        sendTelegram('N5_RETURN',
          '🚨 QAYTARILDI: <b>' + payload.order_id + '</b>\n' +
          'Səbəb: ' + (payload.return_reason || '—')
        );
      }

      batchAuditLog(payload.session_id || Utilities.getUuid(), [{
        entity_type: 'DELIVERY', entity_id: payload.order_id,
        action: 'STATUS_CHANGE', field_name: 'Delivery_Status',
        old_value: String(oldStatus), new_value: newStatus,
        user_email: session.email, user_role: session.role, note: null
      }]);
      return response.setContent(JSON.stringify({ status: 'success', order_id: payload.order_id }));
    }
  }
  return response.setContent(JSON.stringify({ status: 'error', code: 'DELIVERY_ROW_NOT_FOUND' }));
}

// ─────────────────────────────────────────────────────────────────────────────
// doGet helpers
// ─────────────────────────────────────────────────────────────────────────────

function servePriceTable(response) {
  var sheet = SpreadsheetApp.getActive().getSheetByName('PRICE_TABLE');
  if (!sheet) return response.setContent(JSON.stringify({ status: 'ok', items: [] }));
  var data  = sheet.getDataRange().getValues();
  var items = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][10] !== true) continue; // col K (index 10) = Is_Active
    items.push({
      sku:        data[i][0],
      name:       data[i][1],
      category:   data[i][2],
      unit:       data[i][3],
      perakende:  data[i][4],
      topdan:     data[i][5],
      dimensions: data[i][7]
    });
  }
  return response.setContent(JSON.stringify({ status: 'ok', items: items }));
}

function serveOrder(orderId, session, response) {
  if (!orderId)
    return response.setContent(JSON.stringify({ status: 'error', code: 'ORDER_ID_REQUIRED' }));
  var orderRow = findOrderRow(orderId);
  if (!orderRow)
    return response.setContent(JSON.stringify({ status: 'not_found' }));

  var sheet   = SpreadsheetApp.getActive().getSheetByName('ORDERS');
  var rowData = sheet.getRange(orderRow, 1, 1, 44).getValues()[0];
  var COL     = getOrderColumnMap();

  // Managers can only view their own orders
  if (session.role === 'Manager') {
    var createdBy = rowData[COL['Created_By'] - 1];
    if (createdBy !== session.email) {
      logSecurityEvent('PERMISSION_DENIED', session.email, 'getOrder_other_manager', session.role);
      return response.setContent(JSON.stringify({
        status: 'error', code: 'FORBIDDEN',
        message: 'Bu sifariş sizin hesabınıza aid deyil.'
      }));
    }
  }

  return response.setContent(JSON.stringify({
    status:                   'ok',
    order_id:                 rowData[COL['Order_ID']                 - 1],
    client_id:                rowData[COL['Client_ID']                - 1],
    client_name:              rowData[COL['Client_Name']              - 1],
    order_status:             rowData[COL['Order_Status']             - 1],
    payment_status:           rowData[COL['Payment_Status']           - 1],
    discount_approval_status: rowData[COL['Discount_Approval_Status'] - 1],
    grand_total:              rowData[COL['Grand_Total']              - 1],
    total_paid:               rowData[COL['Total_Paid']               - 1],
    debt_amount:              rowData[COL['Debt_Amount']              - 1],
    price_type:               rowData[COL['Price_Type']               - 1],
    payment_terms:            rowData[COL['Payment_Terms']            - 1],
    notes:                    rowData[COL['Notes']                    - 1],
    delivery_address:         rowData[COL['Delivery_Address']         - 1],
    mode:                     getModeFromStatus(rowData[COL['Order_Status'] - 1]),
    created_by:               rowData[COL['Created_By']              - 1]
  }));
}

function serveDashboard(session, response) {
  var dash = SpreadsheetApp.getActive().getSheetByName('DASHBOARD');
  if (!dash) return response.setContent(JSON.stringify({ status: 'error', code: 'NO_DASHBOARD' }));
  return response.setContent(JSON.stringify({
    status:           'ok',
    new_today:        dash.getRange('B2').getValue(),
    wait_payment:     dash.getRange('B3').getValue(),
    paid_awaiting:    dash.getRange('B4').getValue(),
    in_production:    dash.getRange('B5').getValue(),
    overdue_prod:     dash.getRange('B6').getValue(),
    delivered_mtd:    dash.getRange('B7').getValue(),
    closed_mtd:       dash.getRange('B8').getValue(),
    revenue_mtd:      dash.getRange('B10').getValue(),
    paid_mtd:         dash.getRange('B11').getValue(),
    total_debt:       dash.getRange('B12').getValue(),
    discounts_mtd:    dash.getRange('B13').getValue(),
    problem_orders:   dash.getRange('B14').getValue(),
    pending_discounts: dash.getRange('B15').getValue()
  }));
}

function getModeFromStatus(orderStatus) {
  var map = {
    'Yeni':             'new_request',
    'Təsdiqləndi':      'price_confirmed',
    'Avans Alındı':     'avans_alindi',
    'İstehsalda':       'istehsalda',
    'Hazırdır':         'hazirdir',
    'Çatdırıldı':       'catdirildi',
    'Bağlandı':         'baglandı',
    'Ləğv edildi':      'legv_edildi',
    'Problem/Mübahisə': 'problem'
  };
  return map[orderStatus] || 'new_request';
}

// ─────────────────────────────────────────────────────────────────────────────
// FUNCTION 7: enforceServerLock
// Blocks price/quantity edits after order reaches confirmed statuses.
// ─────────────────────────────────────────────────────────────────────────────

function enforceServerLock(orderId, payload, userEmail) {
  var LOCKED_AFTER_CONFIRMED = [
    'client_name', 'client_id', 'client_type', 'client_phone', 'client_voen',
    'price_type', 'subtotal', 'net_amount', 'vat_pct', 'vat_amount',
    'delivery_fee', 'grand_total', 'discount_pct', 'discount_amount', 'line_items'
  ];

  var currentStatus = getOrderField(orderId, 'Order_Status');
  var lockedStatuses = [
    'Təsdiqləndi', 'Avans Alındı', 'İstehsalda', 'Hazırdır', 'Çatdırıldı',
    // legacy names from spec
    'Qiymət Təsdiqləndi', 'İnvoys Hazırdır', 'İnvoys Göndərildi'
  ];

  if (lockedStatuses.indexOf(currentStatus) === -1) return { blocked: false };

  var violations = [];
  LOCKED_AFTER_CONFIRMED.forEach(function (field) {
    if (payload.hasOwnProperty(field)) {
      violations.push({ field: field, code: 'FIELD_LOCKED_AFTER_CONFIRMATION' });
    }
  });

  return violations.length > 0 ? { blocked: true, violations: violations } : { blocked: false };
}

// ─────────────────────────────────────────────────────────────────────────────
// FUNCTION 8: getDiscountApprovalStatus
// ─────────────────────────────────────────────────────────────────────────────

function getDiscountApprovalStatus(priceType, discountPct) {
  if (!discountPct || discountPct <= 0) return 'N/A';
  // CRITICAL: Topdan + ANY discount → Director escalation
  if (priceType === 'Topdan') return 'PENDING_DIRECTOR';
  var max = parseFloat(getSettingValue('MAX_DISCOUNT_NO_APPROVAL') || 10);
  if (discountPct > max) return 'PENDING_DIRECTOR';
  if (discountPct > 5)   return 'PENDING_SR_MGR';
  return 'N/A'; // 0–5% on Perakende: no approval needed
}

// ─────────────────────────────────────────────────────────────────────────────
// FUNCTION 9: onEditTrigger
// Install as: Spreadsheet → On edit trigger.
// Watches ORDERS and PAYMENTS for manual status changes and validates them.
// ─────────────────────────────────────────────────────────────────────────────

function onEditTrigger(e) {
  if (!e || !e.range) return;
  var sheet     = e.range.getSheet();
  var sheetName = sheet.getName();
  if (sheetName !== 'ORDERS' && sheetName !== 'PAYMENTS') return;

  var col = e.range.getColumn();
  var row = e.range.getRow();
  if (row < 2) return; // header row

  var orders  = SpreadsheetApp.getActive().getSheetByName('ORDERS');
  var COL_MAP = getOrderColumnMap();

  // ── ORDERS sheet events ───────────────────────────────────────────────────
  if (sheetName === 'ORDERS') {

    // N (col 14) = Payment_Status
    if (col === COL_MAP['Payment_Status']) {
      var newPayStatus = e.range.getValue();
      var orderId      = orders.getRange(row, 1).getValue();

      if (newPayStatus === 'CONFIRMED') {
        // Verify at least one confirmed PAYMENTS row exists
        var payments = getPaymentsForOrder(orderId);
        if (payments.length === 0) {
          e.range.setValue(e.oldValue || 'Ödənilmədi'); // revert
          sendTelegram('N9_UNMATCHED_PAYMENT',
            '⚠️ <b>' + orderId + '</b> — Ödəniş qeydi olmadan CONFIRMED qoyuldu. Status geri qaytarıldı.'
          );
          return;
        }
        checkAndUnlockProduction(orderId, row);
      }
    }

    // O (col 15) = Production_Status
    if (col === COL_MAP['Production_Status']) {
      var orderId2   = orders.getRange(row, 1).getValue();
      var newProdSt  = e.range.getValue();
      if (newProdSt === 'Hazırdır') {
        createDeliveryRow(orderId2, row);
        var clientName = orders.getRange(row, COL_MAP['Client_Name']).getValue();
        sendTelegram('N4_ORDER_READY',
          '✅ Hazırdır: <b>' + orderId2 + '</b>\n' +
          '👤 ' + clientName + '\nÇatdırılma planlanmalıdır.'
        );
      }
      if (newProdSt === 'İstehsalda') {
        orders.getRange(row, COL_MAP['Order_Status']).setValue('İstehsalda');
      }
    }

    // P (col 16) = Delivery_Status
    if (col === COL_MAP['Delivery_Status']) {
      var orderId3 = orders.getRange(row, 1).getValue();
      var delVal   = e.range.getValue();
      if (delVal === 'Qaytarıldı') {
        orders.getRange(row, COL_MAP['Order_Status']).setValue('Problem/Mübahisə');
        sendTelegram('N5_RETURN',
          '🚨 QAYTARILDI: <b>' + orderId3 + '</b>\nDərhal araşdırılmalıdır.'
        );
      }
      if (delVal === 'Çatdırıldı') {
        orders.getRange(row, COL_MAP['Order_Status']).setValue('Çatdırıldı');
        orders.getRange(row, COL_MAP['Actual_Delivery_Date']).setValue(new Date());
      }
    }
  }

  // ── PAYMENTS sheet events ─────────────────────────────────────────────────
  if (sheetName === 'PAYMENTS') {
    // D (col 4) = Payment_Date — must be a real Date, not text
    if (col === 4) {
      var dateVal = e.range.getValue();
      if (dateVal && !(dateVal instanceof Date)) {
        e.range.setValue(''); // clear invalid entry
        batchAuditLog(Utilities.getUuid(), [{
          entity_type: 'PAYMENT', entity_id: 'NEW',
          action:      'VALIDATION_ERROR',
          field_name:  'Payment_Date',
          old_value:   '',
          new_value:   String(dateVal),
          user_email:  Session.getActiveUser().getEmail() || 'manual_edit',
          user_role:   '',
          note:        'Invalid date format — cell cleared'
        }]);
      }
    }
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// FUNCTION 10: checkAndUnlockProduction
// ─────────────────────────────────────────────────────────────────────────────

function checkAndUnlockProduction(orderId, orderRow) {
  var orders  = SpreadsheetApp.getActive().getSheetByName('ORDERS');
  var COL_MAP = getOrderColumnMap();

  var productionLock = orders.getRange(orderRow, COL_MAP['Production_Lock']).getValue();
  // Production_Lock formula: FALSE = unlocked (payment confirmed, avans met)
  if (productionLock === false || productionLock === 'FALSE' || productionLock === 'false') {
    createProductionRow(orderId, orderRow);
    var clientName = orders.getRange(orderRow, COL_MAP['Client_Name']).getValue();
    sendTelegram('N3_PRODUCTION_UNLOCK',
      '🏭 İstehsal açıldı: <b>' + orderId + '</b>\n' +
      '👤 ' + clientName + '\n📦 Materiallar: hazırlanır'
    );
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// FUNCTION 11: dailyMorningReport — Time trigger: 09:00 daily
// ─────────────────────────────────────────────────────────────────────────────

function dailyMorningReport() {
  var dash = SpreadsheetApp.getActive().getSheetByName('DASHBOARD');
  if (!dash) return;

  var newToday      = Number(dash.getRange('B2').getValue())  || 0;
  var waitPayment   = Number(dash.getRange('B3').getValue())  || 0;
  var inProduction  = Number(dash.getRange('B5').getValue())  || 0;
  var overdueProd   = Number(dash.getRange('B6').getValue())  || 0;
  var revenueMTD    = Number(dash.getRange('B10').getValue()) || 0;
  var paidMTD       = Number(dash.getRange('B11').getValue()) || 0;
  var totalDebt     = Number(dash.getRange('B12').getValue()) || 0;
  var problemOrders = Number(dash.getRange('B14').getValue()) || 0;

  var today = Utilities.formatDate(new Date(), 'Asia/Baku', 'dd.MM.yyyy');
  var msg   = '📊 <b>MFN — ' + today + '</b>\n\n';
  msg += '💰 <b>Maliyyə:</b>\n';
  msg += '  Ay üzrə gəlir: ' + revenueMTD.toFixed(2)  + ' AZN\n';
  msg += '  Toplanan: '      + paidMTD.toFixed(2)      + ' AZN\n';
  msg += '  Ümumi borc: '    + totalDebt.toFixed(2)    + ' AZN\n\n';
  msg += '⚡ <b>Risklər:</b>\n';
  msg += '  Ödəniş gözlənilir: ' + waitPayment + ' sifariş\n';
  if (problemOrders > 0) msg += '  🔴 Problem sifarişlər: ' + problemOrders + '\n';
  msg += '\n🏭 <b>İstehsal:</b>\n';
  msg += '  İşdə: ' + inProduction;
  if (overdueProd > 0) msg += '  |  🔴 Gecikmiş: ' + overdueProd;
  if (newToday > 0)    msg += '\n\n📋 Bu gün yeni: ' + newToday + ' sifariş';
  if (problemOrders === 0 && overdueProd === 0) msg += '\n\n✅ Kritik problem yoxdur.';

  sendTelegram('DAILY_REPORT', msg);
}

// ─────────────────────────────────────────────────────────────────────────────
// FUNCTION 12: dailyOverdueCheck — Time trigger: 09:15 daily
// ─────────────────────────────────────────────────────────────────────────────

function dailyOverdueCheck() {
  var orders  = SpreadsheetApp.getActive().getSheetByName('ORDERS');
  var COL_MAP = getOrderColumnMap();
  var data    = orders.getDataRange().getValues();
  var now     = new Date();

  var overduePayments   = [];
  var overdueProduction = [];

  for (var i = 1; i < data.length; i++) {
    var row     = data[i];
    var orderId = row[0];
    if (!orderId) continue;
    var status = row[COL_MAP['Order_Status'] - 1];
    if (status === 'Bağlandı' || status === 'Ləğv edildi') continue;

    // Payment overdue
    var debt    = row[COL_MAP['Debt_Amount']      - 1];
    var dueDate = row[COL_MAP['Payment_Due_Date'] - 1];
    if (debt > 0 && dueDate instanceof Date && dueDate < now) {
      var days = Math.floor((now - dueDate) / 86400000);
      overduePayments.push(orderId + ': ' + Number(debt).toFixed(0) + ' AZN (' + days + ' gün)');
    }
  }

  // Production overdue (reads PRODUCTION!J = Overdue_Prod_Flag formula result)
  var prodSheet = SpreadsheetApp.getActive().getSheetByName('PRODUCTION');
  var prodData  = prodSheet.getDataRange().getValues();
  for (var j = 1; j < prodData.length; j++) {
    if (prodData[j][9] === true) { // col J (index 9) = Overdue_Prod_Flag
      var planned = prodData[j][7]; // col H (index 7) = Planned_Ready
      var daysLate = planned instanceof Date
        ? Math.floor((now - planned) / 86400000)
        : '?';
      overdueProduction.push(prodData[j][1] + ': ' + daysLate + ' gün gecikmiş');
    }
  }

  if (overduePayments.length > 0) {
    sendTelegram('N6_OVERDUE_PAYMENT',
      '💸 <b>Gecikmiş ödənişlər (' + overduePayments.length + '):</b>\n' +
      overduePayments.slice(0, 10).join('\n')
    );
  }
  if (overdueProduction.length > 0) {
    sendTelegram('N7_OVERDUE_PRODUCTION',
      '🏭 <b>İstehsal gecikməsi (' + overdueProduction.length + '):</b>\n' +
      overdueProduction.slice(0, 10).join('\n')
    );
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// FUNCTION 13: weeklyBackup — Time trigger: Sunday 23:00
// ─────────────────────────────────────────────────────────────────────────────

function weeklyBackup() {
  var ss      = SpreadsheetApp.getActive();
  var dateStr = Utilities.formatDate(new Date(), 'Asia/Baku', 'yyyy-MM-dd');

  // Find or create MFN-Backups folder
  var folder;
  var folders = DriveApp.getFoldersByName('MFN-Backups');
  folder = folders.hasNext() ? folders.next() : DriveApp.createFolder('MFN-Backups');
  var subFolder = folder.createFolder('MFN-Backup-' + dateStr);

  var SHEETS_TO_BACKUP = ['ORDERS', 'CLIENTS', 'PAYMENTS', 'DISCOUNT_LOG', 'AUDIT_LOG'];
  SHEETS_TO_BACKUP.forEach(function (name) {
    var sheet = ss.getSheetByName(name);
    if (!sheet || sheet.getLastRow() < 1) return;
    var data = sheet.getDataRange().getValues();
    var csv  = data.map(function (r) {
      return r.map(function (c) {
        return '"' + String(c).replace(/"/g, '""') + '"';
      }).join(',');
    }).join('\n');
    subFolder.createFile(name + '_' + dateStr + '.csv', csv, MimeType.CSV);
  });

  // Prune backups older than 90 days
  var allFolders = folder.getFolders();
  var cutoff     = new Date(Date.now() - 90 * 24 * 3600 * 1000);
  while (allFolders.hasNext()) {
    var f = allFolders.next();
    try {
      if (f.getDateCreated() < cutoff) f.setTrashed(true);
    } catch (err) { /* skip locked folders */ }
  }

  sendTelegram('DAILY_REPORT',
    '✅ Həftəlik yedek tamamlandı: ' + dateStr + '\n5 cədvəl Google Drive-a ixrac edildi.'
  );
}

// ─────────────────────────────────────────────────────────────────────────────
// FUNCTION 14: sendTelegram
// ─────────────────────────────────────────────────────────────────────────────

function sendTelegram(type, message) {
  var token = getSettingValue('TELEGRAM_BOT_TOKEN');
  if (!token || token.trim() === '') {
    Logger.log('[Telegram] Not configured — skipped: ' + type);
    return;
  }

  var route     = NOTIFICATION_ROUTING[type] || 'OPS';
  var chatIdKey = route === 'CRITICAL' ? 'TELEGRAM_CRITICAL_CHAT_ID' : 'TELEGRAM_OPS_CHAT_ID';
  var chatId    = getSettingValue(chatIdKey);
  if (!chatId || chatId.trim() === '') {
    Logger.log('[Telegram] Chat ID missing for route: ' + route);
    return;
  }

  try {
    UrlFetchApp.fetch(
      'https://api.telegram.org/bot' + token + '/sendMessage',
      {
        method:            'post',
        contentType:       'application/json',
        payload:           JSON.stringify({ chat_id: chatId, text: message, parse_mode: 'HTML' }),
        muteHttpExceptions: true
      }
    );
  } catch (err) {
    Logger.log('[Telegram] Send failed: ' + err.toString());
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// FUNCTION 15: batchAuditLog
// Appends 1–N rows to AUDIT_LOG in a single setValues call (efficient).
// Filters to LOGGABLE_ACTIONS only — passive events never written.
// ─────────────────────────────────────────────────────────────────────────────

function batchAuditLog(sessionId, changes) {
  var logSheet = SpreadsheetApp.getActive().getSheetByName('AUDIT_LOG');
  if (!logSheet) return;

  var rows = changes
    .filter(function (c) { return LOGGABLE_ACTIONS.indexOf(c.action) !== -1; })
    .map(function (c, i) {
      var ts = Date.now();
      return [
        'LOG-' + ts + '-' + String(i).padStart(3, '0'),     // A Log_ID
        new Date(),                                           // B Timestamp
        c.user_email  || '',                                  // C User_Email
        c.user_role   || '',                                  // D User_Role
        c.entity_type || 'ORDER',                             // E Entity_Type
        c.entity_id   || '',                                  // F Entity_ID
        c.action,                                             // G Action
        c.field_name  || '',                                  // H Field_Name
        c.old_value !== undefined && c.old_value !== null     // I Old_Value
          ? String(c.old_value) : '',
        c.new_value !== undefined && c.new_value !== null     // J New_Value
          ? String(c.new_value) : '',
        sessionId || '',                                      // K Session_ID
        c.note || '',                                         // L Note
        ''                                                    // M (reserved)
      ];
    });

  if (rows.length === 0) return;
  logSheet.getRange(logSheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
}

// ─────────────────────────────────────────────────────────────────────────────
// FUNCTION 16: Helper Functions
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Returns a setting value.
 * Secret keys → Script Properties (not visible in Sheets).
 * Operational settings → SETTINGS sheet named ranges.
 */
function getSettingValue(name) {
  var PROP_KEYS = [
    'TELEGRAM_BOT_TOKEN', 'APP_SECRET_KEY',
    'TELEGRAM_CRITICAL_CHAT_ID', 'TELEGRAM_OPS_CHAT_ID'
  ];
  if (PROP_KEYS.indexOf(name) !== -1) {
    return PropertiesService.getScriptProperties().getProperty(name) || '';
  }
  try {
    var range = SpreadsheetApp.getActive().getRangeByName(name);
    return range ? range.getValue() : '';
  } catch (e) { return ''; }
}

/** Returns role string for an email from MANAGERS tab. */
function getUserRole(email) {
  if (!email) return 'Unknown';
  var sheet = SpreadsheetApp.getActive().getSheetByName('MANAGERS');
  if (!sheet || sheet.getLastRow() < 2) return 'Unknown';
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][2]).trim().toLowerCase() === email.toLowerCase())
      return data[i][3]; // col D = Role
  }
  return 'Unknown';
}

/** Returns value of a single ORDERS field for a given orderId. */
function getOrderField(orderId, fieldName) {
  var sheet  = SpreadsheetApp.getActive().getSheetByName('ORDERS');
  var colMap = getOrderColumnMap();
  var col    = colMap[fieldName];
  if (!col) return null;
  var orderRow = findOrderRow(orderId);
  if (!orderRow) return null;
  return sheet.getRange(orderRow, col).getValue();
}

/** Returns 1-based row number for an orderId in ORDERS, or null. */
function findOrderRow(orderId) {
  var sheet   = SpreadsheetApp.getActive().getSheetByName('ORDERS');
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;
  var ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
  var idx = ids.indexOf(orderId);
  return idx === -1 ? null : idx + 2;
}

/** Returns array of confirmed PAYMENTS rows for a given orderId. */
function getPaymentsForOrder(orderId) {
  var sheet = SpreadsheetApp.getActive().getSheetByName('PAYMENTS');
  if (!sheet || sheet.getLastRow() < 2) return [];
  var data = sheet.getDataRange().getValues();
  return data.slice(1).filter(function (row) {
    return row[1] === orderId && row[7] !== ''; // col B=Order_ID, col H=Confirmed_By
  });
}

/**
 * Returns object mapping ORDERS field names → 1-based column index.
 * 44 columns, A through AR.
 * DO NOT reorder columns in ORDERS — update this map if columns are ever added.
 */
function getOrderColumnMap() {
  return {
    'Order_ID':                 1,  // A
    'Request_ID':               2,  // B
    'Created_At':               3,  // C
    'Created_By':               4,  // D
    'Updated_At':               5,  // E
    'Updated_By':               6,  // F
    'Edit_Count':               7,  // G
    'Invoice_Link':             8,  // H
    'Client_ID':                9,  // I
    'Client_Name':              10, // J (FORMULA)
    'Client_Type':              11, // K (FORMULA)
    'Price_Type':               12, // L
    'Order_Status':             13, // M
    'Payment_Status':           14, // N
    'Production_Status':        15, // O
    'Delivery_Status':          16, // P
    'Discount_Approval_Status': 17, // Q
    'Payment_Terms':            18, // R
    'Subtotal_AZN':             19, // S
    'Discount_Pct':             20, // T
    'Discount_Amount':          21, // U (FORMULA)
    'Net_Amount':               22, // V (FORMULA)
    'VAT_Pct':                  23, // W (FORMULA)
    'VAT_Amount':               24, // X (FORMULA)
    'Delivery_Zone':            25, // Y
    'Delivery_Fee':             26, // Z
    'Grand_Total':              27, // AA (FORMULA)
    'Price_Snapshot_JSON':      28, // AB
    'Avans_Required':           29, // AC (FORMULA)
    'Total_Paid':               30, // AD (FORMULA)
    'Debt_Amount':              31, // AE (FORMULA)
    'Payment_Due_Date':         32, // AF
    'Receipt_Link':             33, // AG
    'Production_Lock':          34, // AH (FORMULA)
    'Problem_Flag':             35, // AI (FORMULA)
    'Overdue_Flag':             36, // AJ (FORMULA)
    'Price_Below_Min_Flag':     37, // AK
    'Manager_ID':               38, // AL (FORMULA)
    'Notes':                    39, // AM
    'Delivery_Address':         40, // AN
    'Delivery_Contact':         41, // AO
    'Delivery_Phone':           42, // AP
    'Planned_Delivery_Date':    43, // AQ
    'Actual_Delivery_Date':     44  // AR
  };
}

/** Appends a row to PRODUCTION and writes the Overdue_Prod_Flag formula. */
function createProductionRow(orderId, orderRow) {
  var orders = SpreadsheetApp.getActive().getSheetByName('ORDERS');
  var prod   = SpreadsheetApp.getActive().getSheetByName('PRODUCTION');
  var COL    = getOrderColumnMap();
  var prodId = 'PROD-' + String(prod.getLastRow()).padStart(5, '0');

  prod.appendRow([
    prodId,                                                           // A Prod_ID
    orderId,                                                          // B Order_ID
    orders.getRange(orderRow, COL['Client_Name']).getValue(),         // C Client_Name
    orders.getRange(orderRow, COL['Price_Snapshot_JSON']).getValue(), // D Items_JSON
    'Normal',                                                         // E Priority
    'Sıraya alındı',                                                  // F Prod_Status
    '',                                                               // G Planned_Start
    '',                                                               // H Planned_Ready
    '',                                                               // I Actual_Ready
    false,                                                            // J Overdue_Prod_Flag ← FORMULA below
    ''                                                                // K Prod_Notes
  ]);

  // Write Overdue_Prod_Flag formula for the appended row
  var newRow = prod.getLastRow();
  prod.getRange(newRow, 10).setFormula(
    '=IF((F' + newRow + '<>"Hazırdır")*(F' + newRow + '<>"Ləğv")*' +
    'ISNUMBER(H' + newRow + ')*(H' + newRow + '+2<TODAY()),TRUE,FALSE)'
  );
}

/** Appends a row to DELIVERY. Skips and alerts if address is empty. */
function createDeliveryRow(orderId, orderRow) {
  var orders = SpreadsheetApp.getActive().getSheetByName('ORDERS');
  var del    = SpreadsheetApp.getActive().getSheetByName('DELIVERY');
  var COL    = getOrderColumnMap();

  var addr = orders.getRange(orderRow, COL['Delivery_Address']).getValue();
  if (!addr || String(addr).trim() === '') {
    sendTelegram('N8_PROBLEM_ESCALATION',
      '⚠️ ' + orderId + ': Çatdırılma ünvanı boşdur. Menecerlə əlaqə saxlayın.'
    );
    return; // Never create a DELIVERY row without an address
  }

  var delId = 'DEL-' + String(del.getLastRow()).padStart(5, '0');
  del.appendRow([
    delId,                                                                    // A Delivery_ID
    orderId,                                                                  // B Order_ID
    addr,                                                                     // C Address
    orders.getRange(orderRow, COL['Delivery_Contact']).getValue(),            // D Contact
    orders.getRange(orderRow, COL['Delivery_Phone']).getValue(),              // E Phone
    orders.getRange(orderRow, COL['Delivery_Zone']).getValue(),               // F Zone
    'Planlandı',                                                              // G Delivery_Status
    orders.getRange(orderRow, COL['Planned_Delivery_Date']).getValue() || '', // H Scheduled_Date
    '',                                                                       // I Actual_Date
    '',                                                                       // J Driver_Note
    ''                                                                        // K Return_Reason
  ]);
}

/** Appends a row to DISCOUNT_LOG. */
function createDiscountLogRow(orderId, payload, userEmail, status) {
  var discLog  = SpreadsheetApp.getActive().getSheetByName('DISCOUNT_LOG');
  var discId   = 'DISC-' + String(discLog.getLastRow()).padStart(5, '0');
  var subtotal = parseFloat(payload.subtotal    || 0);
  var discPct  = parseFloat(payload.discount_pct || 0);

  discLog.appendRow([
    discId,                        // A Approval_ID
    orderId,                       // B Order_ID
    payload.client_name || '',     // C Client_Name
    userEmail,                     // D Requested_By
    new Date(),                    // E Requested_At
    subtotal,                      // F Order_Subtotal
    discPct,                       // G Requested_Discount_Pct
    subtotal * discPct / 100,      // H Discount_Amount
    payload.discount_reason || '', // I Manager_Justification
    status,                        // J Approval_Status
    '',                            // K Approved_By
    '',                            // L Approved_At
    '',                            // M Rejection_Reason
    discPct                        // N Final_Approved_Pct (initially = requested)
  ]);
}

/** Sends Telegram + email alert for discount requiring approval. */
function sendDiscountAlert(orderId, payload, userEmail) {
  var approvalEmail = String(getSettingValue('APPROVAL_EMAIL') || '');
  var subtotal = parseFloat(payload.subtotal    || 0);
  var discPct  = parseFloat(payload.discount_pct || 0);
  var lostAzn  = (subtotal * discPct / 100).toFixed(2);

  sendTelegram('N2_DISCOUNT_ALERT',
    '🔒 Endirim təsdiqi: <b>' + orderId + '</b>\n' +
    '👤 ' + payload.client_name + '\n' +
    '💸 ' + discPct + '% endirim (' + lostAzn + ' AZN itkisi)\n' +
    '📝 Səbəb: ' + (payload.discount_reason || 'yoxdur') + '\n' +
    '👨‍💼 Müraciət: ' + userEmail
  );

  if (approvalEmail && approvalEmail.indexOf('@') !== -1) {
    try {
      GmailApp.sendEmail(
        approvalEmail,
        'MFN: Endirim Təsdiqi Tələb Olunur — ' + orderId,
        'Sifariş ' + orderId + ' üçün ' + discPct + '% endirim sorğusu.\n' +
        'Müştəri: ' + payload.client_name + '\nİtki: ' + lostAzn + ' AZN\n' +
        'Səbəb: ' + (payload.discount_reason || '—') + '\nMenecer: ' + userEmail
      );
    } catch (err) {
      Logger.log('[Email] send failed: ' + err.toString());
    }
  }
}

/** Keeps Apps Script warm — prevents 8–12 second cold start on first doPost. */
function warmupPing() {
  SpreadsheetApp.getActive().getName();
}

// ─────────────────────────────────────────────────────────────────────────────
// ONE-TIME SETUP: setupScriptProperties
//
// Run ONCE manually from Apps Script editor after first deployment.
// Then DELETE this function (or comment it out) to prevent accidental reruns.
// Secrets stored here are NOT visible in any Sheet cell.
// ─────────────────────────────────────────────────────────────────────────────

function setupScriptProperties() {
  var props = PropertiesService.getScriptProperties();
  props.setProperties({
    'TELEGRAM_BOT_TOKEN':         'YOUR_BOT_TOKEN_HERE',         // from @BotFather
    'APP_SECRET_KEY':             Utilities.getUuid(),           // auto-generated
    'TELEGRAM_CRITICAL_CHAT_ID':  'YOUR_CRITICAL_CHAT_ID_HERE', // MFN-Kritik group
    'TELEGRAM_OPS_CHAT_ID':       'YOUR_OPS_CHAT_ID_HERE'       // MFN-Əməliyyat group
  });
  Logger.log('✅ Script properties set. DELETE this function now.');
}
