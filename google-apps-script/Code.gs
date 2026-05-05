/**
 * SCHULHELFER – Primarstufe Rittergasse Basel
 * Google Apps Script Backend
 * 
 * WICHTIG: Nach jeder Änderung:
 * 1. Speichern (Ctrl+S)
 * 2. Bereitstellen → Neue Bereitstellung → Web-App
 * 3. Zugriff: "Jeder" (nicht "Jeder mit Google-Konto")
 */

// === Configuration ===
var RATE_LIMIT_WINDOW = 60; // seconds
var RATE_LIMIT_MAX_REQUESTS = 10; // default cap per window
// Per-action caps. GET getEvents is the high-traffic public read, so give
// it a generous budget. POST and admin endpoints stay tight.
var RATE_LIMIT_CAPS = {
  'getEvents': 120,
  'export': 20,
  'getHelferList': 30,
  'POST': 10
};
// Shared bucket for failed admin auth attempts (brute-force mitigation,
// independent of the per-identifier GET limit).
var ADMIN_BRUTEFORCE_CAP = 15;

// Input length caps (characters). Prevents abuse and keeps sheet clean.
var MAX_NAME_LEN = 100;
var MAX_TEL_LEN = 30;
var MAX_DESC_LEN = 500;
var MAX_EVENT_NAME_LEN = 120;

var ADMIN_EMAIL = ''; // Set this to receive notifications (optional)

// Admin key for the Helferliste download. Reads from Script Properties
// first (recommended), falls back to this constant. To set up:
//   1. In the Apps Script editor → Project Settings → Script Properties
//   2. Add property: Key = ADMIN_KEY, Value = <long random string>
//   3. Redeploy. The constant below can then stay empty.
var ADMIN_KEY_FALLBACK = '';

function getAdminKeyValue() {
  try {
    var stored = PropertiesService.getScriptProperties().getProperty('ADMIN_KEY');
    if (stored) return stored;
  } catch (e) {
    Logger.log('Could not read Script Properties: ' + e);
  }
  return ADMIN_KEY_FALLBACK;
}

// === Utility Functions ===

/**
 * Sanitize input to prevent XSS
 */
function sanitizeInput(str) {
  if (!str) return '';
  return String(str).trim();
}

/**
 * Wraps a value for safe writing to a spreadsheet cell. Google Sheets
 * (and Excel on CSV import) treat cells starting with =, +, -, @, \t
 * or \r as formulas. This historically corrupted Swiss phone numbers
 * like "+41 79 123 45 67" — Sheets evaluates "+41..." as a unary-plus
 * formula and stores the number (losing the "+") or, when the
 * expression is invalid, stores "#ERROR!".
 *
 * For string values matching that pattern we prepend a single quote,
 * which Sheets consumes as a text-literal indicator: the cell stores
 * and displays the value unchanged, but no longer tries to parse it.
 * Non-strings (numbers, Dates, booleans, null) pass through unchanged
 * so numeric and date columns keep their types.
 *
 * Reference: OWASP "CSV Injection" (formula-injection) mitigation.
 */
function sheetSafe(value) {
  if (value === null || value === undefined) return value;
  if (typeof value !== 'string') return value;
  if (value === '') return value;
  if (/^[=+\-@\t\r]/.test(value)) return "'" + value;
  return value;
}

/**
 * Normalize a Swiss phone number to canonical international format
 * "+41 XX XXX XX XX". Parents type phone numbers inconsistently – with
 * or without the country code, with various separators – and without
 * normalization the spreadsheet ends up with every variant of the same
 * number. This function harmonises all the common Swiss-formatted
 * inputs to one stored-and-printed form.
 *
 * Handles:
 *   "079 123 45 67"     → "+41 79 123 45 67"
 *   "79 123 45 67"      → "+41 79 123 45 67"   (subscriber digits only)
 *   "61 999 00 00"      → "+41 61 999 00 00"   (subscriber digits only)
 *   "+41 79 123 45 67"  → "+41 79 123 45 67"
 *   "+41791234567"      → "+41 79 123 45 67"
 *   "0041 79 123 45 67" → "+41 79 123 45 67"
 *   "0041791234567"     → "+41 79 123 45 67"
 *   "(079) 123-45-67"   → "+41 79 123 45 67"
 *   "061 999 00 00"     → "+41 61 999 00 00"   (Swiss landline)
 *   "+49 30 12345678"   → "+49 30 12345678"    (German landline)
 *   "0049 30 12345678"  → "+49 30 12345678"
 *   "49 30 12345678"    → "+49 30 12345678"    (subscriber digits only)
 *   "+49 151 1234567"   → "+49 151 1234567"    (German mobile, 3-digit area)
 *   "+33 1 23 45 67 89" → "+33 1 23 45 67 89"  (other foreign: untouched)
 *   ""                  → ""
 *   "Bitte anrufen"     → "Bitte anrufen"      (unparseable: untouched)
 */
function normalizePhone(raw) {
  if (raw === null || raw === undefined) return '';
  var original = String(raw).trim();
  if (!original) return '';

  var hasPlus = original.charAt(0) === '+';
  var digits = original.replace(/\D/g, '');
  if (!digits) return original;

  // 00xx international prefix → +xx
  if (digits.indexOf('00') === 0 && !hasPlus) {
    digits = digits.slice(2);
    hasPlus = true;
  }

  // Swiss national format: leading 0 with 10 digits total (0 + 9)
  if (!hasPlus && digits.charAt(0) === '0' && digits.length === 10) {
    digits = '41' + digits.slice(1);
    hasPlus = true;
  }

  // Swiss subscriber digits without trunk prefix or country code:
  // "79 123 45 67" or "61 999 00 00" — admins / parents commonly drop
  // the leading 0 when typing. 9 digits, first digit non-zero → treat
  // as Swiss subscriber digits and prepend +41.
  if (!hasPlus && digits.length === 9 && digits.charAt(0) !== '0') {
    digits = '41' + digits;
    hasPlus = true;
  }

  // Swiss international format: 41 + 9 subscriber digits = 11 digits
  if (digits.indexOf('41') === 0 && digits.length === 11) {
    return '+41 ' + digits.substr(2, 2) + ' ' + digits.substr(4, 3) +
           ' ' + digits.substr(7, 2) + ' ' + digits.substr(9, 2);
  }

  // German numbers: 49 + 9–11 subscriber digits = 11–13 total. Format
  // as "+49 <area> <rest>" with a 3-digit mobile prefix (1[5-9]…) or a
  // 2-digit landline area (heuristic — actual German area codes vary
  // 2–5 digits but we don't have a database; 2 is the safe fallback).
  if (digits.indexOf('49') === 0 && digits.length >= 11 && digits.length <= 13) {
    var deRest = digits.substring(2);
    var deAreaLen = /^1[5-9]/.test(deRest) ? 3 : 2;
    return '+49 ' + deRest.substring(0, deAreaLen) + ' ' + deRest.substring(deAreaLen);
  }

  // Other international (e.g. "+33...", "+39..."): keep as originally
  // typed (trimmed). We don't know the local grouping convention, so
  // losing the spaces would be worse than leaving them alone.
  if (hasPlus) return original;

  // Nothing matched our Swiss heuristics (no leading 0, no +, unusual
  // length). Return the original trimmed input unchanged – we never
  // discard what the parent typed.
  return original;
}

/**
 * Parse a date safely. Handles:
 *   - Date objects (as returned by Sheets for date-formatted cells)
 *   - ISO-ish strings ("2026-04-14", "2026-04-14T12:00:00")
 *   - Swiss "DD.MM.YYYY" strings (the format the README documents for
 *     manually-entered events; JS's built-in Date() cannot parse these)
 *   - DD/MM/YYYY as a tolerance fallback
 */
function parseDate(dateValue) {
  if (!dateValue) return null;
  if (dateValue instanceof Date) {
    return isNaN(dateValue.getTime()) ? null : dateValue;
  }
  var s = String(dateValue).trim();
  if (!s) return null;
  // DD.MM.YYYY or D.M.YYYY
  var m = s.match(/^(\d{1,2})[\.\/](\d{1,2})[\.\/](\d{4})$/);
  if (m) {
    var day = parseInt(m[1], 10);
    var month = parseInt(m[2], 10) - 1;
    var year = parseInt(m[3], 10);
    var d = new Date(year, month, day);
    if (d.getFullYear() === year && d.getMonth() === month && d.getDate() === day) {
      return d;
    }
    return null;
  }
  // ISO / other formats JS handles natively
  var parsed = new Date(s);
  return isNaN(parsed.getTime()) ? null : parsed;
}

/**
 * Check rate limiting. Uses CacheService (auto-expires) so nothing has
 * to be cleaned up manually – the old ScriptProperties-based approach
 * slowly filled up, and the 10-req default globally rate-limited every
 * anonymous GET because the client never sent an identifier.
 *
 * Fails open on cache errors so a transient Google-side issue doesn't
 * lock the whole site out.
 *
 * @param {string} identifier   Logical bucket (action or email).
 * @param {number} [maxRequests]  Optional override of the default cap.
 * @returns {boolean} true when the request is allowed.
 */
function checkRateLimit(identifier, maxRequests) {
  var cap = maxRequests || RATE_LIMIT_MAX_REQUESTS;
  try {
    var cache = CacheService.getScriptCache();
    var key = 'rl:' + identifier;
    var now = Math.floor(Date.now() / 1000);
    var windowStart = now - RATE_LIMIT_WINDOW;
    var data = cache.get(key);
    var requests = data ? JSON.parse(data) : [];
    requests = requests.filter(function(ts) { return ts > windowStart; });
    if (requests.length >= cap) return false;
    requests.push(now);
    // TTL = 2× window so edge-of-window requests aren't forgotten early.
    cache.put(key, JSON.stringify(requests), RATE_LIMIT_WINDOW * 2);
    return true;
  } catch (e) {
    Logger.log('Rate limit check failed (fail-open): ' + e);
    return true;
  }
}

/**
 * Log audit trail
 */
function logAudit(action, data, success, error) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var logSheet = ss.getSheetByName('Audit-Log');
    
    if (!logSheet) {
      logSheet = ss.insertSheet('Audit-Log');
      logSheet.getRange(1, 1, 1, 6).setValues([
        ['Zeitstempel', 'Aktion', 'Daten', 'Erfolg', 'Fehler', 'IP/User']
      ]);
      logSheet.getRange(1, 1, 1, 6).setBackground('#64748b').setFontColor('white').setFontWeight('bold');
      logSheet.setColumnWidths(1, 1, 150);
      logSheet.setColumnWidths(2, 1, 120);
      logSheet.setColumnWidths(3, 1, 300);
      logSheet.setColumnWidths(4, 1, 80);
      logSheet.setColumnWidths(5, 1, 200);
      logSheet.setColumnWidths(6, 1, 150);
    }
    
    var dataStr = data ? JSON.stringify(data).substring(0, 500) : '';
    var errorStr = error ? String(error).substring(0, 200) : '';
    
    logSheet.appendRow([
      new Date(),
      action,
      dataStr,
      success ? 'Ja' : 'Nein',
      sheetSafe(errorStr),
      Session.getActiveUser().getEmail() || 'Web-App'
    ]);
    
    // Keep only last 1000 entries. Only check every ~50 writes to avoid
    // the overhead of counting rows on every single audit call.
    if (Math.random() < 0.02) {
      var lastRow = logSheet.getLastRow();
      if (lastRow > 1001) {
        logSheet.deleteRows(2, lastRow - 1001);
      }
    }
  } catch (e) {
    // Silent fail for logging
    Logger.log('Audit logging failed: ' + e);
  }
}

/**
 * Send email notification
 */
function sendEmailNotification(to, subject, body, htmlBody) {
  if (!to) return;
  try {
    MailApp.sendEmail({
      to: to,
      subject: subject,
      body: body,
      htmlBody: htmlBody || body.replace(/\n/g, '<br>')
    });
  } catch (e) {
    Logger.log('Email notification failed: ' + e);
  }
}

// === GET Requests ===
function doGet(e) {
  var output;
  var action = (e && e.parameter && e.parameter.action) || 'getEvents';
  // Scope the rate-limit bucket by action. The client still doesn't
  // send an identifier for GETs, so all anonymous traffic shares the
  // 'getEvents'/'export'/'getHelferList' buckets – but the per-action
  // caps are generous enough for real use.
  var identifier = action + ':' + (e.parameter.identifier || 'anon');
  var cap = RATE_LIMIT_CAPS[action] || RATE_LIMIT_MAX_REQUESTS;

  try {
    if (!checkRateLimit(identifier, cap)) {
      logAudit('GET_RATE_LIMIT', { action: action }, false, 'Rate limit exceeded');
      output = JSON.stringify({
        success: false,
        error: 'Zu viele Anfragen. Bitte versuchen Sie es später erneut.',
        errorCode: 'RATE_LIMIT'
      });
      return ContentService
        .createTextOutput(output)
        .setMimeType(ContentService.MimeType.JSON);
    }

    if (action === 'getEvents') {
      var events = getAktiveAnlaesse();
      output = JSON.stringify({ 
        success: true,
        events: events 
      });
      logAudit('GET_EVENTS', { count: events.length }, true, null);
    } else if (action === 'export') {
      // Admin-only PII dump. Same gate as getHelferList: validate the
      // ADMIN_KEY (constant-time) BEFORE running the export, and feed
      // failed attempts into the global brute-force bucket.
      if (!isAdminAuthorized(e.parameter.adminKey)) {
        if (!checkRateLimit('admin-fail', ADMIN_BRUTEFORCE_CAP)) {
          logAudit('EXPORT_AUTH_BRUTEFORCE', {}, false, 'Lockout');
          output = JSON.stringify({ success: false, error: 'Zu viele Fehlversuche. Bitte später erneut versuchen.' });
        } else {
          logAudit('EXPORT_AUTH_FAIL', {}, false, 'Invalid admin key');
          output = JSON.stringify({ success: false, error: 'Keine Berechtigung.' });
        }
      } else {
        var result = exportData();
        output = JSON.stringify(result);
        logAudit('EXPORT_DATA', {}, result.success, result.error);
      }
    } else if (action === 'getHelferList') {
      // Admin-only: return full registration data for a single event so
      // the frontend can build a .docx Helferliste. Requires ADMIN_KEY.
      // Validate the key first so failed attempts feed the brute-force
      // bucket BEFORE the expensive sheet read happens — previously a
      // close-to-brute-forced key still got the data back on the
      // success boundary because the bucket only fired post-query.
      if (!isAdminAuthorized(e.parameter.adminKey)) {
        if (!checkRateLimit('admin-fail', ADMIN_BRUTEFORCE_CAP)) {
          logAudit('ADMIN_AUTH_BRUTEFORCE', { eventId: e.parameter.eventId }, false, 'Lockout');
          output = JSON.stringify({ success: false, error: 'Zu viele Fehlversuche. Bitte später erneut versuchen.' });
        } else {
          logAudit('ADMIN_AUTH_FAIL', { eventId: e.parameter.eventId }, false, 'Invalid admin key');
          output = JSON.stringify({ success: false, error: 'Keine Berechtigung.' });
        }
      } else {
        var result = getHelferList(e.parameter.eventId, e.parameter.adminKey);
        output = JSON.stringify(result);
        logAudit('GET_HELFER_LIST', { eventId: e.parameter.eventId }, result.success, result.error);
      }
    } else {
      output = JSON.stringify({ 
        success: false,
        error: 'Unbekannte Aktion' 
      });
      logAudit('GET_UNKNOWN', { action: action }, false, 'Unknown action');
    }
  } catch (error) {
    var errorMsg = 'Ein unerwarteter Fehler ist aufgetreten.';
    logAudit('GET_ERROR', { action: e.parameter.action }, false, error.toString());
    output = JSON.stringify({ 
      success: false,
      error: errorMsg 
    });
  }
  
  return ContentService
    .createTextOutput(output)
    .setMimeType(ContentService.MimeType.JSON);
}

// === POST Requests ===
function doPost(e) {
  var output;
  var identifier = 'anonymous';

  try {
    var data = JSON.parse(e.postData.contents);
    var action = String(data.action || 'register');

    // Admin actions are dispatched separately. They use the admin
    // bucket for rate limiting and require ADMIN_KEY auth (sent in the
    // POST body, never the URL).
    if (action !== 'register') {
      output = JSON.stringify(handleAdminPost(action, data));
      return ContentService
        .createTextOutput(output)
        .setMimeType(ContentService.MimeType.JSON);
    }

    identifier = 'POST:' + (data.email || 'anonymous').toLowerCase();

    // Honeypot: if the hidden field is filled, silently accept to fool bots
    if (data.website) {
      logAudit('HONEYPOT_TRIGGERED', { email: data.email }, false, 'Bot detected');
      output = JSON.stringify({ success: true, message: 'Vielen Dank für Ihre Anmeldung!' });
      return ContentService.createTextOutput(output).setMimeType(ContentService.MimeType.JSON);
    }

    // Rate limiting
    if (!checkRateLimit(identifier, RATE_LIMIT_CAPS.POST)) {
      logAudit('POST_RATE_LIMIT', { anlassId: data.anlassId }, false, 'Rate limit exceeded');
      output = JSON.stringify({
        success: false,
        message: 'Zu viele Anfragen. Bitte versuchen Sie es später erneut.',
        errorCode: 'RATE_LIMIT'
      });
      return ContentService
        .createTextOutput(output)
        .setMimeType(ContentService.MimeType.JSON);
    }

    var result = registriereHelfer(data);
    logAudit('REGISTRATION', { anlassId: data.anlassId, email: data.email }, result.success, result.message);
    output = JSON.stringify(result);
  } catch (error) {
    var errorMsg = 'Ein unerwarteter Fehler ist aufgetreten. Bitte versuchen Sie es später erneut.';
    logAudit('POST_ERROR', {}, false, error.toString());
    output = JSON.stringify({
      success: false,
      message: errorMsg
    });
  }

  return ContentService
    .createTextOutput(output)
    .setMimeType(ContentService.MimeType.JSON);
}

// =============================================================
// === Admin POST dispatch (PR 3) ==============================
// =============================================================
//
// All admin actions live behind a single ADMIN_KEY check. Failed auth
// attempts share the global brute-force bucket (ADMIN_BRUTEFORCE_CAP),
// which is independent of any per-identifier rate limit. Authenticated
// requests share a 'admin-action' bucket of 60/min — generous for a
// real teacher workflow, tight enough to stop runaway scripts.
//
// Every action returns { success, ... } where unsuccessful results
// carry a human-readable `error` (admin UI surfaces this directly).

function handleAdminPost(action, data) {
  if (!isAdminAuthorized(data && data.adminKey)) {
    if (!checkRateLimit('admin-fail', ADMIN_BRUTEFORCE_CAP)) {
      logAudit('ADMIN_AUTH_BRUTEFORCE', { action: action }, false, 'Lockout');
      return { success: false, error: 'Zu viele Fehlversuche. Bitte später erneut.' };
    }
    logAudit('ADMIN_AUTH_FAIL', { action: action }, false, 'Invalid admin key');
    return { success: false, error: 'Keine Berechtigung.' };
  }

  if (!checkRateLimit('admin-action', 60)) {
    return { success: false, error: 'Zu viele Aktionen in kurzer Zeit. Bitte kurz warten.' };
  }

  try {
    switch (action) {
      case 'getAllEvents':        return getAllEventsAdmin();
      case 'addEvent':            return addEventAdmin(data);
      case 'updateEvent':         return updateEventAdmin(data);
      case 'cancelEvent':         return cancelEventAdmin(data);
      case 'getRegistrations':    return getRegistrationsAdmin(data);
      case 'addRegistration':     return addRegistrationAdmin(data);
      case 'updateRegistration':  return updateRegistrationAdmin(data);
      case 'archiveJahr':    return archiveJahrAdmin(data);
      case 'integrityCheck':      return integrityCheckAdmin();
      case 'availableJahre': return availableJahreAdmin();
      default:
        return { success: false, error: 'Unbekannte Aktion: ' + action };
    }
  } catch (e) {
    logAudit('ADMIN_ACTION_ERROR', { action: action }, false, String(e));
    return { success: false, error: 'Serverfehler: ' + e };
  }
}

function isAdminAuthorized(providedKey) {
  var stored = getAdminKeyValue();
  if (!stored) return false;
  if (!providedKey) return false;
  return secureEquals(String(providedKey), String(stored));
}

/**
 * Return every live event (all statuses, including past dates) so the
 * admin UI can group them. Archived rows are intentionally excluded —
 * they live in their own sheets and the admin dashboard treats them
 * as off-stage history.
 */
function getAllEventsAdmin() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Anlässe');
  if (!sheet) return { success: true, events: [] };
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return { success: true, events: [] };

  var values = sheet.getRange(2, 1, lastRow - 1, ANLASS_HEADERS.length).getValues();
  var events = [];
  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    if (!row[0]) continue;
    var datum = parseDate(row[2]);
    events.push({
      id: String(row[0]),
      name: String(row[1] || ''),
      datum: datum ? datum.toISOString() : '',
      datumDisplay: datum ? formatDatum(datum) : '',
      zeit: String(row[3] || ''),
      maxHelfer: parseInt(row[4]) || 0,
      angemeldete: parseInt(row[5]) || 0,
      beschreibung: String(row[6] || ''),
      status: String(row[7] || 'aktiv').toLowerCase() || 'aktiv',
      schuljahr: String(row[8] || ''),
      sichtbar: String(row[9] || ''),
      notizen: String(row[10] || ''),
      kontaktName: String(row[11] || ''),
      kontaktEmail: String(row[12] || '')
    });
  }
  return { success: true, events: events };
}

function addEventAdmin(data) {
  var validation = validateEventInput(data);
  if (!validation.ok) return { success: false, error: validation.error };

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Anlässe');
  if (!sheet) return { success: false, error: 'Tabellenblatt "Anlässe" fehlt.' };

  var lock = LockService.getScriptLock();
  try {
    if (!lock.tryLock(10000)) {
      return { success: false, error: 'Server gerade ausgelastet. Bitte erneut versuchen.' };
    }
    var newId = nextAnlassId();
    sheet.appendRow([
      newId,
      sheetSafe(validation.name),
      validation.datum,
      sheetSafe(validation.zeit),
      validation.helfer,
      0, // formula written below
      sheetSafe(validation.beschreibung),
      'aktiv',
      jahrFor(validation.datum),
      '',
      '',
      sheetSafe(validation.kontaktName),
      sheetSafe(validation.kontaktEmail)
    ]);
    var newRow = sheet.getLastRow();
    sheet.getRange(newRow, 6).setFormula(angemeldeteFormulaForRow(newRow));
    sheet.getRange(newRow, 10).setFormula(sichtbarFormulaForRow(newRow));
    logAudit('ADMIN_ADD_EVENT', { id: newId, name: validation.name }, true, null);
    return { success: true, event: { id: String(newId), name: validation.name } };
  } finally {
    try { if (lock.hasLock()) lock.releaseLock(); } catch (_) {}
  }
}

function updateEventAdmin(data) {
  var id = String(data && data.id || '').trim();
  if (!id) return { success: false, error: 'Anlass-ID fehlt.' };

  var validation = validateEventInput(data);
  if (!validation.ok) return { success: false, error: validation.error };

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Anlässe');
  if (!sheet) return { success: false, error: 'Tabellenblatt "Anlässe" fehlt.' };

  var lock = LockService.getScriptLock();
  try {
    if (!lock.tryLock(10000)) {
      return { success: false, error: 'Server gerade ausgelastet. Bitte erneut versuchen.' };
    }
    var rowIdx = findEventRow(sheet, id);
    if (rowIdx === -1) return { success: false, error: 'Anlass nicht gefunden.' };

    // Update in-place, preserving Status, Notizen, and the formulas in
    // F (Angemeldete) and J (Sichtbar?). Jahr is recomputed because
    // the date may have moved across the Aug-cutoff.
    sheet.getRange(rowIdx, 2).setValue(sheetSafe(validation.name));
    sheet.getRange(rowIdx, 3).setValue(validation.datum);
    sheet.getRange(rowIdx, 4).setValue(sheetSafe(validation.zeit));
    sheet.getRange(rowIdx, 5).setValue(validation.helfer);
    sheet.getRange(rowIdx, 7).setValue(sheetSafe(validation.beschreibung));
    sheet.getRange(rowIdx, 9).setValue(jahrFor(validation.datum));
    sheet.getRange(rowIdx, 12).setValue(sheetSafe(validation.kontaktName));
    sheet.getRange(rowIdx, 13).setValue(sheetSafe(validation.kontaktEmail));
    // Re-assert the formulas in case the row was edited manually and
    // those cells got clobbered with values.
    sheet.getRange(rowIdx, 6).setFormula(angemeldeteFormulaForRow(rowIdx));
    sheet.getRange(rowIdx, 10).setFormula(sichtbarFormulaForRow(rowIdx));

    logAudit('ADMIN_UPDATE_EVENT', { id: id, name: validation.name }, true, null);
    return { success: true, event: { id: id, name: validation.name } };
  } finally {
    try { if (lock.hasLock()) lock.releaseLock(); } catch (_) {}
  }
}

function cancelEventAdmin(data) {
  var id = String(data && data.id || '').trim();
  if (!id) return { success: false, error: 'Anlass-ID fehlt.' };

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Anlässe');
  if (!sheet) return { success: false, error: 'Tabellenblatt "Anlässe" fehlt.' };

  var lock = LockService.getScriptLock();
  try {
    if (!lock.tryLock(10000)) {
      return { success: false, error: 'Server gerade ausgelastet. Bitte erneut versuchen.' };
    }
    var rowIdx = findEventRow(sheet, id);
    if (rowIdx === -1) return { success: false, error: 'Anlass nicht gefunden.' };

    sheet.getRange(rowIdx, 8).setValue('abgesagt');
    logAudit('ADMIN_CANCEL_EVENT', { id: id }, true, null);
    return { success: true };
  } finally {
    try { if (lock.hasLock()) lock.releaseLock(); } catch (_) {}
  }
}

function findEventRow(sheet, id) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;
  var ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0]).trim() === id) return i + 2;
  }
  return -1;
}

/**
 * Centralised input validation for addEvent/updateEvent. The same
 * caps and rules as the in-Sheet dialog (anlassHinzufuegen) so the
 * two paths can never disagree about what's acceptable.
 */
function validateEventInput(data) {
  if (!data) return { ok: false, error: 'Keine Daten übermittelt.' };
  var name = sanitizeInput(data.name || '');
  var zeit = sanitizeInput(data.zeit || '');
  var beschreibung = sanitizeInput(data.beschreibung || '');
  var helfer = parseInt(data.helfer, 10);
  if (!name) return { ok: false, error: 'Bitte einen Namen eingeben.' };
  if (name.length > MAX_EVENT_NAME_LEN) {
    return { ok: false, error: 'Der Anlass-Name ist zu lang (max. ' + MAX_EVENT_NAME_LEN + ' Zeichen).' };
  }
  if (beschreibung.length > MAX_DESC_LEN) {
    return { ok: false, error: 'Die Beschreibung ist zu lang (max. ' + MAX_DESC_LEN + ' Zeichen).' };
  }
  if (!Number.isFinite(helfer) || helfer < 1) {
    return { ok: false, error: 'Bitte mindestens 1 Helfer angeben.' };
  }
  var datum = parseDate(data.datum);
  if (!datum) return { ok: false, error: 'Ungültiges Datum.' };

  // Optional public-contact fields. Both can be blank. If kontaktEmail
  // is non-empty, validate the format; admins who only have a phone
  // person (no email) can still fill kontaktName alone.
  var kontaktName = sanitizeInput(data.kontaktName || '');
  var kontaktEmail = String(data.kontaktEmail || '').trim().toLowerCase();
  if (kontaktName.length > 80) {
    return { ok: false, error: 'Kontaktperson zu lang (max. 80 Zeichen).' };
  }
  if (kontaktEmail) {
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(kontaktEmail) || kontaktEmail.length > 254) {
      return { ok: false, error: 'Ungültige Kontakt-Email.' };
    }
  }

  return {
    ok: true,
    name: name,
    zeit: zeit,
    beschreibung: beschreibung,
    helfer: helfer,
    datum: datum,
    kontaktName: kontaktName,
    kontaktEmail: kontaktEmail
  };
}

// =============================================================
// === Admin v2: helpers + archive + integrity (PR 4) ==========
// =============================================================

/**
 * Return every registration row for one event, including Status and
 * Notizen. The frontend uses this for the per-event helper drawer.
 * Phone numbers are normalised on read so display stays consistent
 * even for rows registered before the write-side normalisation
 * landed.
 */
function getRegistrationsAdmin(data) {
  var eventId = String(data && data.eventId || '').trim();
  if (!eventId) return { success: false, error: 'Anlass-ID fehlt.' };

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var helferSheet = ss.getSheetByName('Anmeldungen');
  if (!helferSheet) return { success: false, error: 'Tabellenblatt "Anmeldungen" fehlt.' };

  var lastRow = helferSheet.getLastRow();
  if (lastRow < 2) return { success: true, registrations: [] };

  var values = helferSheet.getRange(2, 1, lastRow - 1, ANMELDUNG_HEADERS.length).getValues();
  var registrations = [];
  for (var i = 0; i < values.length; i++) {
    if (String(values[i][1] || '') !== eventId) continue;
    var ts = values[i][0];
    registrations.push({
      // (eventId, email) pair is enforced unique by registriereHelfer's
      // duplicate check, so we use email as the natural row key.
      eventId: eventId,
      email: String(values[i][3] || '').toLowerCase().trim(),
      zeitstempel: ts instanceof Date ? ts.toISOString() : String(ts || ''),
      name: sanitizeInput(String(values[i][2] || '')),
      telefon: normalizePhone(sanitizeInput(String(values[i][4] || ''))),
      anlassName: sanitizeInput(String(values[i][5] || '')),
      status: String(values[i][6] || 'aktiv').toLowerCase() || 'aktiv',
      notizen: String(values[i][7] || '')
    });
  }
  // Stable order: registration time ascending.
  registrations.sort(function(a, b) {
    return new Date(a.zeitstempel || 0).getTime() - new Date(b.zeitstempel || 0).getTime();
  });
  return { success: true, registrations: registrations };
}

/**
 * Admin-driven manual registration. Reaches the same write path as
 * the public form (registriereHelfer → LockService → capacity check
 * → COUNTIFS-aware append) but is exempt from the public POST rate
 * limit (we got here via the admin path which has its own quota).
 *
 * The audit-log line tags source='admin-manual' so the school's
 * compliance trail stays unambiguous.
 */
function addRegistrationAdmin(data) {
  if (!data || !data.anlassId || !data.name || !data.email) {
    return { success: false, error: 'Pflichtfelder fehlen.' };
  }
  var result = registriereHelfer({
    anlassId: data.anlassId,
    name: data.name,
    email: data.email,
    telefon: data.telefon || ''
  });
  logAudit('ADMIN_ADD_REGISTRATION',
    { anlassId: data.anlassId, email: data.email, source: 'admin-manual' },
    !!result.success, result.message);
  // Normalise the response shape so the admin UI sees `error` like
  // every other admin endpoint.
  if (result.success) return { success: true, message: result.message };
  return { success: false, error: result.message };
}

/**
 * Update Status / Notizen of a single registration. The (eventId,
 * email) pair is the natural key – registriereHelfer enforces it
 * unique on the live sheet.
 */
function updateRegistrationAdmin(data) {
  var eventId = String(data && data.eventId || '').trim();
  var email = String(data && data.email || '').toLowerCase().trim();
  if (!eventId || !email) return { success: false, error: 'eventId und email erforderlich.' };

  var newStatus = data.status ? String(data.status).toLowerCase() : null;
  if (newStatus && ANMELDUNG_STATUS_VALUES.indexOf(newStatus) === -1) {
    return { success: false, error: 'Ungültiger Status: ' + newStatus };
  }
  var newNotizen = data.notizen != null ? String(data.notizen) : null;
  if (newNotizen != null && newNotizen.length > MAX_DESC_LEN) {
    return { success: false, error: 'Notiz zu lang.' };
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var helferSheet = ss.getSheetByName('Anmeldungen');
  if (!helferSheet) return { success: false, error: 'Tabellenblatt fehlt.' };

  var lock = LockService.getScriptLock();
  try {
    if (!lock.tryLock(10000)) return { success: false, error: 'Server gerade ausgelastet.' };
    var lastRow = helferSheet.getLastRow();
    if (lastRow < 2) return { success: false, error: 'Anmeldung nicht gefunden.' };
    var values = helferSheet.getRange(2, 1, lastRow - 1, ANMELDUNG_HEADERS.length).getValues();
    for (var i = 0; i < values.length; i++) {
      var rowEvent = String(values[i][1] || '').trim();
      var rowEmail = String(values[i][3] || '').toLowerCase().trim();
      if (rowEvent !== eventId || rowEmail !== email) continue;
      var rowIdx = i + 2;
      if (newStatus) helferSheet.getRange(rowIdx, 7).setValue(newStatus);
      if (newNotizen != null) helferSheet.getRange(rowIdx, 8).setValue(sheetSafe(newNotizen));
      logAudit('ADMIN_UPDATE_REGISTRATION',
        { eventId: eventId, email: email, status: newStatus, notesLen: newNotizen ? newNotizen.length : null },
        true, null);
      return { success: true };
    }
    return { success: false, error: 'Anmeldung nicht gefunden.' };
  } finally {
    try { if (lock.hasLock()) lock.releaseLock(); } catch (_) {}
  }
}

/** Wrapper around archiviereJahr() for the admin POST path. */
function archiveJahrAdmin(data) {
  var jahr = String(data && data.jahr || '').trim();
  if (!jahr) return { success: false, error: 'Jahr erforderlich.' };
  var result = archiviereJahr(jahr);
  // archiviereJahr returns {success, message} — admin UI expects
  // {success, error|message}. Normalise.
  if (result.success) return result;
  return { success: false, error: result.message || 'Archiv fehlgeschlagen.' };
}

/** Wrapper around buildIntegrityReport for the admin POST path. */
function integrityCheckAdmin() {
  var report = buildIntegrityReport();
  if (report.success === false) return { success: false, error: report.error };
  logAudit('ADMIN_INTEGRITY_CHECK', { findings: report.findings.length }, true, null);
  return { success: true, findings: report.findings };
}

/**
 * Available school years for the archive dialog. Returns an array of
 * { jahr, events, registrations }. Only past years are returned —
 * the current year cannot be archived.
 */
function availableJahreAdmin() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var anlassSheet = ss.getSheetByName('Anlässe');
  if (!anlassSheet) return { success: true, years: [] };
  var current = jahrFor(new Date());
  var jahre = availableJahre(anlassSheet)
    .filter(function(y) { return y && y !== current; });
  var enriched = jahre.map(function(j) {
    var stats = jahrStats(j);
    return { jahr: j, events: stats.events, registrations: stats.registrations };
  });
  return { success: true, years: enriched, currentYear: current };
}

// === Get Active Events ===
function getAktiveAnlaesse() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Anlässe');
  if (!sheet) return [];

  var data = sheet.getDataRange().getValues();
  var heute = new Date();
  heute.setHours(0, 0, 0, 0);

  // Note: we deliberately do NOT read the Anmeldungen sheet here. The
  // public API response only carries counts, never names. Keeping the
  // personal data out of the GET response reduces the privacy surface
  // (parents' names used to be visible on the event cards to anyone
  // visiting the public site). The per-event count still works via
  // the 'Angemeldete' column of the Anlässe sheet which is maintained
  // transactionally by registriereHelfer().

  var anlaesse = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row[0]) continue;

    // Improved date parsing
    var datum = parseDate(row[2]);
    if (!datum) continue; // Skip invalid dates

    var maxHelfer = parseInt(row[4]) || 0;
    var aktuelleHelfer = parseInt(row[5]) || 0;
    var anlassId = String(row[0]);

    // Status (column H, index 7) is part of the post-PR-1 schema. Treat
    // a blank cell as "aktiv" for backward compatibility with sheets
    // that haven't been migrated yet via setupVerstaerken().
    var status = String(row[7] || 'aktiv').trim().toLowerCase();
    if (status && status !== 'aktiv') continue;

    // Normalize dates for comparison
    var datumNormalized = new Date(datum);
    datumNormalized.setHours(0, 0, 0, 0);

    // Include fully-booked events too – the frontend greys them out and
    // disables the registration CTA, while admins still need to access
    // the Helferliste download for full events.
    if (datumNormalized >= heute) {
      var freiePlaetze = Math.max(0, maxHelfer - aktuelleHelfer);
      anlaesse.push({
        id: anlassId,
        name: sanitizeInput(row[1]),
        datum: formatDatum(datum),
        datumSort: datumNormalized.getTime(),
        zeit: sanitizeInput(row[3] || ''),
        beschreibung: sanitizeInput(row[6] || ''),
        maxHelfer: maxHelfer,
        aktuelleHelfer: aktuelleHelfer,
        freiePlaetze: freiePlaetze,
        voll: aktuelleHelfer >= maxHelfer,
        // Per-event public contact (optional). Admins type these in
        // the admin UI / Sheet under "öffentlich sichtbar"; if empty,
        // the public card simply omits the contact line.
        kontaktName: sanitizeInput(String(row[11] || '')),
        kontaktEmail: sanitizeInput(String(row[12] || ''))
      });
    }
  }

  // Sort by date ascending (earliest first)
  anlaesse.sort(function(a, b) {
    return a.datumSort - b.datumSort;
  });

  return anlaesse;
}

// === Admin-only: Helfer list for a single event ===
// Used by the frontend to build a Word (.docx) Helferliste. Requires
// the ADMIN_KEY config value, compared with constant-time equality.
// Returns event metadata + every registration row for that event.
function getHelferList(eventId, providedKey) {
  var adminKey = getAdminKeyValue();
  if (!adminKey) {
    return { success: false, error: 'Admin-Zugriff ist nicht konfiguriert. Bitte ADMIN_KEY in den Script Properties setzen.' };
  }
  if (!providedKey || !secureEquals(String(providedKey), String(adminKey))) {
    logAudit('ADMIN_AUTH_FAIL', { eventId: eventId }, false, 'Invalid admin key');
    return { success: false, error: 'Keine Berechtigung.' };
  }
  if (!eventId) {
    return { success: false, error: 'Anlass-ID fehlt.' };
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var anlassSheet = ss.getSheetByName('Anlässe');
  var helferSheet = ss.getSheetByName('Anmeldungen');

  if (!anlassSheet || !helferSheet) {
    return { success: false, error: 'Tabellenblätter nicht gefunden.' };
  }

  // Find event row
  var anlassData = anlassSheet.getDataRange().getValues();
  var eventInfo = null;
  for (var i = 1; i < anlassData.length; i++) {
    if (String(anlassData[i][0]) === String(eventId)) {
      var datum = parseDate(anlassData[i][2]);
      eventInfo = {
        id: String(anlassData[i][0]),
        name: sanitizeInput(anlassData[i][1] || ''),
        datum: datum ? formatDatum(datum) : '',
        zeit: sanitizeInput(anlassData[i][3] || ''),
        maxHelfer: parseInt(anlassData[i][4]) || 0,
        aktuelleHelfer: parseInt(anlassData[i][5]) || 0,
        beschreibung: sanitizeInput(anlassData[i][6] || '')
      };
      break;
    }
  }

  if (!eventInfo) {
    return { success: false, error: 'Anlass nicht gefunden.' };
  }

  // Collect registrations for this event, preserving sheet order
  // (which is chronological insertion order).
  var helferData = helferSheet.getDataRange().getValues();
  var helpers = [];
  for (var j = 1; j < helferData.length; j++) {
    if (String(helferData[j][1] || '') === String(eventId)) {
      var ts = helferData[j][0];
      helpers.push({
        zeitstempel: ts instanceof Date ? ts.toISOString() : String(ts || ''),
        name: sanitizeInput(String(helferData[j][2] || '')),
        email: sanitizeInput(String(helferData[j][3] || '')),
        // Normalize on read so pre-existing rows (registered before the
        // write-side normalization was deployed, or edited manually in
        // the sheet) also display in canonical form.
        telefon: normalizePhone(sanitizeInput(String(helferData[j][4] || '')))
      });
    }
  }

  return {
    success: true,
    event: eventInfo,
    helpers: helpers
  };
}

// Constant-time string comparison to avoid timing side channels on the
// admin key. Returns false fast only when lengths differ.
function secureEquals(a, b) {
  if (a.length !== b.length) return false;
  var mismatch = 0;
  for (var i = 0; i < a.length; i++) {
    mismatch |= a.charCodeAt(i) ^ b.charCodeAt(i);
  }
  return mismatch === 0;
}

function formatDatum(datum) {
  var tage = ['Sonntag', 'Montag', 'Dienstag', 'Mittwoch', 'Donnerstag', 'Freitag', 'Samstag'];
  var monate = ['Januar', 'Februar', 'März', 'April', 'Mai', 'Juni', 'Juli', 'August', 'September', 'Oktober', 'November', 'Dezember'];
  return tage[datum.getDay()] + ', ' + datum.getDate() + '. ' + monate[datum.getMonth()] + ' ' + datum.getFullYear();
}

// === Register Helper ===
function registriereHelfer(data) {
  // Sanitize and validate input
  var anlassId = sanitizeInput(data.anlassId);
  var name = sanitizeInput(data.name || '');
  var email = (data.email || '').trim().toLowerCase();
  // Normalize phone to canonical Swiss "+41 XX XXX XX XX" so the
  // stored value is consistent regardless of how the parent typed it.
  var telefon = normalizePhone(sanitizeInput(data.telefon || ''));
  
  // Validation
  if (!anlassId || !name || !email) {
    return { success: false, message: 'Bitte füllen Sie alle Pflichtfelder aus.' };
  }
  
  if (name.length < 2) {
    return { success: false, message: 'Der Name muss mindestens 2 Zeichen lang sein.' };
  }
  if (name.length > MAX_NAME_LEN) {
    return { success: false, message: 'Der Name ist zu lang (max. ' + MAX_NAME_LEN + ' Zeichen).' };
  }

  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
    return { success: false, message: 'Bitte geben Sie eine gültige E-Mail-Adresse ein.' };
  }
  if (email.length > 254) {
    return { success: false, message: 'Die E-Mail-Adresse ist zu lang.' };
  }

  if (telefon.length > MAX_TEL_LEN) {
    return { success: false, message: 'Die Telefonnummer ist zu lang (max. ' + MAX_TEL_LEN + ' Zeichen).' };
  }
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var anlassSheet = ss.getSheetByName('Anlässe');
  var helferSheet = ss.getSheetByName('Anmeldungen');
  
  if (!anlassSheet || !helferSheet) {
    return { success: false, message: 'Der Anlass konnte nicht verarbeitet werden. Bitte versuchen Sie es später erneut.' };
  }
  
  // Use LockService to prevent race conditions. try/finally guarantees
  // the lock is released even if an unexpected error escapes the body.
  var lock = LockService.getScriptLock();
  var anlassName = '';
  try {
    lock.waitLock(10000); // up to 10 s

    // Find event
    var anlassData = anlassSheet.getDataRange().getValues();
    var anlassRow = -1, maxHelfer = 0, aktuelleHelfer = 0;

    for (var i = 1; i < anlassData.length; i++) {
      if (String(anlassData[i][0]) === String(anlassId)) {
        anlassRow = i + 1;
        anlassName = sanitizeInput(anlassData[i][1] || '');
        maxHelfer = parseInt(anlassData[i][4]) || 0;
        aktuelleHelfer = parseInt(anlassData[i][5]) || 0;
        break;
      }
    }

    if (anlassRow === -1) {
      return { success: false, message: 'Der gewählte Anlass wurde nicht gefunden.' };
    }

    // Re-check capacity after lock (prevent race condition)
    if (aktuelleHelfer >= maxHelfer) {
      return { success: false, message: 'Leider sind bereits alle Plätze vergeben.' };
    }

    // Duplicate check: (anlassId, email) pair – name can vary slightly
    // in casing/whitespace between submissions by the same parent, and
    // we don't want to accept two rows for the same event+email.
    var helferData = helferSheet.getDataRange().getValues();
    for (var j = 1; j < helferData.length; j++) {
      var existingEmail = String(helferData[j][3] || '').toLowerCase().trim();
      var existingAnlassId = String(helferData[j][1] || '');
      if (existingAnlassId === anlassId && existingEmail === email) {
        return { success: false, message: 'Sie sind bereits für diesen Anlass angemeldet.' };
      }
    }

    // Register (transaction-safe). sheetSafe() prefixes any string
    // starting with a formula trigger (=, +, -, @, tab, CR) with a
    // single quote so Sheets stores it as text. Without this, "+41 79
    // ..." phone numbers are silently parsed as unary-plus formulas
    // and either lose the leading "+" or become #ERROR!.
    // Append registration. The Anlässe!F (Angemeldete) column is now a
    // COUNTIFS formula maintained by Sheets, so we no longer write the
    // count manually – that would clobber the formula. Capacity remains
    // race-safe because LockService serialises registrations and the
    // formula recomputes synchronously when the next request reads it.
    // The trailing 'aktiv' goes into Anmeldungen!G (Status column added
    // by setupVerstaerken). On unmigrated sheets it sits as data in an
    // unlabelled column G; harmless and forward-compatible.
    helferSheet.appendRow([
      new Date(),
      sheetSafe(anlassId),
      sheetSafe(name),
      sheetSafe(email),
      sheetSafe(telefon),
      sheetSafe(anlassName),
      'aktiv'
    ]);
    // Keep rows grouped by event so admins skimming the sheet see all
    // helpers for one Anlass together. Cheap (<1s for thousands of
    // rows), race-safe inside the existing lock.
    sortAnmeldungenByAnlass(helferSheet);

  } catch (e) {
    Logger.log('registriereHelfer error: ' + e);
    var msg = 'Ein Fehler ist aufgetreten. Bitte versuchen Sie es erneut.';
    if (String(e).indexOf('lock') !== -1 || String(e).indexOf('Lock') !== -1) {
      msg = 'Der Server ist gerade stark ausgelastet. Bitte versuchen Sie es in wenigen Sekunden erneut.';
    }
    return { success: false, message: msg };
  } finally {
    try { if (lock.hasLock()) lock.releaseLock(); } catch (_) {}
  }

  // Outside the lock: optional email notification (non-critical).
  if (ADMIN_EMAIL) {
    var emailBody = 'Neue Anmeldung für Schulhelfer:\n\n' +
                   'Anlass: ' + anlassName + '\n' +
                   'Name: ' + name + '\n' +
                   'E-Mail: ' + email + '\n' +
                   (telefon ? 'Telefon: ' + telefon + '\n' : '') +
                   'Datum: ' + new Date().toLocaleString('de-CH');
    sendEmailNotification(ADMIN_EMAIL, 'Neue Anmeldung: ' + anlassName, emailBody);
  }

  return {
    success: true,
    message: 'Vielen Dank, ' + name + '! Sie sind für «' + anlassName + '» angemeldet.'
  };
}

// === Setup ===
function erstesSetup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Create Anlässe sheet
  var anlassSheet = ss.getSheetByName('Anlässe') || ss.insertSheet('Anlässe');
  anlassSheet.clear();
  anlassSheet.getRange(1, 1, 1, 7).setValues([
    ['ID', 'Name', 'Datum', 'Zeit', 'Benötigte Helfer', 'Angemeldete', 'Beschreibung']
  ]);
  anlassSheet.getRange(1, 1, 1, 7).setBackground('#1e3a5f').setFontColor('white').setFontWeight('bold');
  
  // Echte Anlässe der Primarstufe Rittergasse 2025/26
  var anlaesse = [
    ['1', 'Adventssingen – Aufbau (Probe)', new Date(2025, 11, 8), '07:30-10:00', 8, 0, 'Aufbau für die Probe am Montag'],
    ['2', 'Adventssingen – Aufbau & Licht', new Date(2025, 11, 10), '07:30-10:00 / 18:45-20:00', 8, 0, 'Morgens Aufbau, abends Lichtunterstützung'],
    ['3', 'Adventssingen – Einlass Münster', new Date(2025, 11, 10), '18:30-20:00', 5, 0, 'Einlass am Münster-Eingang (4 Pers.) und innen (1-2 Pers.)'],
    ['4', 'Adventssingen – Abbau', new Date(2025, 11, 10), 'ab 20:00', 20, 0, 'Abbauen und Rücktransport des Materials in den Schulhaus-Keller'],
    ['5', 'Fasnachtsumzug', new Date(2026, 1, 12), 'ganztags', 10, 0, 'Begleitung der Kinder am Fasnachtsumzug'],
    ['6', 'Sporttag Kindergarten & Unterstufe', new Date(2026, 4, 12), 'ab 08:15', 12, 0, 'Kannenfeldpark – Mithilfe bei Posten, Zeitnahme, Ergebnisse'],
    ['7', 'Sporttag Mittelstufe', new Date(2026, 5, 3), '08:00-12:30', 15, 0, 'Leichtathletikstadion St. Jakob – Mithilfe bei Posten, Zeitnahme, Ergebnisse'],
    ['8', 'Schuljahresabschluss – Dekoration', new Date(2026, 5, 25), 'ab 16:00', 8, 0, 'Schulhof PS Rittergasse – Dekoration, Feuerschalen aufstellen'],
    ['9', 'Schuljahresabschluss – Abbau', new Date(2026, 5, 25), 'ab 21:00', 10, 0, 'Schulhof PS Rittergasse – Abbau der Tischgarnituren'],
    ['10', 'Elterncafé am 1. Schultag', new Date(2026, 7, 10), '09:00-10:45', 6, 0, 'Schulhof PS Rittergasse – Begrüssung der Eltern der neuen 1. Klassen']
  ];
  anlassSheet.getRange(2, 1, anlaesse.length, 7).setValues(anlaesse);
  
  // Format columns
  anlassSheet.setColumnWidth(1, 40);
  anlassSheet.setColumnWidth(2, 250);
  anlassSheet.setColumnWidth(3, 100);
  anlassSheet.setColumnWidth(4, 150);
  anlassSheet.setColumnWidth(5, 100);
  anlassSheet.setColumnWidth(6, 90);
  anlassSheet.setColumnWidth(7, 350);
  
  // Create Anmeldungen sheet
  var helferSheet = ss.getSheetByName('Anmeldungen') || ss.insertSheet('Anmeldungen');
  helferSheet.clear();
  helferSheet.getRange(1, 1, 1, 6).setValues([
    ['Zeitstempel', 'Anlass-ID', 'Name', 'E-Mail', 'Telefon', 'Anlass']
  ]);
  helferSheet.getRange(1, 1, 1, 6).setBackground('#1a7d36').setFontColor('white').setFontWeight('bold');
  
  // Remove default sheet
  try { 
    var s1 = ss.getSheetByName('Sheet1') || ss.getSheetByName('Tabelle1'); 
    if (s1) ss.deleteSheet(s1); 
  } catch(e) {}
  
  SpreadsheetApp.getUi().alert(
    '✅ Setup abgeschlossen!\n\n' +
    '10 Anlässe für 2025/26 wurden eingetragen.\n\n' +
    'Nächster Schritt:\n' +
    '1. Klicken Sie auf "Bereitstellen" → "Neue Bereitstellung"\n' +
    '2. Typ: Web-App\n' +
    '3. Ausführen als: Ich\n' +
    '4. Zugriff: JEDER (wichtig!)\n' +
    '5. Kopieren Sie die URL'
  );
}

// === Export Data ===

/**
 * RFC 4180-ish CSV quoting. Wraps cells containing commas, quotes or
 * newlines in double quotes and doubles any embedded quotes. Dates are
 * rendered via toISOString() so they round-trip cleanly.
 */
function csvCell(value) {
  if (value === null || value === undefined) return '';
  var s;
  if (value instanceof Date) {
    s = isNaN(value.getTime()) ? '' : value.toISOString();
  } else {
    s = String(value);
  }
  if (/[",\n\r]/.test(s)) {
    return '"' + s.replace(/"/g, '""') + '"';
  }
  return s;
}

function csvRow(row) {
  return row.map(csvCell).join(',');
}

function exportData() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var anlassSheet = ss.getSheetByName('Anlässe');
    var helferSheet = ss.getSheetByName('Anmeldungen');

    if (!anlassSheet || !helferSheet) {
      return { success: false, error: 'Tabellenblätter nicht gefunden' };
    }

    var anlaesse = anlassSheet.getDataRange().getValues();
    var anmeldungen = helferSheet.getDataRange().getValues();

    // Convert to CSV format (RFC 4180 quoting)
    var lines = [];

    lines.push('=== ANLÄSSE ===');
    for (var i = 0; i < anlaesse.length; i++) {
      lines.push(csvRow(anlaesse[i]));
    }

    lines.push('');
    lines.push('=== ANMELDUNGEN ===');
    for (var j = 0; j < anmeldungen.length; j++) {
      lines.push(csvRow(anmeldungen[j]));
    }

    return { success: true, data: lines.join('\n') + '\n' };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

/**
 * Write both sheets into a consolidated "Export" tab with proper
 * columns (not as a single CSV column – the previous version dumped
 * RFC 4180 CSV text into column A, which is unusable for review).
 * The action=export GET endpoint (exportData()) still returns CSV for
 * any API caller that wants it.
 */
function exportDataAsCSV() {
  var ui = SpreadsheetApp.getUi();
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var anlassSheet = ss.getSheetByName('Anlässe');
    var helferSheet = ss.getSheetByName('Anmeldungen');
    if (!anlassSheet || !helferSheet) {
      ui.alert('Tabellenblätter nicht gefunden.');
      return;
    }
    var anlaesse = anlassSheet.getDataRange().getValues();
    var anmeldungen = helferSheet.getDataRange().getValues();

    var exportSheet = ss.getSheetByName('Export') || ss.insertSheet('Export');
    exportSheet.clear();

    var row = 1;
    exportSheet.getRange(row, 1).setValue('=== ANLÄSSE ===')
      .setFontWeight('bold').setFontColor('#1e3a5f').setFontSize(12);
    row++;
    if (anlaesse.length) {
      exportSheet.getRange(row, 1, anlaesse.length, anlaesse[0].length).setValues(anlaesse);
      exportSheet.getRange(row, 1, 1, anlaesse[0].length)
        .setFontWeight('bold').setBackground('#e8eef5');
      row += anlaesse.length;
    }
    row += 1; // blank separator

    exportSheet.getRange(row, 1).setValue('=== ANMELDUNGEN ===')
      .setFontWeight('bold').setFontColor('#1a7d36').setFontSize(12);
    row++;
    if (anmeldungen.length) {
      exportSheet.getRange(row, 1, anmeldungen.length, anmeldungen[0].length).setValues(anmeldungen);
      exportSheet.getRange(row, 1, 1, anmeldungen[0].length)
        .setFontWeight('bold').setBackground('#e5f3e8');
      row += anmeldungen.length;
    }

    // Make it readable
    var maxCols = Math.max(
      anlaesse.length ? anlaesse[0].length : 0,
      anmeldungen.length ? anmeldungen[0].length : 0,
      1
    );
    exportSheet.autoResizeColumns(1, maxCols);

    ss.setActiveSheet(exportSheet);
    logAudit('EXPORT_DATA_UI', { anlaesse: anlaesse.length - 1, anmeldungen: anmeldungen.length - 1 }, true, null);
    ui.alert('✅ Export erfolgreich. Daten im Tab "Export".');
  } catch (e) {
    logAudit('EXPORT_DATA_UI', {}, false, String(e));
    ui.alert('Fehler beim Export: ' + e);
  }
}

/**
 * Delete registrations and events that are older than `months` months.
 * Respects Swiss DSG purpose limitation: personal data should not be
 * retained beyond the purpose (= the event).
 */
function alteAnlaesseBereinigen() {
  var ui = SpreadsheetApp.getUi();
  var resp = ui.prompt(
    '🗑️ Alte Anlässe bereinigen',
    'Wie viele Monate nach dem Anlass sollen Daten aufbewahrt werden?\n' +
    '(z.B. 3 = Anlässe, die vor mehr als 3 Monaten stattfanden, werden gelöscht)',
    ui.ButtonSet.OK_CANCEL
  );
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  var months = parseInt(resp.getResponseText(), 10);
  if (isNaN(months) || months < 0) {
    ui.alert('Ungültige Eingabe.');
    return;
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var anlassSheet = ss.getSheetByName('Anlässe');
  var helferSheet = ss.getSheetByName('Anmeldungen');
  if (!anlassSheet || !helferSheet) {
    ui.alert('Tabellenblätter nicht gefunden.');
    return;
  }

  var cutoff = new Date();
  cutoff.setMonth(cutoff.getMonth() - months);
  cutoff.setHours(0, 0, 0, 0);

  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    var anlassData = anlassSheet.getDataRange().getValues();
    var expiredIds = {};
    var anlassRowsToDelete = [];

    for (var i = anlassData.length - 1; i >= 1; i--) {
      var datum = parseDate(anlassData[i][2]);
      if (datum && datum < cutoff) {
        expiredIds[String(anlassData[i][0])] = true;
        anlassRowsToDelete.push(i + 1);
      }
    }

    var helferData = helferSheet.getDataRange().getValues();
    var helferRowsToDelete = [];
    for (var j = helferData.length - 1; j >= 1; j--) {
      if (expiredIds[String(helferData[j][1] || '')]) {
        helferRowsToDelete.push(j + 1);
      }
    }

    // Delete bottom-up so row indices stay valid
    for (var k = 0; k < helferRowsToDelete.length; k++) {
      helferSheet.deleteRow(helferRowsToDelete[k]);
    }
    for (var l = 0; l < anlassRowsToDelete.length; l++) {
      anlassSheet.deleteRow(anlassRowsToDelete[l]);
    }

    logAudit('BEREINIGUNG', {
      months: months,
      anlaesse: anlassRowsToDelete.length,
      anmeldungen: helferRowsToDelete.length
    }, true, null);

    ui.alert(
      '✅ Bereinigung abgeschlossen.\n\n' +
      anlassRowsToDelete.length + ' Anlass/Anlässe und ' +
      helferRowsToDelete.length + ' Anmeldung(en) gelöscht.\n' +
      '(Stichtag: älter als ' + months + ' Monate)'
    );
  } catch (e) {
    logAudit('BEREINIGUNG', {}, false, String(e));
    ui.alert('Fehler: ' + e);
  } finally {
    try { if (lock.hasLock()) lock.releaseLock(); } catch (_) {}
  }
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('🏰 Schulhelfer')
    .addItem('Erstes Setup', 'erstesSetup')
    .addItem('Setup verstärken (Validierung & Formeln)', 'setupVerstaerken')
    .addSeparator()
    .addItem('Neuer Anlass hinzufügen', 'neuerAnlassDialog')
    .addItem('Alle Anmeldungen anzeigen', 'zeigeAnmeldungen')
    .addItem('Zähler neu berechnen', 'zaehlerNeuBerechnen')
    .addItem('Anmeldungen nach Anlass sortieren', 'sortiereAnmeldungenMenu')
    .addItem('Telefonnummern normalisieren', 'normalisiereTelefonnummernMenu')
    .addItem('Daten exportieren', 'exportDataAsCSV')
    .addSeparator()
    .addItem('Jahr archivieren…', 'archiviereJahrDialog')
    .addItem('Daten prüfen', 'integritaetspruefung')
    .addSeparator()
    .addItem('Alte Anlässe bereinigen', 'alteAnlaesseBereinigen')
    .addItem('Admin-Status prüfen', 'adminKeyStatus')
    .addItem('Script-Properties aufräumen', 'bereinigeScriptProperties')
    .addItem('Audit-Log anzeigen', 'zeigeAuditLog')
    .addToUi();
}

/**
 * Recompute the "Angemeldete" column of the Anlässe sheet from the
 * actual rows in the Anmeldungen sheet. The transactional counter
 * maintained by registriereHelfer() can drift from reality if an admin
 * manually deletes or edits registration rows – this reconciles it.
 */
function zaehlerNeuBerechnen() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var anlassSheet = ss.getSheetByName('Anlässe');
  var helferSheet = ss.getSheetByName('Anmeldungen');
  if (!anlassSheet || !helferSheet) {
    ui.alert('Tabellenblätter "Anlässe" und/oder "Anmeldungen" fehlen.');
    return;
  }
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var helferData = helferSheet.getDataRange().getValues();
    var counts = {};
    for (var j = 1; j < helferData.length; j++) {
      var id = String(helferData[j][1] || '').trim();
      if (!id) continue;
      counts[id] = (counts[id] || 0) + 1;
    }
    var anlassData = anlassSheet.getDataRange().getValues();
    var changed = 0;
    for (var i = 1; i < anlassData.length; i++) {
      var eventId = String(anlassData[i][0] || '').trim();
      if (!eventId) continue;
      var newCount = counts[eventId] || 0;
      var oldCount = parseInt(anlassData[i][5], 10) || 0;
      if (newCount !== oldCount) {
        anlassSheet.getRange(i + 1, 6).setValue(newCount);
        changed++;
      }
    }
    sortAnmeldungenByAnlass(helferSheet);
    logAudit('ZAEHLER_NEU_BERECHNET', { changed: changed }, true, null);
    ui.alert('✅ Zähler neu berechnet (und Anmeldungen nach Anlass sortiert).\n\n' + changed + ' Anlass-Zeile(n) aktualisiert.');
  } catch (e) {
    logAudit('ZAEHLER_NEU_BERECHNET', {}, false, String(e));
    ui.alert('Fehler: ' + e);
  } finally {
    try { if (lock.hasLock()) lock.releaseLock(); } catch (_) {}
  }
}

/**
 * Show the current admin-key configuration so the organiser can check
 * whether the Helferliste download is enabled, and warn if the key is
 * dangerously short.
 */
function adminKeyStatus() {
  var ui = SpreadsheetApp.getUi();
  var key = getAdminKeyValue();
  var source = '';
  try {
    source = PropertiesService.getScriptProperties().getProperty('ADMIN_KEY') ? 'Script Properties' : 'Code-Konstante';
  } catch (e) { source = 'unbekannt'; }

  if (!key) {
    ui.alert(
      '🔑 Admin-Status',
      'Es ist KEIN Admin-Schlüssel gesetzt.\n\n' +
      'Der "Helferliste (Word)"-Download ist damit deaktiviert.\n\n' +
      'Einrichten:\n' +
      '1. Projekt-Einstellungen → Script Properties\n' +
      '2. Eigenschaft: ADMIN_KEY = <langer zufälliger Wert>\n' +
      '3. Web-App neu bereitstellen.',
      ui.ButtonSet.OK
    );
    return;
  }
  var msg = 'Ein Admin-Schlüssel ist gesetzt (Länge: ' + key.length + ' Zeichen, Quelle: ' + source + ').';
  if (key.length < 16) {
    msg += '\n\n⚠️ HINWEIS: Der Schlüssel ist kurz. Für produktiven Einsatz ' +
           'sollten Sie mindestens 16 zufällige Zeichen verwenden.';
  }
  msg += '\n\nZum Aktivieren auf einem Gerät die Seite einmal mit ?admin=SCHLÜSSEL ' +
         'am Ende der URL öffnen. Der Schlüssel wird dann lokal gespeichert.';
  ui.alert('🔑 Admin-Status', msg, ui.ButtonSet.OK);
}

/**
 * One-shot cleanup of leftover Script Properties from the old
 * PropertiesService-based rate limiter. The current code uses
 * CacheService for rate limits (auto-expires, never written to
 * properties), but old "rate_<identifier>" entries from before
 * v25-improvements still sit in Script Properties and now exceed the
 * 50-row visible limit in the editor – making it impossible to find
 * ADMIN_KEY in the UI.
 *
 * Deletes every property whose key starts with "rate_". Preserves
 * ADMIN_KEY and any other manually-set property. Confirmable: shows
 * the planned counts before doing anything.
 */
function bereinigeScriptProperties() {
  var ui = SpreadsheetApp.getUi();
  var props = PropertiesService.getScriptProperties();
  var all = props.getProperties();
  var keys = Object.keys(all);
  var rateKeys = keys.filter(function(k) { return k.indexOf('rate_') === 0; });
  var keepKeys = keys.filter(function(k) { return k.indexOf('rate_') !== 0; });

  if (rateKeys.length === 0) {
    ui.alert('Aufräumen', 'Keine alten Rate-Limit-Einträge gefunden. Aktuell ' +
             keys.length + ' Properties insgesamt.', ui.ButtonSet.OK);
    return;
  }

  var resp = ui.alert(
    'Aufräumen',
    'Gefunden: ' + rateKeys.length + ' alte "rate_*"-Einträge ' +
    '(verbleibende ' + keepKeys.length + ' Properties bleiben unverändert, darunter ' +
    (keepKeys.indexOf('ADMIN_KEY') !== -1 ? 'ADMIN_KEY' : 'KEIN ADMIN_KEY') + ').\n\n' +
    'Diese Einträge sind Reste der alten Rate-Limit-Implementierung und werden ' +
    'nicht mehr verwendet. Jetzt löschen?',
    ui.ButtonSet.YES_NO
  );
  if (resp !== ui.Button.YES) return;

  // Delete in chunks – Apps Script may not love deleting hundreds at once.
  var deleted = 0;
  for (var i = 0; i < rateKeys.length; i++) {
    try {
      props.deleteProperty(rateKeys[i]);
      deleted++;
    } catch (e) {
      Logger.log('Could not delete ' + rateKeys[i] + ': ' + e);
    }
  }
  logAudit('SCRIPT_PROPERTIES_CLEANED', { deleted: deleted, kept: keepKeys.length }, true, null);
  ui.alert('✅ Aufgeräumt',
    deleted + ' alte Einträge gelöscht.\n' +
    keepKeys.length + ' Properties verbleiben (ADMIN_KEY ' +
    (keepKeys.indexOf('ADMIN_KEY') !== -1 ? 'enthalten' : 'NICHT gefunden – ggf. neu setzen') + ').',
    ui.ButtonSet.OK);
}

function zeigeAnmeldungen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Anmeldungen');
  if (sheet) {
    ss.setActiveSheet(sheet);
  }
}

function zeigeAuditLog() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Audit-Log');
  if (sheet) {
    ss.setActiveSheet(sheet);
  } else {
    SpreadsheetApp.getUi().alert('Noch keine Audit-Logs vorhanden.');
  }
}

function neuerAnlassDialog() {
  var html = HtmlService.createHtmlOutput(`
    <style>
      body { font-family: Arial, sans-serif; padding: 16px; }
      label { display: block; margin-top: 12px; font-weight: bold; color: #333; }
      input, textarea { 
        width: 100%; padding: 10px; margin-top: 4px; 
        border: 1px solid #ddd; border-radius: 6px; 
        box-sizing: border-box; font-size: 14px;
      }
      input:focus, textarea:focus { outline: none; border-color: #1e3a5f; }
      button { 
        margin-top: 20px; padding: 12px 24px; 
        background: #1e3a5f; color: white; 
        border: none; border-radius: 6px; 
        cursor: pointer; font-size: 14px; font-weight: bold;
      }
      button:hover { background: #2d5a8a; }
    </style>
    <form id="f">
      <label>Name des Anlasses *</label>
      <input id="n" required placeholder="z.B. Sommerfest">
      
      <label>Datum *</label>
      <input type="date" id="d" required>
      
      <label>Zeit (optional)</label>
      <input id="z" placeholder="z.B. 14:00-18:00">
      
      <label>Anzahl benötigter Helfer *</label>
      <input type="number" id="h" min="1" required placeholder="z.B. 5">
      
      <label>Beschreibung (optional)</label>
      <textarea id="b" rows="2" placeholder="Was sollen die Helfer tun?"></textarea>

      <label>Kontaktperson (optional, ÖFFENTLICH SICHTBAR)</label>
      <input id="kn" maxlength="80" placeholder="z.B. Frau Müller, Sekretariat">

      <label>Kontakt-Email (optional, ÖFFENTLICH SICHTBAR)</label>
      <input id="ke" type="email" maxlength="254" placeholder="ansprech@schule.bs.ch">

      <p style="font-size:12px;color:#64748b;margin-top:6px;line-height:1.4;">
        Falls ausgefüllt, erscheinen Name &amp; Email auf jeder Anlass-Karte
        und in den iCal-Einträgen. Bitte nur Schulkontakt-Daten verwenden,
        keine privaten Mobilnummern.
      </p>

      <button type="submit">✓ Anlass erstellen</button>
    </form>
    <script>
      document.getElementById("f").onsubmit = function(e) {
        e.preventDefault();
        google.script.run
          .withSuccessHandler(function() {
            alert('Anlass erfolgreich erstellt!');
            google.script.host.close();
          })
          .withFailureHandler(function(err) {
            alert('Fehler: ' + err);
          })
          .anlassHinzufuegen({
            name: document.getElementById("n").value,
            datum: document.getElementById("d").value,
            zeit: document.getElementById("z").value,
            helfer: document.getElementById("h").value,
            beschreibung: document.getElementById("b").value,
            kontaktName: document.getElementById("kn").value,
            kontaktEmail: document.getElementById("ke").value
          });
      };
    </script>
  `).setWidth(420).setHeight(620);

  SpreadsheetApp.getUi().showModalDialog(html, '🏰 Neuer Anlass');
}

function anlassHinzufuegen(data) {
  // Sanitize & validate input first (outside the lock).
  var name = sanitizeInput(data.name || '');
  var zeit = sanitizeInput(data.zeit || '');
  var beschreibung = sanitizeInput(data.beschreibung || '');
  var helfer = parseInt(data.helfer) || 0;
  var kontaktName = sanitizeInput(data.kontaktName || '');
  var kontaktEmail = String(data.kontaktEmail || '').trim().toLowerCase();

  if (!name || helfer < 1) {
    throw new Error('Bitte füllen Sie alle Pflichtfelder aus.');
  }
  if (name.length > MAX_EVENT_NAME_LEN) {
    throw new Error('Der Anlass-Name ist zu lang (max. ' + MAX_EVENT_NAME_LEN + ' Zeichen).');
  }
  if (beschreibung.length > MAX_DESC_LEN) {
    throw new Error('Die Beschreibung ist zu lang (max. ' + MAX_DESC_LEN + ' Zeichen).');
  }
  if (kontaktName.length > 80) {
    throw new Error('Kontaktperson zu lang (max. 80 Zeichen).');
  }
  if (kontaktEmail && (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(kontaktEmail) || kontaktEmail.length > 254)) {
    throw new Error('Ungültige Kontakt-Email.');
  }

  var datum = parseDate(data.datum);
  if (!datum) {
    throw new Error('Ungültiges Datum.');
  }

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Anlässe');
  if (!sheet) throw new Error('Tabellenblatt "Anlässe" nicht gefunden.');

  // Lock guarantees two concurrent invocations can't compute the same
  // newId and collide on the ID column. We also scan archived
  // Anlässe sheets so IDs are never reused across school years –
  // helps audit-trail integrity and iCal UIDs.
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var newId = nextAnlassId();
    // Status="aktiv" by default; Jahr is derived. The COUNTIFS
    // formula for the Angemeldete column is set in a follow-up
    // setFormula() call because the formula needs the row index, which
    // we only know after the append.
    sheet.appendRow([
      newId,
      sheetSafe(name),
      datum,
      sheetSafe(zeit),
      helfer,
      0, // Angemeldete – overwritten with formula immediately below
      sheetSafe(beschreibung),
      'aktiv',
      jahrFor(datum),
      '', // Sichtbar? – formula written by setupVerstaerken/refresh
      '', // Notizen
      sheetSafe(kontaktName),
      sheetSafe(kontaktEmail)
    ]);
    var newRow = sheet.getLastRow();
    sheet.getRange(newRow, 6).setFormula(angemeldeteFormulaForRow(newRow));
    sheet.getRange(newRow, 10).setFormula(sichtbarFormulaForRow(newRow));
    logAudit('ANLASS_HINZUGEFUEGT', { id: newId, name: name, datum: datum }, true, null);
  } finally {
    try { if (lock.hasLock()) lock.releaseLock(); } catch (_) {}
  }
}

// =============================================================
// === Schema foundation (PR 1: Sheet schema robustness) =======
// =============================================================
//
// The Sheet, not the script, is the source of truth. Counts come from
// COUNTIFS over Anmeldungen. Status drives visibility. setupVerstaerken
// is idempotent: re-running it re-applies validations, formulas,
// conditional formatting, and header protection without disturbing data.

var ANLASS_HEADERS = [
  'ID', 'Name', 'Datum', 'Zeit', 'Benötigte Helfer', 'Angemeldete',
  'Beschreibung', 'Status', 'Jahr', 'Sichtbar?', 'Notizen (intern)',
  // Per-event public contact person. Both fields are optional and
  // appear ON THE PUBLIC SITE if filled in. The admin UI labels them
  // "(öffentlich sichtbar)" so admins know not to put private mobile
  // numbers in here.
  'Kontaktperson', 'Kontakt-Email'
];
var ANMELDUNG_HEADERS = [
  'Zeitstempel', 'Anlass-ID', 'Name', 'E-Mail', 'Telefon', 'Anlass',
  'Status', 'Notizen'
];
var ANLASS_STATUS_VALUES = ['aktiv', 'abgesagt', 'archiviert'];
var ANMELDUNG_STATUS_VALUES = ['aktiv', 'storniert', 'nicht erschienen'];

/**
 * Calendar year as a 4-digit string, "YYYY". Returns '' for invalid
 * dates. The school explicitly chose calendar-year archival over
 * Swiss-school-year (Aug→Jul) so that year boundaries are
 * unambiguous and auto-archive can fire on Jan 1.
 */
function jahrFor(date) {
  if (!(date instanceof Date) || isNaN(date.getTime())) return '';
  return String(date.getFullYear());
}

/**
 * Per-row formula for Anlässe!F (Angemeldete). Counts every row in
 * Anmeldungen that has the same Anlass-ID and a Status that is not
 * "storniert". Blank Status counts as active – keeps backward
 * compatibility with rows registered before the Status column existed.
 */
function angemeldeteFormulaForRow(row) {
  return '=IF(A' + row + '="","",COUNTIFS(' +
         'Anmeldungen!$B:$B,A' + row + ',' +
         'Anmeldungen!$G:$G,"<>storniert"))';
}

/**
 * Per-row formula for Anlässe!J (Sichtbar?). Tells the admin in plain
 * German why an event is or is not on the public site. Mirrors the
 * filter logic of getAktiveAnlaesse() but lives in the Sheet so admins
 * see it without opening the website.
 */
function sichtbarFormulaForRow(row) {
  return '=IFS(' +
         'A' + row + '="","",' +
         'H' + row + '="abgesagt","nein – abgesagt",' +
         'H' + row + '="archiviert","nein – archiviert",' +
         'NOT(ISNUMBER(C' + row + ')),"nein – Datum ungültig",' +
         'C' + row + '<TODAY(),"nein – Datum vorbei",' +
         'AND(ISNUMBER(E' + row + '),ISNUMBER(F' + row + '),F' + row + '>=E' + row + '),"voll – wird trotzdem angezeigt",' +
         'TRUE,"ja"' +
         ')';
}

/**
 * Search every Anlass-bearing sheet (live + archive) for the highest
 * numeric ID and return the next one. Prevents ID reuse across school
 * years so iCal UIDs and audit-log references stay unambiguous forever.
 */
function nextAnlassId() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var max = 0;
  ss.getSheets().forEach(function(s) {
    var name = s.getName();
    if (name !== 'Anlässe' && name.indexOf('Archiv Anlässe') !== 0) return;
    var lastRow = s.getLastRow();
    if (lastRow < 2) return;
    var ids = s.getRange(2, 1, lastRow - 1, 1).getValues();
    ids.forEach(function(r) {
      var n = parseInt(r[0], 10);
      if (!isNaN(n) && n > max) max = n;
    });
  });
  return max + 1;
}

/**
 * Idempotent migration / re-application of the robust schema. Adds
 * missing columns, applies validations, sets formulas, applies
 * conditional formatting, protects header rows, back-fills defaults,
 * and ensures the Anleitung sheet exists. Safe to run any number of
 * times.
 */
function setupVerstaerken() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var anlassSheet = ss.getSheetByName('Anlässe');
  var helferSheet = ss.getSheetByName('Anmeldungen');
  if (!anlassSheet || !helferSheet) {
    ui.alert('Bitte zuerst "Erstes Setup" ausführen – die Tabellenblätter "Anlässe" und "Anmeldungen" fehlen.');
    return;
  }

  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);

    ensureHeaders(anlassSheet, ANLASS_HEADERS, '#1e3a5f');
    ensureHeaders(helferSheet, ANMELDUNG_HEADERS, '#1a7d36');

    backfillAnlassDefaults(anlassSheet);
    backfillAnmeldungDefaults(helferSheet);
    sortAnmeldungenByAnlass(helferSheet);
    // Sweep all live + archived phone numbers into the canonical
    // "+41 XX XXX XX XX" form. Idempotent: rows already canonical are
    // skipped. Logs the totals to the audit trail.
    var phoneStats = normalisiereAlleTelefonnummern();
    if (phoneStats.changed) {
      logAudit('PHONES_NORMALIZED_VIA_SETUP', phoneStats, true, null);
    }

    writeAnlassFormulas(anlassSheet);
    applyAnlassValidations(anlassSheet);
    applyAnmeldungValidations(helferSheet);
    applyAnlassConditionalFormatting(anlassSheet);
    protectHeaderRow(anlassSheet);
    protectHeaderRow(helferSheet);
    ensureAnleitungSheet(ss);

    logAudit('SETUP_VERSTAERKT', {
      anlassRows: Math.max(anlassSheet.getLastRow() - 1, 0),
      anmeldungRows: Math.max(helferSheet.getLastRow() - 1, 0)
    }, true, null);
  } catch (e) {
    logAudit('SETUP_VERSTAERKT', {}, false, String(e));
    ui.alert('Fehler beim Verstärken: ' + e);
    return;
  } finally {
    try { if (lock.hasLock()) lock.releaseLock(); } catch (_) {}
  }

  // Install the daily auto-archive trigger outside the lock (the
  // ScriptApp.newTrigger call hits Google's trigger service, no
  // sheet contention). Best-effort: if the user denies the
  // additional auth, surface it but don't undo the rest of setup.
  var triggerStatus = '';
  try {
    installiereAutoArchiv();
    triggerStatus = '\n• Auto-Archiv aktiv: täglich um 03:00 werden Anlässe abgeschlossener Kalenderjahre automatisch archiviert.';
  } catch (e) {
    logAudit('AUTO_ARCHIVE_INSTALL_FAIL', {}, false, String(e));
    triggerStatus = '\n⚠️  Auto-Archiv konnte nicht installiert werden: ' + e +
                    '\n   (Sie können es später über das Menü manuell auslösen.)';
  }

  ui.alert(
    '✅ Setup verstärkt.\n\n' +
    '• Spalten "Status", "Jahr", "Sichtbar?", "Notizen" auf Anlässe.\n' +
    '• Spalten "Status", "Notizen" auf Anmeldungen.\n' +
    '• "Angemeldete" wird jetzt automatisch berechnet (COUNTIFS-Formel).\n' +
    '• Datum, Pflichtfelder und Status werden in Echtzeit validiert.\n' +
    '• Volle/abgesagte/vergangene Anlässe werden farblich markiert.\n' +
    '• Tab "Anleitung" gibt eine Schritt-für-Schritt-Hilfe.' +
    triggerStatus + '\n\n' +
    'Die Aktion ist idempotent – Sie können sie jederzeit erneut ausführen.'
  );
}

/**
 * Make sure a sheet's header row matches the canonical headers. Adds
 * missing columns at the right of the existing row (without disturbing
 * data) and writes any updated header text.
 */
function ensureHeaders(sheet, headers, bg) {
  var maxCol = sheet.getMaxColumns();
  if (maxCol < headers.length) {
    sheet.insertColumnsAfter(maxCol, headers.length - maxCol);
  }
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setBackground(bg).setFontColor('white').setFontWeight('bold');
  sheet.setFrozenRows(1);
}

/**
 * Back-fill Anlässe defaults: Status=aktiv, Jahr from Datum,
 * Notizen blank. Existing values are preserved – we only fill blanks.
 */
function backfillAnlassDefaults(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  var range = sheet.getRange(2, 1, lastRow - 1, ANLASS_HEADERS.length);
  var values = range.getValues();
  for (var i = 0; i < values.length; i++) {
    if (!values[i][0]) continue; // skip blank rows
    if (!values[i][7]) values[i][7] = 'aktiv';      // Status
    if (!values[i][8]) {                              // Jahr
      var d = parseDate(values[i][2]);
      if (d) values[i][8] = jahrFor(d);
    }
    if (values[i][10] == null) values[i][10] = '';   // Notizen
  }
  range.setValues(values);
}

/**
 * Sort the data range of Anmeldungen by Anlass-ID then Zeitstempel.
 * Pure ordering — no values change. Used after every registration
 * (inside the existing lock so it's race-safe), inside
 * setupVerstaerken / zaehlerNeuBerechnen, and from the menu entry
 * so admins can re-sort manually after editing rows by hand.
 *
 * Lexicographic sort on the ID column is acceptable here because
 * the goal is *grouping*, not numeric order: rows for the same
 * event end up adjacent regardless. If a strict numeric sort
 * becomes important later, it'd take a programmatic
 * read-sort-write pass; currently not worth the complication.
 */
function sortAnmeldungenByAnlass(sheet) {
  if (!sheet) return;
  var lastRow = sheet.getLastRow();
  if (lastRow < 3) return; // header + at most one row → nothing to sort
  var lastCol = Math.max(sheet.getLastColumn(), ANMELDUNG_HEADERS.length);
  sheet.getRange(2, 1, lastRow - 1, lastCol).sort([
    { column: 2, ascending: true }, // Anlass-ID
    { column: 1, ascending: true }  // Zeitstempel
  ]);
}

/** Menu wrapper around sortAnmeldungenByAnlass with user-visible feedback. */
function sortiereAnmeldungenMenu() {
  var ui = SpreadsheetApp.getUi();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Anmeldungen');
  if (!sheet) {
    ui.alert('Tabellenblatt "Anmeldungen" fehlt.');
    return;
  }
  var lock = LockService.getScriptLock();
  try {
    if (!lock.tryLock(10000)) {
      ui.alert('Server gerade ausgelastet. Bitte gleich erneut versuchen.');
      return;
    }
    sortAnmeldungenByAnlass(sheet);
    logAudit('SORTIERT_NACH_ANLASS', { rows: Math.max(sheet.getLastRow() - 1, 0) }, true, null);
    ui.alert('✅ Anmeldungen nach Anlass-ID sortiert.');
  } finally {
    try { if (lock.hasLock()) lock.releaseLock(); } catch (_) {}
  }
}

/**
 * Walk every Anmeldungen + "Archiv Anmeldungen *" sheet and rewrite
 * every Telefon cell through normalizePhone. Idempotent: cells whose
 * normalised form already matches the stored value are skipped, so
 * re-running is cheap.
 *
 * Returns { scanned, changed } counts. Used both interactively (menu
 * item) and as part of setupVerstaerken.
 */
function normalisiereAlleTelefonnummern() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var scanned = 0, changed = 0;
  ss.getSheets().forEach(function(sheet) {
    var name = sheet.getName();
    if (name !== 'Anmeldungen' && name.indexOf('Archiv Anmeldungen') !== 0) return;
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return;
    // Telefon = column E (index 5).
    var range = sheet.getRange(2, 5, lastRow - 1, 1);
    var values = range.getValues();
    var dirty = false;
    var newValues = values.map(function(r) {
      var current = r[0];
      if (current === null || current === undefined || current === '') return r;
      var asString = String(current);
      scanned++;
      var normalized = normalizePhone(asString);
      if (normalized !== asString) {
        dirty = true;
        changed++;
        return [sheetSafe(normalized)];
      }
      return r;
    });
    if (dirty) range.setValues(newValues);
  });
  return { scanned: scanned, changed: changed };
}

/** Menu wrapper: runs the bulk migrator with a lock and a summary alert. */
function normalisiereTelefonnummernMenu() {
  var ui = SpreadsheetApp.getUi();
  var lock = LockService.getScriptLock();
  try {
    if (!lock.tryLock(15000)) {
      ui.alert('Server gerade ausgelastet. Bitte gleich erneut versuchen.');
      return;
    }
    var stats = normalisiereAlleTelefonnummern();
    logAudit('PHONES_NORMALIZED', stats, true, null);
    ui.alert(
      '✅ Telefonnummern normalisiert.\n\n' +
      stats.changed + ' von ' + stats.scanned + ' Einträgen wurden ins ' +
      'kanonische Format "+41 XX XXX XX XX" umgeschrieben.'
    );
  } catch (e) {
    logAudit('PHONES_NORMALIZED', {}, false, String(e));
    ui.alert('Fehler: ' + e);
  } finally {
    try { if (lock.hasLock()) lock.releaseLock(); } catch (_) {}
  }
}

/**
 * Live autocorrect of phone numbers typed directly into Anmeldungen.
 * Apps Script wires onEdit() automatically as a "simple trigger" — no
 * registration needed. Filters fast (sheet name + column index check)
 * so the cost on unrelated edits is negligible.
 *
 * sheetSafe() is used when writing back so the leading "+" doesn't
 * trip Sheets' formula-parsing (which historically corrupted "+41"
 * numbers into "#ERROR!").
 */
function onEdit(e) {
  if (!e || !e.range) return;
  var sheet = e.range.getSheet();
  var name = sheet.getName();
  if (name !== 'Anmeldungen' && name.indexOf('Archiv Anmeldungen') !== 0) return;
  if (e.range.getColumn() !== 5) return; // Telefon column only
  if (e.range.getRow() < 2) return;       // skip header
  // Multi-cell paste: e.range may span rows. Re-normalise every cell
  // in the affected range that sits in column 5.
  var height = e.range.getNumRows();
  var values = e.range.getValues();
  var changedAny = false;
  for (var i = 0; i < height; i++) {
    var raw = values[i][0];
    if (raw === null || raw === undefined || raw === '') continue;
    var asString = String(raw);
    var normalized = normalizePhone(asString);
    if (normalized !== asString) {
      values[i][0] = sheetSafe(normalized);
      changedAny = true;
    }
  }
  if (changedAny) e.range.setValues(values);
}

function backfillAnmeldungDefaults(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  var range = sheet.getRange(2, 7, lastRow - 1, 2); // Status, Notizen
  var values = range.getValues();
  for (var i = 0; i < values.length; i++) {
    if (!values[i][0]) values[i][0] = 'aktiv';
    if (values[i][1] == null) values[i][1] = '';
  }
  range.setValues(values);
}

/**
 * Write per-row formulas for Angemeldete (col 6) and Sichtbar? (col 10)
 * across every populated row. Uses setFormulas in two batches.
 */
function writeAnlassFormulas(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  var angemeldete = [];
  var sichtbar = [];
  for (var r = 2; r <= lastRow; r++) {
    angemeldete.push([angemeldeteFormulaForRow(r)]);
    sichtbar.push([sichtbarFormulaForRow(r)]);
  }
  sheet.getRange(2, 6, angemeldete.length, 1).setFormulas(angemeldete);
  sheet.getRange(2, 10, sichtbar.length, 1).setFormulas(sichtbar);
}

function applyAnlassValidations(sheet) {
  var maxRow = Math.max(sheet.getMaxRows() - 1, 1000);
  // Datum (C): real date, reject everything else
  var dateRule = SpreadsheetApp.newDataValidation()
    .requireDate().setAllowInvalid(false)
    .setHelpText('Bitte ein gültiges Datum auswählen (TT.MM.JJJJ).')
    .build();
  sheet.getRange(2, 3, maxRow, 1).setDataValidation(dateRule);

  // Benötigte Helfer (E): integer ≥ 1
  var helferRule = SpreadsheetApp.newDataValidation()
    .requireNumberGreaterThanOrEqualTo(1).setAllowInvalid(false)
    .setHelpText('Mindestens 1 Helfer erforderlich.')
    .build();
  sheet.getRange(2, 5, maxRow, 1).setDataValidation(helferRule);

  // Status (H): dropdown
  var statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(ANLASS_STATUS_VALUES, true).setAllowInvalid(false)
    .setHelpText('Status: aktiv, abgesagt oder archiviert.')
    .build();
  sheet.getRange(2, 8, maxRow, 1).setDataValidation(statusRule);

  // ID (A): unique. Custom-formula validation; reject duplicates.
  var uniqueRule = SpreadsheetApp.newDataValidation()
    .requireFormulaSatisfied('=COUNTIF($A$2:$A,A2)<2').setAllowInvalid(false)
    .setHelpText('Diese Anlass-ID wird bereits verwendet.')
    .build();
  sheet.getRange(2, 1, maxRow, 1).setDataValidation(uniqueRule);
}

function applyAnmeldungValidations(sheet) {
  var maxRow = Math.max(sheet.getMaxRows() - 1, 1000);
  var statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(ANMELDUNG_STATUS_VALUES, true).setAllowInvalid(false)
    .setHelpText('Status: aktiv, storniert oder nicht erschienen.')
    .build();
  sheet.getRange(2, 7, maxRow, 1).setDataValidation(statusRule);
}

/**
 * Conditional formatting rules on Anlässe. Order matters: later rules
 * override earlier ones for the same cell. We want past/cancelled to
 * dominate full-but-active.
 */
function applyAnlassConditionalFormatting(sheet) {
  var range = sheet.getRange(2, 1, Math.max(sheet.getMaxRows() - 1, 1000), ANLASS_HEADERS.length);

  var rules = [];

  // Almost full (>= 50% but not full): subtle yellow
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND($A2<>"",$E2>0,$F2/$E2>=0.5,$F2<$E2)')
    .setBackground('#fef3c7').setRanges([range]).build());

  // Full: muted green
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND($A2<>"",$E2>0,$F2>=$E2)')
    .setBackground('#d1fae5').setRanges([range]).build());

  // Past events: light grey + strikethrough
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND($A2<>"",ISNUMBER($C2),$C2<TODAY())')
    .setBackground('#f1f5f9').setFontColor('#94a3b8').setStrikethrough(true)
    .setRanges([range]).build());

  // Cancelled: red text overrides everything visually
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$H2="abgesagt"')
    .setFontColor('#dc2626').setBold(true)
    .setRanges([range]).build());

  // Archived: light italics, very subtle
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$H2="archiviert"')
    .setFontColor('#64748b').setItalic(true)
    .setRanges([range]).build());

  sheet.setConditionalFormatRules(rules);
}

/**
 * Protect the header row with warning level – the script (running as
 * the spreadsheet owner) can still write, but a human accidentally
 * editing a header gets a confirmation dialog. Header rename was the
 * single most damaging mistake the old schema permitted: it silently
 * broke the public site without any error.
 */
function protectHeaderRow(sheet) {
  var existing = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
  for (var i = 0; i < existing.length; i++) {
    if (existing[i].getDescription() === 'Header (nicht ändern)') {
      existing[i].remove();
    }
  }
  var protection = sheet.getRange(1, 1, 1, sheet.getMaxColumns()).protect()
    .setDescription('Header (nicht ändern)');
  protection.setWarningOnly(true);
}

/**
 * Create or refresh the Anleitung tab. Keeps it as the leftmost sheet
 * so newcomers see it first.
 */
function ensureAnleitungSheet(ss) {
  var sheet = ss.getSheetByName('Anleitung') || ss.insertSheet('Anleitung', 0);
  if (sheet.getIndex() !== 1) ss.setActiveSheet(sheet) && ss.moveActiveSheet(1);
  sheet.clear();
  sheet.setHiddenGridlines(true);
  var lines = [
    ['🏰 Schulhelfer Rittergasse – Anleitung'],
    [''],
    ['Live-Seite:'],
    ['https://ps-rittergasse.github.io/helferliste/'],
    [''],
    ['ANLÄSSE – Spalten'],
    ['  ID                Wird automatisch vergeben (Menü → Neuer Anlass).'],
    ['  Name              Kurz und klar.'],
    ['  Datum             Echtes Datum (das Sheet validiert).'],
    ['  Zeit              Frei: "14:00-18:00", "ab 16:00", "morgens / abends"…'],
    ['  Benötigte Helfer  Mindestens 1.'],
    ['  Angemeldete       AUTOMATISCH (Formel) – nicht überschreiben.'],
    ['  Beschreibung      Was sollen die Helfer tun?'],
    ['  Status            aktiv (öffentlich) / abgesagt / archiviert.'],
    ['  Jahr              AUTOMATISCH aus dem Datum (Kalenderjahr, z.B. "2025").'],
    ['  Sichtbar?         AUTOMATISCH – zeigt, warum ein Anlass evtl. ausgeblendet ist.'],
    ['  Notizen (intern)  Nur intern, wird nicht öffentlich angezeigt.'],
    [''],
    ['EINEN ANLASS ABSAGEN'],
    ['  Status auf "abgesagt" setzen — die Zeile NICHT löschen. So bleiben'],
    ['  die Anmeldungen erhalten und Sie können die Helfer kontaktieren.'],
    [''],
    ['ANMELDUNGEN'],
    ['  Status "storniert" für abgemeldete Helfer (Slot wird wieder frei).'],
    ['  Status "nicht erschienen" für Nachverfolgung nach dem Anlass.'],
    [''],
    ['ARCHIVIERUNG'],
    ['  Anlässe abgeschlossener Kalenderjahre (Jahr < aktuelles Jahr) werden täglich'],
    ['  automatisch um 03:00 in eigene "Archiv …"-Tabs verschoben (schreibgeschützt).'],
    ['  Manuell auslösbar über Menü → "Jahr archivieren…".'],
    [''],
    ['BEI PROBLEMEN'],
    ['  Menü → "Setup verstärken": stellt Validierung & Formeln wieder her.'],
    ['  Menü → "Zähler neu berechnen": nur bei sehr alten Sheets nötig.']
  ];
  sheet.getRange(1, 1, lines.length, 1).setValues(lines);
  sheet.getRange(1, 1).setFontSize(18).setFontWeight('bold').setFontColor('#1e3a5f');
  [6, 19, 23, 27, 32].forEach(function(r) {
    sheet.getRange(r, 1).setFontWeight('bold').setFontSize(12).setFontColor('#1e3a5f');
  });
  sheet.setColumnWidth(1, 760);
}

// =============================================================
// === Year-end archive (PR 2) =================================
// =============================================================
//
// Per Swiss school year (Aug → Jul) we move events and their
// registrations into dedicated "Archiv …" tabs. The live sheets stay
// small year over year. Archive tabs are warning-protected so admins
// don't accidentally edit historical records.
//
// Design invariants:
//   1. Source-of-truth IDs never repeat. nextAnlassId() scans the
//      live sheet AND every archive sheet, guaranteeing iCal UIDs and
//      audit-log refs stay unambiguous forever.
//   2. The archive function copies-then-deletes. A crash mid-run
//      leaves duplicates in the archive on retry, never lost data.
//   3. LockService blocks concurrent registrations during archive.

/**
 * Show a dialog letting the admin pick a school year (with the most
 * recent past year as default), confirm intent, and trigger the move.
 */
function archiviereJahrDialog() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Anlässe');
  if (!sheet) {
    ui.alert('Tabellenblatt "Anlässe" fehlt.');
    return;
  }

  var heuteYear = jahrFor(new Date());
  var available = availableJahre(sheet)
    .filter(function(y) { return y && y !== heuteYear; });

  if (available.length === 0) {
    ui.alert(
      'Jahr archivieren',
      'Es gibt aktuell keine abgeschlossenen Jahre zum Archivieren.\n\n' +
      '(Das laufende Jahr ' + heuteYear + ' kann nicht archiviert werden.)',
      ui.ButtonSet.OK
    );
    return;
  }

  // Build the dialog HTML. Default option = most recent past year.
  var options = available.map(function(y) {
    var stats = jahrStats(y);
    var label = y + '  (' + stats.events + ' Anlass/Anlässe, ' + stats.registrations + ' Anmeldungen)';
    return { value: y, label: label };
  });
  var defaultYear = options[options.length - 1].value;

  var html =
    '<style>' +
    '  body { font-family: -apple-system, system-ui, Arial, sans-serif; padding: 16px; color: #1e293b; }' +
    '  h2 { margin: 0 0 8px 0; font-size: 18px; color: #1e3a5f; }' +
    '  p { margin: 8px 0; line-height: 1.5; }' +
    '  label { display: block; margin-top: 12px; font-weight: bold; }' +
    '  select { width: 100%; padding: 10px; margin-top: 4px; border: 1px solid #cbd5e1; border-radius: 6px; font-size: 14px; }' +
    '  .checkbox-row { margin-top: 14px; display: flex; align-items: flex-start; gap: 8px; font-size: 13px; }' +
    '  .buttons { margin-top: 18px; display: flex; gap: 8px; justify-content: flex-end; }' +
    '  button { padding: 10px 18px; border: none; border-radius: 6px; cursor: pointer; font-size: 14px; font-weight: 600; }' +
    '  .primary { background: #1e3a5f; color: white; }' +
    '  .primary:disabled { background: #cbd5e1; cursor: not-allowed; }' +
    '  .secondary { background: #e2e8f0; color: #1e293b; }' +
    '  .summary { margin-top: 12px; padding: 10px; background: #f8fafc; border-radius: 6px; font-size: 13px; }' +
    '</style>' +
    '<h2>Jahr archivieren</h2>' +
    '<p>Anlässe und Anmeldungen des gewählten Jahrs werden in dedizierte Archiv-Tabs verschoben und dort schreibgeschützt.</p>' +
    '<label>Jahr</label>' +
    '<select id="year">' +
    options.map(function(o) {
      return '<option value="' + o.value + '"' + (o.value === defaultYear ? ' selected' : '') + '>' + o.label + '</option>';
    }).join('') +
    '</select>' +
    '<div class="summary">Die IDs der Anlässe bleiben weltweit eindeutig – auch in zukünftigen Jahren werden archivierte IDs nicht wiederverwendet.</div>' +
    '<div class="checkbox-row">' +
    '  <input type="checkbox" id="confirm">' +
    '  <label for="confirm" style="font-weight: normal; margin: 0;">Ich habe geprüft, dass ich für dieses Jahr nichts mehr ändern muss.</label>' +
    '</div>' +
    '<div class="buttons">' +
    '  <button class="secondary" type="button" onclick="google.script.host.close()">Abbrechen</button>' +
    '  <button class="primary" type="button" id="go" disabled onclick="run()">Archivieren</button>' +
    '</div>' +
    '<script>' +
    '  var c = document.getElementById("confirm");' +
    '  var b = document.getElementById("go");' +
    '  c.addEventListener("change", function(){ b.disabled = !c.checked; });' +
    '  function run(){' +
    '    b.disabled = true; b.textContent = "Wird verschoben…";' +
    '    google.script.run' +
    '      .withSuccessHandler(function(res){' +
    '        if (res && res.success) { alert("✅ " + res.message); google.script.host.close(); }' +
    '        else { alert("Fehler: " + (res && res.message || "Unbekannter Fehler")); b.disabled = false; b.textContent = "Archivieren"; }' +
    '      })' +
    '      .withFailureHandler(function(err){' +
    '        alert("Fehler: " + err); b.disabled = false; b.textContent = "Archivieren";' +
    '      })' +
    '      .archiviereJahr(document.getElementById("year").value);' +
    '  }' +
    '</script>';

  var output = HtmlService.createHtmlOutput(html).setWidth(440).setHeight(360);
  ui.showModalDialog(output, '🗂️ Jahr archivieren');
}

/** Return the unique non-empty Jahr values in Anlässe, sorted. */
function availableJahre(anlassSheet) {
  var lastRow = anlassSheet.getLastRow();
  if (lastRow < 2) return [];
  var values = anlassSheet.getRange(2, 9, lastRow - 1, 1).getValues();
  var set = {};
  for (var i = 0; i < values.length; i++) {
    var v = String(values[i][0] || '').trim();
    if (v) set[v] = true;
  }
  // For rows missing Jahr (un-migrated), derive from Datum.
  var allValues = anlassSheet.getRange(2, 1, lastRow - 1, 9).getValues();
  for (var j = 0; j < allValues.length; j++) {
    if (!allValues[j][0]) continue;
    var sj = String(allValues[j][8] || '').trim();
    if (!sj) {
      var d = parseDate(allValues[j][2]);
      if (d) {
        var derived = jahrFor(d);
        if (derived) set[derived] = true;
      }
    }
  }
  return Object.keys(set).sort();
}

/** Count how many events + registrations would move for a given year. */
function jahrStats(jahr) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var anlassSheet = ss.getSheetByName('Anlässe');
  var helferSheet = ss.getSheetByName('Anmeldungen');
  if (!anlassSheet || !helferSheet) return { events: 0, registrations: 0 };

  var ids = collectAnlassIdsForYear(anlassSheet, jahr);
  var registrations = 0;
  var lastHelfer = helferSheet.getLastRow();
  if (lastHelfer >= 2) {
    var helferData = helferSheet.getRange(2, 1, lastHelfer - 1, 2).getValues();
    for (var i = 0; i < helferData.length; i++) {
      if (ids.indexOf(String(helferData[i][1] || '')) !== -1) registrations++;
    }
  }
  return { events: ids.length, registrations: registrations };
}

/**
 * Find every Anlass-ID belonging to the given school year. Uses the
 * Jahr column when present and falls back to deriving from Datum
 * for un-migrated rows.
 */
function collectAnlassIdsForYear(anlassSheet, jahr) {
  var lastRow = anlassSheet.getLastRow();
  if (lastRow < 2) return [];
  var values = anlassSheet.getRange(2, 1, lastRow - 1, ANLASS_HEADERS.length).getValues();
  var ids = [];
  for (var i = 0; i < values.length; i++) {
    if (!values[i][0]) continue;
    var rowSj = String(values[i][8] || '').trim();
    if (!rowSj) {
      var d = parseDate(values[i][2]);
      if (d) rowSj = jahrFor(d);
    }
    if (rowSj === jahr) ids.push(String(values[i][0]));
  }
  return ids;
}

/**
 * The actual archive operation. Called from the dialog.
 *
 * Returns { success, message } – the dialog displays the message via
 * an alert(). Errors are caught and surfaced with a friendly message.
 */
function archiviereJahr(jahr) {
  jahr = String(jahr || '').trim();
  if (!jahr) return { success: false, message: 'Kein Jahr angegeben.' };
  var current = jahrFor(new Date());
  if (jahr === current) {
    return { success: false, message: 'Das laufende Jahr kann nicht archiviert werden.' };
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var anlassSheet = ss.getSheetByName('Anlässe');
  var helferSheet = ss.getSheetByName('Anmeldungen');
  if (!anlassSheet || !helferSheet) {
    return { success: false, message: 'Tabellenblätter fehlen.' };
  }

  var lock = LockService.getScriptLock();
  try {
    if (!lock.tryLock(15000)) {
      return { success: false, message: 'Server gerade ausgelastet. Bitte gleich erneut versuchen.' };
    }

    var anlassIds = collectAnlassIdsForYear(anlassSheet, jahr);
    if (anlassIds.length === 0) {
      return { success: false, message: 'Keine Anlässe für ' + jahr + ' gefunden.' };
    }

    // Slug used for the archive sheet name. Sheet names can contain "/"
    // but it confuses some downstream tools, so use "-".
    var slug = jahr.replace('/', '-');
    var anlassArchiveName = 'Archiv Anlässe ' + slug;
    var helferArchiveName = 'Archiv Anmeldungen ' + slug;

    // Find which Anlässe + Anmeldungen rows to move.
    var anlassLast = anlassSheet.getLastRow();
    var anlassValues = anlassSheet.getRange(2, 1, anlassLast - 1, ANLASS_HEADERS.length).getValues();
    var anlassToMove = [];
    var anlassRowsToDelete = [];
    for (var i = 0; i < anlassValues.length; i++) {
      if (!anlassValues[i][0]) continue;
      if (anlassIds.indexOf(String(anlassValues[i][0])) !== -1) {
        anlassToMove.push(anlassValues[i]);
        anlassRowsToDelete.push(i + 2); // 1-based row, header on row 1
      }
    }

    var helferLast = helferSheet.getLastRow();
    var helferToMove = [];
    var helferRowsToDelete = [];
    if (helferLast >= 2) {
      var helferValues = helferSheet.getRange(2, 1, helferLast - 1, ANMELDUNG_HEADERS.length).getValues();
      for (var j = 0; j < helferValues.length; j++) {
        if (anlassIds.indexOf(String(helferValues[j][1] || '')) !== -1) {
          helferToMove.push(helferValues[j]);
          helferRowsToDelete.push(j + 2);
        }
      }
    }

    // --- Copy phase (idempotent on retry) ---
    var anlassArchive = ss.getSheetByName(anlassArchiveName);
    var helferArchive = ss.getSheetByName(helferArchiveName);
    var freshAnlassArchive = false;
    var freshHelferArchive = false;
    if (!anlassArchive) {
      anlassArchive = ss.insertSheet(anlassArchiveName);
      anlassArchive.getRange(1, 1, 1, ANLASS_HEADERS.length).setValues([ANLASS_HEADERS]);
      anlassArchive.getRange(1, 1, 1, ANLASS_HEADERS.length)
        .setBackground('#475569').setFontColor('white').setFontWeight('bold');
      anlassArchive.setFrozenRows(1);
      freshAnlassArchive = true;
    }
    if (!helferArchive) {
      helferArchive = ss.insertSheet(helferArchiveName);
      helferArchive.getRange(1, 1, 1, ANMELDUNG_HEADERS.length).setValues([ANMELDUNG_HEADERS]);
      helferArchive.getRange(1, 1, 1, ANMELDUNG_HEADERS.length)
        .setBackground('#475569').setFontColor('white').setFontWeight('bold');
      helferArchive.setFrozenRows(1);
      freshHelferArchive = true;
    }

    // Idempotency on retry: a previous run may have copied (some)
    // rows already if it crashed mid-way. Skip rows whose key is
    // already in the archive so we never double-write the same data.
    //
    //   Anlässe key   = ID (column A)
    //   Anmeldungen key = (Anlass-ID, lower(email)) — registriereHelfer
    //                     enforces this pair unique on the live sheet,
    //                     so it's also unique within a year archive.
    var existingAnlassIds = readArchiveKeySet(anlassArchive, function(row) {
      return String(row[0] || '').trim();
    });
    var existingHelferKeys = readArchiveKeySet(helferArchive, function(row) {
      var aid = String(row[1] || '').trim();
      var em  = String(row[3] || '').trim().toLowerCase();
      return aid && em ? aid + '|' + em : '';
    });

    var anlassToWrite = anlassToMove.filter(function(row) {
      var id = String(row[0] || '').trim();
      return id && !existingAnlassIds[id];
    });
    var helferToWrite = helferToMove.filter(function(row) {
      var aid = String(row[1] || '').trim();
      var em  = String(row[3] || '').trim().toLowerCase();
      var key = aid && em ? aid + '|' + em : '';
      return key && !existingHelferKeys[key];
    });

    if (anlassToWrite.length) {
      var startRow = anlassArchive.getLastRow() + 1;
      anlassArchive.getRange(startRow, 1, anlassToWrite.length, ANLASS_HEADERS.length).setValues(anlassToWrite);
    }
    if (helferToWrite.length) {
      var startRow2 = helferArchive.getLastRow() + 1;
      helferArchive.getRange(startRow2, 1, helferToWrite.length, ANMELDUNG_HEADERS.length).setValues(helferToWrite);
    }

    // --- Delete phase (bottom-up to keep indices valid) ---
    helferRowsToDelete.sort(function(a, b) { return b - a; });
    for (var k = 0; k < helferRowsToDelete.length; k++) helferSheet.deleteRow(helferRowsToDelete[k]);
    anlassRowsToDelete.sort(function(a, b) { return b - a; });
    for (var l = 0; l < anlassRowsToDelete.length; l++) anlassSheet.deleteRow(anlassRowsToDelete[l]);

    // --- Protect archive sheets (warning-only so admins can fix typos
    // without needing to re-run anything; full protection would block
    // even the script owner without re-acknowledgement). ---
    if (freshAnlassArchive) protectArchiveSheet(anlassArchive, jahr);
    if (freshHelferArchive) protectArchiveSheet(helferArchive, jahr);

    logAudit('SCHULJAHR_ARCHIVIERT', {
      jahr: jahr,
      anlaesse: anlassToMove.length,
      anmeldungen: helferToMove.length
    }, true, null);

    return {
      success: true,
      message: 'Jahr ' + jahr + ' archiviert: ' +
               anlassToMove.length + ' Anlass/Anlässe, ' +
               helferToMove.length + ' Anmeldungen.\n\n' +
               'Tabs: "' + anlassArchiveName + '" und "' + helferArchiveName + '".'
    };
  } catch (e) {
    logAudit('SCHULJAHR_ARCHIVIERT', { jahr: jahr }, false, String(e));
    return { success: false, message: String(e) };
  } finally {
    try { if (lock.hasLock()) lock.releaseLock(); } catch (_) {}
  }
}

/**
 * Build a {key: true} map of every existing archive row so the
 * archive function can skip duplicates on a retry. Returns an empty
 * map for nonexistent / empty sheets.
 */
function readArchiveKeySet(sheet, keyFn) {
  var set = {};
  if (!sheet) return set;
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return set;
  var lastCol = Math.max(sheet.getLastColumn(), 1);
  var values = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  for (var i = 0; i < values.length; i++) {
    var k = keyFn(values[i]);
    if (k) set[k] = true;
  }
  return set;
}

function protectArchiveSheet(sheet, jahr) {
  var existing = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
  for (var i = 0; i < existing.length; i++) existing[i].remove();
  var p = sheet.protect()
    .setDescription('Archiv ' + jahr + ' – schreibgeschützt');
  p.setWarningOnly(true);
}

// =============================================================
// === Auto-archive (daily time-driven trigger) ================
// =============================================================
//
// On Jan 1 the previous year's events are eligible to be archived.
// A daily trigger runs autoArchive() at 03:00 (script timezone).
// It scans live Anlässe for any Jahr value strictly less than the
// current calendar year and moves each year through archiviereJahr.
//
// Idempotent: if no past-year rows remain, it's a no-op. Safe to
// run repeatedly.
//
// installiereAutoArchiv() creates the daily trigger and is invoked
// from setupVerstaerken so admins get auto-archive for free without
// having to know about Apps Script triggers.

var AUTO_ARCHIVE_FN = 'autoArchive';

function autoArchive() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Anlässe');
  if (!sheet) return;
  var currentYear = new Date().getFullYear();
  var jahre = availableJahre(sheet)
    .filter(function(y) {
      var n = parseInt(y, 10);
      return !isNaN(n) && n < currentYear;
    });
  jahre.forEach(function(y) {
    var result = archiviereJahr(y);
    logAudit('AUTO_ARCHIVE', { jahr: y, success: !!result.success },
             !!result.success, result.message);
  });
}

/**
 * Idempotent install of the daily auto-archive trigger. Removes any
 * existing autoArchive triggers first so re-running doesn't pile up
 * duplicates. Time of day is 03:00 in the script's timezone — late
 * enough that the previous day's edits have settled, early enough
 * that admins arriving in the morning see the result.
 */
function installiereAutoArchiv() {
  var existing = ScriptApp.getProjectTriggers();
  for (var i = 0; i < existing.length; i++) {
    if (existing[i].getHandlerFunction() === AUTO_ARCHIVE_FN) {
      ScriptApp.deleteTrigger(existing[i]);
    }
  }
  ScriptApp.newTrigger(AUTO_ARCHIVE_FN)
    .timeBased()
    .atHour(3)
    .everyDays(1)
    .create();
}

// =============================================================
// === Integrity check (PR 2) ==================================
// =============================================================
//
// Surfaces every condition that can silently break the public site:
// duplicate event IDs, orphan registrations, formula errors, events
// in the live sheet that should have been archived, etc. Reports
// counts and row references; never auto-fixes. Human judgment owns
// the corrections.

function integritaetspruefung() {
  var ui = SpreadsheetApp.getUi();
  var report = buildIntegrityReport();
  if (report.success === false) {
    ui.alert(report.error || 'Tabellenblätter fehlen.');
    return;
  }
  logAudit('INTEGRITAETSPRUEFUNG', { findings: report.findings.length }, true, null);
  if (report.findings.length === 0) {
    ui.alert('✅ Datenprüfung', 'Alles sauber. Keine Auffälligkeiten gefunden.', ui.ButtonSet.OK);
  } else {
    ui.alert('🔍 Datenprüfung – ' + report.findings.length + ' Befund(e)',
             report.findings.join('\n\n'), ui.ButtonSet.OK);
  }
}

/**
 * Pure-data variant of the integrity check. Used by both the in-Sheet
 * menu (which alerts) and the admin POST endpoint (which serialises
 * the result into JSON for the web UI). Never touches the UI.
 */
function buildIntegrityReport() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var anlassSheet = ss.getSheetByName('Anlässe');
  var helferSheet = ss.getSheetByName('Anmeldungen');
  if (!anlassSheet || !helferSheet) {
    return { success: false, error: 'Tabellenblätter fehlen.' };
  }

  var report = [];
  var anlassLast = anlassSheet.getLastRow();
  var anlassData = anlassLast >= 2
    ? anlassSheet.getRange(2, 1, anlassLast - 1, ANLASS_HEADERS.length).getValues()
    : [];

  // 1. Duplicate IDs in live Anlässe.
  var idSeen = {};
  var duplicateIds = [];
  for (var i = 0; i < anlassData.length; i++) {
    var id = String(anlassData[i][0] || '').trim();
    if (!id) continue;
    if (idSeen[id]) duplicateIds.push({ id: id, row: i + 2 });
    else idSeen[id] = true;
  }
  if (duplicateIds.length) {
    report.push('⚠️  Doppelte IDs (' + duplicateIds.length + '): ' +
      duplicateIds.map(function(d) { return d.id + ' (Zeile ' + d.row + ')'; }).join(', '));
  }

  // 2. Invalid dates.
  var invalidDates = [];
  for (var j = 0; j < anlassData.length; j++) {
    if (!anlassData[j][0]) continue;
    if (!parseDate(anlassData[j][2])) invalidDates.push(j + 2);
  }
  if (invalidDates.length) {
    report.push('⚠️  Ungültiges Datum in Zeilen: ' + invalidDates.join(', '));
  }

  // 3. Past events still active (date older than 60 days, status still 'aktiv').
  var cutoff = new Date(); cutoff.setDate(cutoff.getDate() - 60); cutoff.setHours(0, 0, 0, 0);
  var oldActive = [];
  for (var k = 0; k < anlassData.length; k++) {
    if (!anlassData[k][0]) continue;
    var d = parseDate(anlassData[k][2]);
    if (!d || d >= cutoff) continue;
    var status = String(anlassData[k][7] || 'aktiv').trim().toLowerCase();
    if (status === 'aktiv') oldActive.push({ id: anlassData[k][0], row: k + 2 });
  }
  if (oldActive.length) {
    report.push('ℹ️  ' + oldActive.length + ' alte Anlässe (>60 Tage) noch mit Status "aktiv". ' +
      'Erwägen Sie "Jahr archivieren…".');
  }

  // 4. Missing Status / Jahr (un-migrated rows).
  var missingStatus = 0, missingJahr = 0;
  for (var m = 0; m < anlassData.length; m++) {
    if (!anlassData[m][0]) continue;
    if (!String(anlassData[m][7] || '').trim()) missingStatus++;
    if (!String(anlassData[m][8] || '').trim()) missingJahr++;
  }
  if (missingStatus) report.push('ℹ️  ' + missingStatus + ' Anlass-Zeile(n) ohne Status (führen Sie "Setup verstärken" aus).');
  if (missingJahr)   report.push('ℹ️  ' + missingJahr + ' Anlass-Zeile(n) ohne Jahr (führen Sie "Setup verstärken" aus).');

  // 5. Formula errors in the Angemeldete column.
  var formulaErrors = [];
  if (anlassData.length) {
    var fValues = anlassSheet.getRange(2, 6, anlassData.length, 1).getDisplayValues();
    for (var n = 0; n < fValues.length; n++) {
      var v = String(fValues[n][0] || '');
      if (v.charAt(0) === '#') formulaErrors.push(n + 2);
    }
  }
  if (formulaErrors.length) {
    report.push('⚠️  Formel-Fehler in Angemeldete-Spalte (Zeilen ' + formulaErrors.join(', ') + '). "Setup verstärken" stellt die Formel wieder her.');
  }

  // 6. Orphan registrations: Anlass-ID present in Anmeldungen but not in any Anlässe sheet (live or archive).
  var helferLast = helferSheet.getLastRow();
  if (helferLast >= 2) {
    var allKnownIds = {};
    ss.getSheets().forEach(function(s) {
      var name = s.getName();
      if (name !== 'Anlässe' && name.indexOf('Archiv Anlässe') !== 0) return;
      var sLast = s.getLastRow();
      if (sLast < 2) return;
      s.getRange(2, 1, sLast - 1, 1).getValues().forEach(function(r) {
        if (r[0]) allKnownIds[String(r[0])] = true;
      });
    });
    var helferIds = helferSheet.getRange(2, 2, helferLast - 1, 1).getValues();
    var orphans = [];
    for (var p = 0; p < helferIds.length; p++) {
      var hid = String(helferIds[p][0] || '').trim();
      if (hid && !allKnownIds[hid]) orphans.push({ id: hid, row: p + 2 });
    }
    if (orphans.length) {
      report.push('⚠️  Verwaiste Anmeldungen (' + orphans.length + '): IDs ' +
        orphans.slice(0, 5).map(function(o) { return o.id + ' (Z. ' + o.row + ')'; }).join(', ') +
        (orphans.length > 5 ? '…' : ''));
    }
  }

  // 7. Anmeldungen Status-Spalte (G) leer.
  if (helferLast >= 2) {
    var hStatus = helferSheet.getRange(2, 7, helferLast - 1, 1).getValues();
    var emptyStatus = 0;
    for (var q = 0; q < hStatus.length; q++) {
      if (!String(hStatus[q][0] || '').trim()) emptyStatus++;
    }
    if (emptyStatus) {
      report.push('ℹ️  ' + emptyStatus + ' Anmeldung(en) ohne Status. "Setup verstärken" füllt "aktiv" nach.');
    }
  }

  return { success: true, findings: report };
}


