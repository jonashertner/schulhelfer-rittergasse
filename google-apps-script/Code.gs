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
 *   "+41 79 123 45 67"  → "+41 79 123 45 67"
 *   "+41791234567"      → "+41 79 123 45 67"
 *   "0041 79 123 45 67" → "+41 79 123 45 67"
 *   "0041791234567"     → "+41 79 123 45 67"
 *   "(079) 123-45-67"   → "+41 79 123 45 67"
 *   "061 999 00 00"     → "+41 61 999 00 00"  (Swiss landline)
 *   "+49 30 12345678"   → "+49 30 12345678"    (foreign: untouched format)
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

  // Swiss international format: 41 + 9 subscriber digits = 11 digits
  if (digits.indexOf('41') === 0 && digits.length === 11) {
    return '+41 ' + digits.substr(2, 2) + ' ' + digits.substr(4, 3) +
           ' ' + digits.substr(7, 2) + ' ' + digits.substr(9, 2);
  }

  // Non-Swiss international (e.g. "+49..."): keep as originally typed
  // (trimmed). We don't know the local grouping convention, so losing
  // the spaces would be worse than leaving them alone.
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
      // Data export functionality
      var result = exportData();
      output = JSON.stringify(result);
      logAudit('EXPORT_DATA', {}, result.success, result.error);
    } else if (action === 'getHelferList') {
      // Admin-only: return full registration data for a single event so
      // the frontend can build a .docx Helferliste. Requires ADMIN_KEY.
      var result = getHelferList(e.parameter.eventId, e.parameter.adminKey);
      // Brute-force mitigation: failed auth attempts all share a global
      // bucket independent of the per-identifier GET limit. Attackers
      // rotating identifiers still hit this cap.
      if (!result.success && result.error === 'Keine Berechtigung.') {
        if (!checkRateLimit('admin-fail', ADMIN_BRUTEFORCE_CAP)) {
          result = { success: false, error: 'Zu viele Fehlversuche. Bitte später erneut versuchen.' };
        }
      }
      output = JSON.stringify(result);
      logAudit('GET_HELFER_LIST', { eventId: e.parameter.eventId }, result.success, result.error);
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
        voll: aktuelleHelfer >= maxHelfer
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
    .addItem('Daten exportieren', 'exportDataAsCSV')
    .addSeparator()
    .addItem('Alte Anlässe bereinigen', 'alteAnlaesseBereinigen')
    .addItem('Admin-Status prüfen', 'adminKeyStatus')
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
    logAudit('ZAEHLER_NEU_BERECHNET', { changed: changed }, true, null);
    ui.alert('✅ Zähler neu berechnet.\n\n' + changed + ' Anlass-Zeile(n) aktualisiert.');
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
            beschreibung: document.getElementById("b").value
          });
      };
    </script>
  `).setWidth(380).setHeight(420);
  
  SpreadsheetApp.getUi().showModalDialog(html, '🏰 Neuer Anlass');
}

function anlassHinzufuegen(data) {
  // Sanitize & validate input first (outside the lock).
  var name = sanitizeInput(data.name || '');
  var zeit = sanitizeInput(data.zeit || '');
  var beschreibung = sanitizeInput(data.beschreibung || '');
  var helfer = parseInt(data.helfer) || 0;

  if (!name || helfer < 1) {
    throw new Error('Bitte füllen Sie alle Pflichtfelder aus.');
  }
  if (name.length > MAX_EVENT_NAME_LEN) {
    throw new Error('Der Anlass-Name ist zu lang (max. ' + MAX_EVENT_NAME_LEN + ' Zeichen).');
  }
  if (beschreibung.length > MAX_DESC_LEN) {
    throw new Error('Die Beschreibung ist zu lang (max. ' + MAX_DESC_LEN + ' Zeichen).');
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
    // Status="aktiv" by default; Schuljahr is derived. The COUNTIFS
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
      schuljahrFor(datum),
      '', // Sichtbar? – formula written by setupVerstaerken/refresh
      ''  // Notizen
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
  'Beschreibung', 'Status', 'Schuljahr', 'Sichtbar?', 'Notizen (intern)'
];
var ANMELDUNG_HEADERS = [
  'Zeitstempel', 'Anlass-ID', 'Name', 'E-Mail', 'Telefon', 'Anlass',
  'Status', 'Notizen'
];
var ANLASS_STATUS_VALUES = ['aktiv', 'abgesagt', 'archiviert'];
var ANMELDUNG_STATUS_VALUES = ['aktiv', 'storniert', 'nicht erschienen'];

/**
 * School year for a given date, "YYYY/YY". Swiss school year runs
 * August → July, so August 2025 belongs to 2025/26 and June 2026 still
 * belongs to 2025/26.
 */
function schuljahrFor(date) {
  if (!(date instanceof Date) || isNaN(date.getTime())) return '';
  var y = date.getFullYear();
  var startYear = date.getMonth() >= 7 ? y : y - 1;
  var endShort = ('0' + ((startYear + 1) % 100)).slice(-2);
  return startYear + '/' + endShort;
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

  ui.alert(
    '✅ Setup verstärkt.\n\n' +
    '• Spalten "Status", "Schuljahr", "Sichtbar?", "Notizen" auf Anlässe.\n' +
    '• Spalten "Status", "Notizen" auf Anmeldungen.\n' +
    '• "Angemeldete" wird jetzt automatisch berechnet (COUNTIFS-Formel).\n' +
    '• Datum, Pflichtfelder und Status werden in Echtzeit validiert.\n' +
    '• Volle/abgesagte/vergangene Anlässe werden farblich markiert.\n' +
    '• Tab "Anleitung" gibt eine Schritt-für-Schritt-Hilfe.\n\n' +
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
 * Back-fill Anlässe defaults: Status=aktiv, Schuljahr from Datum,
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
    if (!values[i][8]) {                              // Schuljahr
      var d = parseDate(values[i][2]);
      if (d) values[i][8] = schuljahrFor(d);
    }
    if (values[i][10] == null) values[i][10] = '';   // Notizen
  }
  range.setValues(values);
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
    ['  Schuljahr         AUTOMATISCH aus dem Datum (z.B. "2025/26").'],
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
    ['BEI PROBLEMEN'],
    ['  Menü → "Setup verstärken": stellt Validierung & Formeln wieder her.'],
    ['  Menü → "Zähler neu berechnen": nur bei sehr alten Sheets nötig.']
  ];
  sheet.getRange(1, 1, lines.length, 1).setValues(lines);
  sheet.getRange(1, 1).setFontSize(18).setFontWeight('bold').setFontColor('#1e3a5f');
  [6, 19, 23, 27].forEach(function(r) {
    sheet.getRange(r, 1).setFontWeight('bold').setFontSize(12).setFontColor('#1e3a5f');
  });
  sheet.setColumnWidth(1, 760);
}

