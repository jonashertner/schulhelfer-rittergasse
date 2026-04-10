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
var RATE_LIMIT_MAX_REQUESTS = 10; // max requests per window
var ADMIN_EMAIL = ''; // Set this to receive notifications (optional)

// === Utility Functions ===

/**
 * Sanitize input to prevent XSS
 */
function sanitizeInput(str) {
  if (!str) return '';
  return String(str)
    .replace(/[<>]/g, '')
    .replace(/javascript:/gi, '')
    .replace(/on\w+=/gi, '')
    .trim();
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
 * Parse date safely
 */
function parseDate(dateValue) {
  if (!dateValue) return null;
  try {
    var date = new Date(dateValue);
    if (isNaN(date.getTime())) return null;
    return date;
  } catch (e) {
    return null;
  }
}

/**
 * Check rate limiting
 */
function checkRateLimit(identifier) {
  var props = PropertiesService.getScriptProperties();
  var key = 'rate_' + identifier;
  var now = Math.floor(Date.now() / 1000);
  var windowStart = now - RATE_LIMIT_WINDOW;
  
  var data = props.getProperty(key);
  var requests = data ? JSON.parse(data) : [];
  
  // Remove old requests outside the window
  requests = requests.filter(function(timestamp) {
    return timestamp > windowStart;
  });
  
  if (requests.length >= RATE_LIMIT_MAX_REQUESTS) {
    return false; // Rate limit exceeded
  }
  
  // Add current request
  requests.push(now);
  props.setProperty(key, JSON.stringify(requests));
  return true;
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
    
    // Keep only last 1000 entries
    var lastRow = logSheet.getLastRow();
    if (lastRow > 1001) {
      logSheet.deleteRows(2, lastRow - 1001);
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
  var identifier = e.parameter.identifier || 'anonymous';
  
  try {
    // Rate limiting
    if (!checkRateLimit(identifier)) {
      logAudit('GET_RATE_LIMIT', { action: e.parameter.action }, false, 'Rate limit exceeded');
      output = JSON.stringify({ 
        success: false,
        error: 'Zu viele Anfragen. Bitte versuchen Sie es später erneut.' 
      });
      return ContentService
        .createTextOutput(output)
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    var action = e.parameter.action || 'getEvents';
    
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
    identifier = data.email || 'anonymous';
    
    // Rate limiting
    if (!checkRateLimit(identifier)) {
      logAudit('POST_RATE_LIMIT', { anlassId: data.anlassId }, false, 'Rate limit exceeded');
      output = JSON.stringify({ 
        success: false, 
        message: 'Zu viele Anfragen. Bitte versuchen Sie es später erneut.' 
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

  // Get all registrations to map helper names to events
  var helferSheet = ss.getSheetByName('Anmeldungen');
  var helferByAnlass = {};
  if (helferSheet) {
    var helferData = helferSheet.getDataRange().getValues();
    for (var j = 1; j < helferData.length; j++) {
      var anlassId = String(helferData[j][1] || '');
      var helferName = String(helferData[j][2] || '').trim();
      if (anlassId && helferName) {
        if (!helferByAnlass[anlassId]) {
          helferByAnlass[anlassId] = [];
        }
        helferByAnlass[anlassId].push(sanitizeInput(helferName));
      }
    }
  }

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

    // Normalize dates for comparison
    var datumNormalized = new Date(datum);
    datumNormalized.setHours(0, 0, 0, 0);

    if (datumNormalized >= heute && aktuelleHelfer < maxHelfer) {
      anlaesse.push({
        id: anlassId,
        name: sanitizeInput(row[1]),
        datum: formatDatum(datum),
        datumSort: datumNormalized.getTime(),
        zeit: sanitizeInput(row[3] || ''),
        beschreibung: sanitizeInput(row[6] || ''),
        maxHelfer: maxHelfer,
        aktuelleHelfer: aktuelleHelfer,
        freiePlaetze: maxHelfer - aktuelleHelfer,
        helferNamen: helferByAnlass[anlassId] || []
      });
    }
  }

  // Sort by date ascending (earliest first)
  anlaesse.sort(function(a, b) {
    return a.datumSort - b.datumSort;
  });

  return anlaesse;
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
  var telefon = sanitizeInput(data.telefon || '');
  
  // Validation
  if (!anlassId || !name || !email) {
    return { success: false, message: 'Bitte füllen Sie alle Pflichtfelder aus.' };
  }
  
  if (name.length < 2) {
    return { success: false, message: 'Der Name muss mindestens 2 Zeichen lang sein.' };
  }
  
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
    return { success: false, message: 'Bitte geben Sie eine gültige E-Mail-Adresse ein.' };
  }
  
  // Email length check
  if (email.length > 254) {
    return { success: false, message: 'Die E-Mail-Adresse ist zu lang.' };
  }
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var anlassSheet = ss.getSheetByName('Anlässe');
  var helferSheet = ss.getSheetByName('Anmeldungen');
  
  if (!anlassSheet || !helferSheet) {
    return { success: false, message: 'Der Anlass konnte nicht verarbeitet werden. Bitte versuchen Sie es später erneut.' };
  }
  
  // Use LockService to prevent race conditions
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000); // Wait up to 10 seconds
    
    // Find event
    var anlassData = anlassSheet.getDataRange().getValues();
    var anlassRow = -1, anlassName = '', maxHelfer = 0, aktuelleHelfer = 0;
    
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
      lock.releaseLock();
      return { success: false, message: 'Der gewählte Anlass wurde nicht gefunden.' };
    }
    
    // Re-check capacity after lock (prevent race condition)
    if (aktuelleHelfer >= maxHelfer) {
      lock.releaseLock();
      return { success: false, message: 'Leider sind bereits alle Plätze vergeben.' };
    }
    
    // Improved duplicate check: name + email combination
    var helferData = helferSheet.getDataRange().getValues();
    for (var j = 1; j < helferData.length; j++) {
      var existingEmail = String(helferData[j][3] || '').toLowerCase();
      var existingName = String(helferData[j][2] || '').toLowerCase();
      var existingAnlassId = String(helferData[j][1] || '');
      
      if (existingAnlassId === anlassId &&
          existingEmail === email &&
          existingName === name.toLowerCase()) {
        lock.releaseLock();
        return { success: false, message: 'Sie sind bereits für diesen Anlass angemeldet.' };
      }
    }
    
    // Register (transaction-safe). sheetSafe() prefixes any string
    // starting with a formula trigger (=, +, -, @, tab, CR) with a
    // single quote so Sheets stores it as text. Without this, "+41 79
    // ..." phone numbers are silently parsed as unary-plus formulas
    // and either lose the leading "+" or become #ERROR!.
    helferSheet.appendRow([
      new Date(),
      sheetSafe(anlassId),
      sheetSafe(name),
      sheetSafe(email),
      sheetSafe(telefon),
      sheetSafe(anlassName)
    ]);
    anlassSheet.getRange(anlassRow, 6).setValue(aktuelleHelfer + 1);
    
    lock.releaseLock();
    
    // Send email notification to admin (optional)
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
    
  } catch (e) {
    if (lock.hasLock()) {
      lock.releaseLock();
    }
    return { success: false, message: 'Ein Fehler ist aufgetreten. Bitte versuchen Sie es erneut.' };
  }
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

function exportDataAsCSV() {
  var result = exportData();
  if (result.success) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var exportSheet = ss.getSheetByName('Export') || ss.insertSheet('Export');
    exportSheet.clear();
    
    var lines = result.data.split('\n');
    var data = [];
    for (var i = 0; i < lines.length; i++) {
      if (lines[i]) {
        // sheetSafe() prefixes formula-trigger lines (e.g. the
        // "=== ANLÄSSE ===" separators) with ' so Sheets renders
        // them as text instead of trying to parse them as formulas.
        data.push([sheetSafe(lines[i])]);
      }
    }
    
    if (data.length > 0) {
      exportSheet.getRange(1, 1, data.length, 1).setValues(data);
      ss.setActiveSheet(exportSheet);
      SpreadsheetApp.getUi().alert('Export erfolgreich! Die Daten wurden im Tab "Export" gespeichert.');
    }
  } else {
    SpreadsheetApp.getUi().alert('Fehler beim Export: ' + result.error);
  }
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('🏰 Schulhelfer')
    .addItem('Erstes Setup', 'erstesSetup')
    .addItem('Neuer Anlass hinzufügen', 'neuerAnlassDialog')
    .addSeparator()
    .addItem('Alle Anmeldungen anzeigen', 'zeigeAnmeldungen')
    .addItem('Daten exportieren (CSV)', 'exportDataAsCSV')
    .addSeparator()
    .addItem('Audit-Log anzeigen', 'zeigeAuditLog')
    .addToUi();
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
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Anlässe');
  var values = sheet.getDataRange().getValues();
  
  // Sanitize input
  var name = sanitizeInput(data.name || '');
  var zeit = sanitizeInput(data.zeit || '');
  var beschreibung = sanitizeInput(data.beschreibung || '');
  var helfer = parseInt(data.helfer) || 0;
  
  if (!name || helfer < 1) {
    throw new Error('Bitte füllen Sie alle Pflichtfelder aus.');
  }
  
  // Validate date
  var datum = parseDate(data.datum);
  if (!datum) {
    throw new Error('Ungültiges Datum.');
  }
  
  // Find max ID
  var maxId = 0;
  for (var i = 1; i < values.length; i++) { 
    var id = parseInt(values[i][0]) || 0; 
    if (id > maxId) maxId = id; 
  }
  
  sheet.appendRow([
    maxId + 1,
    sheetSafe(name),
    datum,
    sheetSafe(zeit),
    helfer,
    0,
    sheetSafe(beschreibung)
  ]);
  
  logAudit('ANLASS_HINZUGEFUEGT', { name: name, datum: datum }, true, null);
}
