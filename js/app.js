/**
 * SCHULHELFER – Primarstufe Rittergasse Basel
 * Barrierefreie JavaScript-Interaktion
 */

(function() {
  'use strict';

  const $ = (sel) => document.querySelector(sel);
  const $$ = (sel) => document.querySelectorAll(sel);

  // === Centralized Configuration ===
  const AppConfig = {
    // Retry settings
    MAX_RETRIES: 3,
    RETRY_DELAY: 1000, // ms

    // Timeout settings (in ms)
    TIMEOUT_GET: 10000,   // 10 seconds for GET requests
    TIMEOUT_POST: 15000,  // 15 seconds for POST requests

    // Form persistence
    FORM_EXPIRY: 3600000, // 1 hour in ms

    // Refresh delay after successful registration
    REFRESH_DELAY: 2000,  // 2 seconds

    // Default event time if not specified
    DEFAULT_EVENT_HOUR: 10,
    DEFAULT_EVENT_DURATION: 2  // hours
  };

  const el = {
    loading: $('#loading'),
    errorMessage: $('#error-message'),
    errorText: $('#error-text'),
    successMessage: $('#success-message'),
    successText: $('#success-text'),
    statusMessage: $('#status-message'),
    eventsSection: $('#events-section'),
    eventsList: $('#events-list'),
    noEvents: $('#no-events'),
    registrationSection: $('#registration-section'),
    selectedEventName: $('#selected-event-name'),
    form: $('#registration-form'),
    eventIdInput: $('#event-id'),
    nameInput: $('#name'),
    emailInput: $('#email'),
    phoneInput: $('#phone'),
    submitBtn: $('#submit-btn')
  };

  // Centralized application state
  const AppState = {
    events: [],
    selectedEvent: null,
    lastRegistration: null,
    retryCount: 0,

    reset() {
      this.events = [];
      this.selectedEvent = null;
      this.lastRegistration = null;
      this.retryCount = 0;
    },

    setEvents(events) {
      this.events = (events || []).slice();
    },

    setSelectedEvent(event) {
      this.selectedEvent = event;
    },

    setLastRegistration(registration) {
      this.lastRegistration = registration;
    },

    findEventById(id) {
      // Use String comparison for consistent matching
      return this.events.find(e => String(e.id) === String(id));
    },

    clearSelectedEvent() {
      this.selectedEvent = null;
    }
  };

  // === Admin mode ===
  // Organisers become admin by visiting the site once with ?admin=KEY in
  // the URL. The key is persisted in localStorage and stripped from the
  // URL so it isn't bookmarked, shared, or visible in the address bar.
  // Regular visitors never see the "Helferliste" button.
  const ADMIN_KEY_STORAGE = 'schulhelfer_admin_key';

  function captureAdminKeyFromUrl() {
    try {
      const params = new URLSearchParams(window.location.search);
      const key = params.get('admin');
      if (key) {
        localStorage.setItem(ADMIN_KEY_STORAGE, key);
        params.delete('admin');
        const qs = params.toString();
        const clean = window.location.pathname +
          (qs ? '?' + qs : '') +
          window.location.hash;
        window.history.replaceState({}, document.title, clean);
      }
    } catch (e) {
      console.warn('Could not capture admin key:', e);
    }
  }

  function getAdminKey() {
    try {
      return localStorage.getItem(ADMIN_KEY_STORAGE) || '';
    } catch (e) {
      return '';
    }
  }

  function clearAdminKey() {
    try {
      localStorage.removeItem(ADMIN_KEY_STORAGE);
    } catch (e) { /* ignore */ }
  }

  function isAdmin() {
    return !!getAdminKey();
  }

  function getIdentifier() {
    let id = localStorage.getItem('userIdentifier');
    if (!id) {
      if (window.crypto && typeof crypto.randomUUID === 'function') {
        id = crypto.randomUUID();
      } else if (window.crypto && crypto.getRandomValues) {
        const a = new Uint32Array(4);
        crypto.getRandomValues(a);
        id = Array.from(a, (n) => n.toString(16).padStart(8, '0')).join('-');
      } else {
        id = 'u-' + Math.random().toString(36).substr(2) + Math.random().toString(36).substr(2);
      }
      localStorage.setItem('userIdentifier', id);
    }
    return id;
  }

  document.addEventListener('DOMContentLoaded', init);

  async function init() {
    if (!CONFIG.API_URL || CONFIG.API_URL === 'IHRE_GOOGLE_APPS_SCRIPT_URL_HIER') {
      showError('Bitte konfigurieren Sie die API_URL in index.html mit Ihrer Google Apps Script URL.');
      return;
    }
    captureAdminKeyFromUrl();
    setupAdminIndicator();
    setupFormValidation();
    setupKeyboardNavigation();
    setupFormPersistence();
    setupSpreadsheetLink();
    setupServiceWorkerUpdate();
    await loadEvents();
    restoreFormData();
  }

  function setupServiceWorkerUpdate() {
    if (!('serviceWorker' in navigator)) return;
    navigator.serviceWorker.addEventListener('controllerchange', () => {
      showUpdateToast();
    });
  }

  function showUpdateToast() {
    let toast = document.getElementById('sw-update-toast');
    if (toast) return;
    toast = document.createElement('div');
    toast.id = 'sw-update-toast';
    toast.className = 'sw-update-toast';
    toast.setAttribute('role', 'alert');
    toast.innerHTML =
      '<span>Neue Version verfügbar</span>' +
      '<button type="button" onclick="location.reload()">Aktualisieren</button>' +
      '<button type="button" class="sw-toast-close" aria-label="Schliessen" ' +
        'onclick="this.parentElement.remove()">✕</button>';
    document.body.appendChild(toast);
    requestAnimationFrame(() => toast.classList.add('is-visible'));
  }
  
  // Show a small badge + logout link when the device is in admin mode,
  // so organisers know the key is active and can turn it off.
  function setupAdminIndicator() {
    let badge = document.getElementById('admin-indicator');
    if (!isAdmin()) {
      if (badge) badge.remove();
      return;
    }
    if (!badge) {
      badge = document.createElement('div');
      badge.id = 'admin-indicator';
      badge.className = 'admin-indicator';
      badge.innerHTML =
        '<span class="admin-indicator-label">Admin-Modus</span>' +
        '<button type="button" class="admin-indicator-logout" aria-label="Admin-Modus verlassen">' +
          'Abmelden' +
        '</button>';
      badge.querySelector('.admin-indicator-logout').addEventListener('click', () => {
        clearAdminKey();
        setupAdminIndicator();
        setupSpreadsheetLink();
        renderEvents();
        announce('Admin-Modus beendet');
      });
      document.body.appendChild(badge);
    }
  }

  // Setup spreadsheet link. Admin-only: parents must never see the
  // Sheet URL — it contains every helper's name, e-mail and phone.
  // Re-call this after admin-key capture / logout so the link toggles
  // visibility in step with the admin indicator.
  function setupSpreadsheetLink() {
    const link = $('#spreadsheet-link');
    if (!link) return;
    if (isAdmin() && CONFIG.SPREADSHEET_URL && CONFIG.SPREADSHEET_URL.trim() !== '') {
      link.href = CONFIG.SPREADSHEET_URL;
      link.style.display = 'inline-flex';
    } else {
      link.removeAttribute('href');
      link.style.display = 'none';
    }
  }

  // === Load Events with Retry Logic ===
  async function loadEvents(retryAttempt = 0) {
    showLoading(true);
    hideError();

    try {
      // Build URL with cache-busting and identifier
      const identifier = getIdentifier();
      const url = `${CONFIG.API_URL}?action=getEvents&identifier=${encodeURIComponent(identifier)}&_=${Date.now()}`;
      
      // Create timeout controller for older browsers
      const controller = new AbortController();
      const timeoutId = setTimeout(() => controller.abort(), AppConfig.TIMEOUT_GET);
      
      const response = await fetch(url, {
        method: 'GET',
        redirect: 'follow',
        signal: controller.signal
      });
      
      clearTimeout(timeoutId);

      if (!response.ok) {
        throw new Error(`Server-Fehler: ${response.status}`);
      }

      const data = await response.json();
      
      if (data.error) {
        throw new Error(data.error);
      }

      AppState.setEvents(data.events);
      renderEvents();
      AppState.retryCount = 0; // Reset on success
      announce(`${AppState.events.length} ${AppState.events.length === 1 ? 'Anlass' : 'Anlässe'} gefunden`);

    } catch (error) {
      console.error('Fehler beim Laden:', error);
      
      // Retry logic for network failures
      if (retryAttempt < AppConfig.MAX_RETRIES && (
        error.message.includes('Failed to fetch') ||
        error.message.includes('network') ||
        error.name === 'TimeoutError' ||
        error.name === 'AbortError'
      )) {
        retryAttempt++;
        AppState.retryCount = retryAttempt;
        announce(`Verbindungsfehler. Versuch ${retryAttempt}/${AppConfig.MAX_RETRIES}...`);
        await new Promise(resolve => setTimeout(resolve, AppConfig.RETRY_DELAY * retryAttempt));
        return loadEvents(retryAttempt);
      }
      
      // More helpful error message
      let message = 'Die Anlässe konnten nicht geladen werden.';
      if (error.message.includes('Failed to fetch') || error.message.includes('CORS') || error.name === 'TimeoutError') {
        message = 'Verbindungsfehler. Bitte prüfen Sie:\n' +
                  '1. Ist die Google Apps Script URL korrekt?\n' +
                  '2. Wurde die Web-App als "Jeder" (nicht "Jeder mit Google-Konto") freigegeben?\n' +
                  '3. Wurde nach Code-Änderungen eine NEUE Bereitstellung erstellt?\n' +
                  '4. Ist Ihre Internetverbindung aktiv?';
      } else if (error.message.includes('Rate limit')) {
        message = 'Zu viele Anfragen. Bitte warten Sie einen Moment und versuchen Sie es erneut.';
      }
      showError(message);
    } finally {
      showLoading(false);
    }
  }

  // === Render Events ===
  function renderEvents() {
    if (AppState.events.length === 0) {
      el.eventsList.innerHTML = '';
      el.noEvents.hidden = false;
      return;
    }

    el.noEvents.hidden = true;
    el.eventsList.innerHTML = AppState.events.map((event, i) => createEventCard(event, i)).join('');

    $$('.event-card').forEach(card => {
      card.addEventListener('click', handleEventClick);
      card.addEventListener('keydown', handleEventKeydown);
    });
  }

  function createEventCard(event, index) {
    const voll = !!event.voll;
    const badgeClass = voll ? 'event-badge--full' :
                       event.freiePlaetze <= 1 ? 'event-badge--last' :
                       event.freiePlaetze <= 3 ? 'event-badge--limited' : 'event-badge--available';
    const badgeText = voll ? 'Ausgebucht' :
                      event.freiePlaetze <= 1 ? 'Noch 1 Helfer benötigt' :
                      `Noch ${event.freiePlaetze} Helfer benötigt`;
    
    // Parse date for calendar
    const eventDate = parseEventDate(event.datum);
    const eventTime = parseEventTime(event.zeit);
    
    // Calculate countdown
    const countdown = getCountdown(eventDate);
    
    return `
      <article class="event-card${voll ? ' event-card--full' : ''}" role="listitem" tabindex="0"
        data-id="${esc(event.id)}"
        data-name="${esc(event.name)}"
        data-date="${eventDate ? eventDate.toISOString() : ''}"
        data-time="${eventTime ? esc(JSON.stringify(eventTime)) : ''}"
        data-description="${esc(event.beschreibung || '')}"
        data-kontakt-name="${esc(event.kontaktName || '')}"
        data-kontakt-email="${esc(event.kontaktEmail || '')}"
        data-voll="${voll ? '1' : '0'}"
        aria-selected="false">
        ${countdown ? `<div class="event-countdown">
          <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5">
            <circle cx="12" cy="12" r="10"/><polyline points="12 6 12 12 16 14"/>
          </svg>
          <span>${countdown}</span>
        </div>` : ''}
        <div class="event-header">
          <h3 class="event-name">${esc(event.name)}</h3>
          <span class="event-badge ${badgeClass}">${badgeText}</span>
        </div>
        <div class="event-meta">
          <span class="event-meta-item">
            <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
              <rect x="3" y="4" width="18" height="18" rx="2"/><line x1="16" y1="2" x2="16" y2="6"/>
              <line x1="8" y1="2" x2="8" y2="6"/><line x1="3" y1="10" x2="21" y2="10"/>
            </svg>
            <span>${esc(event.datum)}</span>
          </span>
          ${event.zeit ? `<span class="event-meta-item">
            <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
              <circle cx="12" cy="12" r="10"/><polyline points="12 6 12 12 16 14"/>
            </svg>
            <span>${esc(event.zeit)}</span>
          </span>` : ''}
          <span class="event-meta-item">
            <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
              <path d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2"/><circle cx="9" cy="7" r="4"/>
              <path d="M23 21v-2a4 4 0 0 0-3-3.87"/><path d="M16 3.13a4 4 0 0 1 0 7.75"/>
            </svg>
            <span>${event.aktuelleHelfer}/${event.maxHelfer} Helfer</span>
          </span>
        </div>
        ${event.beschreibung ? `<p class="event-description">${esc(event.beschreibung)}</p>` : ''}
        ${(event.kontaktName || event.kontaktEmail) ? `
          <p class="event-contact" aria-label="Ansprechperson">
            <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" aria-hidden="true">
              <path d="M20 21v-2a4 4 0 0 0-4-4H8a4 4 0 0 0-4 4v2"/><circle cx="12" cy="7" r="4"/>
            </svg>
            <span class="event-contact-label">Kontakt:</span>
            ${event.kontaktName ? `<span>${esc(event.kontaktName)}</span>` : ''}
            ${event.kontaktName && event.kontaktEmail ? `<span aria-hidden="true">·</span>` : ''}
            ${event.kontaktEmail ? `<a href="mailto:${esc(event.kontaktEmail)}">${esc(event.kontaktEmail)}</a>` : ''}
          </p>
        ` : ''}
        <div class="event-actions">
          ${voll ? `
          <div class="event-cta event-cta--disabled" aria-hidden="true">
            <span>Ausgebucht – danke für Ihr Interesse</span>
          </div>` : `
          <div class="event-cta" aria-hidden="true">
            <span>Jetzt anmelden</span>
            <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
              <polyline points="9 18 15 12 9 6"/>
            </svg>
          </div>`}
          <div class="event-downloads">
            ${isAdmin() ? `
            <button type="button" class="download-btn helpers-download-btn"
              onclick="event.stopPropagation(); downloadHelpersList(this.closest('.event-card').dataset.id);"
              aria-label="Helferliste als Word-Datei herunterladen">
              <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
                <polyline points="14 2 14 8 20 8"/>
                <line x1="9" y1="13" x2="15" y2="13"/>
                <line x1="9" y1="17" x2="15" y2="17"/>
                <line x1="9" y1="9" x2="11" y2="9"/>
              </svg>
              <span class="download-btn-label">Helferliste (Word)</span>
            </button>` : ''}
            <button type="button" class="download-btn calendar-download-btn"
              onclick="event.stopPropagation(); downloadCalendarEvent(this.closest('.event-card'));"
              aria-label="Anlass zum Kalender hinzufügen">
              <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                <rect x="3" y="4" width="18" height="18" rx="2" ry="2"></rect>
                <line x1="16" y1="2" x2="16" y2="6"></line>
                <line x1="8" y1="2" x2="8" y2="6"></line>
                <line x1="3" y1="10" x2="21" y2="10"></line>
              </svg>
              <span class="download-btn-label calendar-btn-label">Kalendereintrag speichern</span>
            </button>
          </div>
        </div>
      </article>`;
  }
  
  // German month names mapping
  const GERMAN_MONTHS = {
    'januar': 0, 'februar': 1, 'märz': 2, 'april': 3, 'mai': 4, 'juni': 5,
    'juli': 6, 'august': 7, 'september': 8, 'oktober': 9, 'november': 10, 'dezember': 11
  };

  // Validate that a Date object represents the expected day/month/year
  function isValidDate(date, year, month, day) {
    return date instanceof Date &&
           !isNaN(date.getTime()) &&
           date.getFullYear() === year &&
           date.getMonth() === month &&
           date.getDate() === day;
  }

  // Parse German date string to Date object
  // Supports formats: "Montag, 15. Juli 2025", "15. Juli 2025", "15.7.2025"
  function parseEventDate(dateStr) {
    if (!dateStr || typeof dateStr !== 'string') return null;

    try {
      // Pattern 1: German format with month name "15. Juli 2025" or "Montag, 15. Juli 2025"
      const germanMatch = dateStr.match(/(\d{1,2})\.\s*(\w+)\s+(\d{4})/i);
      if (germanMatch) {
        const day = parseInt(germanMatch[1], 10);
        const monthName = germanMatch[2].toLowerCase().trim();
        const year = parseInt(germanMatch[3], 10);
        const month = GERMAN_MONTHS[monthName];

        if (month !== undefined && day >= 1 && day <= 31 && year >= 2000 && year <= 2100) {
          const date = new Date(year, month, day);
          if (isValidDate(date, year, month, day)) {
            return date;
          }
        }
      }

      // Pattern 2: Numeric format "15.7.2025" or "15.07.2025"
      const numericMatch = dateStr.match(/(\d{1,2})\.(\d{1,2})\.(\d{4})/);
      if (numericMatch) {
        const day = parseInt(numericMatch[1], 10);
        const month = parseInt(numericMatch[2], 10) - 1; // JS months are 0-indexed
        const year = parseInt(numericMatch[3], 10);

        if (month >= 0 && month <= 11 && day >= 1 && day <= 31 && year >= 2000 && year <= 2100) {
          const date = new Date(year, month, day);
          if (isValidDate(date, year, month, day)) {
            return date;
          }
        }
      }

      // Pattern 3: ISO format fallback (for data attributes)
      const isoDate = new Date(dateStr);
      if (!isNaN(isoDate.getTime())) {
        return isoDate;
      }

      console.warn('Could not parse date:', dateStr);
      return null;
    } catch (e) {
      console.error('Date parsing error:', e, dateStr);
      return null;
    }
  }
  
  // Calculate countdown text
  function getCountdown(eventDate) {
    if (!eventDate || isNaN(eventDate.getTime())) return null;
    
    const now = new Date();
    now.setHours(0, 0, 0, 0);
    
    const target = new Date(eventDate);
    target.setHours(0, 0, 0, 0);
    
    const diffTime = target - now;
    const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
    
    if (diffDays < 0) return null;
    if (diffDays === 0) return 'Heute!';
    if (diffDays === 1) return 'Morgen';
    if (diffDays <= 7) return `In ${diffDays} Tagen`;
    if (diffDays <= 14) return 'In 2 Wochen';
    if (diffDays <= 21) return 'In 3 Wochen';
    if (diffDays <= 30) return 'In ca. 1 Monat';
    if (diffDays <= 60) return 'In ca. 2 Monaten';
    return null;
  }

  // Parse time string (e.g., "14:00-18:00")
  function parseEventTime(timeStr) {
    if (!timeStr) return null;
    try {
      const match = timeStr.match(/(\d{1,2}):(\d{2})\s*-\s*(\d{1,2}):(\d{2})/);
      if (match) {
        return {
          start: { hour: parseInt(match[1]), minute: parseInt(match[2]) },
          end: { hour: parseInt(match[3]), minute: parseInt(match[4]) }
        };
      }
      // Try single time
      const singleMatch = timeStr.match(/(\d{1,2}):(\d{2})/);
      if (singleMatch) {
        const hour = parseInt(singleMatch[1]);
        const minute = parseInt(singleMatch[2]);
        return {
          start: { hour, minute },
          end: { hour: hour + AppConfig.DEFAULT_EVENT_DURATION, minute }
        };
      }
      return null;
    } catch (e) {
      return null;
    }
  }

  // === Shared Calendar Utility ===
  const ICAL_TZID = 'Europe/Zurich';

  // Format date for iCal as local time (YYYYMMDDTHHMMSS) – used with TZID
  function formatICalDate(date) {
    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const day = String(date.getDate()).padStart(2, '0');
    const hours = String(date.getHours()).padStart(2, '0');
    const minutes = String(date.getMinutes()).padStart(2, '0');
    const seconds = String(date.getSeconds()).padStart(2, '0');
    return `${year}${month}${day}T${hours}${minutes}${seconds}`;
  }

  // UTC timestamp for DTSTAMP (always UTC per RFC 5545)
  function formatICalDateUTC(date) {
    const year = date.getUTCFullYear();
    const month = String(date.getUTCMonth() + 1).padStart(2, '0');
    const day = String(date.getUTCDate()).padStart(2, '0');
    const hours = String(date.getUTCHours()).padStart(2, '0');
    const minutes = String(date.getUTCMinutes()).padStart(2, '0');
    const seconds = String(date.getUTCSeconds()).padStart(2, '0');
    return `${year}${month}${day}T${hours}${minutes}${seconds}Z`;
  }

  // Escape text for iCal format
  function escapeICal(text) {
    if (!text) return '';
    return String(text)
      .replace(/\\/g, '\\\\')
      .replace(/;/g, '\\;')
      .replace(/,/g, '\\,')
      .replace(/\n/g, '\\n')
      .replace(/\r/g, '');
  }

  // Create iCal content from event data
  function createICalContent(options) {
    const { eventName, eventDate, eventTime, description, attendeeName } = options;

    if (!eventDate || isNaN(eventDate.getTime())) {
      return null;
    }

    // Set start time
    const start = new Date(eventDate);
    if (eventTime && eventTime.start) {
      start.setHours(eventTime.start.hour, eventTime.start.minute, 0, 0);
    } else {
      start.setHours(AppConfig.DEFAULT_EVENT_HOUR, 0, 0, 0);
    }

    // Set end time
    const end = new Date(start);
    if (eventTime && eventTime.end) {
      end.setHours(eventTime.end.hour, eventTime.end.minute, 0, 0);
    } else {
      end.setHours(start.getHours() + AppConfig.DEFAULT_EVENT_DURATION, 0, 0, 0);
    }

    const startStr = formatICalDate(start);
    const endStr = formatICalDate(end);
    const nowStr = formatICalDateUTC(new Date());

    const summary = attendeeName
      ? `${eventName} - ${attendeeName}`
      : eventName;

    return [
      'BEGIN:VCALENDAR',
      'VERSION:2.0',
      'PRODID:-//Schulhelfer Rittergasse//DE',
      'CALSCALE:GREGORIAN',
      'METHOD:PUBLISH',
      'BEGIN:VTIMEZONE',
      `TZID:${ICAL_TZID}`,
      'BEGIN:STANDARD',
      'DTSTART:19701025T030000',
      'RRULE:FREQ=YEARLY;BYMONTH=10;BYDAY=-1SU',
      'TZOFFSETFROM:+0200',
      'TZOFFSETTO:+0100',
      'TZNAME:CET',
      'END:STANDARD',
      'BEGIN:DAYLIGHT',
      'DTSTART:19700329T020000',
      'RRULE:FREQ=YEARLY;BYMONTH=3;BYDAY=-1SU',
      'TZOFFSETFROM:+0100',
      'TZOFFSETTO:+0200',
      'TZNAME:CEST',
      'END:DAYLIGHT',
      'END:VTIMEZONE',
      'BEGIN:VEVENT',
      `UID:${Date.now()}-${Math.random().toString(36).substr(2, 9)}@schulhelfer-rittergasse`,
      `DTSTAMP:${nowStr}`,
      `DTSTART;TZID=${ICAL_TZID}:${startStr}`,
      `DTEND;TZID=${ICAL_TZID}:${endStr}`,
      `SUMMARY:${escapeICal(summary)}`,
      `DESCRIPTION:${escapeICal(description || 'Schulhelfer Anlass - Primarstufe Rittergasse Basel')}`,
      `LOCATION:Primarstufe Rittergasse Basel`,
      'STATUS:CONFIRMED',
      'SEQUENCE:0',
      'END:VEVENT',
      'END:VCALENDAR'
    ].join('\r\n');
  }

  // Download an iCal file
  function downloadICalFile(icalContent, filename) {
    const blob = new Blob([icalContent], { type: 'text/calendar;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = filename;
    link.style.display = 'none';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
  }

  // Generate iCal file from event card
  function generateICal(eventCard) {
    try {
      const name = eventCard.dataset.name;
      const dateStr = eventCard.dataset.date;
      const timeStr = eventCard.dataset.time;
      const baseDescription = eventCard.dataset.description || '';
      const kontaktName = eventCard.dataset.kontaktName || '';
      const kontaktEmail = eventCard.dataset.kontaktEmail || '';
      // If a public contact is configured for this event, append it to
      // the iCal DESCRIPTION so it lands in the user's calendar app
      // alongside the event details.
      const contactLine = (kontaktName || kontaktEmail)
        ? '\n\nKontakt: ' + [kontaktName, kontaktEmail].filter(Boolean).join(' – ')
        : '';
      const description = baseDescription + contactLine;

      if (!name) {
        console.error('No event name found');
        return null;
      }

      // Try to get date from data attribute first
      let eventDate = null;
      if (dateStr) {
        eventDate = new Date(dateStr);
        if (isNaN(eventDate.getTime())) {
          eventDate = null;
        }
      }

      // If no valid date from data attribute, try to parse from the displayed date
      if (!eventDate) {
        const dateElement = eventCard.querySelector('.event-meta-item span');
        if (dateElement) {
          const displayedDate = dateElement.textContent.trim();
          eventDate = parseEventDate(displayedDate);
        }
      }

      // Parse time
      let eventTime = null;
      if (timeStr && timeStr !== 'null' && timeStr !== '') {
        try {
          eventTime = JSON.parse(timeStr);
        } catch (e) {
          // Try to parse from displayed time
          const timeElements = eventCard.querySelectorAll('.event-meta-item');
          if (timeElements.length > 1) {
            const timeText = timeElements[1].textContent.trim();
            eventTime = parseEventTime(timeText);
          }
        }
      } else {
        // Try to parse from displayed time
        const timeElements = eventCard.querySelectorAll('.event-meta-item');
        if (timeElements.length > 1) {
          const timeText = timeElements[1].textContent.trim();
          eventTime = parseEventTime(timeText);
        }
      }

      return createICalContent({
        eventName: name,
        eventDate: eventDate,
        eventTime: eventTime,
        description: description
      });
    } catch (error) {
      console.error('Error generating iCal:', error);
      return null;
    }
  }
  
  // === Helper list (Word .docx) download ===
  // The .docx generator (CRC, ZIP, OOXML, layout) lives in the shared
  // js/helferliste-docx.js module and is exposed as
  // window.HelferListeDocx — both this page and /admin/ use it.
  // The block that used to be inline here was ~360 lines and has been
  // removed; downloadHelpersList below is a thin fetch+build wrapper.
  window.downloadHelpersList = async function(eventId) {
    const adminKey = getAdminKey();
    if (!adminKey) {
      showError('Keine Berechtigung. Öffnen Sie die Seite einmalig mit ?admin=IHR_SCHLÜSSEL in der URL.');
      return;
    }
    if (!window.HelferListeDocx) {
      showError('Word-Generator nicht geladen. Bitte Seite neu laden.');
      return;
    }
    try {
      announce('Helferliste wird geladen …');

      const url = CONFIG.API_URL +
        '?action=getHelferList' +
        '&eventId=' + encodeURIComponent(eventId) +
        '&adminKey=' + encodeURIComponent(adminKey) +
        '&identifier=' + encodeURIComponent(getIdentifier()) +
        '&_=' + Date.now();

      const controller = new AbortController();
      const timeoutId = setTimeout(() => controller.abort(), AppConfig.TIMEOUT_GET);

      const response = await fetch(url, {
        method: 'GET',
        redirect: 'follow',
        signal: controller.signal
      });
      clearTimeout(timeoutId);

      if (!response.ok) throw new Error('Server-Fehler: ' + response.status);
      const data = await response.json();

      if (!data.success) {
        // Only drop the stored key when the backend explicitly rejects it
        // ("Keine Berechtigung"). Other failures — missing ADMIN_KEY config,
        // unknown action, network issues — leave the key in place so the
        // organiser isn't silently logged out.
        if (data.error && /keine berechtigung/i.test(data.error)) {
          clearAdminKey();
          setupAdminIndicator();
          setupSpreadsheetLink();
          renderEvents();
        }
        showError(data.error || 'Helferliste konnte nicht geladen werden.');
        return;
      }

      const event = data.event || AppState.findEventById(eventId) || { name: 'Anlass' };
      const bytes = window.HelferListeDocx.build(event, data.helpers || []);
      window.HelferListeDocx.download(bytes, window.HelferListeDocx.filenameFor(event));
      announce('Helferliste für "' + (event.name || '') + '" wurde heruntergeladen');
    } catch (error) {
      console.error('Error downloading helpers list:', error);
      showError('Fehler beim Herunterladen der Helferliste. Bitte versuchen Sie es erneut.');
    }
  };

  // Download calendar event
  window.downloadCalendarEvent = function(eventCard) {
    try {
      const ical = generateICal(eventCard);
      if (!ical) {
        showError('Kalender-Download nicht möglich. Das Datum konnte nicht erkannt werden.');
        return;
      }

      const name = eventCard.dataset.name || 'Schulhelfer-Anlass';
      const filename = `${name.replace(/[^a-z0-9äöüÄÖÜß]/gi, '_')}.ics`;
      downloadICalFile(ical, filename);
      announce(`Kalender-Eintrag für "${name}" wurde heruntergeladen`);
    } catch (error) {
      console.error('Error downloading calendar:', error);
      showError('Fehler beim Herunterladen des Kalenders. Bitte versuchen Sie es erneut.');
    }
  };

  // === Event Selection ===
  function handleEventClick(e) {
    const card = e.currentTarget;
    if (card.dataset.voll === '1') return; // no registration for full events
    selectEvent(card.dataset.id, card.dataset.name);
  }

  function handleEventKeydown(e) {
    if (e.key === 'Enter' || e.key === ' ') {
      const card = e.currentTarget;
      if (card.dataset.voll === '1') return;
      e.preventDefault();
      selectEvent(card.dataset.id, card.dataset.name);
    }
  }

  function selectEvent(eventId, eventName) {
    AppState.setSelectedEvent(AppState.findEventById(eventId));
    if (!AppState.selectedEvent) return;

    $$('.event-card').forEach(card => {
      card.setAttribute('aria-selected', (card.dataset.id == eventId).toString());
    });

    el.eventIdInput.value = eventId;
    el.selectedEventName.textContent = eventName;
    el.registrationSection.hidden = false;

    setTimeout(() => {
      el.registrationSection.scrollIntoView({ behavior: 'smooth', block: 'start' });
      setTimeout(() => {
        el.nameInput.focus();
        announce(`Anmeldung für ${eventName} geöffnet`);
      }, 400);
    }, 100);
  }

  // === Keyboard Navigation ===
  function setupKeyboardNavigation() {
    el.eventsList.addEventListener('keydown', (e) => {
      const cards = Array.from($$('.event-card'));
      const idx = cards.indexOf(document.activeElement);
      if (idx === -1) return;

      let newIdx;
      switch (e.key) {
        case 'ArrowDown': case 'ArrowRight': e.preventDefault(); newIdx = (idx + 1) % cards.length; break;
        case 'ArrowUp': case 'ArrowLeft': e.preventDefault(); newIdx = (idx - 1 + cards.length) % cards.length; break;
        case 'Home': e.preventDefault(); newIdx = 0; break;
        case 'End': e.preventDefault(); newIdx = cards.length - 1; break;
        default: return;
      }
      cards[newIdx].focus();
    });
  }

  // === Form Persistence ===
  function setupFormPersistence() {
    // Save form data on input
    [el.nameInput, el.emailInput, el.phoneInput].forEach(input => {
      input.addEventListener('input', () => {
        saveFormData();
      });
    });
    
    // Clear saved data on successful submission
    el.form.addEventListener('submit', () => {
      // Will be cleared after successful submission
    });
  }

  function saveFormData() {
    const formData = {
      name: el.nameInput.value,
      email: el.emailInput.value,
      phone: el.phoneInput.value,
      eventId: el.eventIdInput.value,
      timestamp: Date.now()
    };
    localStorage.setItem('schulhelfer_form', JSON.stringify(formData));
  }

  function restoreFormData() {
    try {
      const saved = localStorage.getItem('schulhelfer_form');
      if (saved) {
        const formData = JSON.parse(saved);
        if (Date.now() - formData.timestamp < AppConfig.FORM_EXPIRY) {
          if (formData.name) el.nameInput.value = formData.name;
          if (formData.email) el.emailInput.value = formData.email;
          if (formData.phone) el.phoneInput.value = formData.phone;
        } else {
          localStorage.removeItem('schulhelfer_form');
        }
      }
    } catch (e) {
      console.warn('Could not restore form data:', e);
    }
  }

  function clearFormData() {
    localStorage.removeItem('schulhelfer_form');
  }

  // === Form ===
  function setupFormValidation() {
    el.form.addEventListener('submit', handleSubmit);
    el.nameInput.addEventListener('blur', () => validateField(el.nameInput, validateName));
    el.emailInput.addEventListener('blur', () => validateField(el.emailInput, validateEmail));
    el.phoneInput.addEventListener('blur', () => validateField(el.phoneInput, validatePhone));
    el.nameInput.addEventListener('input', () => {
      clearFieldError(el.nameInput);
      saveFormData();
    });
    el.emailInput.addEventListener('input', () => {
      clearFieldError(el.emailInput);
      saveFormData();
    });
    el.phoneInput.addEventListener('input', () => {
      clearFieldError(el.phoneInput);
      saveFormData();
    });
  }

  function validateField(input, validator) {
    const error = validator(input.value);
    const errorEl = $(`#${input.id}-error`);
    if (error) {
      input.setAttribute('aria-invalid', 'true');
      if (errorEl) { errorEl.textContent = error; errorEl.hidden = false; }
      return false;
    }
    input.removeAttribute('aria-invalid');
    if (errorEl) errorEl.hidden = true;
    return true;
  }

  function clearFieldError(input) {
    const errorEl = $(`#${input.id}-error`);
    if (errorEl) { input.removeAttribute('aria-invalid'); errorEl.hidden = true; }
  }

  function validateName(v) {
    v = v.trim();
    if (!v) return 'Bitte geben Sie Ihren Namen ein.';
    if (v.length < 2) return 'Der Name muss mindestens 2 Zeichen haben.';
    return null;
  }

  function validateEmail(v) {
    v = v.trim();
    if (!v) return 'Bitte geben Sie Ihre E-Mail-Adresse ein.';
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(v)) return 'Bitte geben Sie eine gültige E-Mail-Adresse ein.';
    return null;
  }

  function validatePhone(v) {
    v = v.trim();
    if (!v) return null; // Optional field - empty is valid

    // Remove spaces and dashes for validation
    const cleaned = v.replace(/[\s\-]/g, '');

    // Swiss phone formats: 079XXXXXXX, +41XXXXXXXXX, 0041XXXXXXXXX
    // Also allow general international format
    if (!/^(\+?41|0041|0)\d{8,10}$/.test(cleaned) && !/^\+?\d{10,15}$/.test(cleaned)) {
      return 'Bitte geben Sie eine gültige Telefonnummer ein (z.B. 079 123 45 67).';
    }

    return null;
  }

  // === Submit with Retry Logic ===
  async function handleSubmit(e) {
    e.preventDefault();
    hideError();
    hideSuccess();

    const nameValid = validateField(el.nameInput, validateName);
    const emailValid = validateField(el.emailInput, validateEmail);
    const phoneValid = validateField(el.phoneInput, validatePhone);

    if (!nameValid || !emailValid || !phoneValid) {
      // Focus the first invalid field
      if (!nameValid) el.nameInput.focus();
      else if (!emailValid) el.emailInput.focus();
      else el.phoneInput.focus();
      announce('Bitte korrigieren Sie die markierten Felder.');
      return;
    }

    const honeypot = document.getElementById('website');
    const data = {
      anlassId: el.eventIdInput.value,
      name: el.nameInput.value.trim(),
      email: el.emailInput.value.trim(),
      telefon: el.phoneInput.value.trim(),
      website: honeypot ? honeypot.value : '',
      identifier: getIdentifier()
    };

    setSubmitLoading(true);

    try {
      const result = await submitWithRetry(data, 0);

      if (result.success) {
        // Store registration data for calendar download
        AppState.setLastRegistration({
          name: data.name,
          eventId: data.anlassId,
          eventName: AppState.selectedEvent ? AppState.selectedEvent.name : '',
          eventDate: AppState.selectedEvent ? AppState.selectedEvent.datum : '',
          eventTime: AppState.selectedEvent ? AppState.selectedEvent.zeit : '',
          eventDescription: AppState.selectedEvent ? AppState.selectedEvent.beschreibung : ''
        });

        showSuccess(result.message, AppState.lastRegistration);
        clearFormData(); // Clear saved form data on success
        resetForm();
        announce('Anmeldung erfolgreich!');
        setTimeout(loadEvents, AppConfig.REFRESH_DELAY);
      } else {
        showError(result.message || 'Ein Fehler ist aufgetreten.');
        announce(`Fehler: ${result.message}`);
      }
    } catch (error) {
      console.error('Fehler beim Senden:', error);
      showError('Die Anmeldung konnte nicht gesendet werden. Bitte versuchen Sie es später erneut.');
      announce('Fehler beim Senden.');
    } finally {
      setSubmitLoading(false);
    }
  }

  async function submitWithRetry(data, retryAttempt = 0) {
    try {
      const controller = new AbortController();
      const timeoutId = setTimeout(() => controller.abort(), AppConfig.TIMEOUT_POST);

      const response = await fetch(CONFIG.API_URL, {
        method: 'POST',
        redirect: 'follow',
        headers: { 'Content-Type': 'text/plain' },
        body: JSON.stringify(data),
        signal: controller.signal
      });

      clearTimeout(timeoutId);

      if (!response.ok) {
        throw new Error(`Server-Fehler: ${response.status}`);
      }

      const result = await response.json();

      if (!result.success && retryAttempt < AppConfig.MAX_RETRIES &&
        result.errorCode === 'RATE_LIMIT'
      ) {
        // Retry on rate limit after delay
        retryAttempt++;
        await new Promise(resolve => setTimeout(resolve, AppConfig.RETRY_DELAY * retryAttempt * 2));
        return submitWithRetry(data, retryAttempt);
      }

      return result;
    } catch (error) {
      // Retry on network errors
      if (retryAttempt < AppConfig.MAX_RETRIES && (
        error.name === 'AbortError' ||
        error.message.includes('Failed to fetch') ||
        error.message.includes('network')
      )) {
        retryAttempt++;
        await new Promise(resolve => setTimeout(resolve, AppConfig.RETRY_DELAY * retryAttempt));
        return submitWithRetry(data, retryAttempt);
      }
      throw error;
    }
  }

  function setSubmitLoading(loading) {
    el.submitBtn.disabled = loading;
    const btnContent = el.submitBtn.querySelector('.btn-content');
    const btnLoading = el.submitBtn.querySelector('.btn-loading');
    if (btnContent) btnContent.hidden = loading;
    if (btnLoading) btnLoading.hidden = !loading;
  }

  function resetForm() {
    el.form.reset();
    el.registrationSection.hidden = true;
    AppState.clearSelectedEvent();
    $$('.event-card').forEach(card => card.setAttribute('aria-selected', 'false'));
    clearFormData();
  }

  window.cancelRegistration = function() {
    resetForm();
    announce('Anmeldung abgebrochen');
    el.eventsSection.scrollIntoView({ behavior: 'smooth', block: 'start' });
    setTimeout(() => { const c = $('.event-card'); if (c) c.focus(); }, 400);
  };

  // === Helpers ===
  function showLoading(show) {
    el.loading.hidden = !show;
    el.eventsSection.setAttribute('aria-busy', show ? 'true' : 'false');
  }

  function showError(msg) {
    el.errorText.textContent = msg;
    el.errorMessage.hidden = false;
    el.errorMessage.scrollIntoView({ behavior: 'smooth', block: 'center' });
    // Move focus so keyboard users land on the message after submit.
    // tabindex=-1 + outline:none keeps the visual unchanged.
    if (el.errorMessage.getAttribute('tabindex') == null) {
      el.errorMessage.setAttribute('tabindex', '-1');
      el.errorMessage.style.outline = 'none';
    }
    try { el.errorMessage.focus({ preventScroll: true }); } catch (_) {}
  }

  function hideError() { el.errorMessage.hidden = true; }
  window.closeError = hideError;

  function showSuccess(msg, registrationData = null) {
    el.successText.textContent = msg;
    
    // Add calendar download button if registration data is available
    let calendarBtn = el.successMessage.querySelector('.calendar-download-success');
    if (registrationData) {
      if (!calendarBtn) {
        calendarBtn = document.createElement('button');
        calendarBtn.className = 'calendar-download-success';
        calendarBtn.type = 'button';
        calendarBtn.innerHTML = `
          <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
            <rect x="3" y="4" width="18" height="18" rx="2" ry="2"></rect>
            <line x1="16" y1="2" x2="16" y2="6"></line>
            <line x1="8" y1="2" x2="8" y2="6"></line>
            <line x1="3" y1="10" x2="21" y2="10"></line>
          </svg>
          <span>Kalender-Eintrag herunterladen</span>
        `;
        calendarBtn.addEventListener('click', () => {
          downloadRegistrationCalendar(registrationData);
        });
        el.successMessage.querySelector('.message-content').appendChild(calendarBtn);
      }
      calendarBtn.hidden = false;
    } else if (calendarBtn) {
      calendarBtn.hidden = true;
    }
    
    el.successMessage.hidden = false;
    el.successMessage.scrollIntoView({ behavior: 'smooth', block: 'center' });
    clearTimeout(showSuccess._timer);
    showSuccess._timer = setTimeout(hideSuccess, 10000);
  }
  
  // Download calendar for registered event
  function downloadRegistrationCalendar(registration) {
    try {
      const eventDate = parseEventDate(registration.eventDate);
      if (!eventDate) {
        showError('Kalender-Download nicht möglich. Das Datum konnte nicht erkannt werden.');
        return;
      }

      const eventTime = parseEventTime(registration.eventTime);
      const eventName = registration.eventName || 'Schulhelfer Anlass';
      const description = `Schulhelfer Anlass: ${eventName}\n\nAngemeldet als: ${registration.name}\n\n${registration.eventDescription || 'Primarstufe Rittergasse Basel'}`;

      const ical = createICalContent({
        eventName: eventName,
        eventDate: eventDate,
        eventTime: eventTime,
        description: description,
        attendeeName: registration.name
      });

      if (!ical) {
        showError('Kalender-Download nicht möglich. Das Datum konnte nicht erkannt werden.');
        return;
      }

      const filename = `${eventName.replace(/[^a-z0-9äöüÄÖÜß]/gi, '_')}_${registration.name.replace(/[^a-z0-9äöüÄÖÜß]/gi, '_')}.ics`;
      downloadICalFile(ical, filename);
      announce(`Kalender-Eintrag für "${eventName}" wurde heruntergeladen`);
    } catch (error) {
      console.error('Error downloading calendar:', error);
      showError('Fehler beim Herunterladen des Kalenders. Bitte versuchen Sie es erneut.');
    }
  }

  function hideSuccess() { el.successMessage.hidden = true; }

  function announce(msg) {
    if (el.statusMessage) {
      el.statusMessage.textContent = msg;
      setTimeout(() => { el.statusMessage.textContent = ''; }, 1000);
    }
  }

  // HTML-escape text for safe interpolation in both element content and
  // double/single-quoted attributes. Unlike the old textContent→innerHTML
  // trick, this also escapes quotes, which is required because we inject
  // values like `data-name="${esc(...)}"` in event-card templates.
  function esc(text) {
    if (text == null) return '';
    return String(text)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;');
  }

})();
