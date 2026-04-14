/**
 * SCHULHELFER – Primarstufe Rittergasse Basel
 * Barrierefreie JavaScript-Interaktion
 */

(function() {
  'use strict';

  const $ = (sel) => document.querySelector(sel);
  const $$ = (sel) => document.querySelectorAll(sel);

  // === Performance Metrics ===
  const Metrics = {
    loadEventsTiming: [],
    registrationTiming: [],
    maxSamples: 10,

    recordLoadEvents(durationMs) {
      this.loadEventsTiming.push(durationMs);
      if (this.loadEventsTiming.length > this.maxSamples) {
        this.loadEventsTiming.shift();
      }
      this.log('loadEvents', durationMs);
    },

    recordRegistration(durationMs) {
      this.registrationTiming.push(durationMs);
      if (this.registrationTiming.length > this.maxSamples) {
        this.registrationTiming.shift();
      }
      this.log('registration', durationMs);
    },

    getAverageLoadTime() {
      if (this.loadEventsTiming.length === 0) return 0;
      const sum = this.loadEventsTiming.reduce((a, b) => a + b, 0);
      return Math.round(sum / this.loadEventsTiming.length);
    },

    getAverageRegistrationTime() {
      if (this.registrationTiming.length === 0) return 0;
      const sum = this.registrationTiming.reduce((a, b) => a + b, 0);
      return Math.round(sum / this.registrationTiming.length);
    },

    log(operation, durationMs) {
      if (typeof console !== 'undefined' && console.debug) {
        console.debug(`[Metrics] ${operation}: ${durationMs}ms`);
      }
    },

    getSummary() {
      return {
        loadEvents: {
          average: this.getAverageLoadTime(),
          samples: this.loadEventsTiming.length,
          last: this.loadEventsTiming[this.loadEventsTiming.length - 1] || 0
        },
        registration: {
          average: this.getAverageRegistrationTime(),
          samples: this.registrationTiming.length,
          last: this.registrationTiming[this.registrationTiming.length - 1] || 0
        }
      };
    }
  };

  // Expose metrics for debugging (accessible via window.SchulhelferMetrics)
  window.SchulhelferMetrics = Metrics;

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
      this.events = (events || []).slice().sort((a, b) => {
        // Sort by datumSort if available (from backend), otherwise parse datum string
        const dateA = a.datumSort || parseEventDateForSort(a.datum);
        const dateB = b.datumSort || parseEventDateForSort(b.datum);
        return dateA - dateB;
      });
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

  // Generate unique identifier for rate limiting (more robust than simple localStorage)
  function getIdentifier() {
    let id = localStorage.getItem('userIdentifier');
    if (!id) {
      // Use crypto API for secure random if available, fallback to Math.random
      let randomPart;
      if (window.crypto && window.crypto.getRandomValues) {
        const array = new Uint32Array(2);
        window.crypto.getRandomValues(array);
        randomPart = array[0].toString(36) + array[1].toString(36);
      } else {
        randomPart = Math.random().toString(36).substr(2, 9) + Math.random().toString(36).substr(2, 9);
      }

      // Add browser fingerprint components for additional uniqueness
      const fingerprint = [
        navigator.language || '',
        screen.width + 'x' + screen.height,
        new Date().getTimezoneOffset()
      ].join('|');

      id = 'user_' + Date.now() + '_' + randomPart + '_' + btoa(fingerprint).substr(0, 8);
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
    // loadEvents must finish before restoreFormData so that the saved
    // event selection can be matched against the loaded event list.
    await loadEvents();
    restoreFormData();
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
        renderEvents();
        announce('Admin-Modus beendet');
      });
      document.body.appendChild(badge);
    }
  }

  // Setup spreadsheet link
  function setupSpreadsheetLink() {
    const link = $('#spreadsheet-link');
    if (link && CONFIG.SPREADSHEET_URL && CONFIG.SPREADSHEET_URL.trim() !== '') {
      link.href = CONFIG.SPREADSHEET_URL;
      link.style.display = 'inline-flex';
    }
  }

  // === Load Events with Retry Logic ===
  async function loadEvents(retryAttempt = 0) {
    showLoading(true);
    hideError();
    const startTime = performance.now();

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

      // Record successful load timing
      Metrics.recordLoadEvents(Math.round(performance.now() - startTime));

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
    const badgeClass = event.freiePlaetze <= 1 ? 'event-badge--last' : 
                       event.freiePlaetze <= 3 ? 'event-badge--limited' : 'event-badge--available';
    const badgeText = event.freiePlaetze <= 1 ? 'Noch 1 Helfer benötigt' :
                      `Noch ${event.freiePlaetze} Helfer benötigt`;
    
    // Parse date for calendar
    const eventDate = parseEventDate(event.datum);
    const eventTime = parseEventTime(event.zeit);
    
    // Calculate countdown
    const countdown = getCountdown(eventDate);
    
    return `
      <article class="event-card" role="listitem" tabindex="0"
        data-id="${esc(event.id)}"
        data-name="${esc(event.name)}"
        data-date="${eventDate ? eventDate.toISOString() : ''}"
        data-time="${eventTime ? esc(JSON.stringify(eventTime)) : ''}"
        data-description="${esc(event.beschreibung || '')}"
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
        <div class="event-actions">
          <div class="event-cta" aria-hidden="true">
            <span>Jetzt anmelden</span>
            <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
              <polyline points="9 18 15 12 9 6"/>
            </svg>
          </div>
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

  // Helper for sorting: returns timestamp or Infinity for invalid dates
  function parseEventDateForSort(dateStr) {
    const date = parseEventDate(dateStr);
    return date ? date.getTime() : Infinity;
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
  // Format date for iCal (YYYYMMDDTHHMMSSZ) - UTC format
  function formatICalDate(date) {
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
    const nowStr = formatICalDate(new Date());

    // Build summary (include attendee name if provided)
    const summary = attendeeName
      ? `${eventName} - ${attendeeName}`
      : eventName;

    return [
      'BEGIN:VCALENDAR',
      'VERSION:2.0',
      'PRODID:-//Schulhelfer Rittergasse//DE',
      'CALSCALE:GREGORIAN',
      'METHOD:PUBLISH',
      'BEGIN:VEVENT',
      `UID:${Date.now()}-${Math.random().toString(36).substr(2, 9)}@schulhelfer-rittergasse`,
      `DTSTAMP:${nowStr}`,
      `DTSTART:${startStr}`,
      `DTEND:${endStr}`,
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
      const description = eventCard.dataset.description || '';

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
  // A .docx file is a ZIP of OOXML parts. We package it client-side with a
  // minimal STORE-only (no compression) ZIP writer — no external library.
  // Word, LibreOffice and Google Docs all open the result natively.

  // Precomputed CRC-32 table (IEEE 802.3 polynomial) used by the ZIP writer.
  const CRC32_TABLE = (function() {
    const t = new Uint32Array(256);
    for (let i = 0; i < 256; i++) {
      let c = i;
      for (let k = 0; k < 8; k++) {
        c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1);
      }
      t[i] = c >>> 0;
    }
    return t;
  })();

  function crc32(bytes) {
    let c = 0xFFFFFFFF;
    for (let i = 0; i < bytes.length; i++) {
      c = CRC32_TABLE[(c ^ bytes[i]) & 0xFF] ^ (c >>> 8);
    }
    return (c ^ 0xFFFFFFFF) >>> 0;
  }

  // Build a ZIP archive (STORE / method 0) containing the given files.
  // files: [{ name: 'path/in/zip', data: string|Uint8Array }]
  function zipStore(files) {
    const encoder = new TextEncoder();
    const localChunks = [];
    const centralChunks = [];
    let offset = 0;

    for (const f of files) {
      const nameBytes = encoder.encode(f.name);
      const data = typeof f.data === 'string' ? encoder.encode(f.data) : f.data;
      const crc = crc32(data);
      const size = data.length;

      // Local file header (30 bytes + name)
      const lfh = new Uint8Array(30);
      const lv = new DataView(lfh.buffer);
      lv.setUint32(0, 0x04034b50, true);  // signature
      lv.setUint16(4, 20, true);           // version needed
      lv.setUint16(6, 0, true);            // flags
      lv.setUint16(8, 0, true);            // method (STORE)
      lv.setUint16(10, 0, true);           // mod time
      lv.setUint16(12, 0x21, true);        // mod date (1980-01-01)
      lv.setUint32(14, crc, true);
      lv.setUint32(18, size, true);        // compressed size
      lv.setUint32(22, size, true);        // uncompressed size
      lv.setUint16(26, nameBytes.length, true);
      lv.setUint16(28, 0, true);           // extra length

      localChunks.push(lfh, nameBytes, data);

      // Central directory entry (46 bytes + name)
      const cdh = new Uint8Array(46);
      const cv = new DataView(cdh.buffer);
      cv.setUint32(0, 0x02014b50, true);
      cv.setUint16(4, 20, true);           // version made by
      cv.setUint16(6, 20, true);           // version needed
      cv.setUint16(8, 0, true);
      cv.setUint16(10, 0, true);
      cv.setUint16(12, 0, true);
      cv.setUint16(14, 0x21, true);
      cv.setUint32(16, crc, true);
      cv.setUint32(20, size, true);
      cv.setUint32(24, size, true);
      cv.setUint16(28, nameBytes.length, true);
      cv.setUint16(30, 0, true);           // extra length
      cv.setUint16(32, 0, true);           // comment length
      cv.setUint16(34, 0, true);           // disk start
      cv.setUint16(36, 0, true);           // internal attrs
      cv.setUint32(38, 0, true);           // external attrs
      cv.setUint32(42, offset, true);      // local header offset

      centralChunks.push(cdh, nameBytes);
      offset += 30 + nameBytes.length + size;
    }

    const localTotal = offset;
    let centralTotal = 0;
    for (const c of centralChunks) centralTotal += c.length;

    // End of central directory
    const eocd = new Uint8Array(22);
    const ev = new DataView(eocd.buffer);
    ev.setUint32(0, 0x06054b50, true);
    ev.setUint16(4, 0, true);
    ev.setUint16(6, 0, true);
    ev.setUint16(8, files.length, true);
    ev.setUint16(10, files.length, true);
    ev.setUint32(12, centralTotal, true);
    ev.setUint32(16, localTotal, true);
    ev.setUint16(20, 0, true);

    const out = new Uint8Array(localTotal + centralTotal + 22);
    let pos = 0;
    for (const c of localChunks) { out.set(c, pos); pos += c.length; }
    for (const c of centralChunks) { out.set(c, pos); pos += c.length; }
    out.set(eocd, pos);
    return out;
  }

  // XML-escape text for placement inside <w:t>…</w:t>.
  function xmlEscape(s) {
    return String(s == null ? '' : s)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&apos;');
  }

  // Build a WordprocessingML paragraph. Sizes are in half-points (22 = 11pt).
  function wPara(text, opts) {
    opts = opts || {};
    const rProps = [];
    if (opts.bold) rProps.push('<w:b/><w:bCs/>');
    if (opts.size) rProps.push('<w:sz w:val="' + opts.size + '"/><w:szCs w:val="' + opts.size + '"/>');
    if (opts.color) rProps.push('<w:color w:val="' + opts.color + '"/>');
    const rPr = rProps.length ? '<w:rPr>' + rProps.join('') + '</w:rPr>' : '';

    const pProps = [];
    if (opts.spacingAfter != null) {
      pProps.push('<w:spacing w:after="' + opts.spacingAfter + '"/>');
    }
    const pPr = pProps.length ? '<w:pPr>' + pProps.join('') + '</w:pPr>' : '';

    return '<w:p>' + pPr + '<w:r>' + rPr +
      '<w:t xml:space="preserve">' + xmlEscape(text) + '</w:t>' +
      '</w:r></w:p>';
  }

  // Build a WordprocessingML table cell.
  function wCell(text, opts) {
    opts = opts || {};
    const width = opts.width || 2000;
    const shading = opts.shading
      ? '<w:shd w:val="clear" w:color="auto" w:fill="' + opts.shading + '"/>'
      : '';
    const tcPr = '<w:tcPr><w:tcW w:w="' + width + '" w:type="dxa"/>' + shading + '</w:tcPr>';

    const rProps = [];
    if (opts.bold) rProps.push('<w:b/><w:bCs/>');
    if (opts.textColor) rProps.push('<w:color w:val="' + opts.textColor + '"/>');
    const rPr = rProps.length ? '<w:rPr>' + rProps.join('') + '</w:rPr>' : '';

    const pProps = [];
    if (opts.align) pProps.push('<w:jc w:val="' + opts.align + '"/>');
    const pPr = pProps.length ? '<w:pPr>' + pProps.join('') + '</w:pPr>' : '';

    return '<w:tc>' + tcPr + '<w:p>' + pPr + '<w:r>' + rPr +
      '<w:t xml:space="preserve">' + xmlEscape(text || '') + '</w:t>' +
      '</w:r></w:p></w:tc>';
  }

  // Build the word/document.xml content for the Helferliste.
  // `helpers` is an array of { name, telefon, email, zeitstempel } objects
  // from the backend getHelferList endpoint.
  function buildHelpersDocumentXml(event, helpers) {
    helpers = Array.isArray(helpers) ? helpers : [];
    const maxHelfer = parseInt(event.maxHelfer, 10) || helpers.length;
    const rowCount = Math.max(maxHelfer, helpers.length);

    const COL_NR = 600;      // Nr. column (dxa, 1 pt = 20 dxa)
    const COL_NAME = 3800;   // Name column
    const COL_TEL = 2400;    // Telefon column
    const COL_SIGN = 2200;   // Unterschrift column
    const HEADER_COLOR = '1E3A5F';

    // Header row
    const headerRow = '<w:tr>' +
      '<w:trPr><w:tblHeader/></w:trPr>' +
      wCell('Nr.', { width: COL_NR, bold: true, shading: HEADER_COLOR, textColor: 'FFFFFF', align: 'center' }) +
      wCell('Name', { width: COL_NAME, bold: true, shading: HEADER_COLOR, textColor: 'FFFFFF' }) +
      wCell('Telefon', { width: COL_TEL, bold: true, shading: HEADER_COLOR, textColor: 'FFFFFF' }) +
      wCell('Unterschrift', { width: COL_SIGN, bold: true, shading: HEADER_COLOR, textColor: 'FFFFFF' }) +
      '</w:tr>';

    // Data rows
    const dataRows = [];
    for (let i = 0; i < rowCount; i++) {
      const h = helpers[i] || {};
      dataRows.push('<w:tr>' +
        wCell(String(i + 1), { width: COL_NR, align: 'center' }) +
        wCell(h.name || '', { width: COL_NAME }) +
        wCell(h.telefon || '', { width: COL_TEL }) +
        wCell('', { width: COL_SIGN }) +
        '</w:tr>');
    }

    const tableBorders = '<w:tblPr>' +
      '<w:tblW w:w="5000" w:type="pct"/>' +
      '<w:tblBorders>' +
        '<w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>' +
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>' +
        '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>' +
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>' +
        '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/>' +
        '<w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/>' +
      '</w:tblBorders>' +
      '</w:tblPr>';

    const metaParts = [];
    if (event.datum) metaParts.push(event.datum);
    if (event.zeit) metaParts.push(event.zeit);
    const metaLine = metaParts.join(' · ');

    const bodyParts = [];
    // Title (Heading-like, 18pt bold, brand color)
    bodyParts.push(wPara('Helferliste – ' + (event.name || ''), {
      bold: true, size: 36, color: HEADER_COLOR, spacingAfter: 60
    }));
    if (metaLine) {
      bodyParts.push(wPara(metaLine, { size: 22, spacingAfter: 40 }));
    }
    if (event.beschreibung) {
      bodyParts.push(wPara(String(event.beschreibung), { size: 22, spacingAfter: 40 }));
    }
    bodyParts.push(wPara(
      'Primarstufe Rittergasse Basel · Angemeldet: ' + helpers.length + '/' + maxHelfer,
      { size: 20, color: '666666', spacingAfter: 200 }
    ));

    // Table
    bodyParts.push('<w:tbl>' + tableBorders + headerRow + dataRows.join('') + '</w:tbl>');

    // Final empty paragraph + section properties (A4 portrait, 2cm margins)
    bodyParts.push('<w:p/>');
    bodyParts.push(
      '<w:sectPr>' +
        '<w:pgSz w:w="11906" w:h="16838"/>' +
        '<w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134" w:header="708" w:footer="708" w:gutter="0"/>' +
      '</w:sectPr>'
    );

    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
      '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
        '<w:body>' + bodyParts.join('') + '</w:body>' +
      '</w:document>';
  }

  // Assemble the full .docx package (ZIP of OOXML parts) as a Uint8Array.
  function buildHelpersDocx(event, helpers) {
    const contentTypes =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
        '<Default Extension="xml" ContentType="application/xml"/>' +
        '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
      '</Types>';

    const rootRels =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
      '</Relationships>';

    const documentXml = buildHelpersDocumentXml(event, helpers);

    return zipStore([
      { name: '[Content_Types].xml', data: contentTypes },
      { name: '_rels/.rels', data: rootRels },
      { name: 'word/document.xml', data: documentXml }
    ]);
  }

  function downloadBinaryFile(bytes, filename, mimeType) {
    const blob = new Blob([bytes], { type: mimeType });
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

  window.downloadHelpersList = async function(eventId) {
    const adminKey = getAdminKey();
    if (!adminKey) {
      showError('Keine Berechtigung. Öffnen Sie die Seite einmalig mit ?admin=IHR_SCHLÜSSEL in der URL.');
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

      if (!response.ok) {
        throw new Error('Server-Fehler: ' + response.status);
      }

      const data = await response.json();

      if (!data.success) {
        // Only drop the stored key when the backend explicitly rejects it
        // ("Keine Berechtigung"). Other failures — missing ADMIN_KEY config
        // on the server, unknown action, network issues — leave the key in
        // place so the organiser isn't silently logged out.
        if (data.error && /keine berechtigung/i.test(data.error)) {
          clearAdminKey();
          setupAdminIndicator();
          renderEvents();
        }
        showError(data.error || 'Helferliste konnte nicht geladen werden.');
        return;
      }

      const event = data.event || AppState.findEventById(eventId) || { name: 'Anlass' };
      const bytes = buildHelpersDocx(event, data.helpers || []);
      const safeName = String(event.name || 'Anlass').replace(/[^a-z0-9äöüÄÖÜß]/gi, '_');
      const filename = 'Helferliste_' + safeName + '.docx';
      downloadBinaryFile(
        bytes,
        filename,
        'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
      );
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
    selectEvent(e.currentTarget.dataset.id, e.currentTarget.dataset.name);
  }

  function handleEventKeydown(e) {
    if (e.key === 'Enter' || e.key === ' ') {
      e.preventDefault();
      selectEvent(e.currentTarget.dataset.id, e.currentTarget.dataset.name);
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
        // Only restore if not expired
        if (Date.now() - formData.timestamp < AppConfig.FORM_EXPIRY) {
          if (formData.name) el.nameInput.value = formData.name;
          if (formData.email) el.emailInput.value = formData.email;
          if (formData.phone) el.phoneInput.value = formData.phone;
          if (formData.eventId) {
            // Try to restore the selected event
            const event = AppState.findEventById(formData.eventId);
            if (event) {
              selectEvent(event.id, event.name);
            }
          }
        } else {
          // Clear old data
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

    const data = {
      anlassId: el.eventIdInput.value,
      name: el.nameInput.value.trim(),
      email: el.emailInput.value.trim(),
      telefon: el.phoneInput.value.trim(),
      identifier: getIdentifier()
    };

    setSubmitLoading(true);
    const startTime = performance.now();

    try {
      const result = await submitWithRetry(data, 0);

      // Record registration timing
      Metrics.recordRegistration(Math.round(performance.now() - startTime));

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

      if (!result.success && retryAttempt < AppConfig.MAX_RETRIES && (
        result.message && result.message.includes('Rate limit')
      )) {
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
