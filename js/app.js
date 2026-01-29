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

  function init() {
    if (!CONFIG.API_URL || CONFIG.API_URL === 'IHRE_GOOGLE_APPS_SCRIPT_URL_HIER') {
      showError('Bitte konfigurieren Sie die API_URL in index.html mit Ihrer Google Apps Script URL.');
      return;
    }
    loadEvents();
    setupFormValidation();
    setupKeyboardNavigation();
    restoreFormData();
    setupFormPersistence();
    setupSpreadsheetLink();
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
        data-time="${eventTime ? JSON.stringify(eventTime) : ''}"
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
        ${event.helferNamen && event.helferNamen.length > 0 ? `
        <div class="event-helpers">
          <span class="event-helpers-label">Bereits angemeldet:</span>
          <span class="event-helpers-list">${event.helferNamen.map(name => esc(name)).join(', ')}</span>
        </div>` : ''}
        <div class="event-actions">
          <div class="event-cta" aria-hidden="true">
            <span>Jetzt anmelden</span>
            <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
              <polyline points="9 18 15 12 9 6"/>
            </svg>
          </div>
          <button type="button" class="calendar-download-btn" 
            onclick="event.stopPropagation(); downloadCalendarEvent(this.closest('.event-card'));"
            aria-label="Anlass zum Kalender hinzufügen">
            <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
              <rect x="3" y="4" width="18" height="18" rx="2" ry="2"></rect>
              <line x1="16" y1="2" x2="16" y2="6"></line>
              <line x1="8" y1="2" x2="8" y2="6"></line>
              <line x1="3" y1="10" x2="21" y2="10"></line>
            </svg>
            <span>Kalendereintrag speichern</span>
          </button>
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

  function esc(text) {
    if (!text) return '';
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
  }

})();
