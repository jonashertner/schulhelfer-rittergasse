/**
 * SCHULHELFER · Admin UI v1
 *
 * Vanilla JS, no framework. Runs at /helferliste/admin/. Auth uses an
 * ADMIN_KEY stored in localStorage and sent in the POST body of every
 * request — never as a query parameter (those leak via referer / Pages
 * access logs).
 *
 * Public surface for the backend:
 *   POST { action: 'getAllEvents', adminKey }      -> { success, events }
 *   POST { action: 'addEvent', ..., adminKey }     -> { success, event }
 *   POST { action: 'updateEvent', id, ..., adminKey }
 *   POST { action: 'cancelEvent', id, adminKey }
 *
 * Auth failures clear the stored key and bounce the user back to the
 * login screen so an invalid/rotated key never sticks around.
 */
(function () {
  'use strict';

  const $ = (s) => document.querySelector(s);
  const $$ = (s) => Array.from(document.querySelectorAll(s));

  const STORAGE_KEY = 'schulhelfer_admin_key';

  const View = {
    login: $('#login-view'),
    dashboard: $('#dashboard-view'),
    eventModal: $('#event-modal'),
    confirmModal: $('#confirm-modal'),
  };

  const State = {
    events: [],
    editingEventId: null,
    helpersEvent: null,         // event currently displayed in the helpers modal
    helpersList: []              // last-loaded registrations for that event
  };

  // ============================================================
  // Boot
  // ============================================================
  document.addEventListener('DOMContentLoaded', init);

  async function init() {
    bindStaticHandlers();
    const key = localStorage.getItem(STORAGE_KEY);
    if (key) {
      showView('dashboard');
      const ok = await refreshEvents();
      if (!ok) {
        localStorage.removeItem(STORAGE_KEY);
        showView('login');
      }
    } else {
      showView('login');
    }
  }

  function bindStaticHandlers() {
    $('#login-form').addEventListener('submit', handleLogin);
    $('#logout-btn').addEventListener('click', handleLogout);
    $('#new-event-btn').addEventListener('click', () => openEventModal(null));
    $('#refresh-btn').addEventListener('click', () => refreshEvents());
    $('#event-form').addEventListener('submit', handleEventSubmit);

    // PR 4 additions
    if (CONFIG.SPREADSHEET_URL) $('#sheet-link').href = CONFIG.SPREADSHEET_URL;
    $('#integrity-btn').addEventListener('click', openIntegrity);
    $('#archive-btn').addEventListener('click', openArchive);
    $('#archive-confirm').addEventListener('change', (e) => { $('#archive-go-btn').disabled = !e.target.checked; });
    $('#archive-go-btn').addEventListener('click', handleArchiveSubmit);
    $('#manual-form').addEventListener('submit', handleManualSubmit);
    $('#helpers-mailto-btn').addEventListener('click', helpersMailto);
    $('#helpers-add-btn').addEventListener('click', openManualForCurrent);

    const allModals = ['event-modal', 'confirm-modal', 'helpers-modal', 'manual-modal', 'archive-modal', 'integrity-modal'];
    allModals.forEach(id => {
      $$('#' + id + ' [data-close]').forEach(el =>
        el.addEventListener('click', () => closeModal(id))
      );
    });

    document.addEventListener('keydown', (e) => {
      if (e.key === 'Escape') allModals.forEach(closeModal);
    });
  }

  function showView(name) {
    View.login.hidden = name !== 'login';
    View.dashboard.hidden = name !== 'dashboard';
  }

  // ============================================================
  // Auth
  // ============================================================
  async function handleLogin(e) {
    e.preventDefault();
    const keyInput = $('#login-key');
    const key = keyInput.value.trim();
    if (!key) return;

    setBusy('#login-btn', true);
    hide('#login-error');

    const result = await callAdmin({ action: 'getAllEvents' }, key);
    setBusy('#login-btn', false);

    if (result && result.success) {
      localStorage.setItem(STORAGE_KEY, key);
      keyInput.value = '';
      State.events = result.events || [];
      showView('dashboard');
      renderEvents();
    } else {
      show('#login-error', pickError(result, 'Anmeldung fehlgeschlagen.'));
    }
  }

  function handleLogout() {
    localStorage.removeItem(STORAGE_KEY);
    State.events = [];
    showView('login');
  }

  // ============================================================
  // Dashboard fetch + render
  // ============================================================
  async function refreshEvents() {
    showLoading(true);
    hideBanner('error-banner');

    const key = localStorage.getItem(STORAGE_KEY);
    if (!key) return false;

    const result = await callAdmin({ action: 'getAllEvents' }, key);
    showLoading(false);

    if (!result || !result.success) {
      const msg = pickError(result, 'Daten konnten nicht geladen werden.');
      if (result && result.error === 'Keine Berechtigung.') return false;
      showBanner('error-banner', msg);
      return true;
    }

    State.events = result.events || [];
    renderEvents();
    return true;
  }

  function renderEvents() {
    const today = todayMidnightTs();
    const upcoming = [];
    const cancelled = [];
    const past = [];

    State.events.forEach(ev => {
      const dateTs = ev.datum ? new Date(ev.datum).getTime() : NaN;
      if (ev.status === 'abgesagt') {
        cancelled.push(ev);
      } else if (!isNaN(dateTs) && dateTs < today) {
        past.push(ev);
      } else {
        upcoming.push(ev);
      }
    });

    upcoming.sort((a, b) => dateTsOf(a) - dateTsOf(b));
    past.sort((a, b) => dateTsOf(b) - dateTsOf(a));
    cancelled.sort((a, b) => dateTsOf(a) - dateTsOf(b));

    fillSection('upcoming', upcoming);
    fillSection('cancelled', cancelled);
    fillSection('past', past);
  }

  function fillSection(name, list) {
    const container = $('#' + name + '-list');
    const empty = $('#' + name + '-empty');
    const count = $('#count-' + name);
    container.replaceChildren();
    count.textContent = list.length;
    if (list.length === 0) {
      empty.hidden = false;
      return;
    }
    empty.hidden = true;
    list.forEach(ev => container.appendChild(buildEventCard(ev, name)));
  }

  // Build the card via explicit DOM construction (no innerHTML on
  // user-supplied content). Static structure is created with
  // createElement; user fields go through textContent.
  function buildEventCard(ev, section) {
    const card = el('article', 'event-card');
    if (ev.status === 'abgesagt') card.classList.add('is-cancelled');
    if (section === 'past') card.classList.add('is-past');

    const main = el('div', 'event-card-main');
    main.appendChild(textNode('h3', 'event-card-title', ev.name));

    const meta = el('div', 'event-card-meta');
    meta.appendChild(metaItem('📅 ' + (ev.datumDisplay || (ev.datum ? new Date(ev.datum).toLocaleDateString('de-CH') : '?'))));
    if (ev.zeit) meta.appendChild(metaItem('⏰ ' + ev.zeit));
    meta.appendChild(buildFillIndicator(ev));
    meta.appendChild(buildStatusBadge(ev));
    if (ev.sichtbar && ev.sichtbar !== 'ja' && ev.status !== 'abgesagt') {
      const hint = metaItem('⚠ ' + ev.sichtbar);
      hint.title = 'Auf der öffentlichen Seite: ' + ev.sichtbar;
      meta.appendChild(hint);
    }
    main.appendChild(meta);

    if (ev.beschreibung) {
      main.appendChild(textNode('p', 'event-card-desc', ev.beschreibung));
    }

    const actions = el('div', 'event-card-actions');
    // The helpers button is always available – the school master needs
    // the contact list even for past or cancelled events so they can
    // follow up with attendees.
    const helpersBtn = buildActionBtn(
      '👥 Helfer (' + ev.angemeldete + ')',
      'btn-secondary',
      () => openHelpers(ev)
    );
    actions.appendChild(helpersBtn);

    if (ev.status !== 'abgesagt') {
      actions.appendChild(buildActionBtn('Bearbeiten', 'btn-secondary', () => openEventModal(ev)));
      if (section !== 'past') {
        actions.appendChild(buildActionBtn('Absagen', 'btn-danger', () => confirmCancel(ev)));
      }
    }

    card.appendChild(main);
    card.appendChild(actions);
    return card;
  }

  function buildFillIndicator(ev) {
    const wrap = el('span', 'event-card-fill');
    wrap.title = 'Angemeldete Helfer';
    const bar = el('span', 'fill-bar');
    const inner = el('span', 'fill-bar-inner');
    const pct = ev.maxHelfer > 0 ? Math.min(1, ev.angemeldete / ev.maxHelfer) : 0;
    inner.style.width = (pct * 100).toFixed(0) + '%';
    if (pct >= 1) inner.classList.add('is-full');
    else if (pct < 0.5) inner.classList.add('is-low');
    bar.appendChild(inner);
    wrap.appendChild(bar);
    const label = document.createElement('span');
    label.textContent = ev.angemeldete + '/' + ev.maxHelfer;
    wrap.appendChild(label);
    return wrap;
  }

  function buildStatusBadge(ev) {
    const status = ev.status === 'abgesagt' ? 'abgesagt'
      : ev.status === 'archiviert' ? 'archiviert'
      : 'aktiv';
    const badge = el('span', 'event-card-status status-' + status);
    badge.textContent = status;
    return badge;
  }

  function buildActionBtn(label, cssClass, onClick) {
    const btn = el('button', 'btn ' + cssClass);
    btn.type = 'button';
    btn.textContent = label;
    btn.addEventListener('click', onClick);
    return btn;
  }

  function metaItem(text) {
    const span = el('span', 'event-card-meta-item');
    span.textContent = text;
    return span;
  }

  function textNode(tag, cssClass, text) {
    const node = document.createElement(tag);
    if (cssClass) node.className = cssClass;
    node.textContent = text;
    return node;
  }

  function el(tag, cssClass) {
    const node = document.createElement(tag);
    if (cssClass) node.className = cssClass;
    return node;
  }

  // ============================================================
  // Add / Edit modal
  // ============================================================
  function openEventModal(event) {
    State.editingEventId = event ? event.id : null;
    $('#event-modal-title').textContent = event ? 'Anlass bearbeiten' : 'Neuer Anlass';
    $('#event-id').value = event ? event.id : '';
    $('#event-name').value = event ? event.name : '';
    $('#event-date').value = event && event.datum ? event.datum.slice(0, 10) : '';
    $('#event-time').value = event ? event.zeit || '' : '';
    $('#event-helfer').value = event ? event.maxHelfer || '' : '';
    $('#event-desc').value = event ? event.beschreibung || '' : '';
    hide('#event-form-error');
    setBusy('#event-save-btn', false);
    openModal('event-modal');
    setTimeout(() => $('#event-name').focus(), 50);
  }

  async function handleEventSubmit(e) {
    e.preventDefault();
    const data = {
      name: $('#event-name').value.trim(),
      datum: $('#event-date').value,
      zeit: $('#event-time').value.trim(),
      helfer: parseInt($('#event-helfer').value, 10),
      beschreibung: $('#event-desc').value.trim()
    };
    if (!data.name)  return show('#event-form-error', 'Bitte einen Namen eingeben.');
    if (!data.datum) return show('#event-form-error', 'Bitte ein Datum wählen.');
    if (!Number.isFinite(data.helfer) || data.helfer < 1) {
      return show('#event-form-error', 'Bitte eine Helferzahl ≥ 1 angeben.');
    }

    setBusy('#event-save-btn', true);
    hide('#event-form-error');

    const key = localStorage.getItem(STORAGE_KEY);
    let payload, action;
    if (State.editingEventId) {
      action = 'updateEvent';
      payload = Object.assign({ action, id: State.editingEventId }, data);
    } else {
      action = 'addEvent';
      payload = Object.assign({ action }, data);
    }

    const result = await callAdmin(payload, key);
    setBusy('#event-save-btn', false);

    if (!result || !result.success) {
      show('#event-form-error', pickError(result, 'Speichern fehlgeschlagen.'));
      return;
    }

    closeModal('event-modal');
    flashSuccess(action === 'addEvent' ? 'Anlass erstellt.' : 'Anlass aktualisiert.');
    await refreshEvents();
  }

  // ============================================================
  // Cancel event
  // ============================================================
  function confirmCancel(event) {
    $('#confirm-title').textContent = 'Anlass absagen';
    $('#confirm-text').textContent =
      'Soll der Anlass «' + event.name + '» wirklich abgesagt werden? ' +
      'Bestehende Anmeldungen bleiben erhalten und Sie können die Helfer per E-Mail informieren. ' +
      'Sie können den Status später im Tabellenblatt zurück auf «aktiv» setzen, falls die Absage rückgängig gemacht werden soll.';
    const okBtn = $('#confirm-ok-btn');
    okBtn.textContent = 'Anlass absagen';
    const fresh = okBtn.cloneNode(true);
    okBtn.parentNode.replaceChild(fresh, okBtn);
    fresh.addEventListener('click', async () => {
      setBusy('#confirm-ok-btn', true);
      const key = localStorage.getItem(STORAGE_KEY);
      const result = await callAdmin({ action: 'cancelEvent', id: event.id }, key);
      setBusy('#confirm-ok-btn', false);
      if (result && result.success) {
        closeModal('confirm-modal');
        flashSuccess('Anlass abgesagt.');
        if (event.angemeldete > 0) maybeOfferEmailDraft(event);
        await refreshEvents();
      } else {
        showBanner('error-banner', pickError(result, 'Absage fehlgeschlagen.'));
        closeModal('confirm-modal');
      }
    });
    openModal('confirm-modal');
  }

  function maybeOfferEmailDraft(event) {
    showBanner('success-banner',
      'Anlass abgesagt. ' + event.angemeldete + ' angemeldete Helfer ' +
      'können Sie per Helferliste-Download (Word) auf der öffentlichen Seite kontaktieren.');
  }

  // ============================================================
  // Backend
  // ============================================================
  async function callAdmin(payload, key) {
    if (!key) return { success: false, error: 'Keine Berechtigung.' };
    const body = Object.assign({ adminKey: key }, payload);
    let res;
    try {
      res = await fetch(CONFIG.API_URL, {
        method: 'POST',
        redirect: 'follow',
        headers: { 'Content-Type': 'text/plain' },
        body: JSON.stringify(body)
      });
    } catch (err) {
      console.error('admin fetch failed', err);
      return { success: false, error: 'Netzwerkfehler. Bitte erneut versuchen.' };
    }
    if (!res.ok) {
      return { success: false, error: 'Server-Fehler ' + res.status };
    }
    // Distinguish JSON-parse failure (server returned HTML, e.g. an
    // out-of-date Apps Script deployment that doesn't recognise our
    // admin actions) from genuine `{success: false, error}` payloads.
    const text = await res.text();
    try {
      return JSON.parse(text);
    } catch (_) {
      console.error('non-JSON admin response:', text.slice(0, 300));
      return {
        success: false,
        error: 'Apps Script antwortet mit HTML statt JSON. Bitte das Script neu bereitstellen ' +
               '(Editor → Bereitstellung verwalten → Version: Neue Version).'
      };
    }
  }

  // Surface whichever message field the backend filled in. Older
  // doPost paths return {success:false, message:…} (parents-form
  // shape); admin paths return {success:false, error:…}. Showing
  // either keeps the diagnostic readable when the deployment is
  // mid-upgrade.
  function pickError(result, fallback) {
    if (!result) return fallback;
    return result.error || result.message || fallback;
  }

  // ============================================================
  // UI helpers
  // ============================================================
  function openModal(id) {
    document.getElementById(id).hidden = false;
    document.body.style.overflow = 'hidden';
  }
  function closeModal(id) {
    document.getElementById(id).hidden = true;
    document.body.style.overflow = '';
  }
  function setBusy(sel, busy) {
    const btn = $(sel);
    if (!btn) return;
    btn.disabled = busy;
    const label = btn.querySelector('.btn-label');
    const spinner = btn.querySelector('.btn-spinner');
    if (label) label.style.opacity = busy ? '0.6' : '1';
    if (spinner) spinner.hidden = !busy;
  }
  function show(sel, msg) {
    const el = $(sel);
    if (!el) return;
    if (msg !== undefined) el.textContent = msg;
    el.hidden = false;
  }
  function hide(sel) {
    const el = $(sel);
    if (el) el.hidden = true;
  }
  function showLoading(on) {
    $('#loading-banner').hidden = !on;
  }
  function showBanner(id, msg) {
    const el = document.getElementById(id);
    if (!el) return;
    el.textContent = msg;
    el.hidden = false;
    el.scrollIntoView({ behavior: 'smooth', block: 'center' });
  }
  function hideBanner(id) {
    const el = document.getElementById(id);
    if (el) el.hidden = true;
  }
  function flashSuccess(msg) {
    showBanner('success-banner', msg);
    setTimeout(() => hideBanner('success-banner'), 3500);
  }
  function todayMidnightTs() {
    const d = new Date(); d.setHours(0, 0, 0, 0); return d.getTime();
  }
  function dateTsOf(ev) {
    const t = ev.datum ? new Date(ev.datum).getTime() : NaN;
    return isNaN(t) ? Infinity : t;
  }

  // ============================================================
  // Helpers modal (per-event drawer)
  // ============================================================
  async function openHelpers(event) {
    State.helpersEvent = event;
    State.helpersList = [];
    $('#helpers-title').textContent = 'Helfer · ' + event.name;
    $('#helpers-event-meta').textContent =
      (event.datumDisplay || '') +
      (event.zeit ? ' · ' + event.zeit : '') +
      ' · ' + event.angemeldete + '/' + event.maxHelfer + ' angemeldet';
    $('#helpers-list').replaceChildren();
    hide('#helpers-error');
    $('#helpers-empty').hidden = true;
    openModal('helpers-modal');

    const key = localStorage.getItem(STORAGE_KEY);
    const result = await callAdmin({ action: 'getRegistrations', eventId: event.id }, key);
    if (!result || !result.success) {
      show('#helpers-error', pickError(result, 'Helferdaten konnten nicht geladen werden.'));
      return;
    }
    State.helpersList = result.registrations || [];
    renderHelpers();
  }

  function renderHelpers() {
    const container = $('#helpers-list');
    container.replaceChildren();
    if (State.helpersList.length === 0) {
      $('#helpers-empty').hidden = false;
      return;
    }
    $('#helpers-empty').hidden = true;
    State.helpersList.forEach(h => container.appendChild(buildHelperRow(h)));
  }

  function buildHelperRow(h) {
    const row = el('div', 'helper-row');
    if (h.status === 'storniert') row.classList.add('is-storniert');
    if (h.status === 'nicht erschienen') row.classList.add('is-noshow');

    const name = el('span', 'h-name');
    name.textContent = h.name;
    if (h.notizen) name.title = 'Notiz: ' + h.notizen;

    const emailWrap = el('span', 'h-email');
    if (h.email) {
      const a = document.createElement('a');
      a.href = 'mailto:' + h.email;
      a.textContent = h.email;
      emailWrap.appendChild(a);
    } else {
      emailWrap.textContent = '—';
    }

    const telWrap = el('span', 'h-tel');
    if (h.telefon) {
      const a = document.createElement('a');
      a.href = 'tel:' + h.telefon.replace(/\s/g, '');
      a.textContent = h.telefon;
      telWrap.appendChild(a);
    } else {
      telWrap.textContent = '—';
    }

    const select = document.createElement('select');
    ['aktiv', 'storniert', 'nicht erschienen'].forEach(s => {
      const o = document.createElement('option');
      o.value = s; o.textContent = s;
      if (h.status === s) o.selected = true;
      select.appendChild(o);
    });
    select.addEventListener('change', () => updateHelperStatus(h, select.value, select));

    const noteBtn = document.createElement('button');
    noteBtn.type = 'button';
    noteBtn.className = 'h-noteBtn';
    if (h.notizen) noteBtn.classList.add('has-note');
    noteBtn.textContent = h.notizen ? '✎ Notiz' : '+ Notiz';
    noteBtn.addEventListener('click', () => toggleNoteEditor(row, h));

    row.appendChild(name);
    row.appendChild(emailWrap);
    row.appendChild(telWrap);
    row.appendChild(select);
    row.appendChild(noteBtn);
    return row;
  }

  function toggleNoteEditor(row, h) {
    const existing = row.querySelector('.h-note-edit');
    if (existing) { existing.remove(); return; }

    const editor = el('div', 'h-note-edit');
    const input = document.createElement('input');
    input.type = 'text';
    input.value = h.notizen || '';
    input.placeholder = 'Notiz (optional)';
    input.maxLength = 500;
    const save = el('button', 'btn btn-secondary');
    save.type = 'button';
    save.textContent = 'Speichern';
    save.addEventListener('click', async () => {
      save.disabled = true;
      const key = localStorage.getItem(STORAGE_KEY);
      const result = await callAdmin({
        action: 'updateRegistration',
        eventId: h.eventId, email: h.email,
        notizen: input.value
      }, key);
      save.disabled = false;
      if (!result || !result.success) {
        show('#helpers-error', pickError(result, 'Notiz konnte nicht gespeichert werden.'));
        return;
      }
      h.notizen = input.value;
      renderHelpers();
    });
    editor.appendChild(input);
    editor.appendChild(save);
    row.appendChild(editor);
    setTimeout(() => input.focus(), 30);
  }

  async function updateHelperStatus(h, newStatus, selectEl) {
    if (newStatus === h.status) return;
    selectEl.disabled = true;
    const key = localStorage.getItem(STORAGE_KEY);
    const result = await callAdmin({
      action: 'updateRegistration',
      eventId: h.eventId, email: h.email,
      status: newStatus
    }, key);
    selectEl.disabled = false;
    if (!result || !result.success) {
      show('#helpers-error', pickError(result, 'Status konnte nicht geändert werden.'));
      selectEl.value = h.status;
      return;
    }
    h.status = newStatus;
    renderHelpers();
    // Refresh dashboard – capacity counts may have changed (storniert
    // frees a slot, aktiv re-claims one).
    refreshEvents();
  }

  function helpersMailto() {
    const emails = State.helpersList
      .filter(h => h.email && h.status !== 'storniert')
      .map(h => h.email);
    if (emails.length === 0) {
      show('#helpers-error', 'Keine aktiven Anmeldungen mit E-Mail-Adresse.');
      return;
    }
    const subject = encodeURIComponent('Schulhelfer · ' + (State.helpersEvent ? State.helpersEvent.name : ''));
    // Use BCC so addresses stay private from each other.
    const href = 'mailto:?bcc=' + emails.join(',') + '&subject=' + subject;
    window.location.href = href;
  }

  function openManualForCurrent() {
    if (!State.helpersEvent) return;
    $('#manual-event-id').value = State.helpersEvent.id;
    $('#manual-event-name').textContent = 'Anlass: ' + State.helpersEvent.name;
    $('#manual-name').value = '';
    $('#manual-email').value = '';
    $('#manual-tel').value = '';
    hide('#manual-error');
    setBusy('#manual-save-btn', false);
    openModal('manual-modal');
    setTimeout(() => $('#manual-name').focus(), 50);
  }

  async function handleManualSubmit(e) {
    e.preventDefault();
    const data = {
      action: 'addRegistration',
      anlassId: $('#manual-event-id').value,
      name: $('#manual-name').value.trim(),
      email: $('#manual-email').value.trim().toLowerCase(),
      telefon: $('#manual-tel').value.trim()
    };
    if (!data.name)  return show('#manual-error', 'Bitte einen Namen eingeben.');
    if (!data.email) return show('#manual-error', 'Bitte eine E-Mail-Adresse eingeben.');

    setBusy('#manual-save-btn', true);
    hide('#manual-error');
    const key = localStorage.getItem(STORAGE_KEY);
    const result = await callAdmin(data, key);
    setBusy('#manual-save-btn', false);

    if (!result || !result.success) {
      show('#manual-error', pickError(result, 'Eintrag fehlgeschlagen.'));
      return;
    }
    closeModal('manual-modal');
    flashSuccess('Helfer eingetragen.');
    // Reload the helper list and dashboard
    if (State.helpersEvent) await openHelpers(State.helpersEvent);
    refreshEvents();
  }

  // ============================================================
  // Archive
  // ============================================================
  async function openArchive() {
    $('#archive-confirm').checked = false;
    $('#archive-go-btn').disabled = true;
    setBusy('#archive-go-btn', false);
    hide('#archive-error');
    const select = $('#archive-year');
    select.replaceChildren();
    const placeholder = document.createElement('option');
    placeholder.textContent = 'Wird geladen…';
    placeholder.disabled = true; placeholder.selected = true;
    select.appendChild(placeholder);
    openModal('archive-modal');

    const key = localStorage.getItem(STORAGE_KEY);
    const result = await callAdmin({ action: 'availableJahre' }, key);
    select.replaceChildren();
    if (!result || !result.success) {
      show('#archive-error', pickError(result, 'Jahre konnten nicht geladen werden.'));
      return;
    }
    if (!result.years || result.years.length === 0) {
      const opt = document.createElement('option');
      opt.textContent = 'Keine archivierbaren Jahre vorhanden.';
      opt.disabled = true; opt.selected = true;
      select.appendChild(opt);
      $('#archive-confirm').disabled = true;
      return;
    }
    $('#archive-confirm').disabled = false;
    result.years.forEach((y, idx) => {
      const opt = document.createElement('option');
      opt.value = y.jahr;
      opt.textContent = y.jahr + ' · ' + y.events + ' Anlässe, ' + y.registrations + ' Anmeldungen';
      // Default to the most recent past year (last in the sorted array).
      if (idx === result.years.length - 1) opt.selected = true;
      select.appendChild(opt);
    });
  }

  async function handleArchiveSubmit() {
    const jahr = $('#archive-year').value;
    if (!jahr) return;
    setBusy('#archive-go-btn', true);
    hide('#archive-error');
    const key = localStorage.getItem(STORAGE_KEY);
    const result = await callAdmin({ action: 'archiveJahr', jahr: jahr }, key);
    setBusy('#archive-go-btn', false);
    if (!result || !result.success) {
      show('#archive-error', pickError(result, 'Archivierung fehlgeschlagen.'));
      return;
    }
    closeModal('archive-modal');
    flashSuccess(result.message || ('Jahr ' + jahr + ' archiviert.'));
    refreshEvents();
  }

  // ============================================================
  // Integrity check
  // ============================================================
  async function openIntegrity() {
    $('#integrity-loading').hidden = false;
    $('#integrity-result').replaceChildren();
    openModal('integrity-modal');
    const key = localStorage.getItem(STORAGE_KEY);
    const result = await callAdmin({ action: 'integrityCheck' }, key);
    $('#integrity-loading').hidden = true;
    const target = $('#integrity-result');
    if (!result || !result.success) {
      const div = el('div', 'integrity-line is-warn');
      div.textContent = pickError(result, 'Prüfung fehlgeschlagen.');
      target.appendChild(div);
      return;
    }
    if (!result.findings || result.findings.length === 0) {
      const div = el('div', 'integrity-line is-ok');
      div.textContent = '✅ Alles sauber. Keine Auffälligkeiten gefunden.';
      target.appendChild(div);
      return;
    }
    result.findings.forEach(line => {
      const div = el('div', 'integrity-line');
      if (line.indexOf('⚠') === 0) div.classList.add('is-warn');
      div.textContent = line;
      target.appendChild(div);
    });
  }
})();
