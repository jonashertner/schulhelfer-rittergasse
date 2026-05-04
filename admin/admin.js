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
    editingEventId: null
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

    $$('#event-modal [data-close]').forEach(el => el.addEventListener('click', () => closeModal('event-modal')));
    $$('#confirm-modal [data-close]').forEach(el => el.addEventListener('click', () => closeModal('confirm-modal')));

    document.addEventListener('keydown', (e) => {
      if (e.key === 'Escape') {
        closeModal('event-modal');
        closeModal('confirm-modal');
      }
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
      show('#login-error', (result && result.error) || 'Anmeldung fehlgeschlagen.');
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
      const msg = (result && result.error) || 'Daten konnten nicht geladen werden.';
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
    if (ev.status !== 'abgesagt') {
      actions.appendChild(buildActionBtn('Bearbeiten', 'btn-secondary', () => openEventModal(ev)));
      if (section !== 'past') {
        actions.appendChild(buildActionBtn('Absagen', 'btn-danger', () => confirmCancel(ev)));
      }
    } else {
      const tag = el('span', 'admin-count');
      tag.textContent = ev.angemeldete + ' Anmeldungen';
      actions.appendChild(tag);
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
      const msg = (result && (result.error || result.message)) || 'Speichern fehlgeschlagen.';
      show('#event-form-error', msg);
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
        const msg = (result && (result.error || result.message)) || 'Absage fehlgeschlagen.';
        showBanner('error-banner', msg);
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
    try {
      const res = await fetch(CONFIG.API_URL, {
        method: 'POST',
        redirect: 'follow',
        headers: { 'Content-Type': 'text/plain' },
        body: JSON.stringify(body)
      });
      if (!res.ok) {
        return { success: false, error: 'Server-Fehler ' + res.status };
      }
      return await res.json();
    } catch (err) {
      console.error('admin call failed', err);
      return { success: false, error: 'Netzwerkfehler. Bitte erneut versuchen.' };
    }
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
})();
