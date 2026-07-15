/**
 * Firebase Realtime Database presence — đếm tab đang online + list.
 * Config: window.PRESENCE_FIREBASE_CONFIG (presence-config.js). Hướng dẫn: PRESENCE.md
 */
(function () {
    'use strict';

    const HEARTBEAT_MS = 20000;
    const OFFLINE_HOLD_MS = 45000;
    const PATH = 'presence';

    const recentlyOffline = new Map();
    let sessionId = '';
    let sessionRef = null;
    let heartbeatTimer = 0;
    let publicIp = '';
    let db = null;
    let started = false;

    function el(id) {
        return document.getElementById(id);
    }

    function shortCode(id) {
        return String(id || '').replace(/\W/g, '').slice(-4) || '----';
    }

    function makeSessionId() {
        try {
            if (crypto && typeof crypto.randomUUID === 'function') {
                return crypto.randomUUID();
            }
        } catch (e) { /* ignore */ }
        return 's' + Date.now().toString(36) + Math.random().toString(36).slice(2, 10);
    }

    function isConfigReady(cfg) {
        if (!cfg || typeof cfg !== 'object') {
            return false;
        }
        const key = String(cfg.apiKey || '');
        const url = String(cfg.databaseURL || '');
        if (!key || key.indexOf('PASTE_') === 0) {
            return false;
        }
        if (!url || url.indexOf('PASTE_') !== -1) {
            return false;
        }
        return true;
    }

    function setWidgetVisible(on) {
        const widget = el('presenceWidget');
        if (widget) {
            widget.hidden = !on;
        }
    }

    function showSetupHint(msg) {
        const count = el('presenceCount');
        const btn = el('presenceBtn');
        setWidgetVisible(true);
        if (count) {
            count.textContent = '—';
        }
        if (btn) {
            btn.title = msg || 'Chưa cấu hình Firebase — xem PRESENCE.md';
            btn.setAttribute('aria-label', btn.title);
        }
        const list = el('presenceList');
        if (list) {
            list.innerHTML = '<li class="presence-item presence-item--hint">' +
                escapeHtml(msg || 'Chưa cấu hình Firebase. Xem PRESENCE.md') + '</li>';
        }
    }

    function escapeHtml(text) {
        return String(text || '')
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;');
    }

    function buildLabel() {
        const code = shortCode(sessionId);
        if (publicIp) {
            return publicIp + ' · ' + code;
        }
        return 'tab · ' + code;
    }

    function fetchPublicIp() {
        return fetch('https://api.ipify.org?format=json', { cache: 'no-store' })
            .then((r) => (r.ok ? r.json() : null))
            .then((j) => {
                if (j && j.ip) {
                    publicIp = String(j.ip);
                }
            })
            .catch(() => { /* offline / blocked */ });
    }

    function writePresence(online) {
        if (!sessionRef) {
            return Promise.resolve();
        }
        return sessionRef.set({
            online: !!online,
            label: buildLabel(),
            updatedAt: Date.now()
        });
    }

    function pruneRecentlyOffline() {
        const now = Date.now();
        recentlyOffline.forEach((entry, id) => {
            if (!entry || now - entry.at > OFFLINE_HOLD_MS) {
                recentlyOffline.delete(id);
            }
        });
    }

    function renderPresence(snapVal) {
        pruneRecentlyOffline();
        const onlineRows = [];
        const data = snapVal && typeof snapVal === 'object' ? snapVal : {};
        const nextOnlineIds = new Set();
        const labelById = renderPresence._labelById || {};

        Object.keys(data).forEach((id) => {
            const row = data[id] || {};
            const label = String(row.label || id);
            labelById[id] = label;
            if (row.online) {
                nextOnlineIds.add(id);
                onlineRows.push({
                    id,
                    label,
                    updatedAt: Number(row.updatedAt) || 0
                });
                recentlyOffline.delete(id);
            }
        });

        const prevOnline = renderPresence._prevOnlineIds || new Set();
        prevOnline.forEach((id) => {
            if (!nextOnlineIds.has(id) && !recentlyOffline.has(id)) {
                recentlyOffline.set(id, {
                    label: labelById[id] || id,
                    at: Date.now()
                });
            }
        });
        renderPresence._prevOnlineIds = nextOnlineIds;
        renderPresence._labelById = labelById;

        onlineRows.sort((a, b) => a.label.localeCompare(b.label) || a.id.localeCompare(b.id));

        const offlineRows = [];
        recentlyOffline.forEach((entry, id) => {
            if (nextOnlineIds.has(id)) {
                return;
            }
            offlineRows.push({ id, label: entry.label, at: entry.at });
        });
        offlineRows.sort((a, b) => b.at - a.at);

        const countEl = el('presenceCount');
        if (countEl) {
            countEl.textContent = String(onlineRows.length);
        }
        const btn = el('presenceBtn');
        if (btn) {
            const title = onlineRows.length + ' người đang online — bấm để xem list';
            btn.title = title;
            btn.setAttribute('aria-label', title);
        }

        const list = el('presenceList');
        if (!list) {
            return;
        }
        let html = '';
        onlineRows.forEach((r) => {
            html += '<li class="presence-item presence-item--online">' +
                '<span class="presence-dot" aria-hidden="true"></span>' +
                '<span class="presence-label">' + escapeHtml(r.label) + ' đang online</span></li>';
        });
        offlineRows.forEach((r) => {
            html += '<li class="presence-item presence-item--offline">' +
                '<span class="presence-dot presence-dot--offline" aria-hidden="true"></span>' +
                '<span class="presence-label">' + escapeHtml(r.label) + ' vừa offline</span></li>';
        });
        if (!html) {
            html = '<li class="presence-item presence-item--hint">Chưa có ai online</li>';
        }
        list.innerHTML = html;
    }

    function bindUi() {
        const btn = el('presenceBtn');
        const panel = el('presencePanel');
        if (!btn || !panel || btn.dataset.presenceBound === '1') {
            return;
        }
        btn.dataset.presenceBound = '1';
        btn.addEventListener('click', (e) => {
            e.stopPropagation();
            panel.classList.toggle('hidden');
            btn.setAttribute('aria-expanded', panel.classList.contains('hidden') ? 'false' : 'true');
        });
        document.addEventListener('click', (e) => {
            const widget = el('presenceWidget');
            if (!widget || !panel || panel.classList.contains('hidden')) {
                return;
            }
            if (widget.contains(e.target)) {
                return;
            }
            panel.classList.add('hidden');
            btn.setAttribute('aria-expanded', 'false');
        });
    }

    function startHeartbeat() {
        if (heartbeatTimer) {
            clearInterval(heartbeatTimer);
        }
        heartbeatTimer = setInterval(() => {
            writePresence(true).catch(() => { /* ignore */ });
        }, HEARTBEAT_MS);
    }

    function startPresence() {
        if (started) {
            return;
        }
        const cfg = window.PRESENCE_FIREBASE_CONFIG;
        if (!isConfigReady(cfg)) {
            bindUi();
            showSetupHint('Chưa cấu hình Firebase — mở PRESENCE.md để tạo project và dán config vào presence-config.js');
            return;
        }
        if (typeof firebase === 'undefined' || !firebase.initializeApp) {
            bindUi();
            showSetupHint('Không tải được Firebase SDK');
            return;
        }

        started = true;
        bindUi();
        setWidgetVisible(true);

        try {
            if (!firebase.apps.length) {
                firebase.initializeApp(cfg);
            }
        } catch (e) {
            showSetupHint('Firebase config lỗi: ' + (e && e.message ? e.message : e));
            return;
        }

        const afterAuth = () => {
            db = firebase.database();
            sessionId = makeSessionId();
            sessionRef = db.ref(PATH + '/' + sessionId);

            const connectedRef = db.ref('.info/connected');
            connectedRef.on('value', (snap) => {
                if (snap.val() !== true) {
                    return;
                }
                sessionRef.onDisconnect().remove()
                    .then(() => writePresence(true))
                    .then(() => startHeartbeat())
                    .catch(() => { /* ignore */ });
            });

            db.ref(PATH).on('value', (snap) => {
                renderPresence(snap.val());
            });

            fetchPublicIp().then(() => writePresence(true)).catch(() => { /* ignore */ });

            window.addEventListener('pagehide', () => {
                try {
                    if (sessionRef) {
                        sessionRef.onDisconnect().cancel();
                        sessionRef.remove();
                    }
                } catch (e) { /* ignore */ }
            });
        };

        if (firebase.auth) {
            firebase.auth().signInAnonymously()
                .then(afterAuth)
                .catch((err) => {
                    // Test-mode rules may allow unauthenticated access
                    console.warn('[presence] Anonymous auth failed, trying without auth', err);
                    afterAuth();
                });
        } else {
            afterAuth();
        }
    }

    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', startPresence);
    } else {
        startPresence();
    }
})();
