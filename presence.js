/**
 * Firebase Realtime Database presence — đếm tab đang online + list.
 * Config: window.PRESENCE_FIREBASE_CONFIG (presence-config.js). Hướng dẫn: PRESENCE.md
 */
(function () {
    'use strict';

    const HEARTBEAT_MS = 20000;
    /** Không heartbeat quá lâu → coi tab chết (onDisconnect đôi khi không chạy). */
    const STALE_ONLINE_MS = 90 * 1000;
    const OFFLINE_HOLD_MS = 5 * 60 * 1000;
    const RENDER_TICK_MS = 15000;
    const PATH = 'presence';

    const recentlyOffline = new Map();
    let sessionId = '';
    let displayCode = '';
    let startedAt = 0;
    let sessionRef = null;
    let heartbeatTimer = 0;
    let renderTickTimer = 0;
    let publicIp = '';
    let lastSnapVal = null;
    let db = null;
    let started = false;

    function el(id) {
        return document.getElementById(id);
    }

    /** Mã hiển thị 3 ký tự: 2 chữ cái + 1 chữ số (vd. AG3), ổn định theo sessionId. */
    function makeDisplayCode(id) {
        const s = String(id || '').replace(/\W/g, '') || 'x';
        let h = 2166136261;
        for (let i = 0; i < s.length; i++) {
            h ^= s.charCodeAt(i);
            h = Math.imul(h, 16777619);
        }
        h = h >>> 0;
        const a = String.fromCharCode(65 + (h % 26));
        const b = String.fromCharCode(65 + ((h >>> 5) % 26));
        const d = String((h >>> 10) % 10);
        return a + b + d;
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

    function pad2(n) {
        return n < 10 ? '0' + n : String(n);
    }

    function formatAccessTime(ts) {
        const t = Number(ts) || 0;
        if (!t) {
            return '—';
        }
        const d = new Date(t);
        return pad2(d.getDate()) + '/' + pad2(d.getMonth() + 1) +
            ' ' + pad2(d.getHours()) + ':' + pad2(d.getMinutes());
    }

    function formatDuration(ms) {
        const totalSec = Math.max(0, Math.floor(Number(ms) / 1000));
        if (totalSec < 60) {
            return totalSec + ' giây';
        }
        const totalMin = Math.floor(totalSec / 60);
        if (totalMin < 60) {
            return totalMin + ' phút';
        }
        const h = Math.floor(totalMin / 60);
        const m = totalMin % 60;
        if (h < 24) {
            return m ? (h + ' giờ ' + m + ' phút') : (h + ' giờ');
        }
        const days = Math.floor(h / 24);
        const remH = h % 24;
        return remH ? (days + ' ngày ' + remH + ' giờ') : (days + ' ngày');
    }

    function buildLabel() {
        const code = displayCode || makeDisplayCode(sessionId);
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
            code: displayCode || makeDisplayCode(sessionId),
            startedAt: startedAt || Date.now(),
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

    function metaLine(started, endedAt) {
        const start = Number(started) || 0;
        const end = Number(endedAt) || Date.now();
        const access = formatAccessTime(start);
        const dur = start ? formatDuration(end - start) : '—';
        return 'Vào ' + access + ' · đã ' + dur;
    }

    function renderPresence(snapVal) {
        lastSnapVal = snapVal;
        pruneRecentlyOffline();
        const onlineRows = [];
        const data = snapVal && typeof snapVal === 'object' ? snapVal : {};
        const nextOnlineIds = new Set();
        const metaById = renderPresence._metaById || {};
        const now = Date.now();

        Object.keys(data).forEach((id) => {
            const row = data[id] || {};
            const label = String(row.label || id);
            const rowStarted = Number(row.startedAt) || Number(row.updatedAt) || 0;
            const updatedAt = Number(row.updatedAt) || 0;
            metaById[id] = { label: label, startedAt: rowStarted };
            const alive = !!row.online && updatedAt > 0 && (now - updatedAt) <= STALE_ONLINE_MS;
            if (row.online && !alive) {
                // Ghost: tab chết / code cũ / onDisconnect không xóa → gỡ khỏi Firebase
                if (db && id !== sessionId) {
                    try {
                        db.ref(PATH + '/' + id).remove();
                    } catch (e) { /* ignore */ }
                }
            }
            if (alive) {
                nextOnlineIds.add(id);
                onlineRows.push({
                    id: id,
                    label: label,
                    startedAt: rowStarted,
                    updatedAt: updatedAt
                });
                recentlyOffline.delete(id);
            }
        });

        const prevOnline = renderPresence._prevOnlineIds || new Set();
        prevOnline.forEach((id) => {
            if (!nextOnlineIds.has(id) && !recentlyOffline.has(id)) {
                const meta = metaById[id] || {};
                recentlyOffline.set(id, {
                    label: meta.label || id,
                    startedAt: meta.startedAt || 0,
                    at: Date.now()
                });
            }
        });
        renderPresence._prevOnlineIds = nextOnlineIds;
        renderPresence._metaById = metaById;

        // Vào sớm nhất lên đầu; vào sau append xuống dưới
        onlineRows.sort((a, b) => {
            const sa = Number(a.startedAt) || Number(a.updatedAt) || 0;
            const sb = Number(b.startedAt) || Number(b.updatedAt) || 0;
            if (sa !== sb) {
                return sa - sb;
            }
            return a.id.localeCompare(b.id);
        });

        const offlineRows = [];
        recentlyOffline.forEach((entry, id) => {
            if (nextOnlineIds.has(id)) {
                return;
            }
            offlineRows.push({
                id: id,
                label: entry.label,
                startedAt: entry.startedAt || 0,
                at: entry.at
            });
        });
        offlineRows.sort((a, b) => {
            if (a.at !== b.at) {
                return a.at - b.at;
            }
            return a.id.localeCompare(b.id);
        });

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
                '<span class="presence-text">' +
                '<span class="presence-label">' + escapeHtml(r.label) + ' đang online</span>' +
                '<span class="presence-meta">' + escapeHtml(metaLine(r.startedAt)) + '</span>' +
                '</span></li>';
        });
        offlineRows.forEach((r) => {
            html += '<li class="presence-item presence-item--offline">' +
                '<span class="presence-dot presence-dot--offline" aria-hidden="true"></span>' +
                '<span class="presence-text">' +
                '<span class="presence-label">' + escapeHtml(r.label) + ' vừa offline</span>' +
                '<span class="presence-meta">' + escapeHtml(metaLine(r.startedAt, r.at)) + '</span>' +
                '</span></li>';
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

    function startRenderTick() {
        if (renderTickTimer) {
            clearInterval(renderTickTimer);
        }
        renderTickTimer = setInterval(() => {
            if (lastSnapVal != null) {
                renderPresence(lastSnapVal);
            }
        }, RENDER_TICK_MS);
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
            displayCode = makeDisplayCode(sessionId);
            startedAt = Date.now();
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
            startRenderTick();

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
