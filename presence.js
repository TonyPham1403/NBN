/**
 * Firebase Realtime Database presence — đếm tab đang online + list.
 * Config: window.PRESENCE_FIREBASE_CONFIG (presence-config.js). Hướng dẫn: PRESENCE.md
 *
 * Offline ~5 phút ghi trên Firebase (online:false + offlineAt). Tab còn mở có thể
 * tái tạo tombstone nếu peer bị xóa (code cũ / cancel onDisconnect) để người vào sau vẫn thấy.
 * GeoIP (geojs.io): quốc gia + lat/lon ghi kèm presence; khoảng cách Haversine tới (You).
 * Device ID bền (localStorage) — mọi tab cùng trình duyệt/máy chung một deviceId (chat sau này).
 */
(function () {
    'use strict';

    const HEARTBEAT_MS = 20000;
    const STALE_ONLINE_MS = 90 * 1000;
    const OFFLINE_HOLD_MS = 5 * 60 * 1000;
    const RENDER_TICK_MS = 15000;
    const PATH = 'presence';
    const DEVICE_STORAGE_KEY = 'presenceDeviceId';
    const GEO_SELF_URL = 'https://get.geojs.io/v1/ip/geo.json';
    const GEO_IP_URL = 'https://get.geojs.io/v1/ip/geo/';

    let sessionId = '';
    let displayCode = '';
    /** UUID bền theo trình duyệt/máy — chung mọi tab trên thiết bị này. */
    let deviceId = '';
    /** Mã ngắn hiển thị thiết bị (vd. D7K2), khác mã tab OT6. */
    let deviceTag = '';
    let startedAt = 0;
    let sessionRef = null;
    let heartbeatTimer = 0;
    let renderTickTimer = 0;
    let publicIp = '';
    /** @type {{ country: string, countryCode: string, region: string, city: string, lat: number, lon: number }|null} */
    let selfGeo = null;
    let lastSnapVal = null;
    let db = null;
    let started = false;
    const pruneRequested = new Set();
    const tombstoneRequested = new Set();
    /** meta gần nhất theo id — để tái tạo tombstone khi node bị xóa khỏi RTDB */
    let metaById = {};
    let prevOnlineIds = new Set();
    /** Offline tạm khi peer bị remove — giữ đến khi RTDB có tombstone */
    const pendingGone = new Map();
    /** Cache GeoIP theo IP (lookup peer thiếu country/lat trên Firebase) */
    const geoCache = new Map();
    const geoFetchInflight = new Set();

    function el(id) {
        return document.getElementById(id);
    }

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

    /** Mã thiết bị ngắn (D + 3 ký tự) từ deviceId — ổn định, khác mã tab 3 ký tự. */
    function makeDeviceTag(id) {
        const s = String(id || '').replace(/\W/g, '') || 'x';
        let h = 2166136261;
        for (let i = 0; i < s.length; i++) {
            h ^= s.charCodeAt(i);
            h = Math.imul(h, 16777619);
        }
        h = h >>> 0;
        const alphabet = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
        let tag = 'D';
        for (let i = 0; i < 3; i++) {
            tag += alphabet[(h >>> (i * 5)) % alphabet.length];
        }
        return tag;
    }

    function getOrCreateDeviceId() {
        try {
            let id = localStorage.getItem(DEVICE_STORAGE_KEY);
            if (id && String(id).length >= 8) {
                return String(id);
            }
            id = makeSessionId();
            localStorage.setItem(DEVICE_STORAGE_KEY, id);
            return id;
        } catch (e) {
            return 'd' + Date.now().toString(36) + Math.random().toString(36).slice(2, 12);
        }
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

    /** Icon người — trước mã tên (OT6); online xanh / offline đỏ. */
    function personIconHtml(offline) {
        const cls = offline ? 'presence-person presence-person--offline' : 'presence-person';
        return '<span class="' + cls + '" aria-hidden="true">' +
            '<svg viewBox="0 0 24 24" width="12" height="12" focusable="false">' +
            '<circle cx="12" cy="8" r="3.2" fill="currentColor"/>' +
            '<path d="M5.5 19.2c.6-3.4 3.2-5.2 6.5-5.2s5.9 1.8 6.5 5.2" ' +
            'fill="none" stroke="currentColor" stroke-width="2.2" ' +
            'stroke-linecap="round"/></svg></span>';
    }

    /** `IP · OT6` → `IP · [icon]OT6` */
    function formatLabelWithPerson(label, offline) {
        const raw = String(label || '');
        const sep = ' · ';
        const idx = raw.lastIndexOf(sep);
        const icon = personIconHtml(!!offline);
        if (idx === -1) {
            return icon + escapeHtml(raw);
        }
        return escapeHtml(raw.slice(0, idx)) + sep + icon + escapeHtml(raw.slice(idx + sep.length));
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

    function parseGeoFromPayload(j) {
        if (!j || typeof j !== 'object') {
            return null;
        }
        const lat = Number(j.latitude != null ? j.latitude : j.lat);
        const lon = Number(j.longitude != null ? j.longitude : j.lon);
        if (!Number.isFinite(lat) || !Number.isFinite(lon)) {
            return null;
        }
        return {
            country: String(j.country || '').trim(),
            countryCode: String(j.country_code || j.countryCode || '').trim(),
            region: String(j.region || j.regionName || '').trim(),
            city: String(j.city || '').trim(),
            lat: lat,
            lon: lon
        };
    }

    /** Quốc gia · tỉnh/thành (region ưu tiên, không trùng city). */
    function formatGeoPlace(geo) {
        if (!geo) {
            return '—';
        }
        const country = geo.country || geo.countryCode || '—';
        const region = String(geo.region || '').trim();
        const city = String(geo.city || '').trim();
        let area = region || city;
        if (region && city) {
            const r = region.toLowerCase();
            const c = city.toLowerCase();
            if (c !== r && c.indexOf(r) === -1 && r.indexOf(c) === -1) {
                area = region + ', ' + city;
            }
        }
        return area ? (country + ' · ' + area) : country;
    }

    function haversineKm(lat1, lon1, lat2, lon2) {
        const toRad = (d) => (d * Math.PI) / 180;
        const R = 6371;
        const dLat = toRad(lat2 - lat1);
        const dLon = toRad(lon2 - lon1);
        const a = Math.sin(dLat / 2) * Math.sin(dLat / 2) +
            Math.cos(toRad(lat1)) * Math.cos(toRad(lat2)) *
            Math.sin(dLon / 2) * Math.sin(dLon / 2);
        return R * 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
    }

    function formatCoord(n) {
        const v = Number(n);
        if (!Number.isFinite(v)) {
            return '—';
        }
        return v.toFixed(2);
    }

    function formatDistanceKm(km) {
        if (!Number.isFinite(km)) {
            return '';
        }
        if (km < 1) {
            return Math.round(km * 1000) + ' m';
        }
        if (km < 100) {
            return km.toFixed(1).replace(/\.0$/, '') + ' km';
        }
        return Math.round(km) + ' km';
    }

    function extractIpFromLabel(label) {
        const m = String(label || '').match(/\b(\d{1,3}(?:\.\d{1,3}){3})\b/);
        return m ? m[1] : '';
    }

    function geoFromRow(row) {
        if (!row || typeof row !== 'object') {
            return null;
        }
        const lat = Number(row.lat);
        const lon = Number(row.lon);
        if (Number.isFinite(lat) && Number.isFinite(lon)) {
            return {
                country: String(row.country || '').trim(),
                countryCode: String(row.countryCode || '').trim(),
                region: String(row.region || '').trim(),
                city: String(row.city || '').trim(),
                lat: lat,
                lon: lon
            };
        }
        const ip = extractIpFromLabel(row.label) || String(row.ip || '');
        if (ip && geoCache.has(ip)) {
            return geoCache.get(ip);
        }
        return null;
    }

    function ensureGeoForIp(ip) {
        const key = String(ip || '').trim();
        if (!key || geoCache.has(key) || geoFetchInflight.has(key)) {
            return;
        }
        geoFetchInflight.add(key);
        fetch(GEO_IP_URL + encodeURIComponent(key) + '.json', { cache: 'no-store' })
            .then((r) => (r.ok ? r.json() : null))
            .then((j) => {
                const g = parseGeoFromPayload(j);
                if (g) {
                    geoCache.set(key, g);
                    if (lastSnapVal != null) {
                        renderPresence(lastSnapVal);
                    }
                }
            })
            .catch(() => { /* ignore */ })
            .then(() => {
                geoFetchInflight.delete(key);
            });
    }

    function attachGeoFields(payload) {
        const out = payload || {};
        if (deviceId) {
            out.deviceId = deviceId;
            out.deviceTag = deviceTag || makeDeviceTag(deviceId);
        }
        if (publicIp) {
            out.ip = publicIp;
        }
        if (selfGeo) {
            out.country = selfGeo.country || '';
            out.countryCode = selfGeo.countryCode || '';
            out.region = selfGeo.region || '';
            out.city = selfGeo.city || '';
            out.lat = selfGeo.lat;
            out.lon = selfGeo.lon;
        }
        return out;
    }

    function deviceMetaLine(row, deviceCounts) {
        const tag = String((row && row.deviceTag) || '').trim()
            || (row && row.deviceId ? makeDeviceTag(row.deviceId) : '');
        if (!tag) {
            return '';
        }
        const did = String((row && row.deviceId) || '');
        const count = did && deviceCounts ? (deviceCounts[did] || 0) : 0;
        let line = 'Thiết bị ' + tag;
        if (did && deviceId && did === deviceId) {
            line += ' · máy này';
        }
        if (count > 1) {
            line += ' · ' + count + ' tab cùng máy';
        }
        return line;
    }

    function buildLabel() {
        const code = displayCode || makeDisplayCode(sessionId);
        if (publicIp) {
            return publicIp + ' · ' + code;
        }
        return 'tab · ' + code;
    }

    function fetchSelfGeo() {
        return fetch(GEO_SELF_URL, { cache: 'no-store' })
            .then((r) => (r.ok ? r.json() : null))
            .then((j) => {
                if (!j) {
                    return;
                }
                if (j.ip) {
                    publicIp = String(j.ip);
                }
                const g = parseGeoFromPayload(j);
                if (g) {
                    selfGeo = g;
                    if (publicIp) {
                        geoCache.set(publicIp, g);
                    }
                }
            })
            .catch(() => { /* offline / blocked */ });
    }

    function offlinePayload(offlineAt) {
        const now = Date.now();
        return attachGeoFields({
            online: false,
            label: buildLabel(),
            code: displayCode || makeDisplayCode(sessionId),
            startedAt: startedAt || now,
            offlineAt: Number(offlineAt) || now,
            updatedAt: now
        });
    }

    function writePresence(online) {
        if (!sessionRef) {
            return Promise.resolve();
        }
        const now = Date.now();
        if (!online) {
            return sessionRef.set(offlinePayload(now));
        }
        return sessionRef.set(attachGeoFields({
            online: true,
            label: buildLabel(),
            code: displayCode || makeDisplayCode(sessionId),
            startedAt: startedAt || now,
            updatedAt: now
        }));
    }

    function requestRemoveNode(id) {
        if (!db || !id || id === sessionId || pruneRequested.has(id)) {
            return;
        }
        pruneRequested.add(id);
        db.ref(PATH + '/' + id).remove()
            .catch(() => { /* ignore */ })
            .then(() => {
                setTimeout(() => pruneRequested.delete(id), 5000);
            });
    }

    /** Ghi tombstone offline lên RTDB (chính mình hoặc giúp peer bị xóa). */
    function requestTombstone(id, meta, offlineAt) {
        if (!db || !id || id === sessionId || tombstoneRequested.has(id)) {
            return;
        }
        const m = meta || {};
        const at = Number(offlineAt) || Date.now();
        tombstoneRequested.add(id);
        const payload = {
            online: false,
            label: String(m.label || id),
            code: String(m.code || ''),
            startedAt: Number(m.startedAt) || at,
            offlineAt: at,
            updatedAt: Date.now()
        };
        if (m.country) {
            payload.country = String(m.country);
        }
        if (m.countryCode) {
            payload.countryCode = String(m.countryCode);
        }
        if (m.region) {
            payload.region = String(m.region);
        }
        if (m.city) {
            payload.city = String(m.city);
        }
        if (Number.isFinite(Number(m.lat)) && Number.isFinite(Number(m.lon))) {
            payload.lat = Number(m.lat);
            payload.lon = Number(m.lon);
        }
        if (m.ip) {
            payload.ip = String(m.ip);
        }
        if (m.deviceId) {
            payload.deviceId = String(m.deviceId);
        }
        if (m.deviceTag) {
            payload.deviceTag = String(m.deviceTag);
        }
        db.ref(PATH + '/' + id).set(payload)
            .catch(() => { /* ignore */ })
            .then(() => {
                setTimeout(() => tombstoneRequested.delete(id), 8000);
            });
    }

    function metaLine(started, endedAt, row, isSelf) {
        const start = Number(started) || 0;
        const end = Number(endedAt) || Date.now();
        const access = formatAccessTime(start);
        const dur = start ? formatDuration(end - start) : '—';
        let line = 'Vào ' + access + ' · đã ' + dur;
        const geo = geoFromRow(row) || (isSelf ? selfGeo : null);
        if (geo && selfGeo && Number.isFinite(selfGeo.lat) && Number.isFinite(selfGeo.lon)
            && Number.isFinite(geo.lat) && Number.isFinite(geo.lon)) {
            const km = haversineKm(selfGeo.lat, selfGeo.lon, geo.lat, geo.lon);
            line += ' · ~' + formatDistanceKm(km) + ' từ You';
        }
        return line;
    }

    function geoMetaLine(row, isSelf) {
        const geo = geoFromRow(row) || (isSelf ? selfGeo : null);
        if (!geo) {
            const ip = extractIpFromLabel(row && row.label) || (row && row.ip) || '';
            if (ip) {
                ensureGeoForIp(ip);
            }
            return '';
        }
        const place = formatGeoPlace(geo);
        const coords = '(' + formatCoord(geo.lat) + ', ' + formatCoord(geo.lon) + ')';
        return place + ' · ' + coords;
    }

    function rowGeoFields(row) {
        return {
            country: String(row.country || '').trim(),
            countryCode: String(row.countryCode || '').trim(),
            region: String(row.region || '').trim(),
            city: String(row.city || '').trim(),
            lat: Number(row.lat),
            lon: Number(row.lon),
            ip: String(row.ip || '').trim(),
            label: String(row.label || '')
        };
    }

    function armOnDisconnect() {
        if (!sessionRef || typeof firebase === 'undefined') {
            return Promise.resolve();
        }
        const SERVER_TS = firebase.database.ServerValue.TIMESTAMP;
        return sessionRef.onDisconnect().set(attachGeoFields({
            online: false,
            label: buildLabel(),
            code: displayCode || makeDisplayCode(sessionId),
            startedAt: startedAt || Date.now(),
            offlineAt: SERVER_TS,
            updatedAt: SERVER_TS
        }));
    }

    function renderPresence(snapVal) {
        lastSnapVal = snapVal;
        const onlineRows = [];
        const offlineRows = [];
        const data = snapVal && typeof snapVal === 'object' ? snapVal : {};
        const now = Date.now();
        const nextOnlineIds = new Set();
        const nextMeta = {};

        Object.keys(data).forEach((id) => {
            const row = data[id] || {};
            const label = String(row.label || id);
            const code = String(row.code || '');
            const rowStarted = Number(row.startedAt) || Number(row.updatedAt) || 0;
            const updatedAt = Number(row.updatedAt) || 0;
            const geoFields = rowGeoFields(Object.assign({}, row, { label: label }));
            const rowDeviceId = String(row.deviceId || '').trim();
            const rowDeviceTag = String(row.deviceTag || '').trim()
                || (rowDeviceId ? makeDeviceTag(rowDeviceId) : '');
            nextMeta[id] = {
                label: label,
                code: code,
                startedAt: rowStarted,
                country: geoFields.country,
                countryCode: geoFields.countryCode,
                region: geoFields.region,
                city: geoFields.city,
                lat: geoFields.lat,
                lon: geoFields.lon,
                ip: geoFields.ip,
                deviceId: rowDeviceId,
                deviceTag: rowDeviceTag
            };

            const alive = !!row.online && updatedAt > 0 && (now - updatedAt) <= STALE_ONLINE_MS;
            if (alive) {
                nextOnlineIds.add(id);
                onlineRows.push(Object.assign({
                    id: id,
                    label: label,
                    startedAt: rowStarted,
                    updatedAt: updatedAt,
                    deviceId: rowDeviceId,
                    deviceTag: rowDeviceTag
                }, geoFields));
                return;
            }

            let offlineAt = Number(row.offlineAt) || 0;
            if (row.online && !alive) {
                offlineAt = offlineAt || updatedAt || now;
                requestTombstone(id, nextMeta[id], offlineAt);
            } else if (!row.online) {
                offlineAt = offlineAt || updatedAt || 0;
            }

            if (!offlineAt) {
                requestRemoveNode(id);
                return;
            }
            if (now - offlineAt > OFFLINE_HOLD_MS) {
                requestRemoveNode(id);
                return;
            }
            offlineRows.push(Object.assign({
                id: id,
                label: label,
                startedAt: rowStarted,
                at: offlineAt,
                deviceId: rowDeviceId,
                deviceTag: rowDeviceTag
            }, geoFields));
            pendingGone.delete(id);
        });

        prevOnlineIds.forEach((id) => {
            if (nextOnlineIds.has(id) || Object.prototype.hasOwnProperty.call(data, id)) {
                return;
            }
            if (id === sessionId) {
                return;
            }
            const meta = metaById[id] || nextMeta[id];
            if (!meta) {
                return;
            }
            const at = now;
            if (!pendingGone.has(id)) {
                pendingGone.set(id, {
                    label: meta.label || id,
                    startedAt: Number(meta.startedAt) || at,
                    at: at,
                    country: meta.country,
                    countryCode: meta.countryCode,
                    region: meta.region,
                    city: meta.city,
                    lat: meta.lat,
                    lon: meta.lon,
                    ip: meta.ip,
                    deviceId: meta.deviceId,
                    deviceTag: meta.deviceTag
                });
            }
            requestTombstone(id, meta, at);
        });

        pendingGone.forEach((entry, id) => {
            if (nextOnlineIds.has(id) || Object.prototype.hasOwnProperty.call(data, id)) {
                pendingGone.delete(id);
                return;
            }
            if (!entry || now - entry.at > OFFLINE_HOLD_MS) {
                pendingGone.delete(id);
                return;
            }
            offlineRows.push({
                id: id,
                label: entry.label,
                startedAt: entry.startedAt,
                at: entry.at,
                country: entry.country,
                countryCode: entry.countryCode,
                region: entry.region,
                city: entry.city,
                lat: entry.lat,
                lon: entry.lon,
                ip: entry.ip,
                deviceId: entry.deviceId,
                deviceTag: entry.deviceTag
            });
        });

        metaById = nextMeta;
        prevOnlineIds = nextOnlineIds;

        const deviceCounts = {};
        onlineRows.forEach((r) => {
            const did = String(r.deviceId || '');
            if (!did) {
                return;
            }
            deviceCounts[did] = (deviceCounts[did] || 0) + 1;
        });

        // Vào sớm nhất lên đầu; vào sau append xuống dưới
        onlineRows.sort((a, b) => {
            const sa = Number(a.startedAt) || Number(a.updatedAt) || 0;
            const sb = Number(b.startedAt) || Number(b.updatedAt) || 0;
            if (sa !== sb) {
                return sa - sb;
            }
            const da = String(a.deviceId || '');
            const db_ = String(b.deviceId || '');
            if (da !== db_) {
                return da.localeCompare(db_);
            }
            return a.id.localeCompare(b.id);
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
            const isSelf = r.id === sessionId;
            const sameDevice = !!(r.deviceId && deviceId && r.deviceId === deviceId);
            const you = isSelf ? ' <span class="presence-you">(You)</span>' : '';
            const geoLine = geoMetaLine(r, isSelf);
            const deviceLine = deviceMetaLine(r, deviceCounts);
            html += '<li class="presence-item presence-item--online' +
                (isSelf ? ' presence-item--self' : '') +
                (sameDevice ? ' presence-item--same-device' : '') + '"' +
                ' data-device-id="' + escapeHtml(r.deviceId || '') + '"' +
                ' data-device-tag="' + escapeHtml(r.deviceTag || '') + '"' +
                ' data-label="' + escapeHtml(r.label || '') + '"' +
                ' data-online="1">' +
                personIconHtml(false) +
                '<span class="presence-text">' +
                '<span class="presence-label">' + escapeHtml(r.label) +
                ' đang online' + you + '</span>' +
                (deviceLine
                    ? '<span class="presence-meta presence-meta--device">' + escapeHtml(deviceLine) + '</span>'
                    : '') +
                '<span class="presence-meta">' +
                escapeHtml(metaLine(r.startedAt, null, r, isSelf)) + '</span>' +
                (geoLine
                    ? '<span class="presence-meta presence-meta--geo">' + escapeHtml(geoLine) + '</span>'
                    : '') +
                (!r.deviceId
                    ? ''
                    : sameDevice
                        ? (isSelf
                            ? ''
                            : '<span class="presence-same-device-note">Cùng máy — không chat</span>')
                        : '<button type="button" class="presence-chat-btn" data-chat-open="1" title="Nhắn tin (theo thiết bị)">Nhắn</button>') +
                '</span></li>';
        });
        offlineRows.forEach((r) => {
            const isSelf = r.id === sessionId;
            const sameDevice = !!(r.deviceId && deviceId && r.deviceId === deviceId);
            const you = isSelf ? ' <span class="presence-you">(You)</span>' : '';
            const geoLine = geoMetaLine(r, isSelf);
            const deviceLine = deviceMetaLine(r, null);
            html += '<li class="presence-item presence-item--offline' +
                (isSelf ? ' presence-item--self' : '') +
                (sameDevice ? ' presence-item--same-device' : '') + '"' +
                ' data-device-id="' + escapeHtml(r.deviceId || '') + '"' +
                ' data-device-tag="' + escapeHtml(r.deviceTag || '') + '"' +
                ' data-label="' + escapeHtml(r.label || '') + '"' +
                ' data-online="0">' +
                personIconHtml(true) +
                '<span class="presence-text">' +
                '<span class="presence-label">' + escapeHtml(r.label) +
                ' vừa offline' + you + '</span>' +
                (deviceLine
                    ? '<span class="presence-meta presence-meta--device">' + escapeHtml(deviceLine) + '</span>'
                    : '') +
                '<span class="presence-meta">' +
                escapeHtml(metaLine(r.startedAt, r.at, r, isSelf)) + '</span>' +
                (geoLine
                    ? '<span class="presence-meta presence-meta--geo">' + escapeHtml(geoLine) + '</span>'
                    : '') +
                (!r.deviceId
                    ? ''
                    : sameDevice
                        ? (isSelf
                            ? ''
                            : '<span class="presence-same-device-note">Cùng máy — không chat</span>')
                        : '<button type="button" class="presence-chat-btn" data-chat-open="1" title="Nhắn tin (theo thiết bị)">Nhắn</button>') +
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
        const list = el('presenceList');
        if (!btn || !panel || btn.dataset.presenceBound === '1') {
            return;
        }
        btn.dataset.presenceBound = '1';
        btn.addEventListener('click', (e) => {
            e.stopPropagation();
            if (window.DeviceChat && typeof window.DeviceChat.isOpen === 'function' && window.DeviceChat.isOpen()) {
                window.DeviceChat.close();
            }
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
        if (list && list.dataset.chatDelegate !== '1') {
            list.dataset.chatDelegate = '1';
            list.addEventListener('click', (e) => {
                const openBtn = e.target && e.target.closest
                    ? e.target.closest('[data-chat-open]')
                    : null;
                if (!openBtn) {
                    return;
                }
                e.preventDefault();
                e.stopPropagation();
                const row = openBtn.closest('li.presence-item');
                if (!row) {
                    return;
                }
                const peerId = String(row.getAttribute('data-device-id') || '').trim();
                if (!peerId || !deviceId || peerId === deviceId) {
                    return;
                }
                if (window.DeviceChat && typeof window.DeviceChat.open === 'function') {
                    window.DeviceChat.open({
                        deviceId: peerId,
                        deviceTag: String(row.getAttribute('data-device-tag') || ''),
                        label: String(row.getAttribute('data-label') || ''),
                        online: row.getAttribute('data-online') === '1'
                    });
                }
            });
        }
    }

    function startHeartbeat() {
        if (heartbeatTimer) {
            clearInterval(heartbeatTimer);
        }
        heartbeatTimer = setInterval(() => {
            writePresence(true)
                .then(() => armOnDisconnect())
                .catch(() => { /* ignore */ });
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

    function markSelfOfflineBestEffort() {
        try {
            if (heartbeatTimer) {
                clearInterval(heartbeatTimer);
                heartbeatTimer = 0;
            }
            if (!sessionRef) {
                return;
            }
            writePresence(false);
        } catch (e) { /* ignore */ }
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
            deviceId = getOrCreateDeviceId();
            deviceTag = makeDeviceTag(deviceId);
            sessionId = makeSessionId();
            displayCode = makeDisplayCode(sessionId);
            startedAt = Date.now();
            sessionRef = db.ref(PATH + '/' + sessionId);

            const connectedRef = db.ref('.info/connected');
            connectedRef.on('value', (snap) => {
                if (snap.val() !== true) {
                    return;
                }
                armOnDisconnect()
                    .then(() => writePresence(true))
                    .then(() => startHeartbeat())
                    .catch(() => { /* ignore */ });
            });

            db.ref(PATH).on('value', (snap) => {
                renderPresence(snap.val());
            });
            startRenderTick();

            fetchSelfGeo().then(() => {
                writePresence(true).then(() => armOnDisconnect());
            }).catch(() => { /* ignore */ });

            window.addEventListener('pagehide', markSelfOfflineBestEffort);
            window.addEventListener('beforeunload', markSelfOfflineBestEffort);

            window.PresenceBridge = {
                getDeviceId: () => deviceId,
                getDeviceTag: () => deviceTag,
                getSessionId: () => sessionId
            };
        };

        if (firebase.auth) {
            firebase.auth().signInAnonymously()
                .then(afterAuth)
                .catch((err) => {
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
