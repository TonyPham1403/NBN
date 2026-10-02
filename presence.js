/**
 * Firebase Realtime Database presence — đếm tab đang online + list.
 * Config: window.PRESENCE_FIREBASE_CONFIG (presence-config.js). Hướng dẫn: PRESENCE.md
 *
 * Offline ~5 phút ghi trên Firebase (online:false + offlineAt). Tab còn mở có thể
 * tái tạo tombstone nếu peer bị xóa (code cũ / cancel onDisconnect) để người vào sau vẫn thấy.
 * GeoIP (ipwho.is → geojs): quốc gia/tỉnh theo IP công khai (fallback).
 * Ưu tiên tọa độ trình duyệt (GPS/Wi‑Fi) khi user cho phép — chi tiết hơn tâm thành phố.
 * Khoảng cách Haversine tới (You).
 * Device ID bền (localStorage). Mã máy 3 chữ (AYU); tab cùng máy AYU1, AYU2…
 * Đếm online = số device unique. Máy khác chỉ hiện tên đại diện AYU (không liệt kê từng tab).
 * Mã máy đổi sau ~10 phút không còn tab (cùng cơ chế idle với chat local).
 */
(function () {
    'use strict';

    const HEARTBEAT_MS = 20000;
    const STALE_ONLINE_MS = 90 * 1000;
    const OFFLINE_HOLD_MS = 5 * 60 * 1000;
    const RENDER_TICK_MS = 15000;
    const IDLE_WIPE_MS = 10 * 60 * 1000;
    const TAB_SLOT_STALE_MS = 90 * 1000;
    const PATH = 'presence';
    const DEVICE_STORAGE_KEY = 'presenceDeviceId';
    const DEVICE_CODE_KEY = 'presenceDeviceCode';
    const TAB_SLOTS_KEY = 'presenceTabSlots';
    const ALIVE_KEY = 'deviceChatAliveAt';
    const AWAY_KEY = 'deviceChatAwayAt';
    const GEO_PROVIDERS = {
        ipwhoSelf: 'https://ipwho.is/',
        ipwhoIp: 'https://ipwho.is/',
        geojsSelf: 'https://get.geojs.io/v1/ip/geo.json',
        geojsIp: 'https://get.geojs.io/v1/ip/geo/'
    };
    /** Chuẩn hóa tên tỉnh/thành VN thường gặp từ GeoIP (tiếng Anh → tiếng Việt). */
    const VN_PLACE_MAP = {
        'ha noi': 'Hà Nội',
        'hanoi': 'Hà Nội',
        'ho chi minh': 'TP. Hồ Chí Minh',
        'ho chi minh city': 'TP. Hồ Chí Minh',
        'hcm': 'TP. Hồ Chí Minh',
        'saigon': 'TP. Hồ Chí Minh',
        'da nang': 'Đà Nẵng',
        'danang': 'Đà Nẵng',
        'hai phong': 'Hải Phòng',
        'haiphong': 'Hải Phòng',
        'can tho': 'Cần Thơ',
        'cantho': 'Cần Thơ',
        'dong nai': 'Đồng Nai',
        'binh duong': 'Bình Dương',
        'ba ria - vung tau': 'Bà Rịa - Vũng Tàu',
        'ba ria vung tau': 'Bà Rịa - Vũng Tàu',
        'vung tau': 'Bà Rịa - Vũng Tàu',
        'khanh hoa': 'Khánh Hòa',
        'lam dong': 'Lâm Đồng',
        'thua thien hue': 'Thừa Thiên Huế',
        'hue': 'Thừa Thiên Huế',
        'quang nam': 'Quảng Nam',
        'quang ninh': 'Quảng Ninh',
        'nghe an': 'Nghệ An',
        'thanh hoa': 'Thanh Hóa',
        'binh dinh': 'Bình Định',
        'gia lai': 'Gia Lai',
        'dak lak': 'Đắk Lắk',
        'daklak': 'Đắk Lắk',
        'an giang': 'An Giang',
        'kien giang': 'Kiên Giang',
        'long an': 'Long An',
        'tay ninh': 'Tây Ninh',
        'tien giang': 'Tiền Giang',
        'ben tre': 'Bến Tre',
        'vinh long': 'Vĩnh Long',
        'dong thap': 'Đồng Tháp',
        'soc trang': 'Sóc Trăng',
        'bac ninh': 'Bắc Ninh',
        'bac giang': 'Bắc Giang',
        'hung yen': 'Hưng Yên',
        'hai duong': 'Hải Dương',
        'nam dinh': 'Nam Định',
        'ninh binh': 'Ninh Bình',
        'phu tho': 'Phú Thọ',
        'thai nguyen': 'Thái Nguyên',
        'vinh phuc': 'Vĩnh Phúc'
    };

    let sessionId = '';
    /** Hiển thị tab: AYU1, AYU2… */
    let displayCode = '';
    /** UUID bền theo trình duyệt/máy. */
    let deviceId = '';
    /** Mã máy 3 chữ cái: AYU (không có số). */
    let deviceTag = '';
    /** Số thứ tự tab trên máy này (1, 2, 11…). */
    let tabIndex = 1;
    let startedAt = 0;
    let aliveTimer = 0;
    let sessionRef = null;
    let heartbeatTimer = 0;
    let renderTickTimer = 0;
    let publicIp = '';
    /** Id kỳ sheet đang focus trên tab này (đồng bộ qua presence). */
    let localFocusRowId = '';
    let focusPublishTimer = 0;
    let lastPublishedFocusRowId = '';
    /** @type {{ country: string, countryCode: string, region: string, city: string, lat: number, lon: number, precise?: boolean, geoSource?: string, accuracyM?: number }|null} */
    let selfGeo = null;
    let lastSnapVal = {};
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
    /** Ước lượng loại máy / HĐH / trình duyệt (từ UA + Client Hints). */
    let selfClientDevice = null;

    function el(id) {
        return document.getElementById(id);
    }

    function detectClientDeviceSync() {
        const ua = String((typeof navigator !== 'undefined' && navigator.userAgent) || '');
        const uaData = (typeof navigator !== 'undefined' && navigator.userAgentData) || null;
        let form = 'desktop';
        let os = '';
        let browser = '';
        let group = '';
        let model = '';

        const maxTouch = Number((typeof navigator !== 'undefined' && navigator.maxTouchPoints) || 0);
        const platform = String((typeof navigator !== 'undefined' && navigator.platform) || '');
        const isIpad = /iPad/i.test(ua) || (platform === 'MacIntel' && maxTouch > 1);
        const isIphone = /iPhone|iPod/i.test(ua);
        const isAndroid = /Android/i.test(ua);

        if (isIpad) {
            form = 'tablet';
            os = 'iPadOS';
            group = 'iPad';
        } else if (isIphone) {
            form = 'mobile';
            os = 'iOS';
            group = 'iPhone';
        } else if (isAndroid) {
            form = /Mobile/i.test(ua) ? 'mobile' : 'tablet';
            os = 'Android';
            group = 'Android';
            const am = ua.match(/\((?:Linux; )?Android [^;]+;\s*([^;)]+)/i);
            if (am && am[1]) {
                const name = String(am[1]).replace(/\s*Build.*$/i, '').trim();
                if (name && !/^Android$/i.test(name) && name !== 'Linux' && !/^U$/i.test(name)) {
                    model = name;
                }
            }
        } else if (/Windows Phone|IEMobile/i.test(ua)) {
            form = 'mobile';
            os = 'Windows';
        } else if (/Mobile|Mobi/i.test(ua)) {
            form = 'mobile';
        }

        if (!os) {
            if (/Windows NT/i.test(ua)) {
                os = 'Windows';
            } else if (/Mac OS X|Macintosh/i.test(ua)) {
                os = 'macOS';
            } else if (/CrOS/i.test(ua)) {
                os = 'Chrome OS';
            } else if (/Linux/i.test(ua)) {
                os = 'Linux';
            }
        }

        if (uaData) {
            if (uaData.mobile && form === 'desktop') {
                form = 'mobile';
            }
            const p = String(uaData.platform || '');
            if (/Win/i.test(p)) {
                os = 'Windows';
            } else if (/macOS|Mac OS/i.test(p)) {
                os = os === 'iPadOS' ? os : 'macOS';
            } else if (/Android/i.test(p)) {
                os = 'Android';
                group = group || 'Android';
            } else if (/iOS/i.test(p)) {
                os = os || 'iOS';
            } else if (/Linux/i.test(p) && !os) {
                os = 'Linux';
            } else if (/Chrome OS/i.test(p)) {
                os = 'Chrome OS';
            }
        }

        if (/Edg\/|EdgiOS\//i.test(ua)) {
            browser = 'Edge';
        } else if (/OPR\/|Opera/i.test(ua)) {
            browser = 'Opera';
        } else if (/SamsungBrowser\//i.test(ua)) {
            browser = 'Samsung Internet';
        } else if (/FxiOS\/|Firefox\//i.test(ua)) {
            browser = 'Firefox';
        } else if (/CriOS\//i.test(ua)) {
            browser = 'Chrome';
        } else if (/Chrome\//i.test(ua) && !/Edg\//i.test(ua)) {
            browser = 'Chrome';
        } else if (/Safari\//i.test(ua) && !/Chrome\//i.test(ua) && !/Chromium\//i.test(ua)) {
            browser = 'Safari';
        }

        return { form: form, os: os, browser: browser, group: group, model: model };
    }

    function applyClientHints(h) {
        if (!h || !selfClientDevice) {
            return false;
        }
        let changed = false;
        if (h.mobile && selfClientDevice.form === 'desktop') {
            selfClientDevice.form = 'mobile';
            changed = true;
        }
        const p = String(h.platform || '');
        if (p) {
            let nextOs = selfClientDevice.os;
            if (/Win/i.test(p)) {
                nextOs = 'Windows';
            } else if (/macOS|Mac OS/i.test(p)) {
                nextOs = selfClientDevice.os === 'iPadOS' ? selfClientDevice.os : 'macOS';
            } else if (/Android/i.test(p)) {
                nextOs = 'Android';
                if (!selfClientDevice.group) {
                    selfClientDevice.group = 'Android';
                    changed = true;
                }
            } else if (/iOS/i.test(p)) {
                nextOs = selfClientDevice.os || 'iOS';
            } else if (/Linux/i.test(p)) {
                nextOs = selfClientDevice.os || 'Linux';
            } else if (/Chrome OS/i.test(p)) {
                nextOs = 'Chrome OS';
            }
            if (nextOs && nextOs !== selfClientDevice.os) {
                selfClientDevice.os = nextOs;
                changed = true;
            }
        }
        const model = String(h.model || '').trim();
        if (model && model !== selfClientDevice.model) {
            selfClientDevice.model = model;
            changed = true;
            if (/iPhone/i.test(model)) {
                selfClientDevice.group = 'iPhone';
            } else if (/iPad/i.test(model)) {
                selfClientDevice.group = 'iPad';
            }
        }
        return changed;
    }

    function enrichClientDeviceHints() {
        if (!navigator.userAgentData || typeof navigator.userAgentData.getHighEntropyValues !== 'function') {
            return Promise.resolve(false);
        }
        return navigator.userAgentData
            .getHighEntropyValues(['platform', 'platformVersion', 'model', 'mobile'])
            .then((h) => applyClientHints(h))
            .catch(() => false);
    }

    function pickClientDevice(row) {
        const src = row || {};
        return {
            clientForm: String(src.clientForm || '').trim(),
            clientOs: String(src.clientOs || '').trim(),
            clientBrowser: String(src.clientBrowser || '').trim(),
            clientGroup: String(src.clientGroup || '').trim(),
            clientModel: String(src.clientModel || '').trim()
        };
    }

    function clientDevicePayloadFields(info) {
        const d = info || selfClientDevice;
        if (!d) {
            return {};
        }
        const out = {};
        if (d.form || d.clientForm) {
            out.clientForm = String(d.form || d.clientForm || '').trim();
        }
        if (d.os || d.clientOs) {
            out.clientOs = String(d.os || d.clientOs || '').trim();
        }
        if (d.browser || d.clientBrowser) {
            out.clientBrowser = String(d.browser || d.clientBrowser || '').trim();
        }
        if (d.group || d.clientGroup) {
            out.clientGroup = String(d.group || d.clientGroup || '').trim();
        }
        if (d.model || d.clientModel) {
            out.clientModel = String(d.model || d.clientModel || '').trim();
        }
        return out;
    }

    function formatClientDeviceTip(row) {
        const d = pickClientDevice(row);
        const formLabel = d.clientForm === 'mobile'
            ? 'Điện thoại'
            : (d.clientForm === 'tablet'
                ? 'Máy tính bảng'
                : (d.clientForm === 'desktop' ? 'Máy tính' : ''));
        const lines = [];
        if (formLabel) {
            lines.push('Loại: ' + formLabel);
        }
        if (d.clientOs) {
            lines.push('HĐH: ' + d.clientOs);
        }
        if (d.clientBrowser) {
            lines.push('Trình duyệt: ' + d.clientBrowser);
        }
        if (d.clientGroup) {
            lines.push('Nhóm: ' + d.clientGroup);
        }
        if (d.clientModel) {
            lines.push('Model: ' + d.clientModel);
        }
        if (!lines.length) {
            return '';
        }
        return lines.join('\n');
    }

    function deviceMetaHtml(row, deviceCounts, isSelf) {
        let line = deviceMetaLine(row, deviceCounts);
        if (!line && isSelf) {
            line = 'Thiết bị';
        }
        if (!line) {
            return '';
        }
        const tip = formatClientDeviceTip(row);
        if (!tip) {
            return '<span class="presence-device-hint">' + escapeHtml(line) + '</span>';
        }
        /* &#10; tránh xuống dòng thô trong attribute (một số môi trường làm lộ text ra DOM) */
        const tipAttr = escapeHtml(tip).replace(/\r?\n/g, '&#10;');
        return '<span class="presence-device-hint" data-tip="' + tipAttr +
            '" tabindex="0">' + escapeHtml(line) + '</span>';
    }

    function makeSessionId() {
        try {
            if (crypto && typeof crypto.randomUUID === 'function') {
                return crypto.randomUUID();
            }
        } catch (e) { /* ignore */ }
        return 's' + Date.now().toString(36) + Math.random().toString(36).slice(2, 10);
    }

    /** 3 chữ cái A–Z (không số) — tên thiết bị đại diện. */
    function makeRandomDeviceCode() {
        let out = '';
        for (let i = 0; i < 3; i++) {
            let n = 0;
            try {
                if (crypto && crypto.getRandomValues) {
                    const buf = new Uint8Array(1);
                    crypto.getRandomValues(buf);
                    n = buf[0];
                } else {
                    n = Math.floor(Math.random() * 256);
                }
            } catch (e) {
                n = Math.floor(Math.random() * 256);
            }
            out += String.fromCharCode(65 + (n % 26));
        }
        return out;
    }

    function isValidDeviceCode(code) {
        return /^[A-Z]{3}$/.test(String(code || ''));
    }

    function normalizeDeviceCode(code) {
        const s = String(code || '').toUpperCase().replace(/[^A-Z]/g, '');
        if (s.length >= 3) {
            return s.slice(0, 3);
        }
        return '';
    }

    /** Fallback cũ: hash → 3 chữ (không số). */
    function makeDeviceTag(id) {
        const s = String(id || '').replace(/\W/g, '') || 'x';
        let h = 2166136261;
        for (let i = 0; i < s.length; i++) {
            h ^= s.charCodeAt(i);
            h = Math.imul(h, 16777619);
        }
        h = h >>> 0;
        return String.fromCharCode(65 + (h % 26))
            + String.fromCharCode(65 + ((h >>> 5) % 26))
            + String.fromCharCode(65 + ((h >>> 10) % 26));
    }

    function touchAlive() {
        try {
            localStorage.setItem(ALIVE_KEY, String(Date.now()));
            localStorage.removeItem(AWAY_KEY);
        } catch (e) { /* ignore */ }
    }

    function markAway() {
        try {
            /* Chỉ ghi lần đầu ẩn/đóng — không refresh mỗi tick (nếu refresh thì idle wipe không bao giờ chạy). */
            if (!localStorage.getItem(AWAY_KEY)) {
                localStorage.setItem(AWAY_KEY, String(Date.now()));
            }
        } catch (e) { /* ignore */ }
    }

    function wasIdleTooLong() {
        const now = Date.now();
        let aliveAt = 0;
        let awayAt = 0;
        try {
            aliveAt = Number(localStorage.getItem(ALIVE_KEY)) || 0;
            awayAt = Number(localStorage.getItem(AWAY_KEY)) || 0;
        } catch (e) { /* ignore */ }
        const last = Math.max(aliveAt, awayAt);
        if (last <= 0) {
            return false;
        }
        return (now - last) > IDLE_WIPE_MS;
    }

    function onPresenceIdleTick() {
        if (document.hidden) {
            markAway();
            return;
        }
        if (!wasIdleTooLong()) {
            touchAlive();
        }
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

    function getOrCreateDeviceCode() {
        if (wasIdleTooLong()) {
            try {
                localStorage.removeItem(DEVICE_CODE_KEY);
                localStorage.removeItem(TAB_SLOTS_KEY);
            } catch (e) { /* ignore */ }
        }
        try {
            let code = localStorage.getItem(DEVICE_CODE_KEY);
            if (isValidDeviceCode(code)) {
                return String(code).toUpperCase();
            }
            code = makeRandomDeviceCode();
            localStorage.setItem(DEVICE_CODE_KEY, code);
            return code;
        } catch (e) {
            return makeRandomDeviceCode();
        }
    }

    /** deviceId → mã 3 chữ trên /presence. Nhiều tab cùng máy = một deviceId → cùng mã (AYU1, AYU2). */
    function extractDeviceCodesFromPresence(data) {
        const map = new Map();
        if (!data || typeof data !== 'object') {
            return map;
        }
        Object.keys(data).forEach((sid) => {
            const row = data[sid];
            if (!row || typeof row !== 'object') {
                return;
            }
            const did = String(row.deviceId || '').trim();
            const code = normalizeDeviceCode(row.deviceCode || row.deviceTag);
            if (!did || !code) {
                return;
            }
            map.set(did, code);
        });
        return map;
    }

    /** Mã đang được thiết bị **khác** (khác deviceId) dùng — bỏ qua tab/máy mình. */
    function getUsedCodesByOthers(deviceCodeMap) {
        const used = new Set();
        deviceCodeMap.forEach((code, did) => {
            if (did && deviceId && did !== deviceId) {
                used.add(code);
            }
        });
        return used;
    }

    /** Thiết bị online khác đang dùng cùng mã (xử lý race khi 2 máy mới cùng lúc). */
    function onlineDeviceIdsWithCode(data, code) {
        const holders = new Set();
        const want = normalizeDeviceCode(code);
        if (!want || !data || typeof data !== 'object') {
            return holders;
        }
        const now = Date.now();
        Object.keys(data).forEach((sid) => {
            const row = data[sid];
            if (!row || typeof row !== 'object') {
                return;
            }
            const rowCode = normalizeDeviceCode(row.deviceCode || row.deviceTag);
            if (rowCode !== want) {
                return;
            }
            const updatedAt = Number(row.updatedAt) || 0;
            const alive = !!row.online && updatedAt > 0 && (now - updatedAt) <= STALE_ONLINE_MS;
            if (!alive) {
                return;
            }
            const did = String(row.deviceId || '').trim();
            if (did) {
                holders.add(did);
            }
        });
        return holders;
    }

    function codeAtIndex(i) {
        const n = Math.max(0, Math.min(17575, Math.floor(Number(i) || 0)));
        const a = 65 + Math.floor(n / 676) % 26;
        const b = 65 + Math.floor(n / 26) % 26;
        const c = 65 + n % 26;
        return String.fromCharCode(a, b, c);
    }

    /** Chọn mã chưa có thiết bị khác trên Firebase. */
    function pickAvailableDeviceCode(usedSet) {
        const used = usedSet || new Set();
        for (let i = 0; i < 400; i++) {
            const code = makeRandomDeviceCode();
            if (!used.has(code)) {
                return code;
            }
        }
        for (let i = 0; i < 17576; i++) {
            const code = codeAtIndex(i);
            if (!used.has(code)) {
                return code;
            }
        }
        return makeRandomDeviceCode();
    }

    function setDeviceCode(code) {
        const next = normalizeDeviceCode(code) || makeRandomDeviceCode();
        deviceTag = next;
        displayCode = deviceTag + String(tabIndex || 1);
        try {
            localStorage.setItem(DEVICE_CODE_KEY, deviceTag);
        } catch (e) { /* ignore */ }
    }

    /**
     * Mã 3 chữ unique theo deviceId (máy), không theo tab.
     * Cùng máy: AYU1 + AYU2 OK. Hai máy khác: không được cùng AYU.
     */
    function reconcileUniqueDeviceCode(data) {
        if (!deviceId) {
            return false;
        }
        const deviceMap = extractDeviceCodesFromPresence(data);
        const usedByOthers = getUsedCodesByOthers(deviceMap);
        let myCode = normalizeDeviceCode(deviceTag) || '';

        if (!isValidDeviceCode(myCode)) {
            setDeviceCode(pickAvailableDeviceCode(usedByOthers));
            return true;
        }

        const onlineHolders = onlineDeviceIdsWithCode(data, myCode);
        if (onlineHolders.size > 1) {
            const winner = Array.from(onlineHolders).sort()[0];
            if (winner !== deviceId) {
                usedByOthers.add(myCode);
                setDeviceCode(pickAvailableDeviceCode(usedByOthers));
                return true;
            }
        }

        if (usedByOthers.has(myCode)) {
            setDeviceCode(pickAvailableDeviceCode(usedByOthers));
            return true;
        }

        return false;
    }

    function readTabSlots() {
        try {
            const j = JSON.parse(localStorage.getItem(TAB_SLOTS_KEY) || '{}');
            return j && typeof j === 'object' ? j : {};
        } catch (e) {
            return {};
        }
    }

    function writeTabSlots(slots) {
        try {
            localStorage.setItem(TAB_SLOTS_KEY, JSON.stringify(slots || {}));
        } catch (e) { /* ignore */ }
    }

    function claimTabIndex(sid) {
        const now = Date.now();
        const slots = readTabSlots();
        Object.keys(slots).forEach((k) => {
            const row = slots[k];
            if (!row || !row.sid || (now - (Number(row.at) || 0)) > TAB_SLOT_STALE_MS) {
                delete slots[k];
            }
        });
        const keys = Object.keys(slots);
        for (let i = 0; i < keys.length; i++) {
            const k = keys[i];
            if (slots[k] && slots[k].sid === sid) {
                slots[k].at = now;
                writeTabSlots(slots);
                return Math.max(1, parseInt(k, 10) || 1);
            }
        }
        let n = 1;
        while (slots[String(n)]) {
            n += 1;
        }
        slots[String(n)] = { sid: sid, at: now };
        writeTabSlots(slots);
        return n;
    }

    function refreshTabSlot(sid) {
        if (!sid) {
            return;
        }
        const now = Date.now();
        const slots = readTabSlots();
        const key = String(tabIndex);
        if (slots[key] && slots[key].sid === sid) {
            slots[key].at = now;
            writeTabSlots(slots);
            return;
        }
        tabIndex = claimTabIndex(sid);
        displayCode = deviceTag + String(tabIndex);
    }

    function releaseTabSlot(sid) {
        const slots = readTabSlots();
        Object.keys(slots).forEach((k) => {
            if (slots[k] && slots[k].sid === sid) {
                delete slots[k];
            }
        });
        writeTabSlots(slots);
    }

    function formatUserLabel(ip, userCode) {
        const code = String(userCode || '').trim() || '???';
        if (ip) {
            return String(ip) + ' · ' + code;
        }
        return 'tab · ' + code;
    }

    function extractIpFromRow(row) {
        return rowIp(row);
    }

    function pickBestGeoRow(rows) {
        if (!rows || !rows.length) {
            return null;
        }
        const scored = rows.slice().sort((a, b) => {
            const score = (r) => {
                let s = 0;
                if (rowIp(r)) {
                    s += 4;
                }
                if (Number.isFinite(Number(r.lat)) && Number.isFinite(Number(r.lon))) {
                    s += 2;
                }
                if (String(r.region || r.city || '').trim()) {
                    s += 1;
                }
                s += (Number(r.updatedAt) || Number(r.at) || 0) / 1e15;
                return s;
            };
            return score(b) - score(a);
        });
        return scored[0];
    }

    function resolvePeerDeviceCode(row) {
        const tagged = normalizeDeviceCode(row && (row.deviceCode || row.deviceTag));
        if (tagged) {
            return tagged;
        }
        const code = String((row && row.code) || '');
        const m = code.match(/^([A-Za-z]{3})\d+$/);
        if (m) {
            return m[1].toUpperCase();
        }
        if (row && row.deviceId) {
            return makeDeviceTag(row.deviceId);
        }
        return normalizeDeviceCode(code) || 'XXX';
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

    /** Avatar chữ cái (mã thiết bị); fallback nếu chat.js chưa sẵn sàng. */
    function personAvatarHtml(deviceId, deviceTag, offline) {
        const cls = 'device-avatar device-avatar--sm' + (offline ? ' device-avatar--offline' : '');
        if (window.DeviceChat && typeof window.DeviceChat.avatarHtml === 'function') {
            return window.DeviceChat.avatarHtml(deviceId, deviceTag, { className: cls });
        }
        const raw = String(deviceTag || '').toUpperCase().replace(/[^A-Z0-9]/g, '');
        const letter = raw.charAt(0) || '?';
        return '<span class="' + cls + '" style="background:#2196F3" aria-hidden="true">' +
            escapeHtml(letter) + '</span>';
    }

    /** Avatar tròn + badge unread; bấm để chat (peer khác máy). */
    function personIconHtml(offline, opts) {
        const o = opts || {};
        const canChat = !!o.canChat;
        const unread = Math.max(0, Math.floor(Number(o.unread) || 0));
        const cls = offline ? 'presence-person presence-person--offline' : 'presence-person';
        const badge = unread > 0
            ? '<span class="presence-person-badge">' + (unread > 99 ? '99+' : String(unread)) + '</span>'
            : '';
        const icon = '<span class="' + cls + '" aria-hidden="true">' +
            personAvatarHtml(o.deviceId, o.deviceTag, offline) + badge + '</span>';
        if (!canChat) {
            return icon;
        }
        const title = unread > 0
            ? ('Nhắn tin — ' + unread + ' tin chưa đọc')
            : 'Nhắn tin với thiết bị này';
        return '<button type="button" class="presence-person-btn" data-chat-open="1" title="' +
            escapeHtml(title) + '" aria-label="' + escapeHtml(title) + '">' + icon + '</button>';
    }

    function getUnreadForDevice(peerDeviceId) {
        if (!peerDeviceId || !window.DeviceChat || typeof window.DeviceChat.getUnread !== 'function') {
            return 0;
        }
        return Number(window.DeviceChat.getUnread(peerDeviceId)) || 0;
    }

    /** active = phản hồi gần nhất (xanh); closed = đã phản hồi trong session (vàng); '' = chưa phản hồi */
    function chatBorderClass(peerDeviceId) {
        if (!peerDeviceId || !window.DeviceChat || typeof window.DeviceChat.getBorderState !== 'function') {
            return '';
        }
        const state = window.DeviceChat.getBorderState(peerDeviceId);
        if (state === 'active') {
            return ' presence-item--chat-active';
        }
        if (state === 'closed') {
            return ' presence-item--chat-closed';
        }
        return '';
    }

    function chatBorderAttr(peerDeviceId) {
        if (!peerDeviceId || !window.DeviceChat || typeof window.DeviceChat.getBorderState !== 'function') {
            return '';
        }
        const state = window.DeviceChat.getBorderState(peerDeviceId);
        return state ? (' data-chat-border="' + state + '"') : '';
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

    function normalizeVnPlace(name) {
        const raw = String(name || '').trim();
        if (!raw) {
            return '';
        }
        const key = raw.toLowerCase()
            .normalize('NFD')
            .replace(/[\u0300-\u036f]/g, '')
            .replace(/đ/g, 'd')
            .replace(/\s+/g, ' ')
            .trim();
        return VN_PLACE_MAP[key] || raw;
    }

    function parseGeoFromPayload(j) {
        if (!j || typeof j !== 'object') {
            return null;
        }
        if (j.success === false) {
            return null;
        }
        const lat = Number(j.latitude != null ? j.latitude : j.lat);
        const lon = Number(j.longitude != null ? j.longitude : j.lon);
        if (!Number.isFinite(lat) || !Number.isFinite(lon)) {
            return null;
        }
        const countryCode = String(j.country_code || j.countryCode || '').trim().toUpperCase();
        const country = String(j.country || '').trim() || countryCode;
        let region = String(j.region || j.regionName || j.region_name || '').trim();
        let city = String(j.city || '').trim();
        if (countryCode === 'VN' || /viet\s*nam/i.test(country)) {
            region = normalizeVnPlace(region);
            city = normalizeVnPlace(city);
        }
        return {
            country: country,
            countryCode: countryCode,
            region: region,
            city: city,
            lat: lat,
            lon: lon,
            ip: String(j.ip || '').trim()
        };
    }

    /** Quốc gia · tỉnh/thành (ưu tiên tỉnh/region; bỏ trùng city). */
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

    function formatCoord(n, digits) {
        const v = Number(n);
        if (!Number.isFinite(v)) {
            return '—';
        }
        const d = Number.isFinite(digits) ? digits : 4;
        return v.toFixed(d);
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

    function rowIp(row) {
        if (!row || typeof row !== 'object') {
            return '';
        }
        return String(row.ip || '').trim() || extractIpFromLabel(row.label);
    }

    /** Ưu tiên lat/lon peer đã publish (GPS); GeoIP cache chỉ bổ sung nhãn / fallback. */
    function geoFromRow(row) {
        if (!row || typeof row !== 'object') {
            return null;
        }
        const ip = rowIp(row);
        const cached = (ip && geoCache.has(ip)) ? geoCache.get(ip) : null;
        const lat = Number(row.lat);
        const lon = Number(row.lon);
        const precise = !!(row.geoPrecise || row.geoSource === 'gps'
            || (row.precise === true));
        if (Number.isFinite(lat) && Number.isFinite(lon)) {
            return {
                country: String(row.country || '').trim() || (cached && cached.country) || '',
                countryCode: String(row.countryCode || '').trim()
                    || (cached && cached.countryCode) || '',
                region: normalizeVnPlace(row.region)
                    || (cached && cached.region) || '',
                city: normalizeVnPlace(row.city)
                    || (cached && cached.city) || '',
                lat: lat,
                lon: lon,
                precise: precise,
                geoSource: precise ? 'gps' : String(row.geoSource || 'ip'),
                accuracyM: Number(row.geoAccuracyM) || undefined
            };
        }
        if (cached) {
            return Object.assign({}, cached, {
                precise: false,
                geoSource: 'ip'
            });
        }
        return null;
    }

    function fetchJson(url) {
        return fetch(url, { cache: 'no-store' })
            .then((r) => (r.ok ? r.json() : null))
            .catch(() => null);
    }

    /** ipwho.is (tỉnh/thành tốt hơn) → geojs fallback. */
    function lookupGeoByIp(ip) {
        const key = String(ip || '').trim();
        if (!key) {
            return Promise.resolve(null);
        }
        return fetchJson(GEO_PROVIDERS.ipwhoIp + encodeURIComponent(key))
            .then((j) => parseGeoFromPayload(j))
            .then((g) => {
                if (g) {
                    return g;
                }
                return fetchJson(GEO_PROVIDERS.geojsIp + encodeURIComponent(key) + '.json')
                    .then((j2) => parseGeoFromPayload(j2));
            });
    }

    function ensureGeoForIp(ip) {
        const key = String(ip || '').trim();
        if (!key || geoCache.has(key) || geoFetchInflight.has(key)) {
            return;
        }
        geoFetchInflight.add(key);
        lookupGeoByIp(key)
            .then((g) => {
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
            out.deviceTag = deviceTag;
            out.deviceCode = deviceTag;
            out.tabIndex = tabIndex;
        }
        const focusId = getLocalFocusRowId();
        if (focusId) {
            out.focusRowId = focusId;
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
            if (selfGeo.precise) {
                out.geoPrecise = true;
                out.geoSource = 'gps';
                if (Number.isFinite(Number(selfGeo.accuracyM))) {
                    out.geoAccuracyM = Math.round(Number(selfGeo.accuracyM));
                }
            } else {
                out.geoPrecise = false;
                out.geoSource = 'ip';
            }
        }
        Object.assign(out, clientDevicePayloadFields(selfClientDevice));
        return out;
    }

    function normalizeFocusRowId(raw) {
        const s = String(raw == null ? '' : raw).trim();
        if (!s) {
            return '';
        }
        // Chỉ giữ id số / chuỗi ngắn (tránh payload lạ)
        if (!/^[0-9A-Za-z._-]{1,24}$/.test(s)) {
            return '';
        }
        return s;
    }

    function getLocalFocusRowId() {
        if (typeof window.getSheetFocusRowId === 'function') {
            const live = normalizeFocusRowId(window.getSheetFocusRowId());
            if (live) {
                localFocusRowId = live;
                return live;
            }
        }
        return normalizeFocusRowId(localFocusRowId);
    }

    function publishFocusRowIdSoon() {
        const next = getLocalFocusRowId();
        if (next === lastPublishedFocusRowId && sessionRef) {
            return;
        }
        localFocusRowId = next;
        if (!sessionRef) {
            return;
        }
        if (focusPublishTimer) {
            clearTimeout(focusPublishTimer);
        }
        focusPublishTimer = setTimeout(() => {
            focusPublishTimer = 0;
            lastPublishedFocusRowId = getLocalFocusRowId();
            writePresence(true).catch(() => { /* ignore */ });
        }, 280);
    }

    function focusSheetRowById(rawId) {
        const id = normalizeFocusRowId(rawId);
        if (!id) {
            return false;
        }
        if (typeof window.focusSheetByRowId === 'function') {
            try {
                return !!window.focusSheetByRowId(id);
            } catch (eFocus) {
                return false;
            }
        }
        return false;
    }

    function formatFocusRowIdLabel(raw) {
        const id = normalizeFocusRowId(raw);
        return id || '—';
    }

    /** Cột avatar + id kỳ đang focus (bấm để nhảy tới). */
    function personColHtml(offline, opts, focusRowId) {
        const id = normalizeFocusRowId(focusRowId);
        const label = formatFocusRowIdLabel(id);
        const emptyCls = id ? '' : ' is-empty';
        const title = id
            ? ('Focus kỳ id ' + id)
            : 'Chưa có id focus';
        return '<span class="presence-avatar-col">' +
            personIconHtml(offline, opts) +
            '<button type="button" class="presence-focus-id' + emptyCls + '"' +
            (id ? (' data-focus-row-id="' + escapeHtml(id) + '"') : ' disabled') +
            ' title="' + escapeHtml(title) + '"' +
            ' aria-label="' + escapeHtml(title) + '">' +
            escapeHtml(label) +
            '</button></span>';
    }

    function deviceMetaLine(row, deviceCounts) {
        const tag = resolvePeerDeviceCode(row);
        if (!tag) {
            return '';
        }
        const did = String((row && row.deviceId) || '');
        const count = did && deviceCounts ? (deviceCounts[did] || 0) : 0;
        let line = 'Thiết bị ' + tag;
        if (did && deviceId && did === deviceId) {
            line += ' · máy này [' + (count > 0 ? count : 1) + ']';
        }
        return line;
    }

    /** Khoảng cách tới You — GPS nếu có; không thì ước lượng GeoIP. */
    function distanceFromYouLine(row, skip) {
        if (skip) {
            return '';
        }
        const ip = rowIp(row);
        if (ip) {
            ensureGeoForIp(ip);
        }
        const geo = geoFromRow(row);
        if (!geo || !selfGeo || !Number.isFinite(selfGeo.lat) || !Number.isFinite(selfGeo.lon)
            || !Number.isFinite(geo.lat) || !Number.isFinite(geo.lon)) {
            return '';
        }
        const km = haversineKm(selfGeo.lat, selfGeo.lon, geo.lat, geo.lon);
        if (publicIp && ip && publicIp === ip) {
            return '~cùng IP công khai';
        }
        const bothPrecise = !!(selfGeo.precise && geo.precise);
        if (!bothPrecise && km < 3) {
            return '~cùng khu vực (ước lượng IP)';
        }
        return '~' + formatDistanceKm(km) + ' từ You'
            + (bothPrecise ? '' : ' (ước lượng IP)');
    }

    function buildLabel() {
        return formatUserLabel(publicIp, displayCode || (deviceTag + String(tabIndex || 1)));
    }

    /** Tọa độ trình duyệt (GPS/Wi‑Fi) — chi tiết hơn tâm thành phố GeoIP. */
    function fetchBrowserPreciseGeo() {
        return new Promise((resolve) => {
            if (typeof navigator === 'undefined' || !navigator.geolocation) {
                resolve(null);
                return;
            }
            let settled = false;
            const done = (v) => {
                if (settled) {
                    return;
                }
                settled = true;
                resolve(v);
            };
            try {
                navigator.geolocation.getCurrentPosition(
                    (pos) => {
                        const lat = Number(pos && pos.coords && pos.coords.latitude);
                        const lon = Number(pos && pos.coords && pos.coords.longitude);
                        if (!Number.isFinite(lat) || !Number.isFinite(lon)) {
                            done(null);
                            return;
                        }
                        done({
                            lat: lat,
                            lon: lon,
                            accuracyM: Number(pos.coords.accuracy),
                            precise: true,
                            geoSource: 'gps'
                        });
                    },
                    () => done(null),
                    {
                        enableHighAccuracy: true,
                        timeout: 14000,
                        maximumAge: 120000
                    }
                );
            } catch (e) {
                done(null);
            }
        });
    }

    function applyPreciseOverride(baseGeo, precise) {
        if (!precise || !Number.isFinite(precise.lat) || !Number.isFinite(precise.lon)) {
            return baseGeo || null;
        }
        const base = baseGeo && typeof baseGeo === 'object' ? baseGeo : {};
        return {
            country: base.country || '',
            countryCode: base.countryCode || '',
            region: base.region || '',
            city: base.city || '',
            lat: precise.lat,
            lon: precise.lon,
            precise: true,
            geoSource: 'gps',
            accuracyM: Number.isFinite(precise.accuracyM) ? precise.accuracyM : undefined,
            ip: base.ip
        };
    }

    function fetchSelfGeo() {
        return fetchJson(GEO_PROVIDERS.ipwhoSelf)
            .then((j) => parseGeoFromPayload(j))
            .then((g) => {
                if (g) {
                    return g;
                }
                return fetchJson(GEO_PROVIDERS.geojsSelf).then((j2) => parseGeoFromPayload(j2));
            })
            .then((g) => {
                if (g) {
                    if (g.ip) {
                        publicIp = String(g.ip);
                    }
                    selfGeo = Object.assign({}, g, {
                        precise: false,
                        geoSource: 'ip'
                    });
                    if (publicIp) {
                        geoCache.set(publicIp, Object.assign({}, selfGeo));
                    }
                }
                return fetchBrowserPreciseGeo();
            })
            .then((precise) => {
                if (precise) {
                    selfGeo = applyPreciseOverride(selfGeo, precise);
                }
            })
            .catch(() => { /* offline / blocked */ });
    }

    function offlinePayload(offlineAt) {
        const now = Date.now();
        return attachGeoFields({
            online: false,
            label: buildLabel(),
            code: displayCode || (deviceTag + String(tabIndex || 1)),
            startedAt: startedAt || now,
            offlineAt: Number(offlineAt) || now,
            updatedAt: now
        });
    }

    function writePresence(online) {
        if (!sessionRef) {
            return Promise.resolve();
        }
        refreshTabSlot(sessionId);
        touchAlive();
        const now = Date.now();
        if (!online) {
            return sessionRef.set(offlinePayload(now));
        }
        lastPublishedFocusRowId = getLocalFocusRowId();
        return sessionRef.set(attachGeoFields({
            online: true,
            label: buildLabel(),
            code: displayCode || (deviceTag + String(tabIndex || 1)),
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
        if (m.deviceCode) {
            payload.deviceCode = String(m.deviceCode);
        }
        if (m.tabIndex != null) {
            payload.tabIndex = Number(m.tabIndex) || 0;
        }
        const focusId = normalizeFocusRowId(m.focusRowId);
        if (focusId) {
            payload.focusRowId = focusId;
        }
        Object.assign(payload, clientDevicePayloadFields(m));
        db.ref(PATH + '/' + id).set(payload)
            .catch(() => { /* ignore */ })
            .then(() => {
                setTimeout(() => tombstoneRequested.delete(id), 8000);
            });
    }

    function metaLine(started, endedAt) {
        const start = Number(started) || 0;
        const end = Number(endedAt) || Date.now();
        const access = formatAccessTime(start);
        const dur = start ? formatDuration(end - start) : '—';
        return 'Vào ' + access + ' · đã ' + dur;
    }

    function geoMetaLine(row, isSelf) {
        const ip = rowIp(row) || (isSelf ? publicIp : '');
        if (ip) {
            ensureGeoForIp(ip);
        }
        const geo = geoFromRow(row) || (isSelf ? selfGeo : null);
        if (!geo) {
            return '';
        }
        const place = formatGeoPlace(geo);
        const digs = geo.precise ? 5 : 4;
        const coords = '(' + formatCoord(geo.lat, digs) + ', ' + formatCoord(geo.lon, digs) + ')';
        const src = geo.precise ? 'GPS' : 'ước lượng IP';
        return place + ' · ' + coords + ' · ' + src;
    }

    function rowGeoFields(row) {
        return {
            country: String(row.country || '').trim(),
            countryCode: String(row.countryCode || '').trim(),
            region: normalizeVnPlace(row.region),
            city: normalizeVnPlace(row.city),
            lat: Number(row.lat),
            lon: Number(row.lon),
            ip: String(row.ip || '').trim() || extractIpFromLabel(row.label),
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
            code: displayCode || (deviceTag + String(tabIndex || 1)),
            startedAt: startedAt || Date.now(),
            offlineAt: SERVER_TS,
            updatedAt: SERVER_TS
        }));
    }

    /** Gom session → dòng list: online máy mình từng tab (AYU1…); offline không hiện tab cùng máy. */
    function buildListRows(sessionRows, deviceCounts, isOfflineList) {
        const groups = new Map();
        sessionRows.forEach((r) => {
            const did = String(r.deviceId || '').trim() || ('session:' + r.id);
            if (!groups.has(did)) {
                groups.set(did, []);
            }
            groups.get(did).push(r);
        });

        const out = [];
        groups.forEach((rows, did) => {
            rows.sort((a, b) => {
                const sa = Number(a.startedAt) || Number(a.updatedAt) || Number(a.at) || 0;
                const sb = Number(b.startedAt) || Number(b.updatedAt) || Number(b.at) || 0;
                if (sa !== sb) {
                    return sa - sb;
                }
                return String(a.id).localeCompare(String(b.id));
            });
            const mine = !!(deviceId && did === deviceId);
            if (mine) {
                /* Offline: bỏ hết tab cùng máy; online: vẫn hiện đủ từng tab */
                if (isOfflineList) {
                    return;
                }
                rows.forEach((r) => {
                    const code = String(r.code || '').trim()
                        || (resolvePeerDeviceCode(r) + String(r.tabIndex || ''));
                    const ip = extractIpFromRow(r);
                    out.push(Object.assign({}, r, {
                        listId: r.id,
                        displayLabel: formatUserLabel(ip, code),
                        deviceTag: resolvePeerDeviceCode(r),
                        aggregate: false,
                        tabCount: deviceCounts[did] || rows.length
                    }));
                });
                return;
            }
            const rep = isOfflineList
                ? rows.slice().sort((a, b) => (Number(b.at) || 0) - (Number(a.at) || 0))[0]
                : pickBestGeoRow(rows);
            const tag = resolvePeerDeviceCode(rep);
            const ip = extractIpFromRow(rep);
            out.push(Object.assign({}, rep, {
                listId: 'device:' + did,
                displayLabel: formatUserLabel(ip, tag),
                deviceTag: tag,
                deviceId: String(rep.deviceId || did.replace(/^session:/, '')),
                aggregate: true,
                tabCount: rows.length,
                startedAt: Number(rows[0].startedAt) || Number(rows[0].updatedAt) || 0,
                at: isOfflineList
                    ? Math.max.apply(null, rows.map((x) => Number(x.at) || 0))
                    : rep.at
            }));
        });

        if (isOfflineList) {
            out.sort((a, b) => (Number(a.at) || 0) - (Number(b.at) || 0));
        } else {
            out.sort((a, b) => {
                const sa = Number(a.startedAt) || Number(a.updatedAt) || 0;
                const sb = Number(b.startedAt) || Number(b.updatedAt) || 0;
                if (sa !== sb) {
                    return sa - sb;
                }
                return String(a.listId).localeCompare(String(b.listId));
            });
        }
        return out;
    }

    /** Badge đỏ trên nút presence: tổng tin nhắn chưa đọc từ mọi user khác. */
    function syncUnreadBadge(btn) {
        let total = 0;
        try {
            if (window.DeviceChat && typeof window.DeviceChat.getUnreadMap === 'function') {
                const m = window.DeviceChat.getUnreadMap();
                Object.keys(m || {}).forEach((k) => {
                    total += Number(m[k]) || 0;
                });
            }
        } catch (e) { /* ignore */ }
        let badge = el('presenceUnreadBadge');
        if (!total) {
            if (badge) {
                badge.hidden = true;
            }
            return;
        }
        if (!badge) {
            badge = document.createElement('span');
            badge.id = 'presenceUnreadBadge';
            badge.className = 'presence-unread-badge';
            badge.setAttribute('aria-label', 'Tin nhắn chưa đọc');
            btn.appendChild(badge);
        }
        badge.hidden = false;
        badge.textContent = total > 99 ? '99+' : String(total);
    }

    function renderPresence(snapVal) {
        lastSnapVal = snapVal && typeof snapVal === 'object' ? snapVal : {};
        if (deviceId && reconcileUniqueDeviceCode(lastSnapVal)) {
            if (sessionRef) {
                writePresence(true)
                    .then(() => armOnDisconnect())
                    .catch(() => { /* ignore */ });
            }
        }
        const onlineRows = [];
        const offlineRows = [];
        const data = lastSnapVal;
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
            const rowDeviceTag = resolvePeerDeviceCode(Object.assign({}, row, {
                deviceTag: row.deviceTag || row.deviceCode,
                deviceCode: row.deviceCode || row.deviceTag
            }));
            const rowTabIndex = Number(row.tabIndex) || 0;
            const rowFocusRowId = normalizeFocusRowId(row.focusRowId);
            const clientFields = pickClientDevice(row);
            nextMeta[id] = Object.assign({
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
                deviceTag: rowDeviceTag,
                deviceCode: rowDeviceTag,
                tabIndex: rowTabIndex,
                focusRowId: rowFocusRowId
            }, clientFields);

            const alive = !!row.online && updatedAt > 0 && (now - updatedAt) <= STALE_ONLINE_MS;
            if (alive) {
                nextOnlineIds.add(id);
                onlineRows.push(Object.assign({
                    id: id,
                    label: label,
                    code: code,
                    startedAt: rowStarted,
                    updatedAt: updatedAt,
                    deviceId: rowDeviceId,
                    deviceTag: rowDeviceTag,
                    deviceCode: rowDeviceTag,
                    tabIndex: rowTabIndex,
                    focusRowId: rowFocusRowId
                }, geoFields, clientFields));
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
                code: code,
                startedAt: rowStarted,
                at: offlineAt,
                deviceId: rowDeviceId,
                deviceTag: rowDeviceTag,
                deviceCode: rowDeviceTag,
                tabIndex: rowTabIndex,
                focusRowId: rowFocusRowId
            }, geoFields, clientFields));
            pendingGone.delete(id);
        });

        const onlineDeviceIds = new Set();
        onlineRows.forEach((r) => {
            const did = String(r.deviceId || '').trim();
            if (did) {
                onlineDeviceIds.add(did);
            }
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
            const metaDid = String(meta.deviceId || '').trim();
            if (metaDid && onlineDeviceIds.has(metaDid)) {
                pendingGone.delete(id);
                return;
            }
            const at = now;
            if (!pendingGone.has(id)) {
                pendingGone.set(id, Object.assign({
                    label: meta.label || id,
                    code: meta.code || '',
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
                    deviceTag: meta.deviceTag,
                    deviceCode: meta.deviceCode || meta.deviceTag,
                    tabIndex: meta.tabIndex,
                    focusRowId: meta.focusRowId
                }, pickClientDevice(meta)));
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
            const entryDid = String(entry.deviceId || '').trim();
            if (entryDid && onlineDeviceIds.has(entryDid)) {
                pendingGone.delete(id);
                return;
            }
            offlineRows.push(Object.assign({
                id: id,
                label: entry.label,
                code: entry.code || '',
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
                deviceTag: entry.deviceTag,
                deviceCode: entry.deviceCode || entry.deviceTag,
                tabIndex: entry.tabIndex,
                focusRowId: entry.focusRowId
            }, pickClientDevice(entry)));
        });

        const reconciledOfflineRows = offlineRows.filter((r) => {
            const did = String(r.deviceId || '').trim();
            if (did && onlineDeviceIds.has(did)) {
                requestRemoveNode(r.id);
                pendingGone.delete(r.id);
                return false;
            }
            return true;
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

        const uniqueDeviceCount = Object.keys(deviceCounts).length
            || new Set(onlineRows.map((r) => r.deviceId || r.id)).size;

        const listOnline = buildListRows(onlineRows, deviceCounts, false);
        const listOffline = buildListRows(reconciledOfflineRows, deviceCounts, true);

        const countEl = el('presenceCount');
        if (countEl) {
            countEl.textContent = String(uniqueDeviceCount);
        }
        const btn = el('presenceBtn');
        if (btn) {
            const title = uniqueDeviceCount + ' thiết bị đang online — bấm để xem list';
            btn.title = title;
            btn.setAttribute('aria-label', title);
            syncUnreadBadge(btn);
        }

        const list = el('presenceList');
        if (!list) {
            return;
        }
        hideDeviceTip();
        let html = '';
        listOnline.forEach((r) => {
            const isSelf = r.id === sessionId;
            const sameDevice = !!(r.deviceId && deviceId && r.deviceId === deviceId);
            const canChat = !!(r.deviceId && !sameDevice);
            const you = isSelf ? ' <span class="presence-you">(You)</span>' : '';
            const geoLine = geoMetaLine(r, isSelf || sameDevice);
            const place = formatGeoPlace(geoFromRow(r) || ((isSelf || sameDevice) ? selfGeo : null));
            const distLine = distanceFromYouLine(r, isSelf || sameDevice);
            const deviceHint = deviceMetaHtml(r, deviceCounts, isSelf);
            let deviceMetaInner = deviceHint;
            if (distLine) {
                deviceMetaInner += (deviceMetaInner ? ' · ' : '') + escapeHtml(distLine);
            }
            deviceMetaInner += you;
            const focusIdShow = isSelf
                ? (getLocalFocusRowId() || r.focusRowId)
                : r.focusRowId;
            html += '<li class="presence-item presence-item--online' +
                (isSelf ? ' presence-item--self' : '') +
                (sameDevice ? ' presence-item--same-device' : '') +
                (!sameDevice ? chatBorderClass(r.deviceId) : '') + '"' +
                (!sameDevice ? chatBorderAttr(r.deviceId) : '') +
                ' data-device-id="' + escapeHtml(r.deviceId || '') + '"' +
                ' data-device-tag="' + escapeHtml(r.deviceTag || '') + '"' +
                ' data-label="' + escapeHtml(r.displayLabel || r.label || '') + '"' +
                ' data-place="' + escapeHtml(place && place !== '—' ? place : '') + '"' +
                ' data-online="1">' +
                personColHtml(false, {
                    canChat: canChat,
                    unread: canChat ? getUnreadForDevice(r.deviceId) : 0,
                    deviceId: r.deviceId || '',
                    deviceTag: r.deviceTag || resolvePeerDeviceCode(r) || ''
                }, focusIdShow) +
                '<span class="presence-text">' +
                '<span class="presence-label">' + escapeHtml(r.displayLabel || r.label) +
                ' <span class="presence-status presence-status--online">đang online</span></span>' +
                (geoLine
                    ? '<span class="presence-meta presence-meta--geo">' + escapeHtml(geoLine) + '</span>'
                    : '') +
                (deviceMetaInner
                    ? '<span class="presence-meta presence-meta--device">' + deviceMetaInner + '</span>'
                    : '') +
                '<span class="presence-meta">' +
                escapeHtml(metaLine(r.startedAt, null)) + '</span>' +
                '</span></li>';
        });
        listOffline.forEach((r) => {
            const isSelf = r.id === sessionId;
            const sameDevice = !!(r.deviceId && deviceId && r.deviceId === deviceId);
            const canChat = !!(r.deviceId && !sameDevice);
            const you = isSelf ? ' <span class="presence-you">(You)</span>' : '';
            const geoLine = geoMetaLine(r, isSelf || sameDevice);
            const place = formatGeoPlace(geoFromRow(r) || ((isSelf || sameDevice) ? selfGeo : null));
            const distLine = distanceFromYouLine(r, isSelf || sameDevice);
            const deviceHint = deviceMetaHtml(r, null, isSelf);
            let deviceMetaInner = deviceHint;
            if (distLine) {
                deviceMetaInner += (deviceMetaInner ? ' · ' : '') + escapeHtml(distLine);
            }
            deviceMetaInner += you;
            html += '<li class="presence-item presence-item--offline' +
                (isSelf ? ' presence-item--self' : '') +
                (sameDevice ? ' presence-item--same-device' : '') +
                (!sameDevice ? chatBorderClass(r.deviceId) : '') + '"' +
                (!sameDevice ? chatBorderAttr(r.deviceId) : '') +
                ' data-device-id="' + escapeHtml(r.deviceId || '') + '"' +
                ' data-device-tag="' + escapeHtml(r.deviceTag || '') + '"' +
                ' data-label="' + escapeHtml(r.displayLabel || r.label || '') + '"' +
                ' data-place="' + escapeHtml(place && place !== '—' ? place : '') + '"' +
                ' data-online="0">' +
                personColHtml(true, {
                    canChat: canChat,
                    unread: canChat ? getUnreadForDevice(r.deviceId) : 0,
                    deviceId: r.deviceId || '',
                    deviceTag: r.deviceTag || resolvePeerDeviceCode(r) || ''
                }, r.focusRowId) +
                '<span class="presence-text">' +
                '<span class="presence-label">' + escapeHtml(r.displayLabel || r.label) +
                ' <span class="presence-status presence-status--offline">vừa offline</span></span>' +
                (geoLine
                    ? '<span class="presence-meta presence-meta--geo">' + escapeHtml(geoLine) + '</span>'
                    : '') +
                (deviceMetaInner
                    ? '<span class="presence-meta presence-meta--device">' + deviceMetaInner + '</span>'
                    : '') +
                '<span class="presence-meta">' +
                escapeHtml(metaLine(r.startedAt, r.at)) + '</span>' +
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
            panel.classList.toggle('hidden');
            const closed = panel.classList.contains('hidden');
            btn.setAttribute('aria-expanded', closed ? 'false' : 'true');
            if (closed) {
                hideDeviceTip();
            }
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
            hideDeviceTip();
        });
        if (list && list.dataset.chatDelegate !== '1') {
            list.dataset.chatDelegate = '1';
            list.addEventListener('click', (e) => {
                const focusBtn = e.target && e.target.closest
                    ? e.target.closest('[data-focus-row-id]')
                    : null;
                if (focusBtn) {
                    e.preventDefault();
                    e.stopPropagation();
                    const rid = String(focusBtn.getAttribute('data-focus-row-id') || '').trim();
                    if (rid) {
                        focusSheetRowById(rid);
                    }
                    return;
                }
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
                        place: String(row.getAttribute('data-place') || ''),
                        online: row.getAttribute('data-online') === '1'
                    });
                }
            });
        }
        bindDeviceTip(list, panel);
        if (!window.__presenceFocusRowBound) {
            window.__presenceFocusRowBound = '1';
            window.addEventListener('rowClicked', (ev) => {
                const detail = (ev && ev.detail) || {};
                const fromId = detail.clickedRowId != null
                    ? detail.clickedRowId
                    : (detail.focusRowId != null ? detail.focusRowId : '');
                const next = normalizeFocusRowId(fromId);
                if (next) {
                    localFocusRowId = next;
                } else if (typeof detail.focusRowIndex === 'number' && typeof window.getSheetFocusRowId === 'function') {
                    localFocusRowId = normalizeFocusRowId(window.getSheetFocusRowId());
                }
                publishFocusRowIdSoon();
                if (lastSnapVal != null) {
                    renderPresence(lastSnapVal);
                }
            });
        }
    }

    function ensureDeviceTipEl() {
        let tip = el('presenceDeviceTip');
        if (tip) {
            return tip;
        }
        tip = document.createElement('div');
        tip.id = 'presenceDeviceTip';
        tip.className = 'presence-device-tip-float';
        tip.setAttribute('role', 'tooltip');
        tip.setAttribute('aria-hidden', 'true');
        /* Inline style: không phụ thuộc CSS cache trên GitHub Pages */
        tip.style.cssText = [
            'position:fixed',
            'z-index:2147483646',
            'display:none',
            'max-width:260px',
            'padding:8px 10px',
            'border:1px solid #334155',
            'border-radius:8px',
            'background:#0f172a',
            'color:#f8fafc',
            'font:11px/1.45 Segoe UI,Arial,sans-serif',
            'white-space:pre-line',
            'box-shadow:0 10px 24px rgba(15,23,42,0.28)',
            'pointer-events:none',
            'left:0',
            'top:0'
        ].join(';');
        document.body.appendChild(tip);
        return tip;
    }

    function hideDeviceTip() {
        const tip = el('presenceDeviceTip');
        if (!tip) {
            return;
        }
        tip.style.display = 'none';
        tip.setAttribute('aria-hidden', 'true');
        tip.textContent = '';
        tip.classList.add('hidden');
    }

    function showDeviceTipFor(hint) {
        const text = hint && hint.getAttribute('data-tip');
        if (!text) {
            return;
        }
        const tip = ensureDeviceTipEl();
        tip.textContent = text;
        tip.classList.remove('hidden');
        tip.style.display = 'block';
        tip.setAttribute('aria-hidden', 'false');
        const rect = hint.getBoundingClientRect();
        const tipW = tip.offsetWidth || 180;
        const tipH = tip.offsetHeight || 80;
        let left = rect.left;
        let top = rect.bottom + 6;
        if (left + tipW > window.innerWidth - 8) {
            left = window.innerWidth - tipW - 8;
        }
        if (left < 8) {
            left = 8;
        }
        if (top + tipH > window.innerHeight - 8) {
            top = rect.top - tipH - 6;
        }
        if (top < 8) {
            top = 8;
        }
        tip.style.left = Math.round(left) + 'px';
        tip.style.top = Math.round(top) + 'px';
    }

    function bindDeviceTip(list, panel) {
        if (!list || list.dataset.deviceTipBound === '1') {
            return;
        }
        list.dataset.deviceTipBound = '1';
        ensureDeviceTipEl();

        const onEnter = (e) => {
            const hint = e.target && e.target.closest
                ? e.target.closest('.presence-device-hint[data-tip]')
                : null;
            if (!hint || !list.contains(hint)) {
                return;
            }
            showDeviceTipFor(hint);
        };
        const onLeave = (e) => {
            const hint = e.target && e.target.closest
                ? e.target.closest('.presence-device-hint[data-tip]')
                : null;
            if (!hint || !list.contains(hint)) {
                return;
            }
            const next = e.relatedTarget;
            if (next && hint.contains(next)) {
                return;
            }
            hideDeviceTip();
        };
        list.addEventListener('mouseover', onEnter);
        list.addEventListener('mouseout', onLeave);
        list.addEventListener('focusin', onEnter);
        list.addEventListener('focusout', onLeave);
        if (panel) {
            panel.addEventListener('scroll', hideDeviceTip, { passive: true });
        }
        window.addEventListener('scroll', hideDeviceTip, { passive: true });
        window.addEventListener('resize', hideDeviceTip);
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
            markAway();
            releaseTabSlot(sessionId);
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

    function startAliveLoop() {
        onPresenceIdleTick();
        if (aliveTimer) {
            clearInterval(aliveTimer);
        }
        aliveTimer = setInterval(() => {
            onPresenceIdleTick();
            if (!document.hidden && sessionId) {
                refreshTabSlot(sessionId);
            }
        }, 5000);
        if (!window.__presenceIdleBound) {
            window.__presenceIdleBound = true;
            document.addEventListener('visibilitychange', onPresenceIdleTick);
        }
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
        selfClientDevice = detectClientDeviceSync();
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
            sessionId = makeSessionId();
            tabIndex = claimTabIndex(sessionId);
            startedAt = Date.now();

            const beginPresence = (initialData) => {
                deviceTag = getOrCreateDeviceCode();
                reconcileUniqueDeviceCode(initialData && typeof initialData === 'object' ? initialData : {});
                displayCode = deviceTag + String(tabIndex);
                sessionRef = db.ref(PATH + '/' + sessionId);
                startAliveLoop();

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

                lastSnapVal = initialData && typeof initialData === 'object' ? initialData : {};
                renderPresence(lastSnapVal);

                fetchSelfGeo().then(() => {
                    writePresence(true).then(() => armOnDisconnect());
                }).catch(() => { /* ignore */ });

                enrichClientDeviceHints().then((changed) => {
                    if (!changed || !sessionRef) {
                        return;
                    }
                    writePresence(true).then(() => armOnDisconnect());
                });

                window.addEventListener('pagehide', markSelfOfflineBestEffort);
                window.addEventListener('beforeunload', markSelfOfflineBestEffort);

                window.PresenceBridge = {
                    getDeviceId: () => deviceId,
                    getDeviceTag: () => deviceTag,
                    getSessionId: () => sessionId,
                    getTabIndex: () => tabIndex,
                    getDisplayCode: () => displayCode,
                    getFocusRowId: () => getLocalFocusRowId(),
                    setFocusRowId: (rawId) => {
                        localFocusRowId = normalizeFocusRowId(rawId);
                        publishFocusRowIdSoon();
                    },
                    rerender: () => {
                        renderPresence(lastSnapVal && typeof lastSnapVal === 'object' ? lastSnapVal : {});
                    }
                };
            };

            db.ref(PATH).once('value')
                .then((snap) => beginPresence(snap.val()))
                .catch(() => beginPresence({}));
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
