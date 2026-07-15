/**
 * Chat theo deviceId — lịch sử chỉ localStorage (không lưu bền trên Firebase).
 *
 * - Tin nằm trên máy khi thiết bị còn “sống” (có tab / < 10 phút kể từ tab cuối).
 * - A online, B offline: A vẫn giữ lịch sử local với B.
 * - B online lại + A còn online: A đẩy transcript qua mailbox tạm → B merge rồi xóa mailbox.
 * - Máy tắt hết tab > 10 phút rồi mở lại: xóa sạch lịch sử local trên máy đó.
 * - Cùng deviceId (hai tab cùng máy) không chat với nhau.
 *
 * Firebase chỉ dùng path /mailbox tạm (set rồi xóa), không phải kho chats.
 */
(function () {
    'use strict';

    const DEVICE_STORAGE_KEY = 'presenceDeviceId';
    const CHAT_STORE_KEY = 'deviceChatStore';
    const ALIVE_KEY = 'deviceChatAliveAt';
    const AWAY_KEY = 'deviceChatAwayAt';
    const READ_KEY = 'deviceChatReadAt';
    const MAILBOX_PATH = 'mailbox';
    const PRESENCE_PATH = 'presence';
    const MAX_TEXT = 500;
    const MSG_LIMIT = 120;
    const IDLE_WIPE_MS = 10 * 60 * 1000;
    const ALIVE_TICK_MS = 5000;
    const BC_NAME = 'device-chat-local-v1';

    let db = null;
    let myDeviceId = '';
    let myDeviceTag = '';
    let peer = null;
    let started = false;
    let aliveTimer = 0;
    let bc = null;
    let mailboxRef = null;
    let presenceRef = null;
    /** peerDeviceId → last push at (tránh spam mailbox) */
    const lastPushAt = {};
    let unreadMap = {};

    function el(id) {
        return document.getElementById(id);
    }

    function escapeHtml(text) {
        return String(text || '')
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;');
    }

    function makeSessionId() {
        try {
            if (crypto && typeof crypto.randomUUID === 'function') {
                return crypto.randomUUID();
            }
        } catch (e) { /* ignore */ }
        return 's' + Date.now().toString(36) + Math.random().toString(36).slice(2, 10);
    }

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

    function pad2(n) {
        return n < 10 ? '0' + n : String(n);
    }

    function formatMsgTime(ts) {
        const d = new Date(Number(ts) || 0);
        if (!Number.isFinite(d.getTime()) || d.getTime() <= 0) {
            return '';
        }
        return pad2(d.getHours()) + ':' + pad2(d.getMinutes());
    }

    function readJson(key, fallback) {
        try {
            const raw = localStorage.getItem(key);
            if (!raw) {
                return fallback;
            }
            const j = JSON.parse(raw);
            return j && typeof j === 'object' ? j : fallback;
        } catch (e) {
            return fallback;
        }
    }

    function writeJson(key, val) {
        try {
            localStorage.setItem(key, JSON.stringify(val));
        } catch (e) { /* quota */ }
    }

    function touchAlive() {
        try {
            localStorage.setItem(ALIVE_KEY, String(Date.now()));
            localStorage.removeItem(AWAY_KEY);
        } catch (e) { /* ignore */ }
    }

    function markAway() {
        try {
            localStorage.setItem(AWAY_KEY, String(Date.now()));
        } catch (e) { /* ignore */ }
    }

    /** Nếu máy này đã “chết” > 10 phút (không có tab sống) → xóa hết chat local. */
    function wipeIfDeviceIdleTooLong() {
        const now = Date.now();
        let aliveAt = 0;
        let awayAt = 0;
        try {
            aliveAt = Number(localStorage.getItem(ALIVE_KEY)) || 0;
            awayAt = Number(localStorage.getItem(AWAY_KEY)) || 0;
        } catch (e) { /* ignore */ }
        const last = Math.max(aliveAt, awayAt);
        if (last > 0 && (now - last) > IDLE_WIPE_MS) {
            try {
                localStorage.removeItem(CHAT_STORE_KEY);
                localStorage.removeItem(READ_KEY);
            } catch (e) { /* ignore */ }
            unreadMap = {};
            setUnreadBadge(0);
            return true;
        }
        return false;
    }

    function loadStore() {
        const store = readJson(CHAT_STORE_KEY, {});
        return store && typeof store === 'object' ? store : {};
    }

    function saveStore(store) {
        writeJson(CHAT_STORE_KEY, store || {});
        broadcast({ type: 'store' });
    }

    function getThread(peerId) {
        const store = loadStore();
        const t = store[peerId];
        if (!t || typeof t !== 'object') {
            return { messages: [], peerTag: '', peerLabel: '', updatedAt: 0 };
        }
        return {
            messages: Array.isArray(t.messages) ? t.messages.slice() : [],
            peerTag: String(t.peerTag || ''),
            peerLabel: String(t.peerLabel || ''),
            updatedAt: Number(t.updatedAt) || 0
        };
    }

    function setThread(peerId, thread) {
        if (!peerId || peerId === myDeviceId) {
            return;
        }
        const store = loadStore();
        const msgs = (thread.messages || []).slice(-MSG_LIMIT);
        store[peerId] = {
            messages: msgs,
            peerTag: String(thread.peerTag || ''),
            peerLabel: String(thread.peerLabel || ''),
            updatedAt: Number(thread.updatedAt) || Date.now()
        };
        saveStore(store);
    }

    function mergeMessages(localMsgs, incoming) {
        const map = new Map();
        (localMsgs || []).forEach((m) => {
            if (!m || !m.id) {
                return;
            }
            map.set(m.id, m);
        });
        (incoming || []).forEach((m) => {
            if (!m || !m.id || !m.text) {
                return;
            }
            if (!map.has(m.id)) {
                map.set(m.id, {
                    id: String(m.id),
                    fromDeviceId: String(m.fromDeviceId || ''),
                    toDeviceId: String(m.toDeviceId || ''),
                    text: String(m.text || '').slice(0, MAX_TEXT),
                    at: Number(m.at) || 0
                });
            }
        });
        return Array.from(map.values())
            .sort((a, b) => a.at - b.at || a.id.localeCompare(b.id))
            .slice(-MSG_LIMIT);
    }

    function msgId() {
        return 'm' + Date.now().toString(36) + '_' + Math.random().toString(36).slice(2, 10);
    }

    function setUnreadBadge(n) {
        const badge = el('chatUnreadBadge');
        if (!badge) {
            return;
        }
        const count = Math.max(0, Math.floor(Number(n) || 0));
        if (count <= 0) {
            badge.hidden = true;
            badge.textContent = '0';
            return;
        }
        badge.hidden = false;
        badge.textContent = count > 99 ? '99+' : String(count);
    }

    function refreshUnreadBadge() {
        let n = 0;
        Object.keys(unreadMap).forEach((k) => {
            if (peer && peer.deviceId === k) {
                return;
            }
            n += Number(unreadMap[k]) || 0;
        });
        setUnreadBadge(n);
    }

    function markPeerRead(peerId) {
        if (!peerId) {
            return;
        }
        unreadMap[peerId] = 0;
        const reads = readJson(READ_KEY, {});
        reads[peerId] = Date.now();
        writeJson(READ_KEY, reads);
        refreshUnreadBadge();
    }

    function bumpUnread(peerId) {
        if (!peerId || peerId === myDeviceId) {
            return;
        }
        if (peer && peer.deviceId === peerId && isChatOpen()) {
            markPeerRead(peerId);
            return;
        }
        unreadMap[peerId] = (Number(unreadMap[peerId]) || 0) + 1;
        refreshUnreadBadge();
    }

    function isChatOpen() {
        const panel = el('chatPanel');
        return !!(panel && !panel.classList.contains('hidden'));
    }

    function setChatOpen(open) {
        const panel = el('chatPanel');
        if (!panel) {
            return;
        }
        panel.classList.toggle('hidden', !open);
        if (open) {
            const presencePanel = el('presencePanel');
            const presenceBtn = el('presenceBtn');
            if (presencePanel) {
                presencePanel.classList.add('hidden');
            }
            if (presenceBtn) {
                presenceBtn.setAttribute('aria-expanded', 'false');
            }
            const input = el('chatInput');
            if (input) {
                setTimeout(() => {
                    try {
                        input.focus();
                    } catch (e) { /* ignore */ }
                }, 0);
            }
        }
    }

    function renderMessages(rows) {
        const box = el('chatMessages');
        if (!box) {
            return;
        }
        if (!rows.length) {
            box.innerHTML = '<div class="chat-empty">Chưa có tin — chỉ lưu trên máy khi còn tab; tắt máy ~10 phút sẽ mất.</div>';
            return;
        }
        let html = '';
        rows.forEach((m) => {
            const mine = m.fromDeviceId === myDeviceId;
            html += '<div class="chat-bubble-row' + (mine ? ' chat-bubble-row--mine' : '') + '">' +
                '<div class="chat-bubble' + (mine ? ' chat-bubble--mine' : '') + '">' +
                '<div class="chat-bubble-text">' + escapeHtml(m.text) + '</div>' +
                '<div class="chat-bubble-time">' + escapeHtml(formatMsgTime(m.at)) +
                (mine ? ' · You' : '') + '</div>' +
                '</div></div>';
        });
        box.innerHTML = html;
        box.scrollTop = box.scrollHeight;
    }

    function updateHeader() {
        const title = el('chatTitle');
        const sub = el('chatSubtitle');
        if (!peer) {
            return;
        }
        if (title) {
            title.textContent = peer.deviceTag
                ? ('Chat · ' + peer.deviceTag)
                : 'Chat';
        }
        if (sub) {
            const bits = [];
            if (peer.label) {
                bits.push(peer.label);
            }
            bits.push(peer.online ? 'đang online' : 'đối phương offline — tin giữ trên máy bạn');
            bits.push('local · hết hạn ~10 phút khi tắt máy');
            sub.textContent = bits.join(' · ');
        }
    }

    function refreshOpenThread() {
        if (!peer || !isChatOpen()) {
            return;
        }
        const t = getThread(peer.deviceId);
        renderMessages(t.messages);
    }

    function openChat(opts) {
        const peerId = String((opts && opts.deviceId) || '').trim();
        if (!peerId || !myDeviceId || peerId === myDeviceId) {
            console.warn('[chat] blocked: same device cannot chat with itself');
            return;
        }
        peer = {
            deviceId: peerId,
            deviceTag: String((opts && opts.deviceTag) || makeDeviceTag(peerId)),
            label: String((opts && opts.label) || ''),
            online: !!(opts && opts.online)
        };
        const t = getThread(peerId);
        t.peerTag = peer.deviceTag || t.peerTag;
        t.peerLabel = peer.label || t.peerLabel;
        setThread(peerId, t);
        updateHeader();
        setChatOpen(true);
        renderMessages(t.messages);
        markPeerRead(peerId);
        // Nếu mình còn lịch sử và peer online → đẩy sang peer
        if (peer.online && t.messages.length) {
            pushMailboxToPeer(peerId, t.messages);
        }
        broadcast({ type: 'open', peer: peer });
    }

    function closeChat(opts) {
        const silent = !!(opts && opts.silent);
        setChatOpen(false);
        peer = null;
        if (!silent) {
            broadcast({ type: 'close' });
        }
    }

    function sendMessage(text) {
        const raw = String(text || '').trim();
        if (!raw || !peer || !myDeviceId) {
            return Promise.resolve(false);
        }
        if (!peer.deviceId || peer.deviceId === myDeviceId) {
            return Promise.resolve(false);
        }
        if (raw.length > MAX_TEXT) {
            return Promise.resolve(false);
        }
        const now = Date.now();
        const m = {
            id: msgId(),
            fromDeviceId: myDeviceId,
            toDeviceId: peer.deviceId,
            text: raw,
            at: now
        };
        const t = getThread(peer.deviceId);
        t.messages = mergeMessages(t.messages, [m]);
        t.peerTag = peer.deviceTag || t.peerTag;
        t.peerLabel = peer.label || t.peerLabel;
        t.updatedAt = now;
        setThread(peer.deviceId, t);
        renderMessages(t.messages);
        markPeerRead(peer.deviceId);
        // Mailbox tạm: peer online thì nhận ngay; offline thì khi họ online lại + mình còn sống sẽ push lại
        return pushMailboxToPeer(peer.deviceId, t.messages).then(() => true);
    }

    function pushMailboxToPeer(peerId, messages) {
        if (!db || !peerId || peerId === myDeviceId) {
            return Promise.resolve();
        }
        const now = Date.now();
        if (lastPushAt[peerId] && (now - lastPushAt[peerId]) < 800) {
            return Promise.resolve();
        }
        lastPushAt[peerId] = now;
        const payload = {
            fromDeviceId: myDeviceId,
            fromDeviceTag: myDeviceTag,
            at: now,
            messages: (messages || []).slice(-MSG_LIMIT).map((m) => ({
                id: m.id,
                fromDeviceId: m.fromDeviceId,
                toDeviceId: m.toDeviceId,
                text: m.text,
                at: m.at
            }))
        };
        // Ghi tạm → peer đọc xong sẽ xóa. Không phải kho lịch sử.
        return db.ref(MAILBOX_PATH + '/' + peerId + '/' + myDeviceId).set(payload)
            .catch((err) => {
                console.warn('[chat] mailbox push failed', err);
            });
    }

    function normalizeMessageList(raw) {
        if (Array.isArray(raw)) {
            return raw;
        }
        if (raw && typeof raw === 'object') {
            return Object.keys(raw)
                .sort((a, b) => Number(a) - Number(b) || String(a).localeCompare(String(b)))
                .map((k) => raw[k]);
        }
        return [];
    }

    function consumeMailboxEntry(fromId, payload) {
        if (!fromId || fromId === myDeviceId || !payload) {
            return;
        }
        const incoming = normalizeMessageList(payload.messages);
        if (!incoming.length) {
            db.ref(MAILBOX_PATH + '/' + myDeviceId + '/' + fromId).remove().catch(() => { /* ignore */ });
            return;
        }
        const t = getThread(fromId);
        const before = t.messages.length;
        t.messages = mergeMessages(t.messages, incoming);
        t.peerTag = String(payload.fromDeviceTag || t.peerTag || makeDeviceTag(fromId));
        t.updatedAt = Date.now();
        setThread(fromId, t);
        const added = t.messages.length - before;
        if (added > 0) {
            if (peer && peer.deviceId === fromId && isChatOpen()) {
                renderMessages(t.messages);
                markPeerRead(fromId);
            } else {
                bumpUnread(fromId);
            }
        }
        // Xóa mailbox ngay sau khi merge — không để lâu trên Firebase
        db.ref(MAILBOX_PATH + '/' + myDeviceId + '/' + fromId).remove().catch(() => { /* ignore */ });

        // Nếu mình đang mở chat với họ và mình có bản đầy hơn → đẩy lại (hai chiều merge)
        if (peer && peer.deviceId === fromId && isChatOpen()) {
            pushMailboxToPeer(fromId, t.messages);
        }
    }

    function startMailboxListener() {
        if (!db || !myDeviceId) {
            return;
        }
        mailboxRef = db.ref(MAILBOX_PATH + '/' + myDeviceId);
        mailboxRef.on('value', (snap) => {
            const val = snap.val() || {};
            Object.keys(val).forEach((fromId) => {
                consumeMailboxEntry(fromId, val[fromId]);
            });
        });
    }

    /** Khi peer device xuất hiện online trên presence → đẩy transcript local (nếu còn). */
    function onPresenceDevices(onlineDeviceIds) {
        const set = onlineDeviceIds instanceof Set ? onlineDeviceIds : new Set(onlineDeviceIds || []);
        if (peer && peer.deviceId) {
            peer.online = set.has(peer.deviceId);
            updateHeader();
        }
        const store = loadStore();
        Object.keys(store).forEach((peerId) => {
            if (!set.has(peerId) || peerId === myDeviceId) {
                return;
            }
            const t = store[peerId];
            if (t && Array.isArray(t.messages) && t.messages.length) {
                pushMailboxToPeer(peerId, t.messages);
            }
        });
    }

    function collectOnlineDeviceIds(snapVal) {
        const ids = new Set();
        const data = snapVal && typeof snapVal === 'object' ? snapVal : {};
        const now = Date.now();
        Object.keys(data).forEach((sid) => {
            const row = data[sid] || {};
            const updatedAt = Number(row.updatedAt) || 0;
            const alive = !!row.online && updatedAt > 0 && (now - updatedAt) <= 90000;
            const did = String(row.deviceId || '').trim();
            if (alive && did) {
                ids.add(did);
            }
        });
        return ids;
    }

    function startPresenceWatch() {
        if (!db) {
            return;
        }
        presenceRef = db.ref(PRESENCE_PATH);
        presenceRef.on('value', (snap) => {
            onPresenceDevices(collectOnlineDeviceIds(snap.val()));
        });
    }

    function bindUi() {
        const panel = el('chatPanel');
        const closeBtn = el('chatCloseBtn');
        const form = el('chatForm');
        const input = el('chatInput');
        if (!panel || panel.dataset.chatBound === '1') {
            return;
        }
        panel.dataset.chatBound = '1';

        if (closeBtn) {
            closeBtn.addEventListener('click', (e) => {
                e.stopPropagation();
                closeChat();
            });
        }

        if (form && input) {
            form.addEventListener('submit', (e) => {
                e.preventDefault();
                e.stopPropagation();
                const text = input.value;
                input.value = '';
                sendMessage(text).then((ok) => {
                    if (!ok && text.trim()) {
                        input.value = text;
                    }
                });
            });
            input.addEventListener('keydown', (e) => e.stopPropagation());
            input.addEventListener('mousedown', (e) => e.stopPropagation());
        }

        panel.addEventListener('click', (e) => e.stopPropagation());

        document.addEventListener('click', (e) => {
            if (!isChatOpen()) {
                return;
            }
            const widget = el('presenceWidget');
            if (widget && widget.contains(e.target)) {
                return;
            }
            closeChat();
        });

        window.addEventListener('storage', (e) => {
            if (!e) {
                return;
            }
            if (e.key === CHAT_STORE_KEY) {
                refreshOpenThread();
            }
        });
    }

    function broadcast(payload) {
        try {
            if (bc) {
                bc.postMessage(payload);
            }
        } catch (e) { /* ignore */ }
    }

    function bindBroadcast() {
        try {
            if (typeof BroadcastChannel === 'undefined') {
                return;
            }
            bc = new BroadcastChannel(BC_NAME);
            bc.onmessage = (ev) => {
                const data = ev && ev.data;
                if (!data || typeof data !== 'object') {
                    return;
                }
                if (data.type === 'store') {
                    refreshOpenThread();
                    return;
                }
                if (data.type === 'open' && data.peer && data.peer.deviceId) {
                    if (data.peer.deviceId === myDeviceId) {
                        return;
                    }
                    if (peer && peer.deviceId === data.peer.deviceId && isChatOpen()) {
                        return;
                    }
                    openChat(data.peer);
                } else if (data.type === 'close') {
                    if (isChatOpen()) {
                        closeChat({ silent: true });
                    }
                }
            };
        } catch (e) { /* ignore */ }
    }

    function startAliveLoop() {
        touchAlive();
        if (aliveTimer) {
            clearInterval(aliveTimer);
        }
        aliveTimer = setInterval(touchAlive, ALIVE_TICK_MS);
        window.addEventListener('pagehide', markAway);
        window.addEventListener('beforeunload', markAway);
    }

    function startChat() {
        if (started) {
            return;
        }
        const cfg = window.PRESENCE_FIREBASE_CONFIG;
        if (!isConfigReady(cfg)) {
            return;
        }
        if (typeof firebase === 'undefined' || !firebase.database) {
            return;
        }

        started = true;
        wipeIfDeviceIdleTooLong();
        myDeviceId = getOrCreateDeviceId();
        myDeviceTag = makeDeviceTag(myDeviceId);
        bindUi();
        bindBroadcast();
        startAliveLoop();

        const boot = () => {
            try {
                if (!firebase.apps.length && isConfigReady(cfg)) {
                    firebase.initializeApp(cfg);
                }
            } catch (e) { /* already init */ }
            db = firebase.database();
            if (window.PresenceBridge && typeof window.PresenceBridge.getDeviceId === 'function') {
                const id = window.PresenceBridge.getDeviceId();
                const tag = window.PresenceBridge.getDeviceTag && window.PresenceBridge.getDeviceTag();
                if (id) {
                    myDeviceId = id;
                    myDeviceTag = tag || makeDeviceTag(id);
                }
            }
            startMailboxListener();
            startPresenceWatch();
        };

        if (firebase.auth && firebase.auth().currentUser) {
            boot();
        } else if (firebase.auth) {
            firebase.auth().onAuthStateChanged(() => boot());
            firebase.auth().signInAnonymously().catch(() => boot());
        } else {
            boot();
        }
    }

    window.DeviceChat = {
        open: openChat,
        close: closeChat,
        getMyDeviceId: () => myDeviceId,
        isOpen: isChatOpen
    };

    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', startChat);
    } else {
        startChat();
    }
})();
