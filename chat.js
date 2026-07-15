/**
 * Chat 1–1 theo deviceId (không theo tab/session).
 * Cùng máy / trình duyệt → cùng deviceId → mọi tab đồng bộ tin qua Firebase.
 * UI mở từ list presence (nút Nhắn). Hướng dẫn rules: PRESENCE.md
 */
(function () {
    'use strict';

    const DEVICE_STORAGE_KEY = 'presenceDeviceId';
    const CHAT_PATH = 'chats';
    const INBOX_PATH = 'inbox';
    const MAX_TEXT = 500;
    const MSG_LIMIT = 80;
    const BC_NAME = 'device-chat-v1';

    let db = null;
    let myDeviceId = '';
    let myDeviceTag = '';
    let peer = null;
    let messagesRef = null;
    let messagesHandler = null;
    let inboxRef = null;
    let started = false;
    let lastMessages = [];
    let bc = null;

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

    function pairId(a, b) {
        const x = String(a || '');
        const y = String(b || '');
        if (!x || !y) {
            return '';
        }
        return x < y ? (x + '__' + y) : (y + '__' + x);
    }

    function readReceipts() {
        try {
            const raw = localStorage.getItem('chatReadAt');
            const j = raw ? JSON.parse(raw) : {};
            return j && typeof j === 'object' ? j : {};
        } catch (e) {
            return {};
        }
    }

    function writeReceipts(map) {
        try {
            localStorage.setItem('chatReadAt', JSON.stringify(map || {}));
        } catch (e) { /* ignore */ }
    }

    function markPairRead(pid, at) {
        if (!pid) {
            return;
        }
        const map = readReceipts();
        map[pid] = Math.max(Number(map[pid]) || 0, Number(at) || Date.now());
        writeReceipts(map);
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

    function refreshUnreadFromInbox(inboxVal) {
        const data = inboxVal && typeof inboxVal === 'object' ? inboxVal : {};
        const receipts = readReceipts();
        let unread = 0;
        Object.keys(data).forEach((fromId) => {
            if (fromId === myDeviceId) {
                return;
            }
            const row = data[fromId] || {};
            const at = Number(row.at) || 0;
            const pid = String(row.pairId || pairId(myDeviceId, fromId));
            const readAt = Number(receipts[pid]) || 0;
            if (at > readAt) {
                unread += 1;
            }
        });
        // Nếu đang mở đúng peer đó thì không tính unread của họ
        if (peer && peer.deviceId && data[peer.deviceId]) {
            const row = data[peer.deviceId] || {};
            const at = Number(row.at) || 0;
            const pid = String(row.pairId || pairId(myDeviceId, peer.deviceId));
            const readAt = Number(receipts[pid]) || 0;
            if (at > readAt) {
                unread = Math.max(0, unread - 1);
            }
        }
        setUnreadBadge(unread);
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

    function detachMessages() {
        if (messagesRef && messagesHandler) {
            try {
                messagesRef.off('value', messagesHandler);
            } catch (e) { /* ignore */ }
        }
        messagesRef = null;
        messagesHandler = null;
        lastMessages = [];
    }

    function renderMessages(rows) {
        const box = el('chatMessages');
        if (!box) {
            return;
        }
        if (!rows.length) {
            box.innerHTML = '<div class="chat-empty">Chưa có tin nhắn — gửi tin đầu tiên.</div>';
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
            bits.push(peer.online ? 'đang online' : 'vừa offline / offline');
            bits.push('đồng bộ mọi tab trên máy bạn');
            sub.textContent = bits.join(' · ');
        }
    }

    function subscribeMessages(pid) {
        detachMessages();
        if (!db || !pid) {
            return;
        }
        messagesRef = db.ref(CHAT_PATH + '/' + pid + '/messages')
            .orderByChild('at')
            .limitToLast(MSG_LIMIT);
        messagesHandler = (snap) => {
            const val = snap.val() || {};
            const rows = Object.keys(val).map((id) => {
                const row = val[id] || {};
                return {
                    id: id,
                    fromDeviceId: String(row.fromDeviceId || ''),
                    toDeviceId: String(row.toDeviceId || ''),
                    text: String(row.text || ''),
                    at: Number(row.at) || 0
                };
            }).filter((m) => m.text);
            rows.sort((a, b) => a.at - b.at || a.id.localeCompare(b.id));
            lastMessages = rows;
            renderMessages(rows);
            if (rows.length) {
                markPairRead(pid, rows[rows.length - 1].at);
            } else {
                markPairRead(pid, Date.now());
            }
            if (inboxRef) {
                inboxRef.once('value').then((s) => refreshUnreadFromInbox(s.val()));
            }
        };
        messagesRef.on('value', messagesHandler);
    }

    function openChat(opts) {
        const peerId = String((opts && opts.deviceId) || '').trim();
        // Cấm chat với chính thiết bị mình (kể cả tab khác cùng máy)
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
        updateHeader();
        setChatOpen(true);
        const pid = pairId(myDeviceId, peerId);
        subscribeMessages(pid);
        markPairRead(pid, Date.now());
        if (db) {
            db.ref(INBOX_PATH + '/' + myDeviceId + '/' + peerId).remove().catch(() => { /* ignore */ });
        }
        if (inboxRef) {
            inboxRef.once('value').then((s) => refreshUnreadFromInbox(s.val()));
        }
        broadcast({ type: 'open', peer: peer });
    }

    function closeChat(opts) {
        const silent = !!(opts && opts.silent);
        setChatOpen(false);
        detachMessages();
        peer = null;
        if (!silent) {
            broadcast({ type: 'close' });
        }
    }

    function sendMessage(text) {
        const raw = String(text || '').trim();
        if (!raw || !peer || !db || !myDeviceId) {
            return Promise.resolve(false);
        }
        if (!peer.deviceId || peer.deviceId === myDeviceId) {
            console.warn('[chat] blocked send: same device');
            return Promise.resolve(false);
        }
        if (raw.length > MAX_TEXT) {
            return Promise.resolve(false);
        }
        const pid = pairId(myDeviceId, peer.deviceId);
        if (!pid || pid.indexOf(myDeviceId + '__' + myDeviceId) === 0) {
            return Promise.resolve(false);
        }
        const now = Date.now();
        const msg = {
            fromDeviceId: myDeviceId,
            toDeviceId: peer.deviceId,
            fromDeviceTag: myDeviceTag,
            text: raw,
            at: now
        };
        const msgRef = db.ref(CHAT_PATH + '/' + pid + '/messages').push();
        return msgRef.set(msg)
            .then(() => db.ref(CHAT_PATH + '/' + pid + '/meta').update({
                a: myDeviceId < peer.deviceId ? myDeviceId : peer.deviceId,
                b: myDeviceId < peer.deviceId ? peer.deviceId : myDeviceId,
                updatedAt: now,
                lastFrom: myDeviceId,
                lastPreview: raw.slice(0, 80)
            }))
            .then(() => db.ref(INBOX_PATH + '/' + peer.deviceId + '/' + myDeviceId).set({
                pairId: pid,
                fromDeviceId: myDeviceId,
                fromDeviceTag: myDeviceTag,
                preview: raw.slice(0, 80),
                at: now
            }))
            .then(() => {
                markPairRead(pid, now);
                return true;
            })
            .catch((err) => {
                console.warn('[chat] send failed', err);
                return false;
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
            input.addEventListener('keydown', (e) => {
                e.stopPropagation();
            });
            input.addEventListener('mousedown', (e) => {
                e.stopPropagation();
            });
        }

        panel.addEventListener('click', (e) => {
            e.stopPropagation();
        });

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

    function startInbox() {
        if (!db || !myDeviceId) {
            return;
        }
        inboxRef = db.ref(INBOX_PATH + '/' + myDeviceId);
        inboxRef.on('value', (snap) => {
            refreshUnreadFromInbox(snap.val());
        });
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
        myDeviceId = getOrCreateDeviceId();
        myDeviceTag = makeDeviceTag(myDeviceId);
        bindUi();
        bindBroadcast();

        const boot = () => {
            try {
                if (!firebase.apps.length && isConfigReady(cfg)) {
                    firebase.initializeApp(cfg);
                }
            } catch (e) { /* already init */ }
            db = firebase.database();
            // Đồng bộ device id với PresenceBridge nếu đã sẵn
            if (window.PresenceBridge && typeof window.PresenceBridge.getDeviceId === 'function') {
                const id = window.PresenceBridge.getDeviceId();
                const tag = window.PresenceBridge.getDeviceTag && window.PresenceBridge.getDeviceTag();
                if (id) {
                    myDeviceId = id;
                    myDeviceTag = tag || makeDeviceTag(id);
                }
            }
            startInbox();
        };

        if (firebase.auth && firebase.auth().currentUser) {
            boot();
        } else if (firebase.auth) {
            firebase.auth().onAuthStateChanged(() => {
                boot();
            });
            // Presence thường đã sign-in; nếu chưa thì cố gắng anonymous
            firebase.auth().signInAnonymously().catch(() => {
                boot();
            });
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
