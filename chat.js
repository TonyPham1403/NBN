/**
 * Chat theo deviceId — text + file.
 * - Meta/history: localStorage; chuyển tạm qua /mailbox rồi xóa.
 * - File blob: Firebase Storage (chat-files/…); cache máy: IndexedDB.
 * - Idle ~10 phút: xóa local (+ IDB); file trên Storage có thể còn đến khi xóa tay / lifecycle.
 * UI: dock Messenger đáy trang, tối đa 3 người.
 */
(function () {
    'use strict';

    const DEVICE_STORAGE_KEY = 'presenceDeviceId';
    const CHAT_STORE_KEY = 'deviceChatStore';
    const AVATAR_STORE_KEY = 'deviceChatAvatars';
    const ALIVE_KEY = 'deviceChatAliveAt';
    const AWAY_KEY = 'deviceChatAwayAt';
    const READ_KEY = 'deviceChatReadAt';
    const SEEN_KEY = 'deviceChatPeerSeenAt';
    /** Material 500 (kiểu avatar Gmail/Contacts) — chỉ màu no đậm, không xám/nâu nhạt. */
    const AVATAR_COLORS = [
        '#F44336', '#E91E63', '#9C27B0', '#673AB7',
        '#3F51B5', '#2196F3', '#03A9F4', '#00BCD4',
        '#009688', '#4CAF50', '#8BC34A', '#FFC107',
        '#FF9800', '#FF5722'
    ];
    const MAILBOX_PATH = 'mailbox';
    const SENT_FLASH_MS = 1400;
    const RECV_FLASH_MS = 1400;
    const MAX_PINS = 20;
    const RECALL_LABEL = 'Tin nhắn đã thu hồi';
    const PRESENCE_PATH = 'presence';
    const STORAGE_ROOT = 'chat-files';
    const IDB_NAME = 'deviceChatFilesDb';
    const IDB_STORE = 'blobs';
    const MAX_TEXT = 500;
    const MSG_LIMIT = 120;
    const MAX_DOCK = 3;
    const MAX_FILE_BYTES = 5 * 1024 * 1024;
    const MAX_INLINE_IMAGE_BYTES = 350 * 1024;
    const PREFER_INLINE_IMAGE_SEND = true;
    const MAX_PENDING_FILES = 10;
    const INPUT_MIN_H = 32;
    const INPUT_MAX_H = 110;
    const IDLE_WIPE_MS = 10 * 60 * 1000;
    const ALIVE_TICK_MS = 5000;
    const BC_NAME = 'device-chat-local-v1';

    let db = null;
    let storage = null;
    let myDeviceId = '';
    let myDeviceTag = '';
    let started = false;
    let aliveTimer = 0;
    let bc = null;
    let mailboxRef = null;
    let presenceRef = null;
    const lastPushAt = {};
    let unreadMap = {};
    /** peerId → thời điểm họ đã đọc tin của mình */
    let peerSeenAt = {};
    /** peerId → hết hạn flash "Đã nhận" */
    const receiveFlashUntil = {};
    let statusTicker = 0;
    /** peerId → true nếu đang neo đáy (auto-scroll); false khi user scroll xem tin cũ */
    const scrollPinned = {};
    const SCROLL_BOTTOM_SLACK = 56;
    /** @type {Array<{deviceId:string, deviceTag:string, label:string, place:string, online:boolean, minimized:boolean, openedAt:number}>} */
    let dockSessions = [];
    let onlineDeviceSet = new Set();
    const objectUrlCache = new Map();
    /** msgId → blob URL preview khi đang upload ảnh */
    const uploadPreviewUrls = new Map();
    const UPLOAD_TIMEOUT_MS = 120000;
    /** @type {Object.<string, Array<{id:string, file:File, previewUrl:string, name:string, size:number, mime:string}>>} */
    const pendingAttach = {};
    /** @type {Object.<string, {id:string, text:string, fromDeviceId:string}>} */
    const pendingReply = {};
    const REPLIED_STORE_KEY = 'deviceChatRepliedSession';
    /** Session: peers đã được phản hồi (gửi tin/file) — viền vàng */
    const repliedPeers = new Set();
    /** Peer phản hồi gần nhất — viền xanh (thay vàng) */
    let lastRepliedPeerId = '';
    /** deviceId → { color, letter } — giữ cùng cơ chế idle wipe với chat */
    let avatarStore = {};
    let avatarStoreLoaded = false;

    function loadRepliedSession() {
        try {
            const raw = sessionStorage.getItem(REPLIED_STORE_KEY);
            if (!raw) {
                return;
            }
            const j = JSON.parse(raw);
            if (!j || typeof j !== 'object') {
                return;
            }
            (Array.isArray(j.peers) ? j.peers : []).forEach((id) => {
                if (id) {
                    repliedPeers.add(String(id));
                }
            });
            if (j.last) {
                lastRepliedPeerId = String(j.last);
            }
        } catch (e) { /* ignore */ }
    }

    function saveRepliedSession() {
        try {
            sessionStorage.setItem(REPLIED_STORE_KEY, JSON.stringify({
                peers: Array.from(repliedPeers),
                last: lastRepliedPeerId || ''
            }));
        } catch (e) { /* ignore */ }
    }

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
        return String.fromCharCode(65 + (h % 26))
            + String.fromCharCode(65 + ((h >>> 5) % 26))
            + String.fromCharCode(65 + ((h >>> 10) % 26));
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

    function formatFileSize(bytes) {
        const n = Number(bytes) || 0;
        if (n < 1024) {
            return n + ' B';
        }
        if (n < 1024 * 1024) {
            return (n / 1024).toFixed(n < 10 * 1024 ? 1 : 0).replace(/\.0$/, '') + ' KB';
        }
        return (n / (1024 * 1024)).toFixed(1).replace(/\.0$/, '') + ' MB';
    }

    function formatChatPlace(place) {
        return String(place || '')
            .replace(/\s*·\s*/g, ' · ')
            .replace(/[ \t]+/g, ' ')
            .trim();
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

    function openIdb() {
        return new Promise((resolve, reject) => {
            if (typeof indexedDB === 'undefined') {
                reject(new Error('no idb'));
                return;
            }
            const req = indexedDB.open(IDB_NAME, 1);
            req.onupgradeneeded = () => {
                const idb = req.result;
                if (!idb.objectStoreNames.contains(IDB_STORE)) {
                    idb.createObjectStore(IDB_STORE);
                }
            };
            req.onsuccess = () => resolve(req.result);
            req.onerror = () => reject(req.error || new Error('idb open failed'));
        });
    }

    function idbPut(key, blob) {
        return openIdb().then((idb) => new Promise((resolve, reject) => {
            const tx = idb.transaction(IDB_STORE, 'readwrite');
            tx.objectStore(IDB_STORE).put(blob, key);
            tx.oncomplete = () => resolve();
            tx.onerror = () => reject(tx.error);
        })).catch(() => { /* ignore */ });
    }

    function idbGet(key) {
        return openIdb().then((idb) => new Promise((resolve, reject) => {
            const tx = idb.transaction(IDB_STORE, 'readonly');
            const req = tx.objectStore(IDB_STORE).get(key);
            req.onsuccess = () => resolve(req.result || null);
            req.onerror = () => reject(req.error);
        })).catch(() => null);
    }

    function idbClear() {
        return openIdb().then((idb) => new Promise((resolve, reject) => {
            const tx = idb.transaction(IDB_STORE, 'readwrite');
            tx.objectStore(IDB_STORE).clear();
            tx.oncomplete = () => resolve();
            tx.onerror = () => reject(tx.error);
        })).catch(() => { /* ignore */ });
    }

    function revokeObjectUrls() {
        objectUrlCache.forEach((url) => {
            try {
                URL.revokeObjectURL(url);
            } catch (e) { /* ignore */ }
        });
        objectUrlCache.clear();
        uploadPreviewUrls.forEach((url) => {
            try {
                URL.revokeObjectURL(url);
            } catch (e) { /* ignore */ }
        });
        uploadPreviewUrls.clear();
    }

    /** Hoạt động gần nhất: heartbeat, đóng tab, hoặc tin trong store. */
    function getLastActivityAt() {
        let aliveAt = 0;
        let awayAt = 0;
        try {
            aliveAt = Number(localStorage.getItem(ALIVE_KEY)) || 0;
            awayAt = Number(localStorage.getItem(AWAY_KEY)) || 0;
        } catch (e) { /* ignore */ }
        let chatAt = 0;
        try {
            const store = readJson(CHAT_STORE_KEY, {});
            Object.keys(store || {}).forEach((peerId) => {
                const t = store[peerId];
                if (!t || typeof t !== 'object') {
                    return;
                }
                chatAt = Math.max(chatAt, Number(t.updatedAt) || 0);
                (Array.isArray(t.messages) ? t.messages : []).forEach((m) => {
                    chatAt = Math.max(chatAt, Number(m.at) || 0);
                });
            });
        } catch (e) { /* ignore */ }
        return Math.max(aliveAt, awayAt, chatAt);
    }

    function hasLocalChatData() {
        try {
            const raw = localStorage.getItem(CHAT_STORE_KEY);
            return !!(raw && raw.length > 2 && raw !== '{}' && raw !== 'null');
        } catch (e) {
            return false;
        }
    }

    function wipeIfDeviceIdleTooLong() {
        if (!hasLocalChatData()) {
            return false;
        }
        const now = Date.now();
        const last = getLastActivityAt();
        if (last <= 0 || (now - last) <= IDLE_WIPE_MS) {
            return false;
        }
        try {
            localStorage.removeItem(CHAT_STORE_KEY);
            localStorage.removeItem(AVATAR_STORE_KEY);
            localStorage.removeItem(READ_KEY);
            localStorage.removeItem(SEEN_KEY);
        } catch (e) { /* ignore */ }
        avatarStore = {};
        avatarStoreLoaded = true;
        idbClear();
        revokeObjectUrls();
        Object.keys(pendingAttach).forEach(clearPendingAttach);
        Object.keys(pendingReply).forEach((k) => { delete pendingReply[k]; });
        repliedPeers.clear();
        lastRepliedPeerId = '';
        try {
            sessionStorage.removeItem(REPLIED_STORE_KEY);
        } catch (e2) { /* ignore */ }
        unreadMap = {};
        peerSeenAt = {};
        Object.keys(receiveFlashUntil).forEach((k) => { delete receiveFlashUntil[k]; });
        dockSessions = [];
        renderDock();
        notifyUnreadUi();
        touchAlive();
        return true;
    }

    function onIdleVisibilityOrTick() {
        if (document.hidden) {
            markAway();
            return;
        }
        if (wipeIfDeviceIdleTooLong()) {
            loadAvatarStore();
            loadPeerSeen();
            loadRepliedSession();
        } else {
            touchAlive();
        }
    }

    function bindIdleLifecycle() {
        if (window.__deviceChatIdleBound === '1') {
            return;
        }
        window.__deviceChatIdleBound = '1';
        document.addEventListener('visibilitychange', onIdleVisibilityOrTick);
        setInterval(onIdleVisibilityOrTick, 60000);
    }

    function loadPeerSeen() {
        peerSeenAt = readJson(SEEN_KEY, {});
        if (!peerSeenAt || typeof peerSeenAt !== 'object') {
            peerSeenAt = {};
        }
    }

    function savePeerSeen() {
        writeJson(SEEN_KEY, peerSeenAt || {});
    }

    function applyPeerReadReceipt(fromId, readAt) {
        const id = String(fromId || '');
        if (!id) {
            return;
        }
        const t = Number(readAt) || Date.now();
        const prev = Number(peerSeenAt[id]) || 0;
        if (t <= prev) {
            return;
        }
        peerSeenAt[id] = t;
        savePeerSeen();
        refreshStatusLine(id);
    }

    function flashReceived(peerId) {
        const id = String(peerId || '');
        if (!id) {
            return;
        }
        receiveFlashUntil[id] = Date.now() + RECV_FLASH_MS;
        refreshStatusLine(id);
        setTimeout(() => {
            refreshStatusLine(id);
        }, RECV_FLASH_MS + 40);
    }

    function statusTextForPeer(peerId) {
        const id = String(peerId || '');
        const now = Date.now();
        if (receiveFlashUntil[id] && now < receiveFlashUntil[id]) {
            return { text: 'Đã nhận', kind: 'recv' };
        }
        const t = getThread(id);
        const msgs = t.messages || [];
        if (!msgs.length) {
            return null;
        }
        const last = msgs[msgs.length - 1];
        if (!last || last.fromDeviceId !== myDeviceId) {
            return null;
        }
        const seen = Number(peerSeenAt[id]) || 0;
        if (seen >= (Number(last.at) || 0)) {
            return { text: 'Đã xem', kind: 'seen' };
        }
        if ((now - (Number(last.at) || 0)) < SENT_FLASH_MS) {
            return { text: 'Đã gửi', kind: 'sent' };
        }
        return { text: 'Chưa xem', kind: 'unseen' };
    }

    function renderStatusInto(line, peerId) {
        if (!line) {
            return;
        }
        const st = statusTextForPeer(peerId);
        if (!st) {
            line.textContent = '';
            line.setAttribute('hidden', 'hidden');
            line.className = 'chat-status-line';
            return;
        }
        line.removeAttribute('hidden');
        line.textContent = st.text;
        line.className = 'chat-status-line chat-status-line--' + st.kind;
    }

    function refreshStatusLine(peerId) {
        const dock = el('chatDock');
        if (!dock || !peerId) {
            return;
        }
        const line = dock.querySelector('[data-chat-status="' + peerId + '"]');
        renderStatusInto(line, peerId);
        const box = dock.querySelector('[data-chat-messages="' + peerId + '"]');
        scrollChatToBottom(box, peerId, false);
    }

    function ensureStatusTicker() {
        if (statusTicker) {
            return;
        }
        statusTicker = setInterval(() => {
            dockSessions.forEach((sess) => {
                if (sess && !sess.minimized) {
                    refreshStatusLine(sess.deviceId);
                }
            });
        }, 400);
    }

    function isChatNearBottom(box) {
        if (!box) {
            return true;
        }
        return (box.scrollHeight - box.scrollTop - box.clientHeight) <= SCROLL_BOTTOM_SLACK;
    }

    function scrollChatToBottom(box, peerId, force) {
        if (!box) {
            return;
        }
        const id = String(peerId || box.getAttribute('data-chat-messages') || '');
        if (!force && id && scrollPinned[id] === false) {
            return;
        }
        if (id && force) {
            scrollPinned[id] = true;
        }
        const go = () => {
            box.scrollTop = box.scrollHeight;
        };
        go();
        requestAnimationFrame(go);
        setTimeout(go, 0);
    }

    function noteReplied(peerId) {
        const id = String(peerId || '');
        if (!id || id === myDeviceId) {
            return;
        }
        repliedPeers.add(id);
        lastRepliedPeerId = id;
        saveRepliedSession();
        // Cập nhật viền list ngay; tách khỏi renderDock để tránh lỗi UI che mất rerender
        if (window.PresenceBridge && typeof window.PresenceBridge.rerender === 'function') {
            window.PresenceBridge.rerender();
        }
        setTimeout(() => {
            if (window.PresenceBridge && typeof window.PresenceBridge.rerender === 'function') {
                window.PresenceBridge.rerender();
            }
        }, 0);
    }

    function getBorderState(peerId) {
        const id = String(peerId || '');
        if (!id) {
            return '';
        }
        if (id === lastRepliedPeerId) {
            return 'active';
        }
        if (repliedPeers.has(id)) {
            return 'closed';
        }
        return '';
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
            return { messages: [], peerTag: '', peerLabel: '', pinnedIds: [], peerPinnedIds: [], updatedAt: 0 };
        }
        const pinnedIds = Array.isArray(t.pinnedIds)
            ? t.pinnedIds.map((id) => String(id)).filter(Boolean).slice(0, MAX_PINS)
            : [];
        const peerPinnedIds = Array.isArray(t.peerPinnedIds)
            ? t.peerPinnedIds.map((id) => String(id)).filter(Boolean).slice(0, MAX_PINS)
            : [];
        return {
            messages: Array.isArray(t.messages) ? t.messages.slice() : [],
            peerTag: String(t.peerTag || ''),
            peerLabel: String(t.peerLabel || ''),
            pinnedIds: pinnedIds,
            peerPinnedIds: peerPinnedIds,
            updatedAt: Number(t.updatedAt) || 0
        };
    }

    function setThread(peerId, thread) {
        if (!peerId || peerId === myDeviceId) {
            return;
        }
        const store = loadStore();
        const prev = store[peerId];
        let messages = (thread.messages || []).slice(-MSG_LIMIT);
        if (prev && Array.isArray(prev.messages) && Array.isArray(thread.messages)) {
            messages = mergeMessages(prev.messages, thread.messages);
        }
        store[peerId] = {
            messages: messages,
            peerTag: String(thread.peerTag || (prev && prev.peerTag) || ''),
            peerLabel: String(thread.peerLabel || (prev && prev.peerLabel) || ''),
            pinnedIds: (thread.pinnedIds || (prev && prev.pinnedIds) || []).map((id) => String(id)).filter(Boolean).slice(0, MAX_PINS),
            peerPinnedIds: (thread.peerPinnedIds || (prev && prev.peerPinnedIds) || []).map((id) => String(id)).filter(Boolean).slice(0, MAX_PINS),
            updatedAt: Number(thread.updatedAt) || Date.now()
        };
        saveStore(store);
    }

    function waitForAuth(maxMs) {
        const limit = Math.max(3000, Number(maxMs) || 15000);
        return new Promise((resolve, reject) => {
            if (typeof firebase === 'undefined' || !firebase.auth) {
                resolve(null);
                return;
            }
            const auth = firebase.auth();
            if (auth.currentUser) {
                resolve(auth.currentUser);
                return;
            }
            let done = false;
            const finish = (fn, val) => {
                if (done) {
                    return;
                }
                done = true;
                clearTimeout(timer);
                try {
                    unsub();
                } catch (e) { /* ignore */ }
                fn(val);
            };
            const timer = setTimeout(() => {
                finish(reject, new Error('Chưa đăng nhập Firebase (Anonymous Auth).'));
            }, limit);
            const unsub = auth.onAuthStateChanged((user) => {
                if (user) {
                    finish(resolve, user);
                }
            });
            auth.signInAnonymously().catch((err) => {
                finish(reject, err || new Error('Anonymous Auth thất bại.'));
            });
        });
    }

    function ensureStorage() {
        if (storage) {
            return storage;
        }
        if (typeof firebase === 'undefined' || !firebase.storage) {
            return null;
        }
        try {
            storage = firebase.storage();
        } catch (e) {
            storage = null;
            console.warn('[chat] Firebase Storage init failed', e);
        }
        return storage;
    }

    function getStorageDownloadUrl(path) {
        const p = String(path || '').trim();
        const st = ensureStorage();
        if (!p || !st) {
            return Promise.resolve('');
        }
        return waitForAuth(10000).then(() => st.ref(p).getDownloadURL()).catch(() => '');
    }

    function hydrateFileUrls(box) {
        if (!box) {
            return;
        }
        box.querySelectorAll('[data-chat-file-path]:not([data-url-hydrated])').forEach((node) => {
            const path = node.getAttribute('data-chat-file-path');
            if (!path) {
                return;
            }
            node.setAttribute('data-url-hydrated', '1');
            getStorageDownloadUrl(path).then((url) => {
                if (!url) {
                    return;
                }
                if (node.tagName === 'IMG') {
                    node.src = url;
                    const link = node.closest('a.chat-file-link');
                    if (link) {
                        link.href = url;
                    }
                    return;
                }
                if (node.classList && node.classList.contains('chat-file-link')) {
                    node.href = url;
                }
            });
        });
    }

    function setUploadPreview(mid, file) {
        if (!mid || !file || String(file.type || '').indexOf('image/') !== 0) {
            return '';
        }
        try {
            const old = uploadPreviewUrls.get(mid);
            if (old) {
                URL.revokeObjectURL(old);
            }
            const url = URL.createObjectURL(file);
            uploadPreviewUrls.set(mid, url);
            return url;
        } catch (e) {
            return '';
        }
    }

    function clearUploadPreview(mid) {
        const url = uploadPreviewUrls.get(mid);
        if (url) {
            try {
                URL.revokeObjectURL(url);
            } catch (e) { /* ignore */ }
            uploadPreviewUrls.delete(mid);
        }
    }

    function fileToDataUrl(file) {
        return new Promise((resolve, reject) => {
            try {
                const reader = new FileReader();
                reader.onload = () => resolve(String(reader.result || ''));
                reader.onerror = () => reject(reader.error || new Error('Đọc ảnh thất bại.'));
                reader.readAsDataURL(file);
            } catch (e) {
                reject(e);
            }
        });
    }

    function normalizeFileMeta(f) {
        if (!f || typeof f !== 'object') {
            return null;
        }
        const name = String(f.name || 'file').slice(0, 120);
        const size = Number(f.size) || 0;
        const mime = String(f.mime || f.contentType || '').slice(0, 120);
        const path = String(f.path || '').slice(0, 400);
        let url = String(f.url || '');
        if (url.indexOf('data:image/') === 0) {
            url = url.slice(0, 900000);
        } else {
            url = url.slice(0, 2000);
        }
        if (!name && !path && !url) {
            return null;
        }
        return { name: name || 'file', size: size, mime: mime, path: path, url: url };
    }

    function normalizeMessage(m) {
        if (!m || !m.id) {
            return null;
        }
        const recalled = !!m.recalled;
        const replyTo = normalizeReplyTo(m.replyTo);
        if (recalled) {
            return {
                id: String(m.id),
                fromDeviceId: String(m.fromDeviceId || ''),
                toDeviceId: String(m.toDeviceId || ''),
                text: '',
                at: Number(m.at) || 0,
                kind: 'text',
                file: null,
                status: 'ready',
                recalled: true,
                replyTo: replyTo
            };
        }
        const file = normalizeFileMeta(m.file);
        const text = String(m.text || '').slice(0, MAX_TEXT);
        if (!text && !file) {
            return null;
        }
        const kind = file ? 'file' : 'text';
        const status = String(m.status || (file && file.url ? 'ready' : (file ? 'uploading' : 'ready')));
        return {
            id: String(m.id),
            fromDeviceId: String(m.fromDeviceId || ''),
            toDeviceId: String(m.toDeviceId || ''),
            text: text,
            at: Number(m.at) || 0,
            kind: kind,
            file: file,
            status: kind === 'file' ? status : 'ready',
            recalled: false,
            replyTo: replyTo
        };
    }

    function normalizeReplyTo(raw) {
        if (!raw || typeof raw !== 'object' || !raw.id) {
            return null;
        }
        return {
            id: String(raw.id),
            text: String(raw.text || '').slice(0, 160),
            fromDeviceId: String(raw.fromDeviceId || '')
        };
    }

    function mergeMessages(localMsgs, incoming) {
        const map = new Map();
        (localMsgs || []).forEach((raw) => {
            const m = normalizeMessage(raw);
            if (m) {
                map.set(m.id, m);
            }
        });
        (incoming || []).forEach((raw) => {
            const m = normalizeMessage(raw);
            if (!m) {
                return;
            }
            const prev = map.get(m.id);
            if (!prev) {
                map.set(m.id, m);
                return;
            }
            if (prev.recalled || m.recalled) {
                map.set(m.id, normalizeMessage(Object.assign({}, prev, m, {
                    recalled: true,
                    text: '',
                    file: null,
                    kind: 'text'
                })) || m);
                return;
            }
            // Giữ bản đầy đủ hơn (có url/path/status tốt hơn)
            const merged = Object.assign({}, prev, m);
            if (prev.file || m.file) {
                merged.file = Object.assign({}, prev.file || {}, m.file || {});
                if (!merged.file.name) {
                    merged.file.name = 'file';
                }
            }
            if (prev.status === 'ready' && m.status !== 'ready') {
                merged.status = 'ready';
            } else if ((m.file && m.file.url) || m.status === 'ready') {
                merged.status = 'ready';
            }
            map.set(m.id, normalizeMessage(merged) || merged);
        });
        return Array.from(map.values())
            .sort((a, b) => a.at - b.at || a.id.localeCompare(b.id))
            .slice(-MSG_LIMIT);
    }

    function msgId() {
        return 'm' + Date.now().toString(36) + '_' + Math.random().toString(36).slice(2, 10);
    }

    function bubbleLetter(tag) {
        const t = String(tag || 'X').toUpperCase().replace(/[^A-Z0-9]/g, '');
        return t.charAt(0) || 'X';
    }

    function loadAvatarStore() {
        const raw = readJson(AVATAR_STORE_KEY, {});
        avatarStore = raw && typeof raw === 'object' ? raw : {};
        avatarStoreLoaded = true;
    }

    function saveAvatarStore() {
        writeJson(AVATAR_STORE_KEY, avatarStore || {});
    }

    function ensureAvatarStore() {
        if (!avatarStoreLoaded) {
            loadAvatarStore();
        }
    }

    function hashAvatarColor(deviceId) {
        const s = String(deviceId || '');
        let h = 2166136261;
        for (let i = 0; i < s.length; i++) {
            h ^= s.charCodeAt(i);
            h = Math.imul(h, 16777619);
        }
        return AVATAR_COLORS[Math.abs(h >>> 0) % AVATAR_COLORS.length];
    }

    /** Avatar ổn định theo deviceId; chữ = ký tự đầu mã thiết bị (AYU → A). */
    function getAvatar(deviceId, deviceTag) {
        ensureAvatarStore();
        const id = String(deviceId || '').trim();
        const letter = bubbleLetter(deviceTag || id);
        if (!id) {
            return { letter: letter, color: AVATAR_COLORS[0] };
        }
        let entry = avatarStore[id];
        const palette = AVATAR_COLORS;
        const inPalette = (c) => palette.some((x) => x.toLowerCase() === String(c || '').toLowerCase());
        if (!entry || typeof entry !== 'object') {
            entry = { color: hashAvatarColor(id), letter: letter };
            avatarStore[id] = entry;
            saveAvatarStore();
        } else {
            let dirty = false;
            if (!entry.color || !inPalette(entry.color)) {
                entry.color = hashAvatarColor(id);
                dirty = true;
            }
            if (letter && entry.letter !== letter) {
                entry.letter = letter;
                dirty = true;
            }
            if (dirty) {
                avatarStore[id] = entry;
                saveAvatarStore();
            }
        }
        return {
            letter: String(entry.letter || letter).charAt(0) || '?',
            color: String(entry.color || hashAvatarColor(id))
        };
    }

    function avatarHtml(deviceId, deviceTag, opts) {
        const o = opts || {};
        const av = getAvatar(deviceId, deviceTag);
        const cls = 'device-avatar' + (o.className ? ' ' + o.className : '');
        return '<span class="' + cls + '" style="background:' + escapeHtml(av.color) + '"' +
            ' aria-hidden="true">' + escapeHtml(av.letter) + '</span>';
    }

    function notifyUnreadUi() {
        if (window.PresenceBridge && typeof window.PresenceBridge.rerender === 'function') {
            window.PresenceBridge.rerender();
        }
        renderDock();
    }

    function markPeerRead(peerId) {
        if (!peerId) {
            return;
        }
        unreadMap[peerId] = 0;
        const reads = readJson(READ_KEY, {});
        reads[peerId] = Date.now();
        writeJson(READ_KEY, reads);
        sendReadReceipt(peerId);
        if (window.PresenceBridge && typeof window.PresenceBridge.rerender === 'function') {
            window.PresenceBridge.rerender();
        }
    }

    function sendReadReceipt(peerId) {
        if (!peerId || !myDeviceId || peerId === myDeviceId) {
            return;
        }
        const t = getThread(peerId);
        pushMailboxToPeer(peerId, t.messages, { readReceipt: true, force: true });
    }

    function bumpUnread(peerId) {
        if (!peerId || peerId === myDeviceId) {
            return;
        }
        const sess = dockSessions.find((s) => s.deviceId === peerId);
        if (sess && !sess.minimized) {
            markPeerRead(peerId);
            return;
        }
        unreadMap[peerId] = (Number(unreadMap[peerId]) || 0) + 1;
        notifyUnreadUi();
    }

    function isChatOpen(peerId) {
        if (peerId) {
            return dockSessions.some((s) => s.deviceId === peerId);
        }
        return dockSessions.length > 0;
    }

    function findSession(peerId) {
        return dockSessions.find((s) => s.deviceId === peerId) || null;
    }

    function safeFileName(name) {
        return String(name || 'file')
            .replace(/[^\w.\-()+ ]+/g, '_')
            .replace(/\s+/g, '_')
            .slice(0, 80) || 'file';
    }

    function buildStoragePath(peerId, mid, fileName) {
        return STORAGE_ROOT + '/' + myDeviceId + '/' + peerId + '/' + mid + '_' + safeFileName(fileName);
    }

    function serializeMsgForMailbox(m) {
        const n = normalizeMessage(m);
        if (!n) {
            return null;
        }
        if (n.recalled) {
            const outR = {
                id: n.id,
                fromDeviceId: n.fromDeviceId,
                toDeviceId: n.toDeviceId,
                text: '',
                at: n.at,
                kind: 'text',
                recalled: true
            };
            if (n.replyTo) {
                outR.replyTo = n.replyTo;
            }
            return outR;
        }
        if (n.kind === 'file' && n.status !== 'ready') {
            return null;
        }
        if (n.kind === 'file' && !(n.file && (n.file.url || n.file.path))) {
            return null;
        }
        const out = {
            id: n.id,
            fromDeviceId: n.fromDeviceId,
            toDeviceId: n.toDeviceId,
            text: n.text || '',
            at: n.at,
            kind: n.kind,
            recalled: false
        };
        if (n.file) {
            out.file = {
                name: n.file.name,
                size: n.file.size,
                mime: n.file.mime,
                path: n.file.path,
                url: n.file.url
            };
        }
        if (n.replyTo) {
            out.replyTo = n.replyTo;
        }
        return out;
    }

    function fileBodyHtml(m) {
        const f = m.file || {};
        let status = m.status || 'ready';
        const previewUrl = uploadPreviewUrls.get(m.id) || '';
        const pathAttr = f.path
            ? ' data-chat-file-path="' + escapeHtml(f.path) + '"'
            : '';
        if (status === 'uploading' && m.at && (Date.now() - Number(m.at)) > UPLOAD_TIMEOUT_MS) {
            status = 'error';
        }
        const href = f.url ? escapeHtml(f.url) : (previewUrl ? escapeHtml(previewUrl) : '');
        let body = '<div class="chat-file-card">';
        if (status === 'uploading') {
            body += '<div class="chat-file-status">Đang tải lên…</div>';
        } else if (status === 'error') {
            body += '<div class="chat-file-status chat-file-status--err">Gửi file thất bại</div>';
        }
        const isImg = String(f.mime || '').indexOf('image/') === 0 && (href || f.path);
        if (isImg) {
            const src = href || 'data:image/gif;base64,R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7';
            body += '<button type="button" class="chat-file-thumb-btn" data-chat-image-preview="1" ' +
                'title="Xem ảnh" aria-label="Xem ảnh">' +
                '<img class="chat-file-thumb" src="' + src + '"' + pathAttr +
                ' alt="' + escapeHtml(f.name || 'image') + '" />' +
                '</button>';
            body += '<div class="chat-file-name">🖼 ' + escapeHtml(f.name || 'image') + '</div>';
        } else if (f.url) {
            body += '<a class="chat-file-link" href="' + escapeHtml(f.url) +
                '" target="_blank" rel="noopener noreferrer" download="' +
                escapeHtml(f.name || 'file') + '">📎 ' + escapeHtml(f.name || 'file') + '</a>';
        } else if (!isImg) {
            body += '<div class="chat-file-link"' + pathAttr + '>📎 ' + escapeHtml(f.name || 'file') + '</div>';
        }
        body += '<div class="chat-file-meta">' + escapeHtml(formatFileSize(f.size)) +
            (f.mime ? ' · ' + escapeHtml(f.mime.split(';')[0]) : '') + '</div>';
        if (m.text) {
            body += '<div class="chat-bubble-text">' + escapeHtml(m.text) + '</div>';
        }
        body += '</div>';
        return body;
    }

    function peerTagForRender(peerId) {
        const sess = findSession(peerId);
        if (sess && sess.deviceTag) {
            return String(sess.deviceTag);
        }
        const t = getThread(peerId);
        if (t && t.peerTag) {
            return String(t.peerTag);
        }
        return peerId ? makeDeviceTag(peerId) : '';
    }

    function pinBannerHtml(peerId) {
        const t = getThread(peerId);
        const myN = (t.pinnedIds || []).length;
        const peerN = (t.peerPinnedIds || []).length;
        if (!myN && !peerN) {
            return '<div class="chat-pin-banner" data-chat-pin-banner="' + escapeHtml(peerId) + '"></div>';
        }
        let label = '';
        if (myN && peerN) {
            label = 'Bạn ghim ' + myN + ' · đối phương ghim ' + peerN + '.';
        } else if (myN === 1) {
            label = 'Bạn đã ghim một tin nhắn.';
        } else if (myN > 1) {
            label = 'Bạn đã ghim ' + myN + ' tin nhắn.';
        } else if (peerN === 1) {
            label = 'Đối phương đã ghim một tin nhắn.';
        } else {
            label = 'Đối phương đã ghim ' + peerN + ' tin nhắn.';
        }
        return '<div class="chat-pin-banner is-on" data-chat-pin-banner="' + escapeHtml(peerId) + '">' +
            '<span class="chat-pin-banner-text">' + escapeHtml(label) + '</span>' +
            '<button type="button" class="chat-pin-banner-link" data-chat-pins-open="' +
            escapeHtml(peerId) + '">Xem tất cả</button></div>';
    }

    function closeAllMsgMenus(exceptMenu) {
        const dock = el('chatDock');
        if (!dock) {
            return;
        }
        dock.querySelectorAll('.chat-msg-menu.is-open').forEach((menu) => {
            if (exceptMenu && menu === exceptMenu) {
                return;
            }
            menu.classList.remove('is-open');
            const row = menu.closest('.chat-bubble-row');
            const btn = row && row.querySelector('.chat-msg-more');
            if (btn) {
                btn.classList.remove('is-open');
            }
        });
    }

    function replyQuoteHtml(m, peerId, peerTag) {
        if (!m || !m.replyTo || !m.replyTo.id) {
            return '';
        }
        const t = getThread(peerId);
        const orig = (t.messages || []).find((x) => x.id === m.replyTo.id);
        let who = m.replyTo.fromDeviceId === myDeviceId ? 'You' : (peerTag || 'Them');
        let preview = String(m.replyTo.text || '');
        if (orig) {
            who = orig.fromDeviceId === myDeviceId ? 'You' : (peerTag || 'Them');
            preview = orig.recalled ? RECALL_LABEL : previewPinnedText(orig);
        }
        if (!preview) {
            preview = 'Tin nhắn';
        }
        return '<div class="chat-reply-quote">' +
            '<div class="chat-reply-quote-who">' + escapeHtml(who) + '</div>' +
            '<div class="chat-reply-quote-text">' + escapeHtml(preview) + '</div></div>';
    }

    function replyDraftHtml(peerId) {
        const r = pendingReply[peerId];
        if (!r) {
            return '<div class="chat-reply-draft" data-chat-reply-draft="' + escapeHtml(peerId) + '"></div>';
        }
        const who = r.fromDeviceId === myDeviceId ? 'You' : peerTagForRender(peerId);
        return '<div class="chat-reply-draft is-on" data-chat-reply-draft="' + escapeHtml(peerId) + '">' +
            '<div class="chat-reply-draft-body"><strong>Đang trả lời ' + escapeHtml(who) + '</strong>' +
            '<span>' + escapeHtml(r.text || 'Tin nhắn') + '</span></div>' +
            '<button type="button" class="chat-reply-draft-cancel" data-chat-reply-cancel="' +
            escapeHtml(peerId) + '" title="Hủy trả lời" aria-label="Hủy trả lời">×</button></div>';
    }

    function setReplyTarget(peerId, msgId) {
        const t = getThread(peerId);
        const m = (t.messages || []).find((x) => x.id === msgId);
        if (!m || m.recalled || !findSession(peerId)) {
            return;
        }
        pendingReply[peerId] = {
            id: m.id,
            text: previewPinnedText(m).slice(0, 160),
            fromDeviceId: m.fromDeviceId
        };
        renderDock();
        focusInput(peerId);
    }

    function clearReplyTarget(peerId) {
        delete pendingReply[peerId];
    }

    function takeReplyTarget(peerId) {
        const r = pendingReply[peerId] || null;
        delete pendingReply[peerId];
        return r;
    }

    function repairStuckUploads(peerId) {
        const t = getThread(peerId);
        const now = Date.now();
        let dirty = false;
        const next = (t.messages || []).map((m) => {
            if (m.status === 'uploading' && Number(m.at) && (now - Number(m.at)) > UPLOAD_TIMEOUT_MS) {
                dirty = true;
                return Object.assign({}, m, { status: 'error' });
            }
            return m;
        });
        if (dirty) {
            t.messages = next;
            t.updatedAt = now;
            setThread(peerId, t);
        }
        return dirty ? next : t.messages;
    }

    function renderMessagesInto(box, rows, peerId) {
        if (!box) {
            return;
        }
        if (peerId) {
            rows = repairStuckUploads(peerId);
        }
        if (!rows.length) {
            box.innerHTML = '<div class="chat-empty">Chưa có tin — lưu trên máy; tắt máy ~10 phút sẽ mất.</div>';
            return;
        }
        const peerTag = peerTagForRender(peerId);
        const shouldStick = !peerId || scrollPinned[peerId] !== false;
        const t = getThread(peerId);
        const pinnedSet = new Set(t.pinnedIds || []);
        let html = '';
        rows.forEach((m) => {
            const mine = m.fromDeviceId === myDeviceId;
            const who = mine ? 'You' : (peerTag || 'Them');
            const recalled = !!m.recalled;
            let inner;
            if (recalled) {
                inner = '<div class="chat-bubble-text chat-bubble-text--recalled">' +
                    escapeHtml(RECALL_LABEL) + '</div>';
            } else if (m.kind === 'file' || m.file) {
                inner = fileBodyHtml(m);
            } else {
                inner = '<div class="chat-bubble-text">' + escapeHtml(m.text) + '</div>';
            }
            const canRecall = mine && !recalled;
            const canPin = !recalled;
            const canReply = !recalled;
            const pinned = pinnedSet.has(m.id);
            const tools = (canReply || canRecall || canPin)
                ? ('<div class="chat-msg-tools">' +
                    (canReply
                        ? '<button type="button" class="chat-msg-reply" data-chat-msg-reply="1" ' +
                        'title="Trả lời" aria-label="Trả lời tin nhắn">↩</button>'
                        : '') +
                    ((canRecall || canPin)
                        ? ('<button type="button" class="chat-msg-more" data-chat-msg-more="1" ' +
                        'title="Tuỳ chọn" aria-label="Tuỳ chọn tin nhắn">⋮</button>' +
                        '<div class="chat-msg-menu" data-chat-msg-menu="1">' +
                        (canRecall
                            ? '<button type="button" class="chat-msg-menu-item chat-msg-menu-item--danger" ' +
                            'data-chat-recall="1">Thu hồi</button>'
                            : '') +
                        (canPin
                            ? '<button type="button" class="chat-msg-menu-item" data-chat-pin-toggle="1">' +
                            (pinned ? 'Bỏ ghim' : 'Ghim') + '</button>'
                            : '') +
                        '</div>')
                        : '') +
                    '</div>')
                : '';
            html += '<div class="chat-bubble-row' + (mine ? ' chat-bubble-row--mine' : '') +
                (recalled ? ' chat-bubble-row--recalled' : '') +
                '" data-msg-id="' + escapeHtml(m.id) + '" data-msg-peer="' + escapeHtml(peerId) + '">' +
                tools +
                '<div class="chat-bubble' + (mine ? ' chat-bubble--mine' : '') + '">' +
                replyQuoteHtml(m, peerId, peerTag) +
                inner +
                '<div class="chat-bubble-time">' + escapeHtml(formatMsgTime(m.at)) +
                ' · ' + escapeHtml(who) + '</div>' +
                '</div></div>';
        });
        html += '<div class="chat-status-line" data-chat-status="' + escapeHtml(peerId || '') +
            '" hidden></div>';
        box.innerHTML = html;
        hydrateFileUrls(box);
        renderStatusInto(box.querySelector('[data-chat-status]'), peerId);
        ensureStatusTicker();
        if (shouldStick) {
            scrollChatToBottom(box, peerId, true);
        }
    }

    function refreshThreadUi(peerId) {
        const dock = el('chatDock');
        if (!dock || !peerId) {
            return;
        }
        const bannerHost = dock.querySelector('[data-chat-pin-banner="' + peerId + '"]');
        if (bannerHost) {
            const wrap = document.createElement('div');
            wrap.innerHTML = pinBannerHtml(peerId);
            const next = wrap.firstElementChild;
            if (next) {
                bannerHost.replaceWith(next);
            }
        }
        const box = dock.querySelector('[data-chat-messages="' + peerId + '"]');
        if (box) {
            renderMessagesInto(box, getThread(peerId).messages, peerId);
        }
        const modal = el('chatPinModal');
        if (modal && !modal.hidden && modal.getAttribute('data-peer') === peerId) {
            renderPinnedModal(peerId);
        }
    }

    function recallMessage(peerId, msgId) {
        const t = getThread(peerId);
        let found = null;
        t.messages = t.messages.map((m) => {
            if (m.id !== msgId) {
                return m;
            }
            if (m.fromDeviceId !== myDeviceId || m.recalled) {
                return m;
            }
            found = m;
            return normalizeMessage({
                id: m.id,
                fromDeviceId: m.fromDeviceId,
                toDeviceId: m.toDeviceId,
                at: m.at,
                recalled: true
            }) || m;
        });
        if (!found) {
            return;
        }
        t.pinnedIds = (t.pinnedIds || []).filter((id) => id !== msgId);
        t.updatedAt = Date.now();
        setThread(peerId, t);
        refreshThreadUi(peerId);
        pushMailboxToPeer(peerId, t.messages, { force: true });
    }

    function togglePinMessage(peerId, msgId) {
        const t = getThread(peerId);
        const msg = (t.messages || []).find((m) => m.id === msgId);
        if (!msg || msg.recalled) {
            return;
        }
        const pins = (t.pinnedIds || []).slice();
        const idx = pins.indexOf(msgId);
        if (idx >= 0) {
            pins.splice(idx, 1);
        } else {
            if (pins.length >= MAX_PINS) {
                window.alert('Tối đa ' + MAX_PINS + ' tin ghim.');
                return;
            }
            pins.push(msgId);
        }
        t.pinnedIds = pins;
        t.updatedAt = Date.now();
        setThread(peerId, t);
        refreshThreadUi(peerId);
        pushMailboxToPeer(peerId, t.messages, { force: true, pins: true });
    }

    let imagePreviewEscHandler = null;

    function ensureImagePreviewModal() {
        let modal = el('chatImagePreviewModal');
        if (modal) {
            return modal;
        }
        modal = document.createElement('div');
        modal.id = 'chatImagePreviewModal';
        modal.className = 'chat-image-preview-modal';
        modal.hidden = true;
        modal.innerHTML = '<div class="chat-image-preview-card" role="dialog" aria-modal="true" ' +
            'aria-label="Xem ảnh">' +
            '<button type="button" class="chat-image-preview-close" data-chat-image-preview-close="1" ' +
            'aria-label="Đóng">×</button>' +
            '<img class="chat-image-preview-img" data-chat-image-preview-img alt="" />' +
            '<div class="chat-image-preview-caption" data-chat-image-preview-caption></div></div>';
        document.body.appendChild(modal);
        modal.addEventListener('click', (e) => {
            if (e.target === modal || (e.target && e.target.getAttribute &&
                e.target.getAttribute('data-chat-image-preview-close') === '1')) {
                closeImagePreview();
            }
        });
        return modal;
    }

    function closeImagePreview() {
        const modal = el('chatImagePreviewModal');
        if (!modal) {
            return;
        }
        modal.hidden = true;
        const img = modal.querySelector('[data-chat-image-preview-img]');
        if (img) {
            img.removeAttribute('src');
            img.alt = '';
        }
        const cap = modal.querySelector('[data-chat-image-preview-caption]');
        if (cap) {
            cap.textContent = '';
        }
        if (imagePreviewEscHandler) {
            document.removeEventListener('keydown', imagePreviewEscHandler);
            imagePreviewEscHandler = null;
        }
    }

    function openImagePreview(src, alt) {
        if (!src) {
            return;
        }
        const modal = ensureImagePreviewModal();
        const img = modal.querySelector('[data-chat-image-preview-img]');
        const cap = modal.querySelector('[data-chat-image-preview-caption]');
        if (img) {
            img.src = src;
            img.alt = alt || 'image';
        }
        if (cap) {
            cap.textContent = alt || '';
            cap.hidden = !alt;
        }
        modal.hidden = false;
        if (!imagePreviewEscHandler) {
            imagePreviewEscHandler = (e) => {
                if (e.key === 'Escape') {
                    closeImagePreview();
                }
            };
            document.addEventListener('keydown', imagePreviewEscHandler);
        }
    }

    function ensurePinModal() {
        let modal = el('chatPinModal');
        if (modal) {
            return modal;
        }
        modal = document.createElement('div');
        modal.id = 'chatPinModal';
        modal.className = 'chat-pin-modal';
        modal.hidden = true;
        modal.innerHTML = '<div class="chat-pin-modal-card" role="dialog" aria-modal="true" ' +
            'aria-label="Tin nhắn đã ghim">' +
            '<div class="chat-pin-modal-head"><span data-pin-modal-title>Tin nhắn đã ghim</span>' +
            '<button type="button" class="chat-pin-modal-close" data-chat-pin-modal-close="1" ' +
            'aria-label="Đóng">×</button></div>' +
            '<div class="chat-pin-modal-list" data-pin-modal-list></div></div>';
        document.body.appendChild(modal);
        modal.addEventListener('click', (e) => {
            if (e.target === modal || (e.target && e.target.getAttribute &&
                e.target.getAttribute('data-chat-pin-modal-close') === '1')) {
                closePinnedModal();
                return;
            }
            const unpin = e.target && e.target.closest
                ? e.target.closest('[data-chat-pin-unpin]')
                : null;
            if (unpin) {
                const peerId = modal.getAttribute('data-peer');
                const msgId = unpin.getAttribute('data-chat-pin-unpin');
                togglePinMessage(peerId, msgId);
            }
        });
        return modal;
    }

    function previewPinnedText(m) {
        if (!m || m.recalled) {
            return RECALL_LABEL;
        }
        if (m.file) {
            return '📎 ' + (m.file.name || 'file') + (m.text ? (' — ' + m.text) : '');
        }
        return String(m.text || '');
    }

    function renderPinnedModal(peerId) {
        const modal = ensurePinModal();
        const t = getThread(peerId);
        const list = modal.querySelector('[data-pin-modal-list]');
        const title = modal.querySelector('[data-pin-modal-title]');
        const myPins = new Set(t.pinnedIds || []);
        const peerPins = new Set(t.peerPinnedIds || []);
        const allIds = [];
        const seen = new Set();
        (t.pinnedIds || []).concat(t.peerPinnedIds || []).forEach((id) => {
            if (!seen.has(id)) {
                seen.add(id);
                allIds.push(id);
            }
        });
        if (title) {
            title.textContent = allIds.length
                ? ('Tin nhắn đã ghim (' + allIds.length + ')')
                : 'Tin nhắn đã ghim';
        }
        if (!list) {
            return;
        }
        if (!allIds.length) {
            list.innerHTML = '<div class="chat-empty">Chưa ghim tin nào.</div>';
            return;
        }
        const byId = new Map((t.messages || []).map((m) => [m.id, m]));
        const ordered = allIds
            .map((id) => byId.get(id))
            .filter(Boolean)
            .sort((a, b) => a.at - b.at || a.id.localeCompare(b.id));
        const peerTag = peerTagForRender(peerId);
        let html = '';
        ordered.forEach((m) => {
            const mine = m.fromDeviceId === myDeviceId;
            const who = mine ? 'You' : peerTag;
            const iPinned = myPins.has(m.id);
            const theyPinned = peerPins.has(m.id);
            let by = '';
            if (iPinned && theyPinned) {
                by = 'Bạn & đối phương ghim';
            } else if (iPinned) {
                by = 'Bạn ghim';
            } else {
                by = 'Đối phương ghim';
            }
            const unpinBtn = iPinned
                ? ('<button type="button" class="chat-pin-unpin" data-chat-pin-unpin="' +
                    escapeHtml(m.id) + '">Bỏ ghim</button>')
                : '';
            html += '<div class="chat-pin-modal-row' + (mine ? ' chat-pin-modal-row--mine' : '') + '">' +
                (mine
                    ? ('<div class="chat-pin-modal-actions"><span class="chat-pin-by">' +
                        escapeHtml(by) + '</span>' + unpinBtn + '</div>')
                    : '') +
                '<div class="chat-pin-modal-bubble">' +
                '<div class="chat-pin-modal-item-text">' + escapeHtml(previewPinnedText(m)) + '</div>' +
                '<div class="chat-pin-modal-item-meta">' +
                escapeHtml(formatMsgTime(m.at)) + ' · ' + escapeHtml(who) + '</div></div>' +
                (!mine
                    ? ('<div class="chat-pin-modal-actions"><span class="chat-pin-by">' +
                        escapeHtml(by) + '</span>' + unpinBtn + '</div>')
                    : '') +
                '</div>';
        });
        list.innerHTML = html || '<div class="chat-empty">Chưa ghim tin nào.</div>';
    }

    function openPinnedModal(peerId) {
        const modal = ensurePinModal();
        modal.setAttribute('data-peer', peerId);
        modal.hidden = false;
        renderPinnedModal(peerId);
    }

    function closePinnedModal() {
        const modal = el('chatPinModal');
        if (modal) {
            modal.hidden = true;
            modal.removeAttribute('data-peer');
        }
    }

    function sessionTitle(sess) {
        const tag = String(sess.deviceTag || 'Chat');
        const place = formatChatPlace(sess.place);
        return place ? (tag + ' (' + place + ')') : tag;
    }

    function sessionSubtitle(sess) {
        return sess.online ? 'đang online' : 'offline — tin giữ trên máy bạn';
    }

    function autosizeInput(input) {
        if (!input) {
            return;
        }
        input.style.height = 'auto';
        const next = Math.max(INPUT_MIN_H, Math.min(INPUT_MAX_H, input.scrollHeight));
        input.style.height = next + 'px';
    }

    function ensureDock() {
        let dock = el('chatDock');
        if (!dock) {
            dock = document.createElement('div');
            dock.id = 'chatDock';
            dock.className = 'chat-dock';
            document.body.appendChild(dock);
        }
        return dock;
    }

    function captureDrafts() {
        const drafts = {};
        const dock = el('chatDock');
        if (!dock) {
            return drafts;
        }
        dock.querySelectorAll('[data-chat-input]').forEach((input) => {
            const id = input.getAttribute('data-chat-input');
            if (id) {
                drafts[id] = input.value;
            }
        });
        return drafts;
    }

    function pendingList(peerId) {
        if (!pendingAttach[peerId]) {
            pendingAttach[peerId] = [];
        }
        return pendingAttach[peerId];
    }

    function clearPendingAttach(peerId) {
        const list = pendingAttach[peerId] || [];
        list.forEach((item) => {
            if (item && item.previewUrl) {
                try {
                    URL.revokeObjectURL(item.previewUrl);
                } catch (e) { /* ignore */ }
            }
        });
        delete pendingAttach[peerId];
    }

    function removePendingAttachItem(peerId, itemId) {
        const list = pendingAttach[peerId] || [];
        const next = [];
        list.forEach((item) => {
            if (item.id === itemId) {
                if (item.previewUrl) {
                    try {
                        URL.revokeObjectURL(item.previewUrl);
                    } catch (e) { /* ignore */ }
                }
                return;
            }
            next.push(item);
        });
        if (next.length) {
            pendingAttach[peerId] = next;
        } else {
            delete pendingAttach[peerId];
        }
    }

    function stageFile(peerId, file, opts) {
        const silent = !!(opts && opts.silent);
        if (!peerId || !file || !findSession(peerId)) {
            return false;
        }
        if (file.size > MAX_FILE_BYTES) {
            if (!silent) {
                window.alert('File tối đa ' + formatFileSize(MAX_FILE_BYTES) + '.');
            }
            return false;
        }
        const list = pendingList(peerId);
        if (list.length >= MAX_PENDING_FILES) {
            if (!silent) {
                window.alert('Tối đa ' + MAX_PENDING_FILES + ' file mỗi lần gửi.');
            }
            return false;
        }
        const mime = String(file.type || 'application/octet-stream');
        let name = String(file.name || '').trim();
        if (!name) {
            const ext = mime.indexOf('image/') === 0
                ? (mime.split('/')[1] || 'png')
                : 'bin';
            name = 'clipboard-' + Date.now() + '.' + ext;
        }
        const previewUrl = mime.indexOf('image/') === 0
            ? URL.createObjectURL(file)
            : '';
        list.push({
            id: 'p' + Date.now().toString(36) + '_' + Math.random().toString(36).slice(2, 8),
            file: file,
            previewUrl: previewUrl,
            name: name.slice(0, 120),
            size: Number(file.size) || 0,
            mime: mime.slice(0, 120)
        });
        if (!silent) {
            renderDock();
            focusInput(peerId);
        }
        return true;
    }

    function stageFiles(peerId, files) {
        const arr = Array.prototype.slice.call(files || []);
        let ok = 0;
        let blockedSize = false;
        let blockedMax = false;
        arr.forEach((f) => {
            if (!f) {
                return;
            }
            if (f.size > MAX_FILE_BYTES) {
                blockedSize = true;
                return;
            }
            if ((pendingAttach[peerId] || []).length >= MAX_PENDING_FILES) {
                blockedMax = true;
                return;
            }
            if (stageFile(peerId, f, { silent: true })) {
                ok += 1;
            }
        });
        if (blockedSize) {
            window.alert('Một số file vượt quá ' + formatFileSize(MAX_FILE_BYTES) + '.');
        } else if (blockedMax) {
            window.alert('Tối đa ' + MAX_PENDING_FILES + ' file mỗi lần gửi.');
        }
        if (ok > 0) {
            renderDock();
            focusInput(peerId);
        }
        return ok > 0;
    }

    function draftAttachHtml(peerId) {
        const list = pendingAttach[peerId] || [];
        if (!list.length) {
            return '<div class="chat-draft-list" data-chat-draft="' + escapeHtml(peerId) + '"></div>';
        }
        let html = '<div class="chat-draft-list is-on" data-chat-draft="' + escapeHtml(peerId) + '">';
        list.forEach((p) => {
            const thumb = p.previewUrl
                ? '<img class="chat-draft-thumb" src="' + escapeHtml(p.previewUrl) + '" alt="" />'
                : '<span class="chat-draft-icon">📎</span>';
            html += '<div class="chat-draft-attach">' +
                thumb +
                '<div class="chat-draft-meta">' + escapeHtml(p.name) +
                '<small>' + escapeHtml(formatFileSize(p.size)) + ' · chờ Gửi</small></div>' +
                '<button type="button" class="chat-draft-remove" data-chat-draft-remove="' +
                escapeHtml(peerId) + '" data-chat-draft-id="' + escapeHtml(p.id) +
                '" title="Bỏ đính kèm" aria-label="Bỏ đính kèm">×</button>' +
                '</div>';
        });
        html += '</div>';
        return html;
    }

    function renderDock() {
        const dock = ensureDock();
        const drafts = captureDrafts();
        const ordered = dockSessions.slice().sort((a, b) => b.openedAt - a.openedAt);
        let html = '';
        ordered.forEach((sess) => {
            const unread = Number(unreadMap[sess.deviceId]) || 0;
            if (sess.minimized) {
                const av = getAvatar(sess.deviceId, sess.deviceTag);
                html += '<button type="button" class="chat-bubble-btn' +
                    (unread > 0 ? ' chat-bubble-btn--unread' : '') + '"' +
                    ' data-chat-expand="' + escapeHtml(sess.deviceId) + '"' +
                    ' title="' + escapeHtml(sess.deviceTag || 'Chat') + '"' +
                    ' style="background:' + escapeHtml(av.color) + '">' +
                    escapeHtml(av.letter) +
                    (unread > 0
                        ? '<span class="chat-bubble-unread">' + (unread > 99 ? '99+' : String(unread)) + '</span>'
                        : '') +
                    '</button>';
                return;
            }
            html += '<div class="chat-window" data-chat-window="' + escapeHtml(sess.deviceId) + '">' +
                '<div class="chat-panel-header" data-chat-header="' + escapeHtml(sess.deviceId) + '">' +
                avatarHtml(sess.deviceId, sess.deviceTag, { className: 'device-avatar--chat-header' }) +
                '<div class="chat-panel-heading">' +
                '<div class="chat-title" title="' + escapeHtml(sessionTitle(sess)) + '">' +
                escapeHtml(sessionTitle(sess)) + '</div>' +
                '<div class="chat-subtitle">' + escapeHtml(sessionSubtitle(sess)) + '</div>' +
                '</div>' +
                '<div class="chat-header-actions">' +
                '<button type="button" class="chat-min-btn" data-chat-min="' + escapeHtml(sess.deviceId) +
                '" title="Thu nhỏ" aria-label="Thu nhỏ">−</button>' +
                '<button type="button" class="chat-close-btn" data-chat-close="' + escapeHtml(sess.deviceId) +
                '" title="Đóng" aria-label="Đóng">×</button>' +
                '</div></div>' +
                pinBannerHtml(sess.deviceId) +
                '<div class="chat-messages" data-chat-messages="' + escapeHtml(sess.deviceId) + '"></div>' +
                '<form class="chat-form" data-chat-form="' + escapeHtml(sess.deviceId) + '" autocomplete="off">' +
                replyDraftHtml(sess.deviceId) +
                draftAttachHtml(sess.deviceId) +
                '<div class="chat-composer-row">' +
                '<button type="button" class="chat-attach-btn" data-chat-attach="' + escapeHtml(sess.deviceId) +
                '" title="Đính kèm file" aria-label="Đính kèm file">+' +
                (((pendingAttach[sess.deviceId] || []).length)
                    ? '<span class="chat-attach-badge">' +
                    String((pendingAttach[sess.deviceId] || []).length > 99
                        ? '99+'
                        : (pendingAttach[sess.deviceId] || []).length) +
                    '</span>'
                    : '') +
                '</button>' +
                '<input class="chat-file-input" type="file" accept="*/*" multiple data-chat-file-input="' +
                escapeHtml(sess.deviceId) + '" />' +
                '<textarea class="chat-input" rows="1" maxlength="500" placeholder="Gõ tin nhắn hoặc dán ảnh…" ' +
                'aria-label="Nội dung tin nhắn" data-chat-input="' + escapeHtml(sess.deviceId) + '"></textarea>' +
                '<button type="submit" class="chat-send-btn">Gửi</button>' +
                '</div></form></div>';
        });
        dock.innerHTML = html;
        ordered.forEach((sess) => {
            if (sess.minimized) {
                return;
            }
            const box = dock.querySelector('[data-chat-messages="' + sess.deviceId + '"]');
            renderMessagesInto(box, getThread(sess.deviceId).messages, sess.deviceId);
            const input = dock.querySelector('[data-chat-input="' + sess.deviceId + '"]');
            if (input) {
                if (Object.prototype.hasOwnProperty.call(drafts, sess.deviceId)) {
                    input.value = drafts[sess.deviceId];
                }
                autosizeInput(input);
            }
        });
    }

    function focusInput(peerId) {
        const dock = el('chatDock');
        if (!dock) {
            return;
        }
        const input = dock.querySelector('[data-chat-input="' + peerId + '"]');
        if (input) {
            setTimeout(() => {
                try {
                    input.focus();
                } catch (e) { /* ignore */ }
            }, 0);
        }
    }

    function refreshOpenThread() {
        const dock = el('chatDock');
        if (!dock) {
            return;
        }
        dockSessions.forEach((sess) => {
            if (sess.minimized) {
                return;
            }
            const box = dock.querySelector('[data-chat-messages="' + sess.deviceId + '"]');
            if (box) {
                renderMessagesInto(box, getThread(sess.deviceId).messages, sess.deviceId);
            }
        });
    }

    function updateThreadMessage(peerId, mid, patch) {
        const t = getThread(peerId);
        let found = false;
        t.messages = t.messages.map((m) => {
            if (m.id !== mid) {
                return m;
            }
            found = true;
            const next = Object.assign({}, m, patch || {});
            if (patch && patch.file) {
                next.file = Object.assign({}, m.file || {}, patch.file);
            }
            return normalizeMessage(next) || next;
        });
        if (!found) {
            return;
        }
        t.updatedAt = Date.now();
        setThread(peerId, t);
        const dock = el('chatDock');
        const box = dock && dock.querySelector('[data-chat-messages="' + peerId + '"]');
        renderMessagesInto(box, t.messages, peerId);
    }

    function openChat(opts) {
        const peerId = String((opts && opts.deviceId) || '').trim();
        if (!peerId || !myDeviceId || peerId === myDeviceId) {
            console.warn('[chat] blocked: same device cannot chat with itself');
            return;
        }
        const silent = !!(opts && opts.silent);
        const tag = String((opts && opts.deviceTag) || makeDeviceTag(peerId));
        const label = String((opts && opts.label) || '');
        const place = String((opts && opts.place) || '');
        const online = opts && typeof opts.online === 'boolean'
            ? !!opts.online
            : onlineDeviceSet.has(peerId);

        let sess = findSession(peerId);
        if (sess) {
            sess.minimized = false;
            sess.deviceTag = tag || sess.deviceTag;
            sess.label = label || sess.label;
            sess.place = place || sess.place;
            sess.online = online;
            sess.openedAt = Date.now();
        } else {
            while (dockSessions.length >= MAX_DOCK) {
                dockSessions.sort((a, b) => a.openedAt - b.openedAt);
                dockSessions.shift();
            }
            sess = {
                deviceId: peerId,
                deviceTag: tag,
                label: label,
                place: place,
                online: online,
                minimized: false,
                openedAt: Date.now()
            };
            dockSessions.push(sess);
        }

        const t = getThread(peerId);
        t.peerTag = sess.deviceTag || t.peerTag;
        t.peerLabel = sess.label || t.peerLabel;
        setThread(peerId, t);
        markPeerRead(peerId);
        scrollPinned[peerId] = true;
        renderDock();
        focusInput(peerId);
        if (sess.online && t.messages.length) {
            pushMailboxToPeer(peerId, t.messages);
        }
        if (!silent) {
            broadcast({ type: 'open', peer: sess });
        }
    }

    function minimizeChat(peerId, silent) {
        const sess = findSession(peerId);
        if (!sess) {
            return;
        }
        sess.minimized = true;
        renderDock();
        if (!silent) {
            broadcast({ type: 'min', peerId: peerId });
        }
    }

    function expandChat(peerId, silent) {
        const sess = findSession(peerId);
        if (!sess) {
            return;
        }
        sess.minimized = false;
        sess.openedAt = Date.now();
        scrollPinned[peerId] = true;
        markPeerRead(peerId);
        renderDock();
        focusInput(peerId);
        if (!silent) {
            broadcast({ type: 'expand', peerId: peerId });
        }
    }

    function closeChat(opts) {
        const silent = !!(opts && opts.silent);
        const peerId = opts && opts.peerId ? String(opts.peerId) : '';
        if (peerId) {
            clearPendingAttach(peerId);
            clearReplyTarget(peerId);
            dockSessions = dockSessions.filter((s) => s.deviceId !== peerId);
        } else {
            Object.keys(pendingAttach).forEach(clearPendingAttach);
            Object.keys(pendingReply).forEach((k) => { delete pendingReply[k]; });
            dockSessions = [];
        }
        renderDock();
        if (!silent) {
            broadcast({ type: 'close', peerId: peerId || null });
        }
    }

    function submitComposer(peerId) {
        const dock = el('chatDock');
        const input = dock && dock.querySelector('[data-chat-input="' + peerId + '"]');
        const text = input ? input.value : '';
        const pending = (pendingAttach[peerId] || []).slice();
        if (!pending.length && !String(text || '').trim()) {
            return Promise.resolve(false);
        }
        if (input) {
            input.value = '';
            autosizeInput(input);
        }
        if (pending.length) {
            const files = pending.map((p) => p.file);
            clearPendingAttach(peerId);
            const replySnap = takeReplyTarget(peerId);
            renderDock();
            const next = el('chatDock') && el('chatDock').querySelector('[data-chat-input="' + peerId + '"]');
            if (next) {
                next.value = '';
                autosizeInput(next);
            }
            let chain = Promise.resolve(true);
            files.forEach((file, idx) => {
                const caption = idx === 0 ? text : '';
                const reply = idx === 0 ? replySnap : null;
                chain = chain.then(() => sendFile(peerId, file, caption, reply));
            });
            return chain;
        }
        return sendMessage(peerId, text).then((ok) => {
            if (!ok && String(text || '').trim()) {
                const again = el('chatDock') && el('chatDock').querySelector('[data-chat-input="' + peerId + '"]');
                if (again) {
                    again.value = text;
                    autosizeInput(again);
                }
            }
            return ok;
        });
    }

    function sendMessage(peerId, text) {
        const raw = String(text || '').trim();
        const sess = findSession(peerId);
        if (!raw || !sess || !myDeviceId) {
            return Promise.resolve(false);
        }
        if (peerId === myDeviceId) {
            return Promise.resolve(false);
        }
        if (raw.length > MAX_TEXT) {
            return Promise.resolve(false);
        }
        const replyTo = takeReplyTarget(peerId);
        const now = Date.now();
        const m = {
            id: msgId(),
            fromDeviceId: myDeviceId,
            toDeviceId: peerId,
            text: raw,
            at: now,
            kind: 'text',
            status: 'ready',
            replyTo: replyTo
        };
        const t = getThread(peerId);
        t.messages = mergeMessages(t.messages, [m]);
        t.peerTag = sess.deviceTag || t.peerTag;
        t.peerLabel = sess.label || t.peerLabel;
        t.updatedAt = now;
        setThread(peerId, t);
        noteReplied(peerId);
        scrollPinned[peerId] = true;
        const dock = el('chatDock');
        const box = dock && dock.querySelector('[data-chat-messages="' + peerId + '"]');
        renderMessagesInto(box, t.messages, peerId);
        // Cập nhật dải "Đang trả lời" nếu vừa clear
        const draftHost = dock && dock.querySelector('[data-chat-reply-draft="' + peerId + '"]');
        if (draftHost) {
            const wrap = document.createElement('div');
            wrap.innerHTML = replyDraftHtml(peerId);
            if (wrap.firstElementChild) {
                draftHost.replaceWith(wrap.firstElementChild);
            }
        }
        markPeerRead(peerId);
        setTimeout(() => refreshStatusLine(peerId), SENT_FLASH_MS + 40);
        return pushMailboxToPeer(peerId, t.messages).then(() => true);
    }

    function uploadToStorage(file, path) {
        const st = ensureStorage();
        if (!st) {
            return Promise.reject(new Error('Firebase Storage chưa sẵn sàng'));
        }
        const meta = {
            contentType: file.type || 'application/octet-stream',
            customMetadata: {
                fromDeviceId: myDeviceId,
                originalName: String(file.name || '').slice(0, 120)
            }
        };
        return waitForAuth(15000).then(() => {
            const ref = st.ref(path);
            const uploadPromise = ref.put(file, meta);
            const done = typeof uploadPromise.then === 'function'
                ? uploadPromise
                : new Promise((resolve, reject) => {
                    uploadPromise.on(
                        'state_changed',
                        (snap) => {
                            if (snap && snap.state === 'success') {
                                resolve(snap);
                            }
                        },
                        reject,
                        () => resolve(uploadPromise.snapshot)
                    );
                });
            return Promise.race([
                done,
                new Promise((_, reject) => setTimeout(
                    () => reject(new Error('Tải file quá lâu — bật Storage + rules trên Firebase Console.')),
                    UPLOAD_TIMEOUT_MS
                ))
            ]);
        }).then((snapshot) => {
            const ref = (snapshot && snapshot.ref) ? snapshot.ref : st.ref(path);
            return ref.getDownloadURL()
                .then((url) => ({ url: url, path: path }))
                .catch((err) => {
                    console.warn('[chat] getDownloadURL failed — dùng path', err);
                    return { url: '', path: path };
                });
        });
    }

    function sendFile(peerId, file, caption, replyTo) {
        const sess = findSession(peerId);
        if (!sess || !myDeviceId || !file || peerId === myDeviceId) {
            return Promise.resolve(false);
        }
        if (file.size > MAX_FILE_BYTES) {
            window.alert('File tối đa ' + formatFileSize(MAX_FILE_BYTES) + '.');
            return Promise.resolve(false);
        }
        const hasStorage = !!ensureStorage();
        const now = Date.now();
        const mid = msgId();
        const path = buildStoragePath(peerId, mid, file.name);
        const cap = String(caption || '').trim().slice(0, MAX_TEXT);
        setUploadPreview(mid, file);
        const pending = {
            id: mid,
            fromDeviceId: myDeviceId,
            toDeviceId: peerId,
            text: cap,
            at: now,
            kind: 'file',
            status: 'uploading',
            replyTo: replyTo || null,
            file: {
                name: String(file.name || 'file').slice(0, 120),
                size: Number(file.size) || 0,
                mime: String(file.type || 'application/octet-stream').slice(0, 120),
                path: path,
                url: ''
            }
        };
        const t = getThread(peerId);
        t.messages = mergeMessages(t.messages, [pending]);
        t.peerTag = sess.deviceTag || t.peerTag;
        t.peerLabel = sess.label || t.peerLabel;
        t.updatedAt = now;
        setThread(peerId, t);
        noteReplied(peerId);
        scrollPinned[peerId] = true;
        const dock = el('chatDock');
        const box = dock && dock.querySelector('[data-chat-messages="' + peerId + '"]');
        renderMessagesInto(box, t.messages, peerId);
        markPeerRead(peerId);
        setTimeout(() => refreshStatusLine(peerId), SENT_FLASH_MS + 40);

        idbPut(mid, file);
        const isImage = String(file.type || '').indexOf('image/') === 0;
        if (isImage && PREFER_INLINE_IMAGE_SEND) {
            if (file.size > MAX_INLINE_IMAGE_BYTES) {
                clearUploadPreview(mid);
                updateThreadMessage(peerId, mid, { status: 'error' });
                window.alert('Ảnh quá lớn cho chế độ miễn phí. Giới hạn ~350KB khi gửi inline.');
                return Promise.resolve(false);
            }
            return fileToDataUrl(file).then((dataUrl) => {
                clearUploadPreview(mid);
                updateThreadMessage(peerId, mid, {
                    status: 'ready',
                    file: { url: dataUrl, path: '' }
                });
                const readyThread = getThread(peerId);
                return pushMailboxToPeer(peerId, readyThread.messages, { force: true }).then(() => true);
            }).catch((err) => {
                clearUploadPreview(mid);
                updateThreadMessage(peerId, mid, { status: 'error' });
                window.alert((err && err.message) ? err.message : 'Không đọc được ảnh để gửi.');
                return false;
            });
        }

        if (!hasStorage) {
            if (!isImage) {
                clearUploadPreview(mid);
                updateThreadMessage(peerId, mid, { status: 'error' });
                window.alert('Project chưa bật Firebase Storage. Chỉ hỗ trợ paste/gửi ảnh nhỏ dạng inline.');
                return Promise.resolve(false);
            }
        }
        return waitForAuth(15000).then(() => uploadToStorage(file, path)).then((res) => {
            clearUploadPreview(mid);
            updateThreadMessage(peerId, mid, {
                status: 'ready',
                file: { url: res.url || '', path: res.path }
            });
            const readyThread = getThread(peerId);
            return pushMailboxToPeer(peerId, readyThread.messages, { force: true }).then(() => true);
        }).catch((err) => {
            console.warn('[chat] file upload failed', err);
            clearUploadPreview(mid);
            updateThreadMessage(peerId, mid, { status: 'error' });
            const code = err && err.code ? String(err.code) + ': ' : '';
            const msg = err && err.message ? String(err.message) : 'Không gửi được file.';
            window.alert(code + msg + '\n\nKiểm tra: Firebase Console → Storage (đã bật?) → Rules → Anonymous Auth.');
            return false;
        });
    }

    function pushMailboxToPeer(peerId, messages, opts) {
        if (!db || !peerId || peerId === myDeviceId) {
            return Promise.resolve();
        }
        const options = opts || {};
        const now = Date.now();
        if (!options.force && lastPushAt[peerId] && (now - lastPushAt[peerId]) < 800) {
            return Promise.resolve();
        }
        const serialized = (messages || []).map(serializeMsgForMailbox).filter(Boolean);
        if (!serialized.length && !options.readReceipt && !options.pins) {
            return Promise.resolve();
        }
        lastPushAt[peerId] = now;
        const payload = {
            fromDeviceId: myDeviceId,
            fromDeviceTag: myDeviceTag,
            at: now,
            messages: serialized.slice(-MSG_LIMIT)
        };
        if (options.readReceipt) {
            payload.readReceipt = { readAt: now };
        }
        if (options.pins || options.readReceipt || serialized.length) {
            payload.pins = (getThread(peerId).pinnedIds || []).slice(0, MAX_PINS);
        }
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
        if (payload.readReceipt) {
            applyPeerReadReceipt(fromId, payload.readReceipt.readAt);
        }
        if (Object.prototype.hasOwnProperty.call(payload, 'pins')) {
            const tPins = getThread(fromId);
            const nextPins = Array.isArray(payload.pins)
                ? payload.pins.map((id) => String(id)).filter(Boolean).slice(0, MAX_PINS)
                : [];
            const prevSig = (tPins.peerPinnedIds || []).join('|');
            const nextSig = nextPins.join('|');
            if (prevSig !== nextSig) {
                tPins.peerPinnedIds = nextPins;
                tPins.updatedAt = Date.now();
                setThread(fromId, tPins);
                if (findSession(fromId) && !findSession(fromId).minimized) {
                    refreshThreadUi(fromId);
                } else if (window.PresenceBridge && typeof window.PresenceBridge.rerender === 'function') {
                    // banner cập nhật khi mở lại
                }
            }
        }
        const incoming = normalizeMessageList(payload.messages);
        if (!incoming.length) {
            db.ref(MAILBOX_PATH + '/' + myDeviceId + '/' + fromId).remove().catch(() => { /* ignore */ });
            return;
        }
        const t = getThread(fromId);
        const beforeLen = t.messages.length;
        const beforeSig = t.messages.map((m) => m.id + ':' + (m.recalled ? '1' : '0')).join('|');
        t.messages = mergeMessages(t.messages, incoming);
        const recalledIds = new Set(t.messages.filter((m) => m.recalled).map((m) => m.id));
        if (recalledIds.size) {
            t.pinnedIds = (t.pinnedIds || []).filter((id) => !recalledIds.has(id));
        }
        t.peerTag = String(payload.fromDeviceTag || t.peerTag || makeDeviceTag(fromId));
        t.updatedAt = Date.now();
        setThread(fromId, t);
        const afterSig = t.messages.map((m) => m.id + ':' + (m.recalled ? '1' : '0')).join('|');
        const added = t.messages.length - beforeLen;
        const changed = added > 0 || beforeSig !== afterSig;
        const sess = findSession(fromId);
        if (added > 0) {
            if (sess && !sess.minimized) {
                flashReceived(fromId);
                markPeerRead(fromId);
            } else {
                bumpUnread(fromId);
            }
        }
        db.ref(MAILBOX_PATH + '/' + myDeviceId + '/' + fromId).remove().catch(() => { /* ignore */ });
        if (sess && !sess.minimized && changed) {
            refreshThreadUi(fromId);
            if (added > 0) {
                pushMailboxToPeer(fromId, getThread(fromId).messages);
            }
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

    function onPresenceDevices(onlineDeviceIds) {
        onlineDeviceSet = onlineDeviceIds instanceof Set
            ? onlineDeviceIds
            : new Set(onlineDeviceIds || []);
        dockSessions.forEach((sess) => {
            sess.online = onlineDeviceSet.has(sess.deviceId);
        });
        renderDock();
        const store = loadStore();
        Object.keys(store).forEach((peerId) => {
            if (!onlineDeviceSet.has(peerId) || peerId === myDeviceId) {
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
        const dock = ensureDock();
        if (dock.dataset.chatBound === '1') {
            return;
        }
        dock.dataset.chatBound = '1';
        dock.addEventListener('scroll', (e) => {
            const box = e.target && e.target.getAttribute
                ? (e.target.getAttribute('data-chat-messages') ? e.target : null)
                : null;
            if (!box) {
                return;
            }
            const peerId = box.getAttribute('data-chat-messages');
            if (!peerId) {
                return;
            }
            scrollPinned[peerId] = isChatNearBottom(box);
        }, true);
        dock.addEventListener('click', (e) => {
            const t = e.target;
            if (!t || !t.closest) {
                return;
            }
            e.stopPropagation();
            const imagePreview = t.closest('[data-chat-image-preview]');
            if (imagePreview) {
                e.preventDefault();
                const img = imagePreview.querySelector('.chat-file-thumb');
                const src = img && (img.currentSrc || img.src);
                if (src) {
                    openImagePreview(src, img.alt || '');
                }
                return;
            }
            const pinsOpen = t.closest('[data-chat-pins-open]');
            if (pinsOpen) {
                openPinnedModal(pinsOpen.getAttribute('data-chat-pins-open'));
                return;
            }
            const replyCancel = t.closest('[data-chat-reply-cancel]');
            if (replyCancel) {
                clearReplyTarget(replyCancel.getAttribute('data-chat-reply-cancel'));
                renderDock();
                return;
            }
            const replyBtn = t.closest('[data-chat-msg-reply]');
            if (replyBtn) {
                const row = replyBtn.closest('.chat-bubble-row');
                const peerId = row && row.getAttribute('data-msg-peer');
                const msgId = row && row.getAttribute('data-msg-id');
                closeAllMsgMenus();
                if (peerId && msgId) {
                    setReplyTarget(peerId, msgId);
                }
                return;
            }
            const moreBtn = t.closest('[data-chat-msg-more]');
            if (moreBtn) {
                const row = moreBtn.closest('.chat-bubble-row');
                const menu = row && row.querySelector('[data-chat-msg-menu]');
                const willOpen = menu && !menu.classList.contains('is-open');
                closeAllMsgMenus(willOpen ? menu : null);
                if (menu && willOpen) {
                    menu.classList.add('is-open');
                    moreBtn.classList.add('is-open');
                }
                return;
            }
            const recallBtn = t.closest('[data-chat-recall]');
            if (recallBtn) {
                const row = recallBtn.closest('.chat-bubble-row');
                const peerId = row && row.getAttribute('data-msg-peer');
                const msgId = row && row.getAttribute('data-msg-id');
                closeAllMsgMenus();
                if (peerId && msgId) {
                    recallMessage(peerId, msgId);
                }
                return;
            }
            const pinBtn = t.closest('[data-chat-pin-toggle]');
            if (pinBtn) {
                const row = pinBtn.closest('.chat-bubble-row');
                const peerId = row && row.getAttribute('data-msg-peer');
                const msgId = row && row.getAttribute('data-msg-id');
                closeAllMsgMenus();
                if (peerId && msgId) {
                    togglePinMessage(peerId, msgId);
                }
                return;
            }
            if (!t.closest('[data-chat-msg-menu]')) {
                closeAllMsgMenus();
            }
            const draftRm = t.closest('[data-chat-draft-remove]');
            if (draftRm) {
                removePendingAttachItem(
                    draftRm.getAttribute('data-chat-draft-remove'),
                    draftRm.getAttribute('data-chat-draft-id')
                );
                renderDock();
                return;
            }
            const attachBtn = t.closest('[data-chat-attach]');
            if (attachBtn) {
                const peerId = attachBtn.getAttribute('data-chat-attach');
                const input = dock.querySelector('[data-chat-file-input="' + peerId + '"]');
                if (input) {
                    input.click();
                }
                return;
            }
            const closeBtn = t.closest('[data-chat-close]');
            if (closeBtn) {
                closeChat({ peerId: closeBtn.getAttribute('data-chat-close') });
                return;
            }
            const minBtn = t.closest('[data-chat-min]');
            if (minBtn) {
                minimizeChat(minBtn.getAttribute('data-chat-min'));
                return;
            }
            const expandBtn = t.closest('[data-chat-expand]');
            if (expandBtn) {
                expandChat(expandBtn.getAttribute('data-chat-expand'));
                return;
            }
            const header = t.closest('[data-chat-header]');
            if (header && !t.closest('.chat-header-actions')) {
                minimizeChat(header.getAttribute('data-chat-header'));
            }
        });
        document.addEventListener('click', (e) => {
            if (e.target && e.target.closest && e.target.closest('#chatDock')) {
                return;
            }
            closeAllMsgMenus();
        });
        dock.addEventListener('change', (e) => {
            const input = e.target && e.target.closest
                ? e.target.closest('[data-chat-file-input]')
                : null;
            if (!input || !input.files || !input.files.length) {
                return;
            }
            const peerId = input.getAttribute('data-chat-file-input');
            const files = Array.prototype.slice.call(input.files);
            input.value = '';
            stageFiles(peerId, files);
        });
        dock.addEventListener('paste', (e) => {
            const input = e.target && e.target.closest
                ? e.target.closest('[data-chat-input]')
                : null;
            if (!input) {
                return;
            }
            const peerId = input.getAttribute('data-chat-input');
            const cd = e.clipboardData;
            if (!peerId || !cd) {
                return;
            }
            const blobs = [];
            if (cd.items && cd.items.length) {
                for (let i = 0; i < cd.items.length; i++) {
                    const it = cd.items[i];
                    if (it && it.type && it.type.indexOf('image/') === 0) {
                        const b = it.getAsFile();
                        if (b) {
                            blobs.push(b);
                        }
                    }
                }
            }
            if (!blobs.length && cd.files && cd.files.length) {
                for (let j = 0; j < cd.files.length; j++) {
                    const f = cd.files[j];
                    if (f && String(f.type || '').indexOf('image/') === 0) {
                        blobs.push(f);
                    }
                }
            }
            if (!blobs.length) {
                return;
            }
            e.preventDefault();
            e.stopPropagation();
            const files = blobs.map((blob, idx) => {
                const mime = blob.type || 'image/png';
                const ext = (mime.split('/')[1] || 'png').replace(/[^a-z0-9]/gi, '') || 'png';
                if (blob instanceof File && blob.name) {
                    return blob;
                }
                return new File([blob], 'clipboard-' + Date.now() + '-' + idx + '.' + ext, { type: mime });
            });
            stageFiles(peerId, files);
        });
        dock.addEventListener('submit', (e) => {
            const form = e.target && e.target.closest
                ? e.target.closest('[data-chat-form]')
                : null;
            if (!form) {
                return;
            }
            e.preventDefault();
            e.stopPropagation();
            submitComposer(form.getAttribute('data-chat-form'));
        });
        dock.addEventListener('input', (e) => {
            const input = e.target && e.target.closest
                ? e.target.closest('[data-chat-input]')
                : null;
            if (input) {
                autosizeInput(input);
            }
        });
        dock.addEventListener('keydown', (e) => {
            e.stopPropagation();
            const input = e.target && e.target.closest
                ? e.target.closest('[data-chat-input]')
                : null;
            if (!input || e.key !== 'Enter' || e.shiftKey) {
                return;
            }
            e.preventDefault();
            const form = input.closest('[data-chat-form]');
            if (form && typeof form.requestSubmit === 'function') {
                form.requestSubmit();
            } else if (form) {
                form.dispatchEvent(new Event('submit', { cancelable: true, bubbles: true }));
            }
        });
        dock.addEventListener('mousedown', (e) => e.stopPropagation());

        window.addEventListener('storage', (e) => {
            if (e && e.key === CHAT_STORE_KEY) {
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
                    openChat(Object.assign({}, data.peer, { silent: true }));
                } else if (data.type === 'close') {
                    closeChat({ peerId: data.peerId || '', silent: true });
                } else if (data.type === 'min' && data.peerId) {
                    minimizeChat(data.peerId, true);
                } else if (data.type === 'expand' && data.peerId) {
                    expandChat(data.peerId, true);
                }
            };
        } catch (e) { /* ignore */ }
    }

    function startAliveLoop() {
        onIdleVisibilityOrTick();
        if (aliveTimer) {
            clearInterval(aliveTimer);
        }
        aliveTimer = setInterval(() => {
            if (document.hidden) {
                markAway();
                return;
            }
            touchAlive();
        }, ALIVE_TICK_MS);
        window.addEventListener('pagehide', markAway);
        window.addEventListener('beforeunload', markAway);
        bindIdleLifecycle();
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
        loadAvatarStore();
        loadPeerSeen();
        loadRepliedSession();
        myDeviceId = getOrCreateDeviceId();
        myDeviceTag = makeDeviceTag(myDeviceId);
        bindUi();
        bindBroadcast();
        startAliveLoop();
        ensureStatusTicker();
        renderDock();
        if (window.PresenceBridge && typeof window.PresenceBridge.rerender === 'function') {
            window.PresenceBridge.rerender();
        }

        const boot = () => {
            try {
                if (!firebase.apps.length && isConfigReady(cfg)) {
                    firebase.initializeApp(cfg);
                }
            } catch (e) { /* already init */ }
            db = firebase.database();
            ensureStorage();
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
        minimize: minimizeChat,
        expand: expandChat,
        getMyDeviceId: () => myDeviceId,
        isOpen: isChatOpen,
        getUnread: (peerId) => Number(unreadMap[peerId]) || 0,
        getUnreadMap: () => Object.assign({}, unreadMap),
        getBorderState: getBorderState,
        getAvatar: getAvatar,
        avatarHtml: avatarHtml
    };

    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', startChat);
    } else {
        startChat();
    }
})();
