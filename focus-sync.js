/**
 * Sync focus id (+ popup) giữa index.html ↔ trial.html (nhiều tab).
 * CHỈ sync khi hai tab cùng mã định danh (tag ≠ '').
 * BroadcastChannel (realtime) + localStorage theo tag (tab mở sau / fallback).
 *
 * state: { id:number, popup:boolean, mode:'hist'|'sim'|null, tag:string }
 * Aux: sim-pha-picks — pick phá kỳ rỗng cuối (trial → index khoanh).
 * source vẫn dùng 'ticks' để tương thích echo-filter cũ.
 */
(function (global) {
  'use strict';

  const CHANNEL = 'sua-focus-id';
  const STORAGE_KEY = 'sua:focusId';
  const SIM_PHA_STORAGE_KEY = 'sua:simPhaPicks:index';
  const SIM_PHA_ROOM_PREFIX = 'sua:simPhaPicks:';

  /** @type {BroadcastChannel|null} */
  let bc = null;
  /** @type {{ id: number|null, popup: boolean, mode: string|null }} */
  let lastState = { id: null, popup: false, mode: null };
  let applying = false;
  /** @type {((state: object, meta: object) => void)|null} */
  let applyFn = null;
  /** @type {((payload: object, meta: object) => void)|null} */
  let simPhaApplyFn = null;
  /** @type {((payload: object, meta: object) => void)|null} */
  let simPhaJoinFn = null;
  let sourceName = 'page';
  /** Mã định danh của tab này ('' = không sync focus). */
  let roomTag = '';
  /** @type {(() => string)|null} */
  let getRoomTagFn = null;
  /** @type {(() => object|null)|null} focus hiện tại của tab (kể cả trước khi có mã). */
  let getCurrentStateFn = null;
  let joinTimer = 0;
  let joinAwaiting = false;
  let joinResponded = false;

  function normalizeSimRoomTag(raw) {
    return String(raw == null ? '' : raw)
      .replace(/[^a-zA-Z0-9]/g, '')
      .slice(0, 3)
      .toUpperCase();
  }

  function simPhaRoomKey(tag) {
    return SIM_PHA_ROOM_PREFIX + tag;
  }

  function currentTag() {
    if (typeof getRoomTagFn === 'function') {
      try {
        return normalizeSimRoomTag(getRoomTagFn());
      } catch (e) {
        return '';
      }
    }
    return normalizeSimRoomTag(roomTag);
  }

  function storageKeyForTag(tag) {
    const t = normalizeSimRoomTag(tag);
    return t ? STORAGE_KEY + ':' + t : null;
  }

  function setRoomTag(tag) {
    const next = normalizeSimRoomTag(tag);
    if (next === roomTag) return;
    roomTag = next;
    lastState = { id: null, popup: false, mode: null };
  }

  function parseId(raw) {
    const n = Number(raw);
    if (!Number.isFinite(n) || n <= 0) return null;
    return Math.floor(n);
  }

  function normalizeMode(mode, popup) {
    if (!popup) return null;
    return mode === 'sim' ? 'sim' : 'hist';
  }

  function normalizeState(input) {
    if (input == null) return null;
    if (typeof input === 'number' || typeof input === 'string') {
      const id = parseId(input);
      if (id == null) return null;
      return { id: id, popup: false, mode: null };
    }
    const id = parseId(input.id);
    if (id == null) return null;
    const popup = !!input.popup;
    return { id: id, popup: popup, mode: normalizeMode(input.mode, popup) };
  }

  function sameState(a, b) {
    if (!a || !b) return false;
    return a.id === b.id && !!a.popup === !!b.popup && (a.mode || null) === (b.mode || null);
  }

  function readStored() {
    const key = storageKeyForTag(currentTag());
    if (!key) return null;
    try {
      const raw = localStorage.getItem(key);
      if (raw == null || raw === '') return null;
      if (raw.charAt(0) === '{') {
        const o = JSON.parse(raw);
        if (o && o.tag && normalizeSimRoomTag(o.tag) !== currentTag()) return null;
        return normalizeState(o);
      }
      return normalizeState(raw);
    } catch (e) {
      return null;
    }
  }

  function writeStored(state) {
    const tag = currentTag();
    const key = storageKeyForTag(tag);
    if (!key) return;
    try {
      localStorage.setItem(key, JSON.stringify({
        id: state.id,
        popup: !!state.popup,
        mode: state.mode || null,
        tag: tag
      }));
    } catch (e) { /* ignore */ }
  }

  function clearApplyingSoon() {
    // Không dùng rAF: tab ẩn / treo lâu có thể hoãn rAF rất lâu → publish bị nuốt.
    setTimeout(function () {
      applying = false;
    }, 0);
  }

  function localCurrentState() {
    if (lastState.id != null) return lastState;
    if (typeof getCurrentStateFn === 'function') {
      try {
        return normalizeState(getCurrentStateFn());
      } catch (e) {
        return null;
      }
    }
    return null;
  }

  function deliver(state, meta) {
    const tag = currentTag();
    if (!tag) return;
    if (meta && meta.remote) {
      joinResponded = true;
      joinAwaiting = false;
    }
    const inTag = (meta && meta.tag != null)
      ? normalizeSimRoomTag(meta.tag)
      : (state && typeof state === 'object' ? normalizeSimRoomTag(state.tag) : '');
    /* BC/storage focus bắt buộc đúng mã; storage theo key tag thì inTag có thể trống. */
    if (inTag && inTag !== tag) return;
    if (meta && meta.viaChannel && !inTag) return;
    const next = normalizeState(state);
    if (!next) return;
    if (sameState(next, lastState) && !(meta && meta.force)) return;
    lastState = next;
    if (typeof applyFn !== 'function') return;
    applying = true;
    try {
      applyFn(next, meta || {});
    } finally {
      clearApplyingSoon();
    }
  }

  /**
   * @param {number|string|object} rawIdOrState
   * @param {{ popup?: boolean, mode?: string|null, force?: boolean }} [opts]
   *   - không truyền popup → giữ popup/mode trước đó (đổi id từ index vẫn mở popup nếu đang mở)
   *   - popup:true/false → set rõ
   *   - force:true → vẫn broadcast dù state trùng (click lại cùng id trên index → mở lại popup ticks)
   */
  function publish(rawIdOrState, opts) {
    if (applying) return;
    const tag = currentTag();
    if (!tag) return;
    opts = opts || {};
    var force = !!opts.force;
    var base = null;
    if (rawIdOrState && typeof rawIdOrState === 'object' && !Array.isArray(rawIdOrState)) {
      base = normalizeState(rawIdOrState);
      if (rawIdOrState && Object.prototype.hasOwnProperty.call(rawIdOrState, 'force')) {
        force = !!rawIdOrState.force || force;
      }
    } else {
      var id = parseId(rawIdOrState);
      if (id == null) return;
      var popup = Object.prototype.hasOwnProperty.call(opts, 'popup')
        ? !!opts.popup
        : !!lastState.popup;
      var mode = Object.prototype.hasOwnProperty.call(opts, 'mode')
        ? opts.mode
        : (popup ? lastState.mode : null);
      if (Object.prototype.hasOwnProperty.call(opts, 'popup') && !opts.popup) {
        mode = null;
      }
      base = normalizeState({ id: id, popup: popup, mode: mode });
    }
    if (!base) return;
    if (sameState(base, lastState) && !force) return;
    lastState = base;
    writeStored(base);
    if (bc) {
      try {
        bc.postMessage({
          type: 'focus',
          id: base.id,
          popup: base.popup,
          mode: base.mode,
          force: force,
          tag: tag,
          t: Date.now(),
          source: sourceName
        });
      } catch (e) { /* ignore */ }
    }
  }

  /**
   * @param {{ source?: string, getRoomTag?: () => string, roomTag?: string,
   *   getCurrentState?: () => object|null,
   *   apply?: (state: object, meta: object) => void }} opts
   */
  function init(opts) {
    opts = opts || {};
    sourceName = opts.source || 'page';
    applyFn = typeof opts.apply === 'function' ? opts.apply : null;
    if (typeof opts.getRoomTag === 'function') getRoomTagFn = opts.getRoomTag;
    if (typeof opts.getCurrentState === 'function') {
      getCurrentStateFn = opts.getCurrentState;
    }
    if (opts.roomTag != null) roomTag = normalizeSimRoomTag(opts.roomTag);

    if (!bc && typeof BroadcastChannel !== 'undefined') {
      try {
        bc = new BroadcastChannel(CHANNEL);
        bc.onmessage = function (ev) {
          const d = ev && ev.data;
          if (!d || !d.type) return;
          if (d.type === 'sim-pha-picks') {
            deliverSimPha(d, { remote: true, source: d.source });
            return;
          }
          if (d.type === 'sim-pha-join') {
            deliverSimPhaJoin(d, { remote: true, source: d.source });
            return;
          }
          if (d.type === 'focus-join') {
            /* Tab vào mã sau hỏi → tab có mã trước trả focus hiện tại của nó. */
            const tag = currentTag();
            if (!tag || normalizeSimRoomTag(d.tag) !== tag) return;
            const st = localCurrentState();
            if (st) publish(st, { force: true });
            return;
          }
          if (d.type !== 'focus') return;
          deliver(
            { id: d.id, popup: d.popup, mode: d.mode, tag: d.tag },
            {
              remote: true,
              source: d.source,
              force: !!d.force,
              tag: d.tag,
              viaChannel: true
            }
          );
        };
      } catch (e) {
        bc = null;
      }
    }

    window.addEventListener('storage', function (ev) {
      if (!ev) return;
      if (ev.key === SIM_PHA_STORAGE_KEY) {
        if (!ev.newValue) return;
        try {
          const o = JSON.parse(ev.newValue);
          if (o && Array.isArray(o.picks)) {
            deliverSimPha(o, { remote: true, source: 'storage' });
          }
        } catch (e) { /* ignore */ }
        return;
      }
      const key = storageKeyForTag(currentTag());
      if (!key || ev.key !== key) return;
      var st = null;
      var stTag = '';
      try {
        if (ev.newValue && ev.newValue.charAt(0) === '{') {
          const o = JSON.parse(ev.newValue);
          stTag = o ? normalizeSimRoomTag(o.tag) : '';
          st = normalizeState(o);
        } else {
          st = normalizeState(ev.newValue);
        }
      } catch (e) {
        st = normalizeState(ev.newValue);
      }
      deliver(st, { remote: true, source: 'storage', tag: stTag });
    });
  }

  function deliverSimPha(payload, meta) {
    if (typeof simPhaApplyFn !== 'function') return;
    try {
      simPhaApplyFn(payload || {}, meta || {});
    } catch (e) { /* ignore */ }
  }

  function deliverSimPhaJoin(payload, meta) {
    if (typeof simPhaJoinFn !== 'function') return;
    try {
      simPhaJoinFn(payload || {}, meta || {});
    } catch (e) { /* ignore */ }
  }

  /**
   * Publish picks phá sim — chỉ khi có mã phòng (index↔ticks cùng mã).
   * @param {number[]} picks
   * @param {{ tag?: string }} [opts]
   */
  function publishSimPhaPicks(picks, opts) {
    opts = opts || {};
    const tag = normalizeSimRoomTag(opts.tag);
    if (!tag) return;
    const list = [];
    const seen = {};
    const raw = Array.isArray(picks) ? picks : [];
    for (let i = 0; i < raw.length; i++) {
      if (list.length >= 5) break;
      const n = Number(raw[i]);
      if (!Number.isFinite(n) || n < 1 || n > 35 || seen[n]) continue;
      seen[n] = true;
      list.push(n);
    }
    const payload = {
      type: 'sim-pha-picks',
      picks: list,
      tag: tag,
      t: Date.now(),
      source: sourceName
    };
    try {
      localStorage.setItem(SIM_PHA_STORAGE_KEY, JSON.stringify(payload));
    } catch (e) { /* ignore */ }
    try {
      localStorage.setItem(simPhaRoomKey(tag), JSON.stringify(payload));
    } catch (e) { /* ignore */ }
    if (bc) {
      try {
        bc.postMessage(payload);
      } catch (e2) { /* ignore */ }
    }
  }

  /** Tab mới vào phòng → xin trạng thái từ tab đang giữ mã. */
  function publishSimPhaJoin(tag) {
    const t = normalizeSimRoomTag(tag);
    if (!t) return;
    const payload = {
      type: 'sim-pha-join',
      tag: t,
      t: Date.now(),
      source: sourceName
    };
    if (bc) {
      try {
        bc.postMessage(payload);
      } catch (e) { /* ignore */ }
    }
  }

  function readSimPhaRoom(tag) {
    const t = normalizeSimRoomTag(tag);
    if (!t) return null;
    try {
      const raw = localStorage.getItem(simPhaRoomKey(t));
      if (!raw) return null;
      const o = JSON.parse(raw);
      if (!o || !Array.isArray(o.picks)) return null;
      return o;
    } catch (e) {
      return null;
    }
  }

  /** @param {(payload: object, meta: object) => void} fn */
  function onSimPhaPicks(fn) {
    simPhaApplyFn = typeof fn === 'function' ? fn : null;
  }

  /** @param {(payload: object, meta: object) => void} fn */
  function onSimPhaJoin(fn) {
    simPhaJoinFn = typeof fn === 'function' ? fn : null;
  }

  function get() {
    if (!currentTag()) return { id: null, popup: false, mode: null };
    return lastState.id != null ? lastState : readStored();
  }

  function getId() {
    var st = get();
    return st && st.id != null ? st.id : null;
  }

  /** Đọc ?id= / #id= rồi apply + broadcast (chỉ khi đã có mã định danh). */
  function seedFromUrlOrStorage(opts) {
    opts = opts || {};
    if (!currentTag()) return null;
    var id = null;
    try {
      var sp = new URLSearchParams(location.search);
      id = parseId(sp.get('id'));
      if (id == null && location.hash) {
        var m = String(location.hash).match(/(?:^[#?]|&)id=(\d+)/);
        if (m) id = parseId(m[1]);
      }
    } catch (e) { /* ignore */ }

    var stored = opts.useStorage !== false ? readStored() : null;
    var state = null;
    if (id != null) {
      // URL ưu tiên id; giữ popup từ storage nếu đang mở ở tab khác.
      state = normalizeState({
        id: id,
        popup: stored ? !!stored.popup : false,
        mode: stored && stored.popup ? stored.mode : null
      });
    } else if (stored) {
      state = stored;
    }
    if (!state) return null;

    lastState = state;
    writeStored(state);
    if (bc) {
      try {
        bc.postMessage({
          type: 'focus',
          id: state.id,
          popup: state.popup,
          mode: state.mode,
          tag: currentTag(),
          t: Date.now(),
          source: sourceName
        });
      } catch (e) { /* ignore */ }
    }
    if (typeof applyFn === 'function') {
      applying = true;
      try {
        applyFn(state, { seed: true });
      } finally {
        clearApplyingSoon();
      }
    }
    return state;
  }

  /**
   * Vào mã định danh:
   * 1) ăn focus đã lưu của mã (nếu có),
   * 2) hỏi tab đang giữ mã trước → nó trả focus hiện tại (theo thằng trước),
   * 3) không ai trả (≈600ms) → tab này là thằng đầu tiên: publish focus
   *    hiện tại của mình để thằng vào sau theo.
   */
  function joinRoom(tagRaw) {
    setRoomTag(tagRaw);
    const tag = currentTag();
    if (!tag) return null;
    joinAwaiting = true;
    joinResponded = false;
    const seeded = seedFromUrlOrStorage({ useStorage: true });
    if (seeded) {
      /* Đã có focus của phòng trong storage — vẫn chờ bản live mới hơn. */
      joinAwaiting = true;
      joinResponded = false;
    }
    if (bc) {
      try {
        bc.postMessage({
          type: 'focus-join',
          tag: tag,
          t: Date.now(),
          source: sourceName
        });
      } catch (e) { /* ignore */ }
    }
    if (joinTimer) clearTimeout(joinTimer);
    joinTimer = setTimeout(function () {
      joinTimer = 0;
      if (!joinAwaiting || joinResponded) return;
      joinAwaiting = false;
      if (seeded) return;
      const st = localCurrentState();
      if (st) publish(st, { force: true });
    }, 600);
    return seeded;
  }

  global.FocusSync = {
    init: init,
    setRoomTag: setRoomTag,
    joinRoom: joinRoom,
    getRoomTag: currentTag,
    publish: publish,
    publishSimPhaPicks: publishSimPhaPicks,
    publishSimPhaJoin: publishSimPhaJoin,
    readSimPhaRoom: readSimPhaRoom,
    onSimPhaPicks: onSimPhaPicks,
    onSimPhaJoin: onSimPhaJoin,
    normalizeSimRoomTag: normalizeSimRoomTag,
    get: get,
    getId: getId,
    readStored: readStored,
    seedFromUrlOrStorage: seedFromUrlOrStorage,
    isApplying: function () { return applying; },
    SIM_PHA_STORAGE_KEY: SIM_PHA_STORAGE_KEY
  };
})(typeof window !== 'undefined' ? window : this);
