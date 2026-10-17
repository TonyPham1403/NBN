/**
 * Sync focus id (+ popup ticks) giữa index.html ↔ ticks.html (nhiều tab).
 * BroadcastChannel (realtime) + localStorage (tab mở sau / fallback).
 *
 * state: { id:number, popup:boolean, mode:'hist'|'sim'|null }
 * Aux: sim-pha-picks — pick phá kỳ rỗng cuối (ticks → index khoanh).
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

  function normalizeSimRoomTag(raw) {
    return String(raw == null ? '' : raw)
      .replace(/[^a-zA-Z0-9]/g, '')
      .slice(0, 3)
      .toUpperCase();
  }

  function simPhaRoomKey(tag) {
    return SIM_PHA_ROOM_PREFIX + tag;
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
    try {
      const raw = localStorage.getItem(STORAGE_KEY);
      if (raw == null || raw === '') return null;
      if (raw.charAt(0) === '{') {
        return normalizeState(JSON.parse(raw));
      }
      return normalizeState(raw);
    } catch (e) {
      return null;
    }
  }

  function writeStored(state) {
    try {
      localStorage.setItem(STORAGE_KEY, JSON.stringify({
        id: state.id,
        popup: !!state.popup,
        mode: state.mode || null
      }));
    } catch (e) { /* ignore */ }
  }

  function clearApplyingSoon() {
    // Không dùng rAF: tab ẩn / treo lâu có thể hoãn rAF rất lâu → publish bị nuốt.
    setTimeout(function () {
      applying = false;
    }, 0);
  }

  function deliver(state, meta) {
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
          t: Date.now(),
          source: sourceName
        });
      } catch (e) { /* ignore */ }
    }
  }

  /**
   * @param {{ source?: string, apply?: (state: object, meta: object) => void }} opts
   */
  function init(opts) {
    opts = opts || {};
    sourceName = opts.source || 'page';
    applyFn = typeof opts.apply === 'function' ? opts.apply : null;

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
          if (d.type !== 'focus') return;
          deliver(
            { id: d.id, popup: d.popup, mode: d.mode },
            { remote: true, source: d.source, force: !!d.force }
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
      if (ev.key !== STORAGE_KEY) return;
      var st = null;
      try {
        if (ev.newValue && ev.newValue.charAt(0) === '{') {
          st = normalizeState(JSON.parse(ev.newValue));
        } else {
          st = normalizeState(ev.newValue);
        }
      } catch (e) {
        st = normalizeState(ev.newValue);
      }
      deliver(st, { remote: true, source: 'storage' });
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
    return lastState.id != null ? lastState : readStored();
  }

  function getId() {
    var st = get();
    return st && st.id != null ? st.id : null;
  }

  /** Đọc ?id= / #id= rồi apply + broadcast (tab mới mở mang id). */
  function seedFromUrlOrStorage(opts) {
    opts = opts || {};
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

  global.FocusSync = {
    init: init,
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
