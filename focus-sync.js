/**
 * Sync focus id (+ popup ticks) giữa index.html ↔ ticks.html (nhiều tab).
 * BroadcastChannel (realtime) + localStorage (tab mở sau / fallback).
 *
 * state: { id:number, popup:boolean, mode:'hist'|'sim'|null }
 */
(function (global) {
  'use strict';

  const CHANNEL = 'sua-focus-id';
  const STORAGE_KEY = 'sua:focusId';

  /** @type {BroadcastChannel|null} */
  let bc = null;
  /** @type {{ id: number|null, popup: boolean, mode: string|null }} */
  let lastState = { id: null, popup: false, mode: null };
  let applying = false;
  /** @type {((state: object, meta: object) => void)|null} */
  let applyFn = null;
  let sourceName = 'page';

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
      requestAnimationFrame(function () {
        applying = false;
      });
    }
  }

  /**
   * @param {number|string|object} rawIdOrState
   * @param {{ popup?: boolean, mode?: string|null }} [opts]
   *   - không truyền popup → giữ popup/mode trước đó (đổi id từ index vẫn mở popup nếu đang mở)
   *   - popup:true/false → set rõ
   */
  function publish(rawIdOrState, opts) {
    if (applying) return;
    opts = opts || {};
    var base = null;
    if (rawIdOrState && typeof rawIdOrState === 'object' && !Array.isArray(rawIdOrState)) {
      base = normalizeState(rawIdOrState);
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
    if (sameState(base, lastState)) return;
    lastState = base;
    writeStored(base);
    if (bc) {
      try {
        bc.postMessage({
          type: 'focus',
          id: base.id,
          popup: base.popup,
          mode: base.mode,
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
          if (!d || d.type !== 'focus') return;
          deliver(
            { id: d.id, popup: d.popup, mode: d.mode },
            { remote: true, source: d.source }
          );
        };
      } catch (e) {
        bc = null;
      }
    }

    window.addEventListener('storage', function (ev) {
      if (!ev || ev.key !== STORAGE_KEY) return;
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
        requestAnimationFrame(function () {
          applying = false;
        });
      }
    }
    return state;
  }

  global.FocusSync = {
    init: init,
    publish: publish,
    get: get,
    getId: getId,
    readStored: readStored,
    seedFromUrlOrStorage: seedFromUrlOrStorage,
    isApplying: function () { return applying; }
  };
})(typeof window !== 'undefined' ? window : this);
