/**
 * Right Pane Sheet Manager & Styling
 * Inspired by Module1-5 VBA patterns: grouping, frequency analysis, color-coding
 *
 * Trọng số + predictMode (`heuristic` | `globalTrigram` | `blend` | `stackedNgram` | `temporalRank`): đồng bộ proj/scripts/special-tracking-predict-core.mjs
 * (`blend` có thể bật `temporalMix` > 0 để trộn z predict.txt; mặc định 0 trên data.json hiện tại).
 * (tối ưu sâu: `node ... --tune-coord` ~vài phút; nhanh: `--tune-coord-quick`).
 */
const SPECIAL_TRACKING_PREDICT_WT = Object.freeze({
    /** Đồng bộ DEFAULT_SPECIAL_TRACKING_PREDICT_WEIGHTS (đã --tune-coord trên data.json). */
    uw50: 1.9293,
    uw20: 1.843,
    uw100: 0.622,
    c4mul: 1.439,
    bgG: 2.545,
    bgR: 11.197,
    tri: 8.98064,
    gapSqrt: 0.573,
    gapTail: 0.163,
    gapThr: 19,
    vel: 1.365,
    velHiSlot: 7,
    velHi: 1.1648,
    cumCap: 1.78,
    cumK: 0.235,
    medCap: 1.491,
    medK: 0.268,
    w144: 0.433,
    mslotVel: 0.594,
    vel10mul: 0.512,
    vel10a: 0.703,
    vel10b: 0.277,
    hot50rat: 1.669,
    penalHot50: 1.21776,
    hot100rat: 1.525,
    penalHot100: 0.539,
    echo3: 10.904,
    echo5: 1.251,
    penB: 4.509,
    repeat: 8.372,
    echoSwap: 0.05,
    bgAlpha: 0.491,
    triAlpha: 0.266,
    recentWin: 87,
    triWinCap: 107,
    /**
     * heuristic | globalTrigram | blend | stackedNgram | temporalRank — đồng bộ special-tracking-predict-core.mjs
     * blend = heuristic + triGlobal·log1p(honest trigram) + marginal Dirichlet (prefix).
     * temporalRank = predict.txt (rolling/gap/rank-velocity/recovery/z), không random.
     */
    triGlobal: 2.35,
    margAlpha: 2,
    margLong: 0.35,
    margShortWin: 72,
    margShort: 0,
    /** Chỉ blend: cộng temporalMix × z(predict.txt timeline). Đồng bộ DEFAULT trong special-tracking-predict-core.mjs */
    temporalMix: 0,
    predictMode: 'blend'
});

/** Đồng bộ globalTrigram với special-tracking-predict-core.mjs (browser không import module). */
const specialTrackingGlobalTrigramCache = new WeakMap();

function specialTrackingGlobalTrigramCacheKey(len, horizon) {
    return `${len}|${horizon}`;
}

function specialTrackingBuildGlobalTrigramByKey(series, len, horizon = null) {
    const H = horizon != null ? horizon : len;
    const byKey = new Map();
    for (let p = 1; p < len - 1; p++) {
        if (p + 1 >= H) {
            continue;
        }
        const a = series[p - 1];
        const b = series[p];
        const nx = series[p + 1];
        if (a < 1 || a > 12 || b < 1 || b > 12 || nx < 1 || nx > 12) {
            continue;
        }
        const key = a * 16 + b;
        if (!byKey.has(key)) {
            byKey.set(key, new Map());
        }
        const m = byKey.get(key);
        m.set(nx, (m.get(nx) || 0) + 1);
    }
    return byKey;
}

function specialTrackingGetGlobalTrigramByKey(series, len, horizon = null) {
    const H = horizon != null ? horizon : len;
    const key = specialTrackingGlobalTrigramCacheKey(len, H);
    let inner = specialTrackingGlobalTrigramCache.get(series);
    if (!inner) {
        inner = new Map();
        specialTrackingGlobalTrigramCache.set(series, inner);
    }
    if (!inner.has(key)) {
        inner.set(key, specialTrackingBuildGlobalTrigramByKey(series, len, H));
    }
    return inner.get(key);
}

/** @returns {number[] | null} */
function specialTrackingTop3FromGlobalTrigram(series, len, pen, prev, horizon = null) {
    const byKey = specialTrackingGetGlobalTrigramByKey(series, len, horizon);
    const key = pen * 16 + prev;
    const m = byKey.get(key);
    if (!m || m.size === 0) {
        return null;
    }
    const tp = [];
    for (let n = 1; n <= 12; n++) {
        tp.push([n, m.get(n) || 0]);
    }
    tp.sort((a, b) => b[1] - a[1] || a[0] - b[0]);
    return [tp[0][0], tp[1][0], tp[2][0]];
}

/** Đồng bộ computeStackedNgramTop3 trong special-tracking-predict-core.mjs */
function specialTrackingComputeStackedNgramTop3(series, nFull, N, wt) {
    if (N < 1) {
        return [1, 2, 3];
    }
    const alpha =
        typeof wt.stackLaplace === 'number' && Number.isFinite(wt.stackLaplace) && wt.stackLaplace > 0
            ? wt.stackLaplace
            : 0.55;
    const sm = typeof wt.stackMarg === 'number' && Number.isFinite(wt.stackMarg) ? wt.stackMarg : 0.85;
    const sb = typeof wt.stackBi === 'number' && Number.isFinite(wt.stackBi) ? wt.stackBi : 1.15;
    const st = typeof wt.stackTri === 'number' && Number.isFinite(wt.stackTri) ? wt.stackTri : 2.2;

    const pen = N >= 2 ? series[N - 2] : 0;
    const prev = series[N - 1];

    const cM = new Array(13).fill(0);
    for (let j = 0; j < N; j++) {
        const x = series[j];
        if (x >= 1 && x <= 12) {
            cM[x]++;
        }
    }
    const denM = N + 12 * alpha;

    const fb = new Array(13).fill(0);
    let cntB = 0;
    if (prev >= 1 && prev <= 12) {
        for (let j = 0; j <= N - 2; j++) {
            if (series[j] === prev) {
                cntB++;
                const nx = series[j + 1];
                if (nx >= 1 && nx <= 12) {
                    fb[nx]++;
                }
            }
        }
    }
    const denB = cntB + 12 * alpha;

    const ft = new Array(13).fill(0);
    let triTot = 0;
    if (N >= 3 && pen >= 1 && pen <= 12 && prev >= 1 && prev <= 12) {
        const byKey = specialTrackingGetGlobalTrigramByKey(series, nFull, N);
        const triMap = byKey.get(pen * 16 + prev);
        if (triMap) {
            for (let n = 1; n <= 12; n++) {
                const v = triMap.get(n) || 0;
                ft[n] = v;
                triTot += v;
            }
        }
    }
    const denT = triTot + 12 * alpha;

    const scores = new Array(13).fill(0);
    for (let n = 1; n <= 12; n++) {
        const phM = (cM[n] + alpha) / denM;
        scores[n] += sm * Math.log(12 * phM);
        const phB = (fb[n] + alpha) / denB;
        scores[n] += sb * Math.log(12 * phB);
        const phT = (ft[n] + alpha) / denT;
        scores[n] += st * Math.log(12 * phT);
    }

    const pairs = [];
    for (let n = 1; n <= 12; n++) {
        pairs.push([n, scores[n]]);
    }
    pairs.sort((a, b) => b[1] - a[1] || a[0] - b[0]);
    return [pairs[0][0], pairs[1][0], pairs[2][0]];
}

/** @param {number[]} vals13 — chỉ số 1..12 */
function specialTrackingZScore12(vals13) {
    let s = 0;
    let s2 = 0;
    const c = 12;
    for (let n = 1; n <= 12; n++) {
        const v = vals13[n] || 0;
        s += v;
        s2 += v * v;
    }
    const mean = s / c;
    const vr = s2 / c - mean * mean;
    const std = Math.sqrt(Math.max(vr, 1e-12));
    const out = new Array(13).fill(0);
    for (let n = 1; n <= 12; n++) {
        out[n] = ((vals13[n] || 0) - mean) / std;
    }
    return out;
}

/**
 * Điểm composite predict.txt (1..12) — dùng cho temporalMix trong blend.
 * @returns {number[] | null}
 */
function specialTrackingComputeTemporalRankScores13(series, frames, N) {
    const lastIdx = Math.min(frames.length, N) - 1;
    if (lastIdx < 0) {
        return null;
    }
    const last = frames[lastIdx];
    if (!last || !last.slotByNum || !last.byNum) {
        return null;
    }
    const W20 = Math.min(20, N);
    const W50 = Math.min(50, N);
    const W100 = Math.min(100, N);
    const exp20 = W20 / 12;
    const exp50 = W50 / 12;
    const exp100 = W100 / 12;

    const recentCount = (n, W) => {
        let cc = 0;
        const from = Math.max(0, N - W);
        for (let i = from; i < N; i++) {
            if (series[i] === n) {
                cc++;
            }
        }
        return cc;
    };

    const lastAt = new Array(13).fill(-1);
    const maxBetween = new Array(13).fill(0);
    const gapSum = new Array(13).fill(0);
    const gapCnt = new Array(13).fill(0);
    for (let i = 0; i < N; i++) {
        const x = series[i];
        if (x < 1 || x > 12) {
            continue;
        }
        if (lastAt[x] >= 0) {
            const g = i - lastAt[x];
            maxBetween[x] = Math.max(maxBetween[x], g);
            gapSum[x] += g;
            gapCnt[x]++;
        }
        lastAt[x] = i;
    }
    const gapCurrent = new Array(13).fill(N);
    const avgGap = new Array(13).fill(0);
    for (let n = 1; n <= 12; n++) {
        gapCurrent[n] = lastAt[n] >= 0 ? N - 1 - lastAt[n] : N;
        avgGap[n] = gapCnt[n] > 0 ? gapSum[n] / gapCnt[n] : N;
    }

    const under20 = new Array(13).fill(0);
    const under50 = new Array(13).fill(0);
    const under100 = new Array(13).fill(0);
    for (let n = 1; n <= 12; n++) {
        under20[n] = Math.max(0, exp20 - recentCount(n, W20));
        under50[n] = Math.max(0, exp50 - recentCount(n, W50));
        under100[n] = Math.max(0, exp100 - recentCount(n, W100));
    }

    const slotAt = (fIdx, n) => {
        const fr = frames[fIdx];
        return fr && fr.slotByNum ? fr.slotByNum[n] ?? 11 : 11;
    };
    const pastFi = (delta) => Math.max(0, lastIdx - delta);

    const slotNow = new Array(13).fill(0);
    const vel20 = new Array(13).fill(0);
    const vel50 = new Array(13).fill(0);
    const accel = new Array(13).fill(0);
    for (let n = 1; n <= 12; n++) {
        const sn = slotAt(lastIdx, n);
        slotNow[n] = sn;
        const s20 = slotAt(pastFi(20), n);
        const s50 = slotAt(pastFi(50), n);
        const s10 = slotAt(pastFi(10), n);
        vel20[n] = s20 - sn;
        vel50[n] = s50 - sn;
        const v10 = s10 - sn;
        accel[n] = v10 - vel20[n];
    }

    const overdueR = new Array(13).fill(0);
    const gapLog = new Array(13).fill(0);
    const maxgAnom = new Array(13).fill(0);
    for (let n = 1; n <= 12; n++) {
        const gc = gapCurrent[n];
        overdueR[n] = gc / (avgGap[n] + 0.25);
        gapLog[n] = Math.log(1 + gc / (avgGap[n] + 0.2));
        const mx = Math.max(maxBetween[n] || 0, 1);
        maxgAnom[n] = Math.log(1 + gc / (mx + 0.15));
    }

    const ideal = N / 12;
    const cumDef = new Array(13).fill(0);
    const recovery = new Array(13).fill(0);
    const leadWeakPen = new Array(13).fill(0);
    for (let n = 1; n <= 12; n++) {
        const cum = last.byNum[n] || 0;
        cumDef[n] = Math.max(0, ideal - cum);
        recovery[n] = (slotNow[n] >= 6 ? 1 : 0) * Math.max(0, vel20[n]);
        const c50 = recentCount(n, W50);
        if (slotNow[n] <= 1 && c50 > exp50 * 1.08) {
            leadWeakPen[n] = 1;
        }
    }

    const zU20 = specialTrackingZScore12(under20);
    const zU50 = specialTrackingZScore12(under50);
    const zU100 = specialTrackingZScore12(under100);
    const zGap = specialTrackingZScore12(gapLog);
    const zOd = specialTrackingZScore12(overdueR);
    const zMaxg = specialTrackingZScore12(maxgAnom);
    const zV20 = specialTrackingZScore12(vel20);
    const zV50 = specialTrackingZScore12(vel50);
    const zAcc = specialTrackingZScore12(accel);
    const zCumD = specialTrackingZScore12(cumDef);
    const zRec = specialTrackingZScore12(recovery);

    const score = new Array(13).fill(0);
    for (let n = 1; n <= 12; n++) {
        score[n] =
            0.22 * zU20[n]
            + 0.16 * zU50[n]
            + 0.08 * zU100[n]
            + 0.12 * zGap[n]
            + 0.08 * zOd[n]
            + 0.06 * zMaxg[n]
            + 0.12 * zV20[n]
            + 0.04 * zV50[n]
            + 0.06 * zAcc[n]
            + 0.06 * zCumD[n]
            + 0.08 * zRec[n]
            - (leadWeakPen[n] ? 0.72 : 0);
    }

    return score;
}

/** @returns {number[]} top-3 theo predict.txt timeline */
function specialTrackingPredictTxtTemporalTop3(series, frames, N) {
    const score = specialTrackingComputeTemporalRankScores13(series, frames, N);
    if (!score) {
        return [1, 2, 3];
    }
    const pairs = [];
    for (let n = 1; n <= 12; n++) {
        pairs.push([n, score[n]]);
    }
    pairs.sort((a, b) => b[1] - a[1] || a[0] - b[0]);
    return [pairs[0][0], pairs[1][0], pairs[2][0]];
}

/** sessionStorage + sheet.specialTrackingUi: khôi phục timeline / predict khi quay lại sheet */
const SPECIAL_TRACKING_UI_STORAGE_KEY = 'rp-special-tracking-ui-v1';

/** Chuột phải ô nonexist: nhảy tới kỳ có id = id hàng hiện tại + delta (vd 00014 → 00024 khi delta=10). */
const NONEXIST_CONTEXTMENU_ID_DELTA = 10;

class RightPaneSheetManager {
    /** Pool gợi ý Hint: số ô tối đa khoanh (top theo margin Dirichlet). */
    static MAIN_FIVE_HINT_POOL_SIZE = 10;
    /** Số dòng sheet liền trước dòng focus dùng cho cửa sổ (tối đa) — cùng ý multiset với bảng 1 trái khi đủ 10 dòng. */
    static MAIN_FIVE_HINT_SLIDE_LEN = 10;
    /**
     * Trọng số điểm margin Dirichlet trên tần suất cửa sổ (0..1); phần còn lại là tích lũy toàn bộ kỳ trước dòng.
     */
    static MAIN_FIVE_HINT_SLIDE_WEIGHT = 0.45;
    /**
     * Alpha Dirichlet chỉ dùng cho hint — lớn hơn → ít bám sát raw count, giảm “chỉ số hot”.
     */
    static MAIN_FIVE_HINT_DIRICHLET_ALPHA = 1.65;
    /**
     * Độ rộng dải xếp hạng khi chọn pool: band = min(35, max(poolN, ceil(poolN * mult))).
     * Pool lấy theo bước thứ hạng trong band (không lấy liền N ô đầu).
     */
    static MAIN_FIVE_HINT_DIVERSITY_BAND_MULT = 3;
    /**
     * true: lần đầu cần hint, chạy benchmark walk-forward trên toàn sheet1 và chọn thuật overlap TB cao nhất.
     */
    static MAIN_FIVE_HINT_AUTO_SELECT_STRATEGY = true;
    /** Số kỳ hợp lệ gần nhất (trong validIdx) dùng cho margin recent + triple. */
    static MAIN_FIVE_HINT_RECENT_VALID_K = 20;
    /**
     * predict.txt §13: số lần lấy mẫu pool có trọng số (seed deterministic từ ctx) — giữ nhỏ để benchmark không đơ.
     */
    static MAIN_FIVE_HINT_MC_POOL_SAMPLES = 40;

    constructor() {
        this.sheets = {};
        this.activeSheet = 'sheet1';
        this.sourceRows = [];
        this.dataRows = [];
        this.selectedLines = [];
        this.selectedNums = new Set();
        this.activeWindowRange = null;
        this.comboFocusRowId = '';
        this.comboFocusRowIndex = -1;
        this.answerPopupFocusMask = { active: false, rowIndex: -1 };
        this._answerPopupMaskAppliedRow = -1;
        this._answerPopupMaskApplyRaf = 0;
        this.comboG1Enabled = false;
        this.comboH1Text = '';
        this.comboHComments = {};
        this.comboHSelection = null;
        this._comboHDragSelect = null;
        this._comboHExcelWired = false;
        this._comboHMarchingVisible = false;
        this._comboHMarchingRange = null;
        this._comboHCutPending = null;
        this.scrollPositions = {};
        /** Cache filter mode connection (invalid khi refreshDerivedState). */
        this._connectionFilterIndicesCache = null;
        this._connectionFilterIndicesCacheRowLen = 0;
        this._connectionFilterNoteCacheRef = null;
        this.frequencyMap = {};
        this.colorPalette = [
            'rgb(255, 192, 0)',    // Gold
            'rgb(0, 176, 240)',    // Light Blue
            'rgb(255, 0, 0)',      // Red
            'rgb(112, 48, 160)',   // Purple
            'rgb(255, 102, 0)',    // Orange
            'rgb(128, 128, 128)'   // Gray
        ];
        this.init();
    }

    /**
     * Initialize sheets storage
     */
    init() {
        // Restore from localStorage if available, otherwise create new sheet1
        const saved = localStorage.getItem('sheetData');
        if (saved) {
            try {
                const data = JSON.parse(saved);
                this.sheets = data.sheets || { sheet1: { data: [], notes: {} } };
                this.activeSheet = data.activeSheet || 'sheet1';
                this.comboFocusRowId = data.comboFocusRowId || '';
                this.comboFocusRowIndex = Number.isFinite(data.comboFocusRowIndex) ? data.comboFocusRowIndex : -1;
                this.comboG1Enabled = !!data.comboG1Enabled;
                this.comboH1Text = data.comboH1Text || '';
                this.comboHComments = (data.comboHComments && typeof data.comboHComments === 'object')
                    ? data.comboHComments
                    : {};
                this.scrollPositions = data.scrollPositions || {};
            } catch (e) {
                this.sheets = { sheet1: { data: [], notes: {} } };
            }
        } else {
            this.sheets = { sheet1: { data: [], notes: {} } };
        }
    }

    /**
     * Load data into the active sheet
     */
    loadData(rows) {
        this.sourceRows = rows || [];
        this.rebuildSheetsFromSource();
        this.activeSheet = 'sheet1';
        this.dataRows = this.getActiveSheetRows();
        this.refreshDerivedState();
        this.save();
    }

    /**
     * Rebuild the sheet collection so sheet1 remains the raw source data and
     * combo_1..5 are derived from it.
     */
    rebuildSheetsFromSource() {
        const comboSheets = this.buildComboSheetsFromRows(this.sourceRows || []);
        const stMeta = this.buildSpecialTrackingSeriesMeta(this.sourceRows || []);
        const stFrames = this.buildSpecialTrackingFrames(stMeta.series);
        this.sheets = {
            sheet1: {
                kind: 'source',
                data: this.sourceRows || []
            },
            ...comboSheets,
            specialtracking: {
                kind: 'specialtracking',
                data: [],
                series: stMeta.series,
                seriesSourceRowIndices: stMeta.sourceRowIndices,
                frames: stFrames
            }
        };
        try {
            sessionStorage.removeItem(SPECIAL_TRACKING_UI_STORAGE_KEY);
        } catch (e) {
            /* ignore */
        }
    }

    /**
     * Return the rows for the currently active sheet.
     */
    getActiveSheetRows() {
        const sheet = this.sheets[this.activeSheet];
        if (!sheet) {
            return [];
        }
        return sheet.data || [];
    }

    /**
     * Rebuild all derived state from the current sheet data.
     * Notes are generated from result data only, following the Module4 BuildNotes logic.
     */
    refreshDerivedState() {
        this.calculateFrequency(this.sourceRows || []);
        this.noteCache = this.buildNotesFromRows(this.sourceRows || []);
        this.nonexistCache = this.buildNonexistFromRows(this.sourceRows || []);
        this.idFrequencyMap = this.buildIdFrequencyMapFromNotes(this.noteCache);
        this.nonexistGreenFilterCache = null;
        this.nonexistDisplayEntriesCache = null;
        this.datebandFilterIndicesCache = null;
        this.datebandFilterIndicesCacheRowLen = 0;
        this.datebandRowDistCache = null;
        this.datebandRowDistCacheRowLen = 0;
        this._connectionFilterIndicesCache = null;
        this._connectionFilterIndicesCacheRowLen = 0;
        this._connectionFilterNoteCacheRef = null;
        this._mainFiveHintStrategyCache = null;
    }

    /** Rows used for sheet1 / nonexist filter (independent of active combo tab). */
    getSourceSheetRows() {
        const sheet = this.sheets && this.sheets.sheet1;
        if (sheet && Array.isArray(sheet.data)) {
            return sheet.data;
        }
        return this.sourceRows || [];
    }

    /**
     * Per-row green nonexist entries for filter (sheet1 only, built lazily).
     * @returns {{ num: number, kind: string }[][]}
     */
    ensureNonexistGreenFilterCache() {
        const rows = this.getSourceSheetRows();
        if (this.nonexistGreenFilterCache && this.nonexistGreenFilterCache.length === rows.length) {
            return this.nonexistGreenFilterCache;
        }
        if (!this.nonexistCache || this.nonexistCache.length !== rows.length) {
            this.nonexistCache = this.buildNonexistFromRows(rows);
        }
        const cache = new Array(rows.length);
        for (let i = 0; i < rows.length; i++) {
            cache[i] = this.isEmptyResultRow(rows[i])
                ? []
                : this.buildNonexistGreenEntriesForRow(i);
        }
        this.nonexistGreenFilterCache = cache;
        return cache;
    }

    /**
     * Green-highlighted numbers in one row's nonexist column (single visual-state pass).
     */
    buildNonexistGreenEntriesForRow(rowIndex) {
        const row = this.getSourceSheetRows()[rowIndex];
        if (!row) {
            return [];
        }

        const nonexistMeta = this.getNonexistMetaForSourceRow(rowIndex, row);
        const nonexistText = String(nonexistMeta.text || '').trim();
        if (!nonexistText || nonexistText === 'N/A') {
            return [];
        }

        const currentResult = row.result || row.Result || '';
        const state = this.computeNonexistVisualState(rowIndex, nonexistText, currentResult);
        if (!state) {
            return [];
        }

        const entries = [];
        const candidates = this.parseNums(nonexistText);
        for (let i = 0; i < candidates.length; i++) {
            const num = candidates[i];
            const kind = this.getNonexistDisplayKindForNumber(
                rowIndex,
                num,
                nonexistText,
                currentResult,
                state
            );
            if (this.isGreenNonexistDisplayKind(kind)) {
                entries.push({ num, kind });
            }
        }
        return entries;
    }

    /**
     * Map a nonexist display kind to a filter color bucket (green / red / purple / yellow).
     */
    getNonexistColorCategory(kind) {
        if (this.isGreenNonexistDisplayKind(kind)) {
            return 'green';
        }
        if (kind === 'red') {
            return 'red';
        }
        if (kind === 'purple') {
            return 'purple';
        }
        if (kind === 'yellow') {
            return 'yellow';
        }
        return '';
    }

    /**
     * Per-row nonexist numbers with final display color (all highlight kinds).
     */
    buildNonexistDisplayEntriesForRow(rowIndex) {
        const row = this.getSourceSheetRows()[rowIndex];
        if (!row) {
            return [];
        }

        const nonexistMeta = this.getNonexistMetaForSourceRow(rowIndex, row);
        const nonexistText = String(nonexistMeta.text || '').trim();
        if (!nonexistText || nonexistText === 'N/A') {
            return [];
        }

        const currentResult = row.result || row.Result || '';
        const state = this.computeNonexistVisualState(rowIndex, nonexistText, currentResult);
        if (!state) {
            return [];
        }

        const entries = [];
        const candidates = this.parseNums(nonexistText);
        for (let i = 0; i < candidates.length; i++) {
            const num = candidates[i];
            const kind = this.getNonexistDisplayKindForNumber(
                rowIndex,
                num,
                nonexistText,
                currentResult,
                state
            );
            const color = this.getNonexistColorCategory(kind);
            if (color) {
                entries.push({ num, kind, color });
            }
        }
        return entries;
    }

    /**
     * Cached per-row nonexist display entries for color filtering.
     */
    ensureNonexistDisplayEntriesCache() {
        const rows = this.getSourceSheetRows();
        if (this.nonexistDisplayEntriesCache && this.nonexistDisplayEntriesCache.length === rows.length) {
            return this.nonexistDisplayEntriesCache;
        }
        if (!this.nonexistCache || this.nonexistCache.length !== rows.length) {
            this.nonexistCache = this.buildNonexistFromRows(rows);
        }
        const cache = new Array(rows.length);
        for (let i = 0; i < rows.length; i++) {
            cache[i] = this.isEmptyResultRow(rows[i])
                ? []
                : this.buildNonexistDisplayEntriesForRow(i);
        }
        this.nonexistDisplayEntriesCache = cache;
        this.nonexistGreenFilterCache = cache.map((entries) =>
            (entries || []).filter((entry) => entry.color === 'green').map((entry) => ({
                num: entry.num,
                kind: entry.kind
            }))
        );
        return cache;
    }

    /**
     * Calculate frequency map inspired by Module4 BuildNotes pattern
     * Groups numbers by appearance count
     */
    calculateFrequency(rows) {
        this.frequencyMap = {};
        for (const row of rows || []) {
            const nums = this.parseNums(row.result || row.Result || '');
            for (const num of nums) {
                if (!this.frequencyMap[num]) {
                    this.frequencyMap[num] = 0;
                }
                this.frequencyMap[num]++;
            }
        }
    }

    /**
     * Parse number string (comma or pipe separated)
     */
    parseNums(s) {
        if (!s) return [];
        return String(s).split(/[\|,;\s]+/).map(x => parseInt(x, 10)).filter(n => !isNaN(n));
    }

    /**
     * Find a source row by its id.
     */
    getSourceRowById(rawId) {
        const key = this.normalizeNumberKey(rawId);
        if (!key) {
            return null;
        }

        return (this.sourceRows || []).find(row => this.normalizeNumberKey(row.id || row.ID || '') === key) || null;
    }

    /**
     * True when F1 points at the trailing / future row with no result (H1 sim only, G1 off).
     */
    isCombo1FocusEmptyResult() {
        const sourceRows = this.sourceRows || [];
        let row = this.getSourceRowById(this.comboFocusRowId);
        if (!row && this.comboFocusRowIndex >= 0 && this.comboFocusRowIndex < sourceRows.length) {
            row = sourceRows[this.comboFocusRowIndex];
        }
        if (row) {
            return this.isEmptyResultRow(row);
        }
        const focusNum = this.parseRowId(this.comboFocusRowId);
        if (focusNum === null) {
            return false;
        }
        const latestValidRow = this.getLatestValidResultRow(sourceRows);
        if (!latestValidRow) {
            return true;
        }
        const latestNum = this.parseRowId(latestValidRow.id || latestValidRow.ID || '');
        if (latestNum === null) {
            return false;
        }
        return focusNum >= latestNum + 1;
    }

    /**
     * Build the current combo_1 focus, arrow, and styling state.
     */
    buildCombo1StyleContext() {
        const sourceRows = this.sourceRows || [];
        const fallbackRow = this.getLatestValidResultRow(sourceRows);
        let focusRow = this.getSourceRowById(this.comboFocusRowId);
        if (!focusRow && this.comboFocusRowIndex >= 0 && this.comboFocusRowIndex < sourceRows.length) {
            focusRow = sourceRows[this.comboFocusRowIndex];
        }
        if (!focusRow) {
            focusRow = fallbackRow;
        }
        const focusId = focusRow ? (focusRow.id || focusRow.ID || '') : this.comboFocusRowId;
        const targetIndex = focusRow ? sourceRows.findIndex(row => this.normalizeNumberKey(row.id || row.ID || '') === this.normalizeNumberKey(focusRow.id || focusRow.ID || '')) : -1;
        const targetRow = targetIndex >= 0 ? sourceRows[targetIndex] || null : null;
        const latestValidRow = fallbackRow;
        const latestIdNum = latestValidRow ? this.parseRowId(latestValidRow.id || latestValidRow.ID || '') : null;
        const typedFocusIdNum = this.parseRowId(focusId);
        const targetResult = targetRow ? (targetRow.result || targetRow.Result || '') : '';
        const targetNums = targetRow ? this.parseMainNums(targetResult) : [];
        const targetSpecial = targetRow ? this.parseSpecialPart(targetResult) : '';
        let targetNonexistText = '';
        if (targetRow) {
            const rawNonexist = String(targetRow.nonexist || targetRow.Nonexist || '').trim();
            if (rawNonexist) {
                targetNonexistText = rawNonexist;
            } else if (targetIndex >= 0 && this.nonexistCache && this.nonexistCache[targetIndex]) {
                targetNonexistText = String(this.nonexistCache[targetIndex].text || '').trim();
            }
        }
        const targetNonexistSet = new Set(
            (targetNonexistText && targetNonexistText !== 'N/A' ? this.parseNums(targetNonexistText) : []).map(num => String(num))
        );
        const h1Nums = this.parseNums(this.comboH1Text);
        const targetIsEmptyResult = targetRow ? this.isEmptyResultRow(targetRow) : this.isCombo1FocusEmptyResult();
        const targetHasResult = !!String(targetResult || '').trim();
        const showArrowsForTarget = !!this.comboG1Enabled && !targetIsEmptyResult;
        const useH1Sim = targetIsEmptyResult
            || (h1Nums.length > 0 && (!showArrowsForTarget || !targetHasResult));
        const arrowNums = targetIsEmptyResult
            ? h1Nums.slice(0, 5)
            : (showArrowsForTarget && targetHasResult ? targetNums : h1Nums.slice(0, 5));
        const arrowSet = new Set(arrowNums.map(num => String(num)));

        const freqWin = new Map();
        const aggNoteWin = new Set();
        const aggNonexistWin = new Set();
        let windowEnd = targetIndex < 0 ? sourceRows.length - 1 : Math.max(0, targetIndex - 1);
        if (targetIndex < 0 && typedFocusIdNum !== null && latestIdNum !== null && typedFocusIdNum >= latestIdNum + 1) {
            windowEnd = sourceRows.length - 1;
        }
        const windowStart = Math.max(0, windowEnd - 9);

        for (let index = windowStart; index <= windowEnd; index++) {
            const row = sourceRows[index] || {};
            const rowNums = this.parseMainNums(row.result || row.Result || '');
            for (const num of rowNums) {
                const key = String(num);
                freqWin.set(key, (freqWin.get(key) || 0) + 1);
            }

            const noteText = String(row.note || row.Note || '');
            if (noteText && noteText !== '?') {
                const noteParts = noteText.split(/\s{3,}/);
                for (const part of noteParts) {
                    const openBrace = part.indexOf('{');
                    const closeBrace = part.indexOf('}', openBrace + 1);
                    if (openBrace >= 0 && closeBrace > openBrace) {
                        const inside = part.substring(openBrace + 1, closeBrace);
                        for (const item of inside.split(',')) {
                            const key = this.normalizeNumberKey(item);
                            if (key) {
                                aggNoteWin.add(key);
                            }
                        }
                    }
                }
            }

            const nonexistText = String(row.nonexist || row.Nonexist || '');
            if (nonexistText && nonexistText !== 'N/A') {
                for (const item of nonexistText.split(',')) {
                    const key = this.normalizeNumberKey(item);
                    if (key) {
                        aggNonexistWin.add(key);
                    }
                }
            }
        }

        return {
            focusRow,
            focusId,
            targetIndex,
            targetId: targetRow ? (targetRow.id || targetRow.ID || '') : '',
            targetNums,
            targetSpecial,
            targetNonexistSet,
            arrowNums,
            arrowSet,
            showArrowsForTarget,
            targetIsEmptyResult,
            useH1Sim,
            h1Nums,
            freqWin,
            aggNoteWin,
            aggNonexistWin
        };
    }

    /**
     * Store a sheet-specific scroll position for the right pane.
     */
    setScrollPosition(sheetName, scrollTop, scrollLeft) {
        if (!sheetName) {
            return;
        }

        this.scrollPositions[sheetName] = {
            top: Math.max(0, Number(scrollTop) || 0),
            left: Math.max(0, Number(scrollLeft) || 0)
        };
        this.save();
    }

    /**
     * Get a saved scroll position for a sheet.
     */
    getScrollPosition(sheetName) {
        return this.scrollPositions && this.scrollPositions[sheetName] ? this.scrollPositions[sheetName] : { top: 0, left: 0 };
    }

    /**
     * Rebuild the combo_1 appear/special rows using the current F1/G1/H1 state.
     */
    buildCombo1RuntimeRows() {
        const sourceRows = this.sourceRows || [];
        const comboState = this.buildCombo1StyleContext();
        const latestValidRow = this.getLatestValidResultRow(sourceRows);
        const latestIdNum = latestValidRow ? this.parseRowId(latestValidRow.id || latestValidRow.ID || '') : null;
        const focusIdNum = this.parseRowId(comboState.focusId || '');
        const targetRow = comboState.focusRow || latestValidRow;

        const showArrowsForTarget = comboState.showArrowsForTarget;
        const useH1Sim = comboState.useH1Sim;
        const targetIndex = comboState.targetIndex >= 0 ? comboState.targetIndex : (targetRow ? sourceRows.findIndex(row => this.normalizeNumberKey(row.id || row.ID || '') === this.normalizeNumberKey(targetRow.id || targetRow.ID || '')) : -1);
        const targetRowIndex = targetIndex >= 0 ? targetIndex : sourceRows.length - 1;
        let freqEnd = showArrowsForTarget ? targetRowIndex : targetRowIndex - 1;

        if (targetIndex < 0 && focusIdNum !== null && latestIdNum !== null && focusIdNum >= latestIdNum + 1) {
            freqEnd = sourceRows.length - 1;
        }

        const startRow = 0;
        const freq = new Map();
        for (let rowIndex = startRow; rowIndex <= freqEnd; rowIndex++) {
            const row = sourceRows[rowIndex] || {};
            const nums = this.parseMainNums(row.result || row.Result || '');
            for (const num of nums) {
                const key = String(num);
                freq.set(key, (freq.get(key) || 0) + 1);
            }
        }

        if (useH1Sim) {
            for (const num of comboState.h1Nums.slice(0, 5)) {
                const key = String(num);
                freq.set(key, (freq.get(key) || 0) + 1);
            }
        }

        const specialCounts = new Map();
        const specialEnd = targetIndex >= 0 ? targetIndex : (targetRowIndex >= 0 ? targetRowIndex : sourceRows.length - 1);
        for (let rowIndex = startRow; rowIndex <= specialEnd; rowIndex++) {
            const row = sourceRows[rowIndex] || {};
            const special = this.parseSpecialPart(row.result || row.Result || '');
            if (special) {
                specialCounts.set(special, (specialCounts.get(special) || 0) + 1);
            }
        }

        const comboRows = [];
        for (let number = 1; number <= 35; number++) {
            const combo = String(number);
            const appear = freq.get(combo) || 0;
            comboRows.push({ combo, appear, arrow: '' });
        }
        const combo1Reach = new Map();
        for (let ci = 0; ci < comboRows.length; ci++) {
            const row = comboRows[ci];
            combo1Reach.set(
                row.combo,
                this.rowIndexWhenComboAppearReached(sourceRows, 0, freqEnd, 1, row.combo, row.appear)
            );
        }
        comboRows.sort((left, right) => {
            if (right.appear !== left.appear) {
                return right.appear - left.appear;
            }
            const ta = combo1Reach.get(left.combo);
            const tb = combo1Reach.get(right.combo);
            if (ta !== tb) {
                return ta - tb;
            }
            return Number(left.combo) - Number(right.combo);
        });

        const targetArrowSet = comboState.arrowSet;
        for (const row of comboRows) {
            if (targetArrowSet.has(this.normalizeNumberKey(row.combo))) {
                row.arrow = '⬆';
            }
        }

        const specialRows = [];
        for (const [special, count] of specialCounts.entries()) {
            specialRows.push({ special, count, arrow: '' });
        }
        const keysSpecialRt = new Set(specialCounts.keys());
        const special1Reach = this.buildSpecialReachRowMapOnePass(
            sourceRows,
            specialEnd + 1,
            specialCounts,
            keysSpecialRt
        );
        specialRows.sort((left, right) => {
            if (right.count !== left.count) {
                return right.count - left.count;
            }
            const ta = special1Reach.get(left.special);
            const tb = special1Reach.get(right.special);
            if (ta !== tb) {
                return ta - tb;
            }
            return String(left.special).localeCompare(String(right.special));
        });

        const targetSpecialKey = this.normalizeNumberKey(comboState.targetSpecial);
        for (const row of specialRows) {
            if (targetSpecialKey && this.normalizeNumberKey(row.special) === targetSpecialKey) {
                row.arrow = '⬆';
            }
        }

        return {
            comboState,
            comboRows,
            specialRows,
            latestId: comboState.targetId || (latestValidRow ? (latestValidRow.id || latestValidRow.ID || '') : '')
        };
    }

    /**
     * Find a source row by its id.
     */
    getSourceRowById(rawId) {
        const key = this.normalizeNumberKey(rawId);
        if (!key) {
            return null;
        }

        return (this.sourceRows || []).find(row => this.normalizeNumberKey(row.id || row.ID || '') === key) || null;
    }

    /**
     * Get color by frequency (inspired by Module5 highlighting pattern)
     * Higher frequency = different color, uses palette cycling
     */
    getColorByFrequency(num) {
        const freq = this.frequencyMap[num] || 0;
        if (freq === 0) return 'inherit';
        const colorIndex = (freq - 1) % this.colorPalette.length;
        return this.colorPalette[colorIndex];
    }

    /**
     * Render data table with frequency-based styling
     */
    renderTable(tableWrap) {
        if (tableWrap) {
            tableWrap.classList.remove('table-wrap--specialtracking');
        }
        if (tableWrap && typeof tableWrap.__specialTrackingCleanup === 'function') {
            try {
                tableWrap.__specialTrackingCleanup();
            } catch (eSt) {
                /* ignore */
            }
            tableWrap.__specialTrackingCleanup = null;
        }

        const sheet = this.sheets[this.activeSheet];
        if (!sheet) {
            tableWrap.innerHTML = '<div class="sheet-empty">Không có dữ liệu. Tải dữ liệu từ data.json</div>';
            return;
        }

        if (sheet.kind === 'specialtracking') {
            this.ensureSpecialTrackingFrames(sheet);
            tableWrap.classList.add('table-wrap--specialtracking');
            tableWrap.innerHTML = this.renderSpecialTrackingShell(sheet);
            this.wireSpecialTrackingUi(tableWrap, sheet);
            return;
        }

        if (sheet.kind === 'combo') {
            tableWrap.innerHTML = this.renderComboSheetHtml(sheet);
            if (this.activeSheet === 'combo_1') {
                this.wireCombo1HeaderControls();
            }
            return;
        }

        if (!sheet.data || sheet.data.length === 0) {
            tableWrap.innerHTML = '<div class="sheet-empty">Không có dữ liệu. Tải dữ liệu từ data.json</div>';
            return;
        }

        if (this.activeSheet === 'sheet1') {
            this.renderSourceSheet(tableWrap, sheet.data);
            return;
        }

        this.renderSourceSheet(tableWrap, sheet.data);
    }

    /**
     * Infer left-pane mode (pair1, pair2, modeq, triple1, ...) for a source row index.
     * Mirrors ok_left.html syncModeFromAnswerNums using the 11-line window ending at rowIndex.
     */
    inferModeForRowIndex(rowIndex) {
        const rows = this.getSourceSheetRows();
        if (rowIndex < 0 || rowIndex >= rows.length) {
            return null;
        }

        const focusRow = rows[rowIndex];
        const answerNums = this.parseMainNums(focusRow.result || focusRow.Result || '');
        if (answerNums.length < 5) {
            return null;
        }

        const windowStart = Math.max(0, rowIndex - 10);
        let pairLines = 0;
        let quadFound = false;
        let tripleFound = false;
        let singleLines = 0;

        const limit = Math.min(rowIndex - windowStart, 10);
        for (let offset = 0; offset < limit; offset++) {
            const lineRow = rows[windowStart + offset] || {};
            const nums = this.parseMainNums(lineRow.result || lineRow.Result || '');
            const matched = answerNums.filter(num => nums.includes(num));
            if (matched.length >= 4) {
                quadFound = true;
                break;
            }
            if (matched.length >= 3) {
                tripleFound = true;
                break;
            }
            if (matched.length >= 2) {
                pairLines++;
            }
            if (matched.length >= 1) {
                singleLines++;
            }
        }

        if (quadFound) {
            return 'quad1';
        }
        if (tripleFound) {
            return 'triple1';
        }
        if (pairLines >= 5) {
            return 'pair5';
        }
        if (pairLines >= 4) {
            return 'pair4';
        }
        if (pairLines >= 3) {
            return 'pair3';
        }
        if (pairLines === 2) {
            return 'pair2';
        }
        if (pairLines === 1) {
            return 'pair1';
        }
        if (singleLines > 0) {
            return 'modeq';
        }
        return null;
    }

    /**
     * Position + frequency signature của một số (specimen) trong cửa sổ 10 chuỗi kết thúc tại refRowIndex
     * (cùng logic cửa sổ với inferModeForRowIndex: các dòng windowStart .. refRowIndex-1).
     * frequency = tổng số lần xuất hiện specimen trong cửa sổ;
     * positions = multiset nhãn Chuỗi (mỗi lần xuất hiện trên một dòng → một phần tử, có thể lặp cùng số chuỗi),
     *   đã sort để so khớp (cùng f và cùng multiset vị trí ↔ cùng phân bố trên các chuỗi).
     */
    computePosnfreqSignature(rows, refRowIndex, specimenNum) {
        if (!Array.isArray(rows) || rows.length === 0) {
            return null;
        }
        if (!Number.isFinite(specimenNum) || specimenNum < 1 || specimenNum > 35) {
            return null;
        }
        if (refRowIndex < 0 || refRowIndex >= rows.length) {
            return null;
        }
        const windowStart = Math.max(0, refRowIndex - 10);
        const limit = Math.min(refRowIndex - windowStart, 10);
        if (limit <= 0) {
            return null;
        }
        const positions = [];
        let frequency = 0;
        for (let offset = 0; offset < limit; offset++) {
            const lineRow = rows[windowStart + offset] || {};
            const nums = this.parseMainNums(lineRow.result || lineRow.Result || '');
            let lineCount = 0;
            for (let k = 0; k < nums.length; k++) {
                if (nums[k] === specimenNum) {
                    lineCount++;
                }
            }
            if (lineCount > 0) {
                frequency += lineCount;
                const chuoiLabel = limit - offset;
                for (let t = 0; t < lineCount; t++) {
                    positions.push(chuoiLabel);
                }
            }
        }
        positions.sort((a, b) => a - b);
        return { frequency, positions: positions.slice() };
    }

    posnfreqPositionsKey(sig) {
        if (!sig || !Array.isArray(sig.positions)) {
            return '';
        }
        return sig.positions.join(',');
    }

    /**
     * @param {object|null} refSig — chữ ký đầy đủ cửa sổ 10 chuỗi của specimen trên kỳ mẫu (f + multiset nhãn Chuỗi)
     * @param {boolean} specimenStrict — true (Số): chỉ số specimen có cùng refSig;
     *                                   false (Mẫu): tồn tại m ∈ [1..35] có cùng refSig (lục giác có thể “đặt” lên m)
     */
    rowMatchesPosnfreqFilter(rows, rowIndex, specimenNum, refSig, specimenStrict) {
        const row = rows[rowIndex];
        if (!row || this.isEmptyResultRow(row)) {
            return false;
        }
        if (!refSig || !Number.isFinite(refSig.frequency)) {
            return false;
        }
        if (specimenStrict) {
            const sig = this.computePosnfreqSignature(rows, rowIndex, specimenNum);
            if (!sig || sig.frequency === 0) {
                return false;
            }
            return sig.frequency === refSig.frequency
                && this.posnfreqPositionsKey(sig) === this.posnfreqPositionsKey(refSig);
        }
        for (let m = 1; m <= 35; m++) {
            const sig = this.computePosnfreqSignature(rows, rowIndex, m);
            if (!sig || sig.frequency === 0) {
                continue;
            }
            if (sig.frequency === refSig.frequency
                && this.posnfreqPositionsKey(sig) === this.posnfreqPositionsKey(refSig)) {
                return true;
            }
        }
        return false;
    }

    /**
     * Mọi m ∈ [1..35] có chữ ký posnfreq trùng refSig trên rowIndex (đã sort tăng dần).
     */
    findAllPosnfreqMatchingNumbers(rows, rowIndex, refSig) {
        if (!refSig || !Number.isFinite(refSig.frequency)) {
            return [];
        }
        const list = rows || this.getSourceSheetRows();
        if (!Array.isArray(list) || rowIndex < 0 || rowIndex >= list.length) {
            return [];
        }
        const out = [];
        for (let m = 1; m <= 35; m++) {
            const sig = this.computePosnfreqSignature(list, rowIndex, m);
            if (!sig || sig.frequency === 0) {
                continue;
            }
            if (sig.frequency === refSig.frequency
                && this.posnfreqPositionsKey(sig) === this.posnfreqPositionsKey(refSig)) {
                out.push(m);
            }
        }
        return out;
    }

    /**
     * Số nhỏ nhất m sao cho chữ ký posnfreq của m trên rowIndex khớp refSig (dùng cho viền lục giác Mẫu).
     */
    findPosnfreqMatchingNumber(rows, rowIndex, refSig) {
        const all = this.findAllPosnfreqMatchingNumbers(rows, rowIndex, refSig);
        return all.length ? all[0] : null;
    }

    /**
     * Row indices on sheet1 whose inferred mode matches the filter mode.
     */
    rowHasCyanDateBand(rows, rowIndex) {
        return this.shouldHighlightDateByPairWindow(rows, rowIndex);
    }

    /**
     * Cached row indices with cyan date band (#00b0f0) for dateband filter mode.
     */
    ensureDatebandFilterIndicesCache() {
        const rows = this.getSourceSheetRows();
        if (this.datebandFilterIndicesCache && this.datebandFilterIndicesCacheRowLen === rows.length) {
            return this.datebandFilterIndicesCache;
        }

        const indices = [];
        for (let i = 0; i < rows.length; i++) {
            if (!this.isEmptyResultRow(rows[i]) && this.rowHasCyanDateBand(rows, i)) {
                indices.push(i);
            }
        }
        this.datebandFilterIndicesCache = indices;
        this.datebandFilterIndicesCacheRowLen = rows.length;
        this.datebandRowDistCache = null;
        this.datebandRowDistCacheRowLen = 0;
        return indices;
    }

    /**
     * x: groups for a dateband row — only from pair lines shown in the 10-row window
     * (visible pair + match current result), using window pair_to_ids distances.
     */
    computeDatebandNoteDistancesForRow(rowIndex, rows, pairToIds) {
        const dists = new Set();
        const list = rows || this.getSourceSheetRows();
        if (!this.rowHasCyanDateBand(list, rowIndex)) {
            return dists;
        }

        const currentRow = list[rowIndex] || {};
        const currentNums = this.parseMainNums(currentRow.result || currentRow.Result || '');
        const rid = this.normalizeNumberKey(currentRow.id || currentRow.ID || '');
        if (!rid || currentNums.length !== 5 || rowIndex < 10) {
            return dists;
        }

        const windowRows = list.slice(Math.max(0, rowIndex - 10), rowIndex);
        const visiblePairs = this.computePairsForRows(windowRows);
        if (!visiblePairs.length) {
            return dists;
        }

        const map = pairToIds || {};
        for (let p = 0; p < visiblePairs.length; p++) {
            const a = visiblePairs[p][0];
            const b = visiblePairs[p][1];
            if (!this.pairExists(currentNums, a, b)) {
                continue;
            }
            const key = `${a},${b}`;
            const arr = map[key] && map[key][rid];
            if (!Array.isArray(arr) || !arr.length) {
                continue;
            }
            for (let i = 0; i < arr.length; i++) {
                const dist = parseInt(arr[i], 10);
                if (dist >= 1 && dist <= 10) {
                    dists.add(dist);
                }
            }
        }
        return dists;
    }

    ensureDatebandRowDistCache() {
        const rows = this.getSourceSheetRows();
        if (this.datebandRowDistCache && this.datebandRowDistCacheRowLen === rows.length) {
            return this.datebandRowDistCache;
        }

        const cache = new Array(rows.length);
        for (let i = 0; i < rows.length; i++) {
            cache[i] = new Set();
        }
        const pairToIds = {};
        this.accumulatePairToIdsFromRowWindows(rows, pairToIds);
        const base = this.ensureDatebandFilterIndicesCache();
        for (let b = 0; b < base.length; b++) {
            const i = base[b];
            cache[i] = this.computeDatebandNoteDistancesForRow(i, rows, pairToIds);
        }
        this.datebandRowDistCache = cache;
        this.datebandRowDistCacheRowLen = rows.length;
        return cache;
    }

    /**
     * Dateband row belongs to group x: when a visible window pair for that row maps to distance x.
     */
    rowMatchesDatebandNoteDistFilter(rowIndex, dist) {
        if (!Number.isFinite(dist) || dist < 1 || dist > 10) {
            return false;
        }
        const cache = this.ensureDatebandRowDistCache();
        const set = cache[rowIndex];
        return set ? set.has(dist) : false;
    }

    getFilterMatchingIndices(mode, filterOptions = null) {
        const indices = [];
        const rows = this.getSourceSheetRows();
        if (mode === 'nonexist') {
            const opts = filterOptions || {};
            const numRaw = opts.num;
            const num = (numRaw === null || numRaw === undefined || numRaw === '')
                ? null
                : parseInt(numRaw, 10);
            const colors = Array.isArray(opts.colors)
                ? opts.colors.filter((c) => ['green', 'red', 'purple', 'yellow'].includes(c))
                : [];
            const styles = Array.isArray(opts.styles) ? opts.styles : [];

            if (colors.length === 0) {
                return indices;
            }

            if (num === null || !Number.isFinite(num) || num < 1 || num > 35) {
                for (let i = 0; i < rows.length; i++) {
                    if (!this.isEmptyResultRow(rows[i])
                        && this.rowMatchesNonexistColorFilter(i, colors, styles, null)) {
                        indices.push(i);
                    }
                }
                return indices;
            }

            for (let i = 0; i < rows.length; i++) {
                if (this.rowMatchesNonexistColorFilter(i, colors, styles, num)) {
                    indices.push(i);
                }
            }
            return indices;
        }

        const noteTags = Array.isArray((filterOptions || {}).noteTags)
            ? (filterOptions.noteTags || []).filter((n) => Number.isFinite(n) && n >= 1 && n <= 10)
            : [];

        if (mode === 'dateband') {
            const base = this.ensureDatebandFilterIndicesCache();
            const distFilter = noteTags.length > 0 ? noteTags[0] : null;
            for (let b = 0; b < base.length; b++) {
                const i = base[b];
                if (distFilter !== null && !this.rowMatchesDatebandNoteDistFilter(i, distFilter)) {
                    continue;
                }
                indices.push(i);
            }
            return indices;
        }

        if (mode === 'connection') {
            const base = this.ensureConnectionFilterIndicesCache();
            if (noteTags.length === 0) {
                return base.slice();
            }
            const out = [];
            for (let b = 0; b < base.length; b++) {
                const i = base[b];
                if (this.rowMatchesNoteTagFilter(i, noteTags)) {
                    out.push(i);
                }
            }
            return out;
        }

        if (mode === 'intersection') {
            const o = filterOptions || {};
            const kind = o.intersectionKind === 'nearintersect' ? 'nearintersect' : 'intersect';
            let thX = parseInt(o.intersectionMinFreqA, 10);
            let thY = parseInt(o.intersectionMinFreqB, 10);
            if (!Number.isFinite(thX)) {
                thX = 2;
            }
            if (!Number.isFinite(thY)) {
                thY = 2;
            }
            thX = Math.min(7, Math.max(2, thX));
            thY = Math.min(7, Math.max(2, thY));
            const rawOpA = String(o.intersectionFreqOpA || '').trim();
            const rawOpB = String(o.intersectionFreqOpB || '').trim();
            const opA = rawOpA === '=' || rawOpA === '<=' || rawOpA === '>=' ? rawOpA : '>=';
            const opB = rawOpB === '=' || rawOpB === '<=' || rawOpB === '>=' ? rawOpB : '>=';
            for (let i = 0; i < rows.length; i++) {
                if (this.isEmptyResultRow(rows[i])) {
                    continue;
                }
                if (this.rowMatchesIntersectionSubmitWindow(rows, i, kind, thX, thY, opA, opB)) {
                    indices.push(i);
                }
            }
            return indices;
        }

        if (mode === 'posnfreq') {
            const o = filterOptions || {};
            const specimen = parseInt(o.specimenNum, 10);
            if (!Number.isFinite(specimen) || specimen < 1 || specimen > 35) {
                return indices;
            }
            const refRow = Number.isFinite(o.refRowIndex) ? o.refRowIndex : -1;
            const specimenStrict = !!o.specimenStrict;
            let refSig = o.refSignature;
            if (!refSig && refRow >= 0) {
                refSig = this.computePosnfreqSignature(rows, refRow, specimen);
            }
            for (let i = 0; i < rows.length; i++) {
                if (this.rowMatchesPosnfreqFilter(rows, i, specimen, refSig, specimenStrict)) {
                    indices.push(i);
                }
            }
            return indices;
        }

        for (let i = 0; i < rows.length; i++) {
            if (this.isEmptyResultRow(rows[i])) {
                continue;
            }
            if (this.inferModeForRowIndex(i) === mode) {
                if (noteTags.length > 0 && !this.rowMatchesNoteTagFilter(i, noteTags)) {
                    continue;
                }
                indices.push(i);
            }
        }
        return indices;
    }

    /**
     * Cửa sổ 10 chuỗi trước kỳ `rowIndex` (cùng inferMode / posnfreq): nhãn Chuỗi 1 = sát đáp án, L = xa nhất.
     * `nearintersect`: mỗi số phải xuất hiện trên các chuỗi tạo một dải nhãn liên tiếp; hai dải rời nhau và kề (max(A)+1=min(B) hoặc ngược lại), không chỉ “có một cặp nhãn |i−j|=1”.
     * Cặp (a,b) theo thứ tự 5 số submit: `freq[a]` so với `threshX` theo `opX` (`>=` | `=` | `<=`); `freq[b]` so với `threshY` theo `opY` (tổng lần xuất hiện trong cửa sổ tối đa 10 chuỗi).
     * @param {object[]} rows
     * @param {number} rowIndex
     * @param {'intersect' | 'nearintersect'} kind
     * @param {number} [threshX=2]
     * @param {number} [threshY=2]
     * @param {string} [opX='>=']
     * @param {string} [opY='>=']
     */
    rowMatchesIntersectionSubmitWindow(rows, rowIndex, kind, threshX = 2, threshY = 2, opX = '>=', opY = '>=') {
        const row = rows[rowIndex];
        if (!row) {
            return false;
        }
        const answerNums = this.parseMainNums(row.result || row.Result || '');
        if (answerNums.length < 5) {
            return false;
        }
        const windowStart = Math.max(0, rowIndex - 10);
        const limit = Math.min(rowIndex - windowStart, 10);
        if (limit <= 0) {
            return false;
        }

        /** @type {{ label: number, nums: number[]}[]} */
        const lines = [];
        for (let offset = 0; offset < limit; offset++) {
            const lineRow = rows[windowStart + offset] || {};
            const nums = this.parseMainNums(lineRow.result || lineRow.Result || '');
            const label = limit - offset;
            lines.push({ label, nums });
        }

        const freq = new Array(36).fill(0);
        for (let li = 0; li < lines.length; li++) {
            const nums = lines[li].nums;
            for (let k = 0; k < nums.length; k++) {
                const n = nums[k];
                if (n >= 1 && n <= 35) {
                    freq[n]++;
                }
            }
        }

        const hasNum = (nums, x) => nums.indexOf(x) !== -1;
        const labelsForNum = (x) => {
            const s = new Set();
            for (let li = 0; li < lines.length; li++) {
                if (hasNum(lines[li].nums, x)) {
                    s.add(lines[li].label);
                }
            }
            return s;
        };
        /** Tập nhãn chuỗi (số nguyên liên tiếp 1..L) phải là dải liền — không được {1,10} mà thiếu 2..9. */
        const labelsFormContiguousRun = (labelSet) => {
            if (labelSet.size <= 1) {
                return true;
            }
            const arr = Array.from(labelSet).sort((u, v) => u - v);
            for (let k = 1; k < arr.length; k++) {
                if (arr[k] !== arr[k - 1] + 1) {
                    return false;
                }
            }
            return true;
        };
        /** Hai dải [lo,hi] rời nhau và chạm cạnh: hi(A)+1 === lo(B) hoặc hi(B)+1 === lo(A). */
        const twoRunsTouchAdjacent = (sa, sb) => {
            const aLo = Math.min(...sa);
            const aHi = Math.max(...sa);
            const bLo = Math.min(...sb);
            const bHi = Math.max(...sb);
            if (aHi < bLo) {
                return aHi + 1 === bLo;
            }
            if (bHi < aLo) {
                return bHi + 1 === aLo;
            }
            return false;
        };

        const freqMatchesOp = (freq, threshold, op) => {
            if (op === '=') {
                return freq === threshold;
            }
            if (op === '<=') {
                return freq <= threshold;
            }
            return freq >= threshold;
        };

        for (let ai = 0; ai < answerNums.length; ai++) {
            for (let bi = ai + 1; bi < answerNums.length; bi++) {
                const a = answerNums[ai];
                const b = answerNums[bi];
                if (a === b || !freqMatchesOp(freq[a], threshX, opX) || !freqMatchesOp(freq[b], threshY, opY)) {
                    continue;
                }

                if (kind === 'nearintersect') {
                    let sharedLine = false;
                    for (let li = 0; li < lines.length; li++) {
                        const nums = lines[li].nums;
                        if (hasNum(nums, a) && hasNum(nums, b)) {
                            sharedLine = true;
                            break;
                        }
                    }
                    if (sharedLine) {
                        continue;
                    }
                    const Sa = labelsForNum(a);
                    const Sb = labelsForNum(b);
                    if (!labelsFormContiguousRun(Sa) || !labelsFormContiguousRun(Sb)) {
                        continue;
                    }
                    const arrA = Array.from(Sa);
                    const arrB = Array.from(Sb);
                    if (twoRunsTouchAdjacent(arrA, arrB)) {
                        return true;
                    }
                    continue;
                }

                /** intersect: ∃ chuỗi M chứa cả a,b; một số có mặt ở chuỗi nhãn > M, số kia ở nhãn < M. */
                const mainLabels = [];
                for (let li = 0; li < lines.length; li++) {
                    const nums = lines[li].nums;
                    if (hasNum(nums, a) && hasNum(nums, b)) {
                        mainLabels.push(lines[li].label);
                    }
                }
                if (mainLabels.length === 0) {
                    continue;
                }
                for (let mi = 0; mi < mainLabels.length; mi++) {
                    const M = mainLabels[mi];
                    let aUp = false;
                    let aLo = false;
                    let bUp = false;
                    let bLo = false;
                    for (let li = 0; li < lines.length; li++) {
                        const lab = lines[li].label;
                        if (lab === M) {
                            continue;
                        }
                        const nums = lines[li].nums;
                        if (lab > M) {
                            if (hasNum(nums, a)) {
                                aUp = true;
                            }
                            if (hasNum(nums, b)) {
                                bUp = true;
                            }
                        } else if (lab < M) {
                            if (hasNum(nums, a)) {
                                aLo = true;
                            }
                            if (hasNum(nums, b)) {
                                bLo = true;
                            }
                        }
                    }
                    if ((aUp && bLo) || (bUp && aLo)) {
                        return true;
                    }
                }
            }
        }
        return false;
    }

    /**
     * Note có "connection": một số xuất hiện trong ≥2 cặp `{a,b}` (ví dụ 6 trong `{6,23}` và `{4,6}`).
     * @param {string} noteText
     * @returns {boolean}
     */
    noteTextHasConnectionPairing(noteText) {
        const text = String(noteText || '');
        if (text.indexOf('{') === -1 || text.indexOf(',') === -1 || text.indexOf('}') === -1) {
            return false;
        }
        const re = /\{(\d+)\s*,\s*(\d+)\}/g;
        /** @type {Map<number, Set<number>>} */
        const numToPairIdx = new Map();
        let pairIndex = 0;
        let m;
        while ((m = re.exec(text)) !== null) {
            const a = parseInt(m[1], 10);
            const b = parseInt(m[2], 10);
            if (!Number.isFinite(a) || !Number.isFinite(b)) {
                continue;
            }
            for (const n of [a, b]) {
                let set = numToPairIdx.get(n);
                if (!set) {
                    set = new Set();
                    numToPairIdx.set(n, set);
                }
                set.add(pairIndex);
            }
            pairIndex++;
        }
        for (const s of numToPairIdx.values()) {
            if (s.size >= 2) {
                return true;
            }
        }
        return false;
    }

    /**
     * Danh sách chỉ số dòng có note connection — cache theo noteCache + độ dài sheet (tránh quét lặp).
     * @returns {number[]}
     */
    ensureConnectionFilterIndicesCache() {
        const rows = this.getSourceSheetRows();
        const n = rows.length;
        const nc = this.noteCache;
        if (
            this._connectionFilterIndicesCache
            && this._connectionFilterIndicesCacheRowLen === n
            && this._connectionFilterNoteCacheRef === nc
        ) {
            return this._connectionFilterIndicesCache;
        }
        const indices = [];
        for (let i = 0; i < n; i++) {
            if (this.isEmptyResultRow(rows[i])) {
                continue;
            }
            const noteText = this.getNoteTextForRowFilter(i);
            if (this.noteTextHasConnectionPairing(noteText)) {
                indices.push(i);
            }
        }
        this._connectionFilterIndicesCache = indices;
        this._connectionFilterIndicesCacheRowLen = n;
        this._connectionFilterNoteCacheRef = nc;
        return this._connectionFilterIndicesCache;
    }

    /**
     * Combined computed + raw note text for filtering.
     */
    getNoteTextForRowFilter(rowIndex) {
        const rows = this.getSourceSheetRows();
        const row = rows[rowIndex];
        if (!row) {
            return '';
        }
        const meta = this.getComputedNoteMeta(rowIndex, row);
        const computed = (meta.text && meta.text !== '?') ? String(meta.text) : '';
        const raw = String(row.note || row.Note || '').trim();
        return [computed, raw].filter(Boolean).join(' ');
    }

    /**
     * Note contains distance marker N:{...} (same shape as checknote notePat).
     */
    noteContainsDistColon(noteText, dist) {
        if (!noteText || dist < 1 || dist > 10) {
            return false;
        }
        const pat = new RegExp(`(?:^|[^0-9])${dist}\\s*:\\s*\\{`);
        return pat.test(noteText);
    }

    /**
     * Row note must contain every selected distance marker (AND).
     */
    rowMatchesNoteTagFilter(rowIndex, noteTags) {
        const tags = Array.isArray(noteTags) ? noteTags : [];
        if (tags.length === 0) {
            return true;
        }
        const noteText = this.getNoteTextForRowFilter(rowIndex);
        if (!noteText) {
            return false;
        }
        for (let i = 0; i < tags.length; i++) {
            if (!this.noteContainsDistColon(noteText, tags[i])) {
                return false;
            }
        }
        return true;
    }

    /**
     * Green (xanh lá) display kinds in the nonexist column.
     */
    isGreenNonexistDisplayKind(kind) {
        return kind === 'green'
            || kind === 'green-italic'
            || kind === 'green-ul'
            || kind === 'green-strike';
    }

    /**
     * Whether a green nonexist number matches one selected style token.
     * Each letter maps to one green variant only (B ≠ italic/underline/strike).
     */
    nonexistGreenKindMatchesStyle(kind, style) {
        if (!this.isGreenNonexistDisplayKind(kind)) {
            return false;
        }
        if (style === 'bold') {
            return kind === 'green';
        }
        if (style === 'italic') {
            return kind === 'green-italic';
        }
        if (style === 'underline') {
            return kind === 'green-ul';
        }
        if (style === 'strikethrough') {
            return kind === 'green-strike';
        }
        return false;
    }

    /**
     * True when the row's nonexist column has at least one green-highlighted number.
     */
    rowHasAnyGreenNonexistNumber(rowIndex) {
        const entries = this.ensureNonexistGreenFilterCache()[rowIndex];
        return Array.isArray(entries) && entries.length > 0;
    }

    /**
     * Match row nonexist: optional specificNum, or any green number matching a selected style (OR).
     */
    rowMatchesNonexistGreenStyleFilter(rowIndex, styles, specificNum = null) {
        return this.rowMatchesNonexistColorFilter(rowIndex, ['green'], styles, specificNum);
    }

    /**
     * Match row by selected nonexist colors. Green uses B/I/U/S toggles; red/purple/yellow are color-only.
     */
    rowMatchesNonexistColorFilter(rowIndex, colors, styles, specificNum = null) {
        const colorList = Array.isArray(colors) ? colors : [];
        if (colorList.length === 0) {
            return false;
        }

        const row = this.getSourceSheetRows()[rowIndex];
        if (!row || this.isEmptyResultRow(row)) {
            return false;
        }

        const targetNum = specificNum === null || specificNum === undefined || specificNum === ''
            ? null
            : parseInt(specificNum, 10);
        if (targetNum !== null && (!Number.isFinite(targetNum) || targetNum < 1 || targetNum > 35)) {
            return false;
        }

        const styleList = Array.isArray(styles) ? styles : [];
        const entries = this.ensureNonexistDisplayEntriesCache()[rowIndex] || [];

        for (let i = 0; i < entries.length; i++) {
            const entry = entries[i];
            if (targetNum !== null && entry.num !== targetNum) {
                continue;
            }
            if (!colorList.includes(entry.color)) {
                continue;
            }
            if (entry.color === 'green') {
                if (styleList.length === 0) {
                    continue;
                }
                for (let j = 0; j < styleList.length; j++) {
                    if (this.nonexistGreenKindMatchesStyle(entry.kind, styleList[j])) {
                        return true;
                    }
                }
                continue;
            }
            return true;
        }
        return false;
    }

    /**
     * True when nonexist lists the number in green and it matches at least one selected style (OR).
     */
    rowMatchesNonexistStyleFilter(rowIndex, num, styles) {
        return this.rowMatchesNonexistGreenStyleFilter(rowIndex, styles, num);
    }

    /**
     * Display kind for one number on a source row (nonexist column), for filter-popup #005000 observation.
     */
    getNonexistDisplayKindForNumberOnSourceRow(rowIndex, num) {
        const rows = this.getSourceSheetRows();
        const row = rows[rowIndex];
        if (!row) {
            return '';
        }
        if (!this.nonexistCache || this.nonexistCache.length !== rows.length) {
            this.nonexistCache = this.buildNonexistFromRows(rows);
        }
        const nonexistMeta = this.getNonexistMetaForSourceRow(rowIndex, row);
        const nx = String(nonexistMeta.text || '').trim();
        if (!nx || nx === 'N/A') {
            return '';
        }
        if (this.parseNums(nx).indexOf(num) === -1) {
            return '';
        }
        const res = row.result || row.Result || '';
        return this.getNonexistDisplayKindForNumber(rowIndex, num, nx, res, null);
    }

    /**
     * True when y is yellow on rowIndex and ∃ sliding window [w..w+9] containing rowIndex
     * with every row having 5 mains and y ∈ nonexist(bottom row w+9).
     * Bỏ qua cửa sổ có đáy e === rowIndex: khi đó y luôn ∈ nonexist(e) nếu y nằm trong ô nonexist của dòng đó
     * (boost 1.5em chỉ có ý nghĩa khi đáy cửa sổ nằm dưới dòng đang xét).
     */
    nonexistYellowHasBoostInSlidingWindowTen(rowIndex, y) {
        const rows = this.getSourceSheetRows();
        const n = rows.length;
        if (!this.nonexistCache || this.nonexistCache.length !== n) {
            this.nonexistCache = this.buildNonexistFromRows(rows);
        }
        const cache = this.nonexistCache;
        const wMin = Math.max(10, rowIndex - 9);
        const wMax = Math.min(rowIndex, n - 10);
        if (wMin > wMax) {
            return false;
        }
        if (this.getNonexistDisplayKindForNumberOnSourceRow(rowIndex, y) !== 'yellow') {
            return false;
        }
        for (let w = wMin; w <= wMax; w++) {
            const e = w + 9;
            if (e === rowIndex) {
                continue;
            }
            let ok = true;
            for (let i = w; i <= e; i++) {
                const r = rows[i];
                if (!r || this.isEmptyResultRow(r) || this.parseMainNums(r.result || r.Result || '').length !== 5) {
                    ok = false;
                    break;
                }
            }
            if (!ok) {
                continue;
            }
            const bottom = String(cache[e].text || '').trim();
            if (!bottom || bottom === 'N/A') {
                continue;
            }
            if (this.parseNums(bottom).indexOf(y) === -1) {
                continue;
            }
            return true;
        }
        return false;
    }

    /**
     * Số vàng y tại rowIndex: y xuất hiện trong 5 số chính của ít nhất một kỳ trong (rowIndex, rowIndex+10]
     * (mỗi kỳ đó phải đủ 5 số chính). Không yêu cầu cùng cửa sổ 10 với boost — ví dụ 00053 + 10 kỳ sau tới 00063.
     */
    nonexistYellowCalledInNextTenRowsAfter(rowIndex, y) {
        const rows = this.getSourceSheetRows();
        const n = rows.length;
        const hi = Math.min(rowIndex + 10, n - 1);
        for (let k = rowIndex + 1; k <= hi; k++) {
            const r = rows[k];
            if (!r || this.isEmptyResultRow(r)) {
                continue;
            }
            const mains = this.parseMainNums(r.result || r.Result || '');
            if (mains.length !== 5) {
                continue;
            }
            if (mains.indexOf(y) !== -1) {
                return true;
            }
        }
        return false;
    }

    /**
     * Yellow y trên rowIndex: boost như nonexistYellowHasBoostInSlidingWindowTen
     * và "gọi lại" = trúng trong 5 số chính của một trong 10 kỳ liền sau dòng đó.
     */
    nonexistYellowHasBoostAndCalledInSlidingWindowTen(rowIndex, y) {
        return (
            this.nonexistYellowHasBoostInSlidingWindowTen(rowIndex, y) &&
            this.nonexistYellowCalledInNextTenRowsAfter(rowIndex, y)
        );
    }

    /**
     * Filter popup nonexist #005000: viền theo số vàng — boost (đáy > dòng, cửa 10) + gọi lại (10 kỳ sau).
     * @returns {'none'|'green'|'yellow'|'red'}
     */
    getNonexist005000BoostBorderKindForRow(rowIndex) {
        const rows = this.getSourceSheetRows();
        const n = rows.length;
        if (rowIndex < 10 || rowIndex >= n) {
            return 'none';
        }
        const row = rows[rowIndex];
        if (!row || this.isEmptyResultRow(row)) {
            return 'none';
        }
        if (!this.nonexistCache || this.nonexistCache.length !== n) {
            this.nonexistCache = this.buildNonexistFromRows(rows);
        }
        const nonexistMeta = this.getNonexistMetaForSourceRow(rowIndex, row);
        const nx = String(nonexistMeta.text || '').trim();
        if (!nx || nx === 'N/A') {
            return 'none';
        }
        const candidates = this.parseNums(nx);
        const yellows = [];
        for (let i = 0; i < candidates.length; i++) {
            const y = candidates[i];
            if (this.getNonexistDisplayKindForNumberOnSourceRow(rowIndex, y) === 'yellow') {
                yellows.push(y);
            }
        }
        if (yellows.length === 0) {
            return 'none';
        }
        let anyFull = false;
        let allFull = true;
        for (let j = 0; j < yellows.length; j++) {
            const ok = this.nonexistYellowHasBoostAndCalledInSlidingWindowTen(rowIndex, yellows[j]);
            if (ok) {
                anyFull = true;
            } else {
                allFull = false;
            }
        }
        if (!anyFull) {
            return 'red';
        }
        if (allFull) {
            return 'green';
        }
        return 'yellow';
    }

    /**
     * Toàn sheet: các dòng khớp lọc nonexist (num+colors+styles) + thống kê viền quan sát #005000 (boost+gọi).
     */
    computeNonexist005000DatasetObservation(colors, styles, num) {
        const rows = this.getSourceSheetRows();
        const n = rows.length;
        let totalFilter = 0;
        let greenYellowBorder = 0;
        const missIds = [];
        for (let r = 10; r < n; r++) {
            if (!this.rowMatchesNonexistColorFilter(r, colors, styles, num)) {
                continue;
            }
            totalFilter++;
            const kind = this.getNonexist005000BoostBorderKindForRow(r);
            if (kind === 'green' || kind === 'yellow') {
                greenYellowBorder++;
            } else {
                missIds.push(String(rows[r].id || rows[r].ID || r));
            }
        }
        const pctDataset = totalFilter ? (100 * greenYellowBorder) / totalFilter : 0;
        return {
            totalFilter,
            greenYellowBorder,
            missIds,
            pctDataset
        };
    }

    countNonexist005000BorderKindsInIndices(indices) {
        const out = { green: 0, yellow: 0, red: 0, none: 0 };
        if (!Array.isArray(indices)) {
            return out;
        }
        for (let i = 0; i < indices.length; i++) {
            const k = this.getNonexist005000BoostBorderKindForRow(indices[i]);
            if (k === 'green') {
                out.green++;
            } else if (k === 'yellow') {
                out.yellow++;
            } else if (k === 'red') {
                out.red++;
            } else {
                out.none++;
            }
        }
        return out;
    }

    /**
     * Answer popup open + Submit OFF: dim result/note and show empty-style nonexist on focus row.
     */
    setAnswerPopupFocusMask(opts) {
        const o = opts || {};
        const open = !!o.open;
        const rowIndex = Number.isFinite(o.rowIndex) ? o.rowIndex : -1;
        const submitOn = !!o.submitOn;
        this.answerPopupFocusMask = {
            active: open && rowIndex >= 0 && !submitOn,
            rowIndex: open ? rowIndex : -1
        };
    }

    shouldAnswerPopupMaskSheet1Row(rowIndex) {
        const m = this.answerPopupFocusMask || {};
        return !!(m.active && m.rowIndex === rowIndex);
    }

    /**
     * Nonexist HTML for one source row (respects Answer-popup focus mask = empty-result styling).
     */
    renderSourceRowNonexistCellHtml(rowIndex, row) {
        const maskRow = this.shouldAnswerPopupMaskSheet1Row(rowIndex);
        const source = row || {};
        const result = source.result || source.Result || '';
        if (maskRow) {
            const emptyRow = Object.assign({}, source, { result: '', Result: '' });
            const meta = this.getNonexistMetaForSourceRow(rowIndex, emptyRow);
            return this.renderNonexistHtml(rowIndex, meta.text, '');
        }
        const meta = this.getNonexistMetaForSourceRow(rowIndex, source);
        return this.renderNonexistHtml(rowIndex, meta.text, result);
    }

    scheduleApplyAnswerPopupFocusMask(tableWrap) {
        if (this._answerPopupMaskApplyRaf) {
            return;
        }
        const wrap = tableWrap || document.getElementById('tableWrap');
        this._answerPopupMaskApplyRaf = requestAnimationFrame(() => {
            this._answerPopupMaskApplyRaf = 0;
            this.applyAnswerPopupFocusMaskToDom(wrap);
        });
    }

    applyAnswerPopupFocusMaskToDom(tableWrap, options = {}) {
        if (!tableWrap || this.activeSheet !== 'sheet1') {
            return;
        }
        if (options.reset) {
            this._answerPopupMaskAppliedRow = -1;
        }

        const m = this.answerPopupFocusMask || {};
        const prevIdx = this._answerPopupMaskAppliedRow;
        const nextIdx = m.active ? m.rowIndex : -1;

        if (prevIdx >= 0 && prevIdx !== nextIdx) {
            this.setAnswerPopupFocusMaskOnRowDom(tableWrap, prevIdx, false);
        }
        if (nextIdx >= 0 && nextIdx !== prevIdx) {
            this.setAnswerPopupFocusMaskOnRowDom(tableWrap, nextIdx, true);
        }

        this._answerPopupMaskAppliedRow = nextIdx;
    }

    setAnswerPopupFocusMaskOnRowDom(tableWrap, rowIndex, masked) {
        const tr = tableWrap.querySelector(`tbody tr[data-idx="${rowIndex}"]`);
        if (!tr) {
            return;
        }
        tr.classList.toggle('answer-popup-focus-masked', masked);
        const nonexistCell = tr.querySelector('td.cell-nonexist');
        const row = (this.dataRows || [])[rowIndex];
        if (!nonexistCell || !row) {
            return;
        }
        nonexistCell.classList.toggle('answer-popup-focus-nonexist', masked);
        nonexistCell.innerHTML = this.renderSourceRowNonexistCellHtml(rowIndex, row);
    }

    /**
     * Render the raw five-column source sheet.
     * @param {object} [options]
     * @param {number[]} [options.indices] - subset of row indices to render
     * @param {number} [options.highlightIdx] - active row highlight (filter popup)
     * @param {boolean} [options.bindKeyboard=true]
     * @param {boolean} [options.applyWindowSelection=true]
     */
    renderSourceSheet(tableWrap, rows, options = {}) {
        const bindKeyboard = options.bindKeyboard !== false;
        const applyWindowSelection = options.applyWindowSelection !== false;
        const highlightIdx = typeof options.highlightIdx === 'number' ? options.highlightIdx : -1;

        if (bindKeyboard) {
            this.bindSourceSheetKeyboardNavigation(tableWrap);
        }

        let html = '<table class="sheet-data-table"><thead><tr><th>date</th><th>id</th><th>result</th><th>note</th><th>nonexist</th></tr></thead><tbody>';

        const displayRows = rows || [];
        const rowIndices = Array.isArray(options.indices)
            ? options.indices.filter(i => i >= 0 && i < displayRows.length)
            : displayRows.map((_, i) => i);
        const prevRecallFoldStats = this.computePrevPeriodRecallFoldStats(displayRows, rowIndices);
        const prevRecallFoldPctLabel = this.formatPrevPeriodRecallFoldPct(prevRecallFoldStats);

        for (const i of rowIndices) {
            const row = displayRows[i];
            const date = row.date || row.Date || '';
            const id = row.id || row.ID || '';
            const result = row.result || row.Result || '';
            const isEmptyResultRow = this.isEmptyResultRow(row);
            const noteMeta = isEmptyResultRow
                ? { text: '', highlightYellow: false }
                : this.getComputedNoteMeta(i, row);
            const idBg = this.getIdBackgroundByFrequency(id);
            const dateBg = this.shouldHighlightDateByPairWindow(displayRows, i) ? ' style="background:#00b0f0;color:#000;font-weight:bold;"' : '';

            let resultHtml = this.highlightResultByFrequency(result);
            let noteHtml = this.renderNoteHtml(noteMeta.text, noteMeta.highlightYellow);
            const noteStyle = noteMeta.highlightYellow ? ' style="background:#ff0;"' : '';
            const nonexistMeta = this.getNonexistMetaForSourceRow(i, row);
            let nonexistHtml = this.renderNonexistHtml(i, nonexistMeta.text, result);
            const idStyle = idBg ? ` style="background:${idBg};"` : '';
            const activeClass = highlightIdx === i ? ' filter-popup-row-active' : '';
            const prevRecallFold = !isEmptyResultRow && this.recallsAtLeastOneFromImmediatePrevPeriod(displayRows, i);
            const resultCellClass = 'cell-result' + (prevRecallFold ? ' has-prev-period-recall' : '');
            const prevRecallFoldHit = prevRecallFold
                ? `<span class="prev-period-recall-fold" data-pct="${this.escapeHtml(prevRecallFoldPctLabel)}"></span>`
                : '';

            html += `<tr data-idx="${i}" class="data-row${activeClass}" data-has-result="${!!result}" data-empty="${isEmptyResultRow ? '1' : '0'}">
                <td class="cell-date"${dateBg}>${date}</td>
                <td class="cell-id"${idStyle}>${id}</td>
                <td class="${resultCellClass}">${prevRecallFoldHit}${resultHtml}</td>
                <td class="cell-note"${noteStyle}>${noteHtml}</td>
                <td class="cell-nonexist">${nonexistHtml}</td>
            </tr>`;
        }
        html += '</tbody></table>';
        tableWrap.innerHTML = html;

        if (options.skipRowClickBind !== true) {
            tableWrap.querySelectorAll('tbody tr').forEach(tr => {
                tr.style.cursor = 'pointer';
                tr.addEventListener('click', (e) => {
                    this.onRowClick(Number(tr.dataset.idx), tr.dataset.empty === '1', e);
                    try {
                        tableWrap.focus({ preventScroll: true });
                    } catch (err) {
                        // ignore focus failures
                    }
                    if (typeof options.onRowActivated === 'function') {
                        options.onRowActivated(Number(tr.dataset.idx));
                    }
                });
            });
        } else {
            tableWrap.querySelectorAll('tbody tr').forEach(tr => {
                tr.style.cursor = 'pointer';
            });
        }

        if (!tableWrap.dataset.nonexistContextmenuBound) {
            tableWrap.dataset.nonexistContextmenuBound = '1';
            tableWrap.addEventListener('contextmenu', (e) => {
                this.handleNonexistCellContextMenu(e, tableWrap);
            });
        }

        if (applyWindowSelection && this.activeWindowRange) {
            const selectionRoot = options.selectionRoot || tableWrap;
            this.applyWindowSelection(
                this.activeWindowRange.start,
                this.activeWindowRange.end,
                this.activeWindowRange.target,
                selectionRoot
            );
        }

        bindPrevPeriodRecallFoldTooltipGlobal();

        const mainWrap = typeof document !== 'undefined' ? document.getElementById('tableWrap') : null;
        if (options.applyAnswerPopupMask !== false && tableWrap && tableWrap === mainWrap) {
            this.applyAnswerPopupFocusMaskToDom(tableWrap, { reset: true });
        }
    }

    /**
     * Step the active source-sheet row using arrow keys (right pane only).
     * @returns {boolean} true when navigation was handled
     */
    stepSourceSheetRowByArrowKey(key, tableWrap) {
        const normalized = String(key || '').toLowerCase();
        const isStepForward = normalized === 'arrowdown' || normalized === 'arrowright';
        const isStepBackward = normalized === 'arrowup' || normalized === 'arrowleft';
        if (!isStepForward && !isStepBackward) {
            return false;
        }
        const step = isStepForward ? 1 : -1;
        return this.stepSourceSheetRowByDelta(step, tableWrap);
    }

    /**
     * Jump the active source-sheet row by a signed row delta (coalesced arrow bursts).
     * @returns {boolean} true when navigation was handled
     */
    stepSourceSheetRowByDelta(delta, tableWrap) {
        const step = Number(delta) || 0;
        if (!step) {
            return false;
        }

        const wrap = tableWrap || document.getElementById('tableWrap');
        if (!wrap) {
            return false;
        }

        const activeSheetMeta = this.sheets[this.activeSheet] || {};
        if (activeSheetMeta.kind === 'combo') {
            return false;
        }

        const displayRows = this.dataRows || [];
        if (displayRows.length === 0) {
            return false;
        }

        const currentIdx = this.activeWindowRange && typeof this.activeWindowRange.target === 'number'
            ? this.activeWindowRange.target
            : -1;
        if (currentIdx < 0) {
            return false;
        }

        const nextIdx = Math.max(0, Math.min(displayRows.length - 1, currentIdx + step));
        if (nextIdx === currentIdx) {
            return false;
        }

        const nextRow = wrap.querySelector(`tbody tr[data-idx="${nextIdx}"]`);
        if (!nextRow) {
            return false;
        }

        nextRow.click();
        this.centerActiveWindowInView(wrap);
        return true;
    }

    /**
     * Enable keyboard navigation on source sheets (arrow keys + Space).
     */
    bindSourceSheetKeyboardNavigation(tableWrap) {
        if (!tableWrap || tableWrap.dataset.enterNavBound === '1') {
            return;
        }

        tableWrap.dataset.enterNavBound = '1';
        if (!tableWrap.hasAttribute('tabindex')) {
            tableWrap.setAttribute('tabindex', '0');
        }

        tableWrap.addEventListener('keydown', (event) => {
            // Handle Space to toggle submit on iframe
            if (event.code === 'Space') {
                event.preventDefault();
                const frame = document.getElementById('okFrame');
                if (frame && frame.contentWindow) {
                    frame.contentWindow.postMessage({ type: 'toggleSubmit' }, '*');
                }
                return;
            }

            // Arrow keys: handled by index.html (preview per keypress, commit after coalesce).
        });
    }

    /**
     * Center the active sliding window inside the right table viewport.
     */
    centerActiveWindowInView(tableWrap) {
        if (!tableWrap || !this.activeWindowRange) {
            return;
        }

        const startIdx = this.activeWindowRange.start;
        const endIdx = this.activeWindowRange.end;
        if (typeof startIdx !== 'number' || typeof endIdx !== 'number' || endIdx < startIdx) {
            return;
        }

        const startRow = tableWrap.querySelector(`tbody tr[data-idx="${startIdx}"]`);
        const endRow = tableWrap.querySelector(`tbody tr[data-idx="${endIdx}"]`);
        if (!startRow || !endRow) {
            return;
        }

        const applyCentering = () => {
            const wrapRect = tableWrap.getBoundingClientRect();
            const startRect = startRow.getBoundingClientRect();
            const endRect = endRow.getBoundingClientRect();

            const windowCenterY = (startRect.top + endRect.bottom) / 2;
            // 60/40 view split: keep active window a bit lower than center.
            const viewportCenterY = wrapRect.top + (wrapRect.height * 0.6);
            const deltaY = windowCenterY - viewportCenterY;

            if (Math.abs(deltaY) < 1) {
                return;
            }

            const maxScrollTop = Math.max(0, tableWrap.scrollHeight - tableWrap.clientHeight);
            const nextScrollTop = Math.min(maxScrollTop, Math.max(0, tableWrap.scrollTop + deltaY));
            tableWrap.scrollTop = nextScrollTop;
        };

        // Center after click styles settle; run twice to stabilize with sticky header/zoom.
        requestAnimationFrame(() => {
            applyCentering();
            requestAnimationFrame(() => {
                applyCentering();
            });
        });
    }

    isBlankSourceRow(row) {
        const source = row || {};
        return !String(source.date || source.Date || '').trim()
            && !String(source.id || source.ID || '').trim()
            && !String(source.result || source.Result || '').trim()
            && !String(source.note || source.Note || '').trim()
            && !String(source.nonexist || source.Nonexist || '').trim();
    }

    isEmptyResultRow(row) {
        const source = row || {};
        return !String(source.result || source.Result || '').trim();
    }

    /**
     * Render a Module2-style combo sheet.
     */
    renderComboSheetHtml(sheet) {
        if (sheet.comboType === 1) {
            return this.renderCombo1SheetHtml(sheet);
        }

        const hasArrowColumn = sheet.comboType === 1;
        let html = '<div class="combo-sheet-wrap">';
        html += '<table class="sheet-data-table combo-sheet-table"><thead><tr><th>combo</th><th>appear</th>' + (hasArrowColumn ? '<th></th>' : '') + '</tr></thead><tbody>';

        const rows = sheet.data || [];
        if (rows.length === 0) {
            html += `<tr class="empty-data-row"><td colspan="${hasArrowColumn ? 3 : 2}">&nbsp;</td></tr>`;
        } else {
            for (const row of rows) {
                html += '<tr class="data-row">';
                html += `<td class="cell-combo">${this.escapeHtml(row.combo || '')}</td>`;
                html += `<td class="cell-appear">${this.escapeHtml(String(row.appear ?? ''))}</td>`;
                if (hasArrowColumn) {
                    html += `<td class="cell-arrow">${this.escapeHtml(row.arrow || '')}</td>`;
                }
                html += '</tr>';
            }
        }

        html += '</tbody></table>';

        if (sheet.comboType === 1) {
            html += '<table class="sheet-data-table combo-special-table"><thead><tr><th>special</th><th>count</th><th></th></tr></thead><tbody>';
            const specialRows = sheet.specialRows || [];
            if (specialRows.length === 0) {
                html += '<tr class="empty-data-row"><td colspan="3">&nbsp;</td></tr>';
            } else {
                for (const row of specialRows) {
                    html += '<tr class="data-row">';
                    html += `<td class="cell-special">${this.escapeHtml(row.special || '')}</td>`;
                    html += `<td class="cell-special-count">${this.escapeHtml(String(row.count ?? ''))}</td>`;
                    html += `<td class="cell-special-arrow">${this.escapeHtml(row.arrow || '')}</td>`;
                    html += '</tr>';
                }
            }
            html += '</tbody></table>';
        }

        html += '</div>';
        return html;
    }

    /**
     * Render combo_1 as a single Excel-like grid from A to K.
     * A/B/C = combo / appear / arrow
     * D/E = blank separators
     * F/G/H = logic cells from Module2 (F1 latest id, G1/H1 reserved)
     * I/J/K = special / count / arrow
     */
    renderCombo1SheetHtml(sheet) {
        const runtime = this.buildCombo1RuntimeRows();
        const comboRows = runtime.comboRows || [];
        const specialRows = runtime.specialRows || [];
        const comboState = runtime.comboState || this.buildCombo1StyleContext();
        const latestId = runtime.latestId || '';
        const f1DisplayId = comboState.focusId || this.comboFocusRowId || latestId || '';
        const rowCount = Math.max(1 + comboRows.length, 1 + specialRows.length);

        let html = '<div class="combo-sheet-wrap">';
        html += '<table class="sheet-data-table combo-sheet-grid"><colgroup>';
        html += '<col class="col-a"><col class="col-b"><col class="col-c"><col class="col-d"><col class="col-e"><col class="col-f"><col class="col-g"><col class="col-h"><col class="col-i"><col class="col-j"><col class="col-k">';
        html += '</colgroup><tbody>';

        for (let rowIndex = 1; rowIndex <= rowCount; rowIndex++) {
            const comboRow = comboRows[rowIndex - 2] || null;
            const specialRow = specialRows[rowIndex - 2] || null;
            const isHeaderRow = rowIndex === 1;
            const comboKey = comboRow ? this.normalizeNumberKey(comboRow.combo || '') : '';
            const comboRowArrow = comboKey && comboState.arrowSet.has(comboKey) ? '⬆' : (comboRow ? (comboRow.arrow || '') : '');
            const comboRowIsFreqOne = comboKey ? comboState.freqWin.get(comboKey) === 1 : false;
            const comboRowInNote = comboKey ? comboState.aggNoteWin.has(comboKey) : false;
            const comboRowInNonexist = comboKey ? comboState.targetNonexistSet.has(comboKey) : false;
            const comboRowIsTarget = comboKey ? comboState.arrowSet.has(comboKey) : false;

            html += `<tr data-row="${rowIndex}" class="data-row${isHeaderRow ? ' combo-header-row' : ''}">`;

            if (isHeaderRow) {
                html += '<td class="cell-col-a">combo</td>';
                html += '<td class="cell-col-b">appear</td>';
                html += '<td class="cell-col-c"></td>';
            } else {
                let comboCellStyle = '';
                let appearCellStyle = '';
                if (comboRowInNonexist) {
                    comboCellStyle += 'background:rgb(255,0,0);color:rgb(255,255,255);';
                    appearCellStyle += 'background:rgb(255,0,0);color:rgb(255,255,255);';
                }

                // Module2: only freq==1 entries get underline and note/nonexist-window color.
                let freqWinColor = '';
                if (comboRowIsFreqOne) {
                    comboCellStyle += 'text-decoration:underline;';
                    if (comboKey && comboState.aggNonexistWin.has(comboKey)) {
                        freqWinColor = 'rgb(234,184,40)';
                    } else if (comboKey && comboState.aggNoteWin.has(comboKey)) {
                        freqWinColor = 'rgb(0,151,167)';
                    }
                }

                if (freqWinColor) {
                    comboCellStyle += `color:${freqWinColor};`;
                }

                if (comboRowIsTarget || comboRowArrow) {
                    comboCellStyle += 'font-weight:800;';
                }
                if (comboRowInNonexist && comboRowArrow) {
                    comboCellStyle += 'background:rgb(255,255,0);color:rgb(0,0,0);';
                    appearCellStyle += 'background:rgb(255,255,0);color:rgb(0,0,0);';
                }
                if (!comboRowInNonexist && !comboRowInNote && comboRowIsTarget) {
                    comboCellStyle += 'color:rgb(0,100,0);';
                }
                const comboArrowHtml = comboRowArrow ? '<span style="font-weight:800;color:rgb(0,100,0);font-family:Segoe UI Symbol;">⬆</span>' : '';
                html += `<td class="cell-col-a"${comboCellStyle ? ` style="${comboCellStyle}"` : ''}>${comboRow ? this.escapeHtml(comboRow.combo || '') : ''}</td>`;
                html += `<td class="cell-col-b"${appearCellStyle ? ` style="${appearCellStyle}"` : ''}>${comboRow ? this.escapeHtml(String(comboRow.appear ?? '')) : ''}</td>`;
                html += `<td class="cell-col-c">${comboArrowHtml}</td>`;
            }
            html += '<td class="cell-col-d blank-cell"></td>';
            html += '<td class="cell-col-e blank-cell"></td>';
            if (isHeaderRow) {
                html += `<td class="cell-col-f combo-logic-focus-cell"><input id="comboF1CellInput" class="combo-cell-input" type="text" value="${this.escapeHtml(String(f1DisplayId))}" aria-label="F1" /></td>`;
                const g1Disabled = !!comboState.targetIsEmptyResult;
                const g1Checked = !g1Disabled && !!this.comboG1Enabled;
                html += `<td class="cell-col-g" style="background:#f8fafc;"><label style="display:flex;align-items:center;justify-content:center;width:100%;height:100%;"><input id="comboG1CellToggle" class="combo-cell-toggle" type="checkbox" aria-label="G1" ${g1Checked ? 'checked' : ''} ${g1Disabled ? 'disabled' : ''} title="${g1Disabled ? 'G1 tắt khi focus dòng không có result (chỉ giả lập H1)' : 'G1'}" /></label></td>`;
                html += `<td class="cell-col-h blank-cell combo-logic-focus-cell"><input id="comboH1CellInput" class="combo-cell-input" type="text" value="${this.escapeHtml(this.comboH1Text || '')}" aria-label="H1" /></td>`;
            } else {
                html += '<td class="cell-col-f"></td>';
                html += '<td class="cell-col-g"></td>';
                const hComment = (this.comboHComments && this.comboHComments[String(rowIndex)]) || '';
                html += `<td class="cell-col-h blank-cell combo-h-comment-cell"><input type="text" class="combo-cell-input combo-h-comment-input" data-combo-h-row="${rowIndex}" value="${this.escapeHtml(hComment)}" aria-label="H${rowIndex}" title="Enter: đặt pick trái theo chuỗi này (2–5 số 1–35, phân tách bằng dấu phẩy hoặc khoảng trắng, không trùng)" spellcheck="false" /></td>`;
            }
            html += `<td class="cell-col-i">${isHeaderRow ? 'special' : (specialRow ? this.escapeHtml(specialRow.special || '') : '')}</td>`;
            html += `<td class="cell-col-j">${isHeaderRow ? 'count' : (specialRow ? this.escapeHtml(String(specialRow.count ?? '')) : '')}</td>`;
            html += `<td class="cell-col-k">${isHeaderRow ? '' : (specialRow && this.normalizeNumberKey(specialRow.special) === this.normalizeNumberKey(comboState.targetSpecial) ? '<span style="font-weight:800;color:rgb(0,100,0);font-family:Segoe UI Symbol;">⬆</span>' : '')}</td>`;

            html += '</tr>';
        }

        html += '</tbody></table>';
        html += '<div class="combo-h-selection-layer" aria-hidden="true">';
        html += '<div class="combo-h-range-border"></div>';
        html += '<div class="combo-h-marching-ants" aria-hidden="true">';
        html += '<svg class="combo-h-marching-ants-svg" viewBox="0 0 100 100" xmlns="http://www.w3.org/2000/svg" preserveAspectRatio="none">';
        html += '<rect class="combo-h-marching-ants-rect" x="0.75" y="0.75" width="98.5" height="98.5" fill="none" pathLength="100"/>';
        html += '</svg></div>';
        html += '</div></div>';
        return html;
    }

    /**
     * Wire editable F1/G1/H1 cells after rendering combo_1.
     */
    wireCombo1HeaderControls() {
        const tableWrap = document.getElementById('tableWrap');
        if (!tableWrap || this.activeSheet !== 'combo_1') {
            return;
        }

        const f1Input = tableWrap.querySelector('#comboF1CellInput');
        const g1Toggle = tableWrap.querySelector('#comboG1CellToggle');
        const h1Input = tableWrap.querySelector('#comboH1CellInput');

        if (f1Input && !f1Input.dataset.bound) {
            f1Input.dataset.bound = '1';
            f1Input.addEventListener('input', () => {
                this.comboFocusRowId = f1Input.value.trim();
                this.save();
                window.dispatchEvent(new CustomEvent('comboControlsChanged', { detail: { sheet: this.activeSheet } }));
            });
        }

        if (g1Toggle && !g1Toggle.dataset.bound) {
            g1Toggle.dataset.bound = '1';
            g1Toggle.addEventListener('change', () => {
                if (this.isCombo1FocusEmptyResult()) {
                    g1Toggle.checked = false;
                    this.comboG1Enabled = false;
                    return;
                }
                this.comboG1Enabled = !!g1Toggle.checked;
                this.save();
                window.dispatchEvent(new CustomEvent('comboControlsChanged', { detail: { sheet: this.activeSheet } }));
            });
        }

        this.syncCombo1HeaderControlStates(f1Input, g1Toggle, h1Input);

        if (h1Input && !h1Input.dataset.bound) {
            h1Input.dataset.bound = '1';
            h1Input.addEventListener('input', () => {
                this.comboH1Text = h1Input.value;
                this.save();
                this.syncComboHColumnWidth(tableWrap);
            });

            h1Input.addEventListener('change', () => {
                this.comboH1Text = h1Input.value;
                this.save();
                this.syncComboHColumnWidth(tableWrap);
                window.dispatchEvent(new CustomEvent('comboControlsChanged', { detail: { sheet: this.activeSheet } }));
            });
        }

        this.wireComboHCommentInputs(tableWrap);
        this.wireComboHColumnExcel(tableWrap);
        this.syncComboHColumnWidth(tableWrap);
    }

    getComboHCommentRowFromCell(cell) {
        if (!cell) {
            return NaN;
        }
        const input = cell.querySelector('.combo-h-comment-input');
        return Number(input && input.dataset.comboHRow);
    }

    getComboHCommentRowList(tableWrap) {
        const rows = [];
        tableWrap.querySelectorAll('.combo-h-comment-input').forEach((input) => {
            const row = Number(input.dataset.comboHRow);
            if (Number.isFinite(row)) {
                rows.push(row);
            }
        });
        rows.sort((a, b) => a - b);
        return rows;
    }

    getComboHSelectionRowRange() {
        const sel = this.comboHSelection;
        if (!sel || !Number.isFinite(sel.anchorRow) || !Number.isFinite(sel.focusRow)) {
            return null;
        }
        return {
            minRow: Math.min(sel.anchorRow, sel.focusRow),
            maxRow: Math.max(sel.anchorRow, sel.focusRow)
        };
    }

    setComboHCellValue(rowKey, value, tableWrap) {
        if (!this.comboHComments || typeof this.comboHComments !== 'object') {
            this.comboHComments = {};
        }
        const text = value == null ? '' : String(value);
        this.comboHComments[String(rowKey)] = text;
        const wrap = tableWrap || document.getElementById('tableWrap');
        const input = wrap && wrap.querySelector(`.combo-h-comment-input[data-combo-h-row="${rowKey}"]`);
        if (input && input.value !== text) {
            input.value = text;
        }
    }

    applyComboHSelectionVisual(tableWrap) {
        if (!tableWrap) {
            return;
        }
        const range = this.getComboHSelectionRowRange();
        const focusRow = this.comboHSelection ? this.comboHSelection.focusRow : NaN;
        const isBlock = !!(range && range.maxRow > range.minRow);
        tableWrap.querySelectorAll('td.combo-h-comment-cell').forEach((cell) => {
            cell.classList.remove('combo-h-cell-active', 'combo-h-cell-selected');
            const row = this.getComboHCommentRowFromCell(cell);
            if (!Number.isFinite(row) || !range) {
                return;
            }
            if (row >= range.minRow && row <= range.maxRow) {
                cell.classList.add('combo-h-cell-selected');
            }
            if (!isBlock && row === focusRow) {
                cell.classList.add('combo-h-cell-active');
            }
        });
        this.updateComboHSelectionFrame(tableWrap);
    }

    getComboHSelectionLayer(tableWrap) {
        if (!tableWrap) {
            return null;
        }
        return tableWrap.querySelector('.combo-h-selection-layer');
    }

    updateComboHSelectionFrame(tableWrap) {
        const layer = this.getComboHSelectionLayer(tableWrap);
        const sheetWrap = tableWrap && tableWrap.querySelector('.combo-sheet-wrap');
        if (!layer || !sheetWrap) {
            return;
        }
        let range = this.getComboHSelectionRowRange();
        if (this._comboHMarchingVisible && this._comboHMarchingRange) {
            range = this._comboHMarchingRange;
        }
        if (!range) {
            layer.style.display = 'none';
            return;
        }
        const cells = [];
        tableWrap.querySelectorAll('td.combo-h-comment-cell').forEach((cell) => {
            const row = this.getComboHCommentRowFromCell(cell);
            if (Number.isFinite(row) && row >= range.minRow && row <= range.maxRow) {
                cells.push(cell);
            }
        });
        if (!cells.length) {
            layer.style.display = 'none';
            return;
        }
        cells.sort((a, b) => this.getComboHCommentRowFromCell(a) - this.getComboHCommentRowFromCell(b));
        const first = cells[0];
        const last = cells[cells.length - 1];
        const wrapRect = sheetWrap.getBoundingClientRect();
        const firstRect = first.getBoundingClientRect();
        const lastRect = last.getBoundingClientRect();
        const top = firstRect.top - wrapRect.top;
        const left = firstRect.left - wrapRect.left;
        const width = firstRect.width;
        const height = lastRect.bottom - firstRect.top;
        layer.style.display = 'block';
        layer.style.top = `${top}px`;
        layer.style.left = `${left}px`;
        layer.style.width = `${Math.ceil(width)}px`;
        layer.style.height = `${Math.ceil(height)}px`;
        layer.classList.toggle('is-marching', !!this._comboHMarchingVisible);
    }

    hideComboHMarchingAnts(tableWrap) {
        this._comboHMarchingVisible = false;
        this._comboHMarchingRange = null;
        const layer = this.getComboHSelectionLayer(tableWrap || document.getElementById('tableWrap'));
        if (layer) {
            layer.classList.remove('is-marching');
        }
        this.updateComboHSelectionFrame(tableWrap || document.getElementById('tableWrap'));
    }

    showComboHMarchingAnts(tableWrap) {
        const range = this.getComboHSelectionRowRange();
        if (range) {
            this._comboHMarchingRange = { minRow: range.minRow, maxRow: range.maxRow };
        }
        this._comboHMarchingVisible = true;
        this.updateComboHSelectionFrame(tableWrap || document.getElementById('tableWrap'));
    }

    clearComboHRowsInRange(tableWrap, range, excludeRange) {
        if (!tableWrap || !range) {
            return;
        }
        const rowList = this.getComboHCommentRowList(tableWrap);
        let changed = false;
        rowList.forEach((row) => {
            if (row < range.minRow || row > range.maxRow) {
                return;
            }
            if (excludeRange && row >= excludeRange.minRow && row <= excludeRange.maxRow) {
                return;
            }
            this.setComboHCellValue(row, '', tableWrap);
            changed = true;
        });
        if (changed) {
            this.save();
            this.syncComboHColumnWidth(tableWrap);
        }
    }

    finishComboHClipboardOp(tableWrap, pasteDestRange) {
        if (this._comboHCutPending) {
            this.clearComboHRowsInRange(tableWrap, this._comboHCutPending, pasteDestRange);
            this._comboHCutPending = null;
        }
        if (this._comboHMarchingVisible) {
            this.hideComboHMarchingAnts(tableWrap);
        }
    }

    selectComboHCell(row, opts) {
        const o = opts || {};
        const r = Number(row);
        if (!Number.isFinite(r)) {
            return;
        }
        if (o.extend && this.comboHSelection) {
            this.comboHSelection = {
                anchorRow: this.comboHSelection.anchorRow,
                focusRow: r
            };
        } else {
            this.comboHSelection = { anchorRow: r, focusRow: r };
        }
        const tableWrap = document.getElementById('tableWrap');
        this.applyComboHSelectionVisual(tableWrap);
    }

    getComboHCellValue(rowKey, tableWrap) {
        const wrap = tableWrap || document.getElementById('tableWrap');
        const key = String(rowKey);
        const input = wrap && wrap.querySelector(`.combo-h-comment-input[data-combo-h-row="${key}"]`);
        if (input) {
            return String(input.value || '');
        }
        return (this.comboHComments && this.comboHComments[key]) || '';
    }

    getComboHSelectedValues(tableWrap) {
        const range = this.getComboHSelectionRowRange();
        if (!range || !tableWrap) {
            return [];
        }
        const out = [];
        const rowList = this.getComboHCommentRowList(tableWrap);
        rowList.forEach((row) => {
            if (row >= range.minRow && row <= range.maxRow) {
                out.push({ row, value: this.getComboHCellValue(row, tableWrap) });
            }
        });
        return out;
    }

    execCommandCopyText(text) {
        const ta = document.createElement('textarea');
        ta.value = String(text == null ? '' : text);
        ta.setAttribute('readonly', '');
        ta.style.cssText = 'position:fixed;left:-9999px;top:0;width:1px;height:1px;opacity:0;';
        document.body.appendChild(ta);
        ta.focus();
        ta.select();
        let ok = false;
        try {
            ok = document.execCommand('copy');
        } catch (e) { /* ignore */ }
        document.body.removeChild(ta);
        return ok;
    }

    copyComboHSelectionNow(tableWrap) {
        const wrap = tableWrap || document.getElementById('tableWrap');
        const items = this.getComboHSelectedValues(wrap);
        if (!items.length) {
            return '';
        }
        const text = items.map((item) => item.value).join('\r\n');
        this._comboHClipboardText = text;
        if (wrap) {
            items.forEach((item) => {
                if (!this.comboHComments || typeof this.comboHComments !== 'object') {
                    this.comboHComments = {};
                }
                this.comboHComments[String(item.row)] = item.value;
            });
        }
        this.execCommandCopyText(text);
        if (navigator.clipboard && navigator.clipboard.writeText) {
            navigator.clipboard.writeText(text).catch(() => { /* execCommand / internal buffer */ });
        }
        return text;
    }

    clearComboHSelectionValues(tableWrap) {
        const items = this.getComboHSelectedValues(tableWrap);
        items.forEach((item) => {
            this.setComboHCellValue(item.row, '', tableWrap);
        });
        if (items.length) {
            this.save();
            this.syncComboHColumnWidth(tableWrap);
        }
    }

    copyComboHSelectionToClipboard() {
        this.copyComboHSelectionNow(document.getElementById('tableWrap'));
        return Promise.resolve();
    }

    pasteComboHClipboardAtFocus(tableWrap) {
        if (this._comboHClipboardText) {
            this.pasteComboHTextAtFocus(tableWrap, this._comboHClipboardText);
            return Promise.resolve();
        }
        const readText = () => {
            if (navigator.clipboard && navigator.clipboard.readText) {
                return navigator.clipboard.readText().catch(() => '');
            }
            return Promise.resolve('');
        };
        return readText().then((raw) => {
            this.pasteComboHTextAtFocus(tableWrap, raw);
        });
    }

    /**
     * Parse nội dung ô H2+ (combo_1) thành thứ tự pick trái: 2–5 số nguyên 1..35, không trùng.
     * Dấu phân tách: dấu phẩy, chấm phẩy, | hoặc khoảng trắng.
     * @param {string} raw
     * @returns {number[]|null}
     */
    parseComboHCommentTextAsLeftPickOrder(raw) {
        const MIN_PICK = 2;
        const MAX_PICK = 5;
        const t = String(raw == null ? '' : raw).trim();
        if (!t) {
            return null;
        }
        const parts = t.split(/[\s,;|]+/).map((x) => x.trim()).filter(Boolean);
        if (parts.length < MIN_PICK || parts.length > MAX_PICK) {
            return null;
        }
        const out = [];
        const seen = new Set();
        for (const p of parts) {
            const n = parseInt(p, 10);
            if (!Number.isFinite(n) || n < 1 || n > 35) {
                return null;
            }
            if (seen.has(n)) {
                return null;
            }
            seen.add(n);
            out.push(n);
        }
        return out;
    }

    /**
     * Áp dụng chuỗi trong ô Hx làm pickOrder + selected trên nửa trái (iframe ok_left).
     */
    postComboHCommentPickToLeftIframe(nums) {
        if (!nums || !nums.length) {
            return;
        }
        const frame = document.getElementById('okFrame');
        if (!frame || !frame.contentWindow) {
            return;
        }
        try {
            frame.contentWindow.postMessage({ type: 'syncAnswerPickSelection', nums }, '*');
        } catch (e) {
            /* ignore */
        }
    }

    isComboHExcelTarget(target) {
        if (!target || !target.closest) {
            return false;
        }
        return !!(target.closest('.combo-h-comment-input') || target.closest('td.combo-h-comment-cell'));
    }

    pasteComboHTextAtFocus(tableWrap, raw) {
        if (!tableWrap || raw == null) {
            return;
        }
        const sel = this.comboHSelection;
        let startRow = sel && Number.isFinite(sel.anchorRow) && Number.isFinite(sel.focusRow)
            ? Math.min(sel.anchorRow, sel.focusRow)
            : NaN;
        if (!Number.isFinite(startRow)) {
            const activeInput = tableWrap.querySelector('.combo-h-comment-input:focus');
            if (activeInput) {
                startRow = Number(activeInput.dataset.comboHRow);
            }
        }
        if (!Number.isFinite(startRow)) {
            return;
        }
        const lines = String(raw).replace(/\r\n/g, '\n').replace(/\r/g, '\n').split('\n');
        while (lines.length && lines[lines.length - 1] === '') {
            lines.pop();
        }
        const values = lines.map((line) => {
            const tab = line.split('\t');
            return tab[0] != null ? tab[0] : '';
        });
        if (!values.length || (values.length === 1 && values[0] === '' && String(raw).trim() === '')) {
            return;
        }
        const rowList = this.getComboHCommentRowList(tableWrap);
        const startIdx = rowList.indexOf(startRow);
        if (startIdx < 0) {
            return;
        }
        let changed = false;
        for (let i = 0; i < values.length && startIdx + i < rowList.length; i++) {
            const row = rowList[startIdx + i];
            this.setComboHCellValue(row, values[i], tableWrap);
            changed = true;
        }
        if (changed) {
            const endRow = rowList[Math.min(startIdx + values.length - 1, rowList.length - 1)];
            const pasteDestRange = { minRow: startRow, maxRow: endRow };
            this.comboHSelection = { anchorRow: startRow, focusRow: endRow };
            this.applyComboHSelectionVisual(tableWrap);
            this.finishComboHClipboardOp(tableWrap, pasteDestRange);
            this.save();
            this.syncComboHColumnWidth(tableWrap);
        }
    }

    ensureComboHSelectionFromFocus(tableWrap) {
        if (this.comboHSelection) {
            return;
        }
        const active = document.activeElement;
        if (!active || !active.classList || !active.classList.contains('combo-h-comment-input')) {
            return;
        }
        if (!tableWrap || !tableWrap.contains(active)) {
            return;
        }
        const row = Number(active.dataset.comboHRow);
        if (Number.isFinite(row)) {
            this.selectComboHCell(row);
        }
    }

    wireComboHColumnExcel(tableWrap) {
        if (!tableWrap) {
            return;
        }
        this.applyComboHSelectionVisual(tableWrap);
        if (this._comboHExcelWired) {
            return;
        }
        this._comboHExcelWired = true;
        if (!tableWrap.hasAttribute('tabindex')) {
            tableWrap.setAttribute('tabindex', '0');
        }

        if (tableWrap.dataset.comboHScrollBound !== '1') {
            tableWrap.dataset.comboHScrollBound = '1';
            tableWrap.addEventListener('scroll', () => {
                if (this.activeSheet === 'combo_1') {
                    this.updateComboHSelectionFrame(tableWrap);
                }
            }, { passive: true });
        }

        tableWrap.addEventListener('pointerdown', (event) => {
            if (this.activeSheet !== 'combo_1') {
                return;
            }
            const cell = event.target.closest('td.combo-h-comment-cell');
            if (!cell) {
                return;
            }
            const row = this.getComboHCommentRowFromCell(cell);
            if (!Number.isFinite(row)) {
                return;
            }
            const onInput = !!event.target.closest('.combo-h-comment-input');
            if (event.shiftKey && this.comboHSelection) {
                this.selectComboHCell(row, { extend: true, keepMarching: false });
            } else {
                this.selectComboHCell(row, { keepMarching: false });
            }
            this._comboHDragSelect = {
                anchorRow: row,
                active: false,
                pointerId: event.pointerId,
                startX: event.clientX,
                startY: event.clientY,
                onInput
            };
            if (!onInput) {
                const input = cell.querySelector('.combo-h-comment-input');
                if (input) {
                    input.focus();
                }
            }
        });

        tableWrap.addEventListener('pointermove', (event) => {
            if (!this._comboHDragSelect || event.pointerId !== this._comboHDragSelect.pointerId) {
                return;
            }
            if ((event.buttons & 1) === 0) {
                return;
            }
            const drag = this._comboHDragSelect;
            if (!drag.active) {
                const dx = Math.abs(event.clientX - drag.startX);
                const dy = Math.abs(event.clientY - drag.startY);
                if (dx < 4 && dy < 4) {
                    return;
                }
                drag.active = true;
                if (drag.onInput) {
                    event.preventDefault();
                }
            }
            const under = document.elementFromPoint(event.clientX, event.clientY);
            const cell = under && under.closest ? under.closest('td.combo-h-comment-cell') : null;
            if (!cell || !tableWrap.contains(cell)) {
                return;
            }
            const row = this.getComboHCommentRowFromCell(cell);
            if (!Number.isFinite(row)) {
                return;
            }
            this.comboHSelection = { anchorRow: drag.anchorRow, focusRow: row };
            this.hideComboHMarchingAnts(tableWrap);
            this.applyComboHSelectionVisual(tableWrap);
            if (drag.active) {
                event.preventDefault();
            }
        });

        tableWrap.addEventListener('pointerup', (event) => {
            if (this._comboHDragSelect && event.pointerId === this._comboHDragSelect.pointerId) {
                this._comboHDragSelect = null;
            }
        });

        tableWrap.addEventListener('pointercancel', () => {
            this._comboHDragSelect = null;
        });

        tableWrap.addEventListener('keydown', (event) => {
            if (this.activeSheet !== 'combo_1') {
                return;
            }
            if (!this.isComboHExcelTarget(event.target)) {
                return;
            }
            const key = String(event.key || '').toLowerCase();
            const mod = event.ctrlKey || event.metaKey;

            this.ensureComboHSelectionFromFocus(tableWrap);

            if ((key === 'delete' || key === 'backspace') && !mod) {
                if (!this.comboHSelection) {
                    return;
                }
                const range = this.getComboHSelectionRowRange();
                const multi = range && range.maxRow > range.minRow;
                const input = event.target.closest('.combo-h-comment-input');
                if (multi || (input && input.selectionStart === 0 && input.selectionEnd === input.value.length)) {
                    event.preventDefault();
                    this.clearComboHSelectionValues(tableWrap);
                }
                return;
            }

            if (key === 'escape') {
                if (this._comboHMarchingVisible || this._comboHCutPending) {
                    event.preventDefault();
                    this._comboHCutPending = null;
                    this.hideComboHMarchingAnts(tableWrap);
                }
                return;
            }

            if (key === 'enter' && !mod) {
                const input = event.target && event.target.closest
                    ? event.target.closest('.combo-h-comment-input')
                    : null;
                if (!input) {
                    return;
                }
                const row = Number(input.dataset.comboHRow);
                if (!Number.isFinite(row) || row <= 1) {
                    return;
                }
                const nums = this.parseComboHCommentTextAsLeftPickOrder(input.value);
                if (!nums) {
                    return;
                }
                event.preventDefault();
                this.postComboHCommentPickToLeftIframe(nums);
                return;
            }

            if (mod && (key === 'c' || key === 'x')) {
                if (!this.comboHSelection) {
                    return;
                }
                const items = this.getComboHSelectedValues(tableWrap);
                if (!items.length) {
                    return;
                }
                event.preventDefault();
                if (key === 'c') {
                    this._comboHCutPending = null;
                } else {
                    const cutRange = this.getComboHSelectionRowRange();
                    this._comboHCutPending = cutRange
                        ? { minRow: cutRange.minRow, maxRow: cutRange.maxRow }
                        : null;
                }
                this.copyComboHSelectionNow(tableWrap);
                this.showComboHMarchingAnts(tableWrap);
                return;
            }

            if (mod && key === 'v') {
                event.preventDefault();
                const applyPaste = (raw) => {
                    let text = raw == null ? '' : String(raw);
                    if (!text.trim() && this._comboHClipboardText) {
                        text = this._comboHClipboardText;
                    }
                    if (!text.trim() && text !== '0') {
                        return;
                    }
                    this.pasteComboHTextAtFocus(tableWrap, text);
                };
                if (this._comboHClipboardText) {
                    applyPaste(this._comboHClipboardText);
                    if (navigator.clipboard && navigator.clipboard.readText) {
                        navigator.clipboard.readText().then((t) => {
                            if (t && t.trim()) {
                                this._comboHClipboardText = t;
                            }
                        }).catch(() => { /* ignore */ });
                    }
                    return;
                }
                if (navigator.clipboard && navigator.clipboard.readText) {
                    navigator.clipboard.readText().then(applyPaste).catch(() => applyPaste(''));
                    return;
                }
                applyPaste('');
            }
        });

        tableWrap.addEventListener('paste', (event) => {
            if (this.activeSheet !== 'combo_1' || !this.isComboHExcelTarget(event.target)) {
                return;
            }
            event.preventDefault();
            this.ensureComboHSelectionFromFocus(tableWrap);
            let raw = (event.clipboardData && event.clipboardData.getData('text/plain')) || '';
            if (!raw.trim() && this._comboHClipboardText) {
                raw = this._comboHClipboardText;
            }
            this.pasteComboHTextAtFocus(tableWrap, raw);
        });
    }

    wireComboHCommentInputs(tableWrap) {
        if (!tableWrap) {
            return;
        }
        tableWrap.querySelectorAll('.combo-h-comment-input').forEach((input) => {
            if (input.dataset.bound === '1') {
                return;
            }
            input.dataset.bound = '1';
            input.addEventListener('focus', () => {
                const row = Number(input.dataset.comboHRow);
                if (Number.isFinite(row)) {
                    this.selectComboHCell(row);
                }
            });
            input.addEventListener('input', () => {
                const rowKey = String(input.dataset.comboHRow || '');
                if (!rowKey) {
                    return;
                }
                if (!this.comboHComments || typeof this.comboHComments !== 'object') {
                    this.comboHComments = {};
                }
                this.comboHComments[rowKey] = input.value;
                this.save();
                this.syncComboHColumnWidth(tableWrap);
            });
            input.addEventListener('change', () => {
                const rowKey = String(input.dataset.comboHRow || '');
                if (!rowKey) {
                    return;
                }
                if (!this.comboHComments || typeof this.comboHComments !== 'object') {
                    this.comboHComments = {};
                }
                this.comboHComments[rowKey] = input.value;
                this.save();
            });
        });
    }

    syncComboHColumnWidth(tableWrap) {
        if (!tableWrap) {
            return;
        }
        const col = tableWrap.querySelector('col.col-h');
        if (!col) {
            return;
        }
        let maxChars = 14;
        const h1Input = tableWrap.querySelector('#comboH1CellInput');
        if (h1Input) {
            maxChars = Math.max(maxChars, String(h1Input.value || '').length + 1);
        }
        tableWrap.querySelectorAll('.combo-h-comment-input').forEach((input) => {
            maxChars = Math.max(maxChars, String(input.value || '').length + 1);
        });
        maxChars = Math.min(Math.max(maxChars, 14), 48);
        col.style.width = `calc(${maxChars}ch + 12px)`;
    }

    /**
     * Sync F1/G1/H1 control enabled state (G1 locked on empty-result focus row).
     */
    syncCombo1HeaderControlStates(f1Input, g1Toggle, h1Input) {
        const emptyFocus = this.isCombo1FocusEmptyResult();
        if (g1Toggle) {
            g1Toggle.disabled = emptyFocus;
            g1Toggle.checked = emptyFocus ? false : !!this.comboG1Enabled;
        }
        if (f1Input) {
            f1Input.disabled = false;
        }
        if (h1Input) {
            h1Input.disabled = false;
        }
    }

    /**
     * Remove the black border from the previously selected 11-row window.
     */
    clearWindowSelection(tableWrapEl) {
        const tableWrap = tableWrapEl || document.getElementById('tableWrap');
        if (!tableWrap) {
            return;
        }

        tableWrap.querySelectorAll('td.window-selected, td.window-edge-top, td.window-edge-bottom, td.window-edge-left, td.window-edge-right, td.window-divider-left, td.window-divider-right, td.window-focus, .win-label-inline').forEach(cell => {
            cell.classList.remove('window-selected', 'window-edge-top', 'window-edge-bottom', 'window-edge-left', 'window-edge-right', 'window-divider-left', 'window-divider-right', 'window-focus');
            if (cell.classList && cell.classList.contains('win-label-inline')) {
                cell.remove();
            }
        });
    }

    /**
     * Apply a black border to the result/note/nonexist cells for the selected window.
     */
    applyWindowSelection(startIdx, endIdx, targetIdx = null, tableWrapEl, options = {}) {
        const tableWrap = tableWrapEl || document.getElementById('tableWrap');
        if (!tableWrap) {
            return;
        }
        const previewOnly = !!(options && options.previewOnly);

        this.clearWindowSelection(tableWrap);

        if (startIdx === null || endIdx === null || endIdx < startIdx) {
            this.activeWindowRange = null;
            if (!previewOnly) {
                this.refreshNonexistCellsForActiveWindow(tableWrap);
            }
            return;
        }

        for (let rowIdx = startIdx; rowIdx <= endIdx; rowIdx++) {
            const row = tableWrap.querySelector(`tbody tr[data-idx="${rowIdx}"]`);
            if (!row) {
                continue;
            }

            const resultCell = row.querySelector('td.cell-result');
            const noteCell = row.querySelector('td.cell-note');
            const nonexistCell = row.querySelector('td.cell-nonexist');

            [resultCell, noteCell, nonexistCell].forEach(cell => {
                if (cell) {
                    cell.classList.add('window-selected');
                }
            });

            if (resultCell) {
                resultCell.classList.add('window-edge-left');
                resultCell.classList.add('window-divider-right');
            }

            if (nonexistCell) {
                nonexistCell.classList.add('window-edge-right');
                nonexistCell.classList.add('window-divider-left');
            }

            if (noteCell) {
                noteCell.classList.add('window-divider-left');
                noteCell.classList.add('window-divider-right');
            }

            if (rowIdx === startIdx) {
                [resultCell, noteCell, nonexistCell].forEach(cell => {
                    if (cell) {
                        cell.classList.add('window-edge-top');
                    }
                });
            }

            if (rowIdx === endIdx) {
                [resultCell, noteCell, nonexistCell].forEach(cell => {
                    if (cell) {
                        cell.classList.add('window-edge-bottom');
                    }
                });
            }

            if (targetIdx !== null && rowIdx === targetIdx) {
                [resultCell, noteCell, nonexistCell].forEach(cell => {
                    if (cell) {
                        cell.classList.add('window-focus');
                    }
                });
            }
        }

        const prevWindowRange = this.activeWindowRange;
        this.activeWindowRange = { start: startIdx, end: endIdx, target: targetIdx };
        if (previewOnly) {
            return;
        }
        // Refresh nonexist HTML first; renderWindowLabels after so innerHTML does not strip labels.
        const refreshIndices = new Set();
        if (prevWindowRange && typeof prevWindowRange.start === 'number' && typeof prevWindowRange.end === 'number') {
            for (let i = prevWindowRange.start; i <= prevWindowRange.end; i++) {
                refreshIndices.add(i);
            }
        }
        for (let i = startIdx; i <= endIdx; i++) {
            refreshIndices.add(i);
        }
        this.refreshNonexistCellsForRowIndices(tableWrap, refreshIndices);
        this.renderWindowLabels(startIdx, endIdx, tableWrap);
    }

    /**
     * Arrow preview: move window/focus outline one row without loading row data.
     */
    previewSourceSheetRowByStep(step, tableWrap) {
        const delta = Number(step) || 0;
        if (!delta) {
            return false;
        }

        const wrap = tableWrap || document.getElementById('tableWrap');
        if (!wrap) {
            return false;
        }

        const activeSheetMeta = this.sheets[this.activeSheet] || {};
        if (activeSheetMeta.kind === 'combo') {
            return false;
        }

        const displayRows = this.dataRows || [];
        if (displayRows.length === 0) {
            return false;
        }

        const currentIdx = this.activeWindowRange && typeof this.activeWindowRange.target === 'number'
            ? this.activeWindowRange.target
            : -1;
        if (currentIdx < 0) {
            return false;
        }

        const nextIdx = Math.max(0, Math.min(displayRows.length - 1, currentIdx + delta));
        if (nextIdx === currentIdx) {
            return false;
        }

        const start = Math.max(0, nextIdx - 10);
        this.applyWindowSelection(start, nextIdx, nextIdx, wrap, { previewOnly: true });
        this.centerActiveWindowInView(wrap);
        return true;
    }

    /**
     * After coalesced arrow burst: load the row at the current preview target.
     */
    commitSourceSheetRowAtTarget(tableWrap) {
        const wrap = tableWrap || document.getElementById('tableWrap');
        if (!wrap) {
            return false;
        }

        const activeSheetMeta = this.sheets[this.activeSheet] || {};
        if (activeSheetMeta.kind === 'combo') {
            return false;
        }

        const targetIdx = this.activeWindowRange && typeof this.activeWindowRange.target === 'number'
            ? this.activeWindowRange.target
            : -1;
        if (targetIdx < 0) {
            return false;
        }

        const nextRow = wrap.querySelector(`tbody tr[data-idx="${targetIdx}"]`);
        if (!nextRow) {
            return false;
        }

        nextRow.click();
        this.centerActiveWindowInView(wrap);
        return true;
    }

    /**
     * Draw 10 inline labels inside the right side of the selected window rows.
     * Mirrors the VBA WinLabel_01..10 placement at the right of the block.
     */
    renderWindowLabels(startIdx, endIdx, tableWrapEl) {
        const tableWrap = tableWrapEl || document.getElementById('tableWrap');
        if (!tableWrap) {
            return;
        }

        tableWrap.querySelectorAll('.win-label-inline').forEach(label => label.remove());

        const maxLabels = Math.min(10, Math.max(0, endIdx - startIdx + 1));
        for (let offset = 0; offset < maxLabels; offset++) {
            const rowIdx = startIdx + offset;
            const row = tableWrap.querySelector(`tbody tr[data-idx="${rowIdx}"]`);
            if (!row || row.dataset.empty === '1') {
                continue;
            }

            const resultCell = row.querySelector('td.cell-result');
            const noteCell = row.querySelector('td.cell-note');
            const nonexistCell = row.querySelector('td.cell-nonexist');
            if (!resultCell || !noteCell || !nonexistCell) {
                continue;
            }

            const labelText = String(10 - offset);
            for (const cell of [resultCell, noteCell, nonexistCell]) {
                const label = document.createElement('span');
                label.className = 'win-label-inline';
                label.textContent = labelText;
                cell.appendChild(label);
            }
        }
    }

    /**
     * Build Module4-style notes for the current rows, using only result data.
     */
    buildNotesFromRows(rows) {
        const noteCache = [];
        const referenceCounts = new Map();

        for (let rowIndex = 0; rowIndex < rows.length; rowIndex++) {
            noteCache.push(this.buildNoteForRow(rows, rowIndex, referenceCounts));
        }

        return noteCache;
    }

    /**
     * Build Module4-style nonexist values for the current rows using only result data.
     */
    buildNonexistFromRows(rows) {
        const nonexistCache = [];

        for (let rowIndex = 0; rowIndex < rows.length; rowIndex++) {
            nonexistCache.push(this.buildNonexistForRow(rows, rowIndex));
        }

        return nonexistCache;
    }

    /**
     * Build the id frequency map from generated note text.
     * This mirrors Module4 ColorByNoteFrequency, which counts previous ids referenced in notes.
     */
    buildIdFrequencyMapFromNotes(noteCache) {
        const freq = new Map();

        for (const noteMeta of noteCache || []) {
            const txt = noteMeta && noteMeta.text ? String(noteMeta.text) : '';
            if (!txt || txt === '?') {
                continue;
            }

            const parts = txt.split(' ');
            for (const part of parts) {
                if (part.indexOf('-') < 0) {
                    continue;
                }

                const leftPart = part.split('=')[0];
                const idPrevRaw = String(leftPart.split('-')[1] || '').trim();
                const idPrevDigits = this.digitsOnly(idPrevRaw);
                const idPrevNum = parseInt(idPrevDigits, 10);

                if (Number.isNaN(idPrevNum)) {
                    continue;
                }

                const key = String(idPrevNum);
                freq.set(key, (freq.get(key) || 0) + 1);
            }
        }

        return freq;
    }

    /**
     * Build the generated nonexist text for a single row.
     */
    buildNonexistForRow(rows, rowIndex) {
        if (rowIndex < 10) {
            return { text: 'N/A' };
        }

        const seen = new Set();
        const startIndex = Math.max(0, rowIndex - 10);

        for (let prevIndex = startIndex; prevIndex < rowIndex; prevIndex++) {
            const prevRow = rows[prevIndex] || {};
            const prevNums = this.parseMainNums(prevRow.result || prevRow.Result || '');
            for (const num of prevNums) {
                seen.add(num);
            }
        }

        const nonexistNums = [];
        for (let num = 1; num <= 35; num++) {
            if (!seen.has(num)) {
                nonexistNums.push(num);
            }
        }

        return {
            text: nonexistNums.length > 0 ? nonexistNums.join(',') : 'N/A'
        };
    }

    /**
     * Build the generated note text for a single row.
     */
    buildNoteForRow(rows, rowIndex, referenceCounts, options = {}) {
        const currentRow = rows[rowIndex] || {};
        const currentId = this.parseRowId(currentRow.id || currentRow.ID || '');
        const currentNums = this.parseMainNums(currentRow.result || currentRow.Result || '');
        const minMainNums = Number.isFinite(options.minMainNums) ? options.minMainNums : 5;

        if (currentId === null || currentNums.length < minMainNums) {
            return { text: '?', highlightYellow: false };
        }

        const startIndex = Math.max(0, rowIndex - 10);
        const matchedNumbersByPrevId = new Map();
        const sourceRowIndexByPrevId = new Map();

        for (let prevIndex = startIndex; prevIndex < rowIndex; prevIndex++) {
            const prevRow = rows[prevIndex] || {};
            const prevId = this.parseRowId(prevRow.id || prevRow.ID || '');
            const prevNums = this.parseMainNums(prevRow.result || prevRow.Result || '');

            if (prevId === null || prevNums.length !== 5) {
                continue;
            }

            for (let a = 0; a < currentNums.length; a++) {
                for (let b = a + 1; b < currentNums.length; b++) {
                    if (this.pairExists(prevNums, currentNums[a], currentNums[b])) {
                        if (!matchedNumbersByPrevId.has(prevId)) {
                            matchedNumbersByPrevId.set(prevId, new Set());
                            sourceRowIndexByPrevId.set(prevId, prevIndex);
                        }
                        matchedNumbersByPrevId.get(prevId).add(currentNums[a]);
                        matchedNumbersByPrevId.get(prevId).add(currentNums[b]);
                    }
                }
            }
        }

        let noteText = '';
        for (const [prevId, matchedNumberSet] of matchedNumbersByPrevId.entries()) {
            const matchedNumbers = Array.from(matchedNumberSet);
            const sourceIndex = sourceRowIndexByPrevId.get(prevId);
            const prevNums = sourceIndex !== undefined ? this.parseMainNums(rows[sourceIndex].result || rows[sourceIndex].Result || '') : [];
            const idxList = [];

            for (const num of matchedNumbers) {
                for (let prevPos = 0; prevPos < prevNums.length; prevPos++) {
                    if (prevNums[prevPos] === num) {
                        idxList.push(String(prevPos + 1));
                    }
                }
            }

            const previousCount = referenceCounts.get(prevId) || 0;
            const expo = this.toSuperscript(previousCount + 1);
            const diff = currentId - prevId;

            noteText += `${currentId}-${prevId}${expo}=${diff}:{${matchedNumbers.join(',')}}|${idxList.join(';')}|   `;
            referenceCounts.set(prevId, previousCount + 1);
        }

        if (!noteText.trim()) {
            return { text: '?', highlightYellow: false };
        }

        const trimmedText = noteText.trim();
        return {
            text: trimmedText,
            highlightYellow: this.shouldHighlightNote(trimmedText)
        };
    }


    /**
     * Chỉ số dòng < endExclusive có đủ 5 số chính (walk-forward; không gồm kỳ đang dự đoán).
     */
    collectValidMainDrawIndicesBefore(rows, endExclusive) {
        const out = [];
        const rowsArr = rows || [];
        const n = Math.min(endExclusive | 0, rowsArr.length);
        for (let j = 0; j < n; j++) {
            const m = this.parseMainNums((rowsArr[j] && (rowsArr[j].result || rowsArr[j].Result)) || '');
            if (m.length === 5) {
                out.push(j);
            }
        }
        return out;
    }

    mainFiveValidNum(n) {
        const u = Number(n);
        return Number.isFinite(u) && u >= 1 && u <= 35;
    }

    mainFiveMarginalLog(countArr, n, alpha) {
        const a = Number(alpha) > 0 ? alpha : 0.35;
        let tot = 0;
        for (let i = 1; i <= 35; i++) {
            tot += countArr[i] || 0;
        }
        const c = countArr[n] || 0;
        return Math.log((c + a) / (tot + 35 * a));
    }

    /**
     * Đếm multiset main (5 số) trên tail kỳ hợp lệ cuối của validIdx (walk-forward).
     */
    mainFiveHintCountMainOnValidTail(validIdx, rowsArr, tailLen) {
        const c = new Array(36).fill(0);
        const L = validIdx.length;
        const t = Math.max(1, Math.min(L, tailLen | 0));
        const start = Math.max(0, L - t);
        for (let i = start; i < L; i++) {
            const m = this.parseMainNums((rowsArr[validIdx[i]] && (rowsArr[validIdx[i]].result || rowsArr[validIdx[i]].Result)) || '');
            for (let u = 0; u < m.length; u++) {
                const x = m[u];
                if (this.mainFiveValidNum(x)) {
                    c[x]++;
                }
            }
        }
        return c;
    }

    /**
     * gapDraws[n] = số kỳ hợp lệ kể từ lần cuối n xuất hiện đến kỳ hợp lệ cuối; avgGap[n] = TB khoảng cách (theo chỉ số kỳ hợp lệ) giữa các lần xuất hiện.
     */
    mainFiveHintGapStatsOnValid(validIdx, rowsArr) {
        const L = validIdx.length;
        const lastAt = new Array(36).fill(-1);
        const gapSum = new Array(36).fill(0);
        const gapCnt = new Array(36).fill(0);
        for (let i = 0; i < L; i++) {
            const m = this.parseMainNums((rowsArr[validIdx[i]] && (rowsArr[validIdx[i]].result || rowsArr[validIdx[i]].Result)) || '');
            for (let u = 0; u < m.length; u++) {
                const x = m[u];
                if (!this.mainFiveValidNum(x)) {
                    continue;
                }
                if (lastAt[x] >= 0) {
                    gapSum[x] += i - lastAt[x];
                    gapCnt[x]++;
                }
                lastAt[x] = i;
            }
        }
        const gapDraws = new Array(36).fill(0);
        const avgGap = new Array(36).fill(0);
        for (let n = 1; n <= 35; n++) {
            gapDraws[n] = lastAt[n] >= 0 ? L - 1 - lastAt[n] : L;
            avgGap[n] = gapCnt[n] > 0 ? gapSum[n] / gapCnt[n] : L;
        }
        return { gapDraws, avgGap };
    }

    /**
     * Mỗi cạnh (a,b) trong cùng một kỳ (tail hợp lệ): +1 vào bậc hai đỉnh (đại lượng co-occurrence đơn giản).
     */
    mainFiveHintPairDegreeOnValidTail(validIdx, rowsArr, tailLen) {
        const deg = new Array(36).fill(0);
        const L = validIdx.length;
        const t = Math.max(1, Math.min(L, tailLen | 0));
        const start = Math.max(0, L - t);
        for (let i = start; i < L; i++) {
            const m = this.parseMainNums((rowsArr[validIdx[i]] && (rowsArr[validIdx[i]].result || rowsArr[validIdx[i]].Result)) || '');
            const nums = [];
            for (let u = 0; u < m.length; u++) {
                if (this.mainFiveValidNum(m[u])) {
                    nums.push(m[u]);
                }
            }
            for (let a = 0; a < nums.length; a++) {
                for (let b = a + 1; b < nums.length; b++) {
                    deg[nums[a]]++;
                    deg[nums[b]]++;
                }
            }
        }
        return deg;
    }

    /** Chuẩn hóa z-score trên 35 số (chỉ số 1..35). */
    mainFiveHintZScore36(vals36) {
        let s = 0;
        let s2 = 0;
        const c = 35;
        for (let n = 1; n <= 35; n++) {
            const v = vals36[n] || 0;
            s += v;
            s2 += v * v;
        }
        const mean = s / c;
        const vrc = s2 / c - mean * mean;
        const std = Math.sqrt(Math.max(vrc, 1e-12));
        const out = new Array(36).fill(0);
        for (let n = 1; n <= 35; n++) {
            out[n] = ((vals36[n] || 0) - mean) / std;
        }
        return out;
    }

    /**
     * Giống mainFiveHintGapStatsOnValid nhưng thêm max khoảng cách giữa hai lần xuất hiện liên tiếp (theo chỉ số kỳ hợp lệ).
     */
    mainFiveHintGapStatsWithMaxOnValid(validIdx, rowsArr) {
        const L = validIdx.length;
        const lastAt = new Array(36).fill(-1);
        const maxBetween = new Array(36).fill(0);
        const gapSum = new Array(36).fill(0);
        const gapCnt = new Array(36).fill(0);
        for (let i = 0; i < L; i++) {
            const m = this.parseMainNums((rowsArr[validIdx[i]] && (rowsArr[validIdx[i]].result || rowsArr[validIdx[i]].Result)) || '');
            for (let u = 0; u < m.length; u++) {
                const x = m[u];
                if (!this.mainFiveValidNum(x)) {
                    continue;
                }
                if (lastAt[x] >= 0) {
                    const g = i - lastAt[x];
                    gapSum[x] += g;
                    gapCnt[x]++;
                    if (g > maxBetween[x]) {
                        maxBetween[x] = g;
                    }
                }
                lastAt[x] = i;
            }
        }
        const gapDraws = new Array(36).fill(0);
        const avgGap = new Array(36).fill(0);
        for (let n = 1; n <= 35; n++) {
            gapDraws[n] = lastAt[n] >= 0 ? L - 1 - lastAt[n] : L;
            avgGap[n] = gapCnt[n] > 0 ? gapSum[n] / gapCnt[n] : L;
        }
        return { gapDraws, avgGap, maxBetween };
    }

    /** Ma trận đồng xuất hiện cặp (đối xứng) trên tail kỳ hợp lệ. */
    mainFiveHintPairMatrixTail(validIdx, rowsArr, tailLen) {
        const P = [];
        for (let i = 0; i < 36; i++) {
            P[i] = new Array(36).fill(0);
        }
        const L = validIdx.length;
        const t = Math.max(1, Math.min(L, tailLen | 0));
        const start = Math.max(0, L - t);
        for (let i = start; i < L; i++) {
            const m = this.parseMainNums((rowsArr[validIdx[i]] && (rowsArr[validIdx[i]].result || rowsArr[validIdx[i]].Result)) || '');
            const nums = [];
            for (let u = 0; u < m.length; u++) {
                if (this.mainFiveValidNum(m[u])) {
                    nums.push(m[u]);
                }
            }
            for (let a = 0; a < nums.length; a++) {
                for (let b = a + 1; b < nums.length; b++) {
                    const u = nums[a];
                    const v = nums[b];
                    P[u][v]++;
                    P[v][u]++;
                }
            }
        }
        return P;
    }

    /** Hub = tổng trọng số cạnh; tam giác = tổng min(P_na,P_nb) (proxy triple / cụm). */
    mainFiveHintPairHubTriangle36(P) {
        const hub = new Array(36).fill(0);
        const tri = new Array(36).fill(0);
        for (let n = 1; n <= 35; n++) {
            let h = 0;
            let tr = 0;
            for (let a = 1; a <= 35; a++) {
                if (a === n) {
                    continue;
                }
                const pan = P[n][a] || 0;
                h += pan;
                for (let b = a + 1; b <= 35; b++) {
                    if (b === n) {
                        continue;
                    }
                    tr += Math.min(pan, P[n][b] || 0);
                }
            }
            hub[n] = h;
            tri[n] = tr;
        }
        return { hub, tri };
    }

    /** Phương sai tần suất xuất hiện qua các đoạn tail (proxy volatility §5). */
    mainFiveHintSegVolatility36(validIdx, rowsArr, tailDraws, nSeg) {
        const L = validIdx.length;
        const T = Math.max(nSeg, Math.min(L, tailDraws | 0));
        const start = L - T;
        const segLen = Math.max(1, Math.floor(T / Math.max(2, nSeg | 0)));
        const nS = Math.max(2, nSeg | 0);
        const segC = [];
        for (let s = 0; s < nS; s++) {
            segC.push(new Array(36).fill(0));
        }
        for (let j = 0; j < T; j++) {
            const i = start + j;
            const si = Math.min(nS - 1, Math.floor(j / segLen));
            const m = this.parseMainNums((rowsArr[validIdx[i]] && (rowsArr[validIdx[i]].result || rowsArr[validIdx[i]].Result)) || '');
            for (let u = 0; u < m.length; u++) {
                const x = m[u];
                if (this.mainFiveValidNum(x)) {
                    segC[si][x]++;
                }
            }
        }
        const vol = new Array(36).fill(0);
        const stab = new Array(36).fill(0);
        for (let n = 1; n <= 35; n++) {
            let s = 0;
            for (let t = 0; t < nS; t++) {
                s += segC[t][n];
            }
            const mean = s / nS;
            let v = 0;
            for (let t = 0; t < nS; t++) {
                const d = segC[t][n] - mean;
                v += d * d;
            }
            v /= nS;
            vol[n] = v;
            stab[n] = 1 / (1 + 4 * v);
        }
        return { vol, stab };
    }

    /**
     * Khi n xuất hiện trong tail: độ lệch cấu trúc kỳ (|lẻ−2.5| + |thấp(1–17)−2.5|) / 2 — cao hơn = kỳ lệch (§6 proxy).
     * Trả về điểm âm TB (ưu tiên số hay nằm trong kỳ “cân” hơn).
     */
    mainFiveHintHitStructuralSkew36(validIdx, rowsArr, tailLen) {
        const sumSk = new Array(36).fill(0);
        const hitCnt = new Array(36).fill(0);
        const sumSpread = new Array(36).fill(0);
        const L = validIdx.length;
        const t = Math.max(1, Math.min(L, tailLen | 0));
        const start = Math.max(0, L - t);
        for (let i = start; i < L; i++) {
            const m = this.parseMainNums((rowsArr[validIdx[i]] && (rowsArr[validIdx[i]].result || rowsArr[validIdx[i]].Result)) || '');
            const nums = [];
            for (let u = 0; u < m.length; u++) {
                if (this.mainFiveValidNum(m[u])) {
                    nums.push(m[u]);
                }
            }
            if (nums.length < 5) {
                continue;
            }
            let odd = 0;
            let low = 0;
            for (let k = 0; k < nums.length; k++) {
                if (nums[k] % 2 === 1) {
                    odd++;
                }
                if (nums[k] <= 17) {
                    low++;
                }
            }
            const sk = (Math.abs(odd - 2.5) + Math.abs(low - 2.5)) * 0.5;
            const sm = nums.slice().sort((a, b) => a - b);
            const spread = (sm[sm.length - 1] - sm[0]) / 34;
            for (let k = 0; k < nums.length; k++) {
                const n = nums[k];
                sumSk[n] += sk;
                sumSpread[n] += spread;
                hitCnt[n]++;
            }
        }
        const balHit = new Array(36).fill(0);
        const spreadHit = new Array(36).fill(0);
        for (let n = 1; n <= 35; n++) {
            const c = hitCnt[n] || 0;
            if (c > 0) {
                balHit[n] = -(sumSk[n] / c);
                spreadHit[n] = sumSpread[n] / c;
            }
        }
        return { balHit, spreadHit };
    }

    /** PRNG deterministic (§13 Monte Carlo có thể lặp lại theo seed). */
    mainFiveHintMulberry32(seed) {
        let a = seed >>> 0;
        return () => {
            a += 0x6d2b79f5;
            let t = a;
            t = Math.imul(t ^ (t >>> 15), t | 1);
            t ^= t + Math.imul(t ^ (t >>> 7), t | 61);
            return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
        };
    }

    mainFiveHintMcPoolTallyFromComposite(ctx, poolN, sampleCount) {
        const { ptxtComposite, countArr, windowCountArr } = ctx;
        const M = Math.min(35, Math.max(1, poolN | 0));
        const samples = Math.max(8, Math.min(200, sampleCount | 0));
        let seed = (ctx.forecastIdx | 0) * 2654435761;
        for (let n = 1; n <= 35; n++) {
            seed = (Math.imul(seed + (countArr[n] | 0), n + 31)) >>> 0;
        }
        const rnd = this.mainFiveHintMulberry32(seed);
        const tally = new Array(36).fill(0);
        for (let s = 0; s < samples; s++) {
            const avail = [];
            for (let n = 1; n <= 35; n++) {
                avail.push(n);
            }
            for (let pick = 0; pick < M; pick++) {
                let sumw = 0;
                const w = [];
                for (let i = 0; i < avail.length; i++) {
                    const n = avail[i];
                    const sc = Math.max(1e-6, Math.exp(0.85 * (ptxtComposite[n] || 0)));
                    w.push(sc);
                    sumw += sc;
                }
                let r = rnd() * sumw;
                let chosen = 0;
                for (let i = 0; i < avail.length; i++) {
                    r -= w[i];
                    if (r <= 0) {
                        chosen = i;
                        break;
                    }
                }
                const n = avail[chosen];
                tally[n]++;
                avail[chosen] = avail[avail.length - 1];
                avail.pop();
            }
        }
        const score36 = new Array(36).fill(0);
        for (let n = 1; n <= 35; n++) {
            score36[n] = tally[n] + 0.001 * (ptxtComposite[n] || 0);
        }
        return this.mainFiveHintSortCandidatesFromScore36(score36, countArr, windowCountArr).slice(0, M);
    }

    /**
     * Pool §11: ổn định / momentum / overdue / hub theo tỉ lệ 3:3:2:2 (scale theo poolN).
     */
    mainFiveHintPicksStratified3322Top(ctx, poolN) {
        const {
            ptxtStab,
            ptxtMom,
            ptxtOverdueRatio,
            ptxtPairHub,
            ptxtComposite,
            countArr,
            windowCountArr
        } = ctx;
        const M = Math.min(35, Math.max(1, poolN | 0));
        const sortBy = (arr) => this.mainFiveHintSortCandidatesFromScore36(arr, countArr, windowCountArr);
        const nSt = Math.max(1, Math.round((3 * M) / 10));
        const nMo = Math.max(1, Math.round((3 * M) / 10));
        const nOv = Math.max(1, Math.round((2 * M) / 10));
        const nHb = Math.max(1, M - nSt - nMo - nOv);
        const lists = [
            sortBy(ptxtStab),
            sortBy(ptxtMom),
            sortBy(ptxtOverdueRatio),
            sortBy(ptxtPairHub)
        ];
        const quotas = [nSt, nMo, nOv, nHb];
        const ptr = [0, 0, 0, 0];
        const used = new Set();
        const out = [];
        const takeFrom = (lane, need) => {
            const lst = lists[lane];
            let got = 0;
            while (got < need && ptr[lane] < lst.length) {
                const n = lst[ptr[lane]++];
                if (!used.has(n)) {
                    used.add(n);
                    out.push(n);
                    got++;
                }
            }
            return got;
        };
        for (let lane = 0; lane < 4; lane++) {
            takeFrom(lane, quotas[lane]);
        }
        if (out.length < M) {
            const fill = sortBy(ptxtComposite);
            for (let i = 0; i < fill.length && out.length < M; i++) {
                if (!used.has(fill[i])) {
                    used.add(fill[i]);
                    out.push(fill[i]);
                }
            }
        }
        return out.slice(0, M);
    }

    /**
     * Bảng 1 (copy ý hai cột): số lần : các số — chỉ các số được dự đoán (picks),
     * số lần = tần suất xuất hiện trong các kỳ học (countArr).
     */
    mainFiveFormatPredictedFrequencyTableLines(picks, countArr) {
        const byCount = new Map();
        for (let i = 0; i < picks.length; i++) {
            const n = picks[i];
            const c = countArr[n] || 0;
            if (!byCount.has(c)) {
                byCount.set(c, []);
            }
            byCount.get(c).push(n);
        }
        const counts = Array.from(byCount.keys()).sort((a, b) => b - a);
        const out = [];
        for (let j = 0; j < counts.length; j++) {
            const c = counts[j];
            const nums = byCount.get(c).slice().sort((a, b) => a - b);
            out.push(`${c}: ${nums.join(', ')}`);
        }
        return out;
    }

    /**
     * Chọn pool hint từ danh sách 1..35 đã sắp theo điểm (cao → thấp): trải đều theo thứ hạng
     * trong dải top-band — tránh toàn bộ pool là N số freq cao nhất liền nhau.
     * @param {number[]} sortedCandidates
     * @param {number} poolN
     * @param {number} bandMult
     * @returns {number[]}
     */
    mainFiveHintPickSpacedPool(sortedCandidates, poolN, bandMult) {
        const M = Math.min(35, Math.max(1, poolN | 0));
        const mult = Number(bandMult) > 0 ? bandMult : 3;
        const bandSize = Math.min(
            35,
            Math.max(M, Math.ceil(M * mult))
        );
        const band = sortedCandidates.slice(0, bandSize);
        const last = band.length - 1;
        if (M === 1 || last <= 0) {
            return band.slice(0, M);
        }
        const picks = [];
        const used = new Set();
        for (let k = 0; k < M; k++) {
            let idx = Math.floor((k * last) / (M - 1));
            idx = Math.max(0, Math.min(last, idx));
            let guard = 0;
            while (used.has(band[idx]) && guard <= last + 1) {
                idx = (idx + 1) % band.length;
                guard++;
            }
            const n = band[idx];
            used.add(n);
            picks.push(n);
        }
        const rankByNum = new Map();
        for (let i = 0; i < sortedCandidates.length; i++) {
            rankByNum.set(sortedCandidates[i], i);
        }
        picks.sort((a, b) => (rankByNum.get(a) - rankByNum.get(b)));
        return picks;
    }

    /** Thứ tự ưu tiên khi benchmark hòa điểm (cái đứng trước được giữ). */
    static MAIN_FIVE_HINT_STRATEGY_IDS = [
        'blend_spaced',
        'equal_spaced',
        'full_spaced',
        'window_spaced',
        'blend_top',
        'equal_top',
        'full_top',
        'window_top',
        'cold_spaced',
        'sumlog_spaced',
        'sumlog_top',
        'maxlog_spaced',
        'maxlog_top',
        'minlog_spaced',
        'minlog_top',
        'recent_spaced',
        'recent_top',
        'triple_spaced',
        'triple_top',
        'union_half_fill_top',
        'boost_spaced',
        'boost_top',
        'ptxt_roll20_spaced',
        'ptxt_roll20_top',
        'ptxt_composite_spaced',
        'ptxt_composite_top',
        'ptxt_gap_spaced',
        'ptxt_gap_top',
        'ptxt_deg_spaced',
        'ptxt_deg_top',
        'ptxt_strat10_top',
        'ptxt_accel_spaced',
        'ptxt_accel_top',
        'ptxt_overdue_spaced',
        'ptxt_overdue_top',
        'ptxt_maxgap_anom_spaced',
        'ptxt_maxgap_anom_top',
        'ptxt_pairhub_spaced',
        'ptxt_pairhub_top',
        'ptxt_tri_cohesion_spaced',
        'ptxt_tri_cohesion_top',
        'ptxt_volatile_spaced',
        'ptxt_volatile_top',
        'ptxt_stable_spaced',
        'ptxt_stable_top',
        'ptxt_balhit_spaced',
        'ptxt_balhit_top',
        'ptxt_spreadhit_spaced',
        'ptxt_spreadhit_top',
        'ptxt_dyn_composite_spaced',
        'ptxt_dyn_composite_top',
        'ptxt_bayes7030_spaced',
        'ptxt_bayes7030_top',
        'ptxt_diverse_spaced',
        'ptxt_diverse_top',
        'ptxt_multisig_spaced',
        'ptxt_multisig_top',
        'ptxt_strat3322_top',
        'ptxt_mc_pool_top',
        'ptxt_cls_stable_top',
        'ptxt_cls_momentum_top',
        'ptxt_cls_overdue_top',
        'ptxt_cls_hub_top',
        'ptxt_cls_tri_top'
    ];

    /**
     * @param {number[]} score36 chỉ số 1..35 dùng được
     * @returns {number[]} 1..35 sắp cao → thấp
     */
    mainFiveHintSortCandidatesFromScore36(score36, countArr, windowCountArr) {
        const candidates = [];
        for (let n = 1; n <= 35; n++) {
            candidates.push(n);
        }
        candidates.sort((a, b) => {
            const ds = score36[b] - score36[a];
            if (Math.abs(ds) > 1e-12) {
                return ds > 0 ? 1 : -1;
            }
            const dc = (countArr[b] || 0) - (countArr[a] || 0);
            if (dc !== 0) {
                return dc > 0 ? 1 : -1;
            }
            const dwc = (windowCountArr[b] || 0) - (windowCountArr[a] || 0);
            if (dwc !== 0) {
                return dwc > 0 ? 1 : -1;
            }
            return a - b;
        });
        return candidates;
    }

    /**
     * Tính count tích lũy / cửa sổ + log margin (walk-forward, không dùng dòng idx).
     * @returns {object | null}
     */
    mainFiveHintBuildScoresAtIndex(rowsArr, idx) {
        const validIdx = this.collectValidMainDrawIndicesBefore(rowsArr, idx);
        if (!validIdx.length) {
            return null;
        }
        const alphaRaw = Number(this.constructor.MAIN_FIVE_HINT_DIRICHLET_ALPHA);
        const alpha = Number.isFinite(alphaRaw) && alphaRaw > 0 ? alphaRaw : 1.65;
        const L = validIdx.length;
        const slideLen = Math.min(
            10,
            Math.max(1, Number(this.constructor.MAIN_FIVE_HINT_SLIDE_LEN) || 10)
        );
        let slideW = Number(this.constructor.MAIN_FIVE_HINT_SLIDE_WEIGHT);
        if (!Number.isFinite(slideW)) {
            slideW = 0.45;
        }
        slideW = Math.max(0, Math.min(1, slideW));
        const countArr = new Array(36).fill(0);
        for (let i = 0; i < L; i++) {
            const m = this.parseMainNums((rowsArr[validIdx[i]] && (rowsArr[validIdx[i]].result || rowsArr[validIdx[i]].Result)) || '');
            for (let t = 0; t < m.length; t++) {
                const x = m[t];
                if (this.mainFiveValidNum(x)) {
                    countArr[x]++;
                }
            }
        }
        const table1WinEnd = idx - 1;
        const table1WinStart = Math.max(0, idx - slideLen);
        const windowCountArr = new Array(36).fill(0);
        let sheetWinRowCount = 0;
        for (let ri = table1WinStart; ri <= table1WinEnd; ri++) {
            const m = this.parseMainNums((rowsArr[ri] && (rowsArr[ri].result || rowsArr[ri].Result)) || '');
            for (let t = 0; t < m.length; t++) {
                const x = m[t];
                if (this.mainFiveValidNum(x)) {
                    windowCountArr[x]++;
                }
            }
            sheetWinRowCount++;
        }
        const fullLog = new Array(36).fill(0);
        const winLog = new Array(36).fill(0);
        for (let n = 1; n <= 35; n++) {
            fullLog[n] = this.mainFiveMarginalLog(countArr, n, alpha);
            winLog[n] = this.mainFiveMarginalLog(windowCountArr, n, alpha);
        }
        const recentK = Math.min(
            L,
            Math.max(1, Number(this.constructor.MAIN_FIVE_HINT_RECENT_VALID_K) || 20)
        );
        const recentStart = Math.max(0, L - recentK);
        const recentCountArr = new Array(36).fill(0);
        for (let ri = recentStart; ri < L; ri++) {
            const m = this.parseMainNums((rowsArr[validIdx[ri]] && (rowsArr[validIdx[ri]].result || rowsArr[validIdx[ri]].Result)) || '');
            for (let t = 0; t < m.length; t++) {
                const x = m[t];
                if (this.mainFiveValidNum(x)) {
                    recentCountArr[x]++;
                }
            }
        }
        const recentLog = new Array(36).fill(0);
        for (let n = 1; n <= 35; n++) {
            recentLog[n] = this.mainFiveMarginalLog(recentCountArr, n, alpha);
        }
        const cnt20 = this.mainFiveHintCountMainOnValidTail(validIdx, rowsArr, 20);
        const cnt50 = this.mainFiveHintCountMainOnValidTail(validIdx, rowsArr, 50);
        const cnt100 = this.mainFiveHintCountMainOnValidTail(validIdx, rowsArr, 100);
        const ptxtRoll20Log = new Array(36).fill(0);
        const ptxtRoll50Log = new Array(36).fill(0);
        const ptxtRoll100Log = new Array(36).fill(0);
        const ptxtMom = new Array(36).fill(0);
        for (let n = 1; n <= 35; n++) {
            ptxtRoll20Log[n] = this.mainFiveMarginalLog(cnt20, n, alpha);
            ptxtRoll50Log[n] = this.mainFiveMarginalLog(cnt50, n, alpha);
            ptxtRoll100Log[n] = this.mainFiveMarginalLog(cnt100, n, alpha);
            ptxtMom[n] = ptxtRoll20Log[n] - ptxtRoll100Log[n];
        }
        const { gapDraws, avgGap, maxBetween } = this.mainFiveHintGapStatsWithMaxOnValid(validIdx, rowsArr);
        const ptxtGapRaw = new Array(36).fill(0);
        const ptxtOverdueRatio = new Array(36).fill(0);
        const ptxtMaxgapAnom = new Array(36).fill(0);
        for (let n = 1; n <= 35; n++) {
            const ag = avgGap[n] || 1;
            const gd = gapDraws[n] || 0;
            ptxtGapRaw[n] = Math.log(1 + gd / (ag + 0.35));
            ptxtOverdueRatio[n] = gd / (ag + 0.25);
            const mx = Math.max(maxBetween[n] || 0, 1);
            ptxtMaxgapAnom[n] = Math.log(1 + gd / (mx + 0.2));
        }
        const ptxtPairDeg = this.mainFiveHintPairDegreeOnValidTail(validIdx, rowsArr, 50);
        const P50 = this.mainFiveHintPairMatrixTail(validIdx, rowsArr, 50);
        const { hub: ptxtPairHub, tri: ptxtTriCohesion } = this.mainFiveHintPairHubTriangle36(P50);
        const { vol: ptxtVol, stab: ptxtStab } = this.mainFiveHintSegVolatility36(validIdx, rowsArr, 120, 6);
        const { balHit: ptxtBalHit, spreadHit: ptxtSpreadHit } = this.mainFiveHintHitStructuralSkew36(validIdx, rowsArr, 50);
        const ptxtAccel = new Array(36).fill(0);
        for (let n = 1; n <= 35; n++) {
            const v1 = ptxtRoll20Log[n] - ptxtRoll50Log[n];
            const v2 = ptxtRoll50Log[n] - ptxtRoll100Log[n];
            ptxtAccel[n] = v1 - v2;
        }
        const zR20 = this.mainFiveHintZScore36(ptxtRoll20Log);
        const zR50 = this.mainFiveHintZScore36(ptxtRoll50Log);
        const zMom = this.mainFiveHintZScore36(ptxtMom);
        const zGap = this.mainFiveHintZScore36(ptxtGapRaw);
        const zDeg = this.mainFiveHintZScore36(ptxtPairDeg);
        const zFull = this.mainFiveHintZScore36(fullLog);
        const zOver = this.mainFiveHintZScore36(ptxtOverdueRatio);
        const zMaxg = this.mainFiveHintZScore36(ptxtMaxgapAnom);
        const zHub = this.mainFiveHintZScore36(ptxtPairHub);
        const zTri = this.mainFiveHintZScore36(ptxtTriCohesion);
        const zVol = this.mainFiveHintZScore36(ptxtVol);
        const zStab = this.mainFiveHintZScore36(ptxtStab);
        const zBal = this.mainFiveHintZScore36(ptxtBalHit);
        const zSpr = this.mainFiveHintZScore36(ptxtSpreadHit);
        const zAcc = this.mainFiveHintZScore36(ptxtAccel);
        const ptxtComposite = new Array(36).fill(0);
        for (let n = 1; n <= 35; n++) {
            ptxtComposite[n] = 0.32 * zR20[n]
                + 0.18 * zR50[n]
                + 0.16 * zMom[n]
                + 0.12 * zGap[n]
                + 0.17 * zDeg[n]
                + 0.05 * zFull[n];
        }
        const dynT = Math.min(1, L / 480);
        const ptxtCompositeDyn = new Array(36).fill(0);
        for (let n = 1; n <= 35; n++) {
            ptxtCompositeDyn[n] = (0.34 + 0.11 * dynT) * zR20[n]
                + (0.19 - 0.05 * dynT) * zR50[n]
                + 0.14 * zMom[n]
                + 0.11 * zGap[n]
                + (0.16 - 0.03 * dynT) * zDeg[n]
                + 0.05 * zFull[n]
                + 0.04 * zAcc[n]
                + 0.03 * zHub[n];
        }
        const ptxtDiverseMix = new Array(36).fill(0);
        const ptxtBayes7030 = new Array(36).fill(0);
        const ptxtMultiSig = new Array(36).fill(0);
        for (let n = 1; n <= 35; n++) {
            ptxtDiverseMix[n] = ptxtComposite[n] - 0.12 * zDeg[n];
            ptxtBayes7030[n] = 0.7 * fullLog[n] + 0.3 * winLog[n];
            ptxtMultiSig[n] = 0.28 * zOver[n] + 0.28 * zGap[n] + 0.22 * zHub[n] + 0.22 * zAcc[n];
        }
        return {
            forecastIdx: idx,
            validIdx,
            countArr,
            windowCountArr,
            recentCountArr,
            fullLog,
            winLog,
            recentLog,
            ptxtRoll20Log,
            ptxtRoll50Log,
            ptxtRoll100Log,
            ptxtMom,
            ptxtGapRaw,
            ptxtOverdueRatio,
            ptxtMaxgapAnom,
            ptxtPairDeg,
            ptxtPairHub,
            ptxtTriCohesion,
            ptxtVol,
            ptxtStab,
            ptxtBalHit,
            ptxtSpreadHit,
            ptxtAccel,
            ptxtComposite,
            ptxtCompositeDyn,
            ptxtDiverseMix,
            ptxtBayes7030,
            ptxtMultiSig,
            sheetWinRowCount,
            table1WinStart,
            table1WinEnd,
            slideLen,
            slideW,
            alpha,
            L,
            recentK
        };
    }

    /**
     * @param {string} strategyId một trong MAIN_FIVE_HINT_STRATEGY_IDS
     * @param {object} ctx kết quả mainFiveHintBuildScoresAtIndex
     * @returns {number[]}
     */
    mainFiveHintPicksForStrategy(strategyId, ctx, poolN, bandMult) {
        const {
            fullLog,
            winLog,
            slideW,
            countArr,
            windowCountArr,
            recentLog,
            ptxtRoll20Log,
            ptxtComposite,
            ptxtCompositeDyn,
            ptxtGapRaw,
            ptxtPairDeg,
            ptxtOverdueRatio,
            ptxtMaxgapAnom,
            ptxtPairHub,
            ptxtTriCohesion,
            ptxtVol,
            ptxtStab,
            ptxtBalHit,
            ptxtSpreadHit,
            ptxtAccel,
            ptxtDiverseMix,
            ptxtBayes7030,
            ptxtMultiSig
        } = ctx;
        const blend = (n) => (1 - slideW) * fullLog[n] + slideW * winLog[n];
        if (strategyId === 'union_half_fill_top') {
            return this.mainFiveHintPicksUnionHalfFillTop(ctx, poolN, blend);
        }
        if (strategyId === 'ptxt_strat10_top') {
            return this.mainFiveHintPicksStratifiedPtxtTop(ctx, poolN);
        }
        if (strategyId === 'ptxt_strat3322_top') {
            return this.mainFiveHintPicksStratified3322Top(ctx, poolN);
        }
        if (strategyId === 'ptxt_mc_pool_top') {
            const Ctor = this.constructor;
            const ns = Number(Ctor.MAIN_FIVE_HINT_MC_POOL_SAMPLES) || 40;
            return this.mainFiveHintMcPoolTallyFromComposite(ctx, poolN, ns);
        }
        const score36 = new Array(36).fill(0);
        const equal = (n) => 0.5 * fullLog[n] + 0.5 * winLog[n];
        switch (strategyId) {
            case 'blend_spaced':
            case 'blend_top':
                for (let n = 1; n <= 35; n++) {
                    score36[n] = blend(n);
                }
                break;
            case 'equal_spaced':
            case 'equal_top':
                for (let n = 1; n <= 35; n++) {
                    score36[n] = equal(n);
                }
                break;
            case 'full_spaced':
            case 'full_top':
                for (let n = 1; n <= 35; n++) {
                    score36[n] = fullLog[n];
                }
                break;
            case 'window_spaced':
            case 'window_top':
                for (let n = 1; n <= 35; n++) {
                    score36[n] = winLog[n];
                }
                break;
            case 'cold_spaced':
                for (let n = 1; n <= 35; n++) {
                    score36[n] = -fullLog[n];
                }
                break;
            case 'sumlog_spaced':
            case 'sumlog_top':
                for (let n = 1; n <= 35; n++) {
                    score36[n] = fullLog[n] + winLog[n];
                }
                break;
            case 'maxlog_spaced':
            case 'maxlog_top':
                for (let n = 1; n <= 35; n++) {
                    score36[n] = Math.max(fullLog[n], winLog[n]);
                }
                break;
            case 'minlog_spaced':
            case 'minlog_top':
                for (let n = 1; n <= 35; n++) {
                    score36[n] = Math.min(fullLog[n], winLog[n]);
                }
                break;
            case 'recent_spaced':
            case 'recent_top':
                for (let n = 1; n <= 35; n++) {
                    score36[n] = recentLog[n];
                }
                break;
            case 'triple_spaced':
            case 'triple_top':
                for (let n = 1; n <= 35; n++) {
                    score36[n] = (fullLog[n] + winLog[n] + recentLog[n]) / 3;
                }
                break;
            case 'ptxt_roll20_spaced':
            case 'ptxt_roll20_top':
                for (let n = 1; n <= 35; n++) {
                    score36[n] = ptxtRoll20Log[n];
                }
                break;
            case 'ptxt_composite_spaced':
            case 'ptxt_composite_top':
                for (let n = 1; n <= 35; n++) {
                    score36[n] = ptxtComposite[n];
                }
                break;
            case 'ptxt_gap_spaced':
            case 'ptxt_gap_top':
                for (let n = 1; n <= 35; n++) {
                    score36[n] = ptxtGapRaw[n];
                }
                break;
            case 'ptxt_deg_spaced':
            case 'ptxt_deg_top':
                for (let n = 1; n <= 35; n++) {
                    score36[n] = ptxtPairDeg[n];
                }
                break;
            case 'ptxt_accel_spaced':
            case 'ptxt_accel_top':
                for (let n = 1; n <= 35; n++) {
                    score36[n] = ptxtAccel[n];
                }
                break;
            case 'ptxt_overdue_spaced':
            case 'ptxt_overdue_top':
                for (let n = 1; n <= 35; n++) {
                    score36[n] = ptxtOverdueRatio[n];
                }
                break;
            case 'ptxt_maxgap_anom_spaced':
            case 'ptxt_maxgap_anom_top':
                for (let n = 1; n <= 35; n++) {
                    score36[n] = ptxtMaxgapAnom[n];
                }
                break;
            case 'ptxt_pairhub_spaced':
            case 'ptxt_pairhub_top':
                for (let n = 1; n <= 35; n++) {
                    score36[n] = ptxtPairHub[n];
                }
                break;
            case 'ptxt_tri_cohesion_spaced':
            case 'ptxt_tri_cohesion_top':
                for (let n = 1; n <= 35; n++) {
                    score36[n] = ptxtTriCohesion[n];
                }
                break;
            case 'ptxt_volatile_spaced':
            case 'ptxt_volatile_top':
                for (let n = 1; n <= 35; n++) {
                    score36[n] = ptxtVol[n];
                }
                break;
            case 'ptxt_stable_spaced':
            case 'ptxt_stable_top':
                for (let n = 1; n <= 35; n++) {
                    score36[n] = ptxtStab[n];
                }
                break;
            case 'ptxt_balhit_spaced':
            case 'ptxt_balhit_top':
                for (let n = 1; n <= 35; n++) {
                    score36[n] = ptxtBalHit[n];
                }
                break;
            case 'ptxt_spreadhit_spaced':
            case 'ptxt_spreadhit_top':
                for (let n = 1; n <= 35; n++) {
                    score36[n] = ptxtSpreadHit[n];
                }
                break;
            case 'ptxt_dyn_composite_spaced':
            case 'ptxt_dyn_composite_top':
                for (let n = 1; n <= 35; n++) {
                    score36[n] = ptxtCompositeDyn[n];
                }
                break;
            case 'ptxt_bayes7030_spaced':
            case 'ptxt_bayes7030_top':
                for (let n = 1; n <= 35; n++) {
                    score36[n] = ptxtBayes7030[n];
                }
                break;
            case 'ptxt_diverse_spaced':
            case 'ptxt_diverse_top':
                for (let n = 1; n <= 35; n++) {
                    score36[n] = ptxtDiverseMix[n];
                }
                break;
            case 'ptxt_multisig_spaced':
            case 'ptxt_multisig_top':
                for (let n = 1; n <= 35; n++) {
                    score36[n] = ptxtMultiSig[n];
                }
                break;
            case 'ptxt_cls_stable_top':
                for (let n = 1; n <= 35; n++) {
                    score36[n] = ptxtStab[n];
                }
                break;
            case 'ptxt_cls_momentum_top':
                for (let n = 1; n <= 35; n++) {
                    score36[n] = ptxtAccel[n];
                }
                break;
            case 'ptxt_cls_overdue_top':
                for (let n = 1; n <= 35; n++) {
                    score36[n] = ptxtOverdueRatio[n];
                }
                break;
            case 'ptxt_cls_hub_top':
                for (let n = 1; n <= 35; n++) {
                    score36[n] = ptxtPairHub[n];
                }
                break;
            case 'ptxt_cls_tri_top':
                for (let n = 1; n <= 35; n++) {
                    score36[n] = ptxtTriCohesion[n];
                }
                break;
            case 'boost_spaced':
            case 'boost_top':
                for (let n = 1; n <= 35; n++) {
                    score36[n] = blend(n)
                        + 0.12 * Math.min(12, countArr[n] || 0)
                        + 0.08 * Math.min(8, windowCountArr[n] || 0);
                }
                break;
            default:
                for (let n = 1; n <= 35; n++) {
                    score36[n] = blend(n);
                }
        }
        const sorted = this.mainFiveHintSortCandidatesFromScore36(score36, countArr, windowCountArr);
        const spaced = strategyId.endsWith('_spaced');
        if (spaced) {
            return this.mainFiveHintPickSpacedPool(sorted, poolN, bandMult);
        }
        return sorted.slice(0, poolN);
    }

    /**
     * Lấy ceil(poolN/2) số đầu theo full, ceil(poolN/2) theo window (trùng thì bỏ qua), rồi fill theo blend tới đủ poolN.
     * @param {function(number): number} blend
     */
    mainFiveHintPicksUnionHalfFillTop(ctx, poolN, blend) {
        const { fullLog, winLog, countArr, windowCountArr } = ctx;
        const M = Math.min(35, Math.max(1, poolN | 0));
        const half = Math.min(17, Math.ceil(M / 2));
        const scoreF = new Array(36).fill(0);
        const scoreW = new Array(36).fill(0);
        const scoreB = new Array(36).fill(0);
        for (let n = 1; n <= 35; n++) {
            scoreF[n] = fullLog[n];
            scoreW[n] = winLog[n];
            scoreB[n] = blend(n);
        }
        const sortedF = this.mainFiveHintSortCandidatesFromScore36(scoreF, countArr, windowCountArr);
        const sortedW = this.mainFiveHintSortCandidatesFromScore36(scoreW, countArr, windowCountArr);
        const sortedB = this.mainFiveHintSortCandidatesFromScore36(scoreB, countArr, windowCountArr);
        const out = [];
        for (let i = 0; i < sortedF.length && out.length < half; i++) {
            const n = sortedF[i];
            if (!out.includes(n)) {
                out.push(n);
            }
        }
        let winAdded = 0;
        for (let i = 0; i < sortedW.length && winAdded < half && out.length < M; i++) {
            const n = sortedW[i];
            if (!out.includes(n)) {
                out.push(n);
                winAdded++;
            }
        }
        for (let i = 0; i < sortedB.length && out.length < M; i++) {
            const n = sortedB[i];
            if (!out.includes(n)) {
                out.push(n);
            }
        }
        return out.slice(0, M);
    }

    /**
     * Ghép pool theo tỉ lệ 4:3:2:1 (composite : roll20 : gap : pair-degree), lặp chu kỳ 10 slot — top liền, walk-forward.
     */
    mainFiveHintPicksStratifiedPtxtTop(ctx, poolN) {
        const { ptxtComposite, ptxtRoll20Log, ptxtGapRaw, ptxtPairDeg, countArr, windowCountArr } = ctx;
        const M = Math.min(35, Math.max(1, poolN | 0));
        const sortBy = (arr) => this.mainFiveHintSortCandidatesFromScore36(arr, countArr, windowCountArr);
        const lists = [
            sortBy(ptxtComposite),
            sortBy(ptxtRoll20Log),
            sortBy(ptxtGapRaw),
            sortBy(ptxtPairDeg)
        ];
        const laneMap = [0, 0, 0, 0, 1, 1, 1, 2, 2, 3];
        const ptr = [0, 0, 0, 0];
        const used = new Set();
        const out = [];
        const tryPushFrom = (lane) => {
            const lst = lists[lane];
            while (ptr[lane] < lst.length && used.has(lst[ptr[lane]])) {
                ptr[lane]++;
            }
            if (ptr[lane] < lst.length) {
                const n = lst[ptr[lane]++];
                used.add(n);
                out.push(n);
                return true;
            }
            return false;
        };
        for (let si = 0; si < M; si++) {
            const primary = laneMap[si % 10];
            if (tryPushFrom(primary)) {
                continue;
            }
            let filled = false;
            for (let r = 0; r < 4 && !filled; r++) {
                filled = tryPushFrom((primary + 1 + r) % 4);
            }
            if (!filled) {
                break;
            }
        }
        if (out.length < M) {
            const rest = sortBy(ptxtComposite);
            for (let i = 0; i < rest.length && out.length < M; i++) {
                if (!used.has(rest[i])) {
                    used.add(rest[i]);
                    out.push(rest[i]);
                }
            }
        }
        return out;
    }

    /**
     * Walk-forward: với mỗi dòng có đủ 5 số thật, pool dự đoán chỉ từ quá khứ; điểm = số trùng trong pool / 5.
     * @returns {{ winner: string, means: Record<string, number>, evalRows: number }}
     */
    runMainFiveHintStrategyBenchmark(rowsArr) {
        const Ctor = this.constructor;
        const poolN = Math.min(Math.max(1, Number(Ctor.MAIN_FIVE_HINT_POOL_SIZE) || 10), 35);
        const bandMult = Number(Ctor.MAIN_FIVE_HINT_DIVERSITY_BAND_MULT) || 3;
        const ids = Ctor.MAIN_FIVE_HINT_STRATEGY_IDS || ['blend_spaced'];
        const sums = {};
        const counts = {};
        for (let ii = 0; ii < ids.length; ii++) {
            sums[ids[ii]] = 0;
            counts[ids[ii]] = 0;
        }
        let evalRows = 0;
        const list = rowsArr || [];
        for (let idx = 0; idx < list.length; idx++) {
            const truth = this.parseMainNums((list[idx] && (list[idx].result || list[idx].Result)) || '');
            if (truth.length !== 5) {
                continue;
            }
            const ctx = this.mainFiveHintBuildScoresAtIndex(list, idx);
            if (!ctx) {
                continue;
            }
            evalRows++;
            const truthSet = new Set(truth);
            for (let si = 0; si < ids.length; si++) {
                const sid = ids[si];
                const picks = this.mainFiveHintPicksForStrategy(sid, ctx, poolN, bandMult);
                let ov = 0;
                for (let pi = 0; pi < picks.length; pi++) {
                    if (truthSet.has(picks[pi])) {
                        ov++;
                    }
                }
                sums[sid] += ov;
                counts[sid]++;
            }
        }
        const means = {};
        let winner = ids[0];
        let bestMean = -1;
        for (let oi = 0; oi < ids.length; oi++) {
            const sid = ids[oi];
            const c = counts[sid] || 0;
            const mean = c > 0 ? sums[sid] / c : 0;
            means[sid] = mean;
            if (mean > bestMean + 1e-12) {
                bestMean = mean;
                winner = sid;
            }
        }
        if (evalRows < 1) {
            winner = 'blend_spaced';
        }
        return { winner, means, evalRows };
    }

    /** Một dòng giải thích chiến lược đang áp (tránh nhầm λ/cửa sổ với full_* / window_*). */
    mainFiveHintStrategyOneLinerVn(id) {
        switch (id) {
            case 'full_top':
                return 'Giải thích: chỉ xếp hạng theo tích lũy + Dirichlet (α) — lấy đúng 10 số cao nhất liền nhau. Trọng số λ ở dòng “cửa sổ” không tham gia xếp hạng picks của chiến lược này.';
            case 'full_spaced':
                return 'Giải thích: chỉ tích lũy + Dirichlet; pool spaced trong top-band. λ không tham gia xếp hạng.';
            case 'window_top':
                return 'Giải thích: chỉ multiset cửa sổ 10 dòng + Dirichlet — top 10 liền. λ không tham gia xếp hạng.';
            case 'window_spaced':
                return 'Giải thích: chỉ cửa sổ + Dirichlet; pool spaced. λ không tham gia xếp hạng.';
            case 'blend_top':
                return 'Giải thích: điểm = (1−λ)·margin(tích lũy) + λ·margin(cửa sổ); lấy top 10 liền.';
            case 'blend_spaced':
                return 'Giải thích: điểm = (1−λ)·margin(tích lũy) + λ·margin(cửa sổ); pool spaced trong top-band.';
            case 'equal_top':
                return 'Giải thích: điểm = 50% tích lũy + 50% cửa sổ; top 10 liền.';
            case 'equal_spaced':
                return 'Giải thích: điểm = 50% tích lũy + 50% cửa sổ; pool spaced.';
            case 'cold_spaced':
                return 'Giải thích: ưu tiên số có margin tích lũy thấp (ít “hot”), pool spaced.';
            case 'sumlog_spaced':
            case 'sumlog_top':
                return 'Giải thích: điểm = margin(tích lũy) + margin(cửa sổ dòng sheet).';
            case 'maxlog_spaced':
            case 'maxlog_top':
                return 'Giải thích: điểm = max(margin tích lũy, margin cửa sổ) — mạnh ở một trong hai.';
            case 'minlog_spaced':
            case 'minlog_top':
                return 'Giải thích: điểm = min(hai margin) — ưu tiên số không quá yếu ở cả hai nguồn.';
            case 'recent_spaced':
            case 'recent_top':
                return 'Giải thích: chỉ margin trên K kỳ hợp lệ gần nhất (MAIN_FIVE_HINT_RECENT_VALID_K).';
            case 'triple_spaced':
            case 'triple_top':
                return 'Giải thích: trung bình margin tích lũy + cửa sổ + K kỳ hợp lệ gần nhất.';
            case 'union_half_fill_top':
                return 'Giải thích: nửa pool theo tích lũy, nửa theo cửa sổ (trừ trùng), phần còn lại fill theo blend.';
            case 'boost_spaced':
            case 'boost_top':
                return 'Giải thích: blend + nhẹ hệ số theo count tích lũy / cửa sổ (ưu số vừa hot vừa lộ diện nhiều).';
            case 'ptxt_roll20_spaced':
            case 'ptxt_roll20_top':
                return 'Giải thích (predict.txt): margin Dirichlet trên 20 kỳ hợp lệ gần nhất — rolling ngắn.';
            case 'ptxt_composite_spaced':
            case 'ptxt_composite_top':
                return 'Giải thích (predict.txt): điểm tổng hợp z (roll20/50, momentum 20−100, gap, bậc cặp 50 kỳ, nhẹ full margin).';
            case 'ptxt_gap_spaced':
            case 'ptxt_gap_top':
                return 'Giải thích (predict.txt): tín hiệu “gap” — log(1 + khoảng cách kể từ lần cuối / TB khoảng cách).';
            case 'ptxt_deg_spaced':
            case 'ptxt_deg_top':
                return 'Giải thích (predict.txt): bậc đồ thị đơn giản — mỗi cặp trong cùng kỳ (tail 50) cộng 1 cho hai đỉnh.';
            case 'ptxt_strat10_top':
                return 'Giải thích (predict.txt): pool ghép stratified top — chu kỳ 10 slot theo tỉ lệ composite:roll20:gap:deg = 4:3:2:1.';
            case 'ptxt_accel_spaced':
            case 'ptxt_accel_top':
                return 'Giải thích (predict.txt §1/§4): acceleration — (margin20−50) − (margin50−100), proxy tốc độ thay đổi rolling.';
            case 'ptxt_overdue_spaced':
            case 'ptxt_overdue_top':
                return 'Giải thích (predict.txt §2): overdue_ratio ≈ gap hiện tại / TB gap (anomaly, không hứa “sắp ra”).';
            case 'ptxt_maxgap_anom_spaced':
            case 'ptxt_maxgap_anom_top':
                return 'Giải thích (predict.txt §2): gap hiện tại so với max gap lịch sử giữa các lần xuất hiện — tín hiệu bất thường.';
            case 'ptxt_pairhub_spaced':
            case 'ptxt_pairhub_top':
            case 'ptxt_cls_hub_top':
                return 'Giải thích (predict.txt §3): “hub” — tổng trọng số cạnh đồng xuất hiện (ma trận cặp tail 50).';
            case 'ptxt_tri_cohesion_spaced':
            case 'ptxt_tri_cohesion_top':
            case 'ptxt_cls_tri_top':
                return 'Giải thích (predict.txt §3): tam giác/cụm — Σ min(edge(n,a), edge(n,b)) trên các cặp a,b.';
            case 'ptxt_volatile_spaced':
            case 'ptxt_volatile_top':
                return 'Giải thích (predict.txt §5): volatility — phương sai tần suất qua các đoạn tail (số burst / im lặng thay đổi).';
            case 'ptxt_stable_spaced':
            case 'ptxt_stable_top':
            case 'ptxt_cls_stable_top':
                return 'Giải thích (predict.txt §5): stability = 1/(1+4·var) trên các đoạn tail — ưu tiên ổn định.';
            case 'ptxt_balhit_spaced':
            case 'ptxt_balhit_top':
                return 'Giải thích (predict.txt §6): khi số xuất hiện trong tail, kỳ đó thường lệch lẻ/cao-thấp ít hay nhiều (âm TB skew → “cân” hơn).';
            case 'ptxt_spreadhit_spaced':
            case 'ptxt_spreadhit_top':
                return 'Giải thích (predict.txt §6): spread kỳ (max−min)/34 trung bình trên các kỳ số đó có mặt — cấu trúc rộng/hẹp.';
            case 'ptxt_dyn_composite_spaced':
            case 'ptxt_dyn_composite_top':
                return 'Giải thích (predict.txt §9): composite z có trọng số phụ thuộc độ dài L (nhấn rolling gần khi L lớn).';
            case 'ptxt_bayes7030_spaced':
            case 'ptxt_bayes7030_top':
                return 'Giải thích (predict.txt §8): tin cậy mềm — 0.7 margin tích lũy + 0.3 margin cửa sổ (smoothing, không cập nhật từng kỳ riêng).';
            case 'ptxt_diverse_spaced':
            case 'ptxt_diverse_top':
                return 'Giải thích (predict.txt §7): composite z trừ nhẹ bậc cặp (tránh pool quá tập trung “hub” cùng kiểu).';
            case 'ptxt_multisig_spaced':
            case 'ptxt_multisig_top':
                return 'Giải thích (predict.txt §9): tổng có trọng số các z overdue, gap, hub, acceleration.';
            case 'ptxt_strat3322_top':
                return 'Giải thích (predict.txt §11): pool 3+3+2+2 — stable / momentum / overdue / hub (scale theo poolN), fill composite.';
            case 'ptxt_mc_pool_top':
                return 'Giải thích (predict.txt §13): Monte Carlo — nhiều lần lấy mẫu 10 số không lặp theo softmax(composite z), seed deterministic; chọn theo tần suất xuất hiện trong mẫu.';
            case 'ptxt_cls_momentum_top':
                return 'Giải thích (predict.txt §12): lớp “Momentum Rising” — xếp theo acceleration (proxy).';
            case 'ptxt_cls_overdue_top':
                return 'Giải thích (predict.txt §12): lớp “Overdue Candidate” — xếp theo overdue_ratio.';
            default:
                return '';
        }
    }

    /**
     * Chạy benchmark một lần trên tham chiếu mảng dòng (thường sheet1), cache theo rowsRef.
     */
    ensureMainFiveHintStrategyFromBenchmark(rowsArr) {
        const Ctor = this.constructor;
        if (!Ctor.MAIN_FIVE_HINT_AUTO_SELECT_STRATEGY) {
            this._mainFiveHintStrategyCache = {
                rowsRef: rowsArr,
                winner: 'blend_spaced',
                means: {},
                evalRows: 0,
                autoOff: true
            };
            return;
        }
        if (this._mainFiveHintStrategyCache && this._mainFiveHintStrategyCache.rowsRef === rowsArr) {
            return;
        }
        const bench = this.runMainFiveHintStrategyBenchmark(rowsArr);
        this._mainFiveHintStrategyCache = {
            rowsRef: rowsArr,
            winner: bench.winner,
            means: bench.means,
            evalRows: bench.evalRows
        };
    }

    /**
     * Pool top-N số chính: margin Dirichlet (tích lũy + cửa sổ) — tự chọn chiến lược pool
     * (spaced vs top, blend vs full vs window…) theo overlap TB walk-forward trên toàn sheet khi bật AUTO.
     * @returns {{ picks: number[], lines: string[] } | { error: string }}
     */
    predictMainFiveHonest(rows, rowIndex) {
        const rowsArr = rows || [];
        const idx = Number(rowIndex);
        if (!Number.isFinite(idx) || idx < 0) {
            return { error: 'Chỉ số dòng không hợp lệ.' };
        }
        const built = this.mainFiveHintBuildScoresAtIndex(rowsArr, idx);
        if (!built) {
            return { error: 'Chưa có kỳ nào trước dòng này có đủ 5 số chính (trước dấu |) để dự đoán.' };
        }
        this.ensureMainFiveHintStrategyFromBenchmark(rowsArr);
        const cache = this._mainFiveHintStrategyCache || { winner: 'blend_spaced', means: {}, evalRows: 0 };
        const strategyId = cache.winner || 'blend_spaced';
        const poolN = Math.min(
            Math.max(1, Number(this.constructor.MAIN_FIVE_HINT_POOL_SIZE) || 10),
            35
        );
        const bandMult = Number(this.constructor.MAIN_FIVE_HINT_DIVERSITY_BAND_MULT) || 3;
        const picks = this.mainFiveHintPicksForStrategy(strategyId, built, poolN, bandMult);
        const { validIdx, countArr, windowCountArr, sheetWinRowCount, table1WinStart, table1WinEnd, slideLen, slideW, alpha, recentK } = built;
        const lines = [];
        lines.push(`Dự đoán pool ${poolN} số chính (kỳ đang chọn, chỉ số dòng ${idx}; không dùng result/note dòng này):`);
        lines.push(`Chiến lược: ${strategyId}${cache.autoOff ? ' (AUTO tắt, cố định blend_spaced)' : ''}`);
        const oneLiner = this.mainFiveHintStrategyOneLinerVn(strategyId);
        if (oneLiner) {
            lines.push(oneLiner);
        }
        if (!cache.autoOff && cache.evalRows > 0) {
            const wm = cache.means && typeof cache.means[strategyId] === 'number' ? cache.means[strategyId] : 0;
            lines.push(
                `Benchmark sheet1: overlap TB ${wm.toFixed(3)} / 5 (${cache.evalRows} kỳ walk-forward).`
            );
            const parts = [];
            const ids = this.constructor.MAIN_FIVE_HINT_STRATEGY_IDS || [];
            for (let zi = 0; zi < ids.length; zi++) {
                const id = ids[zi];
                if (cache.means && typeof cache.means[id] === 'number') {
                    parts.push(`${id}:${cache.means[id].toFixed(3)}`);
                }
            }
            if (parts.length) {
                lines.push(`Điểm TB các thuật: ${parts.join(' | ')}`);
            }
        }
        lines.push(`Picks: ${picks.join(', ')}`);
        lines.push('');
        lines.push(
            `Tham số nền (α, cửa sổ ${sheetWinRowCount} dòng ${table1WinStart}–${table1WinEnd}, λ=${slideW.toFixed(2)}, recent=${recentK} kỳ hợp lệ): blend/equal/boost/full/window/…; ptxt_* theo predict.txt (rolling, gap+maxgap, đồ thị cặp/tam giác, volatility/ổn định, cấu trúc kỳ, composite động, Bayes 70/30, đa tín hiệu, strat 4:3:2:1 & 3:3:2:2, MC pool MAIN_FIVE_HINT_MC_POOL_SAMPLES).`
        );
        lines.push('');
        lines.push(`Bảng 1 (tích lũy) — các số được dự đoán (số lần trong ${validIdx.length} kỳ học : các số):`);
        lines.push(...this.mainFiveFormatPredictedFrequencyTableLines(picks, countArr));
        lines.push('');
        lines.push(
            `Bảng 1 (cửa sổ ${sheetWinRowCount} dòng) — cùng multiset với bảng 1 trái cho cửa sổ này (số lần : các số):`
        );
        lines.push(...this.mainFiveFormatPredictedFrequencyTableLines(picks, windowCountArr));
        lines.push('');
        lines.push(
            `Mô hình: multiset + Dirichlet margin (α=${alpha}); pool _top / spaced trong top-${Math.min(35, Math.max(poolN, Math.ceil(poolN * bandMult)))}. Tắt AUTO: MAIN_FIVE_HINT_AUTO_SELECT_STRATEGY = false.`
        );
        return { picks, lines };
    }

    /**
     * Gợi ý tham chiếu (textarea + picks để khoanh trong iframe trái).
     * @returns {{ text: string, picks: number[] } | { error: string } | null}
     */
    getNoteReferenceHintMeta(rowIndex) {
        const rows = this.getSourceSheetRows();
        const idx = Number(rowIndex);
        if (!Number.isFinite(idx) || idx < 0 || idx >= rows.length) {
            return null;
        }
        const r = this.predictMainFiveHonest(rows, idx);
        if (r.error) {
            return { error: r.error };
        }
        const cap = Math.min(
            Math.max(1, Number(this.constructor.MAIN_FIVE_HINT_POOL_SIZE) || 10),
            35
        );
        return { text: r.lines.join('\n'), picks: Array.isArray(r.picks) ? r.picks.slice(0, cap) : [] };
    }

    /**
     * Text #referenceHint (iframe trái): dự đoán pool top-N số chính cho kỳ đang focus sheet1.
     */
    getNoteReferenceHintForRowIndex(rowIndex) {
        const m = this.getNoteReferenceHintMeta(rowIndex);
        if (!m) {
            return '';
        }
        if (m.error) {
            return m.error;
        }
        return m.text || '';
    }

    /**
     * Get the computed note for a row, falling back to a raw note only if needed.
     */
    getComputedNoteMeta(rowIndex, row) {
        const sourceRowCount = this.getSourceSheetRows().length;
        if (!this.noteCache || this.noteCache.length !== sourceRowCount) {
            this.refreshDerivedState();
        }

        if (this.noteCache && this.noteCache[rowIndex]) {
            return this.noteCache[rowIndex];
        }

        return {
            text: '?',
            highlightYellow: false
        };
    }

    /**
     * Get the computed nonexist value for a row.
     */
    getComputedNonexistMeta(rowIndex, row) {
        const sourceRowCount = this.getSourceSheetRows().length;
        if (!this.nonexistCache || this.nonexistCache.length !== sourceRowCount) {
            this.refreshDerivedState();
        }

        if (this.nonexistCache && this.nonexistCache[rowIndex]) {
            return this.nonexistCache[rowIndex];
        }

        return {
            text: 'N/A'
        };
    }

    /**
     * Highlight note text using the same rules as Module4 HighlightNoteCell.
     */
    renderNoteHtml(noteText, highlightYellow) {
        const escaped = this.escapeHtml(noteText || '');
        const styledPipeSegments = escaped.replace(/\|([^|]*)\|/g, (match, inner) => {
            const highlightedInner = inner.replace(/\b\d+\b/g, (num) => {
                return `<span style="color:rgb(0,80,0);font-weight:bold">${num}</span>`;
            });
            return `|${highlightedInner}|`;
        });

        if (!highlightYellow) {
            return styledPipeSegments;
        }

        return styledPipeSegments;
    }

    /**
     * Resolve nonexist text for a source sheet row (same rules as renderSourceSheet).
     */
    getNonexistMetaForSourceRow(rowIndex, row) {
        const isEmptyResultRow = this.isEmptyResultRow(row);
        if (isEmptyResultRow) {
            const provided = String(row.nonexist || row.Nonexist || '').trim();
            if (provided.length > 0) {
                return { text: provided };
            }
            if (String(row.id || row.ID || '').trim().length > 0) {
                return this.getComputedNonexistMeta(rowIndex, row);
            }
            return { text: '' };
        }
        return this.getComputedNonexistMeta(rowIndex, row);
    }

    /**
     * Start row of the 10-period lookback used for nonexist at rowIndex (Module4 window).
     */
    getNonexistLookbackStart(rowIndex) {
        return Math.max(0, rowIndex - 10);
    }

    /**
     * Consecutive rows (going up) where `num` appears in generated nonexist, down to minRow.
     */
    countNonexistStreak(rowIndex, num, minRowInclusive) {
        if (!this.nonexistCache || rowIndex < 0) {
            return 0;
        }
        const candidate = parseInt(num, 10);
        if (isNaN(candidate)) {
            return 0;
        }
        const minRow = Math.max(0, minRowInclusive);
        let count = 0;
        for (let r = rowIndex; r >= minRow; r--) {
            const meta = this.nonexistCache[r];
            const text = meta ? String(meta.text || '').trim() : '';
            if (!text || text === 'N/A') {
                break;
            }
            const nums = this.parseNums(text);
            if (nums.indexOf(candidate) === -1) {
                break;
            }
            count++;
        }
        return count;
    }

    /**
     * Streak of `num` in nonexist continues before this row's 10-period lookback window.
     */
    isNonexistLongerOutsideWindow(rowIndex, num) {
        const sourceRowCount = this.getSourceSheetRows().length;
        if (!this.nonexistCache || this.nonexistCache.length !== sourceRowCount) {
            this.refreshDerivedState();
        }
        const windowStart = this.getNonexistLookbackStart(rowIndex);
        const fullStreak = this.countNonexistStreak(rowIndex, num, 0);
        const windowStreak = this.countNonexistStreak(rowIndex, num, windowStart);
        return fullStreak > windowStreak;
    }

    shouldBoostYellowNonexistForWindow(rowIndex, num) {
        const win = this.activeWindowRange;
        if (!win || typeof win.start !== 'number' || typeof win.end !== 'number') {
            return false;
        }
        const start = win.start;
        const end = win.end;
        if (end < start) {
            return false;
        }
        const maxLabels = Math.min(10, Math.max(0, end - start + 1));
        const lastLabeledRow = start + maxLabels - 1;
        if (rowIndex < start || rowIndex > lastLabeledRow) {
            return false;
        }
        const bottomMeta = this.nonexistCache && this.nonexistCache[end];
        if (!bottomMeta) {
            return false;
        }
        const bottomText = String(bottomMeta.text || '').trim();
        if (!bottomText || bottomText === 'N/A') {
            return false;
        }
        const bottomNums = this.parseNums(bottomText);
        return bottomNums.indexOf(num) !== -1;
    }

    /**
     * Re-render nonexist cells for specific row indices (yellow x1.5 tracks active window).
     */
    refreshNonexistCellsForRowIndices(tableWrap, rowIndices) {
        if (!tableWrap || this.activeSheet !== 'sheet1') {
            return;
        }
        const meta = this.sheets[this.activeSheet] || {};
        if (meta.kind === 'combo') {
            return;
        }
        const displayRows = this.dataRows || [];
        if (!this.nonexistCache || this.nonexistCache.length !== displayRows.length) {
            this.refreshDerivedState();
        }

        const indices = rowIndices instanceof Set
            ? rowIndices
            : new Set(Array.isArray(rowIndices) ? rowIndices : []);
        if (indices.size === 0) {
            return;
        }

        for (const i of indices) {
            if (i < 0 || i >= displayRows.length) {
                continue;
            }
            const tr = tableWrap.querySelector(`tbody tr[data-idx="${i}"]`);
            if (!tr) {
                continue;
            }
            const cell = tr.querySelector('td.cell-nonexist');
            if (!cell) {
                continue;
            }
            const row = displayRows[i];
            cell.innerHTML = this.renderSourceRowNonexistCellHtml(i, row);
            tr.classList.toggle('answer-popup-focus-masked', this.shouldAnswerPopupMaskSheet1Row(i));
            cell.classList.toggle('answer-popup-focus-nonexist', this.shouldAnswerPopupMaskSheet1Row(i));
        }
    }

    /**
     * Re-render nonexist column for the active sliding window (and previous window when it moves).
     */
    refreshNonexistCellsForActiveWindow(tableWrap) {
        const refreshIndices = new Set();
        const win = this.activeWindowRange;
        if (win && typeof win.start === 'number' && typeof win.end === 'number') {
            for (let i = win.start; i <= win.end; i++) {
                refreshIndices.add(i);
            }
        }
        this.refreshNonexistCellsForRowIndices(tableWrap, refreshIndices);
    }

    /**
     * Visual state for one nonexist cell (shared by renderNonexistHtml and left-pane freq=0).
     * @returns {{ longestSet: Set<string>, prevNonexistNums: Set<number>, currentNums: Set<number> } | null}
     */
    computeNonexistVisualState(rowIndex, nonexistText, currentResult) {
        if (!nonexistText || nonexistText === 'N/A') {
            return null;
        }

        const currentNums = new Set(this.parseMainNums(currentResult));
        const prevNonexist = rowIndex > 0 && this.nonexistCache && this.nonexistCache[rowIndex - 1]
            ? String(this.nonexistCache[rowIndex - 1].text || '')
            : '';
        const prevNonexistNums = new Set(prevNonexist === 'N/A' ? [] : this.parseNums(prevNonexist));
        const candidateNums = this.parseNums(nonexistText);

        const longestCounts = new Map();
        let longestCount = 0;

        for (const candidate of candidateNums) {
            const candidateText = String(candidate);
            let count = 1;

            for (let previousRow = rowIndex - 1; previousRow >= 0; previousRow--) {
                const previousMeta = this.nonexistCache && this.nonexistCache[previousRow] ? this.nonexistCache[previousRow] : null;
                const previousText = previousMeta ? String(previousMeta.text || '') : '';

                if (!previousText || previousText === 'N/A') {
                    break;
                }

                const previousNums = new Set(this.parseNums(previousText));
                if (previousNums.has(candidate)) {
                    count++;
                } else {
                    break;
                }
            }

            longestCounts.set(candidateText, count);
            if (count > longestCount) {
                longestCount = count;
            }
        }

        const longestSet = new Set();
        for (const [candidateText, count] of longestCounts.entries()) {
            if (count === longestCount) {
                longestSet.add(candidateText);
            }
        }

        return { longestSet, prevNonexistNums, currentNums };
    }

    /**
     * Highlight kind for one number in a nonexist list ('yellow' | 'red' | 'green' | 'green-ul' | 'green-italic' | '').
     */
    getNonexistHighlightKindForNumber(rowIndex, num, nonexistText, currentResult, state = null) {
        const visual = state || this.computeNonexistVisualState(rowIndex, nonexistText, currentResult);
        if (!visual) {
            return '';
        }

        const value = parseInt(num, 10);
        if (isNaN(value)) {
            return '';
        }

        const valueText = String(value);
        const isInDiff = !visual.prevNonexistNums.has(value);
        const isMatch = visual.currentNums.has(value);
        const isLongest = visual.longestSet.has(valueText);

        if (isLongest) {
            return isMatch ? 'green-ul' : 'red';
        }
        if (isMatch && isInDiff) {
            return 'green-italic';
        }
        if (isInDiff) {
            return 'yellow';
        }
        if (isMatch) {
            return 'green';
        }
        return '';
    }

    /**
     * Final display kind for one nonexist number (matches renderNonexistHtml priority).
     */
    getNonexistDisplayKindForNumber(rowIndex, num, nonexistText, currentResult, state = null) {
        const visual = state || this.computeNonexistVisualState(rowIndex, nonexistText, currentResult);
        if (!visual) {
            return '';
        }

        const value = parseInt(num, 10);
        if (isNaN(value)) {
            return '';
        }

        const kind = this.getNonexistHighlightKindForNumber(
            rowIndex,
            num,
            nonexistText,
            currentResult,
            visual
        );
        const isMatch = visual.currentNums.has(value);
        const longerOutside = this.isNonexistLongerOutsideWindow(rowIndex, value);

        if (kind === 'red') {
            return 'red';
        }
        if (kind === 'green-ul') {
            return 'green-ul';
        }
        if (longerOutside && isMatch) {
            return 'green-strike';
        }
        if (longerOutside && !isMatch) {
            return 'purple';
        }
        return kind || '';
    }

    /**
     * Map number -> highlight kind for the clicked/focus row (left table freq=0).
     */
    buildNonexistHighlightMapForRow(rowIndex) {
        const row = (this.dataRows || [])[rowIndex];
        if (!row) {
            return {};
        }
        const nonexistMeta = this.getNonexistMetaForSourceRow(rowIndex, row);
        const nonexistText = String(nonexistMeta.text || '').trim();
        if (!nonexistText || nonexistText === 'N/A') {
            return {};
        }
        const currentResult = row.result || row.Result || '';
        const state = this.computeNonexistVisualState(rowIndex, nonexistText, currentResult);
        const out = {};
        const candidates = this.parseNums(nonexistText);
        for (let i = 0; i < candidates.length; i++) {
            const num = candidates[i];
            const kind = this.getNonexistDisplayKindForNumber(
                rowIndex,
                num,
                nonexistText,
                currentResult,
                state
            );
            if (kind) {
                out[num] = kind;
            }
        }
        return out;
    }

    /**
     * Render nonexist text using the generated values from result data only.
     */
    renderNonexistHtml(rowIndex, nonexistText, currentResult) {
        if (!nonexistText || nonexistText === 'N/A') {
            return this.escapeHtml(nonexistText || '');
        }

        const greenStrikeStyle = 'color:rgb(0,80,0);font-weight:bold;font-size:1.5em;text-decoration:line-through';
        const redLongestStyle = 'color:rgb(255,0,0);font-weight:bold';
        const purpleOutsideStyle = 'color:rgb(148,55,220);font-weight:bold';

        return this.escapeHtml(nonexistText).replace(/\b\d+\b/g, (match) => {
            const value = parseInt(match, 10);
            const displayKind = this.getNonexistDisplayKindForNumber(rowIndex, value, nonexistText, currentResult);

            if (displayKind === 'red') {
                return `<span style="${redLongestStyle}">${value}</span>`;
            }
            if (displayKind === 'green-ul') {
                return `<span style="color:rgb(0,80,0);font-weight:bold;text-decoration:underline;font-size:1.5em">${value}</span>`;
            }
            if (displayKind === 'green-strike') {
                return `<span style="${greenStrikeStyle}">${value}</span>`;
            }
            if (displayKind === 'purple') {
                return `<span style="${purpleOutsideStyle}">${value}</span>`;
            }

            if (!displayKind) {
                return match;
            }
            if (displayKind === 'green-italic') {
                return `<span style="color:rgb(0,80,0);font-weight:bold;font-style:italic;font-size:1.5em">${value}</span>`;
            }
            if (displayKind === 'yellow') {
                const boost = this.shouldBoostYellowNonexistForWindow(rowIndex, value);
                const fs = boost ? 'font-size:1.5em;' : '';
                return `<span style="color:rgb(240,200,64);font-weight:bold;${fs}">${value}</span>`;
            }
            if (displayKind === 'green') {
                return `<span style="color:rgb(0,80,0);font-weight:bold;font-size:1.5em">${value}</span>`;
            }
            return match;
        });
    }

    /**
     * Check whether the note should get the yellow cell background.
     */
    shouldHighlightNote(noteText) {
        if (!noteText || noteText === '?') return false;

        const noteParts = String(noteText).split('   ');
        for (const part of noteParts) {
            const openBrace = part.indexOf('{');
            const closeBrace = part.indexOf('}', openBrace + 1);
            if (openBrace >= 0 && closeBrace > openBrace) {
                const inside = part.substring(openBrace + 1, closeBrace);
                const nums = inside.split(',').map(x => x.trim()).filter(Boolean);
                if (nums.length >= 3) {
                    return true;
                }
            }
        }

        return false;
    }

    /**
     * Parse only the five main numbers before the pipe in a result cell.
     */
    parseMainNums(s) {
        if (!s) return [];
        const leftPart = String(s).split('|')[0];
        return leftPart.split(',').map(x => parseInt(x, 10)).filter(n => !isNaN(n));
    }

    /**
     * Có thể so sánh gọi lại với kỳ liền trước (cả hai có đủ 5 số chính).
     */
    isEligibleForPrevPeriodRecallComparison(rows, rowIndex) {
        const list = rows || [];
        const idx = Number(rowIndex);
        if (!Number.isFinite(idx) || idx < 1 || idx >= list.length) {
            return false;
        }
        const curRow = list[idx];
        const prevRow = list[idx - 1];
        if (this.isEmptyResultRow(curRow) || this.isEmptyResultRow(prevRow)) {
            return false;
        }
        const cur = this.parseMainNums(curRow.result || curRow.Result);
        const prev = this.parseMainNums(prevRow.result || prevRow.Result);
        return cur.length === 5 && prev.length === 5;
    }

    /**
     * Kỳ tại rowIndex có ≥1 số chính (5 số trước |) trùng kỳ liền trước (cùng thứ tự data.json).
     */
    recallsAtLeastOneFromImmediatePrevPeriod(rows, rowIndex) {
        const list = rows || [];
        const idx = Number(rowIndex);
        if (!Number.isFinite(idx) || idx < 1 || idx >= list.length) {
            return false;
        }
        const curRow = list[idx];
        const prevRow = list[idx - 1];
        if (this.isEmptyResultRow(curRow) || this.isEmptyResultRow(prevRow)) {
            return false;
        }
        const cur = this.parseMainNums(curRow.result || curRow.Result);
        const prev = this.parseMainNums(prevRow.result || prevRow.Result);
        if (cur.length !== 5 || prev.length !== 5) {
            return false;
        }
        const prevSet = new Set(prev);
        for (let i = 0; i < cur.length; i++) {
            if (prevSet.has(cur[i])) {
                return true;
            }
        }
        return false;
    }

    /**
     * % kỳ có corner fold trong mẫu rowIndices (bảng đầy đủ hoặc tập lọc Ctrl popup).
     */
    computePrevPeriodRecallFoldStats(rows, rowIndices) {
        const list = rows || [];
        const indices = Array.isArray(rowIndices)
            ? rowIndices.filter(i => i >= 0 && i < list.length)
            : list.map((_, i) => i);
        let eligible = 0;
        let withRecall = 0;
        for (let k = 0; k < indices.length; k++) {
            const i = indices[k];
            if (!this.isEligibleForPrevPeriodRecallComparison(list, i)) {
                continue;
            }
            eligible++;
            if (this.recallsAtLeastOneFromImmediatePrevPeriod(list, i)) {
                withRecall++;
            }
        }
        return {
            eligible,
            withRecall,
            pct: eligible > 0 ? (withRecall / eligible) * 100 : null
        };
    }

    formatPrevPeriodRecallFoldPct(stats) {
        if (!stats || !stats.eligible) {
            return '—';
        }
        const rounded = Math.round(stats.pct * 10) / 10;
        return rounded.toFixed(1) + '%';
    }

    /**
     * Cập nhật tooltip % trên góc fold (sau lọc popup khi tái dùng HTML cache).
     */
    applyPrevPeriodRecallFoldTooltips(tableWrap, rows, rowIndices) {
        if (!tableWrap) {
            return null;
        }
        const stats = this.computePrevPeriodRecallFoldStats(rows, rowIndices);
        const pctLabel = this.formatPrevPeriodRecallFoldPct(stats);
        tableWrap.querySelectorAll('td.cell-result.has-prev-period-recall').forEach(function (cell) {
            let hit = cell.querySelector('.prev-period-recall-fold');
            if (!hit) {
                hit = document.createElement('span');
                hit.className = 'prev-period-recall-fold';
                cell.insertBefore(hit, cell.firstChild);
            }
            hit.setAttribute('data-pct', pctLabel);
            hit.removeAttribute('title');
        });
        return stats;
    }

    /**
     * Parse a row id into a number.
     */
    parseRowId(value) {
        const parsed = parseInt(String(value).trim(), 10);
        return Number.isNaN(parsed) ? null : parsed;
    }

    /**
     * Keep only digit characters from a string.
     */
    digitsOnly(text) {
        return String(text || '').replace(/\D/g, '');
    }

    /**
     * Get the id background color by note frequency.
     */
    getIdBackgroundByFrequency(rawId) {
        const idNum = this.parseRowId(rawId);
        if (idNum === null || !this.idFrequencyMap) {
            return '';
        }

        const freq = this.idFrequencyMap.get(String(idNum)) || 0;
        switch (freq) {
            case 1:
                return 'rgb(235, 255, 235)';
            case 2:
                return 'rgb(200, 255, 200)';
            case 3:
                return 'rgb(120, 230, 120)';
            default:
                return freq > 0 ? 'rgb(0, 180, 0)' : '';
        }
    }

    /**
     * Check whether a pair exists in a row's main result numbers.
     */
    pairExists(nums, a, b) {
        return nums.includes(a) && nums.includes(b);
    }

    /**
     * Highlight the date cell when the current row has at least one pair
     * formed against a row inside the previous 10-row window.
     */
    shouldHighlightDateByPairWindow(rows, rowIndex) {
        const currentRow = rows[rowIndex] || {};
        const currentNums = this.parseMainNums(currentRow.result || currentRow.Result || '');

        if (currentNums.length !== 5) {
            return false;
        }

        if (rowIndex < 10) {
            return false;
        }

        const windowRows = rows.slice(Math.max(0, rowIndex - 10), rowIndex);
        if (windowRows.length < 10) {
            return false;
        }

        const visiblePairs = this.computePairsForRows(windowRows);
        if (!visiblePairs || visiblePairs.length === 0) {
            return false;
        }

        for (const pair of visiblePairs) {
            if (this.pairExists(currentNums, pair[0], pair[1])) {
                return true;
            }
        }

        return false;
    }

    /**
     * Cửa sổ 10 dòng: index tăng theo thời gian (0 = xa nhất, 9 = sát kỳ hiện tại).
     * Đếm biên kề nghịch (top có a, bot có b) và lật nằm dưới mọi dòng main (top > mainIdx — đọc từ sát kỳ, vd. 222 8:{26,32}).
     * ≥ 2 biên → cặp rác.
     */
    countOrderOnlyAdjacentBoundaries(sets, a, b) {
        if (!sets || sets.length < 2 || !Number.isFinite(a) || !Number.isFinite(b)) {
            return 0;
        }
        const mains = [];
        for (let idx = 0; idx < sets.length; idx++) {
            if (sets[idx].has(a) && sets[idx].has(b)) {
                mains.push(idx);
            }
        }
        const mainMax = mains.length ? Math.max(...mains) : -1;
        const soleMainOldest = mains.length === 1 && mains[0] === 0;
        let count = 0;
        for (let top = 0; top < sets.length - 1; top++) {
            const bot = top + 1;
            if ((sets[top].has(a) && sets[top].has(b)) || (sets[bot].has(a) && sets[bot].has(b))) {
                continue;
            }
            const order = sets[top].has(a) && sets[bot].has(b);
            const flip = sets[top].has(b) && sets[bot].has(a);
            let nghich = order && !flip;
            if (flip && !order && !soleMainOldest && top > mainMax && top >= mainMax + 5) {
                nghich = true;
            }
            if (nghich) {
                count++;
            }
        }
        return count;
    }

    crossingKindAtTop(sets, top, a, b) {
        const bot = top + 1;
        if (!sets || top < 0 || bot >= sets.length) {
            return null;
        }
        if ((sets[top].has(a) && sets[top].has(b)) || (sets[bot].has(a) && sets[bot].has(b))) {
            return null;
        }
        const order = sets[top].has(a) && sets[bot].has(b);
        const flip = sets[top].has(b) && sets[bot].has(a);
        if (order && !flip) {
            return 'order';
        }
        if (flip && !order) {
            return 'flip';
        }
        return null;
    }

    numInCrossingAtTop(sets, top, kind, n, a, b) {
        const bot = top + 1;
        if (kind === 'order') {
            return (n === a && sets[top].has(a)) || (n === b && sets[bot].has(b));
        }
        if (kind === 'flip') {
            return (n === b && sets[top].has(b)) || (n === a && sets[bot].has(a));
        }
        return false;
    }

    sameNumRecalledAcrossCrossings(sets, top1, kind1, top2, kind2, a, b) {
        if (
            this.numInCrossingAtTop(sets, top1, kind1, b, a, b) &&
            this.numInCrossingAtTop(sets, top2, kind2, b, a, b)
        ) {
            return true;
        }
        if (
            this.numInCrossingAtTop(sets, top1, kind1, a, a, b) &&
            this.numInCrossingAtTop(sets, top2, kind2, a, a, b)
        ) {
            return true;
        }
        return false;
    }

    junkRecallBelowMain(sets, a, b, mainIdx) {
        let t0 = null;
        let k0 = null;
        for (let top = mainIdx; top <= mainIdx + 1 && top < sets.length - 1; top++) {
            const kind = this.crossingKindAtTop(sets, top, a, b);
            if (!kind) {
                continue;
            }
            t0 = top;
            k0 = kind;
            break;
        }
        if (t0 == null || t0 >= sets.length - 2) {
            return false;
        }
        const k1 = this.crossingKindAtTop(sets, t0 + 1, a, b);
        if (!k1) {
            return false;
        }
        return this.sameNumRecalledAcrossCrossings(sets, t0, k0, t0 + 1, k1, a, b);
    }

    junkRecallAboveMain(sets, a, b, mainIdx) {
        if (mainIdx < 1) {
            return false;
        }
        const t0 = mainIdx - 1;
        const k0 = this.crossingKindAtTop(sets, t0, a, b);
        if (k0) {
            const k1 = this.crossingKindAtTop(sets, t0 - 1, a, b);
            if (k1 && this.sameNumRecalledAcrossCrossings(sets, t0 - 1, k1, t0, k0, a, b)) {
                return true;
            }
        }
        if (mainIdx >= 2 && mainIdx < sets.length - 1) {
            let tFirst = null;
            let kFirst = null;
            for (let top = mainIdx - 1; top >= 0; top--) {
                const kind = this.crossingKindAtTop(sets, top, a, b);
                if (!kind) {
                    continue;
                }
                tFirst = top;
                kFirst = kind;
                break;
            }
            if (tFirst != null && tFirst <= mainIdx - 3 && kFirst === 'order') {
                const rowAbove = mainIdx - 1;
                for (const n of [a, b]) {
                    if (
                        this.numInCrossingAtTop(sets, tFirst, kFirst, n, a, b) &&
                        sets[rowAbove].has(n)
                    ) {
                        return true;
                    }
                }
            }
        }
        return false;
    }

    /**
     * Gọi lại quá sát: hai biên cắt kề (155) hoặc cắt xa hơn + số lại ở hàng ngay kề main (182 3:{12,22}).
     */
    junkPairAdjacentRecallTooClose(sets, a, b) {
        if (!sets || sets.length < 3 || !Number.isFinite(a) || !Number.isFinite(b)) {
            return false;
        }
        const mains = [];
        for (let idx = 0; idx < sets.length; idx++) {
            if (sets[idx].has(a) && sets[idx].has(b)) {
                mains.push(idx);
            }
        }
        for (const mainIdx of mains) {
            if (this.junkRecallBelowMain(sets, a, b, mainIdx)) {
                return true;
            }
            if (this.junkRecallAboveMain(sets, a, b, mainIdx)) {
                return true;
            }
        }
        return false;
    }

    /**
     * Chồng quanh biên: lật top=4 b trên freq4/2; thuận a trên f(a)=3 (top≠6 cần f(b)=2; top=6 rác khi f(b)≠2 và (a,b) cùng dòng idx 0); lật top=6 b trên f(b)=3 f(a)=2; lật b trên f(a)=f(b)=3 chỉ khi main duy nhất dòng sát kỳ; lật a dưới f(a)=3 (top=6 f(b)≤3; top≠6 f(b)=2 — 182 9:{3,11}).
     * Cặp có số ≥6 hoặc cả hai ≥5 trong cửa sổ → bỏ thêm (junkPairIfAnyWindowFreqGe / junkPairIfBothWindowFreqGe).
     */
    junkStackedPairNumAboveCrossingAdjacent(sets, a, b, windowFreq) {
        if (!sets || sets.length < 2 || !Number.isFinite(a) || !Number.isFinite(b)) {
            return false;
        }
        const wf = windowFreq || {};
        const f = (n) => (wf[n] != null ? Number(wf[n]) : 0) || 0;
        const mains = [];
        for (let idx = 0; idx < sets.length; idx++) {
            if (sets[idx].has(a) && sets[idx].has(b)) {
                mains.push(idx);
            }
        }
        const onlyMainOnNewestRow = mains.length === 1 && mains[0] === sets.length - 1;
        for (let top = 1; top < sets.length - 1; top++) {
            const bot = top + 1;
            if ((sets[top].has(a) && sets[top].has(b)) || (sets[bot].has(a) && sets[bot].has(b))) {
                continue;
            }
            const flip = sets[top].has(b) && sets[bot].has(a);
            const order = sets[top].has(a) && sets[bot].has(b);
            if (flip && !order && top === 4) {
                if (sets[top - 1].has(b) && sets[top].has(b) && f(b) === 4 && f(a) === 2) {
                    return true;
                }
            }
            if (order && !flip) {
                if (sets[top - 1].has(a) && sets[top].has(a) && f(a) === 3) {
                    if (top === 6) {
                        if (f(b) !== 2 && sets[0].has(a) && sets[0].has(b)) {
                            return true;
                        }
                    } else if (f(b) === 2) {
                        return true;
                    }
                }
            }
            if (flip && !order && top === 6) {
                if (sets[top - 1].has(b) && sets[top].has(b) && f(b) === 3 && f(a) === 2) {
                    return true;
                }
            }
            if (flip && !order && onlyMainOnNewestRow) {
                if (sets[top - 1].has(b) && sets[top].has(b) && f(b) === 3 && f(a) === 3) {
                    return true;
                }
            }
            if (flip && !order && bot + 1 < sets.length) {
                const aBelow = sets[bot].has(a) && sets[bot + 1].has(a) && f(a) === 3;
                if (aBelow && top === 6 && f(b) <= 3) {
                    return true;
                }
                if (aBelow && top !== 6 && f(b) === 2 && mains.some((m) => top >= m)) {
                    return true;
                }
            }
        }
        return false;
    }

    /**
     * Cặp chứa số xuất hiện ≥ ge lần trong cửa sổ 10 dòng → rác (vd. 609: 35×6).
     */
    junkPairIfAnyWindowFreqGe(windowFreq, a, b, ge) {
        const thr = ge != null ? ge : 6;
        if (!windowFreq || !Number.isFinite(a) || !Number.isFinite(b)) {
            return false;
        }
        const f = (n) => (windowFreq[n] != null ? Number(windowFreq[n]) : 0) || 0;
        return f(a) >= thr || f(b) >= thr;
    }

    /**
     * Cả hai số đều ≥ ge lần trong cửa sổ → rác (thống kê: không cùng freq≥5).
     */
    junkPairIfBothWindowFreqGe(windowFreq, a, b, ge) {
        const thr = ge != null ? ge : 5;
        if (!windowFreq || !Number.isFinite(a) || !Number.isFinite(b)) {
            return false;
        }
        const f = (n) => (windowFreq[n] != null ? Number(windowFreq[n]) : 0) || 0;
        return f(a) >= thr && f(b) >= thr;
    }

    /**
     * Compute visible pair candidates for a 10-row window using the same
     * rules as the left pane pair list.
     */
    computePairsForRows(rows) {
        if (!rows || rows.length < 10) {
            return [];
        }

        const display = rows.slice(0, 10);
        const sets = display.map(row => new Set(this.parseMainNums(row.result || row.Result || '')));
        const windowFreq = {};
        for (const set of sets) {
            for (const n of set) {
                windowFreq[n] = (windowFreq[n] || 0) + 1;
            }
        }
        const allNums = new Set();
        sets.forEach(set => set.forEach(num => allNums.add(num)));
        const nums = Array.from(allNums).sort((left, right) => left - right);
        const adjPairs = [];

        for (let topIdx = 0; topIdx < sets.length - 1; topIdx++) {
            adjPairs.push({ top: topIdx, bottom: topIdx + 1 });
        }

        const out = [];
        for (let i = 0; i < nums.length; i++) {
            for (let j = i + 1; j < nums.length; j++) {
                const a = nums[i];
                const b = nums[j];
                const mains = [];

                for (let idx = 0; idx < sets.length; idx++) {
                    if (sets[idx].has(a) && sets[idx].has(b)) {
                        mains.push(idx);
                    }
                }

                if (mains.length === 0) {
                    continue;
                }

                let allMainsOk = true;
                for (const mainIdx of mains) {
                    let aboveFound = false;
                    for (const pair of adjPairs) {
                        if (pair.bottom < mainIdx) {
                            const topHasPair = sets[pair.top].has(a) && sets[pair.bottom].has(b);
                            const flippedHasPair = sets[pair.top].has(b) && sets[pair.bottom].has(a);
                            if (topHasPair || flippedHasPair) {
                                if (!((sets[pair.top].has(a) && sets[pair.top].has(b)) || (sets[pair.bottom].has(a) && sets[pair.bottom].has(b)))) {
                                    const allowAboveIfFreq2 = (
                                        (windowFreq[a] >= 2 && windowFreq[b] >= 2 && mainIdx === sets.length - 1) ||
                                        ((windowFreq[a] === 3 && windowFreq[b] === 2) || (windowFreq[a] === 2 && windowFreq[b] === 3))
                                    );
                                    if (!allowAboveIfFreq2) {
                                        aboveFound = true;
                                        break;
                                    }
                                }
                            }
                        }
                    }
                    if (aboveFound) {
                        allMainsOk = false;
                        break;
                    }

                    let foundForThisMain = false;
                    for (const pair of adjPairs) {
                        if (!(pair.top > mainIdx && pair.bottom > mainIdx)) {
                            continue;
                        }
                        const topHasPair = sets[pair.top].has(a) && sets[pair.bottom].has(b);
                        const flippedHasPair = sets[pair.top].has(b) && sets[pair.bottom].has(a);
                        if (topHasPair || flippedHasPair) {
                            if ((sets[pair.top].has(a) && sets[pair.top].has(b)) || (sets[pair.bottom].has(a) && sets[pair.bottom].has(b))) {
                                continue;
                            }
                            foundForThisMain = true;
                            break;
                        }
                    }

                    if (!foundForThisMain) {
                        const allowAboveIfFreq2 = (
                            (windowFreq[a] >= 2 && windowFreq[b] >= 2 && mainIdx === sets.length - 1) ||
                            ((windowFreq[a] === 3 && windowFreq[b] === 2) || (windowFreq[a] === 2 && windowFreq[b] === 3))
                        );
                        if (allowAboveIfFreq2) {
                            for (const pair of adjPairs) {
                                if (!(pair.bottom < mainIdx)) {
                                    continue;
                                }
                                const topHasPair = sets[pair.top].has(a) && sets[pair.bottom].has(b);
                                const flippedHasPair = sets[pair.top].has(b) && sets[pair.bottom].has(a);
                                if (topHasPair || flippedHasPair) {
                                    if ((sets[pair.top].has(a) && sets[pair.top].has(b)) || (sets[pair.bottom].has(a) && sets[pair.bottom].has(b))) {
                                        continue;
                                    }
                                    foundForThisMain = true;
                                    break;
                                }
                            }
                        }
                    }

                    if (!foundForThisMain) {
                        allMainsOk = false;
                        break;
                    }
                }

                if (allMainsOk) {
                    if (this.countOrderOnlyAdjacentBoundaries(sets, a, b) >= 2) {
                        continue;
                    }
                    if (this.junkStackedPairNumAboveCrossingAdjacent(sets, a, b, windowFreq)) {
                        continue;
                    }
                    if (this.junkPairIfAnyWindowFreqGe(windowFreq, a, b, 6)) {
                        continue;
                    }
                    if (this.junkPairIfBothWindowFreqGe(windowFreq, a, b, 5)) {
                        continue;
                    }
                    if (this.junkPairAdjacentRecallTooClose(sets, a, b)) {
                        continue;
                    }
                    out.push([a, b]);
                }
            }
        }

        const seen = new Set();
        const uniq = [];
        out.forEach(pair => {
            const key = pair[0] + ',' + pair[1];
            if (!seen.has(key)) {
                seen.add(key);
                uniq.push(pair);
            }
        });
        return uniq;
    }

    /**
     * Convert a non-negative integer to superscript text.
     */
    toSuperscript(n) {
        const map = {
            '0': '\u2070',
            '1': '\u00b9',
            '2': '\u00b2',
            '3': '\u00b3',
            '4': '\u2074',
            '5': '\u2075',
            '6': '\u2076',
            '7': '\u2077',
            '8': '\u2078',
            '9': '\u2079'
        };
        return String(n).split('').map(ch => map[ch] || ch).join('');
    }

    /**
     * Escape text before injecting it into HTML.
     */
    escapeHtml(text) {
        return String(text)
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;')
            .replace(/'/g, '&#39;');
    }

    /**
     * Render result text with Module1-like styling after the pipe.
     */
    highlightResultByFrequency(result) {
        if (!result) return '';

        const pipeIndex = result.indexOf('|');
        if (pipeIndex === -1) {
            return result;
        }

        const beforePipe = result.substring(0, pipeIndex);
        const afterPipe = result.substring(pipeIndex + 1);
        const specialNums = this.parseNums(afterPipe);

        if (specialNums.length === 0) {
            return `${beforePipe}|${afterPipe}`;
        }

        const styledSpecial = specialNums.map(num => {
            return `<span style="font-size:1.5em;font-weight:bold;color:rgb(0,100,0)">${num}</span>`;
        }).join(',');

        return `${beforePipe}|${styledSpecial}`;
    }

    /**
     * Chỉ số dòng sheet1 có id (số) khớp targetNum (so khớp sau normalize).
     */
    findSourceSheetRowIndexByNumericId(targetNum) {
        const rows = this.getSourceSheetRows();
        const targetKey = this.normalizeNumberKey(targetNum);
        if (!targetKey) {
            return -1;
        }
        for (let i = 0; i < rows.length; i++) {
            if (this.normalizeNumberKey(rows[i].id || rows[i].ID || '') === targetKey) {
                return i;
            }
        }
        return -1;
    }

    /**
     * Chuột phải trên ô nonexist: chặn menu trình duyệt, focus hàng có id = id hàng gốc + NONEXIST_CONTEXTMENU_ID_DELTA.
     * @param {MouseEvent} event
     * @param {HTMLElement} tableWrap
     * @returns {boolean} true nếu đã xử lý (đã preventDefault)
     */
    handleNonexistCellContextMenu(event, tableWrap) {
        if (!event || this.activeSheet !== 'sheet1') {
            return false;
        }
        const td = event.target && event.target.closest && event.target.closest('td.cell-nonexist');
        if (!td || !tableWrap || !tableWrap.contains(td)) {
            return false;
        }
        const tr = td.closest('tr[data-idx]');
        if (!tr) {
            return false;
        }
        const idx = Number(tr.dataset.idx);
        if (!Number.isFinite(idx) || idx < 0) {
            return false;
        }
        const rows = this.getSourceSheetRows();
        if (idx >= rows.length) {
            return false;
        }
        const row = rows[idx];
        if (!row || this.isEmptyResultRow(row)) {
            return false;
        }
        event.preventDefault();
        event.stopPropagation();
        const currentIdNum = parseInt(String(row.id || row.ID || '').trim(), 10);
        if (!Number.isFinite(currentIdNum)) {
            return true;
        }
        const targetIdNum = currentIdNum + NONEXIST_CONTEXTMENU_ID_DELTA;
        const targetIdx = this.findSourceSheetRowIndexByNumericId(targetIdNum);
        if (targetIdx < 0 || targetIdx === idx) {
            return true;
        }
        const targetRow = rows[targetIdx];
        const targetEmpty = this.isEmptyResultRow(targetRow);
        this.onRowClick(targetIdx, targetEmpty, event, { fromFilterNav: tableWrap.id === 'filterTableWrap' });
        try {
            tableWrap.focus({ preventScroll: true });
        } catch (err) {
            /* ignore */
        }
        return true;
    }

    /**
     * Handle row click - dispatch to parent
     */
    onRowClick(idx, isEmptyRow, event, options = {}) {
        const windowTop = idx >= 10 ? idx - 10 : 0;
        const contextPrefixCount = idx >= 10 ? Math.min(2, windowTop) : 0;
        const dataStart = Math.max(0, windowTop - contextPrefixCount);
        const slice = isEmptyRow ? this.dataRows.slice(dataStart, idx) : this.dataRows.slice(dataStart, idx + 1);
        const lines = slice.map((r, offset) => {
            const res = r.result || r.Result || '';
            const noteMeta = this.getComputedNoteMeta(dataStart + offset, r);
            const note = noteMeta.text;
            const nonexist = this.isEmptyResultRow(r) ? '' : this.getComputedNonexistMeta(dataStart + offset, r).text;
            return [res, note, nonexist].filter(Boolean).join('\t');
        });

        const rowAtClick = this.dataRows[idx] || {};
        const clickedRowId = String(rowAtClick.id || rowAtClick.ID || '').trim();
        const focusNonexistHighlights = this.buildNonexistHighlightMapForRow(idx);

        // Update selectedLines
        this.selectedLines = slice.map((r, offset) => {
            const noteMeta = this.getComputedNoteMeta(dataStart + offset, r);
            const nonexistMeta = this.isEmptyResultRow(r) ? { text: '' } : this.getComputedNonexistMeta(dataStart + offset, r);
            return {
                date: r.date || '',
                id: r.id || '',
                result: r.result || '',
                note: noteMeta.text,
                nonexist: nonexistMeta.text
            };
        });
        if (isEmptyRow) {
            this.selectedLines.push({ date: '', id: '', result: '', note: '', nonexist: '' });
            lines.push('');

            // Keep the trailing row visually blank.
            try {
                const tableWrap = document.getElementById('tableWrap');
                if (tableWrap) {
                    const tr = tableWrap.querySelector(`tbody tr[data-idx="${idx}"]`);
                    if (tr) {
                        tr.dataset.empty = '1';
                    }
                }
            } catch (e) {
                // ignore DOM update failures
            }
        }

        if (this.activeSheet === 'sheet1') {
            const focusRow = this.dataRows[idx] || rowAtClick;
            const nextFocusId = String(focusRow.id || focusRow.ID || clickedRowId || '').trim();
            const hadG1 = this.comboG1Enabled;
            const comboStateChanged = this.comboFocusRowId !== nextFocusId
                || this.comboFocusRowIndex !== idx
                || (this.isEmptyResultRow(focusRow) && hadG1);
            this.comboFocusRowId = nextFocusId;
            this.comboFocusRowIndex = idx;
            if (this.isEmptyResultRow(focusRow)) {
                this.comboG1Enabled = false;
            }
            if (comboStateChanged) {
                window.dispatchEvent(new CustomEvent('comboControlsChanged', { detail: { sheet: this.activeSheet } }));
            }
        }

        if (this.activeSheet === 'sheet1' && !isEmptyRow) {
            try {
                this.syncSpecialTrackingTimelineFromSheet1Row(idx);
            } catch (eStSync) {
                /* ignore */
            }
        }

        if (!options.skipSave) {
            this.save();
        }

        const windowEnd = idx;
        const targetIdx = idx;
        this.applyWindowSelection(windowTop, windowEnd, targetIdx);

        const tableWrap = document.getElementById('tableWrap');
        const activeSheetMeta = this.sheets[this.activeSheet] || {};
        if (!options.skipCenter && tableWrap && activeSheetMeta.kind !== 'combo') {
            this.centerActiveWindowInView(tableWrap);
        }

        // Dispatch custom event with selected lines
        window.dispatchEvent(new CustomEvent('rowClicked', {
            detail: {
                selectedLines: this.selectedLines,
                selectedNums: this.parseNums(this.selectedLines.length > 0 ? (this.selectedLines[this.selectedLines.length - 1].result || '') : ''),
                sheetName: this.activeSheet,
                clickedRowId,
                focusRowIndex: idx,
                focusNonexistHighlights,
                fromFilterNav: !!options.fromFilterNav,
                light: !!options.light,
                contextPrefixCount
            }
        }));
    }

    /**
     * Get current selected lines
     */
    getSelectedLines() {
        return this.selectedLines || [];
    }

    /**
     * Create a new sheet
     */
    createSheet(sheetName) {
        if (this.sheets[sheetName]) {
            return false; // Already exists
        }
        this.sheets[sheetName] = { data: [], notes: {} };
        this.save();
        return true;
    }

    /**
     * Delete a sheet
     */
    deleteSheet(sheetName) {
        if (sheetName === 'sheet1') {
            return false; // Cannot delete default sheet1
        }
        if (sheetName === 'specialtracking') {
            return false;
        }
        if (!this.sheets[sheetName]) {
            return false; // Not found
        }
        delete this.sheets[sheetName];
        if (this.activeSheet === sheetName) {
            this.activeSheet = 'sheet1';
        }
        this.save();
        return true;
    }

    /**
     * Switch to a different sheet
     */
    switchSheet(sheetName) {
        if (!this.sheets[sheetName]) {
            return false;
        }
        this.activeSheet = sheetName;
        this.dataRows = this.sheets[sheetName].data || [];
        this.save();
        return true;
    }

    /**
     * Get all sheet names
     */
    getSheetNames() {
        return Object.keys(this.sheets);
    }

    /**
     * Render sheet tabs (like Excel)
     */
    renderSheetTabs(container) {
        container.innerHTML = '';
        const tabBar = document.createElement('div');
        tabBar.className = 'sheet-tabs-bar';

        const sheetNames = [
            'sheet1',
            'combo_1',
            'combo_2',
            'combo_3',
            'combo_4',
            'combo_5',
            'specialtracking'
        ];
        for (const name of sheetNames) {
            if (!this.sheets[name]) {
                continue;
            }
            const tab = document.createElement('button');
            tab.className = 'sheet-tab';
            if (name === this.activeSheet) {
                tab.classList.add('active');
            }
            tab.textContent = name === 'specialtracking' ? 'specialtracking' : name;
            tab.title = name === 'specialtracking' ? 'Theo dõi tần suất 12 số đặc biệt theo timeline' : name;
            tab.addEventListener('click', () => {
                const tableWrap = document.getElementById('tableWrap');
                if (tableWrap) {
                    this.setScrollPosition(this.activeSheet, tableWrap.scrollTop, tableWrap.scrollLeft);
                }
                this.switchSheet(name);
                window.dispatchEvent(new CustomEvent('sheetChanged', { detail: { sheet: name } }));
            });

            // Right-click context menu
            tab.addEventListener('contextmenu', (e) => {
                e.preventDefault();
                if (name !== 'sheet1' && name !== 'specialtracking') {
                    const confirmed = confirm(`Xóa sheet "${name}"?`);
                    if (confirmed) {
                        this.deleteSheet(name);
                        this.renderSheetTabs(container);
                        window.dispatchEvent(new CustomEvent('sheetChanged', { detail: { sheet: this.activeSheet } }));
                    }
                }
            });

            tabBar.appendChild(tab);
        }

        container.appendChild(tabBar);
    }

    /**
     * Tích lũy dict combo (1..5) + special từ một đoạn dòng (không chỉnh latestRow).
     */
    _accumulateComboDictsFromRows(rowsSlice, dicts, dictSpecial) {
        for (const row of rowsSlice || []) {
            const result = row.result || row.Result || '';
            if (!result) {
                continue;
            }
            const mainNums = this.parseMainNums(result);
            if (mainNums.length !== 5) {
                continue;
            }
            const special = this.parseSpecialPart(result);
            if (special) {
                dictSpecial.set(special, (dictSpecial.get(special) || 0) + 1);
            }
            for (const num of mainNums) {
                const key = String(num);
                dicts[1].set(key, (dicts[1].get(key) || 0) + 1);
            }
            for (let a = 0; a < 4; a++) {
                for (let b = a + 1; b < 5; b++) {
                    const key = `${mainNums[a]},${mainNums[b]}`;
                    dicts[2].set(key, (dicts[2].get(key) || 0) + 1);
                }
            }
            for (let a = 0; a < 3; a++) {
                for (let b = a + 1; b < 4; b++) {
                    for (let c = b + 1; c < 5; c++) {
                        const key = `${mainNums[a]},${mainNums[b]},${mainNums[c]}`;
                        dicts[3].set(key, (dicts[3].get(key) || 0) + 1);
                    }
                }
            }
            for (let a = 0; a < 2; a++) {
                for (let b = a + 1; b < 3; b++) {
                    for (let c = b + 1; c < 4; c++) {
                        for (let d = c + 1; d < 5; d++) {
                            const key = `${mainNums[a]},${mainNums[b]},${mainNums[c]},${mainNums[d]}`;
                            dicts[4].set(key, (dicts[4].get(key) || 0) + 1);
                        }
                    }
                }
            }
            const combo5Key = mainNums.join(',');
            dicts[5].set(combo5Key, (dicts[5].get(combo5Key) || 0) + 1);
        }
    }

    /**
     * Đóng góp của một key combo trong một dòng (khớp _accumulateComboDictsFromRows).
     */
    countComboSlotKeyContributionInRow(mainNums, comboSlot, comboKey) {
        if (!mainNums || mainNums.length !== 5 || comboKey == null || comboKey === '') {
            return 0;
        }
        if (comboSlot === 1) {
            let c = 0;
            for (let i = 0; i < mainNums.length; i++) {
                if (String(mainNums[i]) === comboKey) {
                    c++;
                }
            }
            return c;
        }
        if (comboSlot === 2) {
            for (let a = 0; a < 4; a++) {
                for (let b = a + 1; b < 5; b++) {
                    if (`${mainNums[a]},${mainNums[b]}` === comboKey) {
                        return 1;
                    }
                }
            }
            return 0;
        }
        if (comboSlot === 3) {
            for (let a = 0; a < 3; a++) {
                for (let b = a + 1; b < 4; b++) {
                    for (let c = b + 1; c < 5; c++) {
                        if (`${mainNums[a]},${mainNums[b]},${mainNums[c]}` === comboKey) {
                            return 1;
                        }
                    }
                }
            }
            return 0;
        }
        if (comboSlot === 4) {
            for (let a = 0; a < 2; a++) {
                for (let b = a + 1; b < 3; b++) {
                    for (let c = b + 1; c < 4; c++) {
                        for (let d = c + 1; d < 5; d++) {
                            if (`${mainNums[a]},${mainNums[b]},${mainNums[c]},${mainNums[d]}` === comboKey) {
                                return 1;
                            }
                        }
                    }
                }
            }
            return 0;
        }
        if (comboSlot === 5) {
            return mainNums.join(',') === comboKey ? 1 : 0;
        }
        return 0;
    }

    /**
     * Chỉ số dòng trong [startRow..endRowInclusive] khi tích lũy appear cho key đạt >= targetAppear.
     * Hòa điểm: ai đạt mốc trước (chỉ số nhỏ hơn) xếp trên — cùng luật special tracking.
     */
    rowIndexWhenComboAppearReached(rows, startRow, endRowInclusive, comboSlot, comboKey, targetAppear) {
        if (targetAppear <= 0) {
            return -1;
        }
        let sum = 0;
        const lo = Math.max(0, startRow | 0);
        const hi = Math.max(lo, endRowInclusive | 0);
        const arr = rows || [];
        for (let r = lo; r <= hi; r++) {
            const row = arr[r] || {};
            const mn = this.parseMainNums(row.result || row.Result || '');
            sum += this.countComboSlotKeyContributionInRow(mn, comboSlot, comboKey);
            if (sum >= targetAppear) {
                return r;
            }
        }
        return hi;
    }

    /** Chỉ số dòng khi tích count cho đúng chuỗi special (khớp dictSpecial). */
    rowIndexWhenSpecialCountReached(rows, startRow, endRowInclusive, specialKey, targetCount) {
        if (targetCount <= 0 || specialKey == null || specialKey === '') {
            return -1;
        }
        let sum = 0;
        const lo = Math.max(0, startRow | 0);
        const hi = Math.max(lo, endRowInclusive | 0);
        const arr = rows || [];
        for (let r = lo; r <= hi; r++) {
            const row = arr[r] || {};
            const sp = this.parseSpecialPart(row.result || row.Result || '');
            if (sp === specialKey) {
                sum++;
            }
            if (sum >= targetCount) {
                return r;
            }
        }
        return hi;
    }

    /**
     * Đóng góp theo key trong một hàng (slot 1..5), khớp _accumulateComboDictsFromRows.
     * @returns {Array<[string, number]>} cặp [key, delta] — số phần tử bị chặn (≤10).
     */
    getComboSlotRowContributionDeltas(mainNums, comboSlot) {
        if (!mainNums || mainNums.length !== 5) {
            return [];
        }
        const mn = mainNums;
        if (comboSlot === 1) {
            const freq = new Map();
            for (let i = 0; i < 5; i++) {
                const k = String(mn[i]);
                freq.set(k, (freq.get(k) || 0) + 1);
            }
            const out = [];
            freq.forEach((delta, k) => {
                out.push([k, delta]);
            });
            return out;
        }
        if (comboSlot === 2) {
            const out = [];
            for (let a = 0; a < 4; a++) {
                for (let b = a + 1; b < 5; b++) {
                    out.push([`${mn[a]},${mn[b]}`, 1]);
                }
            }
            return out;
        }
        if (comboSlot === 3) {
            const out = [];
            for (let a = 0; a < 3; a++) {
                for (let b = a + 1; b < 4; b++) {
                    for (let c = b + 1; c < 5; c++) {
                        out.push([`${mn[a]},${mn[b]},${mn[c]}`, 1]);
                    }
                }
            }
            return out;
        }
        if (comboSlot === 4) {
            const out = [];
            for (let a = 0; a < 2; a++) {
                for (let b = a + 1; b < 3; b++) {
                    for (let c = b + 1; c < 4; c++) {
                        for (let d = c + 1; d < 5; d++) {
                            out.push([`${mn[a]},${mn[b]},${mn[c]},${mn[d]}`, 1]);
                        }
                    }
                }
            }
            return out;
        }
        if (comboSlot === 5) {
            return [[mn.join(','), 1]];
        }
        return [];
    }

    /**
     * Một lượt quét dòng [0, endExclusive): chỉ số dòng đầu tiên mà cộng dồn đạt target trong targetByKey.
     * O(N × hằng số) thay vì O(K × N) khi gọi rowIndexWhenComboAppearReached từng key.
     */
    buildComboReachRowMapOnePass(rows, endExclusive, comboSlot, targetByKey, keysNeedingReach) {
        const reach = new Map();
        if (!keysNeedingReach || keysNeedingReach.size === 0) {
            return reach;
        }
        const running = new Map();
        const n = Math.max(0, endExclusive | 0);
        const hi = n > 0 ? n - 1 : -1;
        for (const key of keysNeedingReach) {
            running.set(key, 0);
            reach.set(key, null);
        }
        const arr = rows || [];
        for (let r = 0; r < n; r++) {
            const row = arr[r] || {};
            const mn = this.parseMainNums(row.result || row.Result || '');
            if (mn.length !== 5) {
                continue;
            }
            const deltas = this.getComboSlotRowContributionDeltas(mn, comboSlot);
            for (let i = 0; i < deltas.length; i++) {
                const key = deltas[i][0];
                const delta = deltas[i][1];
                if (!keysNeedingReach.has(key) || delta <= 0) {
                    continue;
                }
                const target = targetByKey.get(key) | 0;
                if (target <= 0) {
                    continue;
                }
                const prev = running.get(key) || 0;
                const next = prev + delta;
                if (reach.get(key) === null && prev < target && next >= target) {
                    reach.set(key, r);
                }
                running.set(key, next);
            }
        }
        if (hi >= 0) {
            for (const key of keysNeedingReach) {
                if (reach.get(key) === null) {
                    reach.set(key, hi);
                }
            }
        }
        return reach;
    }

    /**
     * Một lượt quét cho special: chỉ số dòng đầu tiên đạt count (khớp rowIndexWhenSpecialCountReached).
     */
    buildSpecialReachRowMapOnePass(rows, endExclusive, countByKey, keysNeedingReach) {
        const reach = new Map();
        if (!keysNeedingReach || keysNeedingReach.size === 0) {
            return reach;
        }
        const running = new Map();
        const n = Math.max(0, endExclusive | 0);
        const hi = n > 0 ? n - 1 : -1;
        for (const key of keysNeedingReach) {
            running.set(key, 0);
            reach.set(key, null);
        }
        const arr = rows || [];
        for (let r = 0; r < n; r++) {
            const row = arr[r] || {};
            const sp = this.parseSpecialPart(row.result || row.Result || '');
            if (!sp || !keysNeedingReach.has(sp)) {
                continue;
            }
            const target = countByKey.get(sp) | 0;
            if (target <= 0) {
                continue;
            }
            const prev = running.get(sp) || 0;
            const next = prev + 1;
            if (reach.get(sp) === null && prev < target && next >= target) {
                reach.set(sp, r);
            }
            running.set(sp, next);
        }
        if (hi >= 0) {
            for (const key of keysNeedingReach) {
                if (reach.get(key) === null) {
                    reach.set(key, hi);
                }
            }
        }
        return reach;
    }

    /**
     * Build Module2-style combo sheets from source rows.
     */
    buildComboSheetsFromRows(rows) {
        return this.buildComboSheetsFromRowsUpTo(rows, (rows || []).length);
    }

    /**
     * Giống buildComboSheetsFromRows nhưng chỉ dùng rows[0..endExclusive).
     */
    buildComboSheetsFromRowsUpTo(rows, endExclusive) {
        const dicts = [null, new Map(), new Map(), new Map(), new Map(), new Map()];
        const dictSpecial = new Map();
        const n = Math.max(0, endExclusive | 0);
        this._accumulateComboDictsFromRows((rows || []).slice(0, n), dicts, dictSpecial);

        const latestRow = this.getLatestValidResultRow((rows || []).slice(0, n));
        const latestNumbers = latestRow ? this.parseMainNums(latestRow.result || latestRow.Result || '') : [];
        const latestSpecial = latestRow ? this.parseSpecialPart(latestRow.result || latestRow.Result || '') : '';

        const comboSheets = {};
        for (let s = 1; s <= 5; s++) {
            const data = [];
            for (const [combo, appear] of dicts[s].entries()) {
                if (appear >= 2) {
                    data.push({ combo, appear, arrow: '' });
                }
            }

            /* Tie-break row index: một lượt O(N), tránh O(K×N) khi gọi rowIndexWhenComboAppearReached từng key. */
            const keysAppear2 = new Set();
            for (let di = 0; di < data.length; di++) {
                keysAppear2.add(data[di].combo);
            }
            const comboReachRow = this.buildComboReachRowMapOnePass(rows, n, s, dicts[s], keysAppear2);

            data.sort((left, right) => {
                if (right.appear !== left.appear) {
                    return right.appear - left.appear;
                }
                const ta = comboReachRow.get(left.combo);
                const tb = comboReachRow.get(right.combo);
                if (ta !== tb) {
                    return ta - tb;
                }
                return String(left.combo).localeCompare(String(right.combo));
            });

            if (s === 1 && latestNumbers.length === 5) {
                const latestSet = new Set(latestNumbers.map(num => String(num)));
                for (const row of data) {
                    if (this.comboKeyMatchesNumbers(row.combo, latestSet, 1)) {
                        row.arrow = '⬆';
                    }
                }
            }

            const sheet = {
                kind: 'combo',
                comboType: s,
                data,
                notes: {},
                latestId: latestRow ? (latestRow.id || latestRow.ID || '') : '',
                latestNumbers,
                latestSpecial
            };

            if (s === 1) {
                const specialRows = [];
                for (const [special, count] of dictSpecial.entries()) {
                    specialRows.push({ special, count, arrow: '' });
                }
                const keysSpecial = new Set(dictSpecial.keys());
                const specialReachRow = this.buildSpecialReachRowMapOnePass(rows, n, dictSpecial, keysSpecial);

                specialRows.sort((left, right) => {
                    if (right.count !== left.count) {
                        return right.count - left.count;
                    }
                    const ta = specialReachRow.get(left.special);
                    const tb = specialReachRow.get(right.special);
                    if (ta !== tb) {
                        return ta - tb;
                    }
                    return String(left.special).localeCompare(String(right.special));
                });

                if (latestSpecial) {
                    const latestSpecialSet = new Set(latestSpecial.split(',').map(item => this.normalizeNumberKey(item)).filter(Boolean));
                    for (const row of specialRows) {
                        if (latestSpecialSet.has(this.normalizeNumberKey(row.special))) {
                            row.arrow = '⬆';
                        }
                    }
                }

                sheet.specialRows = specialRows;
            }

            comboSheets[`combo_${s}`] = sheet;
        }

        return comboSheets;
    }

    /**
     * Find the latest row with a valid 5-number result.
     */
    getLatestValidResultRow(rows) {
        for (let index = (rows || []).length - 1; index >= 0; index--) {
            const row = rows[index] || {};
            const result = row.result || row.Result || '';
            if (this.parseMainNums(result).length === 5) {
                return row;
            }
        }
        return null;
    }

    /**
     * Extract the special part after the pipe in a result cell.
     */
    parseSpecialPart(result) {
        if (!result) {
            return '';
        }

        const parts = String(result).split('|');
        if (parts.length < 2) {
            return '';
        }

        return String(parts[1]).trim().replace(/^[,\s]+|[,\s]+$/g, '');
    }

    /**
     * Normalize number text for exact matching.
     */
    normalizeNumberKey(value) {
        const parsed = parseInt(String(value || '').trim(), 10);
        return Number.isNaN(parsed) ? '' : String(parsed);
    }

    /**
     * Check whether a combo key matches a set of numbers.
     */
    comboKeyMatchesNumbers(comboKey, numberSet, expectedCount) {
        const parts = String(comboKey || '').split(',').map(part => this.normalizeNumberKey(part)).filter(Boolean);
        if (parts.length !== expectedCount) {
            return false;
        }

        for (const part of parts) {
            if (!numberSet.has(part)) {
                return false;
            }
        }

        return true;
    }

    /**
     * Fill pair_to_ids using the same 10-row window + pair logic as buildNoteForRow
     * (diff = currentId - prevId for each prevId block). Matches ok.py semantics for column 2/3.
     */
    accumulatePairToIdsFromRowWindows(rows, pair_to_ids) {
        const list = rows || [];
        for (let rowIndex = 0; rowIndex < list.length; rowIndex++) {
            const currentRow = list[rowIndex] || {};
            const currentId = this.parseRowId(currentRow.id || currentRow.ID || '');
            const currentNums = this.parseMainNums(currentRow.result || currentRow.Result || '');
            if (currentId === null || currentNums.length !== 5) {
                continue;
            }
            const rid = this.normalizeNumberKey(currentRow.id || currentRow.ID || '');
            if (!rid) {
                continue;
            }

            const startIndex = Math.max(0, rowIndex - 10);
            const matchedNumbersByPrevId = new Map();

            for (let prevIndex = startIndex; prevIndex < rowIndex; prevIndex++) {
                const prevRow = list[prevIndex] || {};
                const prevId = this.parseRowId(prevRow.id || prevRow.ID || '');
                const prevNums = this.parseMainNums(prevRow.result || prevRow.Result || '');
                if (prevId === null || prevNums.length !== 5) {
                    continue;
                }

                for (let a = 0; a < 4; a++) {
                    for (let b = a + 1; b < 5; b++) {
                        if (this.pairExists(prevNums, currentNums[a], currentNums[b])) {
                            if (!matchedNumbersByPrevId.has(prevId)) {
                                matchedNumbersByPrevId.set(prevId, new Set());
                            }
                            matchedNumbersByPrevId.get(prevId).add(currentNums[a]);
                            matchedNumbersByPrevId.get(prevId).add(currentNums[b]);
                        }
                    }
                }
            }

            for (const [prevId, matchedNumberSet] of matchedNumbersByPrevId.entries()) {
                const uniq = [...new Set(Array.from(matchedNumberSet))];
                const diff = currentId - prevId;
                for (let i = 0; i < uniq.length; i++) {
                    for (let j = i + 1; j < uniq.length; j++) {
                        const a = Math.min(uniq[i], uniq[j]);
                        const b = Math.max(uniq[i], uniq[j]);
                        const key = `${a},${b}`;
                        if (!pair_to_ids[key]) {
                            pair_to_ids[key] = {};
                        }
                        if (!pair_to_ids[key][rid]) {
                            pair_to_ids[key][rid] = [];
                        }
                        pair_to_ids[key][rid].push(diff);
                    }
                }
            }
        }
    }

    /**
     * Build checknote-shaped data from the right pane source rows (same fields as ok.py __checknote_data from 535.xlsm).
     * id_to_result: draw id -> "n1,n2,n3,n4,n5" from result cell
     * pair_to_ids: "a,b" -> { idStr: [dist, ...] } from (1) regex on computed+raw notes `N:{...}` and
     * (2) structural accumulation from the same 10-row window logic as buildNoteForRow.
     */
    buildChecknoteDataFromSourceRows() {
        this.refreshDerivedState();
        const rows = this.sourceRows || [];
        const notes = (this.noteCache && this.noteCache.length === rows.length)
            ? this.noteCache
            : this.buildNotesFromRows(rows);

        const id_to_result = {};
        const pair_to_ids = {};
        const notePat = /(\d+)\s*:\s*\{([^}]*)\}/g;

        for (let r = 0; r < rows.length; r++) {
            const row = rows[r] || {};
            const rid = this.normalizeNumberKey(row.id || row.ID || '');
            if (!rid) {
                continue;
            }

            const main = this.parseMainNums(row.result || row.Result || '');
            if (main.length === 5) {
                id_to_result[rid] = main.join(',');
            }

            const meta = notes[r] || {};
            const rawNote = String(row.note || row.Note || '');
            const computed = (meta.text && meta.text !== '?') ? String(meta.text) : '';
            const note = [computed, rawNote].filter(Boolean).join(' ');

            let m;
            notePat.lastIndex = 0;
            while ((m = notePat.exec(note)) !== null) {
                const dist = parseInt(m[1], 10);
                if (Number.isNaN(dist)) {
                    continue;
                }
                const group = m[2] || '';
                const innerNums = String(group)
                    .split(/[\s,;:|]+/)
                    .map((x) => parseInt(String(x).trim(), 10))
                    .filter((n) => !Number.isNaN(n));
                const uniq = [...new Set(innerNums)];
                for (let i = 0; i < uniq.length; i++) {
                    for (let j = i + 1; j < uniq.length; j++) {
                        const a = Math.min(uniq[i], uniq[j]);
                        const b = Math.max(uniq[i], uniq[j]);
                        const key = `${a},${b}`;
                        if (!pair_to_ids[key]) {
                            pair_to_ids[key] = {};
                        }
                        if (!pair_to_ids[key][rid]) {
                            pair_to_ids[key][rid] = [];
                        }
                        pair_to_ids[key][rid].push(dist);
                    }
                }
            }
        }

        this.accumulatePairToIdsFromRowWindows(rows, pair_to_ids);

        let max_id = null;
        const ids = Object.keys(id_to_result).map((x) => parseInt(x, 10)).filter((n) => !Number.isNaN(n));
        if (ids.length) {
            max_id = Math.max(...ids);
        }

        return { id_to_result, pair_to_ids, max_id };
    }

    /**
     * Chuỗi các số đặc biệt 1–12 theo thứ tự thời gian (cột result, phần sau |).
     * `sourceRowIndices[k]` = chỉ số dòng sheet1 (0-based) tạo ra `series[k]`.
     */
    buildSpecialTrackingSeriesMeta(rows) {
        const series = [];
        const sourceRowIndices = [];
        const list = rows || [];
        for (let ri = 0; ri < list.length; ri++) {
            const row = list[ri];
            const raw = row.result || row.Result || '';
            const part = this.parseSpecialPart(raw);
            if (!part) {
                continue;
            }
            const tokens = String(part)
                .split(/[\s,;]+/)
                .map((t) => t.trim())
                .filter(Boolean);
            let n = null;
            for (const t of tokens) {
                const v = parseInt(t, 10);
                if (Number.isFinite(v) && v >= 1 && v <= 12) {
                    n = v;
                    break;
                }
            }
            if (n != null) {
                series.push(n);
                sourceRowIndices.push(ri);
            }
        }
        return { series, sourceRowIndices };
    }

    /**
     * Chuỗi các số đặc biệt 1–12 theo thứ tự thời gian (cột result, phần sau |).
     * Tham khảo vid.py: tách special và chỉ giữ kỳ có số hợp lệ.
     */
    buildSpecialTrackingSeries(rows) {
        return this.buildSpecialTrackingSeriesMeta(rows).series;
    }

    /**
     * Sheet1 → specialtracking một chiều: record id đang focus là X thì timeline tua tới
     * mốc đã xử lý đến kỳ có id (X−1) — số bước = số lần có special trên các dòng từ đầu đến dòng đó.
     * Scrub timeline / transport ST không gọi ngược lại sheet1.
     */
    syncSpecialTrackingTimelineFromSheet1Row(focusRowIndex) {
        const st = this.sheets && this.sheets.specialtracking;
        if (!st || st.kind !== 'specialtracking') {
            return;
        }
        const rows = this.sourceRows || [];
        const row = rows[focusRowIndex];
        if (!row) {
            return;
        }
        const idNum = parseInt(String(row.id != null ? row.id : row.ID || '').trim(), 10);
        if (!Number.isFinite(idNum)) {
            return;
        }
        this.ensureSpecialTrackingFrames(st);
        const series = st.series || [];
        const srcIx = st.seriesSourceRowIndices || [];
        const frames = st.frames || [];
        const total = frames.length;
        if (total < 1 || !series.length) {
            return;
        }
        if (srcIx.length !== series.length) {
            return;
        }

        const targetId = idNum - 1;
        let targetRowIdx = -1;
        if (targetId >= 1) {
            for (let i = 0; i < rows.length; i++) {
                const rid = parseInt(String(rows[i].id != null ? rows[i].id : rows[i].ID || '').trim(), 10);
                if (Number.isFinite(rid) && rid === targetId) {
                    targetRowIdx = i;
                    break;
                }
            }
            if (targetRowIdx < 0) {
                for (let i = rows.length - 1; i >= 0; i--) {
                    const rid = parseInt(String(rows[i].id != null ? rows[i].id : rows[i].ID || '').trim(), 10);
                    if (Number.isFinite(rid) && rid <= targetId) {
                        targetRowIdx = i;
                        break;
                    }
                }
            }
        }

        let frameIdx = 0;
        if (targetRowIdx >= 0) {
            let count = 0;
            for (let s = 0; s < srcIx.length; s++) {
                if (srcIx[s] <= targetRowIdx) {
                    count++;
                }
            }
            frameIdx = Math.max(0, count - 1);
        }
        frameIdx = Math.max(0, Math.min(total - 1, frameIdx));

        const tailRow = rows.length ? rows[rows.length - 1] : {};
        const tailId = String(tailRow.id ?? tailRow.ID ?? '');
        const uiSig = `${total}|${series.length}|${rows.length}|${tailId}`;
        const prev = st.specialTrackingUi && typeof st.specialTrackingUi === 'object' ? st.specialTrackingUi : {};
        const speed = Number.isFinite(prev.speed) ? Math.min(3, Math.max(0.5, prev.speed)) : 1;
        const snap = {
            sig: uiSig,
            frameIndex: frameIdx,
            playing: false,
            speed,
            predictNeonOn: !!prev.predictNeonOn,
            focusNum: prev.focusNum != null ? prev.focusNum : null
        };
        st.specialTrackingUi = snap;
        try {
            sessionStorage.setItem(SPECIAL_TRACKING_UI_STORAGE_KEY, JSON.stringify(snap));
        } catch (e) {
            /* ignore */
        }

        const tw = document.getElementById('tableWrap');
        if (tw && tw.classList.contains('table-wrap--specialtracking') && typeof tw.__specialTrackingSeekFrame === 'function') {
            tw.__specialTrackingSeekFrame(frameIdx);
        }
    }

    /**
     * Bước s trong [0..end] mà số n vừa đạt đúng lần xuất hiện thứ v (v≥1).
     */
    static specialTrackingStepOfVthHit(series, n, v, endInclusive) {
        if (v <= 0) {
            return -1;
        }
        let seen = 0;
        for (let s = 0; s <= endInclusive; s++) {
            if (series[s] === n) {
                seen++;
                if (seen === v) {
                    return s;
                }
            }
        }
        return endInclusive;
    }

    /**
     * Mỗi bước: tích lũy count + thứ tự giảm dần theo count.
     * Hòa điểm: ai đạt mốc count hiện tại *trước* (bước nhỏ hơn) đứng trên; ai đến sau đứng dưới.
     */
    buildSpecialTrackingFrames(series) {
        const counts = {};
        for (let i = 1; i <= 12; i++) {
            counts[i] = 0;
        }
        const frames = [];
        const list = series || [];
        for (let f = 0; f < list.length; f++) {
            const just = list[f];
            counts[just] += 1;
            const sorted = Object.keys(counts)
                .map((k) => {
                    const n = Number(k);
                    const v = counts[n];
                    const t = v > 0
                        ? RightPaneSheetManager.specialTrackingStepOfVthHit(list, n, v, f)
                        : -1;
                    return { n, v, t };
                })
                .sort((a, b) => (b.v - a.v) || (a.t - b.t) || (a.n - b.n))
                .map(({ n, v }) => ({ n, v }));
            const maxV = Math.max(1, sorted.length ? sorted[0].v : 1);
            /** @type {number[]} slot 0 = trên cùng (nhiều nhất) … 11 = dưới cùng */
            const slotByNum = new Array(13).fill(11);
            for (let s = 0; s < sorted.length; s++) {
                slotByNum[sorted[s].n] = s;
            }
            const wPctByNum = new Array(13).fill(0);
            for (let n = 1; n <= 12; n++) {
                wPctByNum[n] = (counts[n] / maxV) * 100;
            }
            frames.push({
                step: f + 1,
                justDrawn: just,
                byNum: { ...counts },
                sorted,
                maxV,
                slotByNum,
                wPctByNum
            });
        }
        return frames;
    }

    /** Hue 0…1 cho gradient hạng: hạng 1 (slot 0) = 2/12 (vàng), còn lại (slot+1)/12 */
    static specialTrackingRankHueT(slot) {
        const s = Number(slot);
        const t = Number.isFinite(s) ? Math.max(0, Math.min(11, Math.floor(s))) : 0;
        return t === 0 ? 2 / 12 : Math.min(1, (t + 1) / 12);
    }

    /**
     * Ứ viên “kỳ tiếp” khi timeline ở cuối — heuristic theo predict.txt (mở rộng):
     * rolling 20/50/100 (#4), gap (#7), velocity ~20 bước (#4, #8), bottom đang leo (#5),
     * lệch tích lũy so kỳ vọng N/12 (#6, #8), dưới median count toàn cục (#6),
     * mean rank gần đây × velocity (recovery), gia tốc hạng 10 vs 20 bước (#4),
     * Markov bậc 1 (chuẩn hoá min–max), cụm 4 kỳ, lệch tích lũy có trần, median có trần,
     * rolling 144 kỳ, bonus lặp kỳ liền trước, bonus “hồi âm” số ở N−3 / N−5 (#7).
     * Không random; không đảm bảo xác suất.
     * @param {number[]} series
     * @param {object[]} frames — full frames (length ≥ chuỗi)
     * @param {number|null} [prefixLen] — nếu set (≥1), chỉ dùng series[0..prefixLen-1] như độ dài N (dự đoán kỳ tiếp sau prefix đó)
     * @returns {number[]} 1–3 số: thường luôn top-3 điểm (sau heuristic đổi chỗ echo); t0≤0 thì chỉ #1.
     */
    static computeSpecialTrackingPredictRankStats(series, frames) {
        const nFull = series && series.length ? series.length : 0;
        if (nFull < 4 || !frames || frames.length < nFull) {
            return {
                total: 0,
                n1: 0,
                n2: 0,
                n3: 0,
                pct1: 0,
                pct2: 0,
                pct3: 0
            };
        }
        let n1 = 0;
        let n2 = 0;
        let n3 = 0;
        let tot = 0;
        for (let i = 3; i < nFull; i++) {
            const cand = RightPaneSheetManager.computeSpecialTrackingPredictCandidates(series, frames, i);
            const actual = series[i];
            tot++;
            if (cand[0] === actual) {
                n1++;
            } else if (cand.length > 1 && cand[1] === actual) {
                n2++;
            } else if (cand.length > 2 && cand[2] === actual) {
                n3++;
            }
        }
        return {
            total: tot,
            n1,
            n2,
            n3,
            pct1: tot ? (n1 / tot) * 100 : 0,
            pct2: tot ? (n2 / tot) * 100 : 0,
            pct3: tot ? (n3 / tot) * 100 : 0
        };
    }

    /**
     * Top-3 ứng viên đặc biệt (1–12) cho kỳ sau prefix — `SPECIAL_TRACKING_PREDICT_WT.predictMode`:
     * `temporalRank` (predict.txt: rolling/gap/rank-velocity/z), `blend`/`heuristic` (Markov+tuning), `globalTrigram`, `stackedNgram`.
     * @returns {number[]}
     */
    static computeSpecialTrackingPredictCandidates(series, frames, prefixLen = null) {
        const nFull = series && series.length ? series.length : 0;
        const N = prefixLen != null && prefixLen >= 1 ? Math.min(prefixLen, nFull) : nFull;
        if (!N || !frames || frames.length < 1) {
            return [];
        }
        const lastIdx = Math.min(frames.length, N) - 1;
        const last = frames[lastIdx];
        if (!last || !last.slotByNum || !last.byNum) {
            return [];
        }
        const wt = SPECIAL_TRACKING_PREDICT_WT;
        const predictMode =
            wt.predictMode === 'globalTrigram'
                ? 'globalTrigram'
                : wt.predictMode === 'stackedNgram'
                    ? 'stackedNgram'
                    : wt.predictMode === 'temporalRank'
                        ? 'temporalRank'
                        : wt.predictMode === 'blend'
                            ? 'blend'
                            : 'heuristic';
        if (predictMode === 'stackedNgram') {
            return specialTrackingComputeStackedNgramTop3(series, nFull, N, wt);
        }
        if (predictMode === 'globalTrigram' && nFull >= 3 && N >= 3) {
            const penD = series[N - 2];
            const prD = series[N - 1];
            if (penD >= 1 && penD <= 12 && prD >= 1 && prD <= 12) {
                const tg = specialTrackingTop3FromGlobalTrigram(series, nFull, penD, prD, N);
                if (tg) {
                    return tg;
                }
            }
        }
        if (predictMode === 'temporalRank') {
            return specialTrackingPredictTxtTemporalTop3(series, frames, N);
        }

        const K = Math.min(20, lastIdx);
        const past = K > 0 ? frames[lastIdx - K] : last;
        const K10 = Math.min(10, lastIdx);
        const past10 = K10 > 0 ? frames[lastIdx - K10] : last;

        const W20 = Math.min(20, N);
        const W50 = Math.min(50, N);
        const W100 = Math.min(100, N);
        const scores = new Array(13).fill(0);

        const recentCount = (n, W) => {
            let c = 0;
            const from = Math.max(0, N - W);
            for (let i = from; i < N; i++) {
                if (series[i] === n) {
                    c++;
                }
            }
            return c;
        };

        const gapSince = new Array(13).fill(N);
        for (let i = N - 1; i >= 0; i--) {
            const x = series[i];
            if (x >= 1 && x <= 12 && gapSince[x] === N) {
                gapSince[x] = N - 1 - i;
            }
        }

        const exp20 = W20 / 12;
        const exp50 = W50 / 12;
        const exp100 = W100 / 12;
        const idealCum = N / 12;

        const counts12 = [];
        for (let x = 1; x <= 12; x++) {
            counts12.push(last.byNum[x] || 0);
        }
        counts12.sort((a, b) => a - b);
        const medianCnt = (counts12[5] + counts12[6]) / 2;

        const Lmean = Math.min(60, lastIdx + 1);
        const meanSlotRecent = (n) => {
            if (Lmean < 1) {
                return last.slotByNum[n] ?? 11;
            }
            let sum = 0;
            const from = lastIdx - Lmean + 1;
            for (let j = from; j <= lastIdx; j++) {
                sum += frames[j].slotByNum[n] ?? 11;
            }
            return sum / Lmean;
        };

        const prevDraw = series[N - 1];
        const penDraw = N >= 2 ? series[N - 2] : prevDraw;

        const buildBigramNorm = (iLo, iHi) => {
            const follow = new Array(13).fill(0);
            let cnt = 0;
            const a = Math.max(0, iLo);
            const b = Math.min(N - 2, iHi);
            for (let i = a; i <= b; i++) {
                if (series[i] === prevDraw) {
                    cnt++;
                    const nx = series[i + 1];
                    if (nx >= 1 && nx <= 12) {
                        follow[nx]++;
                    }
                }
            }
            const alpha = wt.bgAlpha;
            const denom = cnt + 12 * alpha;
            let mn = Infinity;
            let mx = -Infinity;
            const raw = new Array(13).fill(0);
            for (let x = 1; x <= 12; x++) {
                const r = (follow[x] + alpha) / denom;
                raw[x] = r;
                if (r < mn) {
                    mn = r;
                }
                if (r > mx) {
                    mx = r;
                }
            }
            const sp = mx - mn || 1;
            const norm = new Array(13).fill(0);
            for (let x = 1; x <= 12; x++) {
                norm[x] = (raw[x] - mn) / sp;
            }
            return norm;
        };

        const normBgGlobal = buildBigramNorm(0, N - 2);
        const loRecent = Math.max(0, N - wt.recentWin);
        const normBgRecent = buildBigramNorm(loRecent, N - 2);

        let triCnt = 0;
        const triFollow = new Array(13).fill(0);
        const triWin = Math.min(wt.triWinCap, Math.max(0, N - 2));
        for (let i = Math.max(0, N - 2 - triWin); i < N - 2; i++) {
            if (series[i] === prevDraw && series[i + 1] === penDraw) {
                triCnt++;
                const nx = series[i + 2];
                if (nx >= 1 && nx <= 12) {
                    triFollow[nx]++;
                }
            }
        }
        const triAlpha = wt.triAlpha;
        const triDenom = triCnt + 12 * triAlpha;
        let triMn = Infinity;
        let triMx = -Infinity;
        const rawTri = new Array(13).fill(0);
        for (let x = 1; x <= 12; x++) {
            const r = (triFollow[x] + triAlpha) / triDenom;
            rawTri[x] = r;
            if (r < triMn) {
                triMn = r;
            }
            if (r > triMx) {
                triMx = r;
            }
        }
        const triSp = triMx - triMn || 1;
        const normTri = new Array(13).fill(0);
        for (let x = 1; x <= 12; x++) {
            normTri[x] = (rawTri[x] - triMn) / triSp;
        }

        const cMarg = new Array(13).fill(0);
        for (let i = 0; i < N; i++) {
            const x = series[i];
            if (x >= 1 && x <= 12) {
                cMarg[x]++;
            }
        }
        const alphaM =
            typeof wt.margAlpha === 'number' && Number.isFinite(wt.margAlpha) && wt.margAlpha > 0 ? wt.margAlpha : 2;
        const denM = N + 12 * alphaM;
        const margW =
            typeof wt.margLong === 'number' && Number.isFinite(wt.margLong) ? wt.margLong : 0;
        const margShortW =
            typeof wt.margShort === 'number' && Number.isFinite(wt.margShort) ? wt.margShort : 0;
        const margShortWin =
            typeof wt.margShortWin === 'number' && Number.isFinite(wt.margShortWin) && wt.margShortWin >= 8
                ? Math.floor(wt.margShortWin)
                : 72;
        const cShort = new Array(13).fill(0);
        let Ws = 0;
        if (margShortW !== 0 && N >= 1) {
            const fromS = Math.max(0, N - margShortWin);
            Ws = N - fromS;
            for (let i = fromS; i < N; i++) {
                const x = series[i];
                if (x >= 1 && x <= 12) {
                    cShort[x]++;
                }
            }
        }
        const denS = Ws + 12 * alphaM;

        for (let n = 1; n <= 12; n++) {
            let s = 0;
            const c20 = recentCount(n, W20);
            const c50 = recentCount(n, W50);
            const c100 = recentCount(n, W100);
            s += Math.max(0, exp50 - c50) * wt.uw50;
            s += Math.max(0, exp20 - c20) * wt.uw20;
            s += Math.max(0, exp100 - c100) * wt.uw100;

            const W4 = Math.min(4, N);
            const c4 = recentCount(n, W4);
            if (c4 >= 2) {
                s += wt.c4mul * (c4 - 1);
            }

            s += normBgGlobal[n] * wt.bgG + normBgRecent[n] * wt.bgR;
            if (N >= 3) {
                s += normTri[n] * wt.tri;
            }

            const g = gapSince[n];
            s += Math.sqrt(g + 1) * wt.gapSqrt;
            if (g > wt.gapThr) {
                s += (g - wt.gapThr) * wt.gapTail;
            }

            const slotNow = last.slotByNum[n] ?? 11;
            const slotPast = past.slotByNum[n] ?? 11;
            const vel = slotPast - slotNow;
            s += vel * wt.vel;

            if (slotNow >= wt.velHiSlot && vel > 0) {
                s += vel * wt.velHi;
            }

            const cum = last.byNum[n] || 0;
            /* Tích lũy dài dễ “đè” một số suốt 500+ kỳ — giới hạn đóng góp */
            s += Math.min(wt.cumCap, Math.max(0, idealCum - cum) * wt.cumK);
            if (cum < medianCnt) {
                s += Math.min(wt.medCap, (medianCnt - cum) * wt.medK);
            }

            const W144 = Math.min(144, N);
            const exp144 = W144 / 12;
            const c144 = recentCount(n, W144);
            s += Math.max(0, exp144 - c144) * wt.w144;

            const mSlot = meanSlotRecent(n);
            s += (mSlot / 11) * Math.max(0, vel) * wt.mslotVel;

            const slotPast10 = past10.slotByNum[n] ?? 11;
            const vel10 = slotPast10 - slotNow;
            if (vel10 > 0 && vel > 0 && vel10 > vel * wt.vel10mul) {
                s += (vel10 - vel) * wt.vel10a + wt.vel10b;
            }

            if (slotNow <= 1 && c50 > exp50 * wt.hot50rat) {
                s -= wt.penalHot50;
            }
            if (slotNow <= 2 && c100 > exp100 * wt.hot100rat) {
                s -= wt.penalHot100;
            }
            /* Số cách 2 kỳ (N−3) khác kỳ vừa ra: kiểu …8, 12, 3 → 8 (#7 cục bộ / “hồi âm”) */
            if (N >= 4) {
                const echo = series[N - 3];
                if (echo >= 1 && echo <= 12 && echo !== prevDraw && n === echo) {
                    s += wt.echo3;
                }
            }
            if (N >= 6) {
                const echo5 = series[N - 5];
                if (
                    echo5 >= 1
                    && echo5 <= 12
                    && echo5 !== prevDraw
                    && echo5 !== series[N - 3]
                    && n === echo5
                ) {
                    s += wt.echo5;
                }
            }
            /* Số ngay kề trước kỳ vừa ra (pen) — hỗ trợ …8→7, …7→8 */
            if (N >= 3 && penDraw >= 1 && penDraw <= 12 && penDraw !== prevDraw && n === penDraw) {
                s += wt.penB;
            }
            if (n === prevDraw) {
                s += wt.repeat;
            }
            if (margW !== 0 && N >= 1) {
                const phat = (cMarg[n] + alphaM) / denM;
                s += margW * Math.log(12 * phat);
            }
            if (margShortW !== 0 && Ws >= 1) {
                const phS = (cShort[n] + alphaM) / denS;
                s += margShortW * Math.log(12 * phS);
            }
            scores[n] = s;
        }

        if (predictMode === 'blend' && N >= 3) {
            const penG = series[N - 2];
            const prvG = series[N - 1];
            if (penG >= 1 && penG <= 12 && prvG >= 1 && prvG <= 12) {
                const byKey = specialTrackingGetGlobalTrigramByKey(series, nFull, N);
                const m = byKey.get(penG * 16 + prvG);
                const eff = typeof wt.triGlobal === 'number' && Number.isFinite(wt.triGlobal) && wt.triGlobal > 0 ? wt.triGlobal : 2.35;
                if (m) {
                    for (let nx = 1; nx <= 12; nx++) {
                        scores[nx] += eff * Math.log1p(m.get(nx) || 0);
                    }
                }
            }
        }

        const tMix =
            predictMode === 'blend' &&
                typeof wt.temporalMix === 'number' &&
                Number.isFinite(wt.temporalMix) &&
                wt.temporalMix !== 0
                ? wt.temporalMix
                : 0;
        if (tMix !== 0) {
            const tSc = specialTrackingComputeTemporalRankScores13(series, frames, N);
            if (tSc) {
                const zT = specialTrackingZScore12(tSc);
                for (let nx = 1; nx <= 12; nx++) {
                    scores[nx] += tMix * zT[nx];
                }
            }
        }

        const pairs = [];
        for (let n = 1; n <= 12; n++) {
            pairs.push([n, scores[n]]);
        }
        pairs.sort((a, b) => b[1] - a[1] || a[0] - b[0]);
        const echoN = N >= 4 ? series[N - 3] : 0;
        if (
            echoN >= 1
            && echoN <= 12
            && echoN !== series[N - 1]
            && pairs.length > 1
            && pairs[0][0] !== echoN
            && pairs[1][0] === echoN
            && pairs[0][1] - pairs[1][1] <= pairs[0][1] * wt.echoSwap
        ) {
            const t = pairs[0];
            pairs[0] = pairs[1];
            pairs[1] = t;
        }
        const t0 = pairs[0][1];
        const out = [pairs[0][0]];
        if (t0 <= 0) {
            return out;
        }
        /* Luôn top-3 (neon 3 bar + nhãn #1 #2 #3) để tối đa xác suất trúng trong 3 ô */
        if (pairs.length > 1) {
            out.push(pairs[1][0]);
        }
        if (pairs.length > 2) {
            out.push(pairs[2][0]);
        }
        return out;
    }

    ensureSpecialTrackingFrames(sheet) {
        if (!sheet || sheet.kind !== 'specialtracking') {
            return;
        }
        const fr0 = sheet.frames && sheet.frames[0];
        const framesOk = !!(sheet.frames && sheet.frames.length && fr0 && Array.isArray(fr0.slotByNum) && Array.isArray(fr0.wPctByNum));
        const srs = sheet.series || [];
        const idxOk = Array.isArray(sheet.seriesSourceRowIndices)
            && sheet.seriesSourceRowIndices.length === srs.length;
        if (framesOk && idxOk) {
            return;
        }
        const meta = this.buildSpecialTrackingSeriesMeta(this.sourceRows || []);
        sheet.series = meta.series;
        sheet.seriesSourceRowIndices = meta.sourceRowIndices;
        sheet.frames = this.buildSpecialTrackingFrames(meta.series);
    }

    renderSpecialTrackingShell(sheet) {
        this.ensureSpecialTrackingFrames(sheet);
        const frames = sheet.frames || [];
        if (!frames.length) {
            return (
                '<div class="special-tracking-root">'
                + '<div class="special-tracking-empty">Chưa có kỳ nào có số đặc biệt 1–12 sau dấu | trong cột result. '
                + 'Tải sheet1 dạng <code>…|7</code>.</div>'
                + '</div>'
            );
        }
        const total = frames.length;
        const fr0 = frames[0];
        let rankBarsHtml = '';
        for (let n = 1; n <= 12; n++) {
            const slot = fr0.slotByNum[n] ?? 0;
            const hueT = RightPaneSheetManager.specialTrackingRankHueT(slot);
            rankBarsHtml += `<div class="special-tracking-rank-bar" data-st-bar="${n}" data-special-num="${n}" style="--st-slot:${slot};--st-hue-t:${hueT}" role="button" tabindex="0" aria-label="Số ${n}, click để tô sáng">`
                + '<div class="special-tracking-rank-bar-main">'
                + '<div class="special-tracking-rank-track">'
                + '<span class="special-tracking-rank-fill" data-fill>'
                + `<span class="special-tracking-rank-num" data-st-num>${n}</span>`
                + '</span>'
                + '</div>'
                + '</div>'
                + '<div class="special-tracking-rank-meta">'
                + '<span class="special-tracking-rank-count" data-count>0</span>'
                + '<span class="special-tracking-rank-prio" data-st-predict-rank aria-hidden="true"></span>'
                + '</div>'
                + '</div>';
        }
        return (
            '<div class="special-tracking-root" data-st-root>'
            + '<div class="special-tracking-stage">'
            + '<div class="special-tracking-rank-wrap">'
            + '<div class="special-tracking-rank-shell">'
            + `<div class="special-tracking-rank-stack" data-st-rank-stack>${rankBarsHtml}</div>`
            + '</div>'
            + '</div>'
            + '</div>'
            + '<div class="special-tracking-controls">'
            + '<div class="special-tracking-controls-row special-tracking-controls-row--timeline">'
            + '<div class="special-tracking-timeline-head" data-st-timeline-head>'
            + '<div class="special-tracking-timeline-wrap">'
            + '<div class="special-tracking-timeline" data-st-timeline role="slider" aria-valuemin="0" aria-valuemax="'
            + (total - 1)
            + '" aria-valuenow="0" aria-label="Timeline">'
            + '<div class="special-tracking-timeline-fill" data-st-tl-fill></div>'
            + '<div class="special-tracking-timeline-thumb" data-st-tl-thumb></div>'
            + '<div class="special-tracking-timeline-rail" data-st-tl-hit></div>'
            + '</div>'
            + '<div class="special-tracking-timeline-meta-row">'
            + '<div class="special-tracking-meta-spacer" aria-hidden="true"></div>'
            + `<p class="special-tracking-step-label" data-st-step><strong>1</strong> / ${total}</p>`
            + '<div class="special-tracking-meta-right">'
            + '<button type="button" class="special-tracking-predict-toggle" data-st-predict-toggle '
            + 'aria-pressed="false" title="Bật/tắt neon dự đoán #1–#3 theo từng vị trí timeline" aria-label="Predict">Predict</button>'
            + '<div class="special-tracking-rank-stats" data-st-rank-stats aria-label="Tỷ lệ trúng theo hạng dự đoán toàn lịch sử"></div>'
            + '</div>'
            + '</div>'
            + '</div>'
            + '</div>'
            + '</div>'
            + '<div class="special-tracking-controls-row special-tracking-controls-row--transport">'
            + '<button type="button" class="special-tracking-icon-btn special-tracking-icon-btn--glyph" data-st-first title="Về đầu" aria-label="Về đầu"><span aria-hidden="true">\u23ee\uFE0E</span></button>'
            + '<button type="button" class="special-tracking-icon-btn special-tracking-icon-btn--glyph" data-st-prev title="Lùi 1 kỳ" aria-label="Lùi một kỳ"><span aria-hidden="true">\u23ea\uFE0E</span></button>'
            + '<button type="button" class="special-tracking-icon-btn special-tracking-icon-btn--play" data-st-play title="Phát" aria-label="Phát">'
            + '<svg class="special-tracking-svg-play" viewBox="0 0 24 24" width="28" height="28" aria-hidden="true"><path fill="currentColor" d="M9 6.5v11L18 12 9 6.5z"/></svg>'
            + '<svg class="special-tracking-svg-pause" viewBox="0 0 24 24" width="28" height="28" aria-hidden="true"><path fill="currentColor" d="M8 7h3v10H8V7zm5 0h3v10h-3V7z"/></svg></button>'
            + '<button type="button" class="special-tracking-icon-btn special-tracking-icon-btn--glyph" data-st-next title="Tiến 1 kỳ" aria-label="Tiến một kỳ"><span aria-hidden="true">\u23e9\uFE0E</span></button>'
            + '<button type="button" class="special-tracking-icon-btn special-tracking-icon-btn--glyph" data-st-last title="Cuối" aria-label="Đến cuối"><span aria-hidden="true">\u23ed\uFE0E</span></button>'
            + '<div class="special-tracking-speed-slider-wrap">'
            + '<span class="special-tracking-speed-hint">Tốc độ</span>'
            + '<input type="range" class="special-tracking-speed-slider" data-st-speed-slider '
            + 'min="0.5" max="3" step="0.5" value="1" aria-valuemin="0.5" aria-valuemax="3" aria-valuenow="1" aria-label="Tốc độ phát" />'
            + '<span class="special-tracking-speed-readout" data-st-speed-val>1×</span>'
            + '</div>'
            + '</div>'
            + '</div>'
            + `<input type="hidden" data-st-total value="${total}" />`
            + '</div>'
        );
    }

    wireSpecialTrackingUi(tableWrap, sheet) {
        this.ensureSpecialTrackingFrames(sheet);
        const frames = sheet.frames || [];
        const root = tableWrap.querySelector('[data-st-root]');
        if (!root || !frames.length) {
            tableWrap.__specialTrackingCleanup = null;
            return;
        }

        const series = Array.isArray(sheet.series) ? sheet.series : [];
        const total = frames.length;
        const srcRows = this.sourceRows || [];
        const tailRow = srcRows.length ? srcRows[srcRows.length - 1] : {};
        const tailId = String(tailRow.id ?? tailRow.ID ?? '');
        const uiSig = `${total}|${series.length}|${srcRows.length}|${tailId}`;

        const readSavedStUi = () => {
            let u = sheet.specialTrackingUi;
            if (u && u.sig === uiSig) {
                return u;
            }
            try {
                const raw = sessionStorage.getItem(SPECIAL_TRACKING_UI_STORAGE_KEY);
                if (!raw) {
                    return null;
                }
                const o = JSON.parse(raw);
                if (o && o.sig === uiSig) {
                    return o;
                }
            } catch (e) {
                /* ignore */
            }
            return null;
        };
        const savedSt = readSavedStUi();
        const clampStIdx = (i) => {
            const n = Number(i);
            if (!Number.isFinite(n)) {
                return 0;
            }
            return Math.max(0, Math.min(total - 1, Math.floor(n)));
        };

        let frameIndex = savedSt ? clampStIdx(savedSt.frameIndex) : 0;
        let playing = savedSt ? !!savedSt.playing : false;
        let speed = 1;
        if (savedSt && Number.isFinite(savedSt.speed)) {
            speed = Math.min(3, Math.max(0.5, savedSt.speed));
        }
        let focusNum = null;
        if (savedSt && savedSt.focusNum != null) {
            const fn = Number(savedSt.focusNum);
            if (Number.isFinite(fn) && fn >= 1 && fn <= 12) {
                focusNum = fn;
            }
        }
        let predictNeonOn = savedSt ? !!savedSt.predictNeonOn : false;
        let lastPredictNeonSyncKey = '';

        const progPct = new Float32Array(total);
        for (let i = 0; i < total; i++) {
            progPct[i] = total <= 1 ? 100 : (i / (total - 1)) * 100;
        }

        let timer = null;
        const BASE_MS = 420;

        const persistSpecialTrackingUi = () => {
            try {
                const snap = {
                    sig: uiSig,
                    frameIndex,
                    playing,
                    speed,
                    predictNeonOn,
                    focusNum
                };
                sheet.specialTrackingUi = snap;
                sessionStorage.setItem(SPECIAL_TRACKING_UI_STORAGE_KEY, JSON.stringify(snap));
            } catch (e) {
                /* ignore quota / private mode */
            }
        };

        const btnFirst = root.querySelector('[data-st-first]');
        const btnPrev = root.querySelector('[data-st-prev]');
        const btnPlay = root.querySelector('[data-st-play]');
        const btnNext = root.querySelector('[data-st-next]');
        const btnLast = root.querySelector('[data-st-last]');
        const syncPlayBtnUi = () => {
            if (!btnPlay) {
                return;
            }
            btnPlay.classList.toggle('is-playing', playing);
            btnPlay.setAttribute('aria-label', playing ? 'Tạm dừng' : 'Phát');
            btnPlay.setAttribute('title', playing ? 'Tạm dừng' : 'Phát');
        };
        const stepEl = root.querySelector('[data-st-step]');
        const predictToggle = root.querySelector('[data-st-predict-toggle]');
        const statsRankEl = root.querySelector('[data-st-rank-stats]');
        if (predictToggle) {
            predictToggle.classList.toggle('is-on', predictNeonOn);
            predictToggle.setAttribute('aria-pressed', predictNeonOn ? 'true' : 'false');
        }

        const formatRankStats = (st) => {
            if (!st || st.total < 1) {
                return '—';
            }
            return `#1 ${st.pct1.toFixed(1)}% · #2 ${st.pct2.toFixed(1)}% · #3 ${st.pct3.toFixed(1)}%`;
        };
        const rankStatsTitle = (st) => {
            if (!st || st.total < 1) {
                return '';
            }
            return (
                '#1/#2/#3 ở đây là thứ tự theo điểm heuristic (trọng số + tie-break + hoán echo #1↔#2 khi gần điểm), '
                + 'không phải thứ tự theo tần suất trúng trên lịch sử. '
                + 'Ba % là ba nhóm loại trừ: mỗi kỳ chỉ đếm một ô — nên % ô #3 hoàn toàn có thể cao hơn % ô #1. '
                + `Backtest: ${st.total} bước.`
            );
        };
        if (statsRankEl && series.length === total && frames.length === total) {
            const st = RightPaneSheetManager.computeSpecialTrackingPredictRankStats(series, frames);
            statsRankEl.textContent = formatRankStats(st);
            statsRankEl.title = rankStatsTitle(st);
        } else if (statsRankEl) {
            statsRankEl.textContent = '—';
            statsRankEl.removeAttribute('title');
        }

        const onPredictToggle = () => {
            predictNeonOn = !predictNeonOn;
            if (predictToggle) {
                predictToggle.classList.toggle('is-on', predictNeonOn);
                predictToggle.setAttribute('aria-pressed', predictNeonOn ? 'true' : 'false');
            }
            paint();
        };
        if (predictToggle) {
            predictToggle.addEventListener('click', onPredictToggle);
        }

        const tl = root.querySelector('[data-st-timeline]');
        const tlFill = root.querySelector('[data-st-tl-fill]');
        const tlThumb = root.querySelector('[data-st-tl-thumb]');
        const tlHit = root.querySelector('[data-st-tl-hit]');
        const speedSlider = root.querySelector('[data-st-speed-slider]');
        const speedVal = root.querySelector('[data-st-speed-val]');
        const formatSpeed = (v) => {
            const n = Number(v);
            if (!Number.isFinite(n)) {
                return '1×';
            }
            return `${Number.isInteger(n) ? n : n.toFixed(1)}×`;
        };
        const syncSpeedUi = () => {
            if (speedSlider) {
                speedSlider.setAttribute('aria-valuenow', String(speed));
            }
            if (speedVal) {
                speedVal.textContent = formatSpeed(speed);
            }
        };
        /** @type {Record<number, HTMLElement>} */
        const barByNum = {};
        root.querySelectorAll('[data-st-bar]').forEach((el) => {
            const n = parseInt(el.getAttribute('data-st-bar'), 10);
            if (Number.isFinite(n)) {
                barByNum[n] = el;
            }
        });

        let tlRectCache = null;
        const refreshTlRect = () => {
            if (tl) {
                tlRectCache = tl.getBoundingClientRect();
            }
        };

        let paintRaf = 0;
        const clearTimer = () => {
            if (timer) {
                clearTimeout(timer);
                timer = null;
            }
        };

        const scheduleNext = () => {
            clearTimer();
            if (!playing || frameIndex >= total - 1) {
                if (frameIndex >= total - 1) {
                    playing = false;
                    if (btnPlay) {
                        syncPlayBtnUi();
                    }
                }
                return;
            }
            const ms = Math.max(40, BASE_MS / speed);
            timer = setTimeout(() => {
                timer = null;
                frameIndex += 1;
                paint();
                scheduleNext();
            }, ms);
        };

        const setScrubbing = (on) => {
            root.classList.toggle('special-tracking-root--scrubbing', on);
        };

        const paint = () => {
            const fr = frames[frameIndex];
            if (!fr) {
                return;
            }
            const p = progPct[frameIndex];
            const canRetroPredict = series.length === total && frames.length === total;
            let predictList = [];
            if (predictNeonOn && canRetroPredict) {
                const prefixLen = frameIndex + 1;
                if (prefixLen >= 3) {
                    predictList = RightPaneSheetManager.computeSpecialTrackingPredictCandidates(
                        series,
                        frames,
                        prefixLen
                    );
                }
            }
            let actualNext = null;
            if (frameIndex + 1 < series.length) {
                actualNext = series[frameIndex + 1];
            }
            /** @type {Map<number, number>} số → thứ hạng dự đoán 1..3 */
            const predictRankByNum = new Map();
            if (predictList.length) {
                predictList.forEach((pn, idx) => {
                    predictRankByNum.set(pn, idx + 1);
                });
            }
            const predictNeonActive = predictNeonOn && predictList.length > 0;
            root.classList.toggle('special-tracking-root--predict-neon-on', predictNeonActive);
            if (!predictNeonActive) {
                lastPredictNeonSyncKey = '';
                root.style.removeProperty('--st-predict-neon-delay');
            } else {
                const syncKey = `${frameIndex}:${predictList.join(',')}:${actualNext == null ? '-' : String(actualNext)}`;
                if (syncKey !== lastPredictNeonSyncKey) {
                    lastPredictNeonSyncKey = syncKey;
                    const periodSec = 2 * 0.72;
                    const phase = (performance.now() / 1000) % periodSec;
                    root.style.setProperty('--st-predict-neon-delay', `${-phase}s`);
                }
            }

            for (let n = 1; n <= 12; n++) {
                const el = barByNum[n];
                if (!el) {
                    continue;
                }
                const slot = fr.slotByNum[n] ?? 0;
                el.style.setProperty('--st-slot', String(slot));
                const hueT = RightPaneSheetManager.specialTrackingRankHueT(slot);
                el.style.setProperty('--st-hue-t', String(hueT));
                const fillEl = el.querySelector('[data-fill]');
                const countEl = el.querySelector('[data-count]');
                const prioEl = el.querySelector('[data-st-predict-rank]');
                if (countEl) {
                    countEl.textContent = String(fr.byNum[n] || 0);
                }
                if (fillEl) {
                    fillEl.style.width = `${fr.wPctByNum[n]}%`;
                }
                const pr = predictRankByNum.get(n);
                const predictHit = Boolean(pr) && actualNext != null && n === actualNext;
                el.classList.toggle('special-tracking-rank-bar--just', n === fr.justDrawn);
                el.classList.toggle('special-tracking-rank-bar--focus', focusNum != null && n === focusNum);
                el.classList.toggle('special-tracking-rank-bar--predict', Boolean(pr));
                el.classList.toggle('special-tracking-rank-bar--predict-hit', predictHit);
                el.classList.toggle('special-tracking-rank-bar--predict-1', pr === 1);
                el.classList.toggle('special-tracking-rank-bar--predict-2', pr === 2);
                el.classList.toggle('special-tracking-rank-bar--predict-3', pr === 3);
                if (prioEl) {
                    prioEl.textContent = pr ? `#${pr}` : '';
                    prioEl.setAttribute('aria-hidden', pr ? 'false' : 'true');
                }
                let aria = `Số ${n}, click để tô sáng`;
                if (pr) {
                    aria = predictHit
                        ? `Số ${n}, ứng viên dự đoán hạng ${pr}, trùng đáp án kỳ tiếp theo`
                        : `Số ${n}, ứng viên dự đoán hạng ${pr}, click để tô sáng`;
                }
                el.setAttribute('aria-label', aria);
            }
            if (stepEl) {
                stepEl.innerHTML = `<strong>${fr.step}</strong> / ${total}`;
            }
            if (tlFill) {
                tlFill.style.width = `${p}%`;
            }
            if (tlThumb) {
                tlThumb.style.left = `${p}%`;
            }
            if (tl) {
                tl.setAttribute('aria-valuenow', String(frameIndex));
            }
        };

        const schedulePaint = () => {
            if (paintRaf) {
                cancelAnimationFrame(paintRaf);
            }
            paintRaf = requestAnimationFrame(() => {
                paintRaf = 0;
                paint();
            });
        };

        const setFrame = (idx) => {
            frameIndex = Math.max(0, Math.min(total - 1, idx));
            paint();
            scheduleNext();
        };

        const setFrameRaf = (idx) => {
            frameIndex = Math.max(0, Math.min(total - 1, idx));
            schedulePaint();
            scheduleNext();
        };

        tableWrap.__specialTrackingSeekFrame = (idx) => {
            setFrameRaf(clampStIdx(idx));
            persistSpecialTrackingUi();
        };

        const togglePlay = () => {
            if (frameIndex >= total - 1) {
                frameIndex = 0;
            }
            playing = !playing;
            if (btnPlay) {
                syncPlayBtnUi();
            }
            scheduleNext();
        };

        const onTimelineSeek = (clientX) => {
            if (!tl) {
                return;
            }
            const r = tlRectCache || tl.getBoundingClientRect();
            const x = Math.min(Math.max(0, clientX - r.left), r.width);
            const t = r.width > 0 ? x / r.width : 0;
            const idx = Math.round(t * (total - 1));
            setFrameRaf(idx);
        };

        let scrubDrag = false;
        const onScrubMove = (ev) => {
            if (!scrubDrag) {
                return;
            }
            const cx = ev.clientX != null ? ev.clientX : (ev.touches && ev.touches[0] ? ev.touches[0].clientX : 0);
            onTimelineSeek(cx);
        };
        const onScrubUp = () => {
            scrubDrag = false;
            setScrubbing(false);
            tlRectCache = null;
            window.removeEventListener('mousemove', onScrubMove, true);
            window.removeEventListener('mouseup', onScrubUp, true);
            window.removeEventListener('touchmove', onScrubTouchMove, true);
            window.removeEventListener('touchend', onScrubTouchEnd, true);
            window.removeEventListener('touchcancel', onScrubTouchEnd, true);
        };

        const onScrubTouchMove = (ev) => {
            if (!scrubDrag || !ev.touches || !ev.touches[0]) {
                return;
            }
            onTimelineSeek(ev.touches[0].clientX);
        };

        const onScrubTouchEnd = () => {
            onScrubUp();
        };

        const onTlDown = (ev) => {
            const cx = ev.clientX != null ? ev.clientX : (ev.touches && ev.touches[0] ? ev.touches[0].clientX : 0);
            scrubDrag = true;
            setScrubbing(true);
            refreshTlRect();
            window.addEventListener('mousemove', onScrubMove, true);
            window.addEventListener('mouseup', onScrubUp, true);
            window.addEventListener('touchmove', onScrubTouchMove, { passive: true, capture: true });
            window.addEventListener('touchend', onScrubTouchEnd, true);
            window.addEventListener('touchcancel', onScrubTouchEnd, true);
            onTimelineSeek(cx);
            playing = false;
            if (btnPlay) {
                syncPlayBtnUi();
            }
            clearTimer();
        };
        if (tlHit) {
            tlHit.addEventListener('mousedown', onTlDown);
            tlHit.addEventListener('touchstart', onTlDown, { passive: true });
        }

        if (btnFirst) {
            btnFirst.addEventListener('click', () => {
                playing = false;
                if (btnPlay) {
                    syncPlayBtnUi();
                }
                setFrame(0);
            });
        }
        if (btnLast) {
            btnLast.addEventListener('click', () => {
                playing = false;
                if (btnPlay) {
                    syncPlayBtnUi();
                }
                setFrame(total - 1);
            });
        }
        if (btnPrev) {
            btnPrev.addEventListener('click', () => {
                setFrame(frameIndex - 1);
            });
        }
        if (btnNext) {
            btnNext.addEventListener('click', () => {
                setFrame(frameIndex + 1);
            });
        }
        if (btnPlay) {
            btnPlay.addEventListener('click', togglePlay);
        }

        if (speedSlider) {
            speedSlider.value = String(speed);
            speedSlider.addEventListener('input', () => {
                const v = Number(speedSlider.value);
                speed = Number.isFinite(v) ? v : 1;
                syncSpeedUi();
                if (playing) {
                    scheduleNext();
                }
            });
            const initV = Number(speedSlider.value);
            speed = Number.isFinite(initV) ? initV : 1;
            syncSpeedUi();
        }

        root.querySelectorAll('[data-st-bar]').forEach((row) => {
            const onPick = () => {
                const n = parseInt(row.dataset.specialNum, 10);
                if (!Number.isFinite(n)) {
                    return;
                }
                focusNum = focusNum === n ? null : n;
                paint();
            };
            row.addEventListener('click', onPick);
            row.addEventListener('keydown', (e) => {
                if (e.key === 'Enter' || e.key === ' ') {
                    e.preventDefault();
                    onPick();
                }
            });
        });

        const onSpecialTrackingArrowNav = (ev) => {
            if (!tableWrap.classList.contains('table-wrap--specialtracking')) {
                return;
            }
            if (ev.key !== 'ArrowLeft' && ev.key !== 'ArrowRight') {
                return;
            }
            if (ev.ctrlKey || ev.metaKey || ev.altKey) {
                return;
            }
            const t = ev.target;
            if (t instanceof Element && t.closest('input, textarea, select, [contenteditable="true"]')) {
                return;
            }
            ev.preventDefault();
            if (ev.key === 'ArrowLeft') {
                setFrame(frameIndex - 1);
            } else {
                setFrame(frameIndex + 1);
            }
        };
        window.addEventListener('keydown', onSpecialTrackingArrowNav, true);

        paint();
        if (btnPlay) {
            syncPlayBtnUi();
        }
        if (playing) {
            scheduleNext();
        }

        tableWrap.__specialTrackingCleanup = () => {
            try {
                delete tableWrap.__specialTrackingSeekFrame;
            } catch (eDelSeek) {
                /* ignore */
            }
            persistSpecialTrackingUi();
            clearTimer();
            scrubDrag = false;
            setScrubbing(false);
            tlRectCache = null;
            if (paintRaf) {
                cancelAnimationFrame(paintRaf);
                paintRaf = 0;
            }
            window.removeEventListener('mousemove', onScrubMove, true);
            window.removeEventListener('mouseup', onScrubUp, true);
            if (tlHit) {
                tlHit.removeEventListener('mousedown', onTlDown);
                tlHit.removeEventListener('touchstart', onTlDown);
            }
            window.removeEventListener('touchmove', onScrubTouchMove, true);
            window.removeEventListener('touchend', onScrubTouchEnd, true);
            window.removeEventListener('touchcancel', onScrubTouchEnd, true);
            if (predictToggle) {
                predictToggle.removeEventListener('click', onPredictToggle);
            }
            window.removeEventListener('keydown', onSpecialTrackingArrowNav, true);
        };
    }

    /**
     * Save state to localStorage
     */
    save() {
        let sheetsForSave = this.sheets;
        const st = this.sheets && this.sheets.specialtracking;
        if (st && st.kind === 'specialtracking' && (st.frames || st.series)) {
            sheetsForSave = { ...this.sheets, specialtracking: { kind: 'specialtracking', data: [] } };
        }
        const data = {
            sheets: sheetsForSave,
            activeSheet: this.activeSheet,
            comboFocusRowId: this.comboFocusRowId,
            comboFocusRowIndex: this.comboFocusRowIndex,
            comboG1Enabled: this.comboG1Enabled,
            comboH1Text: this.comboH1Text,
            comboHComments: this.comboHComments || {},
            scrollPositions: this.scrollPositions
        };
        try {
            localStorage.setItem('sheetData', JSON.stringify(data));
        } catch (e) {
            console.warn('LocalStorage save failed:', e);
        }
    }

    /**
     * Export current sheet as JSON
     */
    exportSheet(sheetName) {
        const sheet = this.sheets[sheetName];
        if (!sheet) return null;
        return JSON.stringify(sheet, null, 2);
    }

    /**
     * Import data into a sheet
     */
    importSheet(sheetName, jsonData) {
        try {
            const data = JSON.parse(jsonData);
            if (!this.sheets[sheetName]) {
                this.sheets[sheetName] = { data: [], notes: {} };
            }
            this.sheets[sheetName] = data;
            this.save();
            return true;
        } catch (e) {
            console.error('Import failed:', e);
            return false;
        }
    }
}

let prevRecallFoldTooltipEl = null;
let prevRecallFoldTooltipShowTimer = null;
let prevRecallFoldTooltipHoverTarget = null;
const PREV_RECALL_FOLD_TOOLTIP_SHOW_MS = 35;

function ensurePrevRecallFoldTooltipEl() {
    if (prevRecallFoldTooltipEl && prevRecallFoldTooltipEl.isConnected) {
        return prevRecallFoldTooltipEl;
    }
    prevRecallFoldTooltipEl = document.getElementById('prevRecallFoldTooltip');
    if (!prevRecallFoldTooltipEl) {
        prevRecallFoldTooltipEl = document.createElement('div');
        prevRecallFoldTooltipEl.id = 'prevRecallFoldTooltip';
        prevRecallFoldTooltipEl.className = 'prev-recall-fold-tooltip';
        prevRecallFoldTooltipEl.setAttribute('role', 'tooltip');
        document.body.appendChild(prevRecallFoldTooltipEl);
    }
    return prevRecallFoldTooltipEl;
}

function hidePrevRecallFoldTooltip() {
    clearTimeout(prevRecallFoldTooltipShowTimer);
    prevRecallFoldTooltipShowTimer = null;
    prevRecallFoldTooltipHoverTarget = null;
    if (prevRecallFoldTooltipEl) {
        prevRecallFoldTooltipEl.classList.remove('is-visible');
    }
}

function showPrevRecallFoldTooltip(hit, text, clientX, clientY) {
    const tip = ensurePrevRecallFoldTooltipEl();
    tip.textContent = text;
    tip.classList.add('is-visible');

    const rect = hit.getBoundingClientRect();
    const offsetX = 16;
    const offsetY = 20;
    let left = (typeof clientX === 'number' ? clientX : rect.right) + offsetX;
    let top = (typeof clientY === 'number' ? clientY : rect.top) + offsetY;

    const pad = 6;
    const tipW = tip.offsetWidth;
    const tipH = tip.offsetHeight;
    const vw = window.innerWidth;
    const vh = window.innerHeight;
    if (left + tipW + pad > vw) {
        left = Math.max(pad, (typeof clientX === 'number' ? clientX : rect.left) - tipW - 8);
    }
    if (top + tipH + pad > vh) {
        top = Math.max(pad, (typeof clientY === 'number' ? clientY : rect.bottom) - tipH - 8);
    }

    tip.style.left = Math.round(left) + 'px';
    tip.style.top = Math.round(top) + 'px';
}

function bindPrevPeriodRecallFoldTooltipGlobal() {
    if (document.documentElement.dataset.prevRecallFoldTipBound === '1') {
        return;
    }
    document.documentElement.dataset.prevRecallFoldTipBound = '1';

    document.addEventListener('mouseover', function (event) {
        const hit = event.target && event.target.closest
            ? event.target.closest('.prev-period-recall-fold')
            : null;
        if (!hit) {
            return;
        }
        const pct = hit.getAttribute('data-pct') || '';
        if (!pct) {
            return;
        }
        if (prevRecallFoldTooltipHoverTarget === hit) {
            return;
        }
        prevRecallFoldTooltipHoverTarget = hit;
        clearTimeout(prevRecallFoldTooltipShowTimer);
        const mx = event.clientX;
        const my = event.clientY;
        prevRecallFoldTooltipShowTimer = setTimeout(function () {
            if (prevRecallFoldTooltipHoverTarget === hit) {
                showPrevRecallFoldTooltip(hit, pct, mx, my);
            }
        }, PREV_RECALL_FOLD_TOOLTIP_SHOW_MS);
    }, true);

    document.addEventListener('mouseout', function (event) {
        const hit = event.target && event.target.closest
            ? event.target.closest('.prev-period-recall-fold')
            : null;
        if (!hit) {
            return;
        }
        const to = event.relatedTarget;
        if (to && hit.contains(to)) {
            return;
        }
        hidePrevRecallFoldTooltip();
    }, true);
}

// Export for use in index.html
if (typeof module !== 'undefined' && module.exports) {
    module.exports = RightPaneSheetManager;
}
