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

/** sessionStorage + sheet.trackingUi: khôi phục timeline / predict khi quay lại sheet */
const TRACKING_UI_STORAGE_KEY = 'rp-tracking-ui-v1';
const LEGACY_TRACKING_UI_STORAGE_KEY = 'rp-special-tracking-ui-v1';
const TRACKING_LABEL_MODE_KEY = 'rp-tracking-label-mode-v1';
const TRACKING_SHEET_ID = 'tracking';
const TRACKING_KIND = 'tracking';

/** Chuột phải ô nonexist: nhảy tới kỳ có id = id hàng hiện tại + delta (vd 00014 → 00024 khi delta=10). */
const NONEXIST_CONTEXTMENU_ID_DELTA = 10;

/** Màu nền ô id (sheet1) theo số lần được tham chiếu trong note — tối đa 10 bậc. */
const ID_REF_COUNT_BG_COLORS = [
    'rgb(235, 255, 235)', // 1 — #EBFFEB
    'rgb(200, 255, 200)', // 2 — #C8FFC8
    'rgb(120, 230, 120)', // 3 — #78E678
    'rgb(0, 180, 0)',     // 4 — #00B400
    'rgb(0, 160, 0)',     // 5
    'rgb(0, 140, 0)',     // 6
    'rgb(0, 115, 0)',     // 7
    'rgb(0, 90, 0)',      // 8
    'rgb(0, 65, 0)',      // 9
    'rgb(0, 40, 0)'       // 10+
];

class RightPaneSheetManager {
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
        /** Id đối ứng Ctrl+Z: qua lại giữa id focus trước và hiện tại. */
        this.comboFocusUndoPeerId = '';
        this._comboFocusUndoBurstAnchorId = '';
        this._comboFocusUndoBurstEndId = '';
        this._comboFocusUndoCommitTimer = 0;
        this.answerPopupFocusMask = { active: false, rowIndex: -1 };
        this._answerPopupMaskAppliedRow = -1;
        this._answerPopupMaskApplyRaf = 0;
        this._idFreqAsOfCacheRow = -1;
        this._idFreqAsOfCache = null;
        this._filterAllModeMaskAppliedRow = -1;
        this._filterAllModeMaskApplyRaf = 0;
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
        this._syncingTrackingFromSheet1 = false;
        this._syncingSheet1FromTracking = false;
        /** Sheet1 arrow spam: bơm setLines sang iframe theo từng bước (giống tracking timeline). */
        this._sheet1LeftPaneTarget = -1;
        this._sheet1LeftPaneCurrent = -1;
        this._sheet1LeftPaneTimer = 0;
        this._sheet1LeftPaneStepMs = 58;
        this._sheet1NavTableWrap = null;
        /** Cache DOM bảng sheet1 khi đổi tab — tránh rebuild toàn bộ rows. */
        this._sheet1DomCache = null;
        /** Submit ON ở nửa màn trái (iframe ok_left) — ảnh hưởng basic tracking. */
        this.leftSubmitActive = false;
        /** Autoring ON (toolbar) — ảnh hưởng basic tracking. */
        this.leftAutoringEnabled = false;
        /** Số khoanh trái — preview freq trên basic tracking (kỳ cuối chưa có đáp án). */
        this.leftBasicPreviewPickNums = [];
        /** Nhớ pick giả lập khi bật Submit (khôi phục khi tắt Submit). */
        this.leftBasicPreviewPickNumsStash = [];
        /** Special tracking: giả lập đúng 1 bar (chỉ panel tracking, không đồng bộ nửa trái). */
        this.leftSpecialPreviewPickNum = null;
        this.leftSpecialPreviewPickNumStash = null;
        /** Special: chuỗi pick giả lập theo thứ tự (để chuột phải về số liền trước khi bỏ bar). */
        this.leftSpecialPreviewPickHistory = [];
        /** Bar giả lập focus cho chuột phải — theo thứ tự chuỗi pick, không phải bar vừa bỏ. */
        this.lastTrackingPreviewBarNum = null;
        /** Ctrl+Shift: stash viền cam quan sát (basic/special) để tắt hết rồi khôi phục. */
        this._trackingObserveFocusStashByMode = { basic: null, special: null };
        /** Tăng mỗi lần setLeftBasicPreviewPickNums đổi — chặn response requestLeftCircledNums cũ. */
        this._leftBasicPreviewPickGeneration = 0;
        /** Cache filter mode connection (invalid khi refreshDerivedState). */
        this._connectionFilterIndicesCache = null;
        this._connectionFilterIndicesCacheRowLen = 0;
        this._connectionFilterNoteCacheRef = null;
        /** Cache filter mode conn3 / 3-connection. */
        this._conn3FilterIndicesCache = null;
        this._conn3FilterIndicesCacheRowLen = 0;
        this._conn3WindowExistIndicesCache = null;
        this._conn3WindowExistIndicesCacheRowLen = 0;
        /** Cache tập mẫu theo kỳ cho lọc header2 filter popup. */
        this._filterRowMauSetsCache = null;
        this._filterRowMauSetsCacheKey = '';
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
                if (this.sheets.specialtracking && !this.sheets[TRACKING_SHEET_ID]) {
                    this.sheets[TRACKING_SHEET_ID] = {
                        ...this.sheets.specialtracking,
                        kind: TRACKING_KIND
                    };
                    delete this.sheets.specialtracking;
                }
                if (this.activeSheet === 'specialtracking') {
                    this.activeSheet = TRACKING_SHEET_ID;
                }
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
        this.invalidateSheet1TableDomCache();
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
        const specialMeta = this.buildSpecialTrackingSeriesMeta(this.sourceRows || []);
        const basicMeta = this.buildBasicTrackingSeriesMeta(this.sourceRows || []);
        const prevTracking = this.sheets[TRACKING_SHEET_ID] || this.sheets.specialtracking;
        const viewMode = prevTracking ? this.getTrackingViewMode(prevTracking) : 'basic';
        const trackingFrames = this.buildTrackingFramesForMode(viewMode, specialMeta, basicMeta);
        this.sheets = {
            sheet1: {
                kind: 'source',
                data: this.sourceRows || []
            },
            ...comboSheets,
            [TRACKING_SHEET_ID]: {
                kind: TRACKING_KIND,
                data: [],
                trackingViewMode: viewMode,
                trackingLabelMode: prevTracking && prevTracking.trackingLabelMode
                    ? RightPaneSheetManager.normalizeTrackingLabelMode(prevTracking.trackingLabelMode)
                    : RightPaneSheetManager.readTrackingLabelModeFromStorage(),
                specialSeries: specialMeta.series,
                specialDrawSteps: specialMeta.drawSteps,
                specialSourceRowIndices: specialMeta.sourceRowIndices,
                basicDraws: basicMeta.draws,
                basicDrawSteps: basicMeta.drawSteps,
                basicSourceRowIndices: basicMeta.sourceRowIndices,
                series: trackingFrames.series,
                seriesSourceRowIndices: trackingFrames.sourceRowIndices,
                frames: trackingFrames.frames
            }
        };
        try {
            sessionStorage.removeItem(TRACKING_UI_STORAGE_KEY);
            sessionStorage.removeItem(LEGACY_TRACKING_UI_STORAGE_KEY);
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
        this._idFreqAsOfCacheRow = -1;
        this._idFreqAsOfCache = null;
        this.nonexistGreenFilterCache = null;
        this.nonexistDisplayEntriesCache = null;
        this.datebandFilterIndicesCache = null;
        this.datebandFilterIndicesCacheRowLen = 0;
        this.datebandRowDistCache = null;
        this.datebandRowDistCacheRowLen = 0;
        this.tailFilterIndicesCache = null;
        this.tailFilterIndicesCacheRowLen = 0;
        this._connectionFilterIndicesCache = null;
        this._connectionFilterIndicesCacheRowLen = 0;
        this._connectionFilterNoteCacheRef = null;
        this._conn3FilterIndicesCache = null;
        this._conn3FilterIndicesCacheRowLen = 0;
        this._conn3WindowExistIndicesCache = null;
        this._conn3WindowExistIndicesCacheRowLen = 0;
        this._filterRowMauSetsCache = null;
        this._filterRowMauSetsCacheKey = '';
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

    getSourceSheetTableSig() {
        const rows = this.getSourceSheetRows();
        const tail = rows.length ? rows[rows.length - 1] : {};
        const tailId = String(tail.id ?? tail.ID ?? '');
        return `${rows.length}|${tailId}`;
    }

    getActiveWindowRangeCacheKey() {
        const r = this.activeWindowRange;
        if (!r || typeof r.start !== 'number' || typeof r.end !== 'number') {
            return 'none';
        }
        const target = typeof r.target === 'number' ? r.target : '';
        const idRefs = Array.isArray(r.idRefHighlightIndices) ? r.idRefHighlightIndices.join(',') : '';
        const noteRefs = Array.isArray(r.focusNoteRefHighlightIndices) ? r.focusNoteRefHighlightIndices.join(',') : '';
        return `${r.start}|${r.end}|${target}|${idRefs}|${noteRefs}`;
    }

    invalidateSheet1TableDomCache() {
        this._sheet1DomCache = null;
    }

    cacheSheet1TableDom(tableWrap) {
        if (!this.isSourceSheet1TableDom(tableWrap)) {
            return;
        }
        this._sheet1DomCache = {
            sig: this.getSourceSheetTableSig(),
            html: tableWrap.innerHTML,
            windowRangeKey: this.getActiveWindowRangeCacheKey()
        };
    }

    isSourceSheet1TableDom(tableWrap) {
        if (!tableWrap) {
            return false;
        }
        const table = tableWrap.querySelector('table.sheet1-source-table');
        if (!table || !table.classList.contains('sheet-data-table')) {
            return false;
        }
        if (table.classList.contains('combo-sheet-table')
            || table.classList.contains('combo-special-table')
            || table.classList.contains('combo-sheet-grid')) {
            return false;
        }
        return true;
    }

    tryRestoreSheet1TableDom(tableWrap, options = {}) {
        if (!tableWrap || !this._sheet1DomCache) {
            return false;
        }
        if (this._sheet1DomCache.sig !== this.getSourceSheetTableSig()) {
            this.invalidateSheet1TableDomCache();
            return false;
        }
        if (!String(this._sheet1DomCache.html || '').includes('sheet1-source-table')) {
            this.invalidateSheet1TableDomCache();
            return false;
        }
        tableWrap.innerHTML = this._sheet1DomCache.html;
        if (options.bindKeyboard !== false) {
            this.bindSourceSheetKeyboardNavigation(tableWrap);
        }
        this.bindSourceSheetTableAfterRender(tableWrap, {
            ...options,
            fromDomRestore: true,
            cachedWindowRangeKey: this._sheet1DomCache.windowRangeKey ?? null
        });
        this._lastSheet1RenderWasDomRestore = true;
        return true;
    }

    bindSourceSheetRowClickDelegation(tableWrap, options = {}) {
        if (!tableWrap) {
            return;
        }
        tableWrap.querySelectorAll('tbody tr').forEach(tr => {
            tr.style.cursor = 'pointer';
        });
        if (options.skipRowClickBind === true) {
            return;
        }
        tableWrap.__rowClickOnActivated = typeof options.onRowActivated === 'function'
            ? options.onRowActivated
            : null;
        if (tableWrap.dataset.rowClickDelegated === '1') {
            return;
        }
        tableWrap.dataset.rowClickDelegated = '1';
        tableWrap.addEventListener('click', (e) => {
            const tr = e.target.closest('tbody tr[data-idx]');
            if (!tr || !tableWrap.contains(tr)) {
                return;
            }
            this.onRowClick(Number(tr.dataset.idx), tr.dataset.empty === '1', e);
            try {
                tableWrap.focus({ preventScroll: true });
            } catch (err) {
                // ignore focus failures
            }
            const onActivated = tableWrap.__rowClickOnActivated;
            if (typeof onActivated === 'function') {
                onActivated(Number(tr.dataset.idx));
            }
        });
    }

    bindSourceSheetTableAfterRender(tableWrap, options = {}) {
        const applyWindowSelection = options.applyWindowSelection !== false;
        const fromDomRestore = options.fromDomRestore === true;
        this.bindSourceSheetRowClickDelegation(tableWrap, options);

        if (!tableWrap.dataset.nonexistContextmenuBound) {
            tableWrap.dataset.nonexistContextmenuBound = '1';
            tableWrap.addEventListener('contextmenu', (e) => {
                this.handleSourceSheetCellContextMenu(e, tableWrap);
            });
        }

        if (!fromDomRestore) {
            if (applyWindowSelection && this.activeWindowRange) {
                const selectionRoot = options.selectionRoot || tableWrap;
                const r = this.activeWindowRange;
                this.applyWindowSelection(
                    r.start,
                    r.end,
                    r.target,
                    selectionRoot,
                    {
                        idRefHighlightIndices: r.idRefHighlightIndices || null,
                        focusNoteRefHighlightIndices: r.focusNoteRefHighlightIndices || null
                    }
                );
            }

            const mainWrap = typeof document !== 'undefined' ? document.getElementById('tableWrap') : null;
            if (options.applyAnswerPopupMask !== false && tableWrap && tableWrap === mainWrap) {
                this.applyAnswerPopupFocusMaskToDom(tableWrap, { reset: true });
            }
            const filterWrap = typeof document !== 'undefined' ? document.getElementById('filterTableWrap') : null;
            if (options.applyAnswerPopupMask !== false && tableWrap && tableWrap === filterWrap) {
                this.applyFilterAllModeFocusMaskToDom(tableWrap, { reset: true });
            }
            if (this.activeWindowRange && Array.isArray(this.activeWindowRange.idRefHighlightIndices)
                && this.activeWindowRange.idRefHighlightIndices.length) {
                this.applyIdRefHighlightToDom(this.activeWindowRange.idRefHighlightIndices, tableWrap);
            }
            if (this.activeWindowRange && Array.isArray(this.activeWindowRange.focusNoteRefHighlightIndices)
                && this.activeWindowRange.focusNoteRefHighlightIndices.length) {
                this.applyFocusNoteRefHighlightToDom(this.activeWindowRange.focusNoteRefHighlightIndices, tableWrap);
            }
        } else if (tableWrap && tableWrap.id === 'tableWrap') {
            const m = this.answerPopupFocusMask || {};
            this._answerPopupMaskAppliedRow = m.active ? m.rowIndex : -1;
            const cachedKey = options.cachedWindowRangeKey ?? null;
            const currentKey = this.getActiveWindowRangeCacheKey();
            const focusChangedWhileAway = cachedKey !== currentKey;
            if (focusChangedWhileAway && applyWindowSelection && this.activeWindowRange) {
                const r = this.activeWindowRange;
                this.applyWindowSelection(
                    r.start,
                    r.end,
                    r.target,
                    tableWrap,
                    {
                        idRefHighlightIndices: r.idRefHighlightIndices || null,
                        focusNoteRefHighlightIndices: r.focusNoteRefHighlightIndices || null
                    }
                );
                if (options.applyAnswerPopupMask !== false) {
                    this.applyAnswerPopupFocusMaskToDom(tableWrap, { reset: true });
                }
                requestAnimationFrame(() => {
                    if (this.activeSheet === 'sheet1') {
                        this.centerActiveWindowInView(tableWrap);
                    }
                });
            }
        }

        bindPrevPeriodRecallFoldTooltipGlobal();
    }

    /**
     * Render data table with frequency-based styling
     */
    renderTable(tableWrap) {
        if (tableWrap && this.isSourceSheet1TableDom(tableWrap)) {
            this.cacheSheet1TableDom(tableWrap);
        }
        if (tableWrap) {
            tableWrap.classList.remove('table-wrap--tracking');
        }
        if (tableWrap && typeof tableWrap.__trackingCleanup === 'function') {
            try {
                tableWrap.__trackingCleanup();
            } catch (eSt) {
                /* ignore */
            }
            tableWrap.__trackingCleanup = null;
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

        if (sheet.kind === TRACKING_KIND) {
            this.ensureTrackingFrames(sheet);
            tableWrap.classList.add('table-wrap--tracking');
            tableWrap.innerHTML = this.renderTrackingShell(sheet);
            this.wireTrackingUi(tableWrap, sheet);
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
            this._lastSheet1RenderWasDomRestore = false;
            if (!this.tryRestoreSheet1TableDom(tableWrap, { bindKeyboard: true })) {
                this.renderSourceSheet(tableWrap, sheet.data);
            }
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
     * Khóa vị trí tương đối: trừ min → cùng “hình” lệch nhau cùng khóa
     * (vd [1,6,9] và [2,7,10] → "0,5,8"; [5,7] và [1,3] → "0,2").
     */
    posnfreqRelativeKey(positions) {
        const arr = this.posnfreqNormalizePositionsList(positions);
        if (!arr.length) {
            return '';
        }
        const min = arr[0];
        let out = '0';
        for (let i = 1; i < arr.length; i++) {
            out += `,${arr[i] - min}`;
        }
        return out;
    }

    /** Copy + số hóa + sort tăng dần nhãn Chuỗi. */
    posnfreqNormalizePositionsList(positions) {
        const src = Array.isArray(positions) ? positions : [];
        const out = [];
        for (let i = 0; i < src.length; i++) {
            const n = Number(src[i]);
            if (Number.isFinite(n)) {
                out.push(n);
            }
        }
        out.sort((a, b) => a - b);
        return out;
    }

    /**
     * Multiset bao hàm tuyệt đối: ref ⊆ cand (vd [1,6,9] ⊆ [1,2,6,9]).
     */
    posnfreqPositionsCover(candPositions, refPositions) {
        const cand = this.posnfreqNormalizePositionsList(candPositions);
        const ref = this.posnfreqNormalizePositionsList(refPositions);
        if (ref.length === 0) {
            return true;
        }
        if (cand.length < ref.length) {
            return false;
        }
        const need = new Map();
        for (let i = 0; i < ref.length; i++) {
            const p = ref[i];
            need.set(p, (need.get(p) || 0) + 1);
        }
        for (let i = 0; i < cand.length; i++) {
            const p = cand[i];
            const left = need.get(p);
            if (left == null) {
                continue;
            }
            if (left <= 1) {
                need.delete(p);
            } else {
                need.set(p, left - 1);
            }
        }
        return need.size === 0;
    }

    /**
     * Bao hàm tương đối (khi tắt f): tồn tại tập con trong cand cùng khóa tương đối với ref.
     * vd ref [5,7] ≡ hình "0,2" → [1,3], [2,4], … và cha [1,3,10] (có [1,3]) đều thỏa.
     */
    posnfreqRelativeCover(candPositions, refPositions) {
        const cand = this.posnfreqNormalizePositionsList(candPositions);
        const ref = this.posnfreqNormalizePositionsList(refPositions);
        if (ref.length === 0) {
            return true;
        }
        if (cand.length < ref.length) {
            return false;
        }
        const refKey = this.posnfreqRelativeKey(ref);
        if (cand.length === ref.length) {
            return this.posnfreqRelativeKey(cand) === refKey;
        }
        // Nhanh: mẫu 2 điểm — chỉ cần một cặp cùng hiệu số (vd [5,7] → diff 2).
        if (ref.length === 2) {
            const diff = ref[1] - ref[0];
            for (let i = 0; i < cand.length; i++) {
                for (let j = i + 1; j < cand.length; j++) {
                    if (cand[j] - cand[i] === diff) {
                        return true;
                    }
                }
            }
            return false;
        }
        const need = ref.length;
        const path = [];
        const dfs = (start) => {
            if (path.length === need) {
                return this.posnfreqRelativeKey(path) === refKey;
            }
            const remain = need - path.length;
            for (let i = start; i <= cand.length - remain; i++) {
                path.push(cand[i]);
                if (dfs(i + 1)) {
                    return true;
                }
                path.pop();
            }
            return false;
        };
        return dfs(0);
    }

    /**
     * @param {{ matchFrequency?: boolean, matchPositionsMode?: string, matchPositions?: boolean|string }|null|undefined} matchOpts
     *   matchPositionsMode: 'off' | 'absolute' | 'relative' (mặc định absolute).
     */
    normalizePosnfreqMatchOpts(matchOpts) {
        const o = matchOpts && typeof matchOpts === 'object' ? matchOpts : {};
        let mode = o.matchPositionsMode;
        if (mode !== 'off' && mode !== 'absolute' && mode !== 'relative') {
            if (o.matchPositions === false || o.matchPositions === 'off') {
                mode = 'off';
            } else if (o.matchPositions === 'relative') {
                mode = 'relative';
            } else {
                mode = 'absolute';
            }
        }
        const hasFreq = Object.prototype.hasOwnProperty.call(o, 'matchFrequency');
        return {
            matchFrequency: hasFreq ? !!o.matchFrequency : true,
            matchPositionsMode: mode
        };
    }

    /**
     * So khớp chữ ký posnfreq theo từng ràng buộc (f / positions) có thể tắt độc lập.
     * - matchFrequency: đúng bằng f của mẫu.
     * - absolute: vị trí tuyệt đối (có f → khớp đúng []; tắt f → ⊆).
     * - relative: cùng hình lệch (vd [5,7]≡[1,3]≡[2,4]); tắt f → ⊆ tương đối (vd [1,3,10] ok).
     */
    posnfreqSignatureMatches(sig, refSig, matchOpts) {
        if (!sig || sig.frequency === 0) {
            return false;
        }
        if (!refSig || !Number.isFinite(refSig.frequency)) {
            return false;
        }
        const opts = this.normalizePosnfreqMatchOpts(matchOpts);
        if (opts.matchFrequency && sig.frequency !== refSig.frequency) {
            return false;
        }
        const mode = opts.matchPositionsMode;
        if (mode === 'off') {
            return true;
        }
        if (mode === 'relative') {
            // Có f: cùng độ dài → cover ≡ khớp đúng hình lệch.
            // Tắt f: cho phép cha dài hơn miễn còn một tập con cùng hình.
            return this.posnfreqRelativeCover(sig.positions, refSig.positions);
        }
        if (opts.matchFrequency) {
            return this.posnfreqPositionsKey(sig) === this.posnfreqPositionsKey(refSig);
        }
        return this.posnfreqPositionsCover(sig.positions, refSig.positions);
    }

    /**
     * @param {object|null} refSig — chữ ký đầy đủ cửa sổ 10 chuỗi của specimen trên kỳ mẫu (f + multiset nhãn Chuỗi)
     * @param {boolean} specimenStrict — true (Số): chỉ số specimen có cùng refSig;
     *                                   false (Mẫu): tồn tại m ∈ [1..35] có cùng refSig (lục giác có thể “đặt” lên m)
     * @param {{ matchFrequency?: boolean, matchPositions?: boolean }|null|undefined} [matchOpts]
     */
    rowMatchesPosnfreqFilter(rows, rowIndex, specimenNum, refSig, specimenStrict, matchOpts) {
        const row = rows[rowIndex];
        if (!row || this.isEmptyResultRow(row)) {
            return false;
        }
        if (!refSig || !Number.isFinite(refSig.frequency)) {
            return false;
        }
        if (specimenStrict) {
            const sig = this.computePosnfreqSignature(rows, rowIndex, specimenNum);
            return this.posnfreqSignatureMatches(sig, refSig, matchOpts);
        }
        for (let m = 1; m <= 35; m++) {
            const sig = this.computePosnfreqSignature(rows, rowIndex, m);
            if (this.posnfreqSignatureMatches(sig, refSig, matchOpts)) {
                return true;
            }
        }
        return false;
    }

    /**
     * Mọi m ∈ [1..35] có chữ ký posnfreq trùng refSig trên rowIndex (đã sort tăng dần).
     * @param {{ matchFrequency?: boolean, matchPositions?: boolean }|null|undefined} [matchOpts]
     */
    findAllPosnfreqMatchingNumbers(rows, rowIndex, refSig, matchOpts) {
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
            if (this.posnfreqSignatureMatches(sig, refSig, matchOpts)) {
                out.push(m);
            }
        }
        return out;
    }

    /**
     * Số nhỏ nhất m sao cho chữ ký posnfreq của m trên rowIndex khớp refSig (dùng cho viền lục giác Mẫu).
     */
    findPosnfreqMatchingNumber(rows, rowIndex, refSig, matchOpts) {
        const all = this.findAllPosnfreqMatchingNumbers(rows, rowIndex, refSig, matchOpts);
        return all.length ? all[0] : null;
    }

    /**
     * Tập số "mẫu" của một kỳ cho lọc header2 popup: posnfreq refSig nếu có, không thì 5 số chính result.
     * @param {number} rowIndex
     * @param {object|null} [refSignature]
     * @param {{ matchFrequency?: boolean, matchPositions?: boolean }|null|undefined} [matchOpts]
     * @returns {number[]}
     */
    getFilterRowMauNumbers(rowIndex, refSignature = null, matchOpts = null) {
        const rows = this.getSourceSheetRows();
        if (!Array.isArray(rows) || rowIndex < 0 || rowIndex >= rows.length) {
            return [];
        }
        if (refSignature && typeof this.findAllPosnfreqMatchingNumbers === 'function') {
            return this.findAllPosnfreqMatchingNumbers(rows, rowIndex, refSignature, matchOpts);
        }
        const row = rows[rowIndex];
        if (!row) {
            return [];
        }
        return this.parseMainNums(row.result || row.Result || '');
    }

    /**
     * Tập mẫu theo kỳ cho lọc header2 — tính trước (posnfreq: tối đa 35 chữ ký / kỳ).
     * @param {number[]} indices
     * @param {object|null} [refSignature]
     * @param {{ matchFrequency?: boolean, matchPositions?: boolean }|null|undefined} [matchOpts]
     * @returns {{ mauByRow: Map<number, Set<number>>, rowsByNum: Set<number>[] }}
     */
    ensureFilterRowMauSetsCache(indices, refSignature = null, matchOpts = null) {
        const rows = this.getSourceSheetRows();
        const list = Array.isArray(indices) ? indices : [];
        const opts = this.normalizePosnfreqMatchOpts(matchOpts);
        const sigPart = refSignature && Number.isFinite(refSignature.frequency)
            ? `pnf:${refSignature.frequency}:${this.posnfreqPositionsKey(refSignature)}:${opts.matchFrequency ? '1' : '0'}:${opts.matchPositionsMode}`
            : 'main';
        const cacheKey = `${rows.length}|${sigPart}|${list.length}:${list[0] ?? ''}:${list[list.length - 1] ?? ''}`;
        if (this._filterRowMauSetsCache && this._filterRowMauSetsCacheKey === cacheKey) {
            return this._filterRowMauSetsCache;
        }
        const mauByRow = new Map();
        const rowsByNum = Array.from({ length: 36 }, () => new Set());
        for (let i = 0; i < list.length; i++) {
            const rowIndex = list[i];
            const nums = this.getFilterRowMauNumbers(rowIndex, refSignature, matchOpts);
            const set = new Set(nums);
            mauByRow.set(rowIndex, set);
            for (let u = 0; u < nums.length; u++) {
                const n = nums[u];
                if (n >= 1 && n <= 35) {
                    rowsByNum[n].add(rowIndex);
                }
            }
        }
        this._filterRowMauSetsCache = { mauByRow, rowsByNum };
        this._filterRowMauSetsCacheKey = cacheKey;
        return this._filterRowMauSetsCache;
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

    /**
     * Dateband #00b0f0: cặp cửa sổ 10 khớp đáp án — hai freq khớp ngưỡng x,y (không phân thứ tự).
     * @param {object[]} rows
     * @param {number} rowIndex
     * @param {number} [threshX=2]
     * @param {number} [threshY=2]
     * @param {string} [opX='>=']
     * @param {string} [opY='>=']
     * @returns {boolean}
     */
    rowMatchesDatebandPairFreqFilter(rows, rowIndex, threshX = 2, threshY = 2, opX = '>=', opY = '>=') {
        const list = rows || this.getSourceSheetRows();
        if (!this.rowHasCyanDateBand(list, rowIndex)) {
            return false;
        }
        const currentRow = list[rowIndex] || {};
        const currentNums = this.parseMainNums(currentRow.result || currentRow.Result || '');
        if (currentNums.length !== 5 || rowIndex < 10) {
            return false;
        }
        const windowRows = list.slice(Math.max(0, rowIndex - 10), rowIndex);
        if (windowRows.length < 10) {
            return false;
        }
        const visiblePairs = this.computePairsForRows(windowRows);
        if (!visiblePairs || !visiblePairs.length) {
            return false;
        }
        const freq = new Array(36).fill(0);
        for (let wi = 0; wi < windowRows.length; wi++) {
            const nums = this.parseMainNums(windowRows[wi].result || windowRows[wi].Result || '');
            for (let k = 0; k < nums.length; k++) {
                const n = nums[k];
                if (n >= 1 && n <= 35) {
                    freq[n]++;
                }
            }
        }
        for (let p = 0; p < visiblePairs.length; p++) {
            const a = visiblePairs[p][0];
            const b = visiblePairs[p][1];
            if (!this.pairExists(currentNums, a, b)) {
                continue;
            }
            if (this.pairFreqMatchesUnorderedThresholds(freq[a], freq[b], threshX, threshY, opX, opY)) {
                return true;
            }
        }
        return false;
    }

    /**
     * Hai freq của cặp khớp ngưỡng x,y — gán x/y cho số nào cũng được.
     */
    pairFreqMatchesUnorderedThresholds(freqA, freqB, threshX, threshY, opX, opY) {
        return (this.freqMatchesComparison(freqA, threshX, opX) && this.freqMatchesComparison(freqB, threshY, opY))
            || (this.freqMatchesComparison(freqA, threshY, opY) && this.freqMatchesComparison(freqB, threshX, opX));
    }

    freqMatchesComparison(freq, threshold, op) {
        if (op === '=') {
            return freq === threshold;
        }
        if (op === '<=') {
            return freq <= threshold;
        }
        return freq >= threshold;
    }

    getFilterMatchingIndices(mode, filterOptions = null) {
        const indices = [];
        const rows = this.getSourceSheetRows();
        if (mode === 'all') {
            const noteTags = Array.isArray((filterOptions || {}).noteTags)
                ? (filterOptions.noteTags || []).filter((n) => Number.isFinite(n) && n >= 1 && n <= 10)
                : [];
            const noteTRefs = this.parseNoteTRefExps(filterOptions);
            for (let i = 0; i < rows.length; i++) {
                if (noteTags.length > 0 && !this.rowMatchesNoteTagFilter(i, noteTags)) {
                    continue;
                }
                if (noteTRefs.length > 0 && !this.rowMatchesNoteTRefFilter(i, noteTRefs)) {
                    continue;
                }
                indices.push(i);
            }
            return indices;
        }
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
            const applyColorFilter = opts.applyColorStyleFilter === true;
            const hasSpecificNum = num !== null && Number.isFinite(num) && num >= 1 && num <= 35;
            let bracketTh = parseInt(opts.nonexistBracketCount, 10);
            if (!Number.isFinite(bracketTh)) {
                bracketTh = 0;
            }
            bracketTh = Math.min(5, Math.max(0, bracketTh));
            const rawBracketOp = String(opts.nonexistBracketOp || '').trim();
            const bracketOp = rawBracketOp === '=' || rawBracketOp === '<=' || rawBracketOp === '>=' ? rawBracketOp : '>=';

            if (colors.length === 0) {
                return indices;
            }

            for (let i = 0; i < rows.length; i++) {
                if (this.isEmptyResultRow(rows[i])) {
                    continue;
                }
                if (!this.rowMatchesNonexistBracketFilter(i, bracketTh, bracketOp, colors, styles)) {
                    continue;
                }
                if (hasSpecificNum) {
                    if (!this.rowMatchesNonexistSpecificNumFilter(i, num, colors, styles)) {
                        continue;
                    }
                } else if (applyColorFilter) {
                    if (colors.length === 0 || !this.rowMatchesNonexistColorFilter(i, colors, styles, null)) {
                        continue;
                    }
                }
                indices.push(i);
            }
            return indices;
        }

        const noteTags = Array.isArray((filterOptions || {}).noteTags)
            ? (filterOptions.noteTags || []).filter((n) => Number.isFinite(n) && n >= 1 && n <= 10)
            : [];
        const noteTRefs = this.parseNoteTRefExps(filterOptions);

        if (mode === 'dateband') {
            const o = filterOptions || {};
            let thX = parseInt(o.datebandMinFreqA, 10);
            let thY = parseInt(o.datebandMinFreqB, 10);
            if (!Number.isFinite(thX)) {
                thX = 2;
            }
            if (!Number.isFinite(thY)) {
                thY = 2;
            }
            thX = Math.min(7, Math.max(2, thX));
            thY = Math.min(7, Math.max(2, thY));
            const rawOpA = String(o.datebandFreqOpA || '').trim();
            const rawOpB = String(o.datebandFreqOpB || '').trim();
            const opA = rawOpA === '=' || rawOpA === '<=' || rawOpA === '>=' ? rawOpA : '>=';
            const opB = rawOpB === '=' || rawOpB === '<=' || rawOpB === '>=' ? rawOpB : '>=';
            const base = this.ensureDatebandFilterIndicesCache();
            const distFilter = noteTags.length > 0 ? noteTags[0] : null;
            for (let b = 0; b < base.length; b++) {
                const i = base[b];
                if (distFilter !== null && !this.rowMatchesDatebandNoteDistFilter(i, distFilter)) {
                    continue;
                }
                if (noteTRefs.length > 0 && !this.rowMatchesNoteTRefFilter(i, noteTRefs)) {
                    continue;
                }
                if (!this.rowMatchesDatebandPairFreqFilter(rows, i, thX, thY, opA, opB)) {
                    continue;
                }
                indices.push(i);
            }
            return indices;
        }

        if (mode === 'tail') {
            const o = filterOptions || {};
            let th = parseInt(o.tailMinCount, 10);
            if (!Number.isFinite(th)) {
                th = 2;
            }
            th = Math.min(5, Math.max(2, th));
            const rawOp = String(o.tailCountOp || '').trim();
            const op = rawOp === '=' || rawOp === '<=' || rawOp === '>=' ? rawOp : '>=';
            for (let i = 0; i < rows.length; i++) {
                if (!this.isEmptyResultRow(rows[i])
                    && this.shouldHighlightDateByTailWindow(rows, i, { tailMinCount: th, tailCountOp: op })) {
                    indices.push(i);
                }
            }
            return indices;
        }

        if (mode === 'conn3') {
            return this.ensureConn3FilterIndicesCache().slice();
        }

        if (mode === 'connection') {
            const base = this.ensureConnectionFilterIndicesCache();
            if (noteTags.length === 0 && noteTRefs.length === 0) {
                return base.slice();
            }
            const out = [];
            for (let b = 0; b < base.length; b++) {
                const i = base[b];
                if (noteTags.length > 0 && !this.rowMatchesNoteTagFilter(i, noteTags)) {
                    continue;
                }
                if (noteTRefs.length > 0 && !this.rowMatchesNoteTRefFilter(i, noteTRefs)) {
                    continue;
                }
                out.push(i);
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
                if (!this.rowMatchesIntersectionSubmitWindow(rows, i, kind, thX, thY, opA, opB)) {
                    continue;
                }
                if (noteTags.length > 0 && !this.rowMatchesNoteTagFilter(i, noteTags)) {
                    continue;
                }
                if (noteTRefs.length > 0 && !this.rowMatchesNoteTRefFilter(i, noteTRefs)) {
                    continue;
                }
                indices.push(i);
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
            const matchOpts = this.normalizePosnfreqMatchOpts(o);
            let refSig = o.refSignature;
            if (!refSig && refRow >= 0) {
                refSig = this.computePosnfreqSignature(rows, refRow, specimen);
            }
            for (let i = 0; i < rows.length; i++) {
                if (this.rowMatchesPosnfreqFilter(rows, i, specimen, refSig, specimenStrict, matchOpts)) {
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
                if (noteTRefs.length > 0 && !this.rowMatchesNoteTRefFilter(i, noteTRefs)) {
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
     * Note có "connection": một số xuất hiện trong ≥2 cặp `{a,b}` — mọi cặp trong mỗi khối `{…}` của note.
     * @param {string} noteText
     * @returns {boolean}
     */
    noteTextHasConnectionPairing(noteText) {
        const text = String(noteText || '');
        if (text.indexOf('{') === -1 || text.indexOf('}') === -1) {
            return false;
        }
        const re = /\{([^}]+)\}/g;
        /** @type {Map<number, Set<number>>} */
        const numToPairIdx = new Map();
        let pairIndex = 0;
        let m;
        while ((m = re.exec(text)) !== null) {
            const nums = m[1].split(',')
                .map((part) => parseInt(part.trim(), 10))
                .filter((n) => Number.isFinite(n));
            for (let i = 0; i < nums.length; i++) {
                for (let j = i + 1; j < nums.length; j++) {
                    for (const n of [nums[i], nums[j]]) {
                        let set = numToPairIdx.get(n);
                        if (!set) {
                            set = new Set();
                            numToPairIdx.set(n, set);
                        }
                        set.add(pairIndex);
                    }
                    pairIndex++;
                }
            }
        }
        for (const s of numToPairIdx.values()) {
            if (s.size >= 2) {
                return true;
            }
        }
        return false;
    }

    /**
     * Freq 1–35 trong cửa sổ trượt 10 chuỗi trước kỳ `rowIndex` (không gồm kỳ đó).
     * Cùng logic với buildPickChainLinesBeforeRow / inferMode.
     * @returns {Record<number, number>}
     */
    computeMainNumsWindow10Freq(rows, rowIndex) {
        const freq = {};
        for (let i = 1; i <= 35; i++) {
            freq[i] = 0;
        }
        if (!Array.isArray(rows) || typeof rowIndex !== 'number' || rowIndex < 0) {
            return freq;
        }
        const windowStart = Math.max(0, rowIndex - 10);
        const limit = Math.min(rowIndex - windowStart, 10);
        for (let offset = 0; offset < limit; offset++) {
            const lineRow = rows[windowStart + offset] || {};
            const nums = this.parseMainNums(lineRow.result || lineRow.Result || '');
            for (let k = 0; k < nums.length; k++) {
                const n = nums[k];
                if (n >= 1 && n <= 35) {
                    freq[n] = (freq[n] || 0) + 1;
                }
            }
        }
        return freq;
    }

    /**
     * Chuỗi 11/12 trước cửa sổ 10 — cùng logic ok_left.getCh11Ch12NumsSet cho rowIndex sheet1/tracking.
     * @returns {Set<number>}
     */
    getCh11Ch12NumsSetForSourceRow(rows, rowIndex) {
        const set = new Set();
        if (!Array.isArray(rows) || typeof rowIndex !== 'number' || rowIndex < 0) {
            return set;
        }
        const windowTop = rowIndex >= 10 ? rowIndex - 10 : 0;
        const contextPrefixCount = rowIndex >= 10 ? Math.min(2, windowTop) : 0;
        const addFromRow = (row) => {
            const nums = this.parseMainNums(row?.result || row?.Result || '');
            for (let k = 0; k < nums.length; k++) {
                const n = nums[k];
                if (n >= 1 && n <= 35) {
                    set.add(n);
                }
            }
        };
        if (contextPrefixCount >= 2) {
            const i11 = windowTop - 2;
            const i12 = windowTop - 1;
            if (i11 >= 0 && i11 < rows.length) {
                addFromRow(rows[i11]);
            }
            if (i12 >= 0 && i12 < rows.length) {
                addFromRow(rows[i12]);
            }
        } else if (contextPrefixCount === 1) {
            const i = windowTop - 1;
            if (i >= 0 && i < rows.length) {
                addFromRow(rows[i]);
            }
        } else if (rows.length >= 13) {
            addFromRow(rows[10]);
            addFromRow(rows[11]);
        } else if (rows.length === 12) {
            addFromRow(rows[10]);
        }
        return set;
    }

    /**
     * Chuỗi 11 = hàng liền trên chuỗi 10 (windowTop − 1) — in nghiêng bar basic tracking.
     * @returns {Set<number>}
     */
    getCh11NumsSetForSourceRow(rows, rowIndex) {
        const set = new Set();
        if (!Array.isArray(rows) || typeof rowIndex !== 'number' || rowIndex < 10) {
            return set;
        }
        const windowTop = rowIndex - 10;
        const i11 = windowTop - 1;
        if (i11 < 0 || i11 >= rows.length) {
            return set;
        }
        const nums = this.parseMainNums(rows[i11]?.result || rows[i11]?.Result || '');
        for (let k = 0; k < nums.length; k++) {
            const n = nums[k];
            if (n >= 1 && n <= 35) {
                set.add(n);
            }
        }
        return set;
    }

    /**
     * Số xuất hiện ở chuỗi 1 hoặc chuỗi 2 (hai kỳ liền trước kỳ nguồn) — nghiêng trái trên bar basic tracking.
     * @returns {Set<number>}
     */
    getCh1Ch2NumsSetForSourceRow(rows, rowIndex) {
        const set = new Set();
        const lines = this.buildPickChainLinesBeforeRow(rows, rowIndex);
        for (let i = 0; i < lines.length; i++) {
            const line = lines[i];
            if (line.label !== 1 && line.label !== 2) {
                continue;
            }
            const nums = Array.isArray(line.nums) ? line.nums : [];
            for (let k = 0; k < nums.length; k++) {
                const n = nums[k];
                if (n >= 1 && n <= 35) {
                    set.add(n);
                }
            }
        }
        return set;
    }

    /**
     * Special tracking: số đặc biệt ở kỳ 1 hoặc 2 liền trước kỳ nguồn — nghiêng trái (giống chuỗi 1/2 basic).
     * @param {Array<number|null>} drawSteps `specialDrawSteps[ri]` = số 1–12 hoặc null
     * @param {number} rowIndex
     * @returns {Set<number>}
     */
    getSpecialCh1Ch2NumsSetForSourceRow(drawSteps, rowIndex) {
        const set = new Set();
        if (!Array.isArray(drawSteps) || typeof rowIndex !== 'number' || rowIndex < 0) {
            return set;
        }
        for (let offset = 1; offset <= 2; offset++) {
            const ri = rowIndex - offset;
            if (ri < 0) {
                continue;
            }
            const step = drawSteps[ri];
            if (Number.isFinite(step) && step >= 1 && step <= 12) {
                set.add(step | 0);
            }
        }
        return set;
    }

    /**
     * Special tracking: số đặc biệt ở kỳ 11 hoặc 12 liền trước kỳ nguồn — nghiêng phải (giống chuỗi 11/12 basic).
     * @param {Array<number|null>} drawSteps `specialDrawSteps[ri]` = số 1–12 hoặc null
     * @param {number} rowIndex
     * @returns {Set<number>}
     */
    getSpecialCh11Ch12NumsSetForSourceRow(drawSteps, rowIndex) {
        const set = new Set();
        if (!Array.isArray(drawSteps) || typeof rowIndex !== 'number' || rowIndex < 0) {
            return set;
        }
        for (let offset = 11; offset <= 12; offset++) {
            const ri = rowIndex - offset;
            if (ri < 0) {
                continue;
            }
            const step = drawSteps[ri];
            if (Number.isFinite(step) && step >= 1 && step <= 12) {
                set.add(step | 0);
            }
        }
        return set;
    }

    /**
     * 10 chuỗi trước kỳ `rowIndex` (Chuỗi 1 = sát kỳ đang xét).
     * @param {object[]} rows
     * @param {number} rowIndex
     * @returns {{ label: number, nums: number[] }[]}
     */
    buildPickChainLinesBeforeRow(rows, rowIndex) {
        const windowStart = Math.max(0, rowIndex - 10);
        const limit = Math.min(rowIndex - windowStart, 10);
        /** @type {{ label: number, nums: number[] }[]} */
        const lines = [];
        for (let offset = 0; offset < limit; offset++) {
            const lineRow = rows[windowStart + offset] || {};
            const nums = this.parseMainNums(lineRow.result || lineRow.Result || '');
            const label = limit - offset;
            lines.push({ label, nums });
        }
        return lines;
    }

    /**
     * Cột follow: số duy nhất freq cao nhất trong 5 số chuỗi 1 (cửa sổ 10 trước kỳ);
     * nếu ≥2 số cùng freq max → ký hiệu `?`. Tính được cả khi dòng hiện tại chưa có result.
     * @param {object[]} rows
     * @param {number} rowIndex
     * @returns {string}
     */
    computeFollowCellValue(rows, rowIndex) {
        if (!Array.isArray(rows) || typeof rowIndex !== 'number' || rowIndex < 0) {
            return '';
        }
        const row = rows[rowIndex];
        if (!row) {
            return '';
        }
        const chainLines = this.buildPickChainLinesBeforeRow(rows, rowIndex);
        let chain1 = null;
        for (let i = 0; i < chainLines.length; i++) {
            if (chainLines[i].label === 1) {
                chain1 = chainLines[i];
                break;
            }
        }
        if (!chain1 || !Array.isArray(chain1.nums) || !chain1.nums.length) {
            return '';
        }
        const chain1Nums = chain1.nums.filter((n) => n >= 1 && n <= 35);
        if (!chain1Nums.length) {
            return '';
        }
        const freq = this.computeMainNumsWindow10Freq(rows, rowIndex);
        let maxFreq = -1;
        for (let j = 0; j < chain1Nums.length; j++) {
            const f = freq[chain1Nums[j]] || 0;
            if (f > maxFreq) {
                maxFreq = f;
            }
        }
        const tied = chain1Nums.filter((n) => (freq[n] || 0) === maxFreq);
        if (tied.length === 1) {
            return String(tied[0]);
        }
        if (tied.length >= 2) {
            return '?';
        }
        return '';
    }

    /**
     * Follow xác định: cột follow có số duy nhất (khác `?` và rỗng).
     * @param {object[]} rows
     * @param {number} rowIndex
     * @returns {boolean}
     */
    rowHasDeterminedFollow(rows, rowIndex) {
        const value = this.computeFollowCellValue(rows, rowIndex);
        if (!value || value === '?') {
            return false;
        }
        const num = parseInt(value, 10);
        return Number.isFinite(num) && num >= 1 && num <= 35;
    }

    rowHasUndeterminedFollow(rows, rowIndex) {
        return this.computeFollowCellValue(rows, rowIndex) === '?';
    }

    /**
     * Cùng định nghĩa với recallsAtLeastOneFromPrevPeriodAtOffset / tooltip fold theo chuỗi.
     * @param {number} rowIndex
     * @param {number[]} selectedLabels — 1..10
     * @param {'or'|'and'} [combineMode='or'] — or: ∪ (U thuận); and: ∩ (U flip)
     */
    rowMatchesFilterChainLabels(rowIndex, selectedLabels, combineMode = 'or') {
        const labels = Array.isArray(selectedLabels)
            ? selectedLabels.filter((n) => Number.isFinite(n) && n >= 1 && n <= 10)
            : [];
        if (!labels.length) {
            return true;
        }
        return this.rowMatchesFilterChainLabelsFromRows(
            this.getSourceSheetRows(),
            rowIndex,
            labels,
            combineMode
        );
    }

    /**
     * @param {object[]} rows
     * @param {number} rowIndex
     * @param {number[]} labels — 1..10
     * @param {'or'|'and'} [combineMode='or']
     */
    rowMatchesFilterChainLabelsFromRows(rows, rowIndex, labels, combineMode = 'or') {
        if (!labels.length) {
            return true;
        }
        const list = rows || [];
        const idx = Number(rowIndex);
        if (!Number.isFinite(idx) || idx < 0 || idx >= list.length) {
            return false;
        }
        const curRow = list[idx];
        if (!curRow || this.isEmptyResultRow(curRow)) {
            return false;
        }
        const cur = this.parseMainNums(curRow.result || curRow.Result || '');
        if (cur.length !== 5) {
            return false;
        }
        const curSet = new Set(cur);
        const matchesOffset = (off) => {
            const prevIdx = idx - off;
            if (prevIdx < 0 || prevIdx >= list.length) {
                return false;
            }
            const prevRow = list[prevIdx];
            if (!prevRow || this.isEmptyResultRow(prevRow)) {
                return false;
            }
            const prev = this.parseMainNums(prevRow.result || prevRow.Result || '');
            if (prev.length !== 5) {
                return false;
            }
            for (let pi = 0; pi < prev.length; pi++) {
                if (curSet.has(prev[pi])) {
                    return true;
                }
            }
            return false;
        };
        if (combineMode === 'and') {
            for (let li = 0; li < labels.length; li++) {
                if (!matchesOffset(labels[li])) {
                    return false;
                }
            }
            return true;
        }
        for (let li = 0; li < labels.length; li++) {
            if (matchesOffset(labels[li])) {
                return true;
            }
        }
        return false;
    }

    /**
     * Lọc indices theo nhãn chuỗi — một lần load rows, tránh gọi lặp getSourceSheetRows().
     * @param {number[]} indices
     * @param {number[]} selectedLabels
     * @param {'or'|'and'} [combineMode='or']
     * @returns {number[]}
     */
    filterIndicesByChainLabels(indices, selectedLabels, combineMode = 'or') {
        const base = Array.isArray(indices) ? indices : [];
        const labels = Array.isArray(selectedLabels)
            ? selectedLabels.filter((n) => Number.isFinite(n) && n >= 1 && n <= 10)
            : [];
        if (!labels.length) {
            return base.slice();
        }
        const rows = this.getSourceSheetRows();
        const out = [];
        for (let i = 0; i < base.length; i++) {
            if (this.rowMatchesFilterChainLabelsFromRows(rows, base[i], labels, combineMode)) {
                out.push(base[i]);
            }
        }
        return out;
    }

    /**
     * Cặp [a,b] theo từng chuỗi: mọi cặp hai số pick cùng nằm trên một dòng (thứ tự trên dòng).
     * @param {object[]} rows
     * @param {number} rowIndex
     * @param {number[]} pickNums
     * @returns {number[][]}
     */
    buildChainPairsFromPickSet(rows, rowIndex, pickNums) {
        const effective = new Set(Array.isArray(pickNums) ? pickNums : []);
        if (effective.size < 2) {
            return [];
        }
        const lines = this.buildPickChainLinesBeforeRow(rows, rowIndex);
        const pairs = [];
        for (let li = 0; li < lines.length; li++) {
            const selInLine = lines[li].nums.filter((n) => effective.has(n));
            for (let i = 0; i < selInLine.length; i++) {
                for (let j = i + 1; j < selInLine.length; j++) {
                    pairs.push([selInLine[i], selInLine[j]]);
                }
            }
        }
        return pairs;
    }

    /**
     * Cặp chuỗi của bộ 3 trên một dòng: hai số đầu tiên (theo thứ tự trên dòng) trong {a,b,c}.
     * @param {number[]} lineNums
     * @param {number} a
     * @param {number} b
     * @param {number} c
     * @param {number} wantX
     * @param {number} wantY
     * @returns {boolean}
     */
    tripletChainPairMatchesOnLine(lineNums, a, b, c, wantX, wantY) {
        const triplet = new Set([a, b, c]);
        const selInLine = (Array.isArray(lineNums) ? lineNums : []).filter((n) => triplet.has(n));
        if (selInLine.length < 2) {
            return false;
        }
        const p0 = selInLine[0];
        const p1 = selInLine[1];
        return (p0 === wantX && p1 === wantY) || (p0 === wantY && p1 === wantX);
    }

    /**
     * Bộ 3 (a,b,c): mỗi cặp là cặp chuỗi trên 3 chuỗi khác nhau, freq từng số ≥ 2.
     * @param {{ label: number, nums: number[] }[]} lines
     * @param {number[]} freq
     * @param {number} a
     * @param {number} b
     * @param {number} c
     * @returns {boolean}
     */
    tripletSatisfiesConn3OnLines(lines, freq, a, b, c) {
        if (freq[a] < 2 || freq[b] < 2 || freq[c] < 2) {
            return false;
        }
        if (!lines || lines.length < 3) {
            return false;
        }
        /** @type {number[]} */
        const chainsAB = [];
        /** @type {number[]} */
        const chainsAC = [];
        /** @type {number[]} */
        const chainsBC = [];
        for (let li = 0; li < lines.length; li++) {
            const { label, nums } = lines[li];
            if (this.tripletChainPairMatchesOnLine(nums, a, b, c, a, b)) {
                chainsAB.push(label);
            }
            if (this.tripletChainPairMatchesOnLine(nums, a, b, c, a, c)) {
                chainsAC.push(label);
            }
            if (this.tripletChainPairMatchesOnLine(nums, a, b, c, b, c)) {
                chainsBC.push(label);
            }
        }
        for (let ab = 0; ab < chainsAB.length; ab++) {
            const la = chainsAB[ab];
            for (let ac = 0; ac < chainsAC.length; ac++) {
                const lac = chainsAC[ac];
                if (lac === la) {
                    continue;
                }
                for (let bc = 0; bc < chainsBC.length; bc++) {
                    const lbc = chainsBC[bc];
                    if (lbc !== la && lbc !== lac) {
                        return true;
                    }
                }
            }
        }
        return false;
    }

    /**
     * Cửa sổ 10 chuỗi trước kỳ: có tồn tại bộ 3-connection (không cần nằm trong đáp án kỳ).
     * @param {object[]} rows
     * @param {number} rowIndex
     * @returns {boolean}
     */
    rowWindowHasAnyConn3(rows, rowIndex) {
        const lines = this.buildPickChainLinesBeforeRow(rows, rowIndex);
        if (lines.length < 3) {
            return false;
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
        /** @type {number[]} */
        const candidates = [];
        for (let n = 1; n <= 35; n++) {
            if (freq[n] >= 2) {
                candidates.push(n);
            }
        }
        const len = candidates.length;
        if (len < 3) {
            return false;
        }
        for (let i = 0; i < len; i++) {
            for (let j = i + 1; j < len; j++) {
                for (let k = j + 1; k < len; k++) {
                    const a = candidates[i];
                    const b = candidates[j];
                    const c = candidates[k];
                    if (this.tripletSatisfiesConn3OnLines(lines, freq, a, b, c)) {
                        return true;
                    }
                }
            }
        }
        return false;
    }

    /**
     * 3-connection: bộ 3 số trong kết quả — mỗi cặp là cặp chuỗi (2 số đầu trên dòng)
     * trên một chuỗi riêng (3 chuỗi khác nhau), freq từng số ≥ 2.
     * @param {object[]} rows
     * @param {number} rowIndex
     * @returns {boolean}
     */
    rowMatchesConn3Filter(rows, rowIndex) {
        const row = rows[rowIndex];
        if (!row || this.isEmptyResultRow(row)) {
            return false;
        }
        const pickNums = this.parseMainNums(row.result || row.Result || '');
        if (pickNums.length < 3) {
            return false;
        }
        const lines = this.buildPickChainLinesBeforeRow(rows, rowIndex);
        if (lines.length < 3) {
            return false;
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
        const len = pickNums.length;
        for (let i = 0; i < len; i++) {
            for (let j = i + 1; j < len; j++) {
                for (let k = j + 1; k < len; k++) {
                    const a = pickNums[i];
                    const b = pickNums[j];
                    const c = pickNums[k];
                    if (this.tripletSatisfiesConn3OnLines(lines, freq, a, b, c)) {
                        return true;
                    }
                }
            }
        }
        return false;
    }

    /**
     * Một bộ gán 3 chuỗi cho cặp AB / AC / BC của bộ {a,b,c}.
     * @param {{ label: number, nums: number[] }[]} lines
     * @param {number} a
     * @param {number} b
     * @param {number} c
     * @returns {{ ab: number, ac: number, bc: number } | null}
     */
    getConn3TripletChainLabels(lines, a, b, c) {
        if (!lines || lines.length < 3) {
            return null;
        }
        /** @type {number[]} */
        const chainsAB = [];
        /** @type {number[]} */
        const chainsAC = [];
        /** @type {number[]} */
        const chainsBC = [];
        for (let li = 0; li < lines.length; li++) {
            const { label, nums } = lines[li];
            if (this.tripletChainPairMatchesOnLine(nums, a, b, c, a, b)) {
                chainsAB.push(label);
            }
            if (this.tripletChainPairMatchesOnLine(nums, a, b, c, a, c)) {
                chainsAC.push(label);
            }
            if (this.tripletChainPairMatchesOnLine(nums, a, b, c, b, c)) {
                chainsBC.push(label);
            }
        }
        for (let ab = 0; ab < chainsAB.length; ab++) {
            const la = chainsAB[ab];
            for (let ac = 0; ac < chainsAC.length; ac++) {
                const lac = chainsAC[ac];
                if (lac === la) {
                    continue;
                }
                for (let bc = 0; bc < chainsBC.length; bc++) {
                    const lbc = chainsBC[bc];
                    if (lbc !== la && lbc !== lac) {
                        return { ab: la, ac: lac, bc: lbc };
                    }
                }
            }
        }
        return null;
    }

    /**
     * Mọi bộ 3-connection trong cửa sổ 10 chuỗi trước kỳ (freq mỗi số ≥2).
     * @param {object[]} rows
     * @param {number} rowIndex
     * @returns {{ a: number, b: number, c: number, sorted: number[], chains: { ab: number, ac: number, bc: number } | null, inAnswer: boolean }[]}
     */
    enumerateConn3TripletsForRow(rows, rowIndex) {
        const chainLines = this.buildPickChainLinesBeforeRow(rows, rowIndex);
        if (chainLines.length < 3) {
            return [];
        }
        const freq = new Array(36).fill(0);
        for (let li = 0; li < chainLines.length; li++) {
            const nums = chainLines[li].nums;
            for (let k = 0; k < nums.length; k++) {
                const n = nums[k];
                if (n >= 1 && n <= 35) {
                    freq[n]++;
                }
            }
        }
        /** @type {number[]} */
        const candidates = [];
        for (let n = 1; n <= 35; n++) {
            if (freq[n] >= 2) {
                candidates.push(n);
            }
        }
        const row = rows[rowIndex];
        const pickNums = row && !this.isEmptyResultRow(row)
            ? this.parseMainNums(row.result || row.Result || '')
            : [];
        const pickSet = new Set(pickNums);
        const len = candidates.length;
        const seen = new Set();
        /** @type {{ a: number, b: number, c: number, sorted: number[], chains: { ab: number, ac: number, bc: number } | null, inAnswer: boolean }[]} */
        const out = [];
        for (let i = 0; i < len; i++) {
            for (let j = i + 1; j < len; j++) {
                for (let k = j + 1; k < len; k++) {
                    const a = candidates[i];
                    const b = candidates[j];
                    const c = candidates[k];
                    if (!this.tripletSatisfiesConn3OnLines(chainLines, freq, a, b, c)) {
                        continue;
                    }
                    const sorted = [a, b, c].sort((x, y) => x - y);
                    const key = sorted.join(',');
                    if (seen.has(key)) {
                        continue;
                    }
                    seen.add(key);
                    out.push({
                        a,
                        b,
                        c,
                        sorted,
                        chains: this.getConn3TripletChainLabels(chainLines, a, b, c),
                        inAnswer: pickSet.has(a) && pickSet.has(b) && pickSet.has(c)
                    });
                }
            }
        }
        out.sort((x, y) => {
            if (x.inAnswer !== y.inAnswer) {
                return x.inAnswer ? -1 : 1;
            }
            for (let t = 0; t < 3; t++) {
                if (x.sorted[t] !== y.sorted[t]) {
                    return x.sorted[t] - y.sorted[t];
                }
            }
            return 0;
        });
        return out;
    }

    /**
     * Textarea iframe trái: liệt kê toàn bộ 3-connection của kỳ đang focus.
     * @param {object[]} rows
     * @param {number} rowIndex
     * @returns {{ lines: string[], headerLines: string[], triplets: object[], footerLine: string }}
     */
    formatConn3ReferenceHint(rows, rowIndex) {
        const row = rows[rowIndex];
        const chainLines = this.buildPickChainLinesBeforeRow(rows, rowIndex);
        const pickNums = row && !this.isEmptyResultRow(row)
            ? this.parseMainNums(row.result || row.Result || '')
            : [];
        const triplets = this.enumerateConn3TripletsForRow(rows, rowIndex);
        const periodId = row ? String(row.id || row.ID || '').trim() : '';
        /** @type {string[]} */
        const lines = [];
        /** @type {string[]} */
        const headerLines = [];
        const idLabel = periodId ? `kỳ ${periodId}` : `dòng ${rowIndex + 1}`;
        const head1 = `3-connection — ${idLabel} (${chainLines.length} chuỗi trước kỳ)`;
        lines.push(head1);
        headerLines.push(head1);
        if (pickNums.length) {
            const ansLine = `Đáp án: ${pickNums.join(', ')}`;
            lines.push(ansLine);
            headerLines.push(ansLine);
        }
        lines.push('');
        if (!triplets.length) {
            const emptyMsg = 'Không có bộ 3-connection (freq mỗi số ≥2, 3 cặp trên 3 chuỗi khác nhau).';
            lines.push(emptyMsg);
            return { lines, headerLines, triplets: [], footerLine: emptyMsg };
        }
        /** @type {{ sorted: number[], inAnswer: boolean, label: string, chains: object | null }[]} */
        const tripletRows = [];
        for (let ti = 0; ti < triplets.length; ti++) {
            const t = triplets[ti];
            const nums = t.sorted.join(', ');
            let chainStr = '';
            if (t.chains) {
                chainStr = ` — AB:C${t.chains.ab} AC:C${t.chains.ac} BC:C${t.chains.bc}`;
            }
            const tag = t.inAnswer ? ' ★ đáp án' : '';
            const label = `${ti + 1}. {${nums}}${chainStr}${tag}`;
            lines.push(label);
            tripletRows.push({
                sorted: t.sorted.slice(),
                inAnswer: t.inAnswer,
                label,
                chains: t.chains
            });
        }
        const answerCount = triplets.filter((t) => t.inAnswer).length;
        const footerLine = `Tổng: ${triplets.length} bộ${answerCount ? ` (${answerCount} nằm trong đáp án)` : ''}.`;
        lines.push('');
        lines.push(footerLine);
        return { lines, headerLines, triplets: tripletRows, footerLine };
    }

    /**
     * Cached row indices for conn3 (3-connection) filter mode.
     * @returns {number[]}
     */
    ensureConn3FilterIndicesCache() {
        const rows = this.getSourceSheetRows();
        const n = rows.length;
        if (this._conn3FilterIndicesCache && this._conn3FilterIndicesCacheRowLen === n) {
            return this._conn3FilterIndicesCache;
        }
        const indices = [];
        for (let i = 0; i < n; i++) {
            if (this.rowMatchesConn3Filter(rows, i)) {
                indices.push(i);
            }
        }
        this._conn3FilterIndicesCache = indices;
        this._conn3FilterIndicesCacheRowLen = n;
        return this._conn3FilterIndicesCache;
    }

    /**
     * Mọi chỉ số dòng có cửa sổ 10 chuỗi trước kỳ chứa ít nhất một bộ 3-connection.
     * @returns {number[]}
     */
    ensureConn3WindowExistIndicesCache() {
        const rows = this.getSourceSheetRows();
        const n = rows.length;
        if (this._conn3WindowExistIndicesCache && this._conn3WindowExistIndicesCacheRowLen === n) {
            return this._conn3WindowExistIndicesCache;
        }
        const indices = [];
        for (let i = 0; i < n; i++) {
            if (!this.isEmptyResultRow(rows[i]) && this.rowWindowHasAnyConn3(rows, i)) {
                indices.push(i);
            }
        }
        this._conn3WindowExistIndicesCache = indices;
        this._conn3WindowExistIndicesCacheRowLen = n;
        return this._conn3WindowExistIndicesCache;
    }

    pairListSatisfiesConnection(pairList) {
        if (!pairList || pairList.length < 2) {
            return false;
        }
        /** @type {Map<number, Set<number>>} */
        const numToPairIdx = new Map();
        for (let pairIndex = 0; pairIndex < pairList.length; pairIndex++) {
            const pr = pairList[pairIndex];
            if (!pr || pr.length < 2) {
                continue;
            }
            const a = pr[0];
            const b = pr[1];
            if (!Number.isFinite(a) || !Number.isFinite(b)) {
                continue;
            }
            for (let k = 0; k < 2; k++) {
                const n = k === 0 ? a : b;
                let set = numToPairIdx.get(n);
                if (!set) {
                    set = new Set();
                    numToPairIdx.set(n, set);
                }
                set.add(pairIndex);
            }
        }
        for (const s of numToPairIdx.values()) {
            if (s.size >= 2) {
                return true;
            }
        }
        return false;
    }

    /**
     * Cặp pick → nhãn chuỗi (cửa sổ 10) chứa cặp đó.
     * @param {object[]} rows
     * @param {number} rowIndex
     * @param {number[]} pickNums
     * @returns {Map<string, Set<number>>}
     */
    getConnectionPairToChainsMap(rows, rowIndex, pickNums) {
        /** @type {Map<string, Set<number>>} */
        const pairToChains = new Map();
        const effective = new Set(Array.isArray(pickNums) ? pickNums : []);
        if (effective.size < 2) {
            return pairToChains;
        }
        const lines = this.buildPickChainLinesBeforeRow(rows, rowIndex);
        for (let li = 0; li < lines.length; li++) {
            const selInLine = lines[li].nums.filter((n) => effective.has(n));
            for (let i = 0; i < selInLine.length; i++) {
                for (let j = i + 1; j < selInLine.length; j++) {
                    const a = selInLine[i];
                    const b = selInLine[j];
                    const key = a < b ? `${a},${b}` : `${b},${a}`;
                    let set = pairToChains.get(key);
                    if (!set) {
                        set = new Set();
                        pairToChains.set(key, set);
                    }
                    set.add(lines[li].label);
                }
            }
        }
        return pairToChains;
    }

    /**
     * Connection duplicate: cùng một cặp {a,b} xuất hiện trên ≥2 chuỗi trong cửa sổ 10.
     * @param {object[]} rows
     * @param {number} rowIndex
     * @param {object} [row]
     * @returns {boolean}
     */
    rowHasDuplicateConnection(rows, rowIndex, row) {
        const r = row || rows[rowIndex];
        if (!r || this.isEmptyResultRow(r)) {
            return false;
        }
        const pickNums = this.parseMainNums(r.result || r.Result || '');
        if (pickNums.length < 2) {
            return false;
        }
        const pairs = this.buildChainPairsFromPickSet(rows, rowIndex, pickNums);
        if (!this.pairListSatisfiesConnection(pairs)) {
            return false;
        }
        const pairToChains = this.getConnectionPairToChainsMap(rows, rowIndex, pickNums);
        for (const chains of pairToChains.values()) {
            if (chains.size >= 2) {
                return true;
            }
        }
        return false;
    }

    /**
     * Connection unique: có connection nhưng không có cặp trùng trên ≥2 chuỗi.
     * @param {object[]} rows
     * @param {number} rowIndex
     * @param {object} [row]
     * @returns {boolean}
     */
    rowHasUniqueConnection(rows, rowIndex, row) {
        const r = row || rows[rowIndex];
        if (!r || this.isEmptyResultRow(r)) {
            return false;
        }
        const pickNums = this.parseMainNums(r.result || r.Result || '');
        if (pickNums.length < 2) {
            return false;
        }
        const pairs = this.buildChainPairsFromPickSet(rows, rowIndex, pickNums);
        if (!this.pairListSatisfiesConnection(pairs)) {
            return false;
        }
        return !this.rowHasDuplicateConnection(rows, rowIndex, r);
    }

    /**
     * Intersection (intersect) trên tập pick + 10 chuỗi trước kỳ — khớp pickSubmitSatisfiesIntersect.
     * @param {object[]} rows
     * @param {number} rowIndex
     * @param {number[]} pickNums
     * @returns {boolean}
     */
    rowPickSetSatisfiesIntersect(rows, rowIndex, pickNums) {
        const answerNums = (Array.isArray(pickNums) ? pickNums : [])
            .filter((n) => Number.isFinite(n) && n >= 1 && n <= 35)
            .sort((a, b) => a - b);
        if (answerNums.length < 2) {
            return false;
        }
        const lines = this.buildPickChainLinesBeforeRow(rows, rowIndex);
        if (!lines.length) {
            return false;
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
        for (let ai = 0; ai < answerNums.length; ai++) {
            for (let bi = ai + 1; bi < answerNums.length; bi++) {
                const a = answerNums[ai];
                const b = answerNums[bi];
                if (a === b || freq[a] < 2 || freq[b] < 2) {
                    continue;
                }
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
     * Nhãn C / ∩ sau chuỗi pick của kỳ (HTML gọn).
     * @param {object[]} rows
     * @param {number} rowIndex
     * @param {object} row
     * @returns {string}
     */
    getRowPickPropertyLabelHtml(rows, rowIndex, row) {
        const flags = this.getRowPickPropertyFlags(rows, rowIndex, row);
        if (!flags.conn3 && !flags.conn && !flags.ix) {
            return '';
        }
        let html = '<span class="row-pick-badges">';
        if (flags.conn3) {
            html += '<span class="row-pick-badge row-pick-badge--conn3" title="3-connection">3C</span>';
        } else {
            if (flags.conn) {
                html += '<span class="row-pick-badge row-pick-badge--conn" title="Connection">C</span>';
            }
            if (flags.ix) {
                html += '<span class="row-pick-badge row-pick-badge--ix" title="Intersection">∩</span>';
            }
        }
        html += '</span>';
        return html;
    }

    /**
     * Nhãn pick của kỳ: 3C đứng một mình (bao gồm C/∩); không 3C thì C và/hoặc ∩.
     * @param {object[]} rows
     * @param {number} rowIndex
     * @param {object} row
     * @returns {{ conn: boolean, conn3: boolean, ix: boolean }}
     */
    getRowPickPropertyFlags(rows, rowIndex, row) {
        if (!row || this.isEmptyResultRow(row)) {
            return { conn: false, conn3: false, ix: false };
        }
        const pickNums = this.parseMainNums(row.result || row.Result || '');
        if (pickNums.length < 2) {
            return { conn: false, conn3: false, ix: false };
        }
        const pairs = this.buildChainPairsFromPickSet(rows, rowIndex, pickNums);
        return {
            conn: this.pairListSatisfiesConnection(pairs),
            conn3: this.rowMatchesConn3Filter(rows, rowIndex),
            ix: this.rowPickSetSatisfiesIntersect(rows, rowIndex, pickNums)
        };
    }

    /**
     * Autoring sameRow: trong tập pick, chỉ khoanh số nào cùng nằm trên một chuỗi cửa sổ 10.
     * @param {object[]} rows
     * @param {number} rowIndex
     * @param {number[]} pickNums
     * @returns {number[]}
     */
    getSameRowCircleNumsFromPickSet(rows, rowIndex, pickNums) {
        const answerNums = (Array.isArray(pickNums) ? pickNums : [])
            .filter((n) => Number.isFinite(n) && n >= 1 && n <= 35);
        if (answerNums.length < 2) {
            return [];
        }
        const lines = this.buildPickChainLinesBeforeRow(rows, rowIndex);
        if (!lines.length) {
            return [];
        }
        const out = new Set();
        for (let li = 0; li < lines.length; li++) {
            const nums = lines[li].nums || [];
            const hits = [];
            for (let ai = 0; ai < answerNums.length; ai++) {
                if (nums.indexOf(answerNums[ai]) !== -1) {
                    hits.push(answerNums[ai]);
                }
            }
            if (hits.length >= 2) {
                for (let hi = 0; hi < hits.length; hi++) {
                    out.add(hits[hi]);
                }
            }
        }
        return Array.from(out).sort((a, b) => a - b);
    }

    /**
     * Số trong đáp án tham gia Connection (≥2 cặp trên 10 chuỗi).
     * @param {object[]} rows
     * @param {number} rowIndex
     * @param {number[]} pickNums
     * @returns {number[]}
     */
    getConnectionCircleNumsFromPickSet(rows, rowIndex, pickNums) {
        const pairs = this.buildChainPairsFromPickSet(rows, rowIndex, pickNums);
        if (!pairs || pairs.length < 2) {
            return [];
        }
        /** @type {Map<number, Set<number>>} */
        const numToPairIdx = new Map();
        for (let pairIndex = 0; pairIndex < pairs.length; pairIndex++) {
            const pr = pairs[pairIndex];
            if (!pr || pr.length < 2) {
                continue;
            }
            for (let k = 0; k < 2; k++) {
                const n = pr[k];
                if (!Number.isFinite(n)) {
                    continue;
                }
                let set = numToPairIdx.get(n);
                if (!set) {
                    set = new Set();
                    numToPairIdx.set(n, set);
                }
                set.add(pairIndex);
            }
        }
        const out = [];
        for (const [n, s] of numToPairIdx.entries()) {
            if (s.size >= 2) {
                out.push(n);
            }
        }
        return out.sort((a, b) => a - b);
    }

    /**
     * Số trong đáp án thuộc ít nhất một cặp Intersection hợp lệ.
     * @param {object[]} rows
     * @param {number} rowIndex
     * @param {number[]} pickNums
     * @returns {number[]}
     */
    getIntersectCircleNumsFromPickSet(rows, rowIndex, pickNums) {
        const answerNums = (Array.isArray(pickNums) ? pickNums : [])
            .filter((n) => Number.isFinite(n) && n >= 1 && n <= 35)
            .sort((a, b) => a - b);
        if (answerNums.length < 2) {
            return [];
        }
        const lines = this.buildPickChainLinesBeforeRow(rows, rowIndex);
        if (!lines.length) {
            return [];
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
        const hit = new Set();
        for (let ai = 0; ai < answerNums.length; ai++) {
            for (let bi = ai + 1; bi < answerNums.length; bi++) {
                const a = answerNums[ai];
                const b = answerNums[bi];
                if (a === b || freq[a] < 2 || freq[b] < 2) {
                    continue;
                }
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
                        hit.add(a);
                        hit.add(b);
                    }
                }
            }
        }
        return Array.from(hit).sort((a, b) => a - b);
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
     * Parse noteTRef filter option → list of exponents 1–10.
     * Multi-select = AND (mọi mũ phải có trong note), giống noteTags 📑.
     */
    parseNoteTRefExps(filterOptions) {
        const raw = (filterOptions || {}).noteTRef;
        if (raw == null || raw === '') {
            return [];
        }
        const list = Array.isArray(raw) ? raw : [raw];
        const out = [];
        for (let i = 0; i < list.length; i++) {
            const n = parseInt(list[i], 10);
            if (Number.isFinite(n) && n >= 1 && n <= 10 && out.indexOf(n) < 0) {
                out.push(n);
            }
        }
        return out;
    }

    /**
     * Decode Unicode superscript digits (¹…⁰) to integer.
     * Empty / invalid → null (không mặc định 1 — tránh khớp nhầm ngày kiểu 12-05).
     */
    decodeNoteRefExponent(supStr) {
        const map = {
            '¹': '1', '²': '2', '³': '3', '⁴': '4', '⁵': '5',
            '⁶': '6', '⁷': '7', '⁸': '8', '⁹': '9', '⁰': '0'
        };
        const raw = String(supStr || '');
        if (!raw.length) {
            return null;
        }
        let digits = '';
        for (let i = 0; i < raw.length; i++) {
            const d = map[raw[i]];
            if (!d) {
                return null;
            }
            digits += d;
        }
        const v = parseInt(digits, 10);
        return Number.isFinite(v) ? v : null;
    }

    /**
     * Note có ít nhất một tham chiếu dạng `id-prevIdⁿ=…` với mũ Unicode = exp (1–10).
     * Chỉ đọc note tính toán (không gộp raw) để tránh khớp nhầm.
     */
    noteContainsRefExponent(noteText, exp) {
        const target = parseInt(exp, 10);
        if (!noteText || !Number.isFinite(target) || target < 1 || target > 10) {
            return false;
        }
        // Bắt buộc có mũ + dấu = ngay sau (cùng shape buildNoteForRow).
        const reRef = /([0-9]+)-([0-9]+)([¹²³⁴⁵⁶⁷⁸⁹⁰]+)=/g;
        let m;
        while ((m = reRef.exec(noteText)) !== null) {
            if (this.decodeNoteRefExponent(m[3]) === target) {
                return true;
            }
        }
        return false;
    }

    /**
     * Row note must contain every selected tRef exponent (AND / ∩), giống 📑 noteTags.
     * Chọn t^1 và t^2 → note phải có cả mũ ¹ và ² (không phải OR).
     * @param {number} rowIndex
     * @param {number|number[]} exps
     */
    rowMatchesNoteTRefFilter(rowIndex, exps) {
        let list;
        if (Array.isArray(exps)) {
            list = [];
            for (let i = 0; i < exps.length; i++) {
                const n = parseInt(exps[i], 10);
                if (Number.isFinite(n) && n >= 1 && n <= 10 && list.indexOf(n) < 0) {
                    list.push(n);
                }
            }
        } else {
            const n = parseInt(exps, 10);
            list = Number.isFinite(n) && n >= 1 && n <= 10 ? [n] : [];
        }
        if (!list.length) {
            return true;
        }
        const rows = this.getSourceSheetRows();
        const row = rows[rowIndex];
        if (!row) {
            return false;
        }
        const meta = this.getComputedNoteMeta(rowIndex, row);
        const noteText = (meta && meta.text && meta.text !== '?') ? String(meta.text) : '';
        if (!noteText) {
            return false;
        }
        // AND: thiếu bất kỳ mũ nào → loại
        for (let i = 0; i < list.length; i++) {
            if (!this.noteContainsRefExponent(noteText, list[i])) {
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
     * Count nonexist numbers matching selected colors (green respects optional B/I/U/S styles).
     */
    countNonexistNumbersForColorFilter(rowIndex, colors, styles) {
        const colorList = Array.isArray(colors) ? colors : [];
        if (colorList.length === 0) {
            return 0;
        }
        const styleList = Array.isArray(styles) ? styles : [];
        const entries = this.ensureNonexistDisplayEntriesCache()[rowIndex] || [];
        let count = 0;
        for (let i = 0; i < entries.length; i++) {
            const entry = entries[i];
            if (!colorList.includes(entry.color)) {
                continue;
            }
            if (entry.color === 'green') {
                if (styleList.length === 0) {
                    count++;
                    continue;
                }
                for (let j = 0; j < styleList.length; j++) {
                    if (this.nonexistGreenKindMatchesStyle(entry.kind, styleList[j])) {
                        count++;
                        break;
                    }
                }
                continue;
            }
            count++;
        }
        return count;
    }

    /**
     * [] filter: row matches when count of numbers in selected color(s) satisfies op vs threshold.
     */
    rowMatchesNonexistBracketFilter(rowIndex, threshold, op = '>=', colors = null, styles = null) {
        const colorList = Array.isArray(colors)
            ? colors.filter((c) => ['green', 'red', 'purple', 'yellow'].includes(c))
            : [];
        const count = colorList.length > 0
            ? this.countNonexistNumbersForColorFilter(rowIndex, colorList, styles)
            : 0;
        return this.freqMatchesComparison(count, threshold, op);
    }

    /**
     * Specific nonexist number (1–35) with selected colors; green uses styles when provided.
     */
    rowMatchesNonexistSpecificNumFilter(rowIndex, num, colors, styles) {
        const targetNum = parseInt(num, 10);
        if (!Number.isFinite(targetNum) || targetNum < 1 || targetNum > 35) {
            return false;
        }
        const colorList = Array.isArray(colors) ? colors : [];
        if (colorList.length === 0) {
            return false;
        }
        const styleList = Array.isArray(styles) ? styles : [];
        if (styleList.length > 0) {
            return this.rowMatchesNonexistColorFilter(rowIndex, colorList, styleList, targetNum);
        }
        const entries = this.ensureNonexistDisplayEntriesCache()[rowIndex] || [];
        for (let i = 0; i < entries.length; i++) {
            const entry = entries[i];
            if (entry.num !== targetNum) {
                continue;
            }
            if (colorList.includes(entry.color)) {
                return true;
            }
        }
        return false;
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
     * ID greens use note-frequency as of that moment (notes before focus only — like chưa có result).
     */
    setAnswerPopupFocusMask(opts) {
        const o = opts || {};
        const open = !!o.open;
        const rowIndex = Number.isFinite(o.rowIndex) ? o.rowIndex : -1;
        const submitOn = !!o.submitOn;
        const nextActive = open && rowIndex >= 0 && !submitOn;
        const nextRow = open ? rowIndex : -1;
        const prev = this.answerPopupFocusMask || {};
        if (prev.active !== nextActive || prev.rowIndex !== nextRow) {
            this._idFreqAsOfCacheRow = -1;
            this._idFreqAsOfCache = null;
        }
        this.answerPopupFocusMask = {
            active: nextActive,
            rowIndex: nextRow
        };
    }

    shouldAnswerPopupMaskSheet1Row(rowIndex) {
        const m = this.answerPopupFocusMask || {};
        return !!(m.active && m.rowIndex === rowIndex);
    }

    /**
     * Id frequency map "at focus moment": only notes of rows before focus
     * (as if focus row chưa có result and later periods do not exist yet).
     */
    getIdFrequencyMapAsOfRow(rowIndex) {
        const idx = Number(rowIndex);
        if (!Number.isFinite(idx) || idx < 0) {
            return this.idFrequencyMap;
        }
        if (!this.noteCache) {
            this.refreshDerivedState();
        }
        if (this._idFreqAsOfCacheRow === idx && this._idFreqAsOfCache) {
            return this._idFreqAsOfCache;
        }
        const truncated = (this.noteCache || []).slice(0, idx);
        const map = this.buildIdFrequencyMapFromNotes(truncated);
        this._idFreqAsOfCacheRow = idx;
        this._idFreqAsOfCache = map;
        return map;
    }

    getEffectiveIdFrequencyMap() {
        const m = this.answerPopupFocusMask || {};
        if (m.active && m.rowIndex >= 0) {
            return this.getIdFrequencyMapAsOfRow(m.rowIndex);
        }
        return this.idFrequencyMap;
    }

    /**
     * Repaint sheet1 id cells from effective frequency map (full vs answer-popup as-of).
     */
    applyIdBackgroundsForAnswerPopupMask(tableWrap) {
        if (!tableWrap) {
            return;
        }
        const rows = tableWrap.id === 'filterTableWrap'
            ? (this.getSourceSheetRows() || [])
            : (this.dataRows || []);
        const freqMap = this.getEffectiveIdFrequencyMap();
        const trs = tableWrap.querySelectorAll('tbody tr[data-idx]');
        for (let t = 0; t < trs.length; t++) {
            const tr = trs[t];
            const idx = Number(tr.dataset.idx);
            if (!Number.isFinite(idx) || idx < 0 || idx >= rows.length) {
                continue;
            }
            const idCell = tr.querySelector('td.cell-id');
            if (!idCell) {
                continue;
            }
            const row = rows[idx] || {};
            const bg = this.getIdBackgroundByFrequency(row.id || row.ID || '', freqMap);
            if (bg) {
                idCell.setAttribute('style', `background:${bg};`);
            } else {
                idCell.removeAttribute('style');
            }
        }
    }

    /**
     * Nonexist HTML for one source row (respects Answer-popup focus mask = empty-result styling).
     */
    renderSourceRowNonexistCellHtml(rowIndex, row, options = {}) {
        const maskRow = this.shouldAnswerPopupMaskSheet1Row(rowIndex);
        const source = row || {};
        const result = source.result || source.Result || '';
        if (maskRow) {
            const emptyRow = Object.assign({}, source, { result: '', Result: '' });
            const meta = this.getNonexistMetaForSourceRow(rowIndex, emptyRow);
            return this.renderNonexistHtml(rowIndex, meta.text, '', options);
        }
        const meta = this.getNonexistMetaForSourceRow(rowIndex, source);
        return this.renderNonexistHtml(rowIndex, meta.text, result, options);
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

    scheduleApplyFilterAllModeFocusMask(tableWrap) {
        if (this._filterAllModeMaskApplyRaf) {
            return;
        }
        const wrap = tableWrap || document.getElementById('filterTableWrap');
        this._filterAllModeMaskApplyRaf = requestAnimationFrame(() => {
            this._filterAllModeMaskApplyRaf = 0;
            this.applyFilterAllModeFocusMaskToDom(wrap);
        });
    }

    applyAnswerPopupFocusMaskToDom(tableWrap, options = {}) {
        if (!tableWrap || this.activeSheet !== 'sheet1') {
            return;
        }
        this._applySourceFocusPreviewMaskToDom(
            tableWrap,
            this.answerPopupFocusMask,
            '_answerPopupMaskAppliedRow',
            options
        );
    }

    applyFilterAllModeFocusMaskToDom(tableWrap, options = {}) {
        if (!tableWrap || tableWrap.id !== 'filterTableWrap') {
            return;
        }
        this._applySourceFocusPreviewMaskToDom(
            tableWrap,
            this.answerPopupFocusMask,
            '_filterAllModeMaskAppliedRow',
            options
        );
    }

    _applySourceFocusPreviewMaskToDom(tableWrap, maskState, appliedRowKey, options = {}) {
        if (!tableWrap) {
            return;
        }

        const m = maskState || {};
        const prevIdx = this[appliedRowKey];
        const nextIdx = m.active ? m.rowIndex : -1;
        const force = !!options.reset;
        const maskChanged = prevIdx !== nextIdx || force;

        if (prevIdx >= 0) {
            if (prevIdx !== nextIdx || (force && nextIdx < 0)) {
                this.setFocusPreviewMaskOnRowDom(tableWrap, prevIdx, false);
            }
        }
        if (nextIdx >= 0 && (nextIdx !== prevIdx || force)) {
            this.setFocusPreviewMaskOnRowDom(tableWrap, nextIdx, true);
        }

        this[appliedRowKey] = nextIdx;

        if (maskChanged) {
            this.applyIdBackgroundsForAnswerPopupMask(tableWrap);
            /* Mask đổi → trail tím/đỏ ngoài chuỗi 10 đổi theo; clear cache + refresh */
            this._focusNonexistTrailKey = null;
            this._focusNonexistTrailSet = null;
            const win = this.activeWindowRange;
            if (win && typeof win.start === 'number' && typeof win.end === 'number') {
                this.refreshNonexistCellsForRowIndices(
                    tableWrap,
                    this.collectNonexistBoostRefreshRowIndices(win),
                    {
                        forFilterPopup: tableWrap.id === 'filterTableWrap',
                        windowRange: win
                    }
                );
            }
        }
    }

    setFocusPreviewMaskOnRowDom(tableWrap, rowIndex, masked) {
        const tr = tableWrap.querySelector(`tbody tr[data-idx="${rowIndex}"]`);
        if (!tr) {
            return;
        }
        tr.classList.toggle('answer-popup-focus-masked', masked);
        const nonexistCell = tr.querySelector('td.cell-nonexist');
        const rows = tableWrap.id === 'filterTableWrap'
            ? (this.getSourceSheetRows() || [])
            : (this.dataRows || []);
        const row = rows[rowIndex];
        if (!nonexistCell || !row) {
            return;
        }
        nonexistCell.classList.toggle('answer-popup-focus-nonexist', masked);
        const win = this.activeWindowRange;
        nonexistCell.innerHTML = this.renderSourceRowNonexistCellHtml(rowIndex, row, {
            windowRange: (win && typeof win.start === 'number' && typeof win.end === 'number')
                ? win
                : null
        });
    }

    setAnswerPopupFocusMaskOnRowDom(tableWrap, rowIndex, masked) {
        this.setFocusPreviewMaskOnRowDom(tableWrap, rowIndex, masked);
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

        let html = '<table class="sheet-data-table sheet1-source-table"><thead><tr><th>date</th><th>id</th><th class="cell-pick-label-h">label</th><th class="cell-follow-h">follow</th><th>result</th><th>note</th><th>nonexist</th></tr></thead><tbody>';

        const displayRows = rows || [];
        const rowIndices = Array.isArray(options.indices)
            ? options.indices.filter(i => i >= 0 && i < displayRows.length)
            : displayRows.map((_, i) => i);
        const prevRecallFoldStatsByChain = this.computePrevPeriodRecallFoldStatsByChain(displayRows, rowIndices);
        const prevRecallFoldPctLabel = this.formatPrevPeriodRecallFoldPctByChain(prevRecallFoldStatsByChain);
        const prevRecallFoldPctAttr = this.encodePrevPeriodRecallFoldTooltipAttr(prevRecallFoldPctLabel);

        for (const i of rowIndices) {
            const row = displayRows[i];
            const date = row.date || row.Date || '';
            const id = row.id || row.ID || '';
            const result = row.result || row.Result || '';
            const isEmptyResultRow = this.isEmptyResultRow(row);
            const noteMeta = isEmptyResultRow
                ? { text: '', highlightYellow: false }
                : this.getComputedNoteMeta(i, row);
            const idBg = this.getIdBackgroundByFrequency(id, this.getEffectiveIdFrequencyMap());
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
                ? `<span class="prev-period-recall-fold" data-pct="${prevRecallFoldPctAttr}"></span>`
                : '';
            const pickLabelHtml = isEmptyResultRow ? '' : this.getRowPickPropertyLabelHtml(displayRows, i, row);
            const followValue = this.computeFollowCellValue(displayRows, i);
            const followHtml = followValue === '?'
                ? '<span class="cell-follow-undetermined" title="≥2 số cùng freq cao nhất ở chuỗi 1">?</span>'
                : this.escapeHtml(followValue);

            html += `<tr data-idx="${i}" class="data-row${activeClass}" data-has-result="${!!result}" data-empty="${isEmptyResultRow ? '1' : '0'}">
                <td class="cell-date"${dateBg}>${date}</td>
                <td class="cell-id"${idStyle}>${id}</td>
                <td class="cell-pick-label">${pickLabelHtml}</td>
                <td class="cell-follow">${followHtml}</td>
                <td class="${resultCellClass}">${prevRecallFoldHit}${resultHtml}</td>
                <td class="cell-note"${noteStyle}>${noteHtml}</td>
                <td class="cell-nonexist">${nonexistHtml}</td>
            </tr>`;
        }
        html += '</tbody></table>';
        tableWrap.innerHTML = html;

        this.bindSourceSheetTableAfterRender(tableWrap, options);
        if (this.activeSheet === 'sheet1' && tableWrap.id === 'tableWrap') {
            this.cacheSheet1TableDom(tableWrap);
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
     * Skips display:none rows (F%/U% subfilters) so geometry is not (0 + end)/2 → jump to bottom.
     */
    centerActiveWindowInView(tableWrap, options = {}) {
        if (!tableWrap) {
            return;
        }

        let startIdx;
        let endIdx;
        if (typeof options.startIdx === 'number' && typeof options.endIdx === 'number') {
            startIdx = options.startIdx;
            endIdx = options.endIdx;
        } else if (this.activeWindowRange) {
            startIdx = this.activeWindowRange.start;
            endIdx = this.activeWindowRange.end;
        } else {
            return;
        }
        if (typeof startIdx !== 'number' || typeof endIdx !== 'number' || endIdx < startIdx) {
            return;
        }

        const isRowLaidOut = (tr) => {
            if (!tr) {
                return false;
            }
            // display:none → no client rects; content-visibility skip still has a rect/placeholder
            return tr.getClientRects().length > 0;
        };

        const resolveVisibleWindowRows = () => {
            let first = null;
            let last = null;
            for (let i = startIdx; i <= endIdx; i++) {
                const tr = tableWrap.querySelector(`tbody tr[data-idx="${i}"]`);
                if (!isRowLaidOut(tr)) {
                    continue;
                }
                if (!first) {
                    first = tr;
                }
                last = tr;
            }
            if (first && last) {
                return { startRow: first, endRow: last };
            }
            const fallback = tableWrap.querySelector(`tbody tr[data-idx="${endIdx}"]`)
                || tableWrap.querySelector(`tbody tr[data-idx="${startIdx}"]`);
            if (isRowLaidOut(fallback)) {
                return { startRow: fallback, endRow: fallback };
            }
            return null;
        };

        const applyCentering = () => {
            const bounds = resolveVisibleWindowRows();
            if (!bounds) {
                return;
            }
            const wrapRect = tableWrap.getBoundingClientRect();
            const startRect = bounds.startRow.getBoundingClientRect();
            const endRect = bounds.endRow.getBoundingClientRect();
            if (startRect.height <= 0 && endRect.height <= 0) {
                return;
            }

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
                html += `<td class="cell-col-h blank-cell combo-h-comment-cell"><input type="text" class="combo-cell-input combo-h-comment-input" data-combo-h-row="${rowIndex}" value="${this.escapeHtml(hComment)}" aria-label="H${rowIndex}" title="Enter: khoanh trái theo phần trước | (2–5 số 1–35, phân tách bằng dấu phẩy; phần sau | không khoanh)" spellcheck="false" /></td>`;
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
                const sourceRows = this.getSourceSheetRows();
                const byId = sourceRows.findIndex((r) => String(r.id || r.ID || '').trim() === this.comboFocusRowId);
                this.comboFocusRowIndex = byId >= 0 ? byId : -1;
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
     * Chỉ phần trước dấu | dùng để khoanh; phần sau | là ghi chú (vd. số đặc biệt), bỏ qua.
     * Dấu phân tách số khoanh: dấu phẩy, chấm phẩy hoặc khoảng trắng.
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
        const pipeIdx = t.indexOf('|');
        const pickSegment = pipeIdx >= 0 ? t.slice(0, pipeIdx).trim() : t;
        if (!pickSegment) {
            return null;
        }
        const parts = pickSegment.split(/[\s,;]+/).map((x) => x.trim()).filter(Boolean);
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
            frame.contentWindow.postMessage({
                type: 'syncAnswerPickSelection',
                nums,
                comboHPickSync: true
            }, '*');
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

        this.clearWindowBorderClassesOnWrap(tableWrap);
        this.clearIdRefHighlightFromDom(tableWrap);
        this.clearFocusNoteRefHighlightFromDom(tableWrap);
    }

    /**
     * Chỉ gỡ viền cửa sổ 10 trên một bảng — không đụng viền cyan/đỏ trên bảng kia.
     * @param {HTMLElement} tableWrapEl
     */
    clearWindowBorderClassesOnWrap(tableWrapEl) {
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
     * Vẽ lại viền cyan/đỏ theo activeWindowRange (sau khi filter popup chỉ đổi focus hàng).
     * @param {HTMLElement} [tableWrapEl]
     */
    reapplyActiveWindowHighlightsToDom(tableWrapEl) {
        const r = this.activeWindowRange;
        if (!r) {
            return;
        }
        if (Array.isArray(r.idRefHighlightIndices) && r.idRefHighlightIndices.length) {
            this.applyIdRefHighlightToDom(r.idRefHighlightIndices, tableWrapEl);
        }
        if (Array.isArray(r.focusNoteRefHighlightIndices) && r.focusNoteRefHighlightIndices.length) {
            this.applyFocusNoteRefHighlightToDom(r.focusNoteRefHighlightIndices, tableWrapEl);
        }
    }

    /**
     * @returns {HTMLElement[]}
     */
    getIdRefHighlightTableWraps(extraWrap) {
        const wraps = [];
        if (extraWrap) {
            wraps.push(extraWrap);
        }
        if (typeof document !== 'undefined') {
            const mainWrap = document.getElementById('tableWrap');
            const filterWrap = document.getElementById('filterTableWrap');
            if (mainWrap && !wraps.includes(mainWrap)) {
                wraps.push(mainWrap);
            }
            if (filterWrap && !wraps.includes(filterWrap)) {
                wraps.push(filterWrap);
            }
        }
        return wraps;
    }

    clearIdRefHighlightFromDom(tableWrapEl) {
        const wraps = this.getIdRefHighlightTableWraps(tableWrapEl);
        for (let w = 0; w < wraps.length; w++) {
            const wrap = wraps[w];
            if (!wrap) {
                continue;
            }
            wrap.querySelectorAll('td.cell-note.id-ref-contextmenu-highlight, td.cell-result.id-ref-contextmenu-highlight').forEach((cell) => {
                cell.classList.remove('id-ref-contextmenu-highlight');
            });
        }
    }

    /**
     * Viền xanh tham chiếu id — gắn cửa sổ trượt 10 (mất khi đổi focus).
     * @param {number[]} refRowIndices
     * @param {HTMLElement} [tableWrapEl]
     */
    applyIdRefHighlightToDom(refRowIndices, tableWrapEl) {
        if (!Array.isArray(refRowIndices) || !refRowIndices.length) {
            return;
        }
        const wraps = this.getIdRefHighlightTableWraps(tableWrapEl);
        for (let w = 0; w < wraps.length; w++) {
            const wrap = wraps[w];
            if (!wrap) {
                continue;
            }
            for (let r = 0; r < refRowIndices.length; r++) {
                const rowIdx = refRowIndices[r];
                const tr = wrap.querySelector(`tbody tr[data-idx="${rowIdx}"]`);
                if (!tr) {
                    continue;
                }
                const noteCell = tr.querySelector('td.cell-note');
                if (noteCell) {
                    noteCell.classList.add('id-ref-contextmenu-highlight');
                }
                const resultCell = tr.querySelector('td.cell-result');
                if (resultCell) {
                    resultCell.classList.add('id-ref-contextmenu-highlight');
                }
            }
        }
    }

    clearFocusNoteRefHighlightFromDom(tableWrapEl) {
        const wraps = this.getIdRefHighlightTableWraps(tableWrapEl);
        for (let w = 0; w < wraps.length; w++) {
            const wrap = wraps[w];
            if (!wrap) {
                continue;
            }
            wrap.querySelectorAll('td.cell-note.id-focus-note-ref-highlight, td.cell-result.id-focus-note-ref-highlight').forEach((cell) => {
                cell.classList.remove('id-focus-note-ref-highlight');
            });
        }
    }

    /**
     * Viền đỏ: các kỳ trong cửa sổ 10 mà note của hàng focus tham chiếu tới.
     * @param {number[]} refRowIndices
     * @param {HTMLElement} [tableWrapEl]
     */
    applyFocusNoteRefHighlightToDom(refRowIndices, tableWrapEl) {
        if (!Array.isArray(refRowIndices) || !refRowIndices.length) {
            return;
        }
        const wraps = this.getIdRefHighlightTableWraps(tableWrapEl);
        for (let w = 0; w < wraps.length; w++) {
            const wrap = wraps[w];
            if (!wrap) {
                continue;
            }
            for (let r = 0; r < refRowIndices.length; r++) {
                const rowIdx = refRowIndices[r];
                const tr = wrap.querySelector(`tbody tr[data-idx="${rowIdx}"]`);
                if (!tr) {
                    continue;
                }
                const noteCell = tr.querySelector('td.cell-note');
                if (noteCell) {
                    noteCell.classList.add('id-focus-note-ref-highlight');
                }
                const resultCell = tr.querySelector('td.cell-result');
                if (resultCell) {
                    resultCell.classList.add('id-focus-note-ref-highlight');
                }
            }
        }
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
        const prevWindowRange = this.activeWindowRange;

        this.clearWindowSelection(tableWrap);

        if (startIdx === null || endIdx === null || endIdx < startIdx) {
            this.activeWindowRange = null;
            if (!previewOnly) {
                this.refreshNonexistCellsForActiveWindow(tableWrap);
            }
            return;
        }

        let idRefHighlightIndices = null;
        if (options && Array.isArray(options.idRefHighlightIndices)) {
            idRefHighlightIndices = options.idRefHighlightIndices.slice();
        } else if (previewOnly && prevWindowRange && Array.isArray(prevWindowRange.idRefHighlightIndices)) {
            idRefHighlightIndices = prevWindowRange.idRefHighlightIndices.slice();
        }
        let focusNoteRefHighlightIndices = null;
        if (typeof targetIdx === 'number' && targetIdx >= 0) {
            if (options && options.skipFocusNoteRef) {
                focusNoteRefHighlightIndices = [];
            } else if (options && Array.isArray(options.focusNoteRefHighlightIndices)) {
                focusNoteRefHighlightIndices = options.focusNoteRefHighlightIndices.slice();
            } else {
                focusNoteRefHighlightIndices = this.findWindowRowIndicesReferencedInFocusNote(targetIdx);
            }
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

        this.activeWindowRange = {
            start: startIdx,
            end: endIdx,
            target: targetIdx,
            idRefHighlightIndices,
            focusNoteRefHighlightIndices
        };
        this._focusNonexistTrailKey = null;
        if (previewOnly) {
            if (idRefHighlightIndices && idRefHighlightIndices.length) {
                this.applyIdRefHighlightToDom(idRefHighlightIndices, tableWrap);
            }
            if (focusNoteRefHighlightIndices && focusNoteRefHighlightIndices.length) {
                this.applyFocusNoteRefHighlightToDom(focusNoteRefHighlightIndices, tableWrap);
            }
            this._focusNonexistTrailKey = null;
            const previewRefresh = new Set();
            if (prevWindowRange) {
                for (const i of this.collectNonexistBoostRefreshRowIndices(prevWindowRange)) {
                    previewRefresh.add(i);
                }
            }
            for (const i of this.collectNonexistBoostRefreshRowIndices(this.activeWindowRange)) {
                previewRefresh.add(i);
            }
            this.refreshNonexistCellsForRowIndices(tableWrap, previewRefresh);
            this.renderWindowLabels(startIdx, endIdx, tableWrap);
            return;
        }
        if (idRefHighlightIndices && idRefHighlightIndices.length) {
            this.applyIdRefHighlightToDom(idRefHighlightIndices, tableWrap);
        }
        if (focusNoteRefHighlightIndices && focusNoteRefHighlightIndices.length) {
            this.applyFocusNoteRefHighlightToDom(focusNoteRefHighlightIndices, tableWrap);
        }
        // Refresh nonexist HTML first; renderWindowLabels after so innerHTML does not strip labels.
        const refreshIndices = new Set();
        if (prevWindowRange) {
            for (const i of this.collectNonexistBoostRefreshRowIndices(prevWindowRange)) {
                refreshIndices.add(i);
            }
        }
        for (const i of this.collectNonexistBoostRefreshRowIndices({
            start: startIdx,
            end: endIdx,
            target: targetIdx !== null ? targetIdx : endIdx
        })) {
            refreshIndices.add(i);
        }
        this.refreshNonexistCellsForRowIndices(tableWrap, refreshIndices);
        this.renderWindowLabels(startIdx, endIdx, tableWrap);
    }

    /**
     * Cập nhật activeWindowRange sheet1 trong bộ nhớ (không vẽ DOM).
     * Dùng khi đang xem sheet khác (tracking/combo) nhưng cần giữ focus sheet1 đồng bộ.
     */
    commitSourceSheetWindowRangeForIndex(idx, options = {}) {
        if (typeof idx !== 'number' || idx < 0) {
            return false;
        }
        const rows = this.getSourceSheetRows();
        if (idx >= rows.length) {
            return false;
        }
        const start = Math.max(0, idx - 10);
        const end = idx;
        const preserveIdRefHighlights = options.preserveIdRefHighlights === true;
        const prevWindowRange = this.activeWindowRange;
        let idRefHighlightIndices = null;
        if (Array.isArray(options.idRefHighlightIndices)) {
            idRefHighlightIndices = options.idRefHighlightIndices.slice();
        } else if (preserveIdRefHighlights && prevWindowRange
            && Array.isArray(prevWindowRange.idRefHighlightIndices)) {
            idRefHighlightIndices = prevWindowRange.idRefHighlightIndices.slice();
        }
        let focusNoteRefHighlightIndices = null;
        if (Array.isArray(options.focusNoteRefHighlightIndices)) {
            focusNoteRefHighlightIndices = options.focusNoteRefHighlightIndices.slice();
        } else if (options.skipFocusNoteRef) {
            focusNoteRefHighlightIndices = [];
        } else {
            focusNoteRefHighlightIndices = this.findWindowRowIndicesReferencedInFocusNote(idx);
        }
        this.activeWindowRange = {
            start,
            end,
            target: idx,
            idRefHighlightIndices,
            focusNoteRefHighlightIndices
        };
        return true;
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
        if (currentIdx < 0 && !(this.activeSheet === 'sheet1' && this._sheet1LeftPaneTarget >= 0)) {
            return false;
        }

        let baseIdx = currentIdx;
        if (this.activeSheet === 'sheet1' && this._sheet1LeftPaneTarget >= 0) {
            baseIdx = this._sheet1LeftPaneTarget;
        }

        const nextIdx = Math.max(0, Math.min(displayRows.length - 1, baseIdx + delta));
        const sheet1PumpBehind = this.activeSheet === 'sheet1'
            && this._sheet1LeftPaneTarget >= 0
            && this._sheet1LeftPaneCurrent !== this._sheet1LeftPaneTarget;
        if (nextIdx === baseIdx && !sheet1PumpBehind) {
            return false;
        }

        if (this.activeSheet === 'sheet1') {
            return this.scheduleSheet1NavToIndex(nextIdx, wrap);
        }

        const start = Math.max(0, nextIdx - 10);
        this.applyWindowSelection(start, nextIdx, nextIdx, wrap, { previewOnly: true });
        this.centerActiveWindowInView(wrap);
        return true;
    }

    /**
     * After arrow burst ends: flush pump, refresh right pane, sync reference hint once.
     */
    finishSheet1ArrowNav(tableWrap) {
        const wrap = tableWrap || document.getElementById('tableWrap');
        if (!wrap || this.activeSheet !== 'sheet1') {
            return false;
        }

        const targetIdx = this._sheet1LeftPaneTarget >= 0
            ? this._sheet1LeftPaneTarget
            : (this.activeWindowRange && typeof this.activeWindowRange.target === 'number'
                ? this.activeWindowRange.target
                : -1);
        if (targetIdx < 0) {
            return false;
        }

        this.flushSheet1NavPump();

        const start = Math.max(0, targetIdx - 10);
        this.applyWindowSelection(start, targetIdx, targetIdx, wrap);
        this.centerActiveWindowInView(wrap);

        const rows = this.getSourceSheetRows();
        const data = this.buildSourceRowLeftPaneData(targetIdx, rows, { lightStep: false });
        if (data) {
            const refBundle = this.buildConn3ReferenceDetailForRow(targetIdx);
            this.applySourceRowFocusState(targetIdx, rows, { asSheet1: true, prefetchedData: data });
            window.dispatchEvent(new CustomEvent('leftPaneSetLines', {
                detail: {
                    ...data,
                    ...refBundle,
                    trackingFrameStep: false,
                    preserveSelection: true,
                    skipReferenceMeta: false
                }
            }));
        }

        this.resetSheet1LeftPanePump();
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

        this.resetSheet1LeftPanePump();
        this.focusSourceSheetRow(targetIdx, {
            skipSave: true,
            light: true
        });
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
     * Gợi ý tham chiếu 3-connection (panel iframe trái).
     * @returns {{ text: string, conn3HeaderLines: string[], conn3Triplets: object[], conn3FooterLine: string } | { error: string } | null}
     */
    getNoteReferenceHintMeta(rowIndex) {
        const rows = this.getSourceSheetRows();
        const idx = Number(rowIndex);
        if (!Number.isFinite(idx) || idx < 0 || idx >= rows.length) {
            return null;
        }
        const r = this.formatConn3ReferenceHint(rows, idx);
        return {
            text: r.lines.join('\n'),
            conn3HeaderLines: r.headerLines || [],
            conn3Triplets: Array.isArray(r.triplets) ? r.triplets : [],
            conn3FooterLine: r.footerLine || ''
        };
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

    /** Hàng đầu streak nonexist liên tiếp của `num` tính ngược từ fromRowIndex. */
    getNonexistStreakStartRow(fromRowIndex, num) {
        if (!this.nonexistCache || fromRowIndex < 0) {
            return -1;
        }
        const streak = this.countNonexistStreak(fromRowIndex, num, 0);
        if (streak <= 0) {
            return -1;
        }
        return fromRowIndex - streak + 1;
    }

    numInNonexistCacheRow(rowIdx, num) {
        const meta = this.nonexistCache && this.nonexistCache[rowIdx];
        if (!meta) {
            return false;
        }
        const text = String(meta.text || '').trim();
        if (!text || text === 'N/A') {
            return false;
        }
        return this.parseNums(text).indexOf(num) !== -1;
    }

    isNonexistNumYellowAtRow(rowIndex, num) {
        const rows = this.getSourceSheetRows();
        if (rowIndex < 0 || rowIndex >= rows.length) {
            return false;
        }
        if (!this.numInNonexistCacheRow(rowIndex, num)) {
            return false;
        }
        const row = rows[rowIndex];
        const nonexistMeta = this.getNonexistMetaForSourceRow(rowIndex, row);
        const nx = String(nonexistMeta.text || '').trim();
        if (!nx || nx === 'N/A') {
            return false;
        }
        const res = row.result || row.Result || '';
        return this.getNonexistDisplayKindForNumber(rowIndex, num, nx, res) === 'yellow';
    }

    /** Mọi số trong cột nonexist hàng focus (kể cả đã trúng — vd 30 @731). */
    getFocusRowNonexistNums(focusRowIndex) {
        const rows = this.getSourceSheetRows();
        if (focusRowIndex < 0 || focusRowIndex >= rows.length) {
            return new Set();
        }
        if (!this.nonexistCache || this.nonexistCache.length !== rows.length) {
            this.refreshDerivedState();
        }
        const meta = this.nonexistCache[focusRowIndex];
        const text = meta ? String(meta.text || '').trim() : '';
        if (!text || text === 'N/A') {
            return new Set();
        }
        return new Set(this.parseNums(text));
    }

    /** Số trên nonexist hàng focus chưa trúng 5 số chính (đỏ / vàng / tím…). */
    getFocusRowUncalledNonexistNums(focusRowIndex) {
        const rows = this.getSourceSheetRows();
        if (focusRowIndex < 0 || focusRowIndex >= rows.length) {
            return new Set();
        }
        if (!this.nonexistCache || this.nonexistCache.length !== rows.length) {
            this.refreshDerivedState();
        }
        const row = rows[focusRowIndex];
        const nonexistMeta = this.getNonexistMetaForSourceRow(focusRowIndex, row);
        const nx = String(nonexistMeta.text || '').trim();
        if (!nx || nx === 'N/A') {
            return new Set();
        }
        const res = row.result || row.Result || '';
        const state = this.computeNonexistVisualState(focusRowIndex, nx, res);
        const out = new Set();
        const candidates = this.parseNums(nx);
        for (let i = 0; i < candidates.length; i++) {
            const num = candidates[i];
            if (!state.currentNums.has(num)) {
                out.add(num);
            }
        }
        return out;
    }

    /**
     * Số nonexist tím/đỏ trên hàng focus (rỗng/mask → tính như chưa result; có result → kind thật).
     * Dùng để tìm “vàng gần nhất” ngoài chuỗi 10 và gắn x1.5.
     */
    getFocusRowPurpleRedNonexistTrailNums(focusRowIndex) {
        const rows = this.getSourceSheetRows();
        if (focusRowIndex < 0 || focusRowIndex >= rows.length) {
            return new Set();
        }
        if (!this.nonexistCache || this.nonexistCache.length !== rows.length) {
            this.refreshDerivedState();
        }
        const row = rows[focusRowIndex];
        if (!row) {
            return new Set();
        }
        const rawRes = String(row.result || row.Result || '').trim();
        const treatAsEmpty = !rawRes
            || this.isEmptyResultRow(row)
            || this.shouldAnswerPopupMaskSheet1Row(focusRowIndex);
        const rowForMeta = treatAsEmpty
            ? Object.assign({}, row, { result: '', Result: '' })
            : row;
        const resForKind = treatAsEmpty ? '' : rawRes;
        const nonexistMeta = this.getNonexistMetaForSourceRow(focusRowIndex, rowForMeta);
        const nx = String(nonexistMeta.text || '').trim();
        if (!nx || nx === 'N/A') {
            return new Set();
        }
        const state = this.computeNonexistVisualState(focusRowIndex, nx, resForKind);
        const out = new Set();
        const candidates = this.parseNums(nx);
        for (let i = 0; i < candidates.length; i++) {
            const num = candidates[i];
            const kind = this.getNonexistDisplayKindForNumber(
                focusRowIndex,
                num,
                nx,
                resForKind,
                state
            );
            if (kind === 'purple' || kind === 'red') {
                out.add(num);
            }
        }
        return out;
    }

    /**
     * Hàng ngoài cửa sổ 10 (idx < windowStart) gần cửa nhất mà `num` đang vàng trong nonexist.
     * VD: focus 761 đỏ 18 → 729 (vàng gần nhất của 18 ngoài cửa).
     */
    findNearestOutsideYellowNonexistRow(num, windowStart) {
        const start = Math.floor(Number(windowStart));
        if (!Number.isFinite(start) || start <= 0) {
            return -1;
        }
        const rows = this.getSourceSheetRows();
        if (!this.nonexistCache || this.nonexistCache.length !== rows.length) {
            this.refreshDerivedState();
        }
        for (let i = start - 1; i >= 0; i--) {
            if (this.isNonexistNumYellowAtRow(i, num)) {
                return i;
            }
        }
        return -1;
    }

    _getFocusNonexistTrailNumsCache() {
        const win = this.activeWindowRange;
        if (!win || typeof win.end !== 'number') {
            return new Set();
        }
        const targetIdx = typeof win.target === 'number' ? win.target : win.end;
        const rows = this.getSourceSheetRows();
        const row = rows[targetIdx];
        const rawRes = row ? String(row.result || row.Result || '').trim() : '';
        const treatAsEmpty = !rawRes
            || (row && this.isEmptyResultRow(row))
            || this.shouldAnswerPopupMaskSheet1Row(targetIdx);
        const key = `${targetIdx}|${win.end}|${win.start}|pr|${treatAsEmpty ? 1 : 0}`;
        if (this._focusNonexistTrailKey === key && this._focusNonexistTrailSet) {
            return this._focusNonexistTrailSet;
        }
        this._focusNonexistTrailKey = key;
        this._focusNonexistTrailSet = this.getFocusRowPurpleRedNonexistTrailNums(targetIdx);
        return this._focusNonexistTrailSet;
    }

    /**
     * Boost vàng x1.5 ngoài cửa 10: số tím/đỏ trên focus → đúng hàng
     * “vàng gần nhất” của số đó phía trên cửa sổ (vd 761 đỏ 18 → 729 vàng 18).
     */
    isOutsideWindowFocusTrailBoost(rowIndex, num, winOverride = null) {
        const win = winOverride || this.activeWindowRange;
        if (!win || typeof win.start !== 'number' || typeof win.end !== 'number') {
            return false;
        }
        const start = win.start;
        if (rowIndex >= start) {
            return false;
        }
        const targetIdx = typeof win.target === 'number' ? win.target : win.end;
        const trailSet = winOverride
            ? this.getFocusRowPurpleRedNonexistTrailNums(targetIdx)
            : this._getFocusNonexistTrailNumsCache();
        if (!trailSet.has(num)) {
            return false;
        }
        const nearestYellow = this.findNearestOutsideYellowNonexistRow(num, start);
        return nearestYellow === rowIndex;
    }

    /** Các hàng cần refresh nonexist khi đổi cửa sổ / focus (gồm vàng gần nhất ngoài cửa 10). */
    collectNonexistBoostRefreshRowIndices(win) {
        const indices = new Set();
        if (!win || typeof win.start !== 'number' || typeof win.end !== 'number') {
            return indices;
        }
        const start = win.start;
        const end = win.end;
        const targetIdx = typeof win.target === 'number' ? win.target : end;
        for (let i = start; i <= end; i++) {
            indices.add(i);
        }
        if (targetIdx >= 0) {
            indices.add(targetIdx);
        }
        if (!this.nonexistCache || this.nonexistCache.length !== this.getSourceSheetRows().length) {
            this.refreshDerivedState();
        }
        const trailNumsOnFocus = this.getFocusRowPurpleRedNonexistTrailNums(targetIdx);
        for (const num of trailNumsOnFocus) {
            const nearestYellow = this.findNearestOutsideYellowNonexistRow(num, start);
            if (nearestYellow >= 0) {
                indices.add(nearestYellow);
            }
        }
        return indices;
    }

    shouldBoostYellowNonexistForWindow(rowIndex, num, winOverride = null) {
        const win = winOverride || this.activeWindowRange;
        if (!win || typeof win.start !== 'number' || typeof win.end !== 'number') {
            return false;
        }
        const start = win.start;
        const end = win.end;
        if (end < start) {
            return false;
        }
        if (this.isOutsideWindowFocusTrailBoost(rowIndex, num, win)) {
            return true;
        }

        if (!this.numInNonexistCacheRow(end, num)) {
            return false;
        }
        const maxLabels = Math.min(10, Math.max(0, end - start + 1));
        const lastLabeledRow = start + maxLabels - 1;
        if (rowIndex < start || rowIndex > lastLabeledRow) {
            return false;
        }
        return true;
    }

    /**
     * Re-render nonexist cells for specific row indices (yellow x1.5 tracks active window).
     * @param {object} [options]
     * @param {boolean} [options.forFilterPopup] — #filterTableWrap + getSourceSheetRows()
     */
    refreshNonexistCellsForRowIndices(tableWrap, rowIndices, options = {}) {
        if (!tableWrap) {
            return;
        }
        const forFilterPopup = options.forFilterPopup === true
            || tableWrap.id === 'filterTableWrap';
        if (!forFilterPopup && this.activeSheet !== 'sheet1') {
            return;
        }
        const meta = forFilterPopup ? null : (this.sheets[this.activeSheet] || {});
        if (!forFilterPopup && meta && meta.kind === 'combo') {
            return;
        }
        const displayRows = forFilterPopup
            ? (this.getSourceSheetRows() || [])
            : (this.dataRows || []);
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
            cell.innerHTML = this.renderSourceRowNonexistCellHtml(i, row, {
                windowRange: options.windowRange || null
            });
            const masked = this.shouldAnswerPopupMaskSheet1Row(i);
            tr.classList.toggle('answer-popup-focus-masked', masked);
            cell.classList.toggle('answer-popup-focus-nonexist', masked);
        }
    }

    /**
     * Re-render nonexist column for the active sliding window (and previous window when it moves).
     */
    refreshNonexistCellsForActiveWindow(tableWrap) {
        const win = this.activeWindowRange;
        this.refreshNonexistCellsForRowIndices(tableWrap, this.collectNonexistBoostRefreshRowIndices(win));
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
    renderNonexistHtml(rowIndex, nonexistText, currentResult, options = {}) {
        if (!nonexistText || nonexistText === 'N/A') {
            return this.escapeHtml(nonexistText || '');
        }

        const windowRange = options.windowRange || null;
        const greenStrikeStyle = 'color:rgb(0,80,0);font-weight:bold;font-size:1.5em;text-decoration:line-through';
        const redLongestStyle = 'color:rgb(255,0,0);font-weight:bold';
        const purpleOutsideStyle = 'color:rgb(148,55,220);font-weight:bold';

        const trailYellowBoostStyle = 'color:rgb(240,200,64);font-weight:bold;font-size:1.5em';

        return this.escapeHtml(nonexistText).replace(/\b\d+\b/g, (match) => {
            const value = parseInt(match, 10);
            const displayKind = this.getNonexistDisplayKindForNumber(rowIndex, value, nonexistText, currentResult);

            /* Focus tím/đỏ (chưa result) → ngoài chuỗi 10: ép vàng x1.5 (kể cả hàng đang tím/đỏ) */
            const isGreenKind = (
                displayKind === 'green'
                || displayKind === 'green-ul'
                || displayKind === 'green-italic'
                || displayKind === 'green-strike'
            );
            if (!isGreenKind && this.isOutsideWindowFocusTrailBoost(rowIndex, value, windowRange)) {
                return `<span style="${trailYellowBoostStyle}">${value}</span>`;
            }

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
                const boost = this.shouldBoostYellowNonexistForWindow(rowIndex, value, windowRange);
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
     * Kỳ tại rowIndex có ≥1 số chính trùng kỳ cách đúng `offset` dòng (chuỗi offset trong cửa sổ 10).
     * offset 1 = chuỗi liền trên (kỳ rowIndex - 1).
     */
    recallsAtLeastOneFromPrevPeriodAtOffset(rows, rowIndex, offset) {
        const list = rows || [];
        const idx = Number(rowIndex);
        const off = Number(offset);
        if (!Number.isFinite(idx) || !Number.isFinite(off) || off < 1 || off > 10) {
            return false;
        }
        const prevIdx = idx - off;
        if (prevIdx < 0 || prevIdx >= list.length) {
            return false;
        }
        const curRow = list[idx];
        const prevRow = list[prevIdx];
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
     * Kỳ tại rowIndex có ≥1 số chính (5 số trước |) trùng kỳ liền trước (cùng thứ tự data.json).
     */
    recallsAtLeastOneFromImmediatePrevPeriod(rows, rowIndex) {
        return this.recallsAtLeastOneFromPrevPeriodAtOffset(rows, rowIndex, 1);
    }

    /**
     * Có thể so sánh gọi lại với kỳ cách đúng `offset` dòng (cả hai có đủ 5 số chính).
     */
    isEligibleForPrevPeriodRecallComparisonAtOffset(rows, rowIndex, offset) {
        const list = rows || [];
        const idx = Number(rowIndex);
        const off = Number(offset);
        if (!Number.isFinite(idx) || !Number.isFinite(off) || off < 1 || off > 10) {
            return false;
        }
        const prevIdx = idx - off;
        if (prevIdx < 0 || prevIdx >= list.length) {
            return false;
        }
        const curRow = list[idx];
        const prevRow = list[prevIdx];
        if (this.isEmptyResultRow(curRow) || this.isEmptyResultRow(prevRow)) {
            return false;
        }
        const cur = this.parseMainNums(curRow.result || curRow.Result);
        const prev = this.parseMainNums(prevRow.result || prevRow.Result);
        return cur.length === 5 && prev.length === 5;
    }

    /**
     * Có thể so sánh gọi lại với kỳ liền trước (cả hai có đủ 5 số chính).
     */
    isEligibleForPrevPeriodRecallComparison(rows, rowIndex) {
        return this.isEligibleForPrevPeriodRecallComparisonAtOffset(rows, rowIndex, 1);
    }

    /**
     * % kỳ có corner fold trong mẫu rowIndices (bảng đầy đủ hoặc tập lọc Ctrl popup).
     */
    computePrevPeriodRecallFoldStats(rows, rowIndices, chainOffset = 1) {
        const off = Math.max(1, Math.min(10, Number(chainOffset) || 1));
        const list = rows || [];
        const indices = Array.isArray(rowIndices)
            ? rowIndices.filter(i => i >= 0 && i < list.length)
            : list.map((_, i) => i);
        let eligible = 0;
        let withRecall = 0;
        for (let k = 0; k < indices.length; k++) {
            const i = indices[k];
            if (!this.isEligibleForPrevPeriodRecallComparisonAtOffset(list, i, off)) {
                continue;
            }
            eligible++;
            if (this.recallsAtLeastOneFromPrevPeriodAtOffset(list, i, off)) {
                withRecall++;
            }
        }
        return {
            chain: off,
            eligible,
            withRecall,
            pct: eligible > 0 ? (withRecall / eligible) * 100 : null
        };
    }

    /**
     * % corner fold theo từng chuỗi 1…10 trong cửa sổ trượt (so với mẫu rowIndices).
     * @returns {{ chain: number, eligible: number, withRecall: number, pct: number | null }[]}
     */
    computePrevPeriodRecallFoldStatsByChain(rows, rowIndices) {
        const out = [];
        for (let chain = 1; chain <= 10; chain++) {
            out.push(this.computePrevPeriodRecallFoldStats(rows, rowIndices, chain));
        }
        return out;
    }

    formatPrevPeriodRecallFoldPct(stats) {
        if (!stats || !stats.eligible) {
            return '—';
        }
        const rounded = Math.round(stats.pct * 10) / 10;
        return rounded.toFixed(1) + '%';
    }

    formatPrevPeriodRecallFoldPctByChain(statsByChain) {
        const lines = [];
        for (let chain = 1; chain <= 10; chain++) {
            const stats = statsByChain && statsByChain[chain - 1]
                ? statsByChain[chain - 1]
                : null;
            lines.push(`${chain}: ${this.formatPrevPeriodRecallFoldPct(stats)}`);
        }
        return lines.join('\n');
    }

    encodePrevPeriodRecallFoldTooltipAttr(text) {
        return this.escapeHtml(String(text || '')).replace(/\n/g, '&#10;');
    }

    /**
     * Cập nhật tooltip % trên góc fold (sau lọc popup khi tái dùng HTML cache).
     */
    applyPrevPeriodRecallFoldTooltips(tableWrap, rows, rowIndices) {
        if (!tableWrap) {
            return null;
        }
        const statsByChain = this.computePrevPeriodRecallFoldStatsByChain(rows, rowIndices);
        const pctLabel = this.formatPrevPeriodRecallFoldPctByChain(statsByChain);
        const pctAttr = this.encodePrevPeriodRecallFoldTooltipAttr(pctLabel);
        tableWrap.querySelectorAll('td.cell-result.has-prev-period-recall').forEach(function (cell) {
            let hit = cell.querySelector('.prev-period-recall-fold');
            if (!hit) {
                hit = document.createElement('span');
                hit.className = 'prev-period-recall-fold';
                cell.insertBefore(hit, cell.firstChild);
            }
            hit.setAttribute('data-pct', pctAttr);
            hit.removeAttribute('title');
        });
        return statsByChain;
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
     * @param {*} rawId
     * @param {Map<string, number>} [freqMap] - defaults to full-sheet idFrequencyMap
     */
    getIdBackgroundByFrequency(rawId, freqMap) {
        const idNum = this.parseRowId(rawId);
        const map = freqMap || this.idFrequencyMap;
        if (idNum === null || !map) {
            return '';
        }

        const freq = map.get(String(idNum)) || 0;
        if (freq <= 0) {
            return '';
        }
        const idx = Math.min(freq, ID_REF_COUNT_BG_COLORS.length) - 1;
        return ID_REF_COUNT_BG_COLORS[idx];
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
     * Compute tail candidates for a 10-row window (same heuristic as left pane tail list).
     * @param {object[]} rows
     * @returns {{a:number,b:number}[]}
     */
    computeTailsForRows(rows) {
        if (!rows || rows.length < 10) {
            return [];
        }
        const display = rows.slice(0, 10);
        const sets = display.map(row => new Set(this.parseMainNums(row.result || row.Result || '')));
        const adjMap = {};
        for (let i = 0; i < sets.length - 1; i++) {
            const top = sets[i];
            const bottom = sets[i + 1];
            top.forEach((a) => {
                bottom.forEach((b) => {
                    if (a === b) {
                        return;
                    }
                    const key = a + ',' + b;
                    if (!adjMap[key]) {
                        adjMap[key] = [];
                    }
                    adjMap[key].push(i);
                });
            });
        }
        const tails = [];
        for (const key in adjMap) {
            const idxs = adjMap[key].slice().sort((a, b) => a - b);
            const parts = key.split(',').map((x) => parseInt(x, 10));
            const a = parts[0];
            const b = parts[1];
            let last = -3;
            let count = 0;
            for (const ii of idxs) {
                if (ii > last + 1) {
                    count++;
                    last = ii;
                }
            }
            if (count >= 2) {
                const mainIdx = idxs.length ? Math.min(...idxs) : Infinity;
                tails.push({ a, b, count, mainIdx });
            }
        }
        tails.sort((x, y) => {
            if (x.mainIdx !== y.mainIdx) {
                return x.mainIdx - y.mainIdx;
            }
            if (x.a !== y.a) {
                return x.a - y.a;
            }
            return x.b - y.b;
        });
        return tails;
    }

    /**
     * Row matches tail filter when answer contains the first tail pair (strip order, same as left pane)
     * from the previous 10-row window; [tail] count applies to that pair only.
     */
    shouldHighlightDateByTailWindow(rows, rowIndex, filterOptions = null) {
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

        const visibleTails = this.computeTailsForRows(windowRows);
        if (!visibleTails || !visibleTails.length) {
            return false;
        }

        let th = 2;
        let op = '>=';
        if (filterOptions) {
            const parsed = parseInt(filterOptions.tailMinCount, 10);
            if (Number.isFinite(parsed)) {
                th = Math.min(5, Math.max(2, parsed));
            }
            const rawOp = String(filterOptions.tailCountOp || '').trim();
            if (rawOp === '=' || rawOp === '<=' || rawOp === '>=') {
                op = rawOp;
            }
        }

        for (let t = 0; t < visibleTails.length; t++) {
            const tail = visibleTails[t];
            if (!this.pairExists(currentNums, tail.a, tail.b)) {
                continue;
            }
            const tailCount = Number.isFinite(tail.count) ? tail.count : 2;
            return this.freqMatchesComparison(tailCount, th, op);
        }

        return false;
    }

    /**
     * Cached row indices whose answer contains a tail pair from the 10-row window before that row.
     */
    ensureTailFilterIndicesCache() {
        const rows = this.getSourceSheetRows();
        if (this.tailFilterIndicesCache && this.tailFilterIndicesCacheRowLen === rows.length) {
            return this.tailFilterIndicesCache;
        }

        const indices = [];
        for (let i = 0; i < rows.length; i++) {
            if (!this.isEmptyResultRow(rows[i]) && this.shouldHighlightDateByTailWindow(rows, i)) {
                indices.push(i);
            }
        }
        this.tailFilterIndicesCache = indices;
        this.tailFilterIndicesCacheRowLen = rows.length;
        return indices;
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
     * Trích prevId được tham chiếu trong note (mỗi part dạng current-prev=...).
     * @returns {number[]}
     */
    static extractReferencedPrevIdsFromNoteText(noteText) {
        const ids = [];
        const txt = String(noteText || '').trim();
        if (!txt || txt === '?') {
            return ids;
        }
        const parts = txt.split(' ');
        for (let p = 0; p < parts.length; p++) {
            const part = parts[p];
            if (part.indexOf('-') < 0) {
                continue;
            }
            const leftPart = part.split('=')[0];
            const idPrevRaw = String(leftPart.split('-')[1] || '').trim();
            const idPrevDigits = String(idPrevRaw).replace(/\D/g, '');
            const idPrevNum = parseInt(idPrevDigits, 10);
            if (Number.isFinite(idPrevNum)) {
                ids.push(idPrevNum);
            }
        }
        return ids;
    }

    /**
     * Các chỉ số hàng sheet1 có note tham chiếu tới prevId = targetIdNum.
     * @returns {number[]}
     */
    findRowIndicesReferencingIdInNotes(targetIdNum) {
        if (!Number.isFinite(targetIdNum)) {
            return [];
        }
        const rows = this.getSourceSheetRows();
        const indices = [];
        for (let i = 0; i < rows.length; i++) {
            const row = rows[i];
            if (!row || this.isEmptyResultRow(row)) {
                continue;
            }
            const noteMeta = this.getComputedNoteMeta(i, row);
            const prevIds = RightPaneSheetManager.extractReferencedPrevIdsFromNoteText(noteMeta.text);
            if (prevIds.includes(targetIdNum)) {
                indices.push(i);
            }
        }
        return indices;
    }

    /**
     * Các hàng trong cửa sổ 10 của focusRowIndex có id được note của hàng focus tham chiếu tới.
     * @param {number} focusRowIndex
     * @returns {number[]}
     */
    findWindowRowIndicesReferencedInFocusNote(focusRowIndex) {
        const rows = this.getSourceSheetRows();
        if (!Number.isFinite(focusRowIndex) || focusRowIndex < 0 || focusRowIndex >= rows.length) {
            return [];
        }
        const row = rows[focusRowIndex];
        if (!row || this.isEmptyResultRow(row)) {
            return [];
        }
        const windowTop = focusRowIndex >= 10 ? focusRowIndex - 10 : 0;
        const windowEnd = focusRowIndex;
        const noteMeta = this.getComputedNoteMeta(focusRowIndex, row);
        const prevIds = RightPaneSheetManager.extractReferencedPrevIdsFromNoteText(noteMeta.text);
        if (!prevIds.length) {
            return [];
        }
        const seen = new Set();
        const indices = [];
        for (let p = 0; p < prevIds.length; p++) {
            const matchIdx = this.findSourceSheetRowIndexByNumericId(prevIds[p]);
            if (matchIdx < 0 || matchIdx < windowTop || matchIdx > windowEnd) {
                continue;
            }
            if (seen.has(matchIdx)) {
                continue;
            }
            seen.add(matchIdx);
            indices.push(matchIdx);
        }
        indices.sort((a, b) => a - b);
        return indices;
    }

    /**
     * Focus hàng có id = id hàng gốc + delta (vd 00014 → 00024 khi delta=10).
     * @param {object} [clickOptions] — truyền thêm vào onRowClick (vd idRefHighlightIndices)
     * @returns {boolean} true nếu đã xử lý (đã preventDefault)
     */
    focusSourceSheetRowByIdDelta(sourceIdx, idDelta, event, tableWrap, clickOptions) {
        const rows = this.getSourceSheetRows();
        if (!Number.isFinite(sourceIdx) || sourceIdx < 0 || sourceIdx >= rows.length) {
            return false;
        }
        const row = rows[sourceIdx];
        if (!row) {
            return false;
        }
        const currentIdNum = this.parseRowId(row.id || row.ID || '');
        if (currentIdNum === null) {
            return false;
        }
        const targetIdx = this.resolveContextMenuTargetRowIndex(sourceIdx, idDelta);
        const highlightIndices = clickOptions && Array.isArray(clickOptions.idRefHighlightIndices)
            ? clickOptions.idRefHighlightIndices.slice()
            : null;

        const applyHighlightsOnly = () => {
            if (!highlightIndices || !highlightIndices.length || !tableWrap) {
                return;
            }
            const endIdx = targetIdx >= 0 ? targetIdx : sourceIdx;
            const startIdx = endIdx >= 10 ? endIdx - 10 : 0;
            this.applyWindowSelection(startIdx, endIdx, endIdx, tableWrap, {
                idRefHighlightIndices: highlightIndices
            });
            try {
                tableWrap.focus({ preventScroll: true });
            } catch (err) {
                /* ignore */
            }
        };

        if (targetIdx < 0 || targetIdx === sourceIdx) {
            applyHighlightsOnly();
            return true;
        }
        const targetRow = rows[targetIdx];
        const targetEmpty = this.isEmptyResultRow(targetRow);
        this.onRowClick(targetIdx, targetEmpty, event, {
            fromFilterNav: tableWrap && tableWrap.id === 'filterTableWrap',
            ...(clickOptions || {}),
            idRefHighlightIndices: highlightIndices
        });
        try {
            if (tableWrap) {
                tableWrap.focus({ preventScroll: true });
            }
        } catch (err) {
            /* ignore */
        }
        return true;
    }

    /**
     * Chuột phải trên ô id / nonexist sheet1.
     * @returns {boolean} true nếu đã xử lý (đã preventDefault)
     */
    handleSourceSheetCellContextMenu(event, tableWrap) {
        if (!event || this.activeSheet !== 'sheet1') {
            return false;
        }
        const idTd = event.target && event.target.closest && event.target.closest('td.cell-id');
        if (idTd && tableWrap && tableWrap.contains(idTd)) {
            return this.handleIdCellContextMenu(event, tableWrap, idTd);
        }
        return this.handleNonexistCellContextMenu(event, tableWrap);
    }

    /**
     * Chỉ số focus sau chuột phải id/nonexist (+delta id). Nếu id đích vượt cuối bảng → hàng cuối.
     * @returns {number}
     */
    resolveContextMenuTargetRowIndex(sourceIdx, idDelta) {
        const rows = this.getSourceSheetRows();
        if (!Number.isFinite(sourceIdx) || sourceIdx < 0 || sourceIdx >= rows.length) {
            return -1;
        }
        const row = rows[sourceIdx];
        const currentIdNum = this.parseRowId(row.id || row.ID || '');
        if (currentIdNum === null) {
            return -1;
        }
        const targetIdNum = currentIdNum + (Number(idDelta) || 0);
        let targetIdx = this.findSourceSheetRowIndexByNumericId(targetIdNum);
        if (targetIdx < 0 && rows.length > 0) {
            const lastIdx = rows.length - 1;
            const lastIdNum = this.parseRowId(rows[lastIdx].id || rows[lastIdx].ID || '');
            if (lastIdNum !== null && targetIdNum > lastIdNum) {
                targetIdx = lastIdx;
            }
        }
        if (targetIdx < 0) {
            return sourceIdx;
        }
        return targetIdx;
    }

    /**
     * Viền cyan chuột phải ô id: các kỳ trong cửa sổ 10 có note tham chiếu id vừa click (không tô hàng click).
     * @returns {number[]}
     */
    buildIdRefContextmenuHighlightIndices(clickedIdx, clickedIdNum, focusRowIdx) {
        if (!Number.isFinite(clickedIdNum)) {
            return [];
        }
        const windowEnd = Number.isFinite(focusRowIdx) && focusRowIdx >= 0 ? focusRowIdx : clickedIdx;
        const windowTop = windowEnd >= 10 ? windowEnd - 10 : 0;
        const indices = [];

        const allRefs = this.findRowIndicesReferencingIdInNotes(clickedIdNum);
        for (let r = 0; r < allRefs.length; r++) {
            const rowIdx = allRefs[r];
            if (rowIdx >= windowTop && rowIdx <= windowEnd && rowIdx !== clickedIdx) {
                indices.push(rowIdx);
            }
        }

        indices.sort((a, b) => a - b);
        return indices;
    }

    /**
     * Viền cyan chuột phải ô id: mọi kỳ có note tham chiếu id vừa click (không giới hạn cửa sổ 10).
     * @returns {number[]}
     */
    buildAllIdRefContextmenuHighlightIndices(clickedIdx, clickedIdNum) {
        if (!Number.isFinite(clickedIdNum)) {
            return [];
        }
        const indices = this.findRowIndicesReferencingIdInNotes(clickedIdNum)
            .filter((rowIdx) => rowIdx !== clickedIdx);
        indices.sort((a, b) => a - b);
        return indices;
    }

    /**
     * Hàng sheet1 có nằm trong cửa sổ trượt 10 đang hiển thị không.
     * @param {number} rowIdx
     * @returns {boolean}
     */
    isSourceSheetRowInActiveWindow(rowIdx) {
        const win = this.activeWindowRange;
        if (!win || typeof win.start !== 'number' || typeof win.end !== 'number') {
            return false;
        }
        if (!Number.isFinite(rowIdx) || rowIdx < 0) {
            return false;
        }
        return rowIdx >= win.start && rowIdx <= win.end;
    }

    /**
     * Chuột phải trên ô id: focus id+delta + viền cyan n ô note + result tham chiếu id vừa click.
     * @returns {boolean}
     */
    handleIdCellContextMenu(event, tableWrap, td) {
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
        if (!row) {
            return false;
        }
        event.preventDefault();
        event.stopPropagation();
        const clickedIdNum = this.parseRowId(row.id || row.ID || '');
        if (clickedIdNum === null) {
            return true;
        }
        const win = this.activeWindowRange;
        if (this.isSourceSheetRowInActiveWindow(idx)) {
            const windowTarget = typeof win.target === 'number' ? win.target : win.end;
            const refRowIndices = this.buildAllIdRefContextmenuHighlightIndices(idx, clickedIdNum);
            this.applyWindowSelection(win.start, win.end, windowTarget, tableWrap, {
                idRefHighlightIndices: refRowIndices,
                focusNoteRefHighlightIndices: Array.isArray(win.focusNoteRefHighlightIndices)
                    ? win.focusNoteRefHighlightIndices.slice()
                    : null
            });
            try {
                tableWrap.focus({ preventScroll: true });
            } catch (err) {
                /* ignore */
            }
            return true;
        }
        const focusIdx = this.resolveContextMenuTargetRowIndex(idx, NONEXIST_CONTEXTMENU_ID_DELTA);
        const refRowIndices = this.buildIdRefContextmenuHighlightIndices(idx, clickedIdNum, focusIdx);
        this.focusSourceSheetRowByIdDelta(idx, NONEXIST_CONTEXTMENU_ID_DELTA, event, tableWrap, {
            idRefHighlightIndices: refRowIndices
        });
        return true;
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
        this.focusSourceSheetRowByIdDelta(idx, NONEXIST_CONTEXTMENU_ID_DELTA, event, tableWrap);
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

        if (this.activeSheet === 'sheet1' || options.asSheet1) {
            const focusRow = this.dataRows[idx] || rowAtClick;
            const nextFocusId = String(focusRow.id || focusRow.ID || clickedRowId || '').trim();
            const prevFocusId = String(this.comboFocusRowId || '').trim();
            const hadG1 = this.comboG1Enabled;
            const comboStateChanged = this.comboFocusRowId !== nextFocusId
                || this.comboFocusRowIndex !== idx
                || (this.isEmptyResultRow(focusRow) && hadG1);
            this.onComboFocusIdChanged(prevFocusId, nextFocusId);
            this.comboFocusRowId = nextFocusId;
            this.comboFocusRowIndex = idx;
            if (this.isEmptyResultRow(focusRow)) {
                this.comboG1Enabled = false;
            }
            if (comboStateChanged) {
                window.dispatchEvent(new CustomEvent('comboControlsChanged', {
                    detail: { sheet: options.asSheet1 ? 'sheet1' : this.activeSheet }
                }));
            }
        }

        if (!options.fromTrackingSync && !options.fromSheet1NavStep) {
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
        if (!options.skipMainWindowSelection) {
            this.applyWindowSelection(windowTop, windowEnd, targetIdx, null, {
                idRefHighlightIndices: options.idRefHighlightIndices || null
            });
        } else if (options.asSheet1) {
            this.commitSourceSheetWindowRangeForIndex(idx, {
                idRefHighlightIndices: options.idRefHighlightIndices || null
            });
        }

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
                sheetName: options.asSheet1 ? 'sheet1' : this.activeSheet,
                clickedRowId,
                focusRowIndex: idx,
                focusNonexistHighlights,
                fromFilterNav: !!options.fromFilterNav,
                fromTrackingSync: !!options.fromTrackingSync,
                fromSheet1NavStep: !!options.fromSheet1NavStep,
                trackingFrameStep: !!options.trackingFrameStep,
                light: !!options.light,
                contextPrefixCount
            }
        }));
    }

    /**
     * Filter popup ALL: đồng bộ focus + cửa sổ 10 trên sheet chính, không setLines/iframe.
     */
    syncComboFocusFromSourceRowIndex(idx, options = {}) {
        const rows = this.getSourceSheetRows();
        if (typeof idx !== 'number' || idx < 0 || idx >= rows.length) {
            return false;
        }
        const row = rows[idx];
        const useComboFocus = this.activeSheet === 'sheet1'
            || this.activeSheet === 'tracking'
            || options.asSheet1;
        if (!useComboFocus) {
            return false;
        }
        const nextFocusId = String(row.id || row.ID || '').trim();
        const prevFocusId = String(this.comboFocusRowId || '').trim();
        const hadG1 = this.comboG1Enabled;
        const comboStateChanged = this.comboFocusRowId !== nextFocusId
            || this.comboFocusRowIndex !== idx
            || (this.isEmptyResultRow(row) && hadG1);
        this.onComboFocusIdChanged(prevFocusId, nextFocusId);
        this.comboFocusRowId = nextFocusId;
        this.comboFocusRowIndex = idx;
        if (this.isEmptyResultRow(row)) {
            this.comboG1Enabled = false;
        }
        if (comboStateChanged) {
            window.dispatchEvent(new CustomEvent('comboControlsChanged', {
                detail: { sheet: this.activeSheet === 'tracking' ? 'tracking' : 'sheet1' }
            }));
        }
        return true;
    }

    syncSourceSheetFocusFromFilter(idx, options = {}) {
        const rows = this.getSourceSheetRows();
        if (typeof idx !== 'number' || idx < 0 || idx >= rows.length) {
            return false;
        }
        this.syncComboFocusFromSourceRowIndex(idx, options);
        if (!options.fromTrackingSync) {
            try {
                this.syncSpecialTrackingTimelineFromSheet1Row(idx);
            } catch (eStSync) {
                /* ignore */
            }
        }
        if (this.activeSheet === 'tracking') {
            this.commitSourceSheetWindowRangeForIndex(idx, {
                preserveIdRefHighlights: options.previewOnly !== false
            });
            return true;
        }
        const tableWrap = options.tableWrapEl || document.getElementById('tableWrap');
        const activeSheetMeta = this.sheets[this.activeSheet] || {};
        if (!tableWrap || activeSheetMeta.kind === 'combo') {
            return false;
        }
        const start = Math.max(0, idx - 10);
        const previewOnly = options.previewOnly !== false;
        this.applyWindowSelection(start, idx, idx, tableWrap, { previewOnly });
        if (!options.skipCenter) {
            this.centerActiveWindowInView(tableWrap);
        }
        return true;
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
        if (sheetName === TRACKING_SHEET_ID) {
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
        const active = sheetName;
        requestAnimationFrame(() => {
            if (this.activeSheet === active) {
                this.save();
            }
        });
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
        const sheetNames = [
            'sheet1',
            TRACKING_SHEET_ID,
            'combo_1',
            'combo_2',
            'combo_3',
            'combo_4',
            'combo_5'
        ];
        const visibleNames = sheetNames.filter(name => this.sheets[name]);
        let tabBar = container.querySelector('.sheet-tabs-bar');
        const zoomToolbar = container.querySelector('.zoom-toolbar');
        if (tabBar) {
            const existingTabs = tabBar.querySelectorAll('.sheet-tab[data-sheet-name]');
            if (existingTabs.length === visibleNames.length
                && visibleNames.every((name, i) => existingTabs[i].dataset.sheetName === name)) {
                existingTabs.forEach(tab => {
                    tab.classList.toggle('active', tab.dataset.sheetName === this.activeSheet);
                });
                return;
            }
        }

        if (!tabBar) {
            tabBar = document.createElement('div');
            tabBar.className = 'sheet-tabs-bar';
            if (zoomToolbar) {
                container.insertBefore(tabBar, zoomToolbar);
            } else {
                container.appendChild(tabBar);
            }
        } else {
            tabBar.innerHTML = '';
        }

        for (const name of visibleNames) {
            const tab = document.createElement('button');
            tab.className = 'sheet-tab';
            tab.dataset.sheetName = name;
            if (name === this.activeSheet) {
                tab.classList.add('active');
            }
            tab.textContent = name === TRACKING_SHEET_ID ? 'tracking' : name;
            tab.title = name === TRACKING_SHEET_ID
                ? 'Theo dõi tần suất số theo timeline (special 1–12 / basic 1–35)'
                : name;
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
                if (name !== 'sheet1' && name !== TRACKING_SHEET_ID) {
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

        if (zoomToolbar) {
            container.appendChild(zoomToolbar);
        }
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
     * Hàng focus mặc định khi mở trang: id ngay trên hàng result rỗng cuối sheet.
     * @returns {number}
     */
    getDefaultSheet1FocusRowIndex() {
        const rows = this.sourceRows || [];
        if (!rows.length) {
            return -1;
        }
        for (let i = rows.length - 1; i >= 0; i--) {
            if (this.isEmptyResultRow(rows[i])) {
                return Math.max(0, i - 1);
            }
        }
        for (let i = rows.length - 1; i >= 0; i--) {
            const result = rows[i].result || rows[i].Result || '';
            if (this.parseMainNums(result).length === 5) {
                return i;
            }
        }
        return rows.length - 1;
    }

    /**
     * Chế độ autoring toolbar theo đặc tính đáp án kỳ (nhãn C/∩/3C, dateband #00b0f0, tail, …).
     * @param {number} rowIndex
     * @param {object} [filterOptions]
     * @returns {string}
     */
    getSubmitRingModeForRowIndex(rowIndex, filterOptions = null) {
        const rows = this.getSourceSheetRows();
        if (typeof rowIndex !== 'number' || rowIndex < 0 || rowIndex >= rows.length) {
            return 'max';
        }
        const row = rows[rowIndex];
        if (this.isEmptyResultRow(row)) {
            return 'max';
        }
        const flags = this.getRowPickPropertyFlags(rows, rowIndex, row);
        if (flags.conn3) {
            return 'conn3';
        }
        if (flags.conn && flags.ix) {
            return 'conn_ix';
        }
        if (flags.conn) {
            return 'conn';
        }
        if (flags.ix) {
            return 'ix';
        }
        const opts = filterOptions || {};
        let thX = parseInt(opts.datebandMinFreqA, 10);
        let thY = parseInt(opts.datebandMinFreqB, 10);
        if (!Number.isFinite(thX)) {
            thX = 2;
        }
        if (!Number.isFinite(thY)) {
            thY = 2;
        }
        thX = Math.min(7, Math.max(2, thX));
        thY = Math.min(7, Math.max(2, thY));
        const rawOpA = String(opts.datebandFreqOpA || '').trim();
        const rawOpB = String(opts.datebandFreqOpB || '').trim();
        const opA = (rawOpA === '=' || rawOpA === '<=' || rawOpA === '>=') ? rawOpA : '>=';
        const opB = (rawOpB === '=' || rawOpB === '<=' || rawOpB === '>=') ? rawOpB : '>=';
        if (this.rowMatchesDatebandPairFreqFilter(rows, rowIndex, thX, thY, opA, opB)
            || this.rowHasCyanDateBand(rows, rowIndex)) {
            return 'dateband';
        }
        let tailTh = parseInt(opts.tailMinCount, 10);
        if (!Number.isFinite(tailTh)) {
            tailTh = 2;
        }
        tailTh = Math.min(5, Math.max(2, tailTh));
        const rawTailOp = String(opts.tailCountOp || '').trim();
        const tailOp = (rawTailOp === '=' || rawTailOp === '<=' || rawTailOp === '>=') ? rawTailOp : '>=';
        if (this.shouldHighlightDateByTailWindow(rows, rowIndex, { tailMinCount: tailTh, tailCountOp: tailOp })) {
            return 'tail';
        }
        return 'max';
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

        return { id_to_result, pair_to_ids, max_id, combo2_pair_stats: this.buildCombo2PairStatsForChecknote() };
    }

    /**
     * Map "a,b" -> { appear, groupRank, totalGroups } for checknote 📑 popup.
     * groupRank = thứ hạng nhóm theo appear tích lũy (combo_2, appear≥2, cùng appear = cùng nhóm).
     */
    buildCombo2PairStatsForChecknote() {
        const rows = this.sourceRows || [];
        const dicts = [null, new Map(), new Map(), new Map(), new Map(), new Map()];
        const dictSpecial = new Map();
        this._accumulateComboDictsFromRows(rows, dicts, dictSpecial);

        const ranked = [];
        for (const [combo, appear] of dicts[2].entries()) {
            if (appear >= 2) {
                ranked.push({ combo, appear });
            }
        }

        const appearToGroupRank = new Map();
        let totalGroups = 0;
        if (ranked.length) {
            const keysAppear2 = new Set(ranked.map((row) => row.combo));
            const comboReachRow = this.buildComboReachRowMapOnePass(rows, rows.length, 2, dicts[2], keysAppear2);
            ranked.sort((left, right) => {
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

            let groupRank = 0;
            let lastAppear = null;
            for (const row of ranked) {
                if (row.appear !== lastAppear) {
                    groupRank += 1;
                    appearToGroupRank.set(row.appear, groupRank);
                    lastAppear = row.appear;
                }
            }
            totalGroups = groupRank;
        }

        const stats = {};
        for (const [combo, appear] of dicts[2].entries()) {
            stats[combo] = {
                appear,
                groupRank: appear >= 2 ? (appearToGroupRank.get(appear) || null) : null,
                totalGroups: totalGroups > 0 ? totalGroups : null
            };
        }
        return stats;
    }

    /**
     * Chuỗi các số đặc biệt 1–12 theo thứ tự thời gian (cột result, phần sau |).
     * Timeline gồm mọi hàng sheet1 (kể cả id/result rỗng); `drawSteps[ri]` = số đặc biệt hoặc null.
     * `sourceRowIndices[k]` = k (một frame / một hàng nguồn).
     */
    buildSpecialTrackingSeriesMeta(rows) {
        const series = [];
        const drawSteps = [];
        const sourceRowIndices = [];
        const list = rows || [];
        for (let ri = 0; ri < list.length; ri++) {
            sourceRowIndices.push(ri);
            const row = list[ri];
            let step = null;
            const raw = row.result || row.Result || '';
            const part = this.parseSpecialPart(raw);
            if (part) {
                const tokens = String(part)
                    .split(/[\s,;]+/)
                    .map((t) => t.trim())
                    .filter(Boolean);
                for (const t of tokens) {
                    const v = parseInt(t, 10);
                    if (Number.isFinite(v) && v >= 1 && v <= 12) {
                        step = v;
                        series.push(v);
                        break;
                    }
                }
            }
            drawSteps.push(step);
        }
        return { series, drawSteps, sourceRowIndices };
    }

    /**
     * Chuỗi các số đặc biệt 1–12 theo thứ tự thời gian (cột result, phần sau |).
     * Tham khảo vid.py: tách special và chỉ giữ kỳ có số hợp lệ.
     */
    buildSpecialTrackingSeries(rows) {
        return this.buildSpecialTrackingSeriesMeta(rows).series;
    }

    /**
     * Sheet1 ↔ tracking: map frame timeline → chỉ số hàng nguồn sheet1.
     */
    getTrackingSourceRowIndexForFrame(sheet, frameIndex) {
        if (!sheet || typeof frameIndex !== 'number' || frameIndex < 0) {
            return -1;
        }
        this.ensureTrackingFrames(sheet);
        const viewMode = this.getTrackingViewMode(sheet);
        const srcIx = viewMode === 'basic'
            ? (sheet.basicSourceRowIndices || [])
            : (sheet.specialSourceRowIndices || []);
        if (frameIndex >= srcIx.length) {
            return -1;
        }
        return srcIx[frameIndex];
    }

    /**
     * Map chỉ số hàng sheet1 → frame timeline (mode hiện tại).
     * Ưu tiên khớp chính xác; không có thì lùi về kỳ gần nhất trước đó.
     */
    getTrackingFrameIndexForSourceRow(sheet, sourceRowIndex) {
        if (!sheet || typeof sourceRowIndex !== 'number' || sourceRowIndex < 0) {
            return -1;
        }
        this.ensureTrackingFrames(sheet);
        const viewMode = this.getTrackingViewMode(sheet);
        const srcIx = viewMode === 'basic'
            ? (sheet.basicSourceRowIndices || [])
            : (sheet.specialSourceRowIndices || []);
        if (!srcIx.length) {
            return -1;
        }
        for (let f = 0; f < srcIx.length; f++) {
            if (srcIx[f] === sourceRowIndex) {
                return f;
            }
        }
        let count = 0;
        for (let s = 0; s < srcIx.length; s++) {
            if (srcIx[s] <= sourceRowIndex) {
                count++;
            }
        }
        return Math.max(0, count - 1);
    }

    /** Id kỳ sheet1 tại frame timeline (dùng cho nhãn timeline). */
    getTrackingPeriodIdForFrame(sheet, frameIndex) {
        const rows = this.sourceRows || [];
        const rowIdx = this.getTrackingSourceRowIndexForFrame(sheet, frameIndex);
        if (rowIdx < 0 || rowIdx >= rows.length) {
            return '';
        }
        const row = rows[rowIdx];
        const raw = String(row.id != null ? row.id : row.ID || '').trim();
        const parsed = this.parseRowId(raw);
        return parsed !== null ? String(parsed) : raw;
    }

    clearSheet1LeftPanePump() {
        if (this._sheet1LeftPaneTimer) {
            clearTimeout(this._sheet1LeftPaneTimer);
            this._sheet1LeftPaneTimer = 0;
        }
    }

    resetSheet1LeftPanePump() {
        this.clearSheet1LeftPanePump();
        this._sheet1LeftPaneTarget = -1;
        this._sheet1LeftPaneCurrent = -1;
    }

    buildSourceRowLeftPaneData(idx, rows, options = {}) {
        if (!Array.isArray(rows) || typeof idx !== 'number' || idx < 0 || idx >= rows.length) {
            return null;
        }
        const windowTop = idx >= 10 ? idx - 10 : 0;
        const contextPrefixCount = idx >= 10 ? Math.min(2, windowTop) : 0;
        const dataStart = Math.max(0, windowTop - contextPrefixCount);
        const isEmptyRow = this.isEmptyResultRow(rows[idx]);
        const rowAtClick = rows[idx] || {};
        const clickedRowId = String(rowAtClick.id || rowAtClick.ID || '').trim();
        const slice = isEmptyRow ? rows.slice(dataStart, idx) : rows.slice(dataStart, idx + 1);
        const savedDataRows = this.dataRows;
        this.dataRows = rows;
        const lines = [];
        let focusNonexistHighlights;
        const selectedLines = [];
        try {
            for (let offset = 0; offset < slice.length; offset++) {
                const r = slice[offset];
                const rowIndex = dataStart + offset;
                const res = r.result || r.Result || '';
                const noteMeta = this.getComputedNoteMeta(rowIndex, r);
                const note = noteMeta.text;
                const nonexist = this.isEmptyResultRow(r) ? '' : this.getComputedNonexistMeta(rowIndex, r).text;
                lines.push([res, note, nonexist].filter(Boolean).join('\t'));
                selectedLines.push({
                    date: r.date || '',
                    id: r.id || '',
                    result: res,
                    note: note,
                    nonexist: nonexist
                });
            }
            focusNonexistHighlights = options.lightStep
                ? {}
                : this.buildNonexistHighlightMapForRow(idx);
        } finally {
            this.dataRows = savedDataRows;
        }
        if (isEmptyRow) {
            selectedLines.push({ date: '', id: '', result: '', note: '', nonexist: '' });
            lines.push('');
        }
        return {
            lines,
            selectedLines,
            clickedRowId,
            focusNonexistHighlights,
            contextPrefixCount,
            focusRowIndex: idx,
            isFocusEmpty: selectedLines.length > 0
                && !(String(selectedLines[selectedLines.length - 1].result || '').trim())
        };
    }

    applySourceRowFocusState(idx, rows, options = {}) {
        if (!Array.isArray(rows) || typeof idx !== 'number' || idx < 0 || idx >= rows.length) {
            return;
        }
        const data = options.prefetchedData
            || this.buildSourceRowLeftPaneData(idx, rows, options);
        if (!data) {
            return;
        }
        this.selectedLines = data.selectedLines;
        if (this.activeSheet === 'sheet1' || options.asSheet1) {
            const focusRow = rows[idx] || {};
            const nextFocusId = String(focusRow.id || focusRow.ID || data.clickedRowId || '').trim();
            const prevFocusId = String(this.comboFocusRowId || '').trim();
            const hadG1 = this.comboG1Enabled;
            const comboStateChanged = this.comboFocusRowId !== nextFocusId
                || this.comboFocusRowIndex !== idx
                || (this.isEmptyResultRow(focusRow) && hadG1);
            this.onComboFocusIdChanged(prevFocusId, nextFocusId);
            this.comboFocusRowId = nextFocusId;
            this.comboFocusRowIndex = idx;
            if (this.isEmptyResultRow(focusRow)) {
                this.comboG1Enabled = false;
            }
            if (comboStateChanged) {
                window.dispatchEvent(new CustomEvent('comboControlsChanged', {
                    detail: { sheet: options.asSheet1 ? 'sheet1' : this.activeSheet }
                }));
            }
        }
    }

    buildConn3ReferenceDetailForRow(idx) {
        if (typeof this.getNoteReferenceHintMeta !== 'function') {
            return {};
        }
        try {
            const meta = this.getNoteReferenceHintMeta(idx);
            if (!meta || meta.error) {
                return {};
            }
            return {
                referenceConn3HeaderLines: Array.isArray(meta.conn3HeaderLines) ? meta.conn3HeaderLines : [],
                referenceConn3Triplets: Array.isArray(meta.conn3Triplets) ? meta.conn3Triplets : [],
                referenceConn3FooterLine: meta.conn3FooterLine || ''
            };
        } catch (eRef) {
            return {};
        }
    }

    postLeftPaneSetLinesForSourceRow(idx, options = {}) {
        const rows = this.getSourceSheetRows();
        const data = this.buildSourceRowLeftPaneData(idx, rows, {
            lightStep: options.lightStep !== false
        });
        if (!data || !data.lines.length) {
            return;
        }
        const skipReferenceMeta = options.skipReferenceMeta === true;
        const refBundle = skipReferenceMeta ? {} : this.buildConn3ReferenceDetailForRow(idx);
        this.applySourceRowFocusState(idx, rows, { asSheet1: true, prefetchedData: data });
        window.dispatchEvent(new CustomEvent('leftPaneSetLines', {
            detail: {
                ...data,
                ...refBundle,
                trackingFrameStep: true,
                preserveSelection: true,
                skipReferenceMeta
            }
        }));
    }

    applySheet1NavPreviewAtIndex(idx, tableWrap, options = {}) {
        const wrap = tableWrap || this._sheet1NavTableWrap || document.getElementById('tableWrap');
        if (!wrap || this.activeSheet !== 'sheet1') {
            return;
        }
        if (typeof this.syncComboFocusFromSourceRowIndex === 'function') {
            this.syncComboFocusFromSourceRowIndex(idx);
        }
        const start = Math.max(0, idx - 10);
        const lightStep = options.lightStep !== false;
        this.applyWindowSelection(start, idx, idx, wrap, {
            previewOnly: true,
            skipFocusNoteRef: lightStep
        });
        if (!options.skipCenter) {
            this.centerActiveWindowInView(wrap);
        }
    }

    pumpSheet1NavStep() {
        this._sheet1LeftPaneTimer = 0;
        if (this._sheet1LeftPaneCurrent === this._sheet1LeftPaneTarget) {
            return;
        }
        this._sheet1LeftPaneCurrent += this._sheet1LeftPaneCurrent < this._sheet1LeftPaneTarget ? 1 : -1;
        const idx = this._sheet1LeftPaneCurrent;
        this.applySheet1NavPreviewAtIndex(idx, this._sheet1NavTableWrap, { lightStep: true });
        this.postLeftPaneSetLinesForSourceRow(idx, { lightStep: true, skipReferenceMeta: true });
        this._sheet1LeftPaneTimer = setTimeout(
            () => this.pumpSheet1NavStep(),
            this._sheet1LeftPaneStepMs
        );
    }

    scheduleSheet1NavToIndex(targetIdx, tableWrap) {
        const rows = this.getSourceSheetRows();
        if (!rows.length) {
            return false;
        }
        const next = Math.max(0, Math.min(rows.length - 1, targetIdx));
        this._sheet1NavTableWrap = tableWrap || this._sheet1NavTableWrap || document.getElementById('tableWrap');
        this._sheet1LeftPaneTarget = next;
        if (this._sheet1LeftPaneCurrent < 0) {
            const cur = this.activeWindowRange && typeof this.activeWindowRange.target === 'number'
                ? this.activeWindowRange.target
                : next;
            this._sheet1LeftPaneCurrent = cur;
        }
        if (!this._sheet1LeftPaneTimer) {
            this.pumpSheet1NavStep();
        }
        return true;
    }

    flushSheet1NavPump() {
        this.clearSheet1LeftPanePump();
        const target = this._sheet1LeftPaneTarget;
        if (target < 0 || this._sheet1LeftPaneCurrent === target) {
            return;
        }
        this._sheet1LeftPaneCurrent = target;
        this.applySheet1NavPreviewAtIndex(target, this._sheet1NavTableWrap);
        this.postLeftPaneSetLinesForSourceRow(target, { lightStep: true, skipReferenceMeta: true });
    }

    /**
     * Sheet1/tracking arrow preview: bơm đồng bộ viền sheet1 + iframe trái (giống tracking timeline).
     */
    syncLeftPaneFromSourceRowIndex(idx, options = {}) {
        if (options.immediate) {
            this.postLeftPaneSetLinesForSourceRow(idx, options);
            return;
        }
        const wrap = document.getElementById('tableWrap');
        this.scheduleSheet1NavToIndex(idx, wrap);
    }

    /**
     * Focus một hàng sheet1 (window, combo, rowClicked) dù đang xem sheet nào.
     */
    focusSourceSheetRow(idx, options = {}) {
        const rows = this.getSourceSheetRows();
        if (typeof idx !== 'number' || idx < 0 || idx >= rows.length) {
            return;
        }
        const row = rows[idx];
        const isEmpty = this.isEmptyResultRow(row);
        const savedDataRows = this.dataRows;
        this.dataRows = rows;
        try {
            this.onRowClick(idx, isEmpty, null, {
                skipSave: true,
                asSheet1: true,
                ...options
            });
        } finally {
            this.dataRows = savedDataRows;
        }
    }

    /**
     * Tracking → sheet1: timeline tới frame X thì focus kỳ id tương ứng.
     */
    syncSheet1FromTrackingFrame(frameIndex, options = {}) {
        if (this._syncingSheet1FromTracking || this._syncingTrackingFromSheet1) {
            return;
        }
        const st = this.sheets && this.sheets[TRACKING_SHEET_ID];
        if (!st || st.kind !== TRACKING_KIND) {
            return;
        }
        const rowIdx = this.getTrackingSourceRowIndexForFrame(st, frameIndex);
        if (rowIdx < 0) {
            return;
        }
        this._syncingTrackingFromSheet1 = true;
        try {
            if (options.trackingFrameStep) {
                this.postLeftPaneSetLinesForSourceRow(rowIdx, { skipReferenceMeta: true });
                return;
            }
            this.focusSourceSheetRow(rowIdx, {
                fromTrackingSync: true,
                trackingFrameStep: false,
                light: true,
                skipCenter: this.activeSheet !== 'sheet1'
            });
        } finally {
            this._syncingTrackingFromSheet1 = false;
        }
    }

    /**
     * Sheet1 → tracking một chiều: record id đang focus là X thì timeline tua tới
     * mốc đã xử lý đến kỳ có id X.
     */
    syncTrackingTimelineFromSheet1Row(focusRowIndex) {
        const st = this.sheets && this.sheets[TRACKING_SHEET_ID];
        if (!st || st.kind !== TRACKING_KIND) {
            return;
        }
        const rows = this.sourceRows || [];
        const row = rows[focusRowIndex];
        if (!row) {
            return;
        }
        this.ensureTrackingFrames(st);
        const viewMode = this.getTrackingViewMode(st);
        const frames = st.frames || [];
        const total = frames.length;
        if (total < 1) {
            return;
        }

        let frameIdx = this.getTrackingFrameIndexForSourceRow(st, focusRowIndex);
        if (frameIdx < 0) {
            frameIdx = 0;
        }
        frameIdx = Math.max(0, Math.min(total - 1, frameIdx));

        const tailRow = rows.length ? rows[rows.length - 1] : {};
        const tailId = String(tailRow.id ?? tailRow.ID ?? '');
        const uiSig = `${viewMode}|${total}|${rows.length}|${rows.length}|${tailId}`;
        const prev = st.trackingUi && typeof st.trackingUi === 'object' ? st.trackingUi : {};
        const speed = Number.isFinite(prev.speed) ? Math.min(3, Math.max(0.5, prev.speed)) : 1;
        const snap = {
            sig: uiSig,
            frameIndex: frameIdx,
            focusSourceRowIndex: focusRowIndex,
            playing: false,
            speed,
            predictNeonOn: !!prev.predictNeonOn,
            focusNumsByMode: RightPaneSheetManager.serializeTrackingFocusNumsByMode(
                RightPaneSheetManager.readTrackingFocusNumsByMode(prev)
            ),
            viewMode,
            labelMode: RightPaneSheetManager.normalizeTrackingLabelMode(
                st.trackingLabelMode || prev.labelMode
            )
        };
        st.trackingUi = snap;
        try {
            sessionStorage.setItem(TRACKING_UI_STORAGE_KEY, JSON.stringify(snap));
        } catch (e) {
            /* ignore */
        }

        const tw = document.getElementById('tableWrap');
        if (tw && tw.classList.contains('table-wrap--tracking') && typeof tw.__trackingSeekFrame === 'function') {
            this._syncingSheet1FromTracking = true;
            try {
                tw.__trackingSeekFrame(frameIdx, { skipSheet1Sync: true });
            } finally {
                this._syncingSheet1FromTracking = false;
            }
        }
    }

    syncSpecialTrackingTimelineFromSheet1Row(focusRowIndex) {
        this.syncTrackingTimelineFromSheet1Row(focusRowIndex);
    }

    /** Arrow keys: ±1 frame timeline tracking (hoạt động kể cả khi focus còn ở nửa trái). */
    stepTrackingTimelineFrame(delta) {
        const step = Number(delta) || 0;
        if (!step || this.activeSheet !== TRACKING_SHEET_ID) {
            return false;
        }
        const tableWrap = document.getElementById('tableWrap');
        if (!tableWrap || !tableWrap.classList.contains('table-wrap--tracking')) {
            return false;
        }
        if (typeof tableWrap.__trackingStepFrame === 'function') {
            tableWrap.__trackingStepFrame(step);
            return true;
        }
        return false;
    }

    /** Basic + id cuối (preview): ↑/↓ chọn bar theo thứ tự freq (slot), giả lập khoanh trái. */
    stepBasicTrackingLastIdBarNav(delta) {
        const step = Number(delta) || 0;
        if (!step || this.activeSheet !== TRACKING_SHEET_ID) {
            return false;
        }
        const tableWrap = document.getElementById('tableWrap');
        if (!tableWrap || !tableWrap.classList.contains('table-wrap--tracking')) {
            return false;
        }
        if (typeof tableWrap.__trackingStepBasicLastIdBar !== 'function') {
            return false;
        }
        return !!tableWrap.__trackingStepBasicLastIdBar(step);
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
    buildSpecialTrackingFrames(drawSteps) {
        const counts = {};
        for (let i = 1; i <= 12; i++) {
            counts[i] = 0;
        }
        const frames = [];
        const list = [];
        const steps = drawSteps || [];
        let drawIndex = -1;
        for (let f = 0; f < steps.length; f++) {
            const just = steps[f];
            const holdFrame = just == null;
            if (!holdFrame) {
                drawIndex += 1;
                list.push(just);
                counts[just] += 1;
            }
            const endDrawIdx = drawIndex;
            const sorted = Object.keys(counts)
                .map((k) => {
                    const n = Number(k);
                    const v = counts[n];
                    const t = v > 0 && endDrawIdx >= 0
                        ? RightPaneSheetManager.specialTrackingStepOfVthHit(list, n, v, endDrawIdx)
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
                drawIndex: endDrawIdx,
                holdFrame,
                justDrawn: just,
                justDrawnNums: just != null ? [just] : [],
                byNum: { ...counts },
                sorted,
                maxV,
                slotByNum,
                wPctByNum
            });
        }
        return frames;
    }

    /** Frame sau khi đã xử lý `drawCount` lần quay hợp lệ (bỏ qua hàng id/result rỗng). */
    static getTrackingFrameIndexAfterDraws(frames, drawCount) {
        if (!frames || !frames.length || drawCount < 1) {
            return -1;
        }
        const target = drawCount - 1;
        for (let f = frames.length - 1; f >= 0; f--) {
            const di = frames[f].drawIndex;
            if (di != null && di >= 0 && di === target) {
                return f;
            }
        }
        for (let f = frames.length - 1; f >= 0; f--) {
            const di = frames[f].drawIndex;
            if (di != null && di >= 0 && di < target) {
                return f;
            }
        }
        return 0;
    }

    static normalizeTrackingViewMode(mode) {
        return mode === 'special' ? 'special' : 'basic';
    }

    static normalizeTrackingLabelMode(mode) {
        return mode === 'out' ? 'out' : 'in';
    }

    /** Shift+click viền cam: đọc một mảng số (lọc theo numMax). */
    static normalizeTrackingFocusNumsList(list, numMax) {
        const out = new Set();
        const max = Number(numMax);
        const inRange = (n) => Number.isFinite(n) && n >= 1 && n <= max;
        if (Array.isArray(list)) {
            list.forEach((fn) => {
                const n = Math.floor(Number(fn));
                if (inRange(n)) {
                    out.add(n);
                }
            });
        }
        return out;
    }

    /** Legacy: focusNums[] hoặc focusNum đơn. */
    static normalizeTrackingFocusNums(saved, numMax) {
        const out = new Set();
        const max = Number(numMax);
        const inRange = (n) => Number.isFinite(n) && n >= 1 && n <= max;
        if (saved && Array.isArray(saved.focusNums)) {
            saved.focusNums.forEach((fn) => {
                const n = Math.floor(Number(fn));
                if (inRange(n)) {
                    out.add(n);
                }
            });
        } else if (saved && Number.isFinite(saved.focusNum)) {
            const n = Math.floor(saved.focusNum);
            if (inRange(n)) {
                out.add(n);
            }
        }
        return out;
    }

    /** Basic / special mỗi mode nhớ focus riêng. */
    static readTrackingFocusNumsByMode(saved) {
        const basicMax = 35;
        const specialMax = 12;
        if (saved && saved.focusNumsByMode && typeof saved.focusNumsByMode === 'object') {
            return {
                basic: RightPaneSheetManager.normalizeTrackingFocusNumsList(
                    saved.focusNumsByMode.basic,
                    basicMax
                ),
                special: RightPaneSheetManager.normalizeTrackingFocusNumsList(
                    saved.focusNumsByMode.special,
                    specialMax
                )
            };
        }
        const legacyMode = RightPaneSheetManager.normalizeTrackingViewMode(saved && saved.viewMode);
        const legacyMax = legacyMode === 'basic' ? basicMax : specialMax;
        const legacy = RightPaneSheetManager.normalizeTrackingFocusNums(saved, legacyMax);
        return {
            basic: legacyMode === 'basic' ? legacy : new Set(),
            special: legacyMode === 'special' ? legacy : new Set()
        };
    }

    static serializeTrackingFocusNumsByMode(byMode) {
        const basic = byMode && byMode.basic instanceof Set ? byMode.basic : new Set();
        const special = byMode && byMode.special instanceof Set ? byMode.special : new Set();
        return {
            basic: Array.from(basic).sort((a, b) => a - b),
            special: Array.from(special).sort((a, b) => a - b)
        };
    }

    static readTrackingLabelModeFromStorage() {
        try {
            const raw = sessionStorage.getItem(TRACKING_LABEL_MODE_KEY);
            if (raw == null || raw === '') {
                return 'out';
            }
            return RightPaneSheetManager.normalizeTrackingLabelMode(raw);
        } catch (e) {
            return 'out';
        }
    }

    static writeTrackingLabelModeToStorage(mode) {
        try {
            sessionStorage.setItem(
                TRACKING_LABEL_MODE_KEY,
                RightPaneSheetManager.normalizeTrackingLabelMode(mode)
            );
        } catch (e) {
            /* ignore */
        }
    }

    getTrackingViewMode(sheet) {
        return RightPaneSheetManager.normalizeTrackingViewMode(sheet && sheet.trackingViewMode);
    }

    getTrackingSlotCount(viewMode) {
        return viewMode === 'basic' ? 35 : 12;
    }

    /**
     * Mỗi hàng sheet1 = một bước timeline; `drawSteps[ri]` = 5 số chính hoặc null nếu không hợp lệ.
     */
    buildBasicTrackingSeriesMeta(rows) {
        const draws = [];
        const drawSteps = [];
        const sourceRowIndices = [];
        const list = rows || [];
        for (let ri = 0; ri < list.length; ri++) {
            sourceRowIndices.push(ri);
            const row = list[ri];
            let step = null;
            if (!this.isEmptyResultRow(row)) {
                const main = this.parseMainNums(row.result || row.Result || '');
                if (main.length === 5) {
                    step = main.slice();
                    draws.push(main.slice());
                }
            }
            drawSteps.push(step);
        }
        return { draws, drawSteps, sourceRowIndices };
    }

    /**
     * Nhóm frequency trùng nhau (≥2 số), xếp theo slot trên stack.
     * @returns {{ freq: number, minSlot: number, maxSlot: number, nums: number[] }[]}
     */
    static buildBasicTrackingFreqTieGroups(byNum, slotByNum, numMax = 35) {
        /** @type {Map<number, { n: number, slot: number }[]>} */
        const groups = new Map();
        for (let n = 1; n <= numMax; n++) {
            const freq = (byNum && byNum[n]) || 0;
            if (!groups.has(freq)) {
                groups.set(freq, []);
            }
            groups.get(freq).push({
                n,
                slot: (slotByNum && slotByNum[n]) ?? numMax
            });
        }
        const out = [];
        groups.forEach((members, freq) => {
            if (members.length < 2) {
                return;
            }
            members.sort((a, b) => a.slot - b.slot || a.n - b.n);
            out.push({
                freq,
                minSlot: members[0].slot,
                maxSlot: members[members.length - 1].slot,
                nums: members.map((m) => m.n)
            });
        });
        out.sort((a, b) => a.minSlot - b.minSlot || a.freq - b.freq);
        return out;
    }

    static getFreqTieGroupBellyKey(nums) {
        if (!Array.isArray(nums) || !nums.length) {
            return '';
        }
        return nums.slice().sort((a, b) => a - b).join(',');
    }

    /** Khóa đếm streak: cùng freq + cùng tập số (đổi freq → đếm lại từ 1). */
    static getFreqTieGroupStreakKey(group) {
        if (!group || !Array.isArray(group.nums) || !group.nums.length) {
            return '';
        }
        const numsKey = RightPaneSheetManager.getFreqTieGroupBellyKey(group.nums);
        if (!numsKey) {
            return '';
        }
        return `${group.freq | 0}:${numsKey}`;
    }

    /** Bụng ghost: chỉ nhóm khác bụng hiện tại (submit / giả lập) tại cùng freq. */
    static filterTrackingFreqGhostGroups(ghostGroups, currentGroups) {
        const bellyKeyOf = RightPaneSheetManager.getFreqTieGroupBellyKey;
        const currentByFreq = new Map();
        for (let i = 0; i < (currentGroups || []).length; i++) {
            const g = currentGroups[i];
            currentByFreq.set(g.freq | 0, g);
        }
        const out = [];
        for (let i = 0; i < (ghostGroups || []).length; i++) {
            const gg = ghostGroups[i];
            const cg = currentByFreq.get(gg.freq | 0);
            if (!cg) {
                out.push(gg);
                continue;
            }
            if (bellyKeyOf(gg.nums) !== bellyKeyOf(cg.nums)
                || gg.minSlot !== cg.minSlot
                || gg.maxSlot !== cg.maxSlot) {
                out.push(gg);
            }
        }
        return out;
    }

    /** Khóa streak của bụng ghost đã thay đổi so với bụng hiện tại (submit / giả lập). */
    static getTrackingFreqGhostChangedKeySet(ghostGroups, currentGroups) {
        const bellyKeyOf = RightPaneSheetManager.getFreqTieGroupBellyKey;
        const streakKeyOf = RightPaneSheetManager.getFreqTieGroupStreakKey;
        const currentByFreq = new Map();
        for (let i = 0; i < (currentGroups || []).length; i++) {
            const g = currentGroups[i];
            currentByFreq.set(g.freq | 0, g);
        }
        const changed = new Set();
        for (let i = 0; i < (ghostGroups || []).length; i++) {
            const gg = ghostGroups[i];
            const key = streakKeyOf(gg);
            if (!key) {
                continue;
            }
            const cg = currentByFreq.get(gg.freq | 0);
            if (!cg
                || bellyKeyOf(gg.nums) !== bellyKeyOf(cg.nums)
                || gg.minSlot !== cg.minSlot
                || gg.maxSlot !== cg.maxSlot) {
                changed.add(key);
            }
        }
        return changed;
    }

    /** Khóa exact bụng: tập số + minSlot + maxSlot. */
    static getTrackingFreqBellyExactKey(group) {
        if (!group || !Array.isArray(group.nums) || group.nums.length < 2) {
            return '';
        }
        const bellyKeyOf = RightPaneSheetManager.getFreqTieGroupBellyKey;
        return `${bellyKeyOf(group.nums)}:${group.minSlot}:${group.maxSlot}`;
    }

    /**
     * Khóa streak của bụng trong source không còn y hệt (nums+slot) trong target.
     * Dùng cho solid: current vs trước giả lập/submit — bụng đổi → nét liền.
     */
    static getTrackingFreqBellyStreakKeysWithoutExactMatchIn(sourceGroups, targetGroups) {
        const streakKeyOf = RightPaneSheetManager.getFreqTieGroupStreakKey;
        const targetExact = new Set();
        for (let i = 0; i < (targetGroups || []).length; i++) {
            const exact = RightPaneSheetManager.getTrackingFreqBellyExactKey(targetGroups[i]);
            if (exact) {
                targetExact.add(exact);
            }
        }
        const out = new Set();
        for (let i = 0; i < (sourceGroups || []).length; i++) {
            const g = sourceGroups[i];
            const key = streakKeyOf(g);
            const exact = RightPaneSheetManager.getTrackingFreqBellyExactKey(g);
            if (key && exact && !targetExact.has(exact)) {
                out.add(key);
            }
        }
        return out;
    }

    /** Bụng solid đổi so với tham chiếu (trước giả lập / submit) → nét liền; còn lại chấm li ti. */
    static getTrackingFreqSolidShiftedKeys(referenceGroups, currentGroups) {
        const streakKeyOf = RightPaneSheetManager.getFreqTieGroupStreakKey;
        const refStreakKeys = new Set();
        const refExact = new Set();
        for (let i = 0; i < (referenceGroups || []).length; i++) {
            const rg = referenceGroups[i];
            const key = streakKeyOf(rg);
            const exact = RightPaneSheetManager.getTrackingFreqBellyExactKey(rg);
            if (key) {
                refStreakKeys.add(key);
            }
            if (exact) {
                refExact.add(exact);
            }
        }
        const shifted = new Set();
        for (let i = 0; i < (currentGroups || []).length; i++) {
            const cg = currentGroups[i];
            const key = streakKeyOf(cg);
            const exact = RightPaneSheetManager.getTrackingFreqBellyExactKey(cg);
            if (!key) {
                continue;
            }
            if (!refStreakKeys.has(key) || !refExact.has(exact)) {
                shifted.add(key);
            }
        }
        return shifted;
    }

    /**
     * Giả lập đủ (basic 5 / special 1): nhãn solid bụng không đổi = trước giả lập + 1.
     * Bụng trong `solidShiftedKeys` là bụng đã đổi — không bump.
     */
    static applyTrackingSolidPreviewStreakBump(groups, streakByKey, beforeTie, solidShiftedKeys) {
        const out = new Map(streakByKey || []);
        if (!beforeTie || !Array.isArray(groups) || !groups.length) {
            return out;
        }
        const bellyKeyOf = RightPaneSheetManager.getFreqTieGroupBellyKey;
        const streakKeyOf = RightPaneSheetManager.getFreqTieGroupStreakKey;
        const beforeGroups = beforeTie.groups || [];
        const beforeStreakByKey = beforeTie.streakByKey || new Map();
        const findBeforeStreak = (group) => {
            const key = streakKeyOf(group);
            if (key && beforeStreakByKey.has(key)) {
                return beforeStreakByKey.get(key);
            }
            const exact = RightPaneSheetManager.getTrackingFreqBellyExactKey(group);
            if (exact) {
                for (let i = 0; i < beforeGroups.length; i++) {
                    const bg = beforeGroups[i];
                    if (RightPaneSheetManager.getTrackingFreqBellyExactKey(bg) === exact) {
                        const bk = streakKeyOf(bg);
                        if (bk && beforeStreakByKey.has(bk)) {
                            return beforeStreakByKey.get(bk);
                        }
                    }
                }
            }
            const bkNums = bellyKeyOf(group.nums);
            if (bkNums) {
                for (let i = 0; i < beforeGroups.length; i++) {
                    const bg = beforeGroups[i];
                    if (bellyKeyOf(bg.nums) === bkNums) {
                        const bk = streakKeyOf(bg);
                        if (bk && beforeStreakByKey.has(bk)) {
                            return beforeStreakByKey.get(bk);
                        }
                    }
                }
            }
            return null;
        };
        for (let i = 0; i < groups.length; i++) {
            const g = groups[i];
            const key = streakKeyOf(g);
            if (!key) {
                continue;
            }
            if (solidShiftedKeys && solidShiftedKeys.has(key)) {
                continue;
            }
            const beforeStreak = findBeforeStreak(g);
            if (beforeStreak != null) {
                out.set(
                    key,
                    Math.min(
                        RightPaneSheetManager.FREQ_BRACE_STREAK_MAX,
                        (beforeStreak | 0) + 1
                    )
                );
            }
        }
        return out;
    }

    /** Special giả lập 1 số: bụng mờ cho nhóm trước preview (kể cả khi pick rời khỏi bụng cũ). */
    static filterTrackingSpecialPreviewGhostGroups(beforeGroups, afterGroups, previewPickNum, beforeCounts) {
        const bellyKeyOf = RightPaneSheetManager.getFreqTieGroupBellyKey;
        const streakKeyOf = RightPaneSheetManager.getFreqTieGroupStreakKey;
        const out = [];
        const seen = new Set();
        const pushGroup = (gg) => {
            const key = streakKeyOf(gg);
            if (!key || seen.has(key)) {
                return;
            }
            seen.add(key);
            out.push(gg);
        };
        const base = RightPaneSheetManager.filterTrackingFreqGhostGroups(beforeGroups, afterGroups);
        for (let i = 0; i < base.length; i++) {
            pushGroup(base[i]);
        }
        const pick = Number.isFinite(previewPickNum) ? (previewPickNum | 0) : 0;
        if (!pick) {
            return out;
        }
        const beforeFreq = (beforeCounts && beforeCounts[pick]) | 0;
        for (let i = 0; i < (beforeGroups || []).length; i++) {
            const gg = beforeGroups[i];
            if (!Array.isArray(gg.nums) || gg.nums.length < 2) {
                continue;
            }
            const atPickFreq = beforeFreq > 0 && (gg.freq | 0) === beforeFreq;
            const containsPick = gg.nums.includes(pick);
            if (!atPickFreq && !containsPick) {
                continue;
            }
            const cg = (afterGroups || []).find((g) => (g.freq | 0) === (gg.freq | 0));
            if (!cg) {
                pushGroup(gg);
                continue;
            }
            if (containsPick && !cg.nums.includes(pick)) {
                pushGroup(gg);
                continue;
            }
            if (bellyKeyOf(gg.nums) !== bellyKeyOf(cg.nums)
                || gg.minSlot !== cg.minSlot
                || gg.maxSlot !== cg.maxSlot) {
                pushGroup(gg);
            }
        }
        return out;
    }

    static FREQ_BRACE_STREAK_NUMERALS = '一二三四五六七八九';
    static FREQ_BRACE_STREAK_MAX = 99;

    /** 1–99 → 一…十, 十一…十九, 二十…九十九 */
    static formatChineseStreakNumeral(streak) {
        const n = Math.max(1, Math.min(RightPaneSheetManager.FREQ_BRACE_STREAK_MAX, streak | 0));
        const digits = RightPaneSheetManager.FREQ_BRACE_STREAK_NUMERALS;
        if (n <= 10) {
            if (n === 10) {
                return '十';
            }
            return digits[n - 1];
        }
        if (n < 20) {
            return `十${digits[n - 11]}`;
        }
        const tens = Math.floor(n / 10);
        const ones = n % 10;
        const tensChar = digits[tens - 1];
        if (ones === 0) {
            return `${tensChar}十`;
        }
        return `${tensChar}十${digits[ones - 1]}`;
    }

    static formatFreqBraceCountLabel(memberCount, unchangedStreak) {
        const count = Math.max(2, memberCount | 0);
        const streak = Math.max(1, Math.min(RightPaneSheetManager.FREQ_BRACE_STREAK_MAX, unchangedStreak | 0));
        if (streak <= 1) {
            return String(count);
        }
        const numeral = RightPaneSheetManager.formatChineseStreakNumeral(streak);
        return `${count}${numeral}`;
    }

    /**
     * Đếm số frame liên tiếp (lùi từ frameIndex) mà bụng cùng freq + cùng tập số vẫn tồn tại.
     * Khi giả lập (preview) làm đổi tập số trong bụng tại frame hiện tại → streak = 1 (vd. 3四 → 2一).
     * @returns {{ groups: object[], streakByKey: Map<string, number> }}
     */
    computeTrackingFreqTieGroupsWithStreaks(sheet, frames, frameIndex, options = {}) {
        const isBasic = !!options.isBasic;
        const numMax = options.numMax || 35;
        const leftSubmitOn = !!options.leftSubmitOn;
        const basicDraws = options.basicDraws || [];
        const streakKeyOf = RightPaneSheetManager.getFreqTieGroupStreakKey;
        const bellyKeyOf = RightPaneSheetManager.getFreqTieGroupBellyKey;
        const previewPickNumsResolved = Array.isArray(options.previewPickNums)
            ? options.previewPickNums
            : (this.leftBasicPreviewPickNums || []);
        const previewAtPaintFrame = options.previewAtPaintFrame != null
            ? !!options.previewAtPaintFrame
            : (isBasic && this.isBasicTrackingFreqPreviewLayoutActive(sheet, frameIndex));
        const previewPickNums = previewAtPaintFrame ? previewPickNumsResolved : [];
        const specialPreviewPickResolved = options.specialPreviewPick != null
            ? options.specialPreviewPick
            : this.leftSpecialPreviewPickNum;
        const specialPreviewAtPaintFrame = options.previewAtPaintFrame != null
            ? !!options.previewAtPaintFrame
            : (!isBasic && this.isSpecialTrackingFreqPreviewLayoutActive(sheet, frameIndex));
        const layoutPreviewAtPaintFrame = isBasic ? previewAtPaintFrame : specialPreviewAtPaintFrame;
        const specialPreviewPick = specialPreviewAtPaintFrame && specialPreviewPickResolved != null
            ? specialPreviewPickResolved
            : null;

        const getGroupsForFrame = (f, usePreview = false) => {
            const fr = frames[f];
            if (!fr) {
                return [];
            }
            if (isBasic) {
                const applyPreview = usePreview && f === frameIndex && previewAtPaintFrame;
                const picks = applyPreview ? previewPickNums : [];
                const display = RightPaneSheetManager.computeBasicTrackingDisplayLayout(
                    basicDraws,
                    fr,
                    leftSubmitOn,
                    picks
                );
                return RightPaneSheetManager.buildBasicTrackingFreqTieGroups(
                    display.counts,
                    display.slotByNum,
                    numMax
                );
            }
            const series = sheet.specialSeries || sheet.series || [];
            const previewForLayout = (usePreview && f === frameIndex && specialPreviewAtPaintFrame)
                ? specialPreviewPick
                : null;
            const display = RightPaneSheetManager.computeSpecialTrackingDisplayLayout(
                series,
                fr,
                leftSubmitOn,
                previewForLayout
            );
            return RightPaneSheetManager.buildBasicTrackingFreqTieGroups(
                display.counts,
                display.slotByNum,
                numMax
            );
        };

        const groups = getGroupsForFrame(frameIndex, true);
        const committedGroupsAtPaint = layoutPreviewAtPaintFrame
            ? getGroupsForFrame(frameIndex, false)
            : [];
        /** freq → tập số bụng đã commit tại frame đang vẽ */
        const committedBellyByFreq = new Map();
        for (let c = 0; c < committedGroupsAtPaint.length; c++) {
            const cg = committedGroupsAtPaint[c];
            committedBellyByFreq.set(cg.freq | 0, bellyKeyOf(cg.nums));
        }
        const committedKeySet = new Set(committedGroupsAtPaint.map((g) => streakKeyOf(g)));

        const findBellyKeyAtFreq = (frameGroups, freq) => {
            const want = freq | 0;
            for (let i = 0; i < frameGroups.length; i++) {
                const g = frameGroups[i];
                if ((g.freq | 0) === want) {
                    return bellyKeyOf(g.nums);
                }
            }
            return null;
        };

        const isPreviewBellyCompositionChanged = (group) => {
            if (!layoutPreviewAtPaintFrame || !group) {
                return false;
            }
            const key = streakKeyOf(group);
            if (committedKeySet.has(key)) {
                return false;
            }
            const freq = group.freq | 0;
            const committedBelly = committedBellyByFreq.get(freq);
            if (committedBelly === undefined) {
                return true;
            }
            return bellyKeyOf(group.nums) !== committedBelly;
        };

        /** Basic đủ 5 số giả lập: đếm streak anchor giống Submit ON (bỏ hold frame). */
        const basicFullPreviewSim = isBasic
            && layoutPreviewAtPaintFrame
            && previewPickNums.length >= 5
            && !leftSubmitOn;
        const streakSubmitOn = leftSubmitOn || basicFullPreviewSim;

        /** Hold frame + Submit ON (hoặc giả lập đủ 5): không tính thêm streak tại hold frame.
         *  Submit OFF + giả lập chưa đủ 5: hold frame vẫn là bước preview hợp lệ. */
        const resolveStreakAnchorFrameIndex = (idx) => {
            if (!streakSubmitOn) {
                return idx;
            }
            let anchor = idx;
            while (anchor > 0 && frames[anchor] && frames[anchor].holdFrame) {
                anchor--;
            }
            return anchor;
        };
        const streakAnchorFrameIndex = resolveStreakAnchorFrameIndex(frameIndex);

        const streakByKey = new Map();
        for (let g = 0; g < groups.length; g++) {
            const group = groups[g];
            const key = streakKeyOf(group);
            if (!key || streakByKey.has(key)) {
                continue;
            }
            const groupFreq = group.freq | 0;
            const bellyKey = bellyKeyOf(group.nums);
            let streak = 1;
            const previewCompositionChanged = isPreviewBellyCompositionChanged(group);
            if (!previewCompositionChanged) {
                for (let f = streakAnchorFrameIndex - 1; f >= 0; f--) {
                    if (streakSubmitOn && frames[f] && frames[f].holdFrame) {
                        continue;
                    }
                    const prevBelly = findBellyKeyAtFreq(getGroupsForFrame(f, false), groupFreq);
                    if (prevBelly !== bellyKey) {
                        break;
                    }
                    streak++;
                    if (streak >= RightPaneSheetManager.FREQ_BRACE_STREAK_MAX) {
                        break;
                    }
                }
            }
            streakByKey.set(key, streak);
        }
        return { groups, streakByKey };
    }

    /** Dấu { cột frequency — nhóm ≥2 số cùng giá trị tích lũy; nhãn = [n] + 一二… (freq + tập số không đổi). */
    static syncTrackingFreqBraces(layer, groups, slotCount, streakByKey, options = {}) {
        if (!layer) {
            return;
        }
        const ghost = !!options.ghost;
        const ghostChangedKeys = options.ghostChangedKeys;
        const solidShiftedKeys = options.solidShiftedKeys;
        const sig = !groups.length
            ? ''
            : groups.map((g) => {
                const key = RightPaneSheetManager.getFreqTieGroupStreakKey(g);
                const streak = (streakByKey && streakByKey.get(key)) || 1;
                const changedFlag = ghost && ghostChangedKeys && ghostChangedKeys.has(key) ? 1 : 0;
                const solidShiftedFlag = !ghost && solidShiftedKeys && solidShiftedKeys.has(key) ? 1 : 0;
                return `${g.freq}:${g.minSlot}-${g.maxSlot}:${key}:${streak}:${changedFlag}:${solidShiftedFlag}`;
            }).join('|');
        const sigKey = ghost ? 'stBraceGhostSig' : 'stBraceSig';
        const fullSig = (ghost ? 'ghost|' : '') + sig;
        if (layer.dataset[sigKey] === fullSig) {
            return;
        }
        layer.dataset[sigKey] = fullSig;
        layer.replaceChildren();
        if (!groups.length || !slotCount) {
            return;
        }
        const frag = document.createDocumentFragment();
        for (const g of groups) {
            const memberCount = Array.isArray(g.nums) ? g.nums.length : 0;
            if (memberCount < 2) {
                continue;
            }
            const streakKey = RightPaneSheetManager.getFreqTieGroupStreakKey(g);
            const unchangedStreak = (streakByKey && streakByKey.get(streakKey)) || 1;
            const countLabel = RightPaneSheetManager.formatFreqBraceCountLabel(memberCount, unchangedStreak);
            const ghostChanged = ghost && ghostChangedKeys && ghostChangedKeys.has(streakKey);
            const solidShifted = !ghost && solidShiftedKeys && solidShiftedKeys.has(streakKey);
            const el = document.createElement('div');
            el.className = ghost
                ? ('special-tracking-freq-brace special-tracking-freq-brace--ghost'
                    + (ghostChanged ? ' special-tracking-freq-brace--ghost-changed' : ''))
                : ('special-tracking-freq-brace'
                    + (solidShifted ? ' special-tracking-freq-brace--solid-shifted' : ''));
            el.setAttribute('data-freq', String(g.freq));
            el.setAttribute('data-member-count', String(memberCount));
            el.setAttribute('data-unchanged-streak', String(unchangedStreak));
            el.title = ghost
                ? `${options.ghostLabel || 'Trước submit'} — tần suất ${g.freq}: ${g.nums.join(', ')} (${memberCount} số, freq + tập số không đổi ${unchangedStreak} kỳ${ghostChanged ? ', đã thay đổi sau submit' : ''})`
                : `Tần suất ${g.freq}: ${g.nums.join(', ')} (${memberCount} số, freq + tập số không đổi ${unchangedStreak} kỳ)`;
            const topPct = (g.minSlot / slotCount) * 100;
            const heightPct = ((g.maxSlot - g.minSlot + 1) / slotCount) * 100;
            el.style.top = `${topPct}%`;
            el.style.height = `${heightPct}%`;
            const bracePathD = ghost
                ? 'M 8,0 C 3,1 1,22 1,50 C 1,78 3,99 8,100'
                : 'M 0,0 C 5,1 7,22 7,50 C 7,78 5,99 0,100';
            el.innerHTML = '<span class="special-tracking-freq-brace-count" aria-hidden="true">'
                + countLabel
                + '</span>'
                + '<svg viewBox="0 0 8 100" preserveAspectRatio="none" aria-hidden="true">'
                + '<path d="' + bracePathD + '" fill="none" '
                + 'stroke="currentColor" stroke-width="1.75" vector-effect="non-scaling-stroke" '
                + 'stroke-linecap="round" stroke-linejoin="round"/>'
                + '</svg>';
            frag.appendChild(el);
        }
        layer.appendChild(frag);
    }

    /** @deprecated */
    static syncBasicTrackingFreqBraces(layer, groups, slotCount) {
        RightPaneSheetManager.syncTrackingFreqBraces(layer, groups, slotCount);
    }

    setLeftSubmitActive(on) {
        const next = !!on;
        const prev = this.leftSubmitActive;
        if (next === prev) {
            return;
        }
        if (next) {
            this.leftBasicPreviewPickNumsStash = (this.leftBasicPreviewPickNums || []).slice();
            this.leftBasicPreviewPickNums = [];
            this.leftSpecialPreviewPickNumStash = this.leftSpecialPreviewPickNum;
            this.leftSpecialPreviewPickNum = null;
        }
        this.leftSubmitActive = next;
        if (!next) {
            if (Array.isArray(this.leftBasicPreviewPickNumsStash)
                && this.leftBasicPreviewPickNumsStash.length) {
                // Basic: iframe khôi phục preSubmit khi Submit OFF.
                this._basicPreviewStashRestoredAt = Date.now();
                try {
                    window.dispatchEvent(new CustomEvent('leftCircledNumsChanged'));
                } catch (eEv) { /* ignore */ }
            }
            this.leftBasicPreviewPickNumsStash = [];
            if (this.leftSpecialPreviewPickNumStash != null) {
                this.leftSpecialPreviewPickNum = this.leftSpecialPreviewPickNumStash;
                this.leftSpecialPreviewPickNumStash = null;
            }
        }
        this.requestTrackingUiRepaintIfActive();
    }

    setLeftAutoringEnabled(on) {
        this.leftAutoringEnabled = !!on;
    }

    setLeftBasicPreviewPickNums(nums, options = {}) {
        const next = [];
        const seen = new Set();
        if (Array.isArray(nums)) {
            for (let i = 0; i < nums.length; i++) {
                const n = nums[i];
                if (n >= 1 && n <= 35 && !seen.has(n)) {
                    seen.add(n);
                    next.push(n);
                }
            }
        }
        const prev = this.leftBasicPreviewPickNums || [];
        const same = prev.length === next.length && prev.every((v, i) => v === next[i]);
        if (!same) {
            this.leftBasicPreviewPickNums = next;
            this._leftBasicPreviewPickGeneration += 1;
            if (!options.skipFocusUpdate) {
                this.applyTrackingPreviewFocusAfterPickNumsChange(prev, next);
            }
        }
        return !same;
    }

    /**
     * Đồng bộ focus chuột phải theo chuỗi pick khi đổi từ nửa trái (hoặc nguồn ngoài).
     * Thêm số → focus = phần tử cuối chuỗi.
     * Bớt số (tắt khoanh trái / sync) → focus = phần tử cuối còn lại.
     */
    applyTrackingPreviewFocusAfterPickNumsChange(prev, next) {
        const prevList = Array.isArray(prev) ? prev : [];
        const nextList = Array.isArray(next) ? next : [];
        const prevSet = new Set(prevList);
        const nextSet = new Set(nextList);
        const added = nextList.some((n) => !prevSet.has(n));
        const removed = prevList.some((n) => !nextSet.has(n));
        if (added || removed || !nextList.length) {
            this.lastTrackingPreviewBarNum = nextList.length
                ? nextList[nextList.length - 1]
                : null;
            return;
        }
        const focus = this.lastTrackingPreviewBarNum;
        if (focus == null || !nextSet.has(focus)) {
            this.lastTrackingPreviewBarNum = nextList.length
                ? nextList[nextList.length - 1]
                : null;
        }
    }

    /** Sau khôi phục stash: bỏ qua leftCircledNumsReady trống từ iframe (tránh xóa pick giả lập). */
    shouldIgnoreEmptyLeftCircledNumsAfterStashRestore(nums) {
        const ts = this._basicPreviewStashRestoredAt || 0;
        if (!ts || Date.now() - ts > 800) {
            return false;
        }
        const incoming = Array.isArray(nums) ? nums : [];
        if (incoming.length) {
            this._basicPreviewStashRestoredAt = 0;
            return false;
        }
        return (this.leftBasicPreviewPickNums || []).length > 0;
    }

    /** Đổi id focus: xóa giả lập bar phải và đồng bộ sạch sang nửa trái (tránh viền đen còn mà khoanh trái đã mất). */
    clearLeftBasicBarPreviewPicksOnFocusChange() {
        const had = Array.isArray(this.leftBasicPreviewPickNums) && this.leftBasicPreviewPickNums.length > 0;
        this.leftBasicPreviewPickNums = [];
        this.lastTrackingPreviewBarNum = null;
        this._leftBasicPreviewPickGeneration += 1;
        if (this.shouldSyncBasicBarPickToLeftPane()) {
            this.syncLeftPickSelectionToIframe([]);
        }
        if (had) {
            try {
                window.dispatchEvent(new CustomEvent('leftCircledNumsChanged'));
            } catch (ePaint) { /* ignore */ }
        }
    }

    /** Special tracking: click bar giả lập — reset khi đổi id (không đụng shift+click quan sát). */
    clearLeftSpecialBarPreviewPickOnFocusChange() {
        this.leftSpecialPreviewPickHistory = [];
        this.lastTrackingPreviewBarNum = null;
        return this.setLeftSpecialPreviewPickNum(null);
    }

    requestTrackingUiRepaintIfActive() {
        const tableWrap = typeof document !== 'undefined' ? document.getElementById('tableWrap') : null;
        if (tableWrap && typeof tableWrap.__trackingRepaint === 'function') {
            tableWrap.__trackingRepaint();
        }
    }

    onComboFocusIdChanged(prevFocusId, nextFocusId) {
        const prev = String(prevFocusId || '').trim();
        const next = String(nextFocusId || '').trim();
        if (!next || prev === next) {
            return;
        }
        this.noteComboFocusUndoTransition(prev, next);
        this.clearLeftBasicBarPreviewPicksOnFocusChange();
        if (this.clearLeftSpecialBarPreviewPickOnFocusChange()) {
            this.requestTrackingUiRepaintIfActive();
        }
    }

    /**
     * Ghi nhận chuỗi đổi focus (click/mũi tên): neo = id đầu, sau khi ổn định → peer Ctrl+Z.
     * Thí dụ 701→…→705: peer = 701 (không phải bước trung gian).
     */
    noteComboFocusUndoTransition(prevFocusId, nextFocusId) {
        const prev = String(prevFocusId || '').trim();
        const next = String(nextFocusId || '').trim();
        if (!prev || !next || prev === next) {
            return;
        }
        if (!this._comboFocusUndoBurstAnchorId) {
            this._comboFocusUndoBurstAnchorId = prev;
        }
        this._comboFocusUndoBurstEndId = next;
        if (this._comboFocusUndoCommitTimer) {
            clearTimeout(this._comboFocusUndoCommitTimer);
        }
        this._comboFocusUndoCommitTimer = setTimeout(() => {
            this._comboFocusUndoCommitTimer = 0;
            this.commitComboFocusUndoBurst();
        }, 150);
    }

    commitComboFocusUndoBurst() {
        const anchor = String(this._comboFocusUndoBurstAnchorId || '').trim();
        const end = String(this._comboFocusUndoBurstEndId || '').trim();
        this._comboFocusUndoBurstAnchorId = '';
        this._comboFocusUndoBurstEndId = '';
        if (anchor && end && anchor !== end) {
            this.comboFocusUndoPeerId = anchor;
        }
    }

    flushComboFocusUndoBurst() {
        if (this._comboFocusUndoCommitTimer) {
            clearTimeout(this._comboFocusUndoCommitTimer);
            this._comboFocusUndoCommitTimer = 0;
            this.commitComboFocusUndoBurst();
            return;
        }
        if (this._comboFocusUndoBurstAnchorId) {
            this.commitComboFocusUndoBurst();
        }
    }

    getSourceRowIndexById(rawId) {
        const key = this.normalizeNumberKey(rawId);
        if (!key) {
            return -1;
        }
        const rows = this.getSourceSheetRows();
        return rows.findIndex((row) => this.normalizeNumberKey(row.id || row.ID || '') === key);
    }

    /**
     * Ctrl+Z: focus id peer (id trước ↔ id hiện tại).
     * @returns {boolean}
     */
    toggleComboFocusUndoPeer() {
        this.flushComboFocusUndoBurst();
        const peerId = String(this.comboFocusUndoPeerId || '').trim();
        if (!peerId) {
            return false;
        }
        const currentId = String(this.comboFocusRowId || '').trim();
        const idx = this.getSourceRowIndexById(peerId);
        if (idx < 0) {
            return false;
        }
        if (this.normalizeNumberKey(peerId) === this.normalizeNumberKey(currentId)) {
            return false;
        }
        this.focusSourceSheetRow(idx, { skipSave: true });
        if (this._comboFocusUndoCommitTimer) {
            clearTimeout(this._comboFocusUndoCommitTimer);
            this._comboFocusUndoCommitTimer = 0;
        }
        this._comboFocusUndoBurstAnchorId = '';
        this._comboFocusUndoBurstEndId = '';
        if (currentId) {
            this.comboFocusUndoPeerId = currentId;
        }
        const wrap = typeof document !== 'undefined' ? document.getElementById('tableWrap') : null;
        if (wrap) {
            try {
                wrap.focus({ preventScroll: true });
            } catch (err) {
                /* ignore */
            }
        }
        return true;
    }

    /** Bar basic tracking → nửa trái: Submit OFF hoặc kỳ cuối trống khi Submit ON. */
    shouldSyncBasicBarPickToLeftPane() {
        if (!this.leftSubmitActive) {
            return true;
        }
        const sheet = this.sheets[TRACKING_SHEET_ID] || this.sheets.specialtracking;
        if (!sheet || this.getTrackingViewMode(sheet) !== 'basic') {
            return false;
        }
        const ui = sheet.trackingUi;
        const frameIndex = ui && typeof ui.frameIndex === 'number' ? ui.frameIndex : -1;
        if (frameIndex < 0) {
            return false;
        }
        return this.isBasicTrackingLastEmptyRowSyncActive(sheet, frameIndex);
    }

    setLeftSpecialPreviewPickNum(n) {
        const next = Number.isFinite(n) && n >= 1 && n <= 12 ? n : null;
        if (this.leftSpecialPreviewPickNum === next) {
            return false;
        }
        this.leftSpecialPreviewPickNum = next;
        return true;
    }

    /**
     * Special tracking: click bar giả lập — chỉ 1 số; click bar khác thay thế, click lại bar cũ thì bỏ.
     * options.retainFocus: chuột phải — tắt giả lập nhưng giữ focus bar đó để bật lại.
     * Chuột trái tắt: focus chuyển về phần tử cuối chuỗi còn lại.
     */
    toggleSpecialTrackingBarPick(n, options = {}) {
        const num = parseInt(n, 10);
        if (!Number.isFinite(num) || num < 1 || num > 12) {
            return false;
        }
        const retainFocus = !!options.retainFocus;
        const current = this.leftSpecialPreviewPickNum;
        if (current === num) {
            const hist = Array.isArray(this.leftSpecialPreviewPickHistory)
                ? this.leftSpecialPreviewPickHistory.slice()
                : [];
            while (hist.length && hist[hist.length - 1] === num) {
                hist.pop();
            }
            this.leftSpecialPreviewPickHistory = hist;
            this.lastTrackingPreviewBarNum = retainFocus
                ? num
                : (hist.length ? hist[hist.length - 1] : null);
            return this.setLeftSpecialPreviewPickNum(null);
        }
        const hist = Array.isArray(this.leftSpecialPreviewPickHistory)
            ? this.leftSpecialPreviewPickHistory.slice()
            : [];
        if (!hist.length || hist[hist.length - 1] !== num) {
            hist.push(num);
        }
        this.leftSpecialPreviewPickHistory = hist;
        this.lastTrackingPreviewBarNum = num;
        return this.setLeftSpecialPreviewPickNum(num);
    }

    /**
     * Basic tracking id cuối: toggle khoanh số (tối đa 5).
     * options.retainFocus: chuột phải — tắt giả lập nhưng giữ focus bar tip để bật lại;
     *   chuỗi pick vẫn cập nhật (bỏ/thêm số), focus không nhảy sang số khác.
     * Chuột trái / nửa trái tắt: focus = phần tử cuối chuỗi còn lại.
     */
    toggleBasicTrackingLastIdBarPick(n, options = {}) {
        const num = parseInt(n, 10);
        if (!Number.isFinite(num) || num < 1 || num > 35) {
            return false;
        }
        const retainFocus = !!options.retainFocus;
        const current = Array.isArray(this.leftBasicPreviewPickNums)
            ? this.leftBasicPreviewPickNums.slice()
            : [];
        const idx = current.indexOf(num);
        let next;
        if (idx >= 0) {
            next = current.filter((x) => x !== num);
        } else if (current.length >= 5) {
            return false;
        } else {
            next = current.concat(num);
        }
        this.setLeftBasicPreviewPickNums(next, { skipFocusUpdate: true });
        if (retainFocus) {
            // Chuột phải: chuỗi đã cập nhật; focus vẫn là bar đang toggle
            this.lastTrackingPreviewBarNum = num;
        } else {
            // Chuột trái trên bar: focus = tip chuỗi còn lại (hoặc số vừa thêm)
            this.lastTrackingPreviewBarNum = next.length
                ? next[next.length - 1]
                : null;
        }
        // Luôn đẩy chuỗi pick (thứ tự) sang nửa trái khi giả lập tracking
        this.syncLeftPickSelectionToIframe(next);
        try {
            window.dispatchEvent(new CustomEvent('leftCircledNumsChanged'));
        } catch (ePaint) { /* ignore */ }
        return true;
    }

    syncLeftPickSelectionToIframe(nums) {
        const frame = document.getElementById('okFrame');
        if (!frame || !frame.contentWindow) {
            return;
        }
        const list = Array.isArray(nums) ? nums.slice() : [];
        const win = frame.contentWindow;
        // Cùng origin: gọi thẳng để khoanh/pickOrder cập nhật ngay (chuột phải toggle).
        try {
            if (typeof win.applyBasicTrackingPreviewPicksFromParent === 'function') {
                win.applyBasicTrackingPreviewPicksFromParent(list);
                return;
            }
        } catch (eDirect) { /* fallback postMessage */ }
        try {
            win.postMessage({
                type: 'syncBasicTrackingPreviewPicks',
                nums: list,
                basicTrackingPreview: true
            }, '*');
        } catch (e) {
            try {
                win.postMessage({
                    type: 'syncAnswerPickSelection',
                    nums: list,
                    basicTrackingPreview: true
                }, '*');
            } catch (e2) { /* ignore */ }
        }
    }

    getBasicTrackingFocusRowIndex() {
        const rows = this.sourceRows || [];
        if (!rows.length) {
            return -1;
        }
        if (typeof this.comboFocusRowIndex === 'number'
            && this.comboFocusRowIndex >= 0
            && this.comboFocusRowIndex < rows.length) {
            return this.comboFocusRowIndex;
        }
        const onTracking = this.activeSheet === TRACKING_SHEET_ID || this.activeSheet === 'specialtracking';
        if (onTracking) {
            const st = this.sheets[TRACKING_SHEET_ID] || this.sheets.specialtracking;
            const ui = st && st.trackingUi;
            if (ui && typeof ui.focusSourceRowIndex === 'number') {
                const fromUi = ui.focusSourceRowIndex;
                if (fromUi >= 0 && fromUi < rows.length) {
                    return fromUi;
                }
            }
        }
        if (this.activeWindowRange && typeof this.activeWindowRange.target === 'number') {
            const t = this.activeWindowRange.target;
            if (t >= 0 && t < rows.length) {
                return t;
            }
        }
        return -1;
    }

    /**
     * Basic tracking: frame timeline khớp hàng nguồn — có thể giả lập pick/freq (không phụ thuộc Submit).
     */
    isBasicTrackingFramePreviewEligible(sheet, frameIndex) {
        if (!sheet || this.getTrackingViewMode(sheet) !== 'basic') {
            return false;
        }
        const rows = this.sourceRows || [];
        if (!rows.length) {
            return false;
        }
        const rowIdx = this.getTrackingSourceRowIndexForFrame(sheet, frameIndex);
        if (rowIdx < 0 || rowIdx >= rows.length) {
            return false;
        }
        const frameForRow = this.getTrackingFrameIndexForSourceRow(sheet, rowIdx);
        if (frameForRow < 0 || frameIndex !== frameForRow) {
            return false;
        }
        const row = rows[rowIdx];
        const lastIdx = rows.length - 1;
        if (rowIdx === lastIdx && row && this.isEmptyResultRow(row)) {
            const focusIdx = this.getBasicTrackingFocusRowIndex();
            return focusIdx === lastIdx;
        }
        return true;
    }

    /**
     * Basic tracking + Submit OFF: giả lập bar pick tại frame timeline hiện tại
     * (kỳ lịch sử hoặc kỳ cuối chưa có đáp án khi focus đúng hàng).
     */
    isBasicTrackingBarPreviewActive(sheet, frameIndex) {
        return !this.leftSubmitActive
            && this.isBasicTrackingFramePreviewEligible(sheet, frameIndex);
    }

    /**
     * Basic tracking + focus hàng cuối chưa có đáp án + timeline đúng frame đó.
     * Dùng đồng bộ viền đen bar (kể cả Submit ON / autoring).
     */
    isBasicTrackingLastEmptyRowSyncActive(sheet, frameIndex) {
        if (!sheet || this.getTrackingViewMode(sheet) !== 'basic') {
            return false;
        }
        const rows = this.sourceRows || [];
        if (!rows.length) {
            return false;
        }
        const lastIdx = rows.length - 1;
        const row = rows[lastIdx];
        if (!row || !this.isEmptyResultRow(row)) {
            return false;
        }
        if (this.getBasicTrackingFocusRowIndex() !== lastIdx) {
            return false;
        }
        const rowIdx = this.getTrackingSourceRowIndexForFrame(sheet, frameIndex);
        if (rowIdx !== lastIdx) {
            return false;
        }
        const frameForRow = this.getTrackingFrameIndexForSourceRow(sheet, lastIdx);
        return frameForRow >= 0 && frameIndex === frameForRow;
    }

    /**
     * Basic tracking + submit OFF + focus hàng cuối chưa có đáp án + timeline đúng frame đó.
     */
    isBasicTrackingEmptyLastRowPreviewActive(sheet, frameIndex) {
        return this.isBasicTrackingBarPreviewActive(sheet, frameIndex)
            && this.isBasicTrackingLastEmptyRowSyncActive(sheet, frameIndex);
    }

    /** Viền đen bar + đồng bộ pick: Submit OFF trên frame hợp lệ, hoặc kỳ cuối trống khi Submit ON. */
    isBasicTrackingLeftPickBarSyncActive(sheet, frameIndex) {
        if (!this.isBasicTrackingFramePreviewEligible(sheet, frameIndex)) {
            return false;
        }
        if (!this.leftSubmitActive) {
            return true;
        }
        return this.isBasicTrackingLastEmptyRowSyncActive(sheet, frameIndex);
    }

    /** Layout giả lập freq (+1 pick, trừ justDrawn) — Submit OFF, hoặc kỳ cuối trống khi Submit ON. */
    isBasicTrackingFreqPreviewLayoutActive(sheet, frameIndex) {
        if (!this.isBasicTrackingFramePreviewEligible(sheet, frameIndex)) {
            return false;
        }
        if (!this.leftSubmitActive) {
            return true;
        }
        return this.isBasicTrackingLastEmptyRowSyncActive(sheet, frameIndex);
    }

    /** Special tracking: frame timeline khớp hàng nguồn — giả lập khi Submit OFF, hoặc kỳ cuối trống khi Submit ON. */
    isSpecialTrackingFramePreviewEligible(sheet, frameIndex) {
        if (this.leftSubmitActive
            && !this.isSpecialTrackingLastEmptyRowSyncActive(sheet, frameIndex)) {
            return false;
        }
        if (!sheet || this.getTrackingViewMode(sheet) !== 'special') {
            return false;
        }
        const rows = this.sourceRows || [];
        if (!rows.length) {
            return false;
        }
        const rowIdx = this.getTrackingSourceRowIndexForFrame(sheet, frameIndex);
        if (rowIdx < 0 || rowIdx >= rows.length) {
            return false;
        }
        const frameForRow = this.getTrackingFrameIndexForSourceRow(sheet, rowIdx);
        if (frameForRow < 0 || frameIndex !== frameForRow) {
            return false;
        }
        const row = rows[rowIdx];
        const lastIdx = rows.length - 1;
        if (rowIdx === lastIdx && row && this.isEmptyResultRow(row)) {
            const focusIdx = this.getBasicTrackingFocusRowIndex();
            return focusIdx === lastIdx;
        }
        return true;
    }

    isSpecialTrackingLastEmptyRowSyncActive(sheet, frameIndex) {
        if (!sheet || this.getTrackingViewMode(sheet) !== 'special') {
            return false;
        }
        const rows = this.sourceRows || [];
        if (!rows.length) {
            return false;
        }
        const lastIdx = rows.length - 1;
        const row = rows[lastIdx];
        if (!row || !this.isEmptyResultRow(row)) {
            return false;
        }
        if (this.getBasicTrackingFocusRowIndex() !== lastIdx) {
            return false;
        }
        const rowIdx = this.getTrackingSourceRowIndexForFrame(sheet, frameIndex);
        if (rowIdx !== lastIdx) {
            return false;
        }
        const frameForRow = this.getTrackingFrameIndexForSourceRow(sheet, lastIdx);
        return frameForRow >= 0 && frameIndex === frameForRow;
    }

    /** Layout giả lập số đặc biệt — Submit OFF, hoặc kỳ cuối trống khi Submit ON. */
    isSpecialTrackingFreqPreviewLayoutActive(sheet, frameIndex) {
        if (!this.isSpecialTrackingFramePreviewEligible(sheet, frameIndex)) {
            return false;
        }
        if (!this.leftSubmitActive) {
            return true;
        }
        return this.isSpecialTrackingLastEmptyRowSyncActive(sheet, frameIndex);
    }

    /** Trạng thái trước click giả lập: cùng submit, không preview pick (+1). */
    computeTrackingBeforePreviewGhostTieResult(sheet, frames, frameIndex, options = {}) {
        const isBasic = !!options.isBasic;
        const numMax = options.numMax || (isBasic ? 35 : 12);
        const basicDraws = options.basicDraws || (isBasic ? (sheet.basicDraws || []) : []);
        const tieResult = this.computeTrackingFreqTieGroupsWithStreaks(sheet, frames, frameIndex, {
            isBasic,
            numMax,
            leftSubmitOn: !!options.leftSubmitOn,
            basicDraws,
            previewAtPaintFrame: false,
            previewPickNums: [],
            specialPreviewPick: null
        });
        return {
            groups: tieResult.groups || [],
            streakByKey: tieResult.streakByKey || new Map()
        };
    }

    /**
     * Bụng mờ: trạng thái trước thay đổi.
     * Submit ON → rollback justDrawn (như Submit OFF).
     * Click giả lập → không cộng preview pick (+1).
     */
    computeTrackingGhostFreqTieResult(sheet, frames, frameIndex, options = {}) {
        const isBasic = !!options.isBasic;
        const leftSubmitOn = !!options.leftSubmitOn;
        const numMax = options.numMax || (isBasic ? 35 : 12);
        const basicDraws = options.basicDraws || (isBasic ? (sheet.basicDraws || []) : []);
        const previewLayout = isBasic && !!options.freqPreviewLayout;
        const specialPreviewLayout = !isBasic && !!options.specialPreviewLayout;
        const previewPickNums = isBasic && previewLayout
            ? (options.previewPickNums || []).slice()
            : [];
        const specialPreviewPick = !isBasic && specialPreviewLayout
            ? options.specialPreviewPick
            : null;
        const hasPreviewSimulation = isBasic
            ? (previewLayout && previewPickNums.length > 0)
            : (specialPreviewLayout && specialPreviewPick != null);

        if (!this.leftSubmitActive && !hasPreviewSimulation) {
            return { groups: [], streakByKey: new Map() };
        }

        const ghostLeftSubmitOn = this.leftSubmitActive ? false : leftSubmitOn;
        const ghostPreviewPickNums = isBasic && previewLayout
            ? (this.leftSubmitActive ? previewPickNums.slice() : [])
            : [];
        const ghostSpecialPreviewPick = !isBasic && specialPreviewLayout
            ? (this.leftSubmitActive ? specialPreviewPick : null)
            : null;

        const tieResult = this.computeTrackingFreqTieGroupsWithStreaks(sheet, frames, frameIndex, {
            isBasic,
            numMax,
            leftSubmitOn: ghostLeftSubmitOn,
            basicDraws,
            previewAtPaintFrame: isBasic ? previewLayout : specialPreviewLayout,
            previewPickNums: ghostPreviewPickNums,
            specialPreviewPick: ghostSpecialPreviewPick
        });
        return {
            groups: tieResult.groups || [],
            streakByKey: tieResult.streakByKey || new Map()
        };
    }

    /** @deprecated */
    computeTrackingPreSubmitGhostFreqGroups(sheet, frames, frameIndex, currentGroups, options = {}) {
        return this.computeTrackingGhostFreqTieResult(sheet, frames, frameIndex, options);
    }

    /**
     * Basic tracking: layout hiển thị theo freq tích lũy đã “commit”.
     * Khi submit trái OFF, 5 số kỳ hiện tại chưa được cộng vào tích lũy.
     * previewPickNums: giả lập +1 freq cho số khoanh trái (kỳ cuối chưa có đáp án).
     */
    static computeBasicTrackingDisplayLayout(draws, fr, submitOn, previewPickNums) {
        const list = draws || [];
        const drawEndIndex = (fr && fr.drawIndex != null && fr.drawIndex >= 0) ? fr.drawIndex : -1;
        const counts = {};
        for (let i = 1; i <= 35; i++) {
            counts[i] = (fr.byNum && fr.byNum[i]) || 0;
        }
        const justDrawnNums = Array.isArray(fr.justDrawnNums) ? fr.justDrawnNums : [];
        const hasPreviewPicks = Array.isArray(previewPickNums) && previewPickNums.length > 0;
        const applyDrawnRollback = !submitOn || hasPreviewPicks;
        if (applyDrawnRollback) {
            for (let u = 0; u < justDrawnNums.length; u++) {
                const n = justDrawnNums[u];
                if (n >= 1 && n <= 35) {
                    counts[n] = Math.max(0, (counts[n] || 0) - 1);
                }
            }
        }
        if (hasPreviewPicks) {
            for (let u = 0; u < previewPickNums.length; u++) {
                const n = previewPickNums[u];
                if (n >= 1 && n <= 35) {
                    counts[n] = (counts[n] || 0) + 1;
                }
            }
        }
        const bottomSlot = 34;
        const previewPickSet = hasPreviewPicks ? new Set(previewPickNums) : null;
        const ranked = [];
        for (let n = 1; n <= 35; n++) {
            const v = counts[n] || 0;
            const t = v > 0 && drawEndIndex >= 0
                ? RightPaneSheetManager.basicTrackingRankTForHit(
                    list,
                    n,
                    v,
                    drawEndIndex,
                    { previewPickSet }
                )
                : -1;
            ranked.push({ n, v, t });
        }
        ranked.sort((a, b) => RightPaneSheetManager.basicTrackingRankCompare(
            a, b, list, drawEndIndex, { previewPickSet }
        ));
        const maxV = Math.max(1, ranked.length && ranked[0].v ? ranked[0].v : 1);
        const slotByNum = new Array(36).fill(bottomSlot);
        for (let s = 0; s < ranked.length; s++) {
            slotByNum[ranked[s].n] = s;
        }
        const wPctByNum = new Array(36).fill(0);
        for (let n = 1; n <= 35; n++) {
            wPctByNum[n] = (counts[n] / maxV) * 100;
        }
        return { counts, slotByNum, wPctByNum };
    }

    /**
     * Special tracking: layout hiển thị khi giả lập 1 số đặc biệt (click bar).
     */
    static computeSpecialTrackingDisplayLayout(series, fr, submitOn, previewPickNum) {
        const counts = {};
        for (let i = 1; i <= 12; i++) {
            counts[i] = (fr && fr.byNum && fr.byNum[i]) || 0;
        }
        const justDrawn = fr && fr.justDrawn != null ? fr.justDrawn : null;
        const hasPreview = Number.isFinite(previewPickNum) && previewPickNum >= 1 && previewPickNum <= 12;
        const applyDrawnRollback = !submitOn || hasPreview;
        if (applyDrawnRollback && justDrawn != null) {
            counts[justDrawn] = Math.max(0, (counts[justDrawn] || 0) - 1);
        }
        if (hasPreview) {
            counts[previewPickNum] = (counts[previewPickNum] || 0) + 1;
        }

        const list = Array.isArray(series) ? series.slice() : [];
        let drawEndIndex = fr && fr.drawIndex != null && fr.drawIndex >= 0 ? fr.drawIndex : -1;
        let virtualList = drawEndIndex >= 0 ? list.slice(0, drawEndIndex + 1) : [];
        if (applyDrawnRollback && justDrawn != null && virtualList.length
            && virtualList[virtualList.length - 1] === justDrawn) {
            virtualList = virtualList.slice(0, -1);
            drawEndIndex = virtualList.length - 1;
        }
        if (hasPreview) {
            virtualList.push(previewPickNum);
            drawEndIndex = virtualList.length - 1;
        }

        const ranked = [];
        for (let n = 1; n <= 12; n++) {
            const v = counts[n] || 0;
            const t = v > 0 && drawEndIndex >= 0
                ? RightPaneSheetManager.specialTrackingStepOfVthHit(virtualList, n, v, drawEndIndex)
                : -1;
            ranked.push({ n, v, t });
        }
        ranked.sort((a, b) => (b.v - a.v) || (a.t - b.t) || (a.n - b.n));
        const maxV = Math.max(1, ranked.length && ranked[0].v ? ranked[0].v : 1);
        const slotByNum = new Array(13).fill(11);
        for (let s = 0; s < ranked.length; s++) {
            if (ranked[s].v > 0) {
                slotByNum[ranked[s].n] = s;
            }
        }
        const wPctByNum = new Array(13).fill(0);
        for (let n = 1; n <= 12; n++) {
            wPctByNum[n] = ((counts[n] || 0) / maxV) * 100;
        }
        return { counts, slotByNum, wPctByNum };
    }

    /**
     * Special tracking: số có freq > 0 mà giá trị freq không ±1 (và không cùng bụng)
     * so với ít nhất một số liền trước/sau trên stack.
     * Số cùng bụng (cùng freq liền kề) luôn được coi là kết nối.
     * @returns {Set<number>}
     */
    static computeSpecialTrackingFreqDisconnectedNums(counts, slotByNum, numMax = 12) {
        const ranked = [];
        for (let n = 1; n <= numMax; n++) {
            const freq = (counts && counts[n]) || 0;
            if (freq <= 0) {
                continue;
            }
            ranked.push({
                n,
                freq: freq | 0,
                slot: (slotByNum && slotByNum[n]) ?? numMax
            });
        }
        if (ranked.length < 2) {
            return new Set();
        }
        ranked.sort((a, b) => a.slot - b.slot || a.n - b.n);
        const out = new Set();
        const isNeighborLinked = (freqA, freqB) => freqA === freqB || Math.abs(freqA - freqB) === 1;
        for (let i = 0; i < ranked.length; i++) {
            const { n, freq } = ranked[i];
            let linked = false;
            if (i > 0 && isNeighborLinked(freq, ranked[i - 1].freq)) {
                linked = true;
            }
            if (i < ranked.length - 1 && isNeighborLinked(freq, ranked[i + 1].freq)) {
                linked = true;
            }
            if (!linked) {
                out.add(n);
            }
        }
        return out;
    }

    /**
     * Đường kẻ ngăn phía solid: giữa 2 thanh liền kề trên stack khi |Δfreq| ≥ minGap.
     * @returns {Array<{ boundarySlot: number, gap: number, aboveNum: number, belowNum: number }>}
     */
    static computeTrackingFreqGapDividers(counts, slotByNum, numMax, minGap = 2) {
        const ranked = [];
        for (let n = 1; n <= numMax; n++) {
            const freq = (counts && counts[n]) || 0;
            if (freq <= 0) {
                continue;
            }
            ranked.push({
                n,
                freq: freq | 0,
                slot: (slotByNum && slotByNum[n]) ?? numMax
            });
        }
        if (ranked.length < 2) {
            return [];
        }
        ranked.sort((a, b) => a.slot - b.slot || a.n - b.n);
        const threshold = Math.max(2, minGap | 0);
        const out = [];
        for (let i = 0; i < ranked.length - 1; i++) {
            const gap = Math.abs(ranked[i].freq - ranked[i + 1].freq);
            if (gap >= threshold) {
                out.push({
                    boundarySlot: ranked[i + 1].slot,
                    gap,
                    aboveNum: ranked[i].n,
                    belowNum: ranked[i + 1].n
                });
            }
        }
        return out;
    }

    /** Vẽ đường kẻ ngăn freq-gap trên cột solid (tail + meta). */
    static syncTrackingFreqGapDividers(layer, dividers, slotCount) {
        if (!layer) {
            return;
        }
        const sig = !dividers.length
            ? ''
            : dividers.map((d) => `${d.boundarySlot}:${d.gap}:${d.aboveNum}-${d.belowNum}`).join('|');
        if (layer.dataset.stFreqGapSig !== sig) {
            layer.dataset.stFreqGapSig = sig;
            layer.replaceChildren();
            if (!dividers.length || !slotCount) {
                return;
            }
            const frag = document.createDocumentFragment();
            for (const d of dividers) {
                const el = document.createElement('div');
                el.className = 'special-tracking-freq-gap-divider'
                    + (d.gap >= 3 ? ' special-tracking-freq-gap-divider--gap-ge3' : '');
                el.dataset.stFreqGapSlot = String(d.boundarySlot);
                el.dataset.stFreqGap = String(d.gap);
                el.title = `Gap freq ${d.gap}: ${d.aboveNum} | ${d.belowNum}`;
                el.setAttribute('aria-hidden', 'true');
                el.innerHTML = '<svg viewBox="0 0 100 4" preserveAspectRatio="none" aria-hidden="true">'
                    + '<line x1="0" y1="2" x2="100" y2="2" fill="none" stroke="currentColor" '
                    + 'stroke-linecap="square" vector-effect="non-scaling-stroke"/></svg>';
                frag.appendChild(el);
            }
            layer.appendChild(frag);
        }
        RightPaneSheetManager.layoutTrackingFreqGapDividers(layer, slotCount);
    }

    /** Snap vị trí freq-gap theo pixel — tránh subpixel làm nét dày mỏn không đều. */
    static layoutTrackingFreqGapDividers(layer, slotCount) {
        if (!layer || !slotCount) {
            return;
        }
        const layerHeight = layer.clientHeight;
        if (layerHeight <= 0) {
            return;
        }
        const kids = layer.children;
        for (let i = 0; i < kids.length; i++) {
            const el = kids[i];
            const slot = parseInt(el.dataset.stFreqGapSlot, 10);
            if (!Number.isFinite(slot)) {
                continue;
            }
            const yPx = Math.round((slot / slotCount) * layerHeight);
            el.style.top = `${yPx}px`;
        }
    }

    /**
     * Bước s trong chuỗi draws mà số n vừa đạt đúng lần xuất hiện thứ v (v≥1).
     */
    static basicTrackingStepOfVthHit(draws, n, v, endInclusive) {
        if (v <= 0) {
            return -1;
        }
        let seen = 0;
        for (let s = 0; s <= endInclusive; s++) {
            const d = draws[s];
            if (!Array.isArray(d)) {
                continue;
            }
            if (d.includes(n)) {
                seen++;
                if (seen === v) {
                    return s;
                }
            }
        }
        return endInclusive;
    }

    static countBasicTrackingHitsInDraws(draws, n, endInclusive) {
        let seen = 0;
        for (let s = 0; s <= endInclusive; s++) {
            const d = draws[s];
            if (Array.isArray(d) && d.includes(n)) {
                seen++;
            }
        }
        return seen;
    }

    /**
     * Bước xếp hạng khi cùng freq (t nhỏ = đạt v sớm hơn, slot cao hơn).
     * Preview +1: lần đạt v giả lập xếp sau mọi số đã đạt v trong draws.
     */
    static basicTrackingRankTForHit(draws, n, v, endInclusive, options = {}) {
        if (v <= 0 || endInclusive < 0) {
            return -1;
        }
        const previewSet = options.previewPickSet;
        const tFromDraws = RightPaneSheetManager.basicTrackingStepOfVthHit(draws, n, v, endInclusive);
        if (!previewSet || !previewSet.has(n)) {
            return tFromDraws;
        }
        const hitsInDraws = RightPaneSheetManager.countBasicTrackingHitsInDraws(draws, n, endInclusive);
        if (hitsInDraws < v) {
            return endInclusive + 1;
        }
        const tAtVminus1 = v > 1
            ? RightPaneSheetManager.basicTrackingStepOfVthHit(draws, n, v - 1, endInclusive)
            : -1;
        if (tFromDraws >= endInclusive && tAtVminus1 < endInclusive) {
            return endInclusive + 1;
        }
        return tFromDraws;
    }

    /**
     * So sánh xếp hạng bar basic: freq cao hơn trước; cùng freq thì t nhỏ hơn (đạt sớm hơn).
     * Cùng kỳ đạt mốc v: lùi so v-1, v-2, … để giữ thứ tự lịch sử (17:87 trước 6:87 → 17:88 vẫn trên 6:88).
     */
    static basicTrackingRankCompare(a, b, draws, endInclusive, options = {}) {
        if (b.v !== a.v) {
            return b.v - a.v;
        }
        const v = a.v | 0;
        for (let k = v; k >= 1; k--) {
            const ta = RightPaneSheetManager.basicTrackingRankTForHit(draws, a.n, k, endInclusive, options);
            const tb = RightPaneSheetManager.basicTrackingRankTForHit(draws, b.n, k, endInclusive, options);
            if (ta !== tb) {
                return ta - tb;
            }
        }
        return a.n - b.n;
    }

    buildBasicTrackingFrames(drawSteps) {
        const counts = {};
        for (let i = 1; i <= 35; i++) {
            counts[i] = 0;
        }
        const frames = [];
        const list = [];
        const steps = drawSteps || [];
        const bottomSlot = 34;
        let drawIndex = -1;
        for (let f = 0; f < steps.length; f++) {
            const step = steps[f];
            const justDrawnNums = Array.isArray(step) ? step.slice() : [];
            const holdFrame = justDrawnNums.length === 0;
            if (!holdFrame) {
                drawIndex += 1;
                list.push(justDrawnNums);
                for (let u = 0; u < justDrawnNums.length; u++) {
                    const n = justDrawnNums[u];
                    if (n >= 1 && n <= 35) {
                        counts[n] += 1;
                    }
                }
            }
            const endDrawIdx = drawIndex;
            const sorted = Object.keys(counts)
                .map((k) => {
                    const n = Number(k);
                    const v = counts[n];
                    const t = v > 0 && endDrawIdx >= 0
                        ? RightPaneSheetManager.basicTrackingStepOfVthHit(list, n, v, endDrawIdx)
                        : -1;
                    return { n, v, t };
                })
                .sort((a, b) => RightPaneSheetManager.basicTrackingRankCompare(a, b, list, endDrawIdx, {}))
                .map(({ n, v }) => ({ n, v }));
            const maxV = Math.max(1, sorted.length ? sorted[0].v : 1);
            const slotByNum = new Array(36).fill(bottomSlot);
            for (let s = 0; s < sorted.length; s++) {
                slotByNum[sorted[s].n] = s;
            }
            const wPctByNum = new Array(36).fill(0);
            for (let n = 1; n <= 35; n++) {
                wPctByNum[n] = (counts[n] / maxV) * 100;
            }
            frames.push({
                step: f + 1,
                drawIndex: endDrawIdx,
                holdFrame,
                justDrawn: justDrawnNums[0] ?? null,
                justDrawnNums,
                byNum: { ...counts },
                sorted,
                maxV,
                slotByNum,
                wPctByNum
            });
        }
        return frames;
    }

    buildTrackingFramesForMode(viewMode, specialMeta, basicMeta) {
        const mode = RightPaneSheetManager.normalizeTrackingViewMode(viewMode);
        if (mode === 'basic') {
            const drawSteps = (basicMeta && basicMeta.drawSteps) || [];
            return {
                series: (basicMeta && basicMeta.draws) || [],
                sourceRowIndices: (basicMeta && basicMeta.sourceRowIndices) || [],
                frames: this.buildBasicTrackingFrames(drawSteps)
            };
        }
        const drawSteps = (specialMeta && specialMeta.drawSteps) || [];
        return {
            series: (specialMeta && specialMeta.series) || [],
            sourceRowIndices: (specialMeta && specialMeta.sourceRowIndices) || [],
            frames: this.buildSpecialTrackingFrames(drawSteps)
        };
    }

    ensureTrackingMetaCaches(sheet) {
        if (!sheet || sheet.kind !== TRACKING_KIND) {
            return;
        }
        const rows = this.sourceRows || [];
        if (!Array.isArray(sheet.specialSeries) || !Array.isArray(sheet.specialSourceRowIndices) || !Array.isArray(sheet.specialDrawSteps)) {
            const sm = this.buildSpecialTrackingSeriesMeta(rows);
            sheet.specialSeries = sm.series;
            sheet.specialDrawSteps = sm.drawSteps;
            sheet.specialSourceRowIndices = sm.sourceRowIndices;
        }
        if (!Array.isArray(sheet.basicDraws) || !Array.isArray(sheet.basicSourceRowIndices) || !Array.isArray(sheet.basicDrawSteps)) {
            const bm = this.buildBasicTrackingSeriesMeta(rows);
            sheet.basicDraws = bm.draws;
            sheet.basicDrawSteps = bm.drawSteps;
            sheet.basicSourceRowIndices = bm.sourceRowIndices;
        }
    }

    ensureTrackingFrames(sheet) {
        if (!sheet || sheet.kind !== TRACKING_KIND) {
            return;
        }
        this.ensureTrackingMetaCaches(sheet);
        const viewMode = this.getTrackingViewMode(sheet);
        const fr0 = sheet.frames && sheet.frames[0];
        const framesOk = !!(sheet.frames && sheet.frames.length && fr0 && Array.isArray(fr0.slotByNum) && Array.isArray(fr0.wPctByNum));
        const expectedLen = (this.sourceRows || []).length;
        const idxList = viewMode === 'basic'
            ? (sheet.basicSourceRowIndices || [])
            : (sheet.specialSourceRowIndices || []);
        const idxOk = Array.isArray(idxList) && idxList.length === expectedLen;
        const modeOk = sheet._trackingFramesMode === viewMode;
        if (framesOk && idxOk && modeOk && sheet.frames.length === expectedLen) {
            sheet.series = viewMode === 'basic' ? (sheet.basicDraws || []) : (sheet.specialSeries || []);
            sheet.seriesSourceRowIndices = idxList.slice();
            return;
        }
        const built = this.buildTrackingFramesForMode(viewMode, {
            series: sheet.specialSeries,
            drawSteps: sheet.specialDrawSteps,
            sourceRowIndices: sheet.specialSourceRowIndices
        }, {
            draws: sheet.basicDraws,
            drawSteps: sheet.basicDrawSteps,
            sourceRowIndices: sheet.basicSourceRowIndices
        });
        sheet.series = built.series;
        sheet.seriesSourceRowIndices = built.sourceRowIndices;
        sheet.frames = built.frames;
        sheet._trackingFramesMode = viewMode;
    }

    /** Hue 0…1 cho gradient hạng: hạng 1 (slot 0) = vàng nhạt, còn lại phân bổ đều. */
    static specialTrackingRankHueT(slot, slotMax = 11) {
        const s = Number(slot);
        const maxS = Math.max(1, Number(slotMax) | 0);
        const t = Number.isFinite(s) ? Math.max(0, Math.min(maxS, Math.floor(s))) : 0;
        return t === 0 ? 2 / (maxS + 1) : Math.min(1, (t + 1) / (maxS + 1));
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
        const lastIdx = RightPaneSheetManager.getTrackingFrameIndexAfterDraws(frames, N);
        if (lastIdx < 0) {
            return [];
        }
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
        this.ensureTrackingFrames(sheet);
    }

    renderTrackingShell(sheet) {
        this.ensureTrackingFrames(sheet);
        const viewMode = this.getTrackingViewMode(sheet);
        const isBasic = viewMode === 'basic';
        const slotCount = this.getTrackingSlotCount(viewMode);
        const numMax = slotCount;
        const frames = sheet.frames || [];
        if (!frames.length) {
            const emptyMsg = isBasic
                ? 'Chưa có kỳ nào có đủ 5 số chính (trước dấu |) trong cột result.'
                : 'Chưa có kỳ nào có số đặc biệt 1–12 sau dấu | trong cột result. Tải sheet1 dạng <code>…|7</code>.';
            return (
                `<div class="special-tracking-root${isBasic ? ' special-tracking-root--basic' : ''}" data-st-view-mode="${viewMode}">`
                + `<div class="special-tracking-empty">${emptyMsg}</div>`
                + '</div>'
            );
        }
        const total = frames.length;
        const fr0 = frames[0];
        let rankBarsHtml = '';
        for (let n = 1; n <= numMax; n++) {
            const slot = fr0.slotByNum[n] ?? 0;
            const hueT = RightPaneSheetManager.specialTrackingRankHueT(slot, slotCount - 1);
            rankBarsHtml += `<div class="special-tracking-rank-bar" data-st-bar="${n}" data-special-num="${n}" style="--st-slot:${slot};--st-slot-count:${slotCount};--st-hue-t:${hueT}" role="button" tabindex="0" aria-label="Số ${n}, click để tô sáng">`
                + '<div class="special-tracking-rank-bar-main">'
                + '<div class="special-tracking-rank-track special-tracking-rank-track--main">'
                + '<span class="special-tracking-rank-fill" data-fill></span>'
                + `<span class="special-tracking-rank-num" data-st-num>${n}</span>`
                + '</div>'
                + '</div>'
                + '<div class="special-tracking-rank-bar-tail">'
                + '<div class="special-tracking-rank-track special-tracking-rank-track--label">'
                + '</div>'
                + '<div class="special-tracking-rank-meta">'
                + '<span class="special-tracking-rank-count" data-count>0</span>'
                + (isBasic ? '' : '<span class="special-tracking-rank-prio" data-st-predict-rank aria-hidden="true"></span>')
                + '</div>'
                + '</div>'
                + '</div>';
        }
        const predictMetaHtml = isBasic
            ? ''
            : (
                '<div class="special-tracking-meta-right special-tracking-meta-predict">'
                + '<button type="button" class="special-tracking-predict-toggle" data-st-predict-toggle '
                + 'aria-pressed="false" title="Bật/tắt neon dự đoán #1–#3 theo từng vị trí timeline" aria-label="Predict">Predict</button>'
                + '<div class="special-tracking-rank-stats" data-st-rank-stats aria-label="Tỷ lệ trúng theo hạng dự đoán toàn lịch sử"></div>'
                + '</div>'
            );
        const viewLabel = isBasic ? 'basic' : 'special';
        return (
            `<div class="special-tracking-root${isBasic ? ' special-tracking-root--basic' : ''}" data-st-root data-st-view-mode="${viewMode}" style="--st-slot-count:${slotCount}">`
            + '<div class="special-tracking-stage">'
            + '<div class="special-tracking-rank-wrap">'
            + '<div class="special-tracking-rank-shell">'
            + `<div class="special-tracking-rank-stack" data-st-rank-stack>${rankBarsHtml}</div>`
            + '<div class="special-tracking-freq-brace-layer special-tracking-freq-brace-layer--ghost" data-st-freq-braces-ghost aria-hidden="true"></div>'
            + '<div class="special-tracking-freq-brace-layer" data-st-freq-braces aria-hidden="true"></div>'
            + '<div class="special-tracking-freq-gap-layer" data-st-freq-gap-dividers aria-hidden="true"></div>'
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
            + predictMetaHtml
            + '</div>'
            + '</div>'
            + '</div>'
            + '</div>'
            + '<div class="special-tracking-controls-row special-tracking-controls-row--transport">'
            + '<button type="button" class="special-tracking-icon-btn special-tracking-icon-btn--glyph" data-st-first title="Về đầu" aria-label="Về đầu"><span aria-hidden="true">\u23ee\uFE0E</span></button>'
            + '<button type="button" class="special-tracking-icon-btn special-tracking-icon-btn--glyph" data-st-prev title="Lùi 1 kỳ" aria-label="Lùi một kỳ"><span aria-hidden="true">\u23ea\uFE0E</span></button>'
            + '<button type="button" class="special-tracking-icon-btn special-tracking-icon-btn--play" data-st-play title="Phát" aria-label="Phát">'
            + '<svg class="special-tracking-svg-play" viewBox="0 0 24 24" width="22" height="22" aria-hidden="true"><path fill="currentColor" d="M9 6.5v11L18 12 9 6.5z"/></svg>'
            + '<svg class="special-tracking-svg-pause" viewBox="0 0 24 24" width="22" height="22" aria-hidden="true"><path fill="currentColor" d="M8 7h3v10H8V7zm5 0h3v10h-3V7z"/></svg></button>'
            + '<button type="button" class="special-tracking-icon-btn special-tracking-icon-btn--glyph" data-st-next title="Tiến 1 kỳ" aria-label="Tiến một kỳ"><span aria-hidden="true">\u23e9\uFE0E</span></button>'
            + '<button type="button" class="special-tracking-icon-btn special-tracking-icon-btn--glyph" data-st-last title="Cuối" aria-label="Đến cuối"><span aria-hidden="true">\u23ed\uFE0E</span></button>'
            + '<div class="special-tracking-speed-slider-wrap">'
            + '<span class="special-tracking-speed-hint" title="1× = 100ms/kỳ khi Phát">Tốc độ</span>'
            + '<div class="special-tracking-speed-slider-inner">'
            + '<input type="range" class="special-tracking-speed-slider" data-st-speed-slider '
            + 'min="0.5" max="3" step="0.5" value="1" aria-valuemin="0.5" aria-valuemax="3" aria-valuenow="1" aria-label="Tốc độ phát" />'
            + '<span class="special-tracking-speed-readout" data-st-speed-val>1×</span>'
            + '</div></div>'
            + `<button type="button" class="special-tracking-label-toggle" data-st-label-toggle aria-pressed="true" `
            + `title="Nhãn số sát biên bar (bấm để in)" aria-label="Nhãn bar: out">out</button>`
            + `<button type="button" class="special-tracking-view-toggle" data-st-view-toggle aria-pressed="true" `
            + `title="Chế độ hiển thị: ${viewLabel} (bấm để đổi)" aria-label="Chế độ ${viewLabel}">${viewLabel}</button>`
            + '</div>'
            + '</div>'
            + `<input type="hidden" data-st-total value="${total}" />`
            + '</div>'
        );
    }

    renderSpecialTrackingShell(sheet) {
        return this.renderTrackingShell(sheet);
    }

    wireTrackingUi(tableWrap, sheet) {
        this.ensureTrackingFrames(sheet);
        const viewMode = this.getTrackingViewMode(sheet);
        const isBasic = viewMode === 'basic';
        const slotCount = this.getTrackingSlotCount(viewMode);
        const numMax = slotCount;
        const frames = sheet.frames || [];
        const root = tableWrap.querySelector('[data-st-root]');
        if (!root || !frames.length) {
            tableWrap.__trackingCleanup = null;
            return;
        }

        const series = Array.isArray(sheet.series) ? sheet.series : [];
        const total = frames.length;
        const srcRows = this.sourceRows || [];
        const tailRow = srcRows.length ? srcRows[srcRows.length - 1] : {};
        const tailId = String(tailRow.id ?? tailRow.ID ?? '');
        const uiSig = `${viewMode}|${total}|${srcRows.length}|${srcRows.length}|${tailId}`;

        const readSavedStUi = () => {
            let u = sheet.trackingUi || sheet.specialTrackingUi;
            if (u && u.sig === uiSig) {
                return u;
            }
            for (const key of [TRACKING_UI_STORAGE_KEY, LEGACY_TRACKING_UI_STORAGE_KEY]) {
                try {
                    const raw = sessionStorage.getItem(key);
                    if (!raw) {
                        continue;
                    }
                    const o = JSON.parse(raw);
                    if (o && o.sig === uiSig) {
                        return o;
                    }
                } catch (e) {
                    /* ignore */
                }
            }
            if (u && typeof u.focusSourceRowIndex === 'number') {
                return u;
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

        const resolveFrameFromSourceRow = (sourceRowIndex) => {
            const anchored = this.getTrackingFrameIndexForSourceRow(sheet, sourceRowIndex);
            return anchored >= 0 ? clampStIdx(anchored) : 0;
        };

        let frameIndex = 0;
        const liveSourceRow = typeof this.comboFocusRowIndex === 'number' && this.comboFocusRowIndex >= 0
            ? this.comboFocusRowIndex
            : (this.activeWindowRange && typeof this.activeWindowRange.target === 'number'
                ? this.activeWindowRange.target
                : -1);
        if (typeof sheet._trackingFocusSourceRowIndex === 'number' && sheet._trackingFocusSourceRowIndex >= 0) {
            frameIndex = resolveFrameFromSourceRow(sheet._trackingFocusSourceRowIndex);
            delete sheet._trackingFocusSourceRowIndex;
        } else if (liveSourceRow >= 0) {
            frameIndex = resolveFrameFromSourceRow(liveSourceRow);
        } else if (savedSt) {
            if (savedSt.sig === uiSig && savedSt.frameIndex != null) {
                frameIndex = clampStIdx(savedSt.frameIndex);
            } else if (typeof savedSt.focusSourceRowIndex === 'number') {
                frameIndex = resolveFrameFromSourceRow(savedSt.focusSourceRowIndex);
            }
        }
        let playing = savedSt ? !!savedSt.playing : false;
        let scrubDrag = false;
        let speed = 1;
        if (savedSt && Number.isFinite(savedSt.speed)) {
            speed = Math.min(3, Math.max(0.5, savedSt.speed));
        }
        const focusNumsByMode = RightPaneSheetManager.readTrackingFocusNumsByMode(
            sheet.trackingUi && sheet.trackingUi.focusNumsByMode
                ? sheet.trackingUi
                : (savedSt || sheet.trackingUi)
        );
        const getFocusNums = () => focusNumsByMode[isBasic ? 'basic' : 'special'];
        let basicLastIdBarNavNum = null;
        let predictNeonOn = !isBasic && savedSt ? !!savedSt.predictNeonOn : false;
        let labelMode = RightPaneSheetManager.normalizeTrackingLabelMode(
            sheet.trackingLabelMode || RightPaneSheetManager.readTrackingLabelModeFromStorage()
        );
        sheet.trackingLabelMode = labelMode;

        const persistLabelMode = () => {
            sheet.trackingLabelMode = labelMode;
            RightPaneSheetManager.writeTrackingLabelModeToStorage(labelMode);
        };

        const syncLabelModeUi = () => {
            root.classList.toggle('special-tracking-root--label-in', labelMode === 'in');
            root.classList.toggle('special-tracking-root--label-out', labelMode === 'out');
            root.dataset.stLabelMode = labelMode;
            const labelToggle = root.querySelector('[data-st-label-toggle]');
            if (labelToggle) {
                labelToggle.textContent = labelMode;
                labelToggle.setAttribute('aria-pressed', labelMode === 'out' ? 'true' : 'false');
                labelToggle.setAttribute('aria-label', `Nhãn bar: ${labelMode}`);
                labelToggle.title = labelMode === 'out'
                    ? 'Nhãn số sát biên bar (bấm để in)'
                    : 'Nhãn số trong bar (bấm để out)';
            }
        };
        syncLabelModeUi();

        const progPct = new Float32Array(total);
        for (let i = 0; i < total; i++) {
            progPct[i] = total <= 1 ? 100 : (i / (total - 1)) * 100;
        }

        let timer = null;
        /** 1× = 100ms mỗi kỳ khi bấm Phát (ms = BASE / speed). */
        const TRACKING_PLAY_BASE_MS = 100;
        const TRACKING_PLAY_MIN_MS = 25;

        const persistTrackingUi = () => {
            try {
                const snap = {
                    sig: uiSig,
                    frameIndex,
                    focusSourceRowIndex: this.getTrackingSourceRowIndexForFrame(sheet, frameIndex),
                    playing,
                    speed,
                    predictNeonOn,
                    focusNumsByMode: RightPaneSheetManager.serializeTrackingFocusNumsByMode(focusNumsByMode),
                    viewMode,
                    labelMode
                };
                sheet.trackingUi = snap;
                persistLabelMode();
                sessionStorage.setItem(TRACKING_UI_STORAGE_KEY, JSON.stringify(snap));
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
            syncMotionClass();
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
        if (statsRankEl && !isBasic && series.length === total && frames.length === total && !Array.isArray(series[0])) {
            const st = RightPaneSheetManager.computeSpecialTrackingPredictRankStats(series, frames);
            statsRankEl.textContent = formatRankStats(st);
            statsRankEl.title = rankStatsTitle(st);
        } else if (statsRankEl) {
            statsRankEl.textContent = '—';
            statsRankEl.removeAttribute('title');
        }

        const onPredictToggle = () => {
            if (isBasic) {
                return;
            }
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
        /** @type {Record<number, { fill: HTMLElement|null, num: HTMLElement|null, count: HTMLElement|null, prio: HTMLElement|null }>} */
        const barPartsByNum = {};
        root.querySelectorAll('[data-st-bar]').forEach((el) => {
            const n = parseInt(el.getAttribute('data-st-bar'), 10);
            if (Number.isFinite(n)) {
                barByNum[n] = el;
                barPartsByNum[n] = {
                    fill: el.querySelector('[data-fill]'),
                    num: el.querySelector('[data-st-num]'),
                    count: el.querySelector('[data-count]'),
                    prio: el.querySelector('[data-st-predict-rank]')
                };
            }
        });
        const freqBraceLayer = root.querySelector('[data-st-freq-braces]');
        const freqBraceGhostLayer = root.querySelector('[data-st-freq-braces-ghost]');
        const freqGapLayer = root.querySelector('[data-st-freq-gap-dividers]');

        let tlRectCache = null;
        const refreshTlRect = () => {
            if (tl) {
                tlRectCache = tl.getBoundingClientRect();
            }
        };

        let paintRaf = 0;
        let frameAnimTarget = frameIndex;
        root.classList.add('special-tracking-root--mount-snap');
        let frameAnimTimer = 0;
        let frameNavOpts = {};
        let frameSteppingActive = false;
        const FRAME_NAV_STEP_MS = 58;
        let sheet1SyncTimer = 0;
        let sheet1SyncRaf = 0;
        let pendingSheet1SyncFrame = -1;
        const clearSheet1SyncTimer = () => {
            if (sheet1SyncTimer) {
                clearTimeout(sheet1SyncTimer);
                sheet1SyncTimer = 0;
            }
            if (sheet1SyncRaf) {
                cancelAnimationFrame(sheet1SyncRaf);
                sheet1SyncRaf = 0;
            }
        };
        /** Paint tracking trước, sync iframe sheet1 frame sau — tránh giật khi Submit ON. */
        const syncSheet1FromTrackingNow = (idx) => {
            clearSheet1SyncTimer();
            pendingSheet1SyncFrame = idx;
            if (sheet1SyncRaf) {
                return;
            }
            sheet1SyncRaf = requestAnimationFrame(() => {
                sheet1SyncRaf = 0;
                const fi = pendingSheet1SyncFrame;
                pendingSheet1SyncFrame = -1;
                if (fi >= 0 && !this._syncingSheet1FromTracking) {
                    this.syncSheet1FromTrackingFrame(fi);
                }
            });
        };
        const scheduleSheet1Sync = (idx, immediate = false) => {
            pendingSheet1SyncFrame = idx;
            if (scrubDrag && !immediate) {
                return;
            }
            if (immediate || !playing) {
                syncSheet1FromTrackingNow(idx);
                return;
            }
            if (sheet1SyncTimer) {
                return;
            }
            const deferMs = this.leftSubmitActive ? 160 : 120;
            sheet1SyncTimer = setTimeout(() => {
                sheet1SyncTimer = 0;
                const fi = pendingSheet1SyncFrame;
                if (fi >= 0) {
                    syncSheet1FromTrackingNow(fi);
                }
            }, deferMs);
        };

        let basicPaintCacheKey = '';
        /** @type {{ basicDisplay: object|null, basicWindow10Freq: Record<number, number>|null, freqTieGroups: object[] }|null} */
        let basicPaintCache = null;
        let paintLastLeftSubmitOn = null;
        let paintLastGhostContextSig = null;
        let specialPreviewAnchorSourceRow = this.getTrackingSourceRowIndexForFrame(sheet, frameIndex);
        const resetSpecialPreviewIfSourceRowChanged = (fi) => {
            if (isBasic) {
                return false;
            }
            const row = this.getTrackingSourceRowIndexForFrame(sheet, fi);
            if (row === specialPreviewAnchorSourceRow) {
                return false;
            }
            specialPreviewAnchorSourceRow = row;
            this.leftSpecialPreviewPickHistory = [];
            this.lastTrackingPreviewBarNum = null;
            return this.setLeftSpecialPreviewPickNum(null);
        };
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
            const ms = Math.max(TRACKING_PLAY_MIN_MS, TRACKING_PLAY_BASE_MS / speed);
            timer = setTimeout(() => {
                timer = null;
                clearFrameAnim();
                frameIndex += 1;
                frameAnimTarget = frameIndex;
                paint();
                if (!this._syncingSheet1FromTracking) {
                    scheduleSheet1Sync(frameIndex);
                }
                scheduleNext();
            }, ms);
        };

        const setScrubbing = (on) => {
            root.classList.toggle('special-tracking-root--scrubbing', on);
            root.classList.toggle('special-tracking-root--playing', !!playing && !on);
        };

        const syncMotionClass = () => {
            root.classList.toggle('special-tracking-root--playing', !!playing && !scrubDrag);
        };

        const paint = () => {
            const fr = frames[frameIndex];
            const leftSubmitOnAtPaintStart = !!this.leftSubmitActive;
            if (freqBraceGhostLayer && paintLastLeftSubmitOn !== null
                && paintLastLeftSubmitOn !== leftSubmitOnAtPaintStart) {
                delete freqBraceGhostLayer.dataset.stBraceGhostSig;
            }
            paintLastLeftSubmitOn = leftSubmitOnAtPaintStart;
            if (!fr) {
                return;
            }
            const p = progPct[frameIndex];
            const canRetroPredict = !isBasic
                && series.length === total
                && frames.length === total
                && !Array.isArray(series[0]);
            let predictList = [];
            if (predictNeonOn && canRetroPredict) {
                const drawPrefix = (fr.drawIndex != null && fr.drawIndex >= 0) ? fr.drawIndex + 1 : 0;
                if (drawPrefix >= 3) {
                    predictList = RightPaneSheetManager.computeSpecialTrackingPredictCandidates(
                        series,
                        frames,
                        drawPrefix
                    );
                }
            }
            let actualNext = null;
            const drawIdx = fr.drawIndex != null ? fr.drawIndex : -1;
            if (!isBasic && drawIdx >= 0 && drawIdx + 1 < series.length) {
                actualNext = series[drawIdx + 1];
            }
            /** @type {Map<number, number>} số → thứ hạng dự đoán 1..3 */
            const predictRankByNum = new Map();
            if (predictList.length) {
                predictList.forEach((pn, idx) => {
                    predictRankByNum.set(pn, idx + 1);
                });
            }
            const predictNeonActive = predictNeonOn && predictList.length > 0;
            root.classList.toggle('special-tracking-root--predict-on', predictNeonOn);
            root.classList.toggle('special-tracking-root--predict-neon-on', predictNeonActive);

            const justNums = Array.isArray(fr.justDrawnNums)
                ? fr.justDrawnNums
                : (fr.justDrawn != null ? [fr.justDrawn] : []);
            const justSet = new Set(justNums);
            const leftSubmitOn = !!this.leftSubmitActive;
            const basicDraws = isBasic ? (sheet.basicDraws || []) : [];
            const freqPreviewLayout = isBasic
                && this.isBasicTrackingFreqPreviewLayoutActive(sheet, frameIndex);
            const previewPickNums = freqPreviewLayout
                ? (this.leftBasicPreviewPickNums || [])
                : [];
            const specialPreviewLayout = !isBasic
                && this.isSpecialTrackingFreqPreviewLayoutActive(sheet, frameIndex);
            const specialPreviewActive = specialPreviewLayout
                && this.leftSpecialPreviewPickNum != null;
            const specialPreviewPick = specialPreviewActive
                ? this.leftSpecialPreviewPickNum
                : null;
            const hasPreviewSimulation = (isBasic && freqPreviewLayout && previewPickNums.length > 0)
                || specialPreviewActive;
            const ghostContextSig = `${leftSubmitOnAtPaintStart ? 1 : 0}|${previewPickNums.join(',')}|${specialPreviewPick ?? ''}`;
            if (freqBraceGhostLayer && paintLastGhostContextSig !== null
                && paintLastGhostContextSig !== ghostContextSig) {
                delete freqBraceGhostLayer.dataset.stBraceGhostSig;
            }
            if (freqBraceLayer && paintLastGhostContextSig !== null
                && paintLastGhostContextSig !== ghostContextSig) {
                delete freqBraceLayer.dataset.stBraceSig;
            }
            paintLastGhostContextSig = ghostContextSig;
            const leftPickBarSyncActive = isBasic
                && this.isBasicTrackingLeftPickBarSyncActive(sheet, frameIndex);
            const leftPickSyncSet = isBasic
                ? (leftPickBarSyncActive
                    ? new Set(this.leftBasicPreviewPickNums || [])
                    : new Set())
                : (specialPreviewActive
                    ? new Set([specialPreviewPick])
                    : new Set());
            const autoringPickSet = this.leftAutoringEnabled && leftSubmitOn
                ? new Set(this.leftBasicPreviewPickNums || [])
                : new Set();
            let basicDisplay = null;
            let basicWindow10Freq = null;
            let basicCh11Ch12Set = null;
            let basicCh11Set = null;
            let basicCh1Ch2Set = null;
            let specialCh1Ch2Set = null;
            let specialCh11Ch12Set = null;
            let freqTieGroups = [];
            let freqTieStreakByKey = new Map();
            let specialDisplay = null;
            if (isBasic) {
                const cacheKey = `${frameIndex}|${leftSubmitOn ? 1 : 0}|${freqPreviewLayout ? 1 : 0}|${previewPickNums.join(',')}`;
                if (basicPaintCacheKey !== cacheKey || !basicPaintCache) {
                    basicDisplay = RightPaneSheetManager.computeBasicTrackingDisplayLayout(
                        basicDraws,
                        fr,
                        leftSubmitOn,
                        previewPickNums
                    );
                    const srcRow = this.getTrackingSourceRowIndexForFrame(sheet, frameIndex);
                    basicWindow10Freq = this.computeMainNumsWindow10Freq(
                        this.sourceRows || [],
                        srcRow
                    );
                    basicCh11Ch12Set = this.getCh11Ch12NumsSetForSourceRow(
                        this.sourceRows || [],
                        srcRow
                    );
                    basicCh11Set = this.getCh11NumsSetForSourceRow(
                        this.sourceRows || [],
                        srcRow
                    );
                    basicCh1Ch2Set = this.getCh1Ch2NumsSetForSourceRow(
                        this.sourceRows || [],
                        srcRow
                    );
                    const tieResult = this.computeTrackingFreqTieGroupsWithStreaks(sheet, frames, frameIndex, {
                        isBasic: true,
                        numMax,
                        leftSubmitOn,
                        basicDraws,
                        previewAtPaintFrame: freqPreviewLayout,
                        previewPickNums
                    });
                    freqTieGroups = tieResult.groups;
                    freqTieStreakByKey = tieResult.streakByKey;
                    basicPaintCacheKey = cacheKey;
                    basicPaintCache = {
                        basicDisplay,
                        basicWindow10Freq,
                        basicCh11Ch12Set,
                        basicCh11Set,
                        basicCh1Ch2Set,
                        freqTieGroups,
                        freqTieStreakByKey
                    };
                } else {
                    basicDisplay = basicPaintCache.basicDisplay;
                    basicWindow10Freq = basicPaintCache.basicWindow10Freq;
                    basicCh11Ch12Set = basicPaintCache.basicCh11Ch12Set;
                    basicCh11Set = basicPaintCache.basicCh11Set;
                    basicCh1Ch2Set = basicPaintCache.basicCh1Ch2Set;
                    freqTieGroups = basicPaintCache.freqTieGroups;
                    freqTieStreakByKey = basicPaintCache.freqTieStreakByKey || new Map();
                }
            } else {
                specialDisplay = RightPaneSheetManager.computeSpecialTrackingDisplayLayout(
                    sheet.specialSeries || sheet.series || [],
                    fr,
                    leftSubmitOn,
                    specialPreviewActive ? specialPreviewPick : null
                );
                const tieResult = this.computeTrackingFreqTieGroupsWithStreaks(sheet, frames, frameIndex, {
                    isBasic: false,
                    numMax,
                    leftSubmitOn,
                    specialPreviewPick,
                    previewAtPaintFrame: specialPreviewActive
                });
                freqTieGroups = tieResult.groups;
                freqTieStreakByKey = tieResult.streakByKey;
                const srcRow = this.getTrackingSourceRowIndexForFrame(sheet, frameIndex);
                specialCh1Ch2Set = this.getSpecialCh1Ch2NumsSetForSourceRow(
                    sheet.specialDrawSteps || [],
                    srcRow
                );
                specialCh11Ch12Set = this.getSpecialCh11Ch12NumsSetForSourceRow(
                    sheet.specialDrawSteps || [],
                    srcRow
                );
            }

            const effectiveJustSet = (!isBasic && specialPreviewActive)
                ? new Set([specialPreviewPick])
                : (!isBasic && !leftSubmitOn)
                    ? new Set()
                    : justSet;

            const specialFreqDisconnected = (!isBasic && specialDisplay)
                ? RightPaneSheetManager.computeSpecialTrackingFreqDisconnectedNums(
                    specialDisplay.counts,
                    specialDisplay.slotByNum,
                    numMax
                )
                : null;

            for (let n = 1; n <= numMax; n++) {
                const el = barByNum[n];
                const parts = barPartsByNum[n];
                if (!el || !parts) {
                    continue;
                }
                const slot = isBasic && basicDisplay
                    ? (basicDisplay.slotByNum[n] ?? 0)
                    : (specialDisplay.slotByNum[n] ?? 0);
                el.style.setProperty('--st-slot', String(slot));
                el.style.setProperty('--st-slot-count', String(slotCount));
                const hueT = RightPaneSheetManager.specialTrackingRankHueT(slot, slotCount - 1);
                el.style.setProperty('--st-hue-t', String(hueT));
                const fillEl = parts.fill;
                const numEl = parts.num;
                const countEl = parts.count;
                const prioEl = parts.prio;
                const wPct = isBasic && basicDisplay
                    ? (basicDisplay.wPctByNum[n] || 0)
                    : (specialDisplay.wPctByNum[n] || 0);
                if (fillEl) {
                    fillEl.style.width = `${wPct}%`;
                }
                const win10Freq = isBasic && basicWindow10Freq ? (basicWindow10Freq[n] || 0) : null;
                const win10FreqZero = isBasic && win10Freq === 0;
                const win10FreqOne = isBasic && win10Freq === 1;
                const win10Ch1Ch2ItalicLeft = isBasic
                    && basicCh1Ch2Set
                    && basicCh1Ch2Set.has(n);
                const specialCh1Ch2ItalicLeft = !isBasic
                    && specialCh1Ch2Set
                    && specialCh1Ch2Set.has(n);
                const ch1Ch2ItalicLeft = win10Ch1Ch2ItalicLeft || specialCh1Ch2ItalicLeft;
                const win10Ch11ItalicRight = isBasic && (
                    (win10FreqOne && basicCh11Ch12Set && basicCh11Ch12Set.has(n))
                    || (win10FreqZero && basicCh11Set && basicCh11Set.has(n))
                );
                const specialCh11Ch12ItalicRight = !isBasic
                    && specialCh11Ch12Set
                    && specialCh11Ch12Set.has(n);
                const ch11ItalicRight = win10Ch11ItalicRight || specialCh11Ch12ItalicRight;
                const ch11Italic = ch1Ch2ItalicLeft ? false : ch11ItalicRight;
                const numItalicSkew = ch1Ch2ItalicLeft
                    ? 'skewX(28deg)'
                    : (ch11Italic ? 'skewX(-12deg)' : '');
                const isJust = effectiveJustSet.has(n);
                const leftPickPreviewLabel = leftPickSyncSet.has(n);
                const actualAnswer = isBasic
                    ? (leftSubmitOn && isJust)
                    : (leftSubmitOn && !predictNeonOn && n === fr.justDrawn);
                const showWin10ZeroTransition = isBasic
                    && win10Freq === 0
                    && ((leftSubmitOn && isJust) || (!leftSubmitOn && leftPickPreviewLabel));
                if (countEl) {
                    const cumFreq = isBasic && basicDisplay
                        ? (basicDisplay.counts[n] || 0)
                        : (specialDisplay.counts[n] || 0);
                    const freqText = String(cumFreq);
                    if (countEl.textContent !== freqText) {
                        countEl.textContent = freqText;
                    }
                }
                if (numEl) {
                    const trackMainEl = el.querySelector('.special-tracking-rank-track--main');
                    const trackLabelEl = el.querySelector('.special-tracking-rank-track--label');
                    if (labelMode === 'out' && trackLabelEl && numEl.parentElement !== trackLabelEl) {
                        trackLabelEl.appendChild(numEl);
                    } else if (labelMode === 'in' && trackMainEl && numEl.parentElement !== trackMainEl) {
                        trackMainEl.appendChild(numEl);
                    }
                    const numBaseTranslate = labelMode === 'out'
                        ? 'translateY(-50%)'
                        : 'translate(-100%, -50%)';
                    if (labelMode === 'out') {
                        numEl.style.left = '';
                        numEl.style.right = '0';
                    } else {
                        const insetPx = isBasic ? 4 : 8;
                        numEl.style.right = 'auto';
                        numEl.style.left = `max(2px, calc(${wPct}% - ${insetPx}px))`;
                    }
                    numEl.style.transform = numItalicSkew
                        ? `${numBaseTranslate} ${numItalicSkew}`
                        : numBaseTranslate;
                    const numText = String(n);
                    if (numEl.textContent !== numText) {
                        numEl.textContent = numText;
                    }
                    numEl.classList.toggle('special-tracking-rank-num--win10-ch11-ch12-italic', ch11Italic);
                    numEl.classList.toggle('special-tracking-rank-num--ch1-ch2-italic-left', ch1Ch2ItalicLeft);
                }
                const autoringBarLabel = isBasic
                    && this.leftAutoringEnabled
                    && leftSubmitOn
                    && isJust
                    && autoringPickSet.has(n)
                    && !actualAnswer;
                el.classList.toggle(
                    'special-tracking-rank-bar--win10-freq-zero',
                    isBasic && win10Freq === 0
                );
                el.classList.toggle(
                    'special-tracking-rank-bar--win10-freq-one',
                    win10FreqOne
                );
                el.classList.toggle(
                    'special-tracking-rank-bar--submit-win10-zero-answer',
                    showWin10ZeroTransition
                );
                el.classList.toggle(
                    'special-tracking-rank-bar--freq-disconnected',
                    !isBasic && specialFreqDisconnected && specialFreqDisconnected.has(n)
                );
                if (countEl) {
                    countEl.classList.toggle('special-tracking-rank-count--win10-ch11-ch12-italic', ch11Italic);
                    countEl.classList.toggle('special-tracking-rank-count--ch1-ch2-italic-left', ch1Ch2ItalicLeft);
                }
                if (isBasic) {
                    el.dataset.stWin10Freq = win10Freq != null ? String(win10Freq) : '';
                } else {
                    delete el.dataset.stWin10Freq;
                }
                const pr = predictRankByNum.get(n);
                const predictHit = Boolean(pr) && actualNext != null && n === actualNext;
                el.classList.toggle(
                    'special-tracking-rank-bar--just',
                    !isBasic && leftSubmitOn && isJust && !actualAnswer && !autoringBarLabel
                );
                const userClickFocus = getFocusNums().has(n);
                el.classList.toggle('special-tracking-rank-bar--click-focus', userClickFocus);
                el.classList.toggle(
                    'special-tracking-rank-bar--focus',
                    autoringBarLabel || leftPickPreviewLabel
                );
                el.classList.toggle('special-tracking-rank-bar--autoring-nonexist', autoringBarLabel);
                el.classList.toggle(
                    'special-tracking-rank-bar--left-pick-preview',
                    leftPickPreviewLabel
                );
                el.classList.toggle('special-tracking-rank-bar--actual-answer', actualAnswer);
                el.classList.toggle('special-tracking-rank-bar--predict', Boolean(pr));
                el.classList.toggle('special-tracking-rank-bar--predict-hit', predictHit);
                el.classList.toggle('special-tracking-rank-bar--predict-1', pr === 1);
                el.classList.toggle('special-tracking-rank-bar--predict-2', pr === 2);
                el.classList.toggle('special-tracking-rank-bar--predict-3', pr === 3);
                if (prioEl) {
                    prioEl.textContent = pr ? `#${pr}` : '';
                    prioEl.setAttribute('aria-hidden', pr ? 'false' : 'true');
                }
                let aria = `Số ${n}, Shift+click để viền cam quan sát`;
                if (showWin10ZeroTransition) {
                    aria = leftPickPreviewLabel && !leftSubmitOn
                        ? `Số ${n}, nonexist cửa sổ 10 — giả lập (xanh lá)`
                        : `Số ${n}, nonexist cửa sổ 10 — đáp án kỳ hiện tại (xanh lá), submit ON`;
                } else if (leftPickPreviewLabel && !leftSubmitOn) {
                    aria = isBasic
                        ? `Số ${n}, click giả lập — preview freq +1 (submit OFF)`
                        : `Số ${n}, click giả lập — preview số đặc biệt kỳ này (submit OFF)`;
                } else if (autoringBarLabel) {
                    aria = `Số ${n}, autoring khoanh trái — đồng bộ với sheet1`;
                } else if (actualAnswer) {
                    aria = isBasic
                        ? `Số ${n}, trong đáp án 5 số chính kỳ hiện tại (id), Shift+click để viền cam`
                        : `Số ${n}, đáp án kỳ hiện tại (id), Shift+click để viền cam`;
                } else if (pr) {
                    aria = predictHit
                        ? `Số ${n}, ứng viên dự đoán hạng ${pr}, trùng đáp án kỳ tiếp theo (id+1)`
                        : `Số ${n}, ứng viên dự đoán hạng ${pr} cho kỳ tiếp theo (id+1), Shift+click để viền cam`;
                }
                if (el.getAttribute('aria-label') !== aria) {
                    el.setAttribute('aria-label', aria);
                }
            }
            const ghostMirrorSolid = !leftSubmitOn && !hasPreviewSimulation;
            let ghostLabel = 'Trước thay đổi';
            if (ghostMirrorSolid) {
                ghostLabel = 'Đồng bộ solid';
            } else if (leftSubmitOn && hasPreviewSimulation) {
                ghostLabel = 'Trước submit / giả lập';
            } else if (leftSubmitOn) {
                ghostLabel = 'Trước submit';
            } else if (hasPreviewSimulation) {
                ghostLabel = 'Trước giả lập';
            }
            let ghostGroupsToDraw = [];
            let ghostStreakToDraw = new Map();
            const basicPreviewUnchangedReady = isBasic
                && freqPreviewLayout
                && previewPickNums.length >= 5;
            const specialPreviewUnchangedReady = !isBasic && specialPreviewActive;
            const previewUnchangedBellyReady = basicPreviewUnchangedReady
                || specialPreviewUnchangedReady;
            let beforePreviewTie = null;
            if (hasPreviewSimulation) {
                beforePreviewTie = this.computeTrackingBeforePreviewGhostTieResult(
                    sheet,
                    frames,
                    frameIndex,
                    {
                        isBasic,
                        numMax,
                        basicDraws: isBasic ? (sheet.basicDraws || []) : [],
                        leftSubmitOn
                    }
                );
            }
            if (ghostMirrorSolid) {
                ghostGroupsToDraw = freqTieGroups;
                ghostStreakToDraw = freqTieStreakByKey;
            } else if (hasPreviewSimulation) {
                ghostGroupsToDraw = beforePreviewTie.groups;
                ghostStreakToDraw = beforePreviewTie.streakByKey;
            } else if (leftSubmitOn) {
                const submitGhostTie = this.computeTrackingGhostFreqTieResult(
                    sheet,
                    frames,
                    frameIndex,
                    {
                        isBasic,
                        numMax,
                        basicDraws: isBasic ? (sheet.basicDraws || []) : [],
                        leftSubmitOn,
                        freqPreviewLayout,
                        specialPreviewLayout: specialPreviewActive,
                        previewPickNums,
                        specialPreviewPick
                    }
                );
                ghostGroupsToDraw = submitGhostTie.groups;
                ghostStreakToDraw = submitGhostTie.streakByKey;
            }
            let referenceGroupsForSolid = null;
            if (hasPreviewSimulation && beforePreviewTie) {
                referenceGroupsForSolid = beforePreviewTie.groups;
            } else if (leftSubmitOn && ghostGroupsToDraw.length) {
                referenceGroupsForSolid = ghostGroupsToDraw;
            }
            const solidShiftedKeys = referenceGroupsForSolid
                ? RightPaneSheetManager.getTrackingFreqSolidShiftedKeys(
                    referenceGroupsForSolid,
                    freqTieGroups
                )
                : null;
            const ghostChangedKeys = ghostMirrorSolid
                ? null
                : RightPaneSheetManager.getTrackingFreqBellyStreakKeysWithoutExactMatchIn(
                    ghostGroupsToDraw,
                    freqTieGroups
                );
            if (freqBraceLayer) {
                const solidStreakByKey = (previewUnchangedBellyReady && beforePreviewTie)
                    ? RightPaneSheetManager.applyTrackingSolidPreviewStreakBump(
                        freqTieGroups,
                        freqTieStreakByKey,
                        beforePreviewTie,
                        solidShiftedKeys
                    )
                    : freqTieStreakByKey;
                RightPaneSheetManager.syncTrackingFreqBraces(
                    freqBraceLayer,
                    freqTieGroups,
                    slotCount,
                    solidStreakByKey,
                    { solidShiftedKeys }
                );
            }
            if (freqBraceGhostLayer) {
                RightPaneSheetManager.syncTrackingFreqBraces(
                    freqBraceGhostLayer,
                    ghostGroupsToDraw,
                    slotCount,
                    ghostStreakToDraw,
                    {
                        ghost: true,
                        ghostLabel,
                        ghostChangedKeys
                    }
                );
            }
            if (freqGapLayer) {
                const gapCounts = isBasic && basicDisplay
                    ? basicDisplay.counts
                    : specialDisplay.counts;
                const gapSlots = isBasic && basicDisplay
                    ? basicDisplay.slotByNum
                    : specialDisplay.slotByNum;
                const gapDividers = RightPaneSheetManager.computeTrackingFreqGapDividers(
                    gapCounts,
                    gapSlots,
                    numMax
                );
                RightPaneSheetManager.syncTrackingFreqGapDividers(
                    freqGapLayer,
                    gapDividers,
                    slotCount
                );
            }
            syncTimelineUi();
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

        const clearFrameAnim = () => {
            if (frameAnimTimer) {
                clearTimeout(frameAnimTimer);
                frameAnimTimer = 0;
            }
        };

        const syncTimelineUi = () => {
            const p = progPct[frameIndex];
            if (stepEl) {
                const periodId = this.getTrackingPeriodIdForFrame(sheet, frameIndex);
                const frStep = frames[frameIndex];
                const stepLabel = periodId || (frStep ? String(frStep.step) : '');
                stepEl.innerHTML = `<strong>${stepLabel}</strong> / ${total}`;
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

        const setFrameStepping = (on) => {
            frameSteppingActive = !!on;
            root.classList.toggle('special-tracking-root--frame-stepping', frameSteppingActive);
        };

        const syncLeftPaneFromFrame = () => {
            if (!this._syncingSheet1FromTracking && !frameNavOpts.skipSheet1Sync) {
                this.syncSheet1FromTrackingFrame(frameIndex, {
                    trackingFrameStep: frameSteppingActive
                });
            }
        };

        const finishFrameNav = () => {
            setFrameStepping(false);
            persistTrackingUi();
            scheduleNext();
            syncLeftPaneFromFrame();
        };

        const pumpFrameAnim = () => {
            frameAnimTimer = 0;
            if (frameIndex === frameAnimTarget) {
                finishFrameNav();
                return;
            }
            setFrameStepping(true);
            frameIndex += frameIndex < frameAnimTarget ? 1 : -1;
            resetSpecialPreviewIfSourceRowChanged(frameIndex);
            syncBasicLastIdBarNavAnchor();
            syncTimelineUi();
            paint();
            syncLeftPaneFromFrame();
            frameAnimTimer = setTimeout(pumpFrameAnim, FRAME_NAV_STEP_MS);
        };

        const applyFrameImmediate = (idx, opts = {}) => {
            const next = Math.max(0, Math.min(total - 1, idx));
            resetSpecialPreviewIfSourceRowChanged(next);
            clearFrameAnim();
            setFrameStepping(false);
            frameNavOpts = opts || {};
            frameAnimTarget = next;
            frameIndex = next;
            syncBasicLastIdBarNavAnchor();
            syncTimelineUi();
            schedulePaint();
            persistTrackingUi();
            scheduleNext();
            syncLeftPaneFromFrame();
        };

        const setFrameNav = (idx, opts = {}) => {
            const next = Math.max(0, Math.min(total - 1, idx));
            frameNavOpts = opts || {};
            frameAnimTarget = next;
            if (next === frameIndex && !frameAnimTimer) {
                return;
            }
            const jump = Math.abs(next - frameIndex);
            if (frameNavOpts.immediate || jump > 12) {
                applyFrameImmediate(next, frameNavOpts);
                return;
            }
            if (!frameAnimTimer) {
                pumpFrameAnim();
            }
        };

        const syncBasicLastIdBarNavAnchor = () => {
            if (!this.isBasicTrackingFramePreviewEligible(sheet, frameIndex)) {
                basicLastIdBarNavNum = null;
            }
        };

        const getBasicLastIdBarNavOrder = () => {
            const fr = frames[frameIndex];
            if (!fr) {
                return [];
            }
            // Thứ tự cố định theo freq gốc — không cộng +1 giả lập của pick hiện tại
            // (tránh hoán vị 22↔23 khi arrow chuyển viền đen giữa hai bar hòa điểm).
            const basicDisplay = RightPaneSheetManager.computeBasicTrackingDisplayLayout(
                sheet.basicDraws || [],
                fr,
                !!this.leftSubmitActive,
                []
            );
            const bottomSlot = numMax - 1;
            const items = [];
            for (let n = 1; n <= numMax; n++) {
                items.push({ n, slot: basicDisplay.slotByNum[n] ?? bottomSlot });
            }
            items.sort((a, b) => a.slot - b.slot || a.n - b.n);
            return items.map((x) => x.n);
        };

        const applyBasicLastIdBarNavPick = (n) => {
            basicLastIdBarNavNum = n;
            this.lastTrackingPreviewBarNum = n;
            this.setLeftBasicPreviewPickNums([n]);
            if (this.shouldSyncBasicBarPickToLeftPane()) {
                this.syncLeftPickSelectionToIframe([n]);
            }
            try {
                window.dispatchEvent(new CustomEvent('leftCircledNumsChanged'));
            } catch (ePaint) { /* ignore */ }
            paint();
        };

        const stepBasicLastIdBarNav = (delta) => {
            if (!this.isBasicTrackingFramePreviewEligible(sheet, frameIndex)) {
                return false;
            }
            const order = getBasicLastIdBarNavOrder();
            if (!order.length) {
                return false;
            }
            const step = Number(delta) || 0;
            if (!step) {
                return false;
            }
            if (basicLastIdBarNavNum == null) {
                applyBasicLastIdBarNavPick(step > 0 ? order[0] : order[order.length - 1]);
                return true;
            }
            let idx = order.indexOf(basicLastIdBarNavNum);
            if (idx < 0) {
                idx = 0;
            }
            const len = order.length;
            let nextIdx = (idx + step) % len;
            if (nextIdx < 0) {
                nextIdx += len;
            }
            applyBasicLastIdBarNavPick(order[nextIdx]);
            return true;
        };

        const setFrame = (idx) => {
            setFrameNav(idx, { immediateSheet1Sync: true });
        };

        const setFrameRaf = (idx, opts = {}) => {
            if (scrubDrag || opts.immediate || this._syncingSheet1FromTracking) {
                applyFrameImmediate(idx, opts);
                return;
            }
            setFrameNav(idx, opts);
        };

        tableWrap.__trackingSeekFrame = (idx, opts = {}) => {
            setFrameRaf(clampStIdx(idx), opts || {});
        };
        tableWrap.__specialTrackingSeekFrame = tableWrap.__trackingSeekFrame;
        tableWrap.__trackingStepFrame = (delta) => {
            setFrameRaf(frameIndex + (Number(delta) || 0));
        };
        tableWrap.__trackingStepBasicLastIdBar = stepBasicLastIdBarNav;

        const togglePlay = () => {
            if (frameIndex >= total - 1) {
                frameIndex = 0;
                frameAnimTarget = 0;
            }
            playing = !playing;
            if (playing) {
                clearFrameAnim();
                frameAnimTarget = frameIndex;
            }
            syncMotionClass();
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
            syncMotionClass();
            tlRectCache = null;
            window.removeEventListener('mousemove', onScrubMove, true);
            window.removeEventListener('mouseup', onScrubUp, true);
            window.removeEventListener('touchmove', onScrubTouchMove, true);
            window.removeEventListener('touchend', onScrubTouchEnd, true);
            window.removeEventListener('touchcancel', onScrubTouchEnd, true);
            if (!this._syncingSheet1FromTracking) {
                syncSheet1FromTrackingNow(frameIndex);
            }
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

        const rememberMainKeyboardFocus = () => {
            if (typeof window.rememberKeyboardFocusTarget === 'function') {
                window.rememberKeyboardFocusTarget('main');
            }
        };

        const onTlDown = (ev) => {
            rememberMainKeyboardFocus();
            const cx = ev.clientX != null ? ev.clientX : (ev.touches && ev.touches[0] ? ev.touches[0].clientX : 0);
            scrubDrag = true;
            setScrubbing(true);
            syncMotionClass();
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
                applyFrameImmediate(0, { immediateSheet1Sync: true });
            });
        }
        if (btnLast) {
            btnLast.addEventListener('click', () => {
                playing = false;
                if (btnPlay) {
                    syncPlayBtnUi();
                }
                applyFrameImmediate(total - 1, { immediateSheet1Sync: true });
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
            const onPick = (ev) => {
                rememberMainKeyboardFocus();
                const n = parseInt(row.dataset.specialNum, 10);
                if (!Number.isFinite(n)) {
                    return;
                }
                const shiftClick = !!(ev && ev.shiftKey);
                if (shiftClick) {
                    toggleObserveFocusNum(n);
                    return;
                }
                if (isBasic && this.isBasicTrackingFramePreviewEligible(sheet, frameIndex)) {
                    if (this.toggleBasicTrackingLastIdBarPick(n)) {
                        basicLastIdBarNavNum = this.lastTrackingPreviewBarNum;
                        paint();
                    }
                    return;
                }
                if (!isBasic && this.isSpecialTrackingFramePreviewEligible(sheet, frameIndex)) {
                    if (this.toggleSpecialTrackingBarPick(n)) {
                        paint();
                    }
                }
            };
            row.addEventListener('mousedown', (e) => {
                if (e.button === 2) {
                    e.preventDefault();
                    return;
                }
                if (e.button === 0) {
                    e.preventDefault();
                }
            });
            row.addEventListener('click', (e) => {
                onPick(e);
                try {
                    row.blur();
                } catch (eBlur) {
                    /* ignore */
                }
            });
            row.addEventListener('keydown', (e) => {
                if (e.key === 'Enter' || e.key === ' ') {
                    e.preventDefault();
                    onPick(e);
                }
            });
        });

        /** Chuột phải nửa màn phải: toggle giả lập bar vừa click gần nhất (nửa trái giữ menu trình duyệt). */
        const rightPaneEl = tableWrap.closest('.pane.right')
            || document.querySelector('.pane.right');
        const resolveLastPreviewBarToToggle = () => {
            let n = this.lastTrackingPreviewBarNum;
            if (isBasic) {
                const maxN = 35;
                const picks = this.leftBasicPreviewPickNums || [];
                if (Number.isFinite(n) && n >= 1 && n <= maxN) {
                    return n;
                }
                if (picks.length) {
                    return picks[picks.length - 1];
                }
                if (Number.isFinite(basicLastIdBarNavNum) && basicLastIdBarNavNum >= 1
                    && basicLastIdBarNavNum <= maxN) {
                    return basicLastIdBarNavNum;
                }
                return null;
            }
            const maxN = 12;
            if (Number.isFinite(n) && n >= 1 && n <= maxN) {
                return n;
            }
            const hist = this.leftSpecialPreviewPickHistory || [];
            if (hist.length) {
                const tip = hist[hist.length - 1];
                if (Number.isFinite(tip) && tip >= 1 && tip <= maxN) {
                    return tip;
                }
            }
            n = this.leftSpecialPreviewPickNum;
            return (Number.isFinite(n) && n >= 1 && n <= maxN) ? n : null;
        };
        const handleRightPanePreviewToggle = (e) => {
            const eligible = isBasic
                ? this.isBasicTrackingFramePreviewEligible(sheet, frameIndex)
                : this.isSpecialTrackingFramePreviewEligible(sheet, frameIndex);
            if (!eligible) {
                return false;
            }
            const n = resolveLastPreviewBarToToggle();
            if (n == null) {
                return false;
            }
            e.preventDefault();
            e.stopPropagation();
            let changed = false;
            if (isBasic) {
                changed = this.toggleBasicTrackingLastIdBarPick(n, { retainFocus: true });
                if (changed) {
                    basicLastIdBarNavNum = this.lastTrackingPreviewBarNum;
                }
            } else {
                changed = this.toggleSpecialTrackingBarPick(n, { retainFocus: true });
            }
            if (changed) {
                paint();
                const ae = document.activeElement;
                if (ae && rightPaneEl && rightPaneEl.contains(ae)
                    && typeof ae.blur === 'function') {
                    try {
                        ae.blur();
                    } catch (eBlur) { /* ignore */ }
                }
            }
            return true;
        };
        /** Chặn menu trình duyệt sớm; chỉ toggle một lần ở contextmenu (tránh mousedown+contextmenu đảo hai lần). */
        const onRightPanePreviewPointerDown = (e) => {
            if (e.button !== 2) {
                return;
            }
            const eligible = isBasic
                ? this.isBasicTrackingFramePreviewEligible(sheet, frameIndex)
                : this.isSpecialTrackingFramePreviewEligible(sheet, frameIndex);
            if (!eligible || resolveLastPreviewBarToToggle() == null) {
                return;
            }
            e.preventDefault();
            e.stopPropagation();
        };
        const onRightPanePreviewContextMenu = (e) => {
            handleRightPanePreviewToggle(e);
        };
        if (rightPaneEl) {
            const cap = true;
            rightPaneEl.addEventListener('mousedown', onRightPanePreviewPointerDown, cap);
            rightPaneEl.addEventListener('pointerdown', onRightPanePreviewPointerDown, cap);
            rightPaneEl.addEventListener('contextmenu', onRightPanePreviewContextMenu, cap);
        }

        const viewToggle = root.querySelector('[data-st-view-toggle]');
        const labelToggle = root.querySelector('[data-st-label-toggle]');
        const onLabelToggle = () => {
            labelMode = labelMode === 'out' ? 'in' : 'out';
            syncLabelModeUi();
            persistLabelMode();
            paint();
            persistTrackingUi();
        };
        if (labelToggle) {
            labelToggle.addEventListener('click', onLabelToggle);
        }
        if (viewToggle) {
            viewToggle.addEventListener('click', () => {
                playing = false;
                clearTimer();
                const anchorRow = this.getTrackingSourceRowIndexForFrame(sheet, frameIndex);
                const nextMode = isBasic ? 'special' : 'basic';
                sheet.trackingViewMode = nextMode;
                sheet._trackingFramesMode = null;
                if (anchorRow >= 0) {
                    sheet._trackingFocusSourceRowIndex = anchorRow;
                }
                this.renderTable(tableWrap);
                try {
                    this.emitTrackingObserveFocusChanged();
                } catch (eModeEmit) { /* ignore */ }
            });
        }

        const onTrackingArrowNav = (ev) => {
            if (!tableWrap.classList.contains('table-wrap--tracking')) {
                return;
            }
            const verticalStep = ev.key === 'ArrowDown'
                ? 1
                : ev.key === 'ArrowUp'
                    ? -1
                    : 0;
            const timelineStep = ev.key === 'ArrowRight'
                ? 1
                : ev.key === 'ArrowLeft'
                    ? -1
                    : 0;
            if (!verticalStep && !timelineStep) {
                return;
            }
            if (ev.ctrlKey || ev.metaKey || ev.altKey) {
                return;
            }
            if (typeof window.shouldRouteRightPaneArrowsToFilter === 'function'
                && window.shouldRouteRightPaneArrowsToFilter()) {
                return;
            }
            if (typeof window.shouldRouteRightPaneArrowsToAnswer === 'function'
                && window.shouldRouteRightPaneArrowsToAnswer()) {
                return;
            }
            const t = ev.target;
            if (t instanceof Element && t.closest('input, textarea, select, [contenteditable="true"]')) {
                return;
            }
            ev.preventDefault();
            ev.stopPropagation();
            rememberMainKeyboardFocus();
            if (verticalStep && stepBasicLastIdBarNav(verticalStep)) {
                return;
            }
            if (timelineStep) {
                setFrame(frameIndex + timelineStep);
            }
        };
        window.addEventListener('keydown', onTrackingArrowNav, true);

        const onLeftSubmitStateChanged = () => {
            paint();
        };
        const onLeftAutoringStateChanged = () => {
            if (isBasic) {
                requestLeftCircledForPreview();
            }
        };
        const onLeftCircledNumsChanged = () => {
            paint();
        };
        const onRowClickedForPreview = () => {
            paint();
        };
        window.addEventListener('leftSubmitStateChanged', onLeftSubmitStateChanged);
        window.addEventListener('leftAutoringStateChanged', onLeftAutoringStateChanged);
        window.addEventListener('leftCircledNumsChanged', onLeftCircledNumsChanged);
        window.addEventListener('rowClicked', onRowClickedForPreview);

        const requestLeftCircledForPreview = () => {
            if (!isBasic) {
                return;
            }
            const frame = document.getElementById('okFrame');
            if (!frame || !frame.contentWindow) {
                return;
            }
            const reqGen = this._leftBasicPreviewPickGeneration;
            const nonce = `trkPrev_${Date.now()}_${Math.random().toString(36).slice(2)}`;
            const onMessage = (ev) => {
                const msg = ev.data || {};
                if (msg.type !== 'leftCircledNums' || msg.nonce !== nonce) {
                    return;
                }
                window.removeEventListener('message', onMessage);
                if (reqGen !== this._leftBasicPreviewPickGeneration) {
                    return;
                }
                if (this.setLeftBasicPreviewPickNums(msg.nums || [])) {
                    paint();
                }
            };
            window.addEventListener('message', onMessage);
            frame.contentWindow.postMessage({ type: 'requestLeftCircledNums', nonce }, '*');
            setTimeout(() => window.removeEventListener('message', onMessage), 400);
        };
        if (isBasic) {
            requestLeftCircledForPreview();
        }

        paint();
        syncMotionClass();
        requestAnimationFrame(() => {
            root.classList.remove('special-tracking-root--mount-snap');
        });
        if (btnPlay) {
            syncPlayBtnUi();
        }
        if (playing) {
            scheduleNext();
        }
        if (!this._syncingSheet1FromTracking) {
            const syncRow = this.getTrackingSourceRowIndexForFrame(sheet, frameIndex);
            const skipSync = liveSourceRow >= 0 && syncRow === liveSourceRow;
            if (!skipSync) {
                this.syncSheet1FromTrackingFrame(frameIndex);
            }
        }

        const cleanupTrackingUi = () => {
            try {
                delete tableWrap.__trackingSeekFrame;
                delete tableWrap.__specialTrackingSeekFrame;
                delete tableWrap.__trackingStepFrame;
                delete tableWrap.__trackingStepBasicLastIdBar;
                delete tableWrap.__trackingRepaint;
                delete tableWrap.__trackingGetFrameIndex;
                delete tableWrap.__trackingToggleObserveFocus;
                delete tableWrap.__trackingToggleObserveFocusAll;
                delete tableWrap.__trackingSetObserveFocusNums;
            } catch (eDelSeek) {
                /* ignore */
            }
            persistTrackingUi();
            clearTimer();
            clearFrameAnim();
            setFrameStepping(false);
            clearSheet1SyncTimer();
            scrubDrag = false;
            setScrubbing(false);
            syncMotionClass();
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
            if (viewToggle) {
                viewToggle.replaceWith(viewToggle.cloneNode(true));
            }
            if (labelToggle) {
                labelToggle.removeEventListener('click', onLabelToggle);
            }
            window.removeEventListener('keydown', onTrackingArrowNav, true);
            window.removeEventListener('leftSubmitStateChanged', onLeftSubmitStateChanged);
            window.removeEventListener('leftAutoringStateChanged', onLeftAutoringStateChanged);
            window.removeEventListener('leftCircledNumsChanged', onLeftCircledNumsChanged);
            window.removeEventListener('rowClicked', onRowClickedForPreview);
            if (rightPaneEl) {
                const cap = true;
                rightPaneEl.removeEventListener('mousedown', onRightPanePreviewPointerDown, cap);
                rightPaneEl.removeEventListener('pointerdown', onRightPanePreviewPointerDown, cap);
                rightPaneEl.removeEventListener('contextmenu', onRightPanePreviewContextMenu, cap);
            }
        };
        const abandonObserveFocusStashForMode = (mode) => {
            if (!this._trackingObserveFocusStashByMode) {
                this._trackingObserveFocusStashByMode = { basic: null, special: null };
            }
            this._trackingObserveFocusStashByMode[mode] = null;
        };

        const emitObserveFocusChanged = () => {
            try {
                this.emitTrackingObserveFocusChanged();
            } catch (eEmit) { /* ignore */ }
        };

        const toggleObserveFocusNum = (rawNum) => {
            const n = Math.floor(Number(rawNum));
            const maxN = isBasic ? 35 : 12;
            if (!Number.isFinite(n) || n < 1 || n > maxN) {
                return false;
            }
            const mode = isBasic ? 'basic' : 'special';
            abandonObserveFocusStashForMode(mode);
            const focusNums = getFocusNums();
            if (focusNums.has(n)) {
                focusNums.delete(n);
            } else {
                focusNums.add(n);
            }
            paint();
            persistTrackingUi();
            emitObserveFocusChanged();
            return true;
        };

        const toggleObserveFocusAll = () => {
            if (!this._trackingObserveFocusStashByMode) {
                this._trackingObserveFocusStashByMode = { basic: null, special: null };
            }
            const mode = isBasic ? 'basic' : 'special';
            const focusNums = getFocusNums();
            const stash = this._trackingObserveFocusStashByMode[mode];
            if (stash != null) {
                focusNums.clear();
                stash.forEach((n) => focusNums.add(n));
                this._trackingObserveFocusStashByMode[mode] = null;
            } else {
                this._trackingObserveFocusStashByMode[mode] = new Set(focusNums);
                focusNums.clear();
            }
            paint();
            persistTrackingUi();
            emitObserveFocusChanged();
            return true;
        };

        tableWrap.__trackingCleanup = cleanupTrackingUi;
        tableWrap.__specialTrackingCleanup = cleanupTrackingUi;
        tableWrap.__trackingRepaint = () => {
            paint();
        };
        tableWrap.__trackingGetFrameIndex = () => frameIndex;
        const setObserveFocusNums = (rawNums, options) => {
            const force = !!(options && options.force);
            const mode = (options && options.mode === 'special')
                ? 'special'
                : (options && options.mode === 'basic')
                    ? 'basic'
                    : (isBasic ? 'basic' : 'special');
            const maxN = mode === 'basic' ? 35 : 12;
            if (!this._trackingObserveFocusStashByMode) {
                this._trackingObserveFocusStashByMode = { basic: null, special: null };
            }
            const next = new Set();
            (Array.isArray(rawNums) ? rawNums : []).forEach((x) => {
                const n = Math.floor(Number(x));
                if (Number.isFinite(n) && n >= 1 && n <= maxN) {
                    next.add(n);
                }
            });
            const focusNums = focusNumsByMode[mode];
            let same = focusNums.size === next.size && next.size > 0;
            if (same) {
                next.forEach((n) => {
                    if (!focusNums.has(n)) {
                        same = false;
                    }
                });
            }
            const toggleOff = !force && same;
            focusNums.clear();
            if (toggleOff) {
                /* Click lại Chuỗi x → tắt, stash để Ctrl+Shift bật lại */
                this._trackingObserveFocusStashByMode[mode] = new Set(next);
            } else {
                abandonObserveFocusStashForMode(mode);
                next.forEach((n) => focusNums.add(n));
            }
            paint();
            persistTrackingUi();
            emitObserveFocusChanged();
            return true;
        };

        tableWrap.__trackingToggleObserveFocus = toggleObserveFocusNum;
        tableWrap.__trackingToggleObserveFocusAll = toggleObserveFocusAll;
        tableWrap.__trackingSetObserveFocusNums = setObserveFocusNums;
    }

    /** Số đang có viền cam quan sát theo mode ('basic' | 'special'). */
    getTrackingBarObserveFocusNumsForMode(mode) {
        const sheet = this.sheets[TRACKING_SHEET_ID] || this.sheets.specialtracking;
        if (!sheet || sheet.kind !== TRACKING_KIND) {
            return [];
        }
        const key = mode === 'special' ? 'special' : 'basic';
        const byMode = RightPaneSheetManager.readTrackingFocusNumsByMode(
            sheet.trackingUi || {}
        );
        return Array.from(byMode[key] || []).sort((a, b) => a - b);
    }

    /** Số đang có viền cam quan sát (mode tracking hiện tại). */
    getTrackingBarObserveFocusNums() {
        return this.getTrackingBarObserveFocusNumsForMode(this.getActiveTrackingViewMode());
    }

    /** basic | special — mode tracking đang xem. */
    getActiveTrackingViewMode() {
        const sheet = this.sheets[TRACKING_SHEET_ID] || this.sheets.specialtracking;
        if (!sheet || sheet.kind !== TRACKING_KIND) {
            return 'basic';
        }
        return this.getTrackingViewMode(sheet);
    }

    /** Báo nửa trái: gạch chân Chuỗi luôn theo basicNums (tách biệt special). */
    emitTrackingObserveFocusChanged() {
        try {
            window.dispatchEvent(new CustomEvent('trackingObserveFocusNumsChanged', {
                detail: {
                    nums: this.getTrackingBarObserveFocusNums(),
                    viewMode: this.getActiveTrackingViewMode(),
                    basicNums: this.getTrackingBarObserveFocusNumsForMode('basic')
                }
            }));
        } catch (e) { /* ignore */ }
    }

    /**
     * Shift+click số nửa trái (hoặc bar tracking): bật/tắt viền cam quan sát trên bar tương ứng.
     * @returns {boolean}
     */
    toggleTrackingBarObserveFocus(rawNum) {
        const tableWrap = typeof document !== 'undefined'
            ? document.getElementById('tableWrap')
            : null;
        if (tableWrap && typeof tableWrap.__trackingToggleObserveFocus === 'function') {
            return !!tableWrap.__trackingToggleObserveFocus(rawNum);
        }
        const sheet = this.sheets[TRACKING_SHEET_ID] || this.sheets.specialtracking;
        if (!sheet || sheet.kind !== TRACKING_KIND) {
            return false;
        }
        const viewMode = this.getTrackingViewMode(sheet);
        const maxN = viewMode === 'basic' ? 35 : 12;
        const n = Math.floor(Number(rawNum));
        if (!Number.isFinite(n) || n < 1 || n > maxN) {
            return false;
        }
        const mode = viewMode === 'basic' ? 'basic' : 'special';
        if (!this._trackingObserveFocusStashByMode) {
            this._trackingObserveFocusStashByMode = { basic: null, special: null };
        }
        this._trackingObserveFocusStashByMode[mode] = null;
        const byMode = RightPaneSheetManager.readTrackingFocusNumsByMode(
            sheet.trackingUi || {}
        );
        const set = byMode[mode];
        if (set.has(n)) {
            set.delete(n);
        } else {
            set.add(n);
        }
        const prev = sheet.trackingUi && typeof sheet.trackingUi === 'object'
            ? sheet.trackingUi
            : {};
        sheet.trackingUi = {
            ...prev,
            viewMode,
            focusNumsByMode: RightPaneSheetManager.serializeTrackingFocusNumsByMode(byMode)
        };
        try {
            sessionStorage.setItem(
                TRACKING_UI_STORAGE_KEY,
                JSON.stringify(sheet.trackingUi)
            );
        } catch (e) {
            /* ignore */
        }
        this.emitTrackingObserveFocusChanged();
        return true;
    }

    /**
     * Thay toàn bộ viền cam quan sát bằng danh sách số (vd. 5 số Chuỗi x → mode basic).
     * @param {number[]} rawNums
     * @param {{ force?: boolean, mode?: 'basic'|'special' }} [options]
     * @returns {boolean}
     */
    setTrackingBarObserveFocusNums(rawNums, options) {
        const tableWrap = typeof document !== 'undefined'
            ? document.getElementById('tableWrap')
            : null;
        if (tableWrap && typeof tableWrap.__trackingSetObserveFocusNums === 'function') {
            return !!tableWrap.__trackingSetObserveFocusNums(rawNums, options);
        }
        const force = !!(options && options.force);
        const sheet = this.sheets[TRACKING_SHEET_ID] || this.sheets.specialtracking;
        if (!sheet || sheet.kind !== TRACKING_KIND) {
            return false;
        }
        if (!this._trackingObserveFocusStashByMode) {
            this._trackingObserveFocusStashByMode = { basic: null, special: null };
        }
        const viewMode = this.getTrackingViewMode(sheet);
        const mode = (options && options.mode === 'special')
            ? 'special'
            : (options && options.mode === 'basic')
                ? 'basic'
                : (viewMode === 'basic' ? 'basic' : 'special');
        const maxN = mode === 'basic' ? 35 : 12;
        const byMode = RightPaneSheetManager.readTrackingFocusNumsByMode(
            sheet.trackingUi || {}
        );
        const set = byMode[mode];
        const next = new Set();
        (Array.isArray(rawNums) ? rawNums : []).forEach((x) => {
            const n = Math.floor(Number(x));
            if (Number.isFinite(n) && n >= 1 && n <= maxN) {
                next.add(n);
            }
        });
        let same = set.size === next.size && next.size > 0;
        if (same) {
            next.forEach((n) => {
                if (!set.has(n)) {
                    same = false;
                }
            });
        }
        const toggleOff = !force && same;
        set.clear();
        if (toggleOff) {
            this._trackingObserveFocusStashByMode[mode] = new Set(next);
        } else {
            this._trackingObserveFocusStashByMode[mode] = null;
            next.forEach((n) => set.add(n));
        }
        const prev = sheet.trackingUi && typeof sheet.trackingUi === 'object'
            ? sheet.trackingUi
            : {};
        sheet.trackingUi = {
            ...prev,
            viewMode,
            focusNumsByMode: RightPaneSheetManager.serializeTrackingFocusNumsByMode(byMode)
        };
        try {
            sessionStorage.setItem(
                TRACKING_UI_STORAGE_KEY,
                JSON.stringify(sheet.trackingUi)
            );
        } catch (e) {
            /* ignore */
        }
        this.emitTrackingObserveFocusChanged();
        return true;
    }

    /**
     * Ctrl+Shift: tắt hết viền cam quan sát (mode hiện tại); lần nữa khôi phục stash vừa tắt.
     * @returns {boolean}
     */
    toggleTrackingBarObserveFocusAll() {
        const tableWrap = typeof document !== 'undefined'
            ? document.getElementById('tableWrap')
            : null;
        if (tableWrap && typeof tableWrap.__trackingToggleObserveFocusAll === 'function') {
            return !!tableWrap.__trackingToggleObserveFocusAll();
        }
        const sheet = this.sheets[TRACKING_SHEET_ID] || this.sheets.specialtracking;
        if (!sheet || sheet.kind !== TRACKING_KIND) {
            return false;
        }
        if (!this._trackingObserveFocusStashByMode) {
            this._trackingObserveFocusStashByMode = { basic: null, special: null };
        }
        const viewMode = this.getTrackingViewMode(sheet);
        const mode = viewMode === 'basic' ? 'basic' : 'special';
        const byMode = RightPaneSheetManager.readTrackingFocusNumsByMode(
            sheet.trackingUi || {}
        );
        const set = byMode[mode];
        const stash = this._trackingObserveFocusStashByMode[mode];
        if (stash != null) {
            set.clear();
            stash.forEach((n) => set.add(n));
            this._trackingObserveFocusStashByMode[mode] = null;
        } else {
            this._trackingObserveFocusStashByMode[mode] = new Set(set);
            set.clear();
        }
        const prev = sheet.trackingUi && typeof sheet.trackingUi === 'object'
            ? sheet.trackingUi
            : {};
        sheet.trackingUi = {
            ...prev,
            viewMode,
            focusNumsByMode: RightPaneSheetManager.serializeTrackingFocusNumsByMode(byMode)
        };
        try {
            sessionStorage.setItem(
                TRACKING_UI_STORAGE_KEY,
                JSON.stringify(sheet.trackingUi)
            );
        } catch (e) {
            /* ignore */
        }
        this.emitTrackingObserveFocusChanged();
        return true;
    }

    wireSpecialTrackingUi(tableWrap, sheet) {
        this.wireTrackingUi(tableWrap, sheet);
    }

    /**
     * Save state to localStorage
     */
    save() {
        let sheetsForSave = this.sheets;
        const st = this.sheets && this.sheets[TRACKING_SHEET_ID];
        if (st && st.kind === TRACKING_KIND && (st.frames || st.series)) {
            sheetsForSave = { ...this.sheets, [TRACKING_SHEET_ID]: { kind: TRACKING_KIND, data: [] } };
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
