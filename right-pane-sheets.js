/**
 * Right Pane Sheet Manager & Styling
 * Inspired by Module1-5 VBA patterns: grouping, frequency analysis, color-coding
 */

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
        this.comboG1Enabled = false;
        this.comboH1Text = '';
        this.scrollPositions = {};
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
        this.sheets = {
            sheet1: {
                kind: 'source',
                data: this.sourceRows || []
            },
            ...comboSheets
        };
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
        comboRows.sort((left, right) => right.appear - left.appear || Number(left.combo) - Number(right.combo));

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
        specialRows.sort((left, right) => right.count - left.count || String(left.special).localeCompare(String(right.special)));

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
        const sheet = this.sheets[this.activeSheet];
        if (!sheet) {
            tableWrap.innerHTML = '<div class="sheet-empty">Không có dữ liệu. Tải dữ liệu từ data.json</div>';
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

        for (const i of rowIndices) {
            const row = displayRows[i];
            const date = row.date || row.Date || '';
            const id = row.id || row.ID || '';
            const result = row.result || row.Result || '';
            const isEmptyResultRow = this.isEmptyResultRow(row);
            const noteMeta = isEmptyResultRow ? { text: '', highlightYellow: false } : this.getComputedNoteMeta(i, row);
            const nonexistMeta = this.getNonexistMetaForSourceRow(i, row);
            const idBg = this.getIdBackgroundByFrequency(id);
            const dateBg = this.shouldHighlightDateByPairWindow(displayRows, i) ? ' style="background:#00b0f0;color:#000;font-weight:bold;"' : '';

            let resultHtml = this.highlightResultByFrequency(result);
            let noteHtml = this.renderNoteHtml(noteMeta.text, noteMeta.highlightYellow);
            const noteStyle = noteMeta.highlightYellow ? ' style="background:#ff0;"' : '';
            let nonexistHtml = this.renderNonexistHtml(i, nonexistMeta.text, result);
            const idStyle = idBg ? ` style="background:${idBg};"` : '';
            const activeClass = highlightIdx === i ? ' filter-popup-row-active' : '';

            html += `<tr data-idx="${i}" class="data-row${activeClass}" data-has-result="${!!result}" data-empty="${isEmptyResultRow ? '1' : '0'}">
                <td class="cell-date"${dateBg}>${date}</td>
                <td class="cell-id"${idStyle}>${id}</td>
                <td class="cell-result">${resultHtml}</td>
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

        if (applyWindowSelection && this.activeWindowRange) {
            const selectionRoot = options.selectionRoot || tableWrap;
            this.applyWindowSelection(
                this.activeWindowRange.start,
                this.activeWindowRange.end,
                this.activeWindowRange.target,
                selectionRoot
            );
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
                html += '<td class="cell-col-h blank-cell"></td>';
            }
            html += `<td class="cell-col-i">${isHeaderRow ? 'special' : (specialRow ? this.escapeHtml(specialRow.special || '') : '')}</td>`;
            html += `<td class="cell-col-j">${isHeaderRow ? 'count' : (specialRow ? this.escapeHtml(String(specialRow.count ?? '')) : '')}</td>`;
            html += `<td class="cell-col-k">${isHeaderRow ? '' : (specialRow && this.normalizeNumberKey(specialRow.special) === this.normalizeNumberKey(comboState.targetSpecial) ? '<span style="font-weight:800;color:rgb(0,100,0);font-family:Segoe UI Symbol;">⬆</span>' : '')}</td>`;

            html += '</tr>';
        }

        html += '</tbody></table></div>';
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
            });

            h1Input.addEventListener('change', () => {
                this.comboH1Text = h1Input.value;
                this.save();
                window.dispatchEvent(new CustomEvent('comboControlsChanged', { detail: { sheet: this.activeSheet } }));
            });
        }
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
    buildNoteForRow(rows, rowIndex, referenceCounts) {
        const currentRow = rows[rowIndex] || {};
        const currentId = this.parseRowId(currentRow.id || currentRow.ID || '');
        const currentNums = this.parseMainNums(currentRow.result || currentRow.Result || '');

        if (currentId === null || currentNums.length !== 5) {
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

            for (let a = 0; a < 4; a++) {
                for (let b = a + 1; b < 5; b++) {
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
            const nonexistMeta = this.getNonexistMetaForSourceRow(i, row);
            const result = row.result || row.Result || '';
            cell.innerHTML = this.renderNonexistHtml(i, nonexistMeta.text, result);
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

        const sheetNames = ['sheet1', 'combo_1', 'combo_2', 'combo_3', 'combo_4', 'combo_5'];
        for (const name of sheetNames) {
            if (!this.sheets[name]) {
                continue;
            }
            const tab = document.createElement('button');
            tab.className = 'sheet-tab';
            if (name === this.activeSheet) {
                tab.classList.add('active');
            }
            tab.textContent = name;
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
                if (name !== 'sheet1') {
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
     * Build Module2-style combo sheets from source rows.
     */
    buildComboSheetsFromRows(rows) {
        const dicts = [null, new Map(), new Map(), new Map(), new Map(), new Map()];
        const dictSpecial = new Map();

        for (const row of rows || []) {
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

        const latestRow = this.getLatestValidResultRow(rows || []);
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

            data.sort((left, right) => right.appear - left.appear || left.combo.localeCompare(right.combo));

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
                specialRows.sort((left, right) => right.count - left.count || left.special.localeCompare(right.special));

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
     * Save state to localStorage
     */
    save() {
        const data = {
            sheets: this.sheets,
            activeSheet: this.activeSheet,
            comboFocusRowId: this.comboFocusRowId,
            comboFocusRowIndex: this.comboFocusRowIndex,
            comboG1Enabled: this.comboG1Enabled,
            comboH1Text: this.comboH1Text,
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

// Export for use in index.html
if (typeof module !== 'undefined' && module.exports) {
    module.exports = RightPaneSheetManager;
}
