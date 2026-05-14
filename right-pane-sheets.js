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
     * Build the current combo_1 focus, arrow, and styling state.
     */
    buildCombo1StyleContext() {
        const sourceRows = this.sourceRows || [];
        const fallbackRow = this.getLatestValidResultRow(sourceRows);
        const focusRow = this.getSourceRowById(this.comboFocusRowId) || fallbackRow;
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
        const showArrowsForTarget = !!this.comboG1Enabled;
        const targetHasResult = !!String(targetResult || '').trim();
        const useH1Sim = h1Nums.length > 0 && (!showArrowsForTarget || !targetHasResult);
        const arrowNums = showArrowsForTarget && targetHasResult ? targetNums : h1Nums.slice(0, 5);
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
     * Render the raw five-column source sheet.
     */
    renderSourceSheet(tableWrap, rows) {
        this.bindSourceSheetKeyboardNavigation(tableWrap);

        let html = '<table class="sheet-data-table"><thead><tr><th>date</th><th>id</th><th>result</th><th>note</th><th>nonexist</th></tr></thead><tbody>';

        const displayRows = rows || [];
        for (let i = 0; i < displayRows.length; i++) {
            const row = displayRows[i];
            const date = row.date || row.Date || '';
            const id = row.id || row.ID || '';
            const result = row.result || row.Result || '';
            const isEmptyResultRow = this.isEmptyResultRow(row);
            const noteMeta = isEmptyResultRow ? { text: '', highlightYellow: false } : this.getComputedNoteMeta(i, row);
            const nonexistMeta = this.getNonexistMetaForSourceRow(i, row);
            const idBg = this.getIdBackgroundByFrequency(id);
            const dateBg = this.shouldHighlightDateByPairWindow(displayRows, i) ? ' style="background:#00b0f0;color:#000;font-weight:bold;"' : '';

            // Build result HTML with frequency coloring (inspired by Module2 highlighting)
            let resultHtml = this.highlightResultByFrequency(result);
            let noteHtml = this.renderNoteHtml(noteMeta.text, noteMeta.highlightYellow);
            const noteStyle = noteMeta.highlightYellow ? ' style="background:#ff0;"' : '';
            let nonexistHtml = this.renderNonexistHtml(i, nonexistMeta.text, result);
            const idStyle = idBg ? ` style="background:${idBg};"` : '';

            html += `<tr data-idx="${i}" class="data-row" data-has-result="${!!result}" data-empty="${isEmptyResultRow ? '1' : '0'}">
                <td class="cell-date"${dateBg}>${date}</td>
                <td class="cell-id"${idStyle}>${id}</td>
                <td class="cell-result">${resultHtml}</td>
                <td class="cell-note"${noteStyle}>${noteHtml}</td>
                <td class="cell-nonexist">${nonexistHtml}</td>
            </tr>`;
        }
        html += '</tbody></table>';
        tableWrap.innerHTML = html;

        // Attach click handlers
        tableWrap.querySelectorAll('tbody tr').forEach(tr => {
            tr.style.cursor = 'pointer';
            tr.addEventListener('click', (e) => {
                this.onRowClick(Number(tr.dataset.idx), tr.dataset.empty === '1', e);
                try {
                    tableWrap.focus({ preventScroll: true });
                } catch (err) {
                    // ignore focus failures
                }
            });
        });

        if (this.activeWindowRange) {
            this.applyWindowSelection(this.activeWindowRange.start, this.activeWindowRange.end, this.activeWindowRange.target);
        }
    }

    /**
     * Enable Enter-to-select-next-row keyboard navigation on source sheets.
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
            const key = String(event.key || '').toLowerCase();
            const isStepForward = key === 'arrowdown' || key === 'arrowright' || key === 's' || key === 'd';
            const isStepBackward = key === 'arrowup' || key === 'arrowleft' || key === 'w' || key === 'a';

            // Handle Space to toggle submit on iframe
            if (event.code === 'Space') {
                event.preventDefault();
                const frame = document.getElementById('okFrame');
                if (frame && frame.contentWindow) {
                    frame.contentWindow.postMessage({ type: 'toggleSubmit' }, '*');
                }
                return;
            }

            if (!isStepForward && !isStepBackward) {
                return;
            }

            const target = event.target;
            const tag = target && target.tagName ? String(target.tagName).toUpperCase() : '';
            if (target && (target.isContentEditable || tag === 'INPUT' || tag === 'TEXTAREA' || tag === 'SELECT' || tag === 'BUTTON')) {
                return;
            }

            const activeEl = document.activeElement;
            if (activeEl && activeEl !== tableWrap && !tableWrap.contains(activeEl)) {
                return;
            }

            const activeSheetMeta = this.sheets[this.activeSheet] || {};
            if (activeSheetMeta.kind === 'combo') {
                return;
            }

            const displayRows = this.dataRows || [];
            if (displayRows.length === 0) {
                return;
            }

            const currentIdx = this.activeWindowRange && typeof this.activeWindowRange.target === 'number'
                ? this.activeWindowRange.target
                : -1;
            if (currentIdx < 0) {
                return;
            }

            const step = isStepForward ? 1 : -1;
            const nextIdx = Math.max(0, Math.min(displayRows.length - 1, currentIdx + step));
            if (nextIdx === currentIdx) {
                return;
            }

            const nextRow = tableWrap.querySelector(`tbody tr[data-idx="${nextIdx}"]`);
            if (!nextRow) {
                return;
            }

            event.preventDefault();
            nextRow.click();
            this.centerActiveWindowInView(tableWrap);
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
                html += `<td class="cell-col-f combo-logic-focus-cell"><input id="comboF1CellInput" class="combo-cell-input" type="text" value="${this.escapeHtml(String(latestId || ''))}" aria-label="F1" /></td>`;
                html += `<td class="cell-col-g" style="background:#f8fafc;"><label style="display:flex;align-items:center;justify-content:center;width:100%;height:100%;"><input id="comboG1CellToggle" class="combo-cell-toggle" type="checkbox" aria-label="G1" ${this.comboG1Enabled ? 'checked' : ''} /></label></td>`;
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
                this.comboG1Enabled = !!g1Toggle.checked;
                this.save();
                window.dispatchEvent(new CustomEvent('comboControlsChanged', { detail: { sheet: this.activeSheet } }));
            });
        }

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
     * Remove the black border from the previously selected 11-row window.
     */
    clearWindowSelection() {
        const tableWrap = document.getElementById('tableWrap');
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
    applyWindowSelection(startIdx, endIdx, targetIdx = null) {
        const tableWrap = document.getElementById('tableWrap');
        if (!tableWrap) {
            return;
        }

        this.clearWindowSelection();

        if (startIdx === null || endIdx === null || endIdx < startIdx) {
            this.activeWindowRange = null;
            this.refreshNonexistCellsForActiveWindow(tableWrap);
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

        this.activeWindowRange = { start: startIdx, end: endIdx, target: targetIdx };
        this.renderWindowLabels(startIdx, endIdx);
        this.refreshNonexistCellsForActiveWindow(tableWrap);
    }

    /**
     * Draw 10 inline labels inside the right side of the selected window rows.
     * Mirrors the VBA WinLabel_01..10 placement at the right of the block.
     */
    renderWindowLabels(startIdx, endIdx) {
        const tableWrap = document.getElementById('tableWrap');
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
        if (!this.noteCache || this.noteCache.length !== this.dataRows.length) {
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
        if (!this.nonexistCache || this.nonexistCache.length !== this.dataRows.length) {
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
     * True when rowIndex is in the WinLabel row range for the active window and `num`
     * appears in the bottom row (chuỗi 11) nonexist — yellow nonexist gets x1.5 in that case.
     */
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
     * Re-render nonexist column so yellow x1.5 tracks the active sliding window (chuỗi 11 nonexist).
     */
    refreshNonexistCellsForActiveWindow(tableWrap) {
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
        for (let i = 0; i < displayRows.length; i++) {
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
     * Render nonexist text using the generated values from result data only.
     */
    renderNonexistHtml(rowIndex, nonexistText, currentResult) {
        if (!nonexistText || nonexistText === 'N/A') {
            return this.escapeHtml(nonexistText || '');
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

        return this.escapeHtml(nonexistText).replace(/\b\d+\b/g, (match) => {
            const value = parseInt(match, 10);
            const valueText = String(value);
            const isInDiff = !prevNonexistNums.has(value);
            const isMatch = currentNums.has(value);
            const isLongest = longestSet.has(valueText);

            if (isLongest) {
                if (isMatch) {
                    return `<span style="color:rgb(0,80,0);font-weight:bold;text-decoration:underline;font-size:1.5em">${value}</span>`;
                }
                return `<span style="color:rgb(180,30,30);font-weight:bold">${value}</span>`;
            }

            if (isMatch && isInDiff) {
                return `<span style="color:rgb(0,80,0);font-weight:bold;font-style:italic;font-size:1.5em">${value}</span>`;
            }

            if (isInDiff) {
                const boost = this.shouldBoostYellowNonexistForWindow(rowIndex, value);
                const fs = boost ? 'font-size:1.5em;' : '';
                return `<span style="color:rgb(240,200,64);font-weight:bold;${fs}">${value}</span>`;
            }

            if (isMatch) {
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
    onRowClick(idx, isEmptyRow, event) {
        const start = Math.max(0, idx - 10);
        const slice = isEmptyRow ? this.dataRows.slice(start, idx) : this.dataRows.slice(start, idx + 1);
        const lines = slice.map((r, offset) => {
            const res = r.result || r.Result || '';
            const noteMeta = this.getComputedNoteMeta(start + offset, r);
            const note = noteMeta.text;
            const nonexist = this.isEmptyResultRow(r) ? '' : this.getComputedNonexistMeta(start + offset, r).text;
            return [res, note, nonexist].filter(Boolean).join('\t');
        });

        const rowAtClick = this.dataRows[idx] || {};
        const clickedRowId = String(rowAtClick.id || rowAtClick.ID || '').trim();

        // Send to parent/iframe
        const frame = document.getElementById('okFrame');
        if (frame && frame.contentWindow) {
            frame.contentWindow.postMessage({
                type: 'setLines',
                lines: lines,
                disableSubmit: !!isEmptyRow,
                sheetName: this.activeSheet,
                focusRowId: clickedRowId
            }, '*');
        }

        // Update selectedLines
        this.selectedLines = slice.map((r, offset) => {
            const noteMeta = this.getComputedNoteMeta(start + offset, r);
            const nonexistMeta = this.isEmptyResultRow(r) ? { text: '' } : this.getComputedNonexistMeta(start + offset, r);
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

        if (this.activeSheet === 'sheet1' && this.selectedLines.length > 0) {
            const focusCandidate = this.selectedLines[this.selectedLines.length - 1] || {};
            this.comboFocusRowId = focusCandidate.id || '';
        }

        this.save();

        const windowEnd = idx;
        const targetIdx = idx;
        this.applyWindowSelection(start, windowEnd, targetIdx);

        const tableWrap = document.getElementById('tableWrap');
        const activeSheetMeta = this.sheets[this.activeSheet] || {};
        if (tableWrap && activeSheetMeta.kind !== 'combo') {
            this.centerActiveWindowInView(tableWrap);
        }

        // Dispatch custom event with selected lines
        window.dispatchEvent(new CustomEvent('rowClicked', {
            detail: {
                selectedLines: this.selectedLines,
                selectedNums: this.parseNums(this.selectedLines.length > 0 ? (this.selectedLines[this.selectedLines.length - 1].result || '') : ''),
                sheetName: this.activeSheet,
                clickedRowId
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
        this.refreshDerivedState();
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
