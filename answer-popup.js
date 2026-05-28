/**
 * Answer popup: phiếu trả lời gắn kỳ đang focus (sheet1), pick số từ nửa trái.
 */
const ANSWER_INITIAL_TICKET_COUNT = 3;
const ANSWER_MIN_TICKET_COUNT = 1;

class AnswerPopupController {
    constructor(deps) {
        this.deps = deps || {};
        this.open = false;
        this.focusRowIndex = -1;
        this.focusDate = '';
        this.focusId = '';
        this.tickets = [];
        this.activeTicketId = null;
        this.checked = false;
        this.answerNums = [];
        this._leftSyncSilent = false;
        this._dragTicketId = null;
        this._rowPointerDrag = null;
        this._bound = false;
        this._nextFormSeq = 0;
        this._submitCheckSyncLock = false;
    }

    allocFormId() {
        const formId = this.formIdFromSequence(this._nextFormSeq);
        this._nextFormSeq += 1;
        return formId;
    }

    formIdFromSequence(n) {
        let seq = Math.max(0, Number(n) || 0) + 1;
        let label = '';
        while (seq > 0) {
            seq -= 1;
            label = String.fromCharCode(65 + (seq % 26)) + label;
            seq = Math.floor(seq / 26);
        }
        return label;
    }

    el(id) {
        return typeof this.deps.el === 'function' ? this.deps.el(id) : document.getElementById(id);
    }

    getSheetManager() {
        return typeof this.deps.getSheetManager === 'function' ? this.deps.getSheetManager() : null;
    }

    parseNums(s) {
        if (typeof this.deps.parseNums === 'function') {
            return this.deps.parseNums(s);
        }
        return String(s || '').split(/[,\s|]+/).map(x => parseInt(x, 10)).filter(n => !isNaN(n));
    }

    isEnabled() {
        const sm = this.getSheetManager();
        return !!(sm && sm.activeSheet === 'sheet1');
    }

    isOpen() {
        return this.open;
    }

    hasActiveTicket() {
        return this.open && !!this.activeTicketId && !this.checked;
    }

    getFocusAnswerNums() {
        const sm = this.getSheetManager();
        if (!sm || this.focusRowIndex < 0) {
            return (this.answerNums || []).slice();
        }
        const row = sm.dataRows[this.focusRowIndex] || {};
        const raw = row.result || row.Result || '';
        if (typeof sm.parseMainNums === 'function') {
            return sm.parseMainNums(raw);
        }
        return this.parseNums(raw);
    }

    syncSubmitWithFocusRow(wantOn) {
        if (typeof this.deps.syncSubmitForAnswer !== 'function') {
            return;
        }
        const nums = wantOn ? this.getFocusAnswerNums() : undefined;
        this.deps.syncSubmitForAnswer(wantOn, nums, { clearManualPicks: true });
    }

    resolveOpenRowIndex(sm) {
        if (!sm) {
            return -1;
        }
        if (sm.comboFocusRowIndex >= 0) {
            return sm.comboFocusRowIndex;
        }
        if (this.focusRowIndex >= 0) {
            return this.focusRowIndex;
        }
        const range = sm.activeWindowRange;
        if (range && typeof range.target === 'number' && range.target >= 0) {
            return range.target;
        }
        const focusId = String(sm.comboFocusRowId || '').trim();
        if (focusId && Array.isArray(sm.dataRows)) {
            const byId = sm.dataRows.findIndex((r) => String(r.id || r.ID || '').trim() === focusId);
            if (byId >= 0) {
                return byId;
            }
        }
        if (Array.isArray(sm.dataRows) && sm.dataRows.length > 0) {
            return sm.dataRows.length - 1;
        }
        return -1;
    }

    clearLeftPickForAnswerOpen() {
        const frame = this.el('okFrame');
        if (!frame || !frame.contentWindow) {
            return;
        }
        this._leftSyncSilent = true;
        try {
            frame.contentWindow.postMessage({ type: 'clearLeftPickForAnswer' }, '*');
        } catch (e) { /* ignore */ }
        setTimeout(() => { this._leftSyncSilent = false; }, 120);
    }

    bindOnce() {
        if (this._bound) {
            return;
        }
        this._bound = true;
        const tableWrap = this.el('answerTableWrap');
        const addBtn = this.el('answerPopupAddBtn');
        const checkBtn = this.el('answerPopupCheckBtn');
        const clearBtn = this.el('answerPopupClearBtn');
        const closeBtn = this.el('answerPopupClose');

        if (closeBtn) {
            closeBtn.addEventListener('click', (e) => {
                e.preventDefault();
                e.stopPropagation();
                this.close();
            });
        }
        if (addBtn) {
            addBtn.addEventListener('click', (e) => {
                e.preventDefault();
                e.stopPropagation();
                this.addTicket();
            });
        }
        if (checkBtn) {
            checkBtn.addEventListener('click', (e) => {
                e.preventDefault();
                e.stopPropagation();
                this.toggleCheck();
            });
        }
        if (clearBtn) {
            clearBtn.addEventListener('click', (e) => {
                e.preventDefault();
                this.clearAllTickets();
            });
        }
        if (tableWrap) {
            tableWrap.addEventListener('click', (e) => this.onTableClick(e));
            tableWrap.addEventListener('pointerdown', (e) => {
                if (!e.target.closest('[data-action="drag"]')) {
                    return;
                }
                e.stopPropagation();
                this.onRowDragPointerDown(e);
            });
        }
        if (typeof this.deps.initDrag === 'function') {
            this.deps.initDrag();
        }
        if (typeof this.deps.initResize === 'function') {
            this.deps.initResize();
        }
    }

    toggle() {
        if (!this.isEnabled()) {
            return;
        }
        if (this.open) {
            this.close();
        } else {
            this.openPopup();
        }
    }

    openPopup() {
        if (!this.isEnabled()) {
            return;
        }
        this.bindOnce();
        const sm = this.getSheetManager();
        const idx = this.resolveOpenRowIndex(sm);
        if (idx < 0) {
            return;
        }
        this.clearLeftPickForAnswerOpen();
        this.resetForFocusRow(idx);
        this.activeTicketId = null;
        this.open = true;
        const dock = this.el('answerPopupDock');
        if (typeof this.deps.prepareLayout === 'function') {
            this.deps.prepareLayout();
        }
        if (dock) {
            dock.classList.remove('hidden');
            dock.setAttribute('aria-hidden', 'false');
        }
        this._submitCheckSyncLock = true;
        try {
            this.checked = false;
            this.syncCheckButtonUi();
            this.syncSubmitWithFocusRow(false);
        } finally {
            this._submitCheckSyncLock = false;
        }
        this.render();
    }

    close() {
        this.open = false;
        this.checked = false;
        this.syncCheckButtonUi();
        const dock = this.el('answerPopupDock');
        if (dock) {
            dock.classList.add('hidden');
            dock.setAttribute('aria-hidden', 'true');
        }
        if (typeof this.deps.saveLayout === 'function') {
            this.deps.saveLayout();
        }
        this._submitCheckSyncLock = true;
        try {
            this.syncSubmitWithFocusRow(false);
        } finally {
            this._submitCheckSyncLock = false;
        }
        this._leftSyncSilent = true;
        this.postLeftSync([]);
        setTimeout(() => { this._leftSyncSilent = false; }, 80);
    }

    onFocusRowChanged(rowIndex) {
        const idx = Number(rowIndex);
        if (!Number.isFinite(idx) || idx < 0) {
            return;
        }
        const changed = idx !== this.focusRowIndex;
        this.focusRowIndex = idx;
        const sm = this.getSheetManager();
        if (sm && sm.dataRows && sm.dataRows[idx]) {
            const row = sm.dataRows[idx];
            this.focusDate = String(row.date || row.Date || '');
            this.focusId = String(row.id || row.ID || '');
        }
        if (this.open && changed) {
            const wasChecked = this.checked;
            this.resetForFocusRow(idx);
            if (wasChecked) {
                this._submitCheckSyncLock = true;
                try {
                    this.syncSubmitWithFocusRow(false);
                } finally {
                    this._submitCheckSyncLock = false;
                }
            }
            this.render();
            if (wasChecked) {
                this.activeTicketId = null;
                this.postLeftTicketPreview([]);
            } else if (this.activeTicketId) {
                this.syncLeftFromActiveTicket();
            }
        }
    }

    resetForFocusRow(rowIndex) {
        const sm = this.getSheetManager();
        const idx = Number(rowIndex);
        if (!sm || !Number.isFinite(idx) || idx < 0) {
            return;
        }
        const row = sm.dataRows[idx] || {};
        this.focusRowIndex = idx;
        this.focusDate = String(row.date || row.Date || '');
        this.focusId = String(row.id || row.ID || '');
        this.checked = false;
        this._nextFormSeq = 0;
        this.answerNums = this.parseNums(sm.parseMainNums ? sm.parseMainNums(row.result || row.Result || '') : (row.result || ''));
        this.tickets = this.createInitialTickets();
        this.activeTicketId = this.tickets[0].id;
        this.syncCheckButtonUi();
    }

    createInitialTickets() {
        const tickets = [];
        for (let i = 0; i < ANSWER_INITIAL_TICKET_COUNT; i++) {
            tickets.push(this.createEmptyTicket());
        }
        return tickets;
    }

    ensureMinTickets() {
        while (this.tickets.length < ANSWER_MIN_TICKET_COUNT) {
            this.tickets.push(this.createEmptyTicket());
        }
    }

    createEmptyTicket() {
        const nonexist = this.buildNonexistForFocus();
        return {
            id: 'ticket-' + Date.now() + '-' + Math.random().toString(36).slice(2, 7),
            formId: this.allocFormId(),
            nums: [],
            note: '',
            nonexist: nonexist,
            matchNums: [],
            winCount: 0,
            isWin: false
        };
    }

    buildNonexistForFocus() {
        const sm = this.getSheetManager();
        if (!sm || this.focusRowIndex < 0) {
            return '';
        }
        try {
            const meta = sm.buildNonexistForRow(sm.dataRows, this.focusRowIndex);
            return meta && meta.text ? String(meta.text) : '';
        } catch (e) {
            return '';
        }
    }

    buildNoteForNums(nums) {
        const sm = this.getSheetManager();
        if (!sm || this.focusRowIndex < 0 || !nums || nums.length !== 5) {
            return '';
        }
        try {
            const rows = sm.dataRows;
            const tempRows = rows.map((r, i) => {
                if (i !== this.focusRowIndex) {
                    return r;
                }
                return Object.assign({}, r, { result: nums.join(',') });
            });
            const meta = sm.buildNoteForRow(tempRows, this.focusRowIndex, new Map());
            return meta && meta.text ? String(meta.text) : '';
        } catch (e) {
            return '';
        }
    }

    refreshTicketDerived(ticket) {
        if (!ticket) {
            return;
        }
        ticket.nonexist = this.buildNonexistForFocus();
        if (ticket.nums.length === 5) {
            ticket.note = this.buildNoteForNums(ticket.nums);
        } else {
            ticket.note = '';
        }
        if (this.checked) {
            this.applyCheckToTicket(ticket);
        }
    }

    findTicket(id) {
        return this.tickets.find(t => t.id === id) || null;
    }

    addTicket() {
        if (this.checked) {
            return;
        }
        const t = this.createEmptyTicket();
        this.tickets.push(t);
        this.activeTicketId = t.id;
        this.render();
        this.syncLeftFromActiveTicket();
    }

    removeTicket(id) {
        if (this.tickets.length <= ANSWER_MIN_TICKET_COUNT) {
            return;
        }
        const idx = this.tickets.findIndex(t => t.id === id);
        if (idx < 0) {
            return;
        }
        this.tickets.splice(idx, 1);
        if (this.activeTicketId === id) {
            this.activeTicketId = this.tickets[Math.min(idx, this.tickets.length - 1)].id;
            this.syncLeftFromActiveTicket();
        }
        this.render();
    }

    setActiveTicket(id) {
        if (!this.findTicket(id)) {
            return;
        }
        this.activeTicketId = id;
        this.render();
        this.syncLeftFromActiveTicket();
    }

    toggleNumOnActiveTicket(num) {
        if (this.checked) {
            return;
        }
        const n = parseInt(num, 10);
        if (isNaN(n)) {
            return;
        }
        const ticket = this.findTicket(this.activeTicketId);
        if (!ticket) {
            return;
        }
        const i = ticket.nums.indexOf(n);
        if (i >= 0) {
            ticket.nums.splice(i, 1);
        } else if (ticket.nums.length < 5) {
            ticket.nums.push(n);
        }
        this.checked = false;
        this.refreshTicketDerived(ticket);
        this.render();
        this.syncLeftFromActiveTicket();
    }

    applyLeftNums(nums) {
        if (this.checked || !this.activeTicketId) {
            return;
        }
        const ticket = this.findTicket(this.activeTicketId);
        if (!ticket) {
            return;
        }
        ticket.nums = (nums || [])
            .map(n => parseInt(n, 10))
            .filter(n => !isNaN(n))
            .slice(0, 5);
        this.checked = false;
        this.refreshTicketDerived(ticket);
        this.render();
    }

    clearAllTickets() {
        this.resetForFocusRow(this.focusRowIndex);
        this.render();
        this.syncLeftFromActiveTicket();
    }

    syncCheckButtonUi() {
        const btn = this.el('answerPopupCheckBtn');
        if (btn) {
            btn.classList.toggle('active', !!this.checked);
        }
    }

    toggleCheck() {
        this.setCheckedState(!this.checked);
    }

    releaseTicketFocusBeforeCheck() {
        this.activeTicketId = null;
        this._leftSyncSilent = true;
        this.postLeftTicketPreview([]);
        this.render();
    }

    setCheckedState(wantOn, options) {
        const opts = options || {};
        const syncSubmit = opts.syncSubmit !== false;
        const want = !!wantOn;
        if (want === this.checked) {
            return;
        }
        if (want) {
            const sm = this.getSheetManager();
            if (!sm || this.focusRowIndex < 0) {
                return;
            }
            this.activeTicketId = null;
            this.postLeftTicketPreview([]);
            const row = sm.dataRows[this.focusRowIndex] || {};
            this.answerNums = sm.parseMainNums(row.result || row.Result || '');
            this.checked = true;
            this.tickets.forEach((t) => this.applyCheckToTicket(t));
            this.syncCheckButtonUi();
            this.render();
            if (syncSubmit) {
                this.syncSubmitWithFocusRow(true);
            } else {
                this._leftSyncSilent = false;
            }
            setTimeout(() => {
                this.postLeftTicketPreview([]);
                if (syncSubmit) {
                    this._leftSyncSilent = false;
                }
            }, 200);
        } else {
            this.checked = false;
            this.activeTicketId = null;
            this._leftSyncSilent = true;
            this.postLeftTicketPreview([]);
            this.tickets.forEach((t) => {
                t.matchNums = [];
                t.winCount = 0;
                t.isWin = false;
            });
            this.syncCheckButtonUi();
            this.render();
            if (syncSubmit) {
                this.syncSubmitWithFocusRow(false);
            }
            setTimeout(() => {
                this.postLeftTicketPreview([]);
                this.postLeftSync([]);
                this._leftSyncSilent = false;
            }, 180);
            return;
        }
        this.syncCheckButtonUi();
        this.render();
    }

    syncCheckFromSubmit(submitOn) {
        if (!this.open || this._submitCheckSyncLock) {
            return;
        }
        const want = !!submitOn;
        if (want === this.checked) {
            return;
        }
        this._submitCheckSyncLock = true;
        try {
            this.setCheckedState(want, { syncSubmit: false });
        } finally {
            this._submitCheckSyncLock = false;
        }
    }

    applyCheckToTicket(ticket) {
        const ans = new Set(this.answerNums);
        ticket.matchNums = ticket.nums.filter(n => ans.has(n));
        ticket.winCount = ticket.matchNums.length;
        ticket.isWin = ticket.winCount >= 3 && this.answerNums.length === 5;
    }

    postLeftSync(nums) {
        const frame = this.el('okFrame');
        if (!frame || !frame.contentWindow) {
            return;
        }
        try {
            frame.contentWindow.postMessage({
                type: 'syncAnswerPickSelection',
                nums: (nums || []).slice(0, 5)
            }, '*');
        } catch (e) { /* ignore */ }
    }

    postLeftTicketPreview(nums) {
        const frame = this.el('okFrame');
        if (!frame || !frame.contentWindow) {
            return;
        }
        try {
            frame.contentWindow.postMessage({
                type: 'syncAnswerTicketPreview',
                nums: (nums || []).slice(0, 5)
            }, '*');
        } catch (e) { /* ignore */ }
    }

    syncLeftFromActiveTicket() {
        const ticket = this.findTicket(this.activeTicketId);
        this._leftSyncSilent = true;
        if (this.checked) {
            this.postLeftTicketPreview(ticket ? ticket.nums : []);
        } else if (ticket) {
            this.postLeftSync(ticket.nums);
        }
        setTimeout(() => { this._leftSyncSilent = false; }, 60);
    }

    onLeftCircledNums(nums) {
        if (!this.open) {
            return false;
        }
        if (this.checked || this._leftSyncSilent || !this.activeTicketId) {
            return true;
        }
        this.applyLeftNums(nums);
        return true;
    }

    formatResultHtml(ticket) {
        const sm = this.getSheetManager();
        const nums = ticket.nums || [];
        if (!nums.length) {
            return '<span class="answer-ticket-empty">—</span>';
        }
        const matchSet = new Set((ticket.matchNums || []).map(n => Number(n)));
        const showMatch = this.checked && this.answerNums.length === 5;
        const parts = nums.map(n => {
            const hit = showMatch && matchSet.has(n);
            const cls = 'answer-pick-num' + (hit ? ' answer-num-hit' : '');
            return `<span class="${cls}">${n}</span>`;
        });
        return parts.join(',');
    }

    escapeHtml(s) {
        return String(s || '')
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;');
    }

    formatNoteHtml(ticket) {
        const sm = this.getSheetManager();
        const text = ticket && ticket.note ? String(ticket.note) : '';
        if (!text) {
            return '<span class="answer-ticket-empty">—</span>';
        }
        if (sm && typeof sm.renderNoteHtml === 'function') {
            return sm.renderNoteHtml(text, false);
        }
        return this.escapeHtml(text);
    }

    formatNonexistHtml(ticket) {
        const sm = this.getSheetManager();
        const text = ticket && ticket.nonexist ? String(ticket.nonexist) : '';
        if (!text) {
            return '<span class="answer-ticket-empty">—</span>';
        }
        if (sm && typeof sm.renderNonexistHtml === 'function') {
            return sm.renderNonexistHtml(this.focusRowIndex, text, ticket.nums.join(','));
        }
        return this.escapeHtml(text);
    }

    render() {
        this.ensureMinTickets();
        const metaEl = this.el('answerPopupFocusMeta');
        const tableWrap = this.el('answerTableWrap');
        if (metaEl) {
            metaEl.textContent = this.focusDate && this.focusId
                ? `${this.focusDate} · ${this.focusId}`
                : (this.focusId || '—');
            metaEl.title = metaEl.textContent;
        }
        if (!tableWrap) {
            return;
        }
        tableWrap.classList.toggle('is-check-mode', !!this.checked);
        const bodyRows = this.tickets.map((t, index) => {
            const active = t.id === this.activeTicketId;
            const winCls = this.checked && t.isWin ? ' is-winning' : '';
            const activeCls = active ? ' is-active' : '';
            const resultHtml = this.formatResultHtml(t);
            const noteHtml = this.formatNoteHtml(t);
            const nonexistHtml = this.formatNonexistHtml(t);
            const canRemove = this.tickets.length > ANSWER_MIN_TICKET_COUNT && !this.checked;
            const formId = this.escapeHtml(t.formId || '');
            return `<tr class="answer-ticket-row${activeCls}${winCls}" data-ticket-id="${this.escapeHtml(t.id)}" data-index="${index}">
                <td class="cell-form-id">${formId}</td>
                <td class="cell-result answer-cell-result" data-action="pick">${resultHtml}</td>
                <td class="cell-note">${noteHtml}</td>
                <td class="cell-nonexist">${nonexistHtml}</td>
                <td class="answer-cell-btn"><button type="button" class="answer-ticket-remove" data-action="remove" title="Hủy phiếu"${canRemove ? '' : ' disabled'}>−</button></td>
                <td class="answer-cell-btn"><span class="answer-ticket-drag" data-action="drag" title="Kéo đổi thứ tự">☰</span></td>
            </tr>`;
        }).join('');
        tableWrap.innerHTML = `<table class="sheet-data-table answer-ticket-table">
            <colgroup>
                <col class="col-form-id">
                <col class="col-result">
                <col class="col-note">
                <col class="col-nonexist">
                <col class="col-action">
                <col class="col-action">
            </colgroup>
            <thead><tr>
            <th>formID</th>
            <th>result</th>
            <th>note</th>
            <th>nonexist</th>
            <th aria-label="Hủy"></th>
            <th aria-label="Kéo"></th>
        </tr></thead><tbody>${bodyRows}</tbody></table>`;
        if (typeof this.deps.applyUiScale === 'function') {
            requestAnimationFrame(() => this.deps.applyUiScale());
        }
    }

    onTableClick(e) {
        const removeBtn = e.target.closest('[data-action="remove"]');
        if (removeBtn) {
            if (this.checked) {
                return;
            }
            e.preventDefault();
            e.stopPropagation();
            const row = removeBtn.closest('tr[data-ticket-id]');
            if (row) {
                this.removeTicket(row.getAttribute('data-ticket-id'));
            }
            return;
        }
        const row = e.target.closest('tr[data-ticket-id]');
        if (!row) {
            return;
        }
        const id = row.getAttribute('data-ticket-id');
        this.setActiveTicket(id);
        if (this.checked) {
            return;
        }
        const numEl = e.target.closest('.answer-pick-num');
        if (numEl && numEl.textContent) {
            const n = parseInt(numEl.textContent, 10);
            if (!isNaN(n)) {
                this.toggleNumOnActiveTicket(n);
            }
        }
    }

    clearRowPointerDragUi() {
        const tableWrap = this.el('answerTableWrap');
        if (tableWrap) {
            tableWrap.classList.remove('is-ticket-row-dragging');
            tableWrap.querySelectorAll('.is-drag-source, .is-dragging-handle').forEach((el) => {
                el.classList.remove('is-drag-source', 'is-dragging-handle');
            });
            tableWrap.querySelectorAll('tr.answer-ticket-row').forEach((row) => {
                row.style.transform = '';
                row.style.transition = '';
                row.style.display = '';
            });
            tableWrap.querySelectorAll('.answer-ticket-drag-placeholder').forEach((el) => {
                try { el.remove(); } catch (err) { /* ignore */ }
            });
        }
        if (this._rowPointerDrag) {
            if (this._rowPointerDrag.raf) {
                cancelAnimationFrame(this._rowPointerDrag.raf);
            }
            if (this._rowPointerDrag.ghostEl) {
                try { this._rowPointerDrag.ghostEl.remove(); } catch (err) { /* ignore */ }
            }
        }
        this._rowPointerDrag = null;
        this._dragTicketId = null;
        document.body.style.userSelect = '';
        document.body.style.cursor = '';
    }

    /** Index among visible rows (0..n-1); n = append after last visible. */
    getVisibleRowInsertIndex(clientY, tableWrap) {
        const rowTargets = tableWrap.querySelectorAll(
            'tbody tr.answer-ticket-row:not(.is-drag-source):not(.answer-ticket-drag-placeholder)'
        );
        for (let i = 0; i < rowTargets.length; i++) {
            const rect = rowTargets[i].getBoundingClientRect();
            if (clientY < rect.top + rect.height * 0.5) {
                return i;
            }
        }
        return rowTargets.length;
    }

    ticketIndexToUiInsert(ticketIdx, fromIdx) {
        if (ticketIdx <= fromIdx) {
            return ticketIdx;
        }
        return ticketIdx - 1;
    }

    uiInsertToTicketIndex(uiInsert, fromIdx, ticketCount) {
        if (uiInsert >= ticketCount - 1) {
            return ticketCount;
        }
        if (uiInsert <= fromIdx) {
            return uiInsert;
        }
        return uiInsert + 1;
    }

    createRowDragPlaceholder(rowHeight) {
        const tr = document.createElement('tr');
        tr.className = 'answer-ticket-drag-placeholder';
        tr.setAttribute('aria-hidden', 'true');
        const td = document.createElement('td');
        td.colSpan = 6;
        td.style.height = `${rowHeight}px`;
        td.style.padding = '0';
        td.style.border = 'none';
        td.style.lineHeight = '0';
        td.style.verticalAlign = 'top';
        tr.appendChild(td);
        return tr;
    }

    repositionRowDragPlaceholder(drag, tableWrap) {
        const tbody = tableWrap.querySelector('tbody');
        const placeholder = drag.placeholder;
        if (!tbody || !placeholder) {
            return;
        }
        const others = tbody.querySelectorAll(
            'tr.answer-ticket-row:not(.is-drag-source):not(.answer-ticket-drag-placeholder)'
        );
        const uiInsert = drag.uiInsert;
        if (uiInsert >= others.length) {
            tbody.appendChild(placeholder);
        } else {
            tbody.insertBefore(placeholder, others[uiInsert]);
        }
    }

    copyAnswerRowGhostCellStyles(srcRow, cloneRow) {
        const props = [
            'width', 'minWidth', 'maxWidth', 'height', 'minHeight',
            'padding', 'paddingTop', 'paddingRight', 'paddingBottom', 'paddingLeft',
            'fontSize', 'fontWeight', 'lineHeight', 'textAlign', 'color', 'backgroundColor',
            'boxShadow', 'border', 'borderTop', 'borderRight', 'borderBottom', 'borderLeft',
            'whiteSpace', 'verticalAlign', 'boxSizing', 'overflow', 'textOverflow'
        ];
        const srcCells = srcRow.querySelectorAll('td');
        const cloneCells = cloneRow.querySelectorAll('td');
        srcCells.forEach((src, i) => {
            const dst = cloneCells[i];
            if (!dst) {
                return;
            }
            const cs = window.getComputedStyle(src);
            props.forEach((prop) => {
                dst.style[prop] = cs[prop];
            });
        });
        srcRow.querySelectorAll('.answer-ticket-remove, .answer-ticket-drag').forEach((srcBtn, i) => {
            const dstBtn = cloneRow.querySelectorAll('.answer-ticket-remove, .answer-ticket-drag')[i];
            if (!dstBtn) {
                return;
            }
            const cs = window.getComputedStyle(srcBtn);
            ['width', 'height', 'minWidth', 'minHeight', 'padding', 'fontSize', 'lineHeight',
                'color', 'backgroundColor', 'border', 'borderRadius', 'boxSizing'].forEach((prop) => {
                dstBtn.style[prop] = cs[prop];
            });
        });
        const srcDesc = srcRow.querySelectorAll('*');
        const cloneDesc = cloneRow.querySelectorAll('*');
        const n = Math.min(srcDesc.length, cloneDesc.length);
        for (let i = 0; i < n; i++) {
            if (srcDesc[i].tagName !== cloneDesc[i].tagName) {
                continue;
            }
            if (!srcDesc[i].className) {
                continue;
            }
            const cs = window.getComputedStyle(srcDesc[i]);
            if (cs.color) {
                cloneDesc[i].style.color = cs.color;
            }
            if (cs.fontWeight && cs.fontWeight !== 'normal' && cs.fontWeight !== '400') {
                cloneDesc[i].style.fontWeight = cs.fontWeight;
            }
        }
    }

    createRowDragGhost(row, table, tableWrap) {
        const ghost = document.createElement('div');
        ghost.className = 'answer-ticket-drag-ghost';
        const tbl = document.createElement('table');
        tbl.className = table.className;
        const colgroup = table.querySelector('colgroup');
        if (colgroup) {
            tbl.appendChild(colgroup.cloneNode(true));
        }
        const tbody = document.createElement('tbody');
        const clone = row.cloneNode(true);
        clone.classList.remove('is-drag-source');
        tbody.appendChild(clone);
        tbl.appendChild(tbody);
        ghost.appendChild(tbl);

        const tableRect = table.getBoundingClientRect();
        ghost.style.width = `${tableRect.width}px`;
        tbl.style.width = '100%';

        this.copyAnswerRowGhostCellStyles(row, clone);

        const dragHandle = clone.querySelector('[data-action="drag"]');
        if (dragHandle) {
            dragHandle.classList.add('is-dragging-handle');
        }

        tableWrap.appendChild(ghost);
        return ghost;
    }

    positionRowDragGhost(drag) {
        const ghost = drag.ghostEl;
        const placeholder = drag.placeholder;
        if (!ghost) {
            return;
        }
        let top = drag.clientY - drag.offsetY;
        let left = drag.tableLeft;
        if (placeholder && placeholder.isConnected) {
            const slot = placeholder.getBoundingClientRect();
            left = slot.left;
            if (drag.uiInsert >= drag.visibleRowCount) {
                const maxTop = slot.bottom - drag.rowHeight;
                if (top > maxTop) {
                    top = maxTop;
                }
                if (top < slot.top) {
                    top = slot.top;
                }
            }
        }
        ghost.style.transform = `translate3d(${left}px, ${top}px, 0)`;
    }

    tickRowDrag() {
        const drag = this._rowPointerDrag;
        const tableWrap = this.el('answerTableWrap');
        if (!drag || !tableWrap) {
            return;
        }
        drag.raf = 0;
        const uiInsert = this.getVisibleRowInsertIndex(drag.clientY, tableWrap);
        if (uiInsert !== drag.uiInsert) {
            drag.uiInsert = uiInsert;
            drag.ticketInsert = this.uiInsertToTicketIndex(uiInsert, drag.fromIdx, this.tickets.length);
            this.repositionRowDragPlaceholder(drag, tableWrap);
        }
        this.positionRowDragGhost(drag);
    }

    onRowDragPointerDown(e) {
        if (this.checked || e.button !== 0) {
            return;
        }
        const handle = e.target.closest('[data-action="drag"]');
        if (!handle) {
            return;
        }
        const row = handle.closest('tr[data-ticket-id]');
        const table = row && row.closest('table');
        const tableWrap = this.el('answerTableWrap');
        if (!row || !table || !tableWrap) {
            return;
        }
        e.preventDefault();
        e.stopPropagation();

        const ticketId = row.getAttribute('data-ticket-id');
        const fromIdx = this.tickets.findIndex((t) => t.id === ticketId);
        if (fromIdx < 0) {
            return;
        }

        const rowRect = row.getBoundingClientRect();
        const rowHeight = row.offsetHeight || rowRect.height || 32;
        const offsetX = e.clientX - rowRect.left;
        const offsetY = e.clientY - rowRect.top;
        const ghostEl = this.createRowDragGhost(row, table, tableWrap);
        const placeholder = this.createRowDragPlaceholder(rowHeight);
        const tbody = table.querySelector('tbody');

        row.classList.add('is-drag-source');
        handle.classList.add('is-dragging-handle');
        tableWrap.classList.add('is-ticket-row-dragging');
        document.body.style.userSelect = 'none';
        document.body.style.cursor = 'grabbing';

        this._dragTicketId = ticketId;
        tbody.appendChild(placeholder);

        const visibleRowCount = this.tickets.length - 1;
        const uiInsert = this.ticketIndexToUiInsert(fromIdx, fromIdx);
        this._rowPointerDrag = {
            ticketId,
            fromIdx,
            uiInsert,
            ticketInsert: fromIdx,
            visibleRowCount,
            ghostEl,
            placeholder,
            rowHeight,
            offsetX,
            offsetY,
            tableLeft: rowRect.left,
            clientX: e.clientX,
            clientY: e.clientY,
            raf: 0
        };
        this.repositionRowDragPlaceholder(this._rowPointerDrag, tableWrap);
        this.positionRowDragGhost(this._rowPointerDrag);

        const move = (ev) => this.onRowDragPointerMove(ev);
        const up = (ev) => {
            handle.releasePointerCapture(ev.pointerId);
            document.removeEventListener('pointermove', move);
            document.removeEventListener('pointerup', up);
            document.removeEventListener('pointercancel', up);
            this.onRowDragPointerUp();
        };
        try {
            handle.setPointerCapture(e.pointerId);
        } catch (err) { /* ignore */ }
        document.addEventListener('pointermove', move);
        document.addEventListener('pointerup', up);
        document.addEventListener('pointercancel', up);
    }

    onRowDragPointerMove(e) {
        const drag = this._rowPointerDrag;
        if (!drag) {
            return;
        }
        drag.clientX = e.clientX;
        drag.clientY = e.clientY;
        if (drag.raf) {
            return;
        }
        drag.raf = requestAnimationFrame(() => this.tickRowDrag());
    }

    onRowDragPointerUp() {
        const drag = this._rowPointerDrag;
        if (!drag) {
            return;
        }
        if (drag.raf) {
            cancelAnimationFrame(drag.raf);
            drag.raf = 0;
        }
        const tableWrap = this.el('answerTableWrap');
        if (tableWrap) {
            const uiInsert = this.getVisibleRowInsertIndex(drag.clientY, tableWrap);
            drag.uiInsert = uiInsert;
            drag.ticketInsert = this.uiInsertToTicketIndex(uiInsert, drag.fromIdx, this.tickets.length);
        }
        const fromIdx = drag.fromIdx;
        const ticketInsert = drag.ticketInsert;
        this.clearRowPointerDragUi();
        if (fromIdx !== ticketInsert) {
            const item = this.tickets.splice(fromIdx, 1)[0];
            let to = ticketInsert;
            if (fromIdx < to) {
                to -= 1;
            }
            this.tickets.splice(Math.max(0, Math.min(to, this.tickets.length)), 0, item);
            this.render();
        }
    }
}

if (typeof module !== 'undefined' && module.exports) {
    module.exports = AnswerPopupController;
}
