/**
 * 基金資料篩選器 — 匯入 → 基金對照 → 彙整校對 → 匯出 Checklist
 */

const ExcelViewer = (() => {
    'use strict';

    const VALID_FILE_EXTENSIONS = ['.xls', '.xlsx', '.ods'];
    const MAX_ROWS_PER_SHEET = 10000; // ponytail: 每張表只讀前 1 萬列，超大檔需要時再放寬

    const state = {
        isProcessing: false,
        rawSheets: [],          // 原始工作表快取 [{filename, label, startRow, startCol, sheet, jsonData}]，供重新套用清洗設定
        undoStack: [],

        fundSortOrder: [],
        fundAliasMap: {},

        // 基金對照 & 彙整
        fileInfos: [],          // 每個檔案的辨識結果 [{filename, detectedFunds, headers, rows}]
        aggregatedRows: [],     // 彙整後的逐列資料
        aggSourceFiles: [],     // 彙整時勾選的檔名，用來逐檔勾稽保留列數
        matchColumn: '1',       // 比對基金名稱的欄位（1/2/3）

        // 自訂比對模式
        compareMode: 'standard',   // 'standard' | 'custom'
        customFundList: [],        // 使用者貼入的自訂名稱清單
    };

    const elements = {};
    let dropAreaInitialHtml = '';

    const utils = {
        debounce(fn, ms) {
            let timer;
            return (...args) => { clearTimeout(timer); timer = setTimeout(() => fn(...args), ms); };
        },
        parsePositionString(str) {
            const indices = new Set();
            str.split(',').map(p => p.trim()).filter(Boolean).forEach(part => {
                if (part.includes('-')) {
                    const [s, e] = part.split('-').map(Number);
                    if (!isNaN(s) && !isNaN(e) && s <= e) {
                        for (let i = s; i <= e; i++) indices.add(i - 1);
                    }
                } else {
                    const n = Number(part);
                    if (!isNaN(n)) indices.add(n - 1);
                }
            });
            return Array.from(indices).sort((a, b) => a - b);
        },
        calculateColumnWidths(data) {
            if (!data.length) return [];
            return data[0].map((_, col) => ({
                wch: Math.min(50, Math.max(10, ...data.map(r => (r[col] ? String(r[col]).length : 0))) + 2),
            }));
        },
        // 匯出時把「看起來是數字」的字串轉成數值；保留前導 0 的代碼（如 00123）為文字
        toNumberOrNull(val) {
            if (typeof val !== 'string' || !/^-?[\d,]*\.?\d+$/.test(val)) return null;
            const clean = val.replace(/,/g, '');
            if (/^-?0\d/.test(clean)) return null;
            const n = Number(clean);
            return isNaN(n) ? null : n;
        }
    };

    const undoManager = {
        push(description, restoreFn) {
            state.undoStack.push({ description, restoreFn });
            if (state.undoStack.length > 5) state.undoStack.shift();
            this.showToast(description);
        },
        undoLast() {
            const action = state.undoStack.pop();
            if (!action) return;
            action.restoreFn();
            this.hideToast();
        },
        showToast(desc) {
            if (!elements.undoToast || !elements.undoText) return;
            elements.undoText.textContent = `已${desc}`;
            elements.undoToast.classList.add('show');
            clearTimeout(this.timer);
            this.timer = setTimeout(() => this.hideToast(), 8000);
        },
        hideToast() {
            if (elements.undoToast) elements.undoToast.classList.remove('show');
        }
    };

    function cacheElements() {
        const mapping = {
            fileInput: 'file-input', dropArea: 'drop-area', dropAreaContent: 'drop-area-content',
            specificSheetNameGroup: 'specific-sheet-name-group', specificSheetNameInput: 'specific-sheet-name-input',
            specificSheetPositionGroup: 'specific-sheet-position-group', specificSheetPositionInput: 'specific-sheet-position-input',
            backToTopBtn: 'back-to-top-btn',

            skipTopRowsCheckbox: 'skip-top-rows-checkbox', discardRowsInput: 'discard-rows-input',
            headerRowsInput: 'header-rows-input', removeEmptyRowsCheckbox: 'remove-empty-rows-checkbox',
            removeKeywordRowsCheckbox: 'remove-keyword-rows-checkbox', removeKeywordRowsInput: 'remove-keyword-rows-input',
            reapplyBanner: 'reapply-banner', reapplySettingsBtn: 'reapply-settings-btn',
            undoToast: 'undo-toast', undoText: 'undo-text', undoBtn: 'undo-btn',

            // Step 2: 基金對照 & 彙整
            stepMapping: 'step-mapping',
            mappingTbody: 'mapping-tbody', mapCheckAll: 'map-check-all',
            mapSelectAllBtn: 'map-select-all-btn', mapDeselectAllBtn: 'map-deselect-all-btn', mapInvertBtn: 'map-invert-btn',
            executeMatchBtn: 'execute-match-btn', exportMatchedBtn: 'export-matched-btn',
            aggregatePanel: 'aggregate-panel', aggregateThead: 'aggregate-thead', aggregateTbody: 'aggregate-tbody',
            aggStatTotal: 'agg-stat-total', aggStatReview: 'agg-stat-review', aggStatChecked: 'agg-stat-checked',
            aggReconcile: 'agg-reconcile',
            aggCheckReviewBtn: 'agg-check-review-btn', aggUncheckAllBtn: 'agg-uncheck-all-btn',
            aggDeleteCheckedBtn: 'agg-delete-checked-btn', aggExportBtn: 'agg-export-btn',
            matchColSelect: 'match-col-select',
            toggleImportSettings: 'toggle-import-settings', importSettingsPanel: 'import-settings-panel',
            topBarActions: 'top-bar-actions', statTableCount: 'stat-table-count', statRowCount: 'stat-row-count',
            clearAllBtn: 'clear-all-btn',

            // 自訂比對
            tabStandard: 'tab-standard', tabCustom: 'tab-custom',
            customListPanel: 'custom-list-panel',
            customListTextarea: 'custom-list-textarea',
            applyCustomListBtn: 'apply-custom-list-btn',
            clearCustomListBtn: 'clear-custom-list-btn',
            customListCount: 'custom-list-count',
            customListNum: 'custom-list-num'
        };

        Object.keys(mapping).forEach(key => { elements[key] = document.getElementById(mapping[key]); });
    }

    // ===== 自訂比對：模式切換 =====
    function switchCompareMode(mode) {
        state.compareMode = mode;
        if (elements.tabStandard) elements.tabStandard.classList.toggle('active', mode === 'standard');
        if (elements.tabCustom) elements.tabCustom.classList.toggle('active', mode === 'custom');
        if (elements.customListPanel) elements.customListPanel.classList.toggle('hidden', mode !== 'custom');
        // 已匯入檔案時，依新模式重新辨識
        redetectFunds();
    }

    // ===== 自訂比對：套用清單 =====
    function applyCustomList() {
        const raw = elements.customListTextarea ? elements.customListTextarea.value : '';
        const list = raw.split('\n').map(line => line.trim()).filter(Boolean);
        state.customFundList = list;

        if (elements.customListNum) elements.customListNum.textContent = list.length;
        if (elements.customListCount) elements.customListCount.classList.toggle('hidden', list.length === 0);
        redetectFunds();
    }

    async function loadFundConfig() {
        try {
            const response = await fetch(`fund-config.json?v=${Date.now()}`);
            if (!response.ok) return;
            const config = await response.json();
            if (config.sortOrder && config.aliasMap) {
                state.fundSortOrder = config.sortOrder;
                state.fundAliasMap = config.aliasMap;
            }
        } catch (err) { console.warn('設定檔讀取失敗', err); }
    }

    // --- 基金辨識引擎 ---
    let _fundSearchKeysCache = null;
    function getFundSearchKeys() {
        if (_fundSearchKeysCache) return _fundSearchKeysCache;
        const items = [];
        const seen = new Set();
        const add = (key, fund) => { if (!key || seen.has(key)) return; items.push({ key, fund }); seen.add(key); };

        Object.keys(state.fundAliasMap).forEach(alias => add(alias, state.fundAliasMap[alias]));
        const suffixes = ['作業基金', '校務基金', '基金'];
        state.fundSortOrder.forEach(fund => {
            add(fund, fund);
            for (const sfx of suffixes) {
                if (fund.endsWith(sfx) && fund.length > sfx.length + 2) { add(fund.slice(0, -sfx.length), fund); break; }
            }
        });
        items.sort((a, b) => b.key.length - a.key.length);
        _fundSearchKeysCache = items;
        return items;
    }

    function detectFundsInText(text) {
        if (!text) return [];

        // 自訂模式：逐一查找檔名是否包含該名稱
        if (state.compareMode === 'custom') {
            const lower = text.toLowerCase();
            return state.customFundList.filter(name => lower.includes(name.toLowerCase()));
        }

        // 標準模式：由長到短比對，已命中的字段遮蔽，避免短名稱重複命中
        let scan = text.toLowerCase();
        const found = new Set();
        for (const { key, fund } of getFundSearchKeys()) {
            const k = key.toLowerCase();
            let idx = scan.indexOf(k);
            while (idx !== -1) {
                found.add(fund);
                scan = scan.slice(0, idx) + '\u0000'.repeat(k.length) + scan.slice(idx + k.length);
                idx = scan.indexOf(k);
            }
        }
        return state.fundSortOrder.filter(f => found.has(f));
    }

    // 判斷儲存格文字是否屬於某標準基金
    // 自訂模式：僅精確比對（trim）
    // 標準模式：精確比對 → 別名比對 → 前4字+字數相同模糊比對
    function cellMatchesFund(cellText, standardFund) {
        if (!cellText || !standardFund) return false;
        const cell = cellText.trim();
        const fund = standardFund.trim();
        if (cell === fund) return true;

        if (state.compareMode === 'custom') return false;

        if (state.fundAliasMap[cell] === standardFund) return true;

        const cellChars = Array.from(cell);
        const fundChars = Array.from(fund);
        return cellChars.length >= 4 &&
               cellChars.length === fundChars.length &&
               cellChars.slice(0, 4).join('') === fundChars.slice(0, 4).join('');
    }

    function escHtml(str) {
        return String(str ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
    }

    // 取得當前模式下的有效基金清單（標準模式用 fundSortOrder；自訂模式用 customFundList）
    function getActiveFundList() {
        return state.compareMode === 'custom' ? state.customFundList : state.fundSortOrder;
    }

    // --- Stage 1: 基金對照總覽 (mapping table) ---
    function buildMappingTable() {
        if (!elements.mappingTbody) return;

        elements.mappingTbody.innerHTML = state.fileInfos.map((info, idx) => {
            const fundNames = info.detectedFunds.length > 0
                ? info.detectedFunds.map(f => `<span style="color:#10b981;">${escHtml(f)}</span>`).join('、')
                : '<span style="color:#f59e0b;">未辨識</span>';
            return `<tr>
                <td class="map-col-check"><input type="checkbox" class="map-row-check" data-index="${idx}" ${info.detectedFunds.length > 0 ? 'checked' : ''}></td>
                <td class="map-col-num">${idx + 1}</td>
                <td class="map-col-file" title="${escHtml(info.filename)}">${escHtml(info.filename)}</td>
                <td class="map-col-fund">${fundNames}</td>
                <td class="map-col-rows">${info.rows.length}</td>
            </tr>`;
        }).join('');

        const hasFiles = state.fileInfos.length > 0;
        if (elements.stepMapping) elements.stepMapping.classList.toggle('hidden', !hasFiles);
        if (elements.topBarActions) elements.topBarActions.classList.toggle('hidden', !hasFiles);
        if (elements.statTableCount) elements.statTableCount.textContent = state.fileInfos.length;
        if (elements.statRowCount) elements.statRowCount.textContent = state.fileInfos.reduce((s, f) => s + f.rows.length, 0);
    }

    // --- Stage 2: 彙整面板 (Aggregation) ---
    // 依有效基金清單順序，把勾選檔案的所有資料列貼入（一個檔案對應多個基金就貼多次）
    function buildAggregation({ silent = false } = {}) {
        const selectedIdxs = elements.mappingTbody
            ? Array.from(elements.mappingTbody.querySelectorAll('.map-row-check:checked'), cb => parseInt(cb.dataset.index, 10))
            : [];
        if (selectedIdxs.length === 0) { if (!silent) alert('請先勾選要彙整的檔案。'); return; }

        if (state.compareMode === 'custom' && state.customFundList.length === 0) {
            if (!silent) alert('自訂比對模式：請先在上方貼入名稱清單並點擊「套用清單」。');
            return;
        }

        const matchColIdx = parseInt(state.matchColumn, 10) - 1;
        const aggregated = [];

        for (const standardFund of getActiveFundList()) {
            for (const idx of selectedIdxs) {
                const info = state.fileInfos[idx];
                if (!info || !info.detectedFunds.includes(standardFund)) continue;

                info.rows.forEach(cells => {
                    const needsReview = !cellMatchesFund((cells[matchColIdx] || '').trim(), standardFund);
                    aggregated.push({
                        standardFund,
                        sourceFile: info.filename,
                        headers: info.headers,
                        cells,
                        needsReview,
                        checked: needsReview,
                    });
                });
            }
        }

        state.aggregatedRows = aggregated;
        state.aggSourceFiles = selectedIdxs.map(i => state.fileInfos[i].filename);
        renderAggregationPanel();
    }

    function renderAggregationPanel() {
        if (!elements.aggregatePanel || !elements.aggregateTbody || !elements.aggregateThead) return;
        if (!state.aggregatedRows.length) {
            elements.aggregatePanel.classList.add('hidden');
            return;
        }
        elements.aggregatePanel.classList.remove('hidden');

        const matchIdx = parseInt(state.matchColumn, 10) - 1;
        elements.aggregateThead.innerHTML = `<tr>
            <th style="width:40px;"><input type="checkbox" id="agg-check-all-cb"></th>
            <th>#</th>
            <th>標準基金</th>
            <th>來源檔案</th>
            <th>第一欄</th>
            <th>第二欄</th>
            <th>第三欄</th>
        </tr>`;

        elements.aggregateTbody.innerHTML = state.aggregatedRows.map((row, idx) => {
            const cls = `${row.needsReview ? 'row-needs-review' : ''} ${row.checked ? 'row-checked-delete' : ''}`;
            const colCells = [0, 1, 2].map(c =>
                `<td class="${c === matchIdx ? 'match-col' : ''}">${escHtml((row.cells[c] || '').trim())}</td>`).join('');
            return `<tr data-agg-index="${idx}" class="${cls}">
                <td><input type="checkbox" class="agg-row-check" ${row.checked ? 'checked' : ''}></td>
                <td>${idx + 1}</td>
                <td>${escHtml(row.standardFund)}</td>
                <td title="${escHtml(row.sourceFile)}" style="max-width:300px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;">${escHtml(row.sourceFile)}</td>
                ${colCells}
            </tr>`;
        }).join('');
        updateAggregationCounts();
    }

    function updateAggregationCounts() {
        if (elements.aggStatTotal) elements.aggStatTotal.textContent = state.aggregatedRows.length;
        if (elements.aggStatReview) elements.aggStatReview.textContent = state.aggregatedRows.filter(r => r.needsReview).length;
        if (elements.aggStatChecked) elements.aggStatChecked.textContent = state.aggregatedRows.filter(r => r.checked).length;
        renderReconcile();
    }

    // 逐檔勾稽：每個勾選的檔案最後保留幾列（未勾選刪除的列），0 列的檔案要回頭查
    function renderReconcile() {
        if (!elements.aggReconcile) return;
        const kept = {};
        state.aggregatedRows.forEach(r => { if (!r.checked) kept[r.sourceFile] = (kept[r.sourceFile] || 0) + 1; });
        const files = state.aggSourceFiles;
        const empty = files.filter(f => !kept[f]);
        const total = Object.values(kept).reduce((a, b) => a + b, 0);
        elements.aggReconcile.classList.toggle('reconcile-warn', empty.length > 0);
        elements.aggReconcile.innerHTML = `勾稽：勾選 <strong>${files.length}</strong> 個檔，保留 <strong>${total}</strong> 列。`
            + (empty.length
                ? `<br>⚠️ 下列 <strong>${empty.length}</strong> 個檔沒有保留任何列，請確認檔內基金名稱或比對欄位：<br>${empty.map(escHtml).join('<br>')}`
                : '每個檔都至少保留 1 列。');
    }

    function deleteCheckedAggregatedRows() {
        const backup = state.aggregatedRows;
        state.aggregatedRows = backup.filter(r => !r.checked);
        const removed = backup.length - state.aggregatedRows.length;
        if (removed === 0) { alert('沒有勾選任何列。'); return; }
        renderAggregationPanel();
        undoManager.push(`刪除 ${removed} 筆彙整資料`, () => { state.aggregatedRows = backup; renderAggregationPanel(); });
    }

    // --- Checklist 匯出 ---
    function exportAggregatedRows() {
        const validRows = state.aggregatedRows.filter(r => !r.checked);

        // 收集所有表頭聯集
        const unionHeaders = Array.from(new Set(validRows.flatMap(r => r.headers)));
        const exportData = [['標準基金', '繳交狀態', '來源檔案', ...unionHeaders]];

        const dataByFund = {};
        validRows.forEach(row => { (dataByFund[row.standardFund] ||= []).push(row); });

        // ⭐ Checklist 核心：巡覽有效基金清單，沒資料的基金也列出
        getActiveFundList().forEach(standardFund => {
            const fundRows = dataByFund[standardFund];
            if (fundRows) {
                fundRows.forEach(row => {
                    const cellMap = {};
                    row.headers.forEach((h, i) => { cellMap[h] = row.cells[i] ?? ''; });
                    exportData.push([standardFund, '✅ 已回傳', row.sourceFile, ...unionHeaders.map(h => cellMap[h] ?? '')]);
                });
            } else {
                exportData.push([standardFund, '❌ 未回傳/無資料', '-', ...unionHeaders.map(() => '')]);
            }
        });

        const ws = XLSX.utils.aoa_to_sheet(exportData);

        // 設定數值格式
        const range = XLSX.utils.decode_range(ws['!ref']);
        for (let R = 1; R <= range.e.r; R++) {
            for (let C = 0; C <= range.e.c; C++) {
                const cell = ws[XLSX.utils.encode_cell({ r: R, c: C })];
                const num = cell ? utils.toNumberOrNull(cell.v) : null;
                if (num !== null) { cell.t = 'n'; cell.v = num; }
            }
        }

        ws['!cols'] = utils.calculateColumnWidths(exportData);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, '基金彙整結果');
        XLSX.writeFile(wb, `基金資料_Checklist_${new Date().toISOString().slice(0, 10)}.xlsx`);
    }

    function setupDragAndDrop() {
        if (!elements.dropArea || !elements.fileInput) return;
        elements.dropArea.addEventListener('click', e => {
            if (e.target === elements.fileInput) return;
            elements.fileInput.click();
        });
        ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
            elements.dropArea.addEventListener(eventName, e => { e.preventDefault(); e.stopPropagation(); });
        });
        ['dragenter', 'dragover'].forEach(e => elements.dropArea.addEventListener(e, () => elements.dropArea.classList.add('highlight')));
        ['dragleave', 'drop'].forEach(e => elements.dropArea.addEventListener(e, () => elements.dropArea.classList.remove('highlight')));

        elements.dropArea.addEventListener('drop', e => {
            if (e.dataTransfer.files && e.dataTransfer.files.length > 0) processFiles(e.dataTransfer.files);
        });
        elements.fileInput.addEventListener('change', e => {
            if (e.target.files && e.target.files.length > 0) processFiles(e.target.files);
        });
    }

    // 🚀 清洗：裁掉格式造成的假空白欄、隱藏欄列，壓縮表頭，移除空白/關鍵字列
    function applyPreprocessing(jsonData, sheet, startRow, startCol) {
        const useHeaderCompression = elements.skipTopRowsCheckbox && elements.skipTopRowsCheckbox.checked;
        const discardCount = useHeaderCompression && elements.discardRowsInput ? parseInt(elements.discardRowsInput.value, 10) || 0 : 0;
        const headerCount = useHeaderCompression && elements.headerRowsInput ? parseInt(elements.headerRowsInput.value, 10) || 1 : 1;
        const totalSkipCount = useHeaderCompression ? (discardCount + headerCount) : 0;

        const removeKeywords = elements.removeKeywordRowsCheckbox ? elements.removeKeywordRowsCheckbox.checked : false;
        const keywords = removeKeywords && elements.removeKeywordRowsInput ? elements.removeKeywordRowsInput.value.split(',').map(k => k.trim().toLowerCase()).filter(Boolean) : [];

        const colProps = sheet['!cols'] || [];
        const rowProps = sheet['!rows'] || [];

        // 1. 找出真正的資料寬度，忽略格式造成的假空白欄位
        let actualMaxCol = -1;
        for (const row of jsonData) {
            if (!row) continue;
            for (let j = row.length - 1; j > actualMaxCol; j--) {
                if (row[j] !== undefined && row[j] !== null && String(row[j]).trim() !== '') { actualMaxCol = j; break; }
            }
        }
        if (actualMaxCol === -1) return []; // 整張表都是空的

        const visibleCols = [];
        for (let c = 0; c <= actualMaxCol; c++) {
            if (!(colProps[startCol + c] && colProps[startCol + c].hidden)) visibleCols.push(c);
        }

        const result = [];
        const usedNames = new Set();

        // 2. 處理壓縮表頭
        if (useHeaderCompression && jsonData.length > discardCount) {
            const scanEnd = Math.min(discardCount + headerCount, jsonData.length);
            result.push(visibleCols.map((colIdx, index) => {
                const headerParts = [];
                for (let r = discardCount; r < scanEnd; r++) {
                    if (jsonData[r] && jsonData[r][colIdx] !== undefined) {
                        const val = String(jsonData[r][colIdx]).trim();
                        if (val !== '') headerParts.push(val.replace(/\r?\n|\r/g, ''));
                    }
                }
                const uniqueParts = [...new Set(headerParts)];
                const colLabel = `(欄位 ${index + 1})`;
                const baseName = uniqueParts.length > 0 ? `${colLabel}\n${uniqueParts.join('\n')}` : colLabel;

                let uniqueName = baseName;
                for (let counter = 2; usedNames.has(uniqueName); counter++) uniqueName = `${baseName}_${counter}`;
                usedNames.add(uniqueName);
                return uniqueName;
            }));
        }

        // 3. 處理資料列
        for (let idx = totalSkipCount; idx < jsonData.length; idx++) {
            const row = jsonData[idx];
            if (!row) continue;
            if (rowProps[startRow + idx] && rowProps[startRow + idx].hidden) continue;

            const newRow = visibleCols.map(c => row[c] !== undefined ? String(row[c]).trim() : '');
            if (newRow.every(val => val === '')) continue; // 全空列一律略過
            if (keywords.length > 0) {
                const rowText = newRow.join(' ').toLowerCase();
                if (keywords.some(k => rowText.includes(k))) continue;
            }
            result.push(newRow);
        }

        return result;
    }

    function setDropMessage(html) {
        if (elements.dropAreaContent) elements.dropAreaContent.innerHTML = html;
    }

    async function processFiles(fileList) {
        if (state.isProcessing) return;
        const files = Array.from(fileList).filter(f => VALID_FILE_EXTENSIONS.some(ext => f.name.toLowerCase().endsWith(ext)));
        if (!files.length) { alert('請上傳 Excel 檔案 (.xls, .xlsx, .ods)'); return; }

        const importMode = document.querySelector('input[name="import-mode"]:checked')?.value || 'first';
        const sheetCriteria = { name: elements.specificSheetNameInput?.value.trim() || '', position: elements.specificSheetPositionInput?.value.trim() || '' };

        state.isProcessing = true;
        state.rawSheets = [];

        try {
            for (let i = 0; i < files.length; i++) {
                const file = files[i];
                setDropMessage(`<h3>⏳ 讀取中 (${i + 1}/${files.length})</h3><p>${escHtml(file.name)}</p>`);
                await new Promise(r => setTimeout(r, 0)); // 讓 UI 有機會更新

                const workbook = XLSX.read(await file.arrayBuffer(), { type: 'array', cellStyles: true });
                for (const sheetName of getSelectedSheetNames(workbook, importMode, sheetCriteria)) {
                    const sheet = workbook.Sheets[sheetName];
                    if (!sheet['!ref']) continue;
                    const range = XLSX.utils.decode_range(sheet['!ref']);
                    range.e.r = Math.min(range.e.r, range.s.r + MAX_ROWS_PER_SHEET);

                    const jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1, range, raw: false });
                    fillMergedCells(jsonData, sheet['!merges'], range);

                    state.rawSheets.push({
                        filename: file.name, label: `${file.name} (${sheetName})`,
                        startRow: range.s.r, startCol: range.s.c, sheet, jsonData,
                    });
                }
            }
            rebuildFileInfos();
        } catch (err) {
            console.error('處理檔案發生錯誤:', err);
            setDropMessage(`<h3 style="color:#ef4444;">❌ 處理檔案錯誤</h3><p>${escHtml(err.message)}</p>`);
        } finally {
            state.isProcessing = false;
            if (elements.fileInput) elements.fileInput.value = '';
        }
    }

    // 合併儲存格：把左上角的值填滿整個合併範圍
    function fillMergedCells(jsonData, merges, range) {
        if (!merges) return;
        const maxCol = jsonData.reduce((m, r) => Math.max(m, r ? r.length : 0), 0);
        merges.forEach(merge => {
            const startR = merge.s.r - range.s.r, startC = merge.s.c - range.s.c;
            if (startR < 0 || startC < 0 || !jsonData[startR] || jsonData[startR][startC] === undefined) return;
            const val = jsonData[startR][startC];
            if (String(val).trim() === '') return;
            const maxR = Math.min(merge.e.r - range.s.r, jsonData.length - 1);
            const maxC = Math.min(merge.e.c - range.s.c, maxCol);
            for (let r = startR; r <= maxR; r++) {
                if (!jsonData[r]) jsonData[r] = [];
                for (let c = startC; c <= maxC; c++) jsonData[r][c] = val;
            }
        });
    }

    // 由原始工作表快取重建 fileInfos（匯入、重新套用清洗設定都走這裡）
    function rebuildFileInfos() {
        state.fileInfos = [];
        state.rawSheets.forEach(raw => {
            const filtered = applyPreprocessing(raw.jsonData, raw.sheet, raw.startRow, raw.startCol);
            if (!filtered.length) return;
            state.fileInfos.push({
                filename: raw.label,
                detectedFunds: detectFundsInText(raw.filename),
                headers: filtered[0].map((h, idx) => String(h).trim() || `(欄位 ${idx + 1})`),
                rows: filtered.slice(1),
            });
        });
        state.aggregatedRows = [];
        state.undoStack = []; // 彙整已重建，舊的復原點失效
        undoManager.hideToast();
        renderAggregationPanel();
        buildMappingTable();
        updateDropAreaDisplay();
        markSettingsClean();
    }

    // 比對模式或自訂清單變動後，重新辨識每個檔案所屬基金
    function redetectFunds() {
        if (!state.fileInfos.length) return;
        const rawByLabel = new Map(state.rawSheets.map(r => [r.label, r.filename]));
        state.fileInfos.forEach(info => { info.detectedFunds = detectFundsInText(rawByLabel.get(info.filename)); });
        state.aggregatedRows = [];
        state.undoStack = []; // 彙整已重建，舊的復原點失效
        undoManager.hideToast();
        renderAggregationPanel();
        buildMappingTable();
    }

    function markSettingsDirty() {
        if (state.rawSheets.length && elements.reapplyBanner) elements.reapplyBanner.classList.remove('hidden');
    }

    function markSettingsClean() {
        if (elements.reapplyBanner) elements.reapplyBanner.classList.add('hidden');
    }

    function getSelectedSheetNames(workbook, mode, criteria) {
        const names = workbook.SheetNames;
        if (!names.length) return [];
        if (mode === 'first') return [names[0]];
        if (mode === 'specific') return names.filter(n => n.toLowerCase().includes(criteria.name.toLowerCase()));
        if (mode === 'position') return utils.parsePositionString(criteria.position).map(i => names[i]).filter(Boolean);
        return names;
    }

    function updateDropAreaDisplay() {
        const hasFiles = state.fileInfos.length > 0;
        if (elements.dropArea) elements.dropArea.classList.toggle('compact', hasFiles);
        setDropMessage(hasFiles
            ? `<h3>📊 已載入 ${state.fileInfos.length} 張表</h3><p>拖放更多檔案或點擊上傳（會取代目前資料）</p>`
            : dropAreaInitialHtml);
    }

    function clearAllFiles() {
        if (!confirm('確定清除所有檔案？')) return;
        state.rawSheets = [];
        state.fileInfos = [];
        state.aggregatedRows = [];
        if (elements.fileInput) elements.fileInput.value = '';
        renderAggregationPanel();
        buildMappingTable();
        updateDropAreaDisplay();
        markSettingsClean();
    }

    function setMapChecks(fn) {
        if (elements.mappingTbody) elements.mappingTbody.querySelectorAll('.map-row-check').forEach(cb => { cb.checked = fn(cb.checked); });
    }

    function bindEvents() {
        setupDragAndDrop();

        if (elements.importSettingsPanel) {
            elements.importSettingsPanel.addEventListener('change', e => {
                if (e.target.name !== 'import-mode') return;
                const mode = e.target.value;
                if (elements.specificSheetNameGroup) elements.specificSheetNameGroup.classList.toggle('hidden', mode !== 'specific');
                if (elements.specificSheetPositionGroup) elements.specificSheetPositionGroup.classList.toggle('hidden', mode !== 'position');
            });
        }

        ['skipTopRowsCheckbox', 'discardRowsInput', 'headerRowsInput', 'removeEmptyRowsCheckbox', 'removeKeywordRowsCheckbox', 'removeKeywordRowsInput'].forEach(id => {
            if (!elements[id]) return;
            elements[id].addEventListener('change', markSettingsDirty);
            elements[id].addEventListener('input', utils.debounce(markSettingsDirty, 500));
        });

        if (elements.skipTopRowsCheckbox) {
            elements.skipTopRowsCheckbox.addEventListener('change', e => {
                if (elements.discardRowsInput) elements.discardRowsInput.disabled = !e.target.checked;
                if (elements.headerRowsInput) elements.headerRowsInput.disabled = !e.target.checked;
            });
        }
        if (elements.removeKeywordRowsCheckbox) elements.removeKeywordRowsCheckbox.addEventListener('change', e => { if (elements.removeKeywordRowsInput) elements.removeKeywordRowsInput.disabled = !e.target.checked; });
        if (elements.reapplySettingsBtn) elements.reapplySettingsBtn.addEventListener('click', rebuildFileInfos);

        if (elements.backToTopBtn) elements.backToTopBtn.addEventListener('click', () => window.scrollTo({ top: 0, behavior: 'smooth' }));
        window.addEventListener('scroll', () => {
            if (elements.backToTopBtn) elements.backToTopBtn.classList.toggle('visible', window.scrollY > window.innerHeight / 2);
        });

        document.addEventListener('keydown', e => {
            if ((e.ctrlKey || e.metaKey) && e.key === 'z' && !e.target.matches('input, textarea')) { e.preventDefault(); undoManager.undoLast(); }
        });
        if (elements.undoBtn) elements.undoBtn.addEventListener('click', () => undoManager.undoLast());

        // --- Step 2: 基金對照 & 彙整 事件 ---
        if (elements.executeMatchBtn) elements.executeMatchBtn.addEventListener('click', handleAutoMatch);
        if (elements.exportMatchedBtn) elements.exportMatchedBtn.addEventListener('click', handleExportMatched);

        if (elements.mapCheckAll) elements.mapCheckAll.addEventListener('change', e => setMapChecks(() => e.target.checked));
        if (elements.mapSelectAllBtn) elements.mapSelectAllBtn.addEventListener('click', () => setMapChecks(() => true));
        if (elements.mapDeselectAllBtn) elements.mapDeselectAllBtn.addEventListener('click', () => setMapChecks(() => false));
        if (elements.mapInvertBtn) elements.mapInvertBtn.addEventListener('click', () => setMapChecks(c => !c));

        // 彙整面板事件
        if (elements.aggCheckReviewBtn) elements.aggCheckReviewBtn.addEventListener('click', () => {
            state.aggregatedRows.forEach(r => { if (r.needsReview) r.checked = true; });
            renderAggregationPanel();
        });
        if (elements.aggUncheckAllBtn) elements.aggUncheckAllBtn.addEventListener('click', () => {
            state.aggregatedRows.forEach(r => { r.checked = false; });
            renderAggregationPanel();
        });
        if (elements.aggDeleteCheckedBtn) elements.aggDeleteCheckedBtn.addEventListener('click', deleteCheckedAggregatedRows);
        if (elements.aggExportBtn) elements.aggExportBtn.addEventListener('click', exportAggregatedRows);

        if (elements.matchColSelect) elements.matchColSelect.addEventListener('change', () => {
            state.matchColumn = elements.matchColSelect.value;
            if (state.aggregatedRows.length) buildAggregation({ silent: true });
        });

        if (elements.aggregatePanel) {
            elements.aggregatePanel.addEventListener('change', e => {
                if (e.target.classList.contains('agg-row-check')) {
                    const tr = e.target.closest('tr');
                    state.aggregatedRows[parseInt(tr.dataset.aggIndex, 10)].checked = e.target.checked;
                    tr.classList.toggle('row-checked-delete', e.target.checked);
                    updateAggregationCounts();
                }
                if (e.target.id === 'agg-check-all-cb') {
                    state.aggregatedRows.forEach(r => { r.checked = e.target.checked; });
                    renderAggregationPanel();
                }
            });
        }

        if (elements.toggleImportSettings) {
            elements.toggleImportSettings.addEventListener('click', () => {
                if (elements.importSettingsPanel) elements.importSettingsPanel.classList.toggle('collapsed');
            });
        }

        if (elements.clearAllBtn) elements.clearAllBtn.addEventListener('click', clearAllFiles);

        // ===== 自訂比對 =====
        if (elements.tabStandard) elements.tabStandard.addEventListener('click', () => switchCompareMode('standard'));
        if (elements.tabCustom) elements.tabCustom.addEventListener('click', () => switchCompareMode('custom'));
        if (elements.applyCustomListBtn) elements.applyCustomListBtn.addEventListener('click', applyCustomList);
        if (elements.clearCustomListBtn) {
            elements.clearCustomListBtn.addEventListener('click', () => {
                if (elements.customListTextarea) elements.customListTextarea.value = '';
                applyCustomList();
            });
        }
    }

    function handleAutoMatch() {
        buildAggregation();
        if (elements.aggregatePanel && !elements.aggregatePanel.classList.contains('hidden')) {
            elements.aggregatePanel.scrollIntoView({ behavior: 'smooth', block: 'start' });
        }
    }

    function handleExportMatched() {
        if (!state.aggregatedRows.length) buildAggregation();
        if (state.aggregatedRows.length) exportAggregatedRows();
        else if (elements.mappingTbody?.querySelector('.map-row-check:checked')) alert('沒有找到可匯出的資料。');
    }

    async function init() {
        cacheElements();
        dropAreaInitialHtml = elements.dropAreaContent ? elements.dropAreaContent.innerHTML : '';
        bindEvents();
        await loadFundConfig();
    }

    return { init };
})();

if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', ExcelViewer.init);
} else {
    ExcelViewer.init();
}
