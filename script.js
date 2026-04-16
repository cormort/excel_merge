/**
 * ExcelViewer — 不卡死極速版 (實體邊界掃描、非同步 UI 釋放、雙軌全功能)
 */

const ExcelViewer = (() => {
    'use strict';

    const CONSTANTS = { VALID_FILE_EXTENSIONS: ['.xls', '.xlsx', '.ods'] };

    const state = {
        originalHtmlString: '',
        isProcessing: false,
        loadedFiles: [],
        loadedTables: 0,
        zoomedCard: null,

        rawSheetsCache: [],
        isSettingsDirty: false,
        undoStack: [],

        columnModalContext: 'main', 
        dedupContext: 'main',
        mainViewHiddenColumns: new Set(),
        isMainEditing: false,

        isMergedView: false,
        isEditing: false,
        showTotalRow: false,
        showSourceColumn: false,
        mergedData: [],
        mergedHeaders: [],

        fundSortOrder: [],
        fundAliasMap: {},
        fundAliasKeys: [],

        // 基金對照 & 彙整
        fileInfos: [],          // 每個檔案的辨識結果 [{filename, detectedFunds, headers, rows}]
        aggregatedRows: [],     // 彙整後的逐列資料
        matchColumn: '1',       // 比對基金名稱的欄位（1=第一欄, 2=第二欄）
    };

    const elements = {};

    const utils = {
        debounce(fn, ms) {
            let timer;
            return (...args) => { clearTimeout(timer); timer = setTimeout(() => fn(...args), ms); };
        },
        formatNumber(value) {
            const num = parseFloat(String(value).replace(/,/g, ''));
            return isNaN(num) ? value : new Intl.NumberFormat('en-US').format(num);
        },
        isStrictNumber(value) {
            const clean = String(value).replace(/,/g, '').trim();
            return clean !== '' && !isNaN(clean);
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
        readFileAsBinary(file) {
            return new Promise((resolve, reject) => {
                const reader = new FileReader();
                reader.onload = e => resolve(e.target.result);
                reader.onerror = reject;
                reader.readAsBinaryString(file);
            });
        },
        calculateColumnWidths(data) {
            if (!data.length) return [];
            return data[0].map((_, col) => ({
                wch: Math.min(50, Math.max(10, ...data.map(r => (r[col] ? String(r[col]).length : 0))) + 2),
            }));
        },
        buildKeywordMatcher(keyword, isRegex) {
            if (!keyword) return null;
            if (isRegex) return text => new RegExp(keyword, 'i').test(text);
            if (keyword.includes(',')) {
                const kws = keyword.split(',').map(k => k.trim().toLowerCase()).filter(Boolean);
                return text => kws.some(k => text.includes(k));
            }
            const kws = keyword.split(/\s+/).map(k => k.trim().toLowerCase()).filter(Boolean);
            return text => kws.every(k => text.includes(k));
        }
    };

    const undoManager = {
        push(description, restoreFn) {
            state.undoStack.push({ description, restoreFn });
            if (state.undoStack.length > 5) state.undoStack.shift();
            this.showToast(description);
        },
        undoLast() {
            if (state.undoStack.length === 0) return;
            const action = state.undoStack.pop();
            action.restoreFn();
            this.hideToast();
        },
        showToast(desc) {
            if (elements.undoToast && elements.undoText) {
                elements.undoText.textContent = `已${desc}`;
                elements.undoToast.classList.add('show');
                clearTimeout(this.timer);
                this.timer = setTimeout(() => this.hideToast(), 8000);
            }
        },
        hideToast() {
            if (elements.undoToast) elements.undoToast.classList.remove('show');
        }
    };

    function cacheElements() {
        const mapping = {
            fileInput: 'file-input', displayArea: 'excel-display-area', searchInput: 'search-input',
            dropArea: 'drop-area', deleteSelectedBtn: 'delete-selected-btn', invertSelectionBtn: 'invert-selection-btn',
            resetViewBtn: 'reset-view-btn', selectEmptyBtn: 'select-empty-btn', showHiddenBtn: 'show-hidden-btn', 
            exportMergedXlsxBtn: 'export-merged-xlsx-btn', selectByKeywordGroup: 'select-by-keyword-group', 
            selectKeywordInput: 'select-keyword-input', selectByKeywordBtn: 'select-by-keyword-btn', 
            selectKeywordRegex: 'select-keyword-regex', loadStatusMessage: 'load-status-message', controlPanel: 'control-panel',
            dropAreaInitial: 'drop-area-initial', dropAreaLoaded: 'drop-area-loaded', fileCount: 'file-count', 
            fileNames: 'file-names', clearFilesBtn: 'clear-files-btn', selectAllBtn: 'select-all-btn', 
            importOptionsContainer: 'import-options-container', specificSheetNameGroup: 'specific-sheet-name-group', 
            specificSheetNameInput: 'specific-sheet-name-input', specificSheetPositionGroup: 'specific-sheet-position-group', 
            specificSheetPositionInput: 'specific-sheet-position-input', selectAllTablesBtn: 'select-all-tables-btn', 
            unselectAllTablesBtn: 'unselect-all-tables-btn', deleteSelectedTablesBtn: 'delete-selected-tables-btn', 
            sortByNameBtn: 'sort-by-fund-name-btn', tableLevelControls: 'table-level-controls', 
            selectedTablesInfo: 'selected-tables-info', selectedTablesList: 'selected-tables-list', 
            listViewBtn: 'list-view-btn', gridViewBtn: 'grid-view-btn', backToTopBtn: 'back-to-top-btn',
            gridScaleControl: 'grid-scale-control', gridScaleSlider: 'grid-scale-slider',
            
            mainColumnOperationsBtn: 'main-column-operations-btn',
            mainSmartDedupBtn: 'main-smart-dedup-btn', mainEditDataBtn: 'main-edit-data-btn',
            mainSaveEditsBtn: 'main-save-edits-btn', mainCopyRowsBtn: 'main-copy-rows-btn',
            
            toggleMainConditionBtn: 'toggle-main-condition-btn',
            mainCondition2Wrapper: 'main-condition-2-wrapper',
            mainColSelect1: 'main-col-select-1', mainColSelect2: 'main-col-select-2',
            mainInputCriteria1: 'main-input-criteria-1', mainInputCriteria2: 'main-input-criteria-2',
            mainExecuteFilterBtn: 'main-execute-filter-btn',

            skipTopRowsCheckbox: 'skip-top-rows-checkbox', discardRowsInput: 'discard-rows-input', 
            headerRowsInput: 'header-rows-input', removeEmptyRowsCheckbox: 'remove-empty-rows-checkbox', 
            removeKeywordRowsCheckbox: 'remove-keyword-rows-checkbox', removeKeywordRowsInput: 'remove-keyword-rows-input', 
            reapplyBanner: 'reapply-banner', reapplySettingsBtn: 'reapply-settings-btn',
            
            mergeViewModal: 'merge-view-modal', closeMergeViewBtn: 'close-merge-view-btn', mergeViewContent: 'merge-view-content',
            mergeViewBtn: 'merge-view-btn', viewCheckedCombinedBtn: 'view-checked-combined-btn',
            toggleToolbarBtn: 'toggle-toolbar-btn', collapsibleToolbar: 'collapsible-toolbar-area',
            searchInputMerged: 'search-input-merged', selectKeywordInputMerged: 'select-keyword-input-merged',
            selectKeywordRegexMerged: 'select-keyword-regex-merged', executeFilterSelectionBtn: 'execute-filter-selection-btn',
            unselectMergedRowsBtn: 'unselect-merged-rows-btn', invertSelectionMergedBtn: 'invert-selection-merged-btn',
            
            toggleMergeConditionBtn: 'toggle-merge-condition-btn',
            mergeCondition2Wrapper: 'merge-condition-2-wrapper',
            colSelect1: 'col-select-1', colSelect2: 'col-select-2', inputCriteria1: 'input-criteria-1', inputCriteria2: 'input-criteria-2',
            
            editDataBtn: 'edit-data-btn', saveEditsBtn: 'save-edits-btn', cancelEditsBtn: 'cancel-edits-btn', addNewRowBtn: 'add-new-row-btn',
            copySelectedRowsBtn: 'copy-selected-rows-btn', deleteMergedRowsBtn: 'delete-merged-rows-btn', toggleTotalRowBtn: 'toggle-total-row-btn',
            toggleSourceColBtn: 'toggle-source-col-btn', exportCurrentMergedXlsxBtn: 'export-current-merged-xlsx-btn',
            sortMergedByNameBtn: 'sort-merged-by-fund-name-btn', columnOperationsBtn: 'column-operations-btn', columnModal: 'column-modal',
            closeColumnModalBtn: 'close-column-modal-btn', columnChecklist: 'column-checklist', applyColumnChangesBtn: 'apply-column-changes-btn',
            modalCheckAll: 'modal-check-all', modalUncheckAll: 'modal-uncheck-all', smartDedupBtn: 'smart-dedup-btn',
            dedupModal: 'dedup-modal', closeDedupModalBtn: 'close-dedup-modal-btn', cancelDedupBtn: 'cancel-dedup-btn', executeDedupBtn: 'execute-dedup-btn',
            dedupColSelect: 'dedup-col-select', dedupResultPanel: 'dedup-result-panel', dedupResultText: 'dedup-result-text', 
            clearDedupMarksBtn: 'clear-dedup-marks-btn', deleteDedupMarksBtn: 'delete-dedup-marks-btn',
            undoToast: 'undo-toast', undoText: 'undo-text', undoBtn: 'undo-btn',

            // Step 2: 基金對照 & 彙整
            stepMapping: 'step-mapping', stepData: 'step-data', stepExport: 'step-export',
            mappingTbody: 'mapping-tbody', mapCheckAll: 'map-check-all',
            mapSelectAllBtn: 'map-select-all-btn', mapDeselectAllBtn: 'map-deselect-all-btn', mapInvertBtn: 'map-invert-btn',
            executeMatchBtn: 'execute-match-btn', exportMatchedBtn: 'export-matched-btn',
            aggregatePanel: 'aggregate-panel', aggregateThead: 'aggregate-thead', aggregateTbody: 'aggregate-tbody',
            aggStatTotal: 'agg-stat-total', aggStatReview: 'agg-stat-review', aggStatChecked: 'agg-stat-checked',
            aggCheckReviewBtn: 'agg-check-review-btn', aggUncheckAllBtn: 'agg-uncheck-all-btn',
            aggDeleteCheckedBtn: 'agg-delete-checked-btn', aggExportBtn: 'agg-export-btn',
            matchColSelect: 'match-col-select',
            toggleImportSettings: 'toggle-import-settings', importSettingsPanel: 'import-settings-panel',
            topBarActions: 'top-bar-actions', statTableCount: 'stat-table-count', statRowCount: 'stat-row-count',
            clearAllBtn: 'clear-all-btn'
        };

        Object.keys(mapping).forEach(key => { elements[key] = document.getElementById(mapping[key]); });
    }

    function getActiveScope() { return state.isMergedView ? elements.mergeViewContent : elements.displayArea; }

    function resetControls(isNewFile) {
        if (!isNewFile) return;
        state.originalHtmlString = '';
        if (elements.searchInput) elements.searchInput.value = '';
        if (elements.selectKeywordInput) elements.selectKeywordInput.value = '';
        if (elements.selectKeywordRegex) elements.selectKeywordRegex.checked = false;
        if (elements.controlPanel) elements.controlPanel.classList.add('hidden');
        updateSelectionInfo();
    }

    async function loadFundConfig() {
        try {
            const response = await fetch(`fund-config.json?v=${Date.now()}`);
            if (!response.ok) return;
            const config = await response.json();
            if (config.sortOrder && config.aliasMap) {
                state.fundSortOrder = config.sortOrder;
                state.fundAliasMap = config.aliasMap;
                state.fundAliasKeys = Object.keys(config.aliasMap).sort((a, b) => b.length - a.length);
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

        state.fundAliasKeys.forEach(alias => add(alias, state.fundAliasMap[alias]));
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

    // 判斷儲存格文字是否屬於某標準基金（別名匹配 或 前4字相同 + 字數相同）
    function cellMatchesFund(cellText, standardFund) {
        if (!cellText || !standardFund) return false;
        const cellTextTrimmed = cellText.trim();
        const fundTextTrimmed = standardFund.trim();
        if (cellTextTrimmed === fundTextTrimmed) return true;
        
        // 檢查別名是否匹配
        const aliasKeys = Object.keys(state.fundAliasMap);
        for (const alias of aliasKeys) {
            if (cellTextTrimmed === alias && state.fundAliasMap[alias] === standardFund) {
                return true;
            }
        }

        const cellChars = Array.from(cellTextTrimmed);
        const fundChars = Array.from(fundTextTrimmed);
        if (cellChars.length < 4 || fundChars.length < 4) return false;
        return cellChars.length === fundChars.length && 
               cellChars[0] === fundChars[0] && 
               cellChars[1] === fundChars[1] &&
               cellChars[2] === fundChars[2] &&
               cellChars[3] === fundChars[3];
    }

    // --- Stage 1: 基金對照總覽 (mapping table) ---
    // 從 state.fileInfos 渲染對照表：檔名 | 對應標準基金（可能多個）| 資料列數
    function buildMappingTable() {
        if (!elements.mappingTbody || !state.fileInfos.length) return;

        const html = state.fileInfos.map((info, idx) => {
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
        elements.mappingTbody.innerHTML = html;

        // 顯示 Step 2
        if (elements.stepMapping) elements.stepMapping.classList.remove('hidden');

        // 更新頂部統計
        if (elements.topBarActions) elements.topBarActions.classList.remove('hidden');
        if (elements.statTableCount) elements.statTableCount.textContent = state.fileInfos.length;
        if (elements.statRowCount) elements.statRowCount.textContent = state.fileInfos.reduce((s, f) => s + f.rows.length, 0);
    }

    function escHtml(str) {
        return String(str ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
    }

    // --- Stage 2: 彙整面板 (Aggregation) ---
    // 依標準基金順序，把勾選的檔案的所有資料列貼入（一個檔案對應多個基金就貼多次）
    function buildAggregation() {
        const selectedIdxs = [];
        if (elements.mappingTbody) {
            elements.mappingTbody.querySelectorAll('.map-row-check:checked').forEach(cb => {
                selectedIdxs.push(parseInt(cb.dataset.index, 10));
            });
        }
        if (selectedIdxs.length === 0) { alert('請先勾選要彙整的檔案。'); return; }

        const aggregated = [];
        let reviewCount = 0;

        // 依標準基金順序巡覽
        for (const standardFund of state.fundSortOrder) {
            // 對每個勾選的檔案，檢查是否對應到此基金
            for (const idx of selectedIdxs) {
                const info = state.fileInfos[idx];
                if (!info || !info.detectedFunds.includes(standardFund)) continue;

                // 把該檔案的全部資料列貼入此基金底下
                info.rows.forEach((cells, rowIdx) => {
                    const firstCellText = (cells[0] || '').trim();
                    const secondCellText = (cells[1] || '').trim();
                    const thirdCellText = (cells[2] || '').trim();
                    const lastCellText = cells.length > 1 ? (cells[cells.length - 1] || '').trim() : '';
                    const matchColIdx = parseInt(state.matchColumn, 10) - 1;
                    const matchCellText = (cells[matchColIdx] || '').trim();
                    const needsReview = !cellMatchesFund(matchCellText, standardFund);
                    if (needsReview) reviewCount++;
                    aggregated.push({
                        standardFund,
                        sourceFile: info.filename,
                        sourceRowIdx: rowIdx,
                        headers: info.headers,
                        cells,
                        firstCellText,
                        secondCellText,
                        thirdCellText,
                        lastCellText,
                        matchCellText,
                        matchColumn: state.matchColumn,
                        needsReview,
                        checked: needsReview,
                    });
                });
            }
        }

        state.aggregatedRows = aggregated;
        renderAggregationPanel();
    }

    // 渲染彙整面板：只顯示「標準基金 | 來源檔名 | 第一欄」+ 勾選框
    function renderAggregationPanel() {
        if (!elements.aggregatePanel || !elements.aggregateTbody || !elements.aggregateThead) return;
        if (!state.aggregatedRows.length) {
            elements.aggregatePanel.classList.add('hidden');
            return;
        }
        elements.aggregatePanel.classList.remove('hidden');

        const matchCol = state.matchColumn || '1';
        const matchColLabel = matchCol === '1' ? '第一欄（基金名稱）' : matchCol === '2' ? '第二欄（基金名稱）' : '第三欄（基金名稱）';
        const getOtherCol = (mc) => {
            if (mc === '1') return [row.secondCellText, row.thirdCellText];
            if (mc === '2') return [row.firstCellText, row.thirdCellText];
            return [row.firstCellText, row.secondCellText];
        };
        elements.aggregateThead.innerHTML = `<tr>
            <th style="width:40px;"><input type="checkbox" id="agg-check-all-cb"></th>
            <th>#</th>
            <th>標準基金</th>
            <th>來源檔案</th>
            <th>${matchColLabel}</th>
            <th>其他欄位</th>
            <th>第三欄</th>
        </tr>`;

        const html = state.aggregatedRows.map((row, idx) => {
            const cls = row.needsReview ? 'row-needs-review' : '';
            const checkedCls = row.checked ? 'row-checked-delete' : '';
            const matchText = row.matchCellText || '';
            const [other1, other2] = getOtherCol(matchCol);
            return `<tr data-agg-index="${idx}" class="${cls} ${checkedCls}">
                <td><input type="checkbox" class="agg-row-check" ${row.checked ? 'checked' : ''}></td>
                <td>${idx + 1}</td>
                <td>${escHtml(row.standardFund)}</td>
                <td title="${escHtml(row.sourceFile)}" style="max-width:300px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;">${escHtml(row.sourceFile)}</td>
                <td class="${row.needsReview ? 'first-col-mismatch' : ''}">${escHtml(matchText)}</td>
                <td>${escHtml(other1 || '')}</td>
                <td>${escHtml(row.thirdCellText || '')}</td>
            </tr>`;
        }).join('');
        elements.aggregateTbody.innerHTML = html;
        updateAggregationCounts();
    }

    function rebuildAggregationWithNewMatchColumn() {
        const selectedIdxs = [];
        if (elements.mappingTbody) {
            elements.mappingTbody.querySelectorAll('.map-row-check:checked').forEach(cb => {
                selectedIdxs.push(parseInt(cb.dataset.index, 10));
            });
        }
        if (selectedIdxs.length === 0) return;

        const aggregated = [];
        let reviewCount = 0;

        for (const standardFund of state.fundSortOrder) {
            for (const idx of selectedIdxs) {
                const info = state.fileInfos[idx];
                if (!info || !info.detectedFunds.includes(standardFund)) continue;

                info.rows.forEach((cells, rowIdx) => {
                    const firstCellText = (cells[0] || '').trim();
                    const secondCellText = (cells[1] || '').trim();
                    const thirdCellText = (cells[2] || '').trim();
                    const lastCellText = cells.length > 1 ? (cells[cells.length - 1] || '').trim() : '';
                    const matchColIdx = parseInt(state.matchColumn, 10) - 1;
                    const matchCellText = (cells[matchColIdx] || '').trim();
                    const needsReview = !cellMatchesFund(matchCellText, standardFund);
                    if (needsReview) reviewCount++;
                    aggregated.push({
                        standardFund,
                        sourceFile: info.filename,
                        sourceRowIdx: rowIdx,
                        headers: info.headers,
                        cells,
                        firstCellText,
                        secondCellText,
                        thirdCellText,
                        lastCellText,
                        matchCellText,
                        matchColumn: state.matchColumn,
                        needsReview,
                        checked: needsReview,
                    });
                });
            }
        }

        state.aggregatedRows = aggregated;
        renderAggregationPanel();
    }

    function updateAggregationCounts() {
        const total = state.aggregatedRows.length;
        const review = state.aggregatedRows.filter(r => r.needsReview).length;
        const checked = state.aggregatedRows.filter(r => r.checked).length;
        if (elements.aggStatTotal) elements.aggStatTotal.textContent = total;
        if (elements.aggStatReview) elements.aggStatReview.textContent = review;
        if (elements.aggStatChecked) elements.aggStatChecked.textContent = checked;
    }

    function deleteCheckedAggregatedRows() {
        const before = state.aggregatedRows.length;
        state.aggregatedRows = state.aggregatedRows.filter(r => !r.checked);
        const removed = before - state.aggregatedRows.length;
        if (removed === 0) { alert('沒有勾選任何列。'); return; }
        renderAggregationPanel();
        undoManager.showToast(`刪除 ${removed} 筆彙整資料`);
    }

    // --- Checklist 匯出 ---
    function exportAggregatedRows() {
        const validRows = state.aggregatedRows.filter(r => !r.checked);

        // 收集所有表頭聯集
        const headerSet = new Set();
        validRows.forEach(r => r.headers.forEach(h => headerSet.add(h)));
        const unionHeaders = Array.from(headerSet);

        // 建立匯出資料（加入繳交狀態欄位）
        const exportData = [['標準基金', '繳交狀態', '來源檔案', ...unionHeaders]];

        // 將有效資料依基金分組
        const dataByFund = {};
        validRows.forEach(row => {
            if (!dataByFund[row.standardFund]) dataByFund[row.standardFund] = [];
            dataByFund[row.standardFund].push(row);
        });

        // ⭐ Checklist 核心：巡覽 fund-config.json 裡全部標準基金
        state.fundSortOrder.forEach(standardFund => {
            const fundRows = dataByFund[standardFund];
            if (fundRows && fundRows.length > 0) {
                fundRows.forEach(row => {
                    const cellMap = {};
                    row.headers.forEach((h, i) => { cellMap[h] = row.cells[i] ?? ''; });
                    exportData.push([
                        standardFund,
                        '✅ 已回傳',
                        row.sourceFile,
                        ...unionHeaders.map(h => cellMap[h] ?? '')
                    ]);
                });
            } else {
                exportData.push([
                    standardFund,
                    '❌ 未回傳/無資料',
                    '-',
                    ...unionHeaders.map(() => '')
                ]);
            }
        });

        const ws = XLSX.utils.aoa_to_sheet(exportData);
        
        // 設定數值格式
        const cellRefs = XLSX.utils.decode_range(ws['!ref']);
        for (let C = 0; C <= cellRefs.e.c; C++) {
            for (let R = 1; R <= cellRefs.e.r; R++) {
                const cellAddr = XLSX.utils.encode_cell({ r: R, c: C });
                const cell = ws[cellAddr];
                if (!cell) continue;
                const val = cell.v;
                if (val !== '' && !isNaN(Number(val)) && String(val).match(/^[\d,.-]+$/)) {
                    cell.t = 'n';
                    cell.v = Number(String(val).replace(/,/g, ''));
                }
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
            if (e.target.id === 'clear-files-btn' || e.target.closest('.btn-clear') || e.target === elements.fileInput) return;
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

    // 🚀 【核心效能引擎】：強制邊界檢測，切斷所有幽靈資料
    function applyPreprocessing(jsonData, sheet, startRow, startCol, endCol) {
        const useHeaderCompression = elements.skipTopRowsCheckbox && elements.skipTopRowsCheckbox.checked;
        const discardCount = useHeaderCompression && elements.discardRowsInput ? parseInt(elements.discardRowsInput.value, 10) || 0 : 0;
        const headerCount = useHeaderCompression && elements.headerRowsInput ? parseInt(elements.headerRowsInput.value, 10) || 1 : 1;
        const totalSkipCount = useHeaderCompression ? (discardCount + headerCount) : 0;

        const removeEmpty = elements.removeEmptyRowsCheckbox ? elements.removeEmptyRowsCheckbox.checked : false;
        const removeKeywords = elements.removeKeywordRowsCheckbox ? elements.removeKeywordRowsCheckbox.checked : false;
        const keywords = removeKeywords && elements.removeKeywordRowsInput ? elements.removeKeywordRowsInput.value.split(',').map(k => k.trim().toLowerCase()).filter(Boolean) : [];

        const colProps = sheet['!cols'] || [];
        const rowProps = sheet['!rows'] || [];

        // 1. 找出真正的資料寬度，忽略格式造成的假空白欄位
        let actualMaxCol = -1;
        for (let i = 0; i < jsonData.length; i++) {
            const row = jsonData[i];
            if (row) {
                for (let j = row.length - 1; j >= 0; j--) {
                    if (row[j] !== undefined && row[j] !== null && String(row[j]).trim() !== '') {
                        if (j > actualMaxCol) actualMaxCol = j;
                        break;
                    }
                }
            }
        }
        if (actualMaxCol === -1) return []; // 整張表都是空的

        const visibleCols = [];
        for (let c = 0; c <= actualMaxCol; c++) {
            const absCol = startCol + c;
            if (!(colProps[absCol] && colProps[absCol].hidden)) {
                visibleCols.push(c);
            }
        }

        const result = [];
        const usedNames = new Set();

        // 2. 處理壓縮表頭
        if (useHeaderCompression && jsonData.length > discardCount) {
            const headerRow = visibleCols.map((colIdx, index) => {
                const headerParts = [];
                const scanEnd = Math.min(discardCount + headerCount, jsonData.length);
                for (let r = discardCount; r < scanEnd; r++) {
                    if (jsonData[r] && jsonData[r][colIdx] !== undefined) {
                        const val = String(jsonData[r][colIdx]).trim();
                        if (val !== '') headerParts.push(val.replace(/\r?\n|\r/g, ''));
                    }
                }
                const uniqueParts = [...new Set(headerParts)];
                const colLabel = `(欄位 ${index + 1})`;
                let baseName = uniqueParts.length > 0 ? `${colLabel}\n${uniqueParts.join('\n')}` : colLabel;
                
                let uniqueName = baseName;
                let counter = 2;
                while (usedNames.has(uniqueName)) {
                    uniqueName = `${baseName}_${counter}`;
                    counter++;
                }
                usedNames.add(uniqueName);
                return uniqueName;
            });
            result.push(headerRow);
        }

        // 3. 處理資料列
        for (let idx = totalSkipCount; idx < jsonData.length; idx++) {
            const row = jsonData[idx];
            if (!row) continue;
            
            const absRow = startRow + idx;
            if (rowProps[absRow] && rowProps[absRow].hidden) continue;

            const newRow = visibleCols.map(c => row[c] !== undefined ? String(row[c]).trim() : '');
            
            if (removeEmpty && newRow.every(val => val === '')) continue;
            
            if (keywords.length > 0 && newRow.some(val => val !== '')) {
                const rowText = newRow.join(' ').toLowerCase();
                if (keywords.some(k => rowText.includes(k))) continue;
            }

            // 只要不是全空就加入
            if (newRow.some(val => val !== '')) {
                result.push(newRow);
            }
        }

        return result;
    }

    async function processFiles(fileList) {
        if (!fileList || fileList.length === 0) return;
        const files = Array.from(fileList).filter(f => CONSTANTS.VALID_FILE_EXTENSIONS.some(ext => f.name.toLowerCase().endsWith(ext)));
        if (!files.length) { alert('請上傳 Excel 檔案 (.xls, .xlsx)'); return; }
        if (state.isProcessing) return;

        const importModeInput = document.querySelector('input[name="import-mode"]:checked');
        const importMode = importModeInput ? importModeInput.value : 'first';
        const sheetCriteria = { name: elements.specificSheetNameInput?.value.trim() || '', position: elements.specificSheetPositionInput?.value.trim() || '' };

        state.isProcessing = true;
        if(elements.displayArea) elements.displayArea.innerHTML = '<div class="loading">正在準備解析...</div>';
        state.fileInfos = []; state.aggregatedRows = [];
        state.rawSheetsCache = []; state.loadedFiles = [];

        try {
            for (let i = 0; i < files.length; i++) {
                const file = files[i];

                if(elements.displayArea) elements.displayArea.innerHTML = `<div class="loading">高速讀取中... 檔案 (${i + 1}/${files.length}) : ${file.name}</div>`;
                await new Promise(r => setTimeout(r, 20));

                const binary = await utils.readFileAsBinary(file);
                const workbook = XLSX.read(binary, { type: 'binary', cellStyles: true });
                const sheetNames = await getSelectedSheetNames(file.name, workbook, importMode, sheetCriteria);

                for (const sheetName of sheetNames) {
                    const sheet = workbook.Sheets[sheetName];
                    const ref = sheet['!ref'];
                    if (!ref) continue;
                    const range = XLSX.utils.decode_range(ref);
                    range.e.r = Math.min(range.e.r, range.s.r + 10000);

                    let jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1, range: range, raw: false });

                    let maxCol = 0;
                    for(let r=0; r<jsonData.length; r++) {
                        if(jsonData[r] && jsonData[r].length > maxCol) maxCol = jsonData[r].length;
                    }

                    if (sheet['!merges']) {
                        sheet['!merges'].forEach(merge => {
                            const startR = merge.s.r - range.s.r, startC = merge.s.c - range.s.c;
                            const endR = merge.e.r - range.s.r, endC = merge.e.c - range.s.c;
                            if (startR >= 0 && startC >= 0 && jsonData[startR] && jsonData[startR][startC] !== undefined) {
                                const val = jsonData[startR][startC];
                                if (String(val).trim() === '') return;
                                const maxR = Math.min(endR, jsonData.length - 1);
                                const maxC = Math.min(endC, maxCol);
                                for (let r = startR; r <= maxR; r++) {
                                    if (!jsonData[r]) jsonData[r] = [];
                                    for (let c = startC; c <= maxC; c++) { jsonData[r][c] = val; }
                                }
                            }
                        });
                    }

                    const label = `${file.name} (${sheetName})`;
                    state.rawSheetsCache.push({ label, startRow: range.s.r, startCol: range.s.c, endCol: range.e.c, sheet, jsonData: jsonData });

                    // 預處理：清洗資料（拋棄列、壓縮表頭、移除空白列等）
                    const filtered = applyPreprocessing(jsonData, sheet, range.s.r, range.s.c, range.e.c);
                    if (filtered.length > 0) {
                        // 直接存入 fileInfos（不渲染到 HTML）
                        const headers = filtered[0].map((h, idx) => String(h).trim() || `(欄位 ${idx + 1})`);
                        const rows = filtered.slice(1); // 資料列（不含表頭）
                        const detectedFunds = detectFundsInText(file.name);

                        state.fileInfos.push({ filename: label, detectedFunds, headers, rows });
                        state.loadedFiles.push(label);
                    }
                }
            }

            state.loadedTables = state.fileInfos.length;
            if(elements.displayArea) elements.displayArea.innerHTML = '';

            // 直接顯示 Stage 1 對照表
            buildMappingTable();
            updateDropAreaDisplay();
            markSettingsClean();
        } catch (err) {
            console.error("處理檔案發生錯誤:", err);
            if(elements.displayArea) elements.displayArea.innerHTML = `<p style="color:red; font-weight:bold;">❌ 處理檔案錯誤：<br>${err.message}</p>`;
        } finally {
            state.isProcessing = false;
            if (elements.fileInput) elements.fileInput.value = '';
        }
    }

    function reapplyPreprocessing() {
        if (!state.rawSheetsCache.length) return;
        state.isProcessing = true;
        if(elements.displayArea) elements.displayArea.innerHTML = '<div class="loading">重新清洗中...</div>';
        
        setTimeout(() => {
            const tablesToRender = [];
            state.loadedFiles = [];
            state.rawSheetsCache.forEach(cache => {
                const filtered = applyPreprocessing(cache.jsonData, cache.sheet, cache.startRow, cache.startCol, cache.endCol);
                if (filtered.length > 0) {
                    const cleanSheet = XLSX.utils.aoa_to_sheet(filtered);
                    tablesToRender.push({ html: XLSX.utils.sheet_to_html(cleanSheet), filename: cache.label });
                    state.loadedFiles.push(cache.label);
                }
            });
            state.loadedTables = tablesToRender.length;
            renderTables(tablesToRender);
            updateDropAreaDisplay();
            markSettingsClean();
            state.isProcessing = false;
            undoManager.showToast('重新套用清洗設定');
        }, 50);
    }

    function markSettingsDirty() {
        if (state.loadedTables > 0) {
            state.isSettingsDirty = true;
            if(elements.reapplyBanner) elements.reapplyBanner.classList.remove('hidden');
        }
    }

    function markSettingsClean() {
        state.isSettingsDirty = false;
        if(elements.reapplyBanner) elements.reapplyBanner.classList.add('hidden');
    }

    async function getSelectedSheetNames(filename, workbook, mode, criteria) {
        const names = workbook.SheetNames;
        if (!names.length) return [];
        if (mode === 'first') return [names[0]];
        if (mode === 'specific') return names.filter(n => n.toLowerCase().includes(criteria.name.toLowerCase()));
        if (mode === 'position') return utils.parsePositionString(criteria.position).map(i => names[i]).filter(Boolean);
        return names;
    }

    function renderTables(tables) {
        if (!elements.displayArea) return;
        if (!tables.length) { elements.displayArea.innerHTML = '<p>沒有找到符合條件的工作表。</p>'; return; }

        const fragment = document.createDocumentFragment();
        tables.forEach(({ html, filename }) => {
            const wrapper = document.createElement('div');
            wrapper.className = 'table-wrapper';
            
            const tempDiv = document.createElement('div');
            tempDiv.innerHTML = html;
            const table = tempDiv.querySelector('table');
            if (table) {
                const tbody = table.querySelector('tbody');
                if (tbody) {
                    const firstRow = tbody.querySelector('tr');
                    if (firstRow) {
                        const thead = document.createElement('thead');
                        thead.appendChild(firstRow);
                        firstRow.querySelectorAll('td').forEach(td => {
                            const th = document.createElement('th');
                            td.innerHTML = td.innerHTML.replace(/<br\s*[\/]?>/gi, '___NEWLINE___');
                            th.textContent = td.textContent.replace(/___NEWLINE___/g, '\n').trim();
                            th.style.cssText = td.style.cssText;
                            td.replaceWith(th);
                        });
                        table.insertBefore(thead, tbody);
                    }
                }
            }

            wrapper.innerHTML = `
                <div class="table-header">
                    <input type="checkbox" class="table-select-checkbox" title="選取此表格">
                    <h4>${filename}</h4>
                    <div class="header-actions">
                        <button class="btn btn-danger btn-sm delete-rows-btn">刪除選取列</button>
                        <button class="btn btn-danger btn-sm delete-table-btn">刪除此表</button>
                    </div>
                    <button class="close-zoom">&times;</button>
                </div>
                <div class="table-content"></div>`;
            
            if (table) wrapper.querySelector('.table-content').appendChild(table);
            fragment.appendChild(wrapper);
        });

        elements.displayArea.innerHTML = '';
        elements.displayArea.appendChild(fragment);
        state.originalHtmlString = elements.displayArea.innerHTML;
        
        injectCheckboxes(elements.displayArea);
        if (state.mainViewHiddenColumns.size > 0) applyMainViewColumnState();
        populateMainDropdowns();
        showControls(detectHiddenElements());
        sortTablesByFundName();
    }

    function injectCheckboxes(scope) {
        scope.querySelectorAll('thead tr').forEach((row, idx) => {
            if (row.querySelector('.checkbox-cell')) return;
            row.insertAdjacentHTML('afterbegin', `<th class="checkbox-cell"><input type="checkbox" id="select-all-cb-${scope.id}-${idx}"></th>`);
        });
        scope.querySelectorAll('tbody tr').forEach(row => {
            if (row.querySelector('.checkbox-cell')) return;
            row.insertAdjacentHTML('afterbegin', '<td class="checkbox-cell"><input type="checkbox" class="row-checkbox"></td>');
        });
    }

    function populateMainDropdowns() {
        const allHeaders = new Set();
        elements.displayArea.querySelectorAll('thead th:not(.checkbox-cell)').forEach(th => {
            allHeaders.add(th.textContent.trim().replace(/\n/g, ' '));
        });
        const options = Array.from(allHeaders).map(h => `<option value="${h}">${h}</option>`).join('');
        if (elements.mainColSelect1) elements.mainColSelect1.innerHTML = '<option value="">-- 選擇欄位 1 --</option>' + options;
        if (elements.mainColSelect2) elements.mainColSelect2.innerHTML = '<option value="">-- 選擇欄位 2 (選填) --</option>' + options;
    }

    function toggleMainEditMode(startEditing) {
        state.isMainEditing = startEditing;
        if(elements.mainEditDataBtn) elements.mainEditDataBtn.classList.toggle('hidden', startEditing);
        if(elements.mainSaveEditsBtn) elements.mainSaveEditsBtn.classList.toggle('hidden', !startEditing);
        elements.displayArea.querySelectorAll('td:not(.checkbox-cell)').forEach(td => {
            td.contentEditable = startEditing;
            td.style.backgroundColor = startEditing ? '#fffbeb' : ''; 
        });
        ['mainCopyRowsBtn', 'mainSmartDedupBtn', 'mainExecuteFilterBtn', 'deleteSelectedBtn', 'mainColumnOperationsBtn'].forEach(id => {
            if (elements[id]) elements[id].disabled = startEditing;
        });
    }

    function saveMainEdits() {
        elements.displayArea.querySelectorAll('td:not(.checkbox-cell)').forEach(td => {
            td.contentEditable = false; td.style.backgroundColor = '';
        });
        state.isMainEditing = false;
        if(elements.mainEditDataBtn) elements.mainEditDataBtn.classList.remove('hidden');
        if(elements.mainSaveEditsBtn) elements.mainSaveEditsBtn.classList.add('hidden');
        ['mainCopyRowsBtn', 'mainSmartDedupBtn', 'mainExecuteFilterBtn', 'deleteSelectedBtn', 'mainColumnOperationsBtn'].forEach(id => {
            if (elements[id]) elements[id].disabled = false;
        });
        state.originalHtmlString = elements.displayArea.innerHTML;
        undoManager.showToast('儲存主畫面編輯');
    }

    function copyMainSelectedRows() {
        const selected = elements.displayArea.querySelectorAll('tbody .row-checkbox:checked');
        if (!selected.length) { alert('請先勾選要複製的資料列。'); return; }
        selected.forEach(cb => {
            const tr = cb.closest('tr');
            const clone = tr.cloneNode(true);
            clone.querySelector('.row-checkbox').checked = false;
            clone.classList.add('new-row-highlight'); 
            tr.parentNode.insertBefore(clone, tr.nextSibling);
        });
        syncCheckboxesInScope();
        state.originalHtmlString = elements.displayArea.innerHTML;
    }

    function executeMainComplexFilter() {
        const col1 = elements.mainColSelect1?.value;
        let col2 = elements.mainColSelect2?.value;
        if (elements.mainCondition2Wrapper && elements.mainCondition2Wrapper.classList.contains('hidden')) {
            col2 = ''; 
        }
        
        const criteria1 = document.querySelector('input[name="main-criteria-1"]:checked')?.value;
        const criteria2 = document.querySelector('input[name="main-criteria-2"]:checked')?.value;
        const logicOp = document.querySelector('input[name="main-logic-op"]:checked')?.value ?? 'and';
        const inputVal1 = elements.mainInputCriteria1?.value;
        const inputVal2 = elements.mainInputCriteria2?.value;

        if (!col1 && !col2) { alert('請輸入至少一個條件'); return; }
        const checkValue = (cellVal, cr, val) => {
            const s = String(cellVal).trim(), v = String(val).trim();
            return cr === 'empty' ? s === '' : cr === 'zero' ? s === '0' : cr === 'value' ? s !== '' : cr === 'exact' ? s === v : cr === 'includes' ? v !== '' && s.toLowerCase().includes(v.toLowerCase()) : false;
        };

        let count = 0;
        elements.displayArea.querySelectorAll('tbody tr:not(.row-hidden-search)').forEach(row => {
            const checkbox = row.querySelector('.row-checkbox');
            if (!checkbox) return;
            const table = row.closest('table');
            const ths = Array.from(table.querySelectorAll('thead th:not(.checkbox-cell)'));
            const getCellText = (colName) => {
                if (!colName) return null;
                const colIdx = ths.findIndex(th => th.textContent.trim().replace(/\n/g, ' ') === colName);
                if (colIdx === -1) return null;
                return row.children[colIdx + 1] ? row.children[colIdx + 1].textContent : '';
            };
            const r1 = col1 ? checkValue(getCellText(col1) ?? '', criteria1, inputVal1) : null;
            const r2 = col2 ? checkValue(getCellText(col2) ?? '', criteria2, inputVal2) : null;
            const cMatch = col1 && col2 ? (logicOp === 'and' ? r1 && r2 : r1 || r2) : (col1 ? r1 : r2);
            if (cMatch) { checkbox.checked = true; count++; }
        });
        alert(count > 0 ? `已勾選 ${count} 筆` : '未找到符合條件的資料');
        syncCheckboxesInScope();
    }

    function openMainColumnModal() {
        const allTables = elements.displayArea.querySelectorAll('.table-wrapper table');
        if (allTables.length === 0) { alert('目前沒有可操作的表格。'); return; }
        const allHeaders = new Set();
        allTables.forEach(table => { table.querySelectorAll('thead th:not(.checkbox-cell)').forEach(th => allHeaders.add(th.textContent.trim())); });

        state.columnModalContext = 'main';
        if(elements.columnChecklist) {
            elements.columnChecklist.innerHTML = Array.from(allHeaders).map(h => {
                const displayH = h.replace(/\n/g, ' '); 
                const isChecked = !state.mainViewHiddenColumns.has(h);
                return `<label><input type="checkbox" value="${h}" ${isChecked ? 'checked' : ''}> ${displayH}</label>`;
            }).join('');
        }
        toggleColumnModal(true);
    }

    function openMergeColumnModal() {
        state.columnModalContext = 'merge';
        updateColumnSelects(state.mergedHeaders);
        toggleColumnModal(true);
    }

    function toggleColumnModal(show) { if (elements.columnModal) elements.columnModal.classList.toggle('hidden', !show); }

    function applyColumnChanges() {
        if (state.columnModalContext === 'main') {
            elements.columnChecklist.querySelectorAll('input').forEach(input => {
                if (!input.checked) state.mainViewHiddenColumns.add(input.value); else state.mainViewHiddenColumns.delete(input.value);
            });
            applyMainViewColumnState();
        } else {
            const mergedTable = elements.mergeViewContent.querySelector('table');
            if (!mergedTable) return;
            const visibility = {};
            elements.columnChecklist.querySelectorAll('input').forEach(input => visibility[input.value] = input.checked);
            
            const allThs = Array.from(mergedTable.querySelectorAll('thead th'));
            const firstDataIdx = allThs.findIndex(th => !th.classList.contains('checkbox-cell') && !th.classList.contains('source-col'));
            if (firstDataIdx === -1) return;
            allThs.slice(firstDataIdx).forEach((th, i) => {
                const colIdx = i + firstDataIdx;
                const headerText = th.textContent.replace('×', '').trim();
                mergedTable.querySelectorAll(`tr > *:nth-child(${colIdx + 1})`).forEach(cell => cell.classList.toggle('column-hidden', !visibility[headerText]));
            });
        }
    }

    function applyMainViewColumnState() {
        const allTables = elements.displayArea.querySelectorAll('.table-wrapper table');
        allTables.forEach(table => {
            const headers = Array.from(table.querySelectorAll('thead th'));
            headers.forEach((th, colIdx) => {
                if (th.classList.contains('checkbox-cell')) return;
                const isHidden = state.mainViewHiddenColumns.has(th.textContent.trim());
                th.classList.toggle('column-hidden', isHidden);
                table.querySelectorAll(`tbody tr`).forEach(tr => {
                    const td = tr.children[colIdx];
                    if (td) td.classList.toggle('column-hidden', isHidden);
                });
            });
        });
    }

    function updateColumnSelects(headers) {
        if(elements.columnChecklist) {
            elements.columnChecklist.innerHTML = headers.map(h => {
                const displayH = h.replace(/\n/g, ' '); 
                return `<label><input type="checkbox" value="${h}" checked> ${displayH}</label>`;
            }).join('');
        }
        const makeOption = (value, text) => { const opt = document.createElement('option'); opt.value = value; opt.textContent = text.replace(/\n/g, ' '); return opt; };
        if(elements.colSelect1) { elements.colSelect1.innerHTML = ''; elements.colSelect1.appendChild(makeOption('', '-- 選擇欄位 1 --')); headers.forEach(h => elements.colSelect1.appendChild(makeOption(h, h))); }
        if(elements.colSelect2) { elements.colSelect2.innerHTML = ''; elements.colSelect2.appendChild(makeOption('', '-- 選擇欄位 2 (選填) --')); headers.forEach(h => elements.colSelect2.appendChild(makeOption(h, h))); }
    }

    function createMergedView(mode = 'all') {
        if (!elements.displayArea || !elements.mergeViewModal) return;
        const tables = Array.from(elements.displayArea.querySelectorAll('.table-wrapper:not([style*="display: none"]) table'));
        if (!tables.length) { alert('沒有可合併的表格。'); return; }

        if (state.mergedData.length > 0 && state.isMergedView === false) {
            if (confirm('偵測到上次的合併紀錄，是否接續上次編輯狀態？\n(按「取消」將放棄舊狀態，重新合併最新主表)')) {
                elements.mergeViewModal.classList.remove('hidden');
                document.body.classList.add('no-scroll');
                state.isMergedView = true;
                renderMergedTable();
                return;
            }
        }

        if (mode === 'checked') {
            const checkedCount = tables.reduce((n, t) => n + t.querySelectorAll('tbody .row-checkbox:checked').length, 0);
            if (!checkedCount) { alert('請先勾選至少一個資料列。'); return; }
        }

        const allHeaders = new Set();
        const tableHeaderMap = new Map();

        tables.forEach(table => {
            let headers = Array.from(table.querySelectorAll('thead th:not(.checkbox-cell)')).map((th, i) => th.textContent.trim() || `(欄位 ${i + 1})`);
            if (!headers.length) headers = Array.from({ length: 10 }, (_, i) => `(欄位 ${i + 1})`);
            headers.forEach(h => allHeaders.add(h));
            tableHeaderMap.set(table, headers);
        });

        const tableData = [];
        tables.forEach(table => {
            const headers = tableHeaderMap.get(table);
            if (!headers) return;
            const filename = table.closest('.table-wrapper')?.querySelector('h4')?.textContent || '未知來源';
            const rows = mode === 'all' ? Array.from(table.querySelectorAll('tbody tr')).filter(r => !r.classList.contains('row-hidden-search')) : Array.from(table.querySelectorAll('tbody .row-checkbox:checked')).map(cb => cb.closest('tr'));

            rows.forEach(row => {
                const rowData = { _sourceFile: filename, _id: Date.now() + Math.random() };
                Array.from(row.querySelectorAll('td:not(.checkbox-cell)')).forEach((td, i) => { if (headers[i]) rowData[headers[i]] = td.textContent; });
                tableData.push(rowData);
            });
        });

        state.mergedHeaders = Array.from(allHeaders);
        state.mergedData = tableData;
        renderMergedTable();
        updateColumnSelects(state.mergedHeaders);
        
        state.isMergedView = true;
        elements.mergeViewModal.classList.remove('hidden');
        document.body.classList.add('no-scroll');
    }

    function closeMergeView() {
        if (!elements.mergeViewModal) return;
        if (state.isEditing && !confirm('確定要放棄目前的編輯狀態？')) return;
        
        if (elements.mergeCondition2Wrapper) {
            elements.mergeCondition2Wrapper.classList.add('hidden');
            if (elements.toggleMergeConditionBtn) elements.toggleMergeConditionBtn.textContent = '+ 新增第二條件';
        }

        elements.mergeViewModal.classList.add('hidden');
        document.body.classList.remove('no-scroll');
        state.isMergedView = false;
        toggleEditMode(false);
    }

    function renderMergedTable() {
        if (!elements.mergeViewContent) return;
        const table = document.createElement('table');
        const thead = table.createTHead();
        const headerRow = thead.insertRow();

        if (state.showSourceColumn) headerRow.insertAdjacentHTML('beforeend', '<th class="source-col">來源檔案</th>');
        state.mergedHeaders.forEach(header => { 
            headerRow.insertAdjacentHTML('beforeend', `<th>${header}<span class="delete-col-btn" data-header="${header}">&times;</span></th>`); 
        });

        const tbody = table.createTBody();
        state.mergedData.forEach((rowData, index) => {
            const tr = tbody.insertRow();
            tr.dataset.rowIndex = index;
            if (rowData._sourceFile) tr.title = `來源: ${rowData._sourceFile}`;
            if (rowData._isNew) tr.classList.add('new-row-highlight');

            if (state.showSourceColumn) {
                const td = document.createElement('td');
                td.textContent = rowData._sourceFile || '';
                td.classList.add('source-col');
                tr.prepend(td);
            }

            state.mergedHeaders.forEach(header => {
                const td = tr.insertCell();
                td.contentEditable = state.isEditing;
                td.dataset.colHeader = header;
                const value = rowData[header] || '';
                if (utils.isStrictNumber(value)) {
                    td.textContent = utils.formatNumber(value);
                    td.classList.add('numeric');
                } else { td.textContent = value; }
            });
        });

        if (state.showTotalRow) {
            const tfoot = table.createTFoot();
            const totalRow = tfoot.insertRow();
            totalRow.className = 'total-row';
            totalRow.insertCell(); 
            if (state.showSourceColumn) totalRow.insertCell().textContent = '';

            const totals = calculateTotals();
            let labelApplied = false;
            state.mergedHeaders.forEach(header => {
                const td = totalRow.insertCell();
                if (totals[header] !== undefined) {
                    td.textContent = utils.formatNumber(totals[header]);
                    td.classList.add('numeric');
                } else if (!labelApplied) { td.textContent = '合計'; labelApplied = true; }
            });
        }

        elements.mergeViewContent.innerHTML = '';
        elements.mergeViewContent.appendChild(table);
        elements.mergeViewContent.classList.toggle('is-editing', state.isEditing);
        injectCheckboxes(elements.mergeViewContent);

        const selectAllCb = elements.mergeViewContent.querySelector('thead input[type="checkbox"]');
        if (selectAllCb) selectAllCb.addEventListener('change', e => { elements.mergeViewContent.querySelectorAll('.row-checkbox').forEach(cb => cb.checked = e.target.checked); });
    }

    function toggleEditMode(startEditing) {
        state.isEditing = startEditing;
        if(elements.editDataBtn) elements.editDataBtn.classList.toggle('hidden', startEditing);
        if(elements.saveEditsBtn) elements.saveEditsBtn.classList.toggle('hidden', !startEditing);
        if(elements.cancelEditsBtn) elements.cancelEditsBtn.classList.toggle('hidden', !startEditing);
        if(elements.mergeViewContent) {
            const cells = elements.mergeViewContent.querySelectorAll('td:not(.checkbox-cell):not(.source-col)');
            cells.forEach(td => { td.contentEditable = startEditing; td.style.backgroundColor = startEditing ? '#fffbeb' : ''; });
        }
    }

    function saveEdits() {
        if(elements.mergeViewContent) {
            const rows = elements.mergeViewContent.querySelectorAll('tbody tr');
            state.mergedData = [];
            rows.forEach(tr => {
                const rowData = { _sourceFile: tr.dataset.sourceFile || '', _id: tr.dataset.rowIndex || Date.now() + Math.random() };
                Array.from(tr.querySelectorAll('td:not(.checkbox-cell):not(.source-col)')).forEach((td, i) => {
                    if (state.mergedHeaders[i]) rowData[state.mergedHeaders[i]] = td.textContent;
                });
                state.mergedData.push(rowData);
            });
        }
        toggleEditMode(false);
        undoManager.showToast('儲存編輯');
    }

    function addNewRow() {
        if (!state.isMergedView) return;
        const newRow = { _sourceFile: '', _id: Date.now() + Math.random() };
        state.mergedHeaders.forEach(h => newRow[h] = '');
        state.mergedData.push(newRow);
        renderMergedTable();
    }

    function copySelectedRows() {
        if (!state.isMergedView) {
            if(elements.displayArea) {
                const selected = elements.displayArea.querySelectorAll('tbody .row-checkbox:checked');
                if (!selected.length) { alert('請先勾選要複製的列'); return; }
                selected.forEach(cb => {
                    const tr = cb.closest('tr');
                    const clone = tr.cloneNode(true);
                    clone.querySelector('.row-checkbox').checked = false;
                    clone.classList.add('new-row-highlight');
                    tr.parentNode.insertBefore(clone, tr.nextSibling);
                });
                syncCheckboxesInScope();
                state.originalHtmlString = elements.displayArea.innerHTML;
            }
            return;
        }
        const selected = elements.mergeViewContent?.querySelectorAll('tbody .row-checkbox:checked');
        if (!selected?.length) { alert('請先勾選要複製的列'); return; }
        selected.forEach(cb => {
            const idx = parseInt(cb.closest('tr').dataset.rowIndex, 10);
            const row = state.mergedData[idx];
            if (row) {
                const copy = { ...row, _id: Date.now() + Math.random() };
                state.mergedData.splice(idx + 1, 0, copy);
            }
        });
        renderMergedTable();
    }

    function toggleSourceColumn() {
        state.showSourceColumn = !state.showSourceColumn;
        if(elements.toggleSourceColBtn) elements.toggleSourceColBtn.classList.toggle('active', state.showSourceColumn);
        renderMergedTable();
    }

    function deleteSelectedRows(specificScope = null) {
        const scope = specificScope || getActiveScope();
        if(!scope) return;
        const selected = scope.querySelectorAll('tbody .row-checkbox:checked');
        if (!selected.length) { alert('請先勾選列'); return; }

        if (state.isMergedView && !specificScope) {
            const backupData = [...state.mergedData];
            const toDelete = new Set(Array.from(selected).map(cb => parseInt(cb.closest('tr').dataset.rowIndex, 10)));
            undoManager.push(`刪除 ${toDelete.size} 列 (合併視圖)`, () => { state.mergedData = backupData; if (state.isMergedView) renderMergedTable(); });
            state.mergedData = state.mergedData.filter((_, i) => !toDelete.has(i));
            renderMergedTable();
            if(elements.dedupResultPanel) elements.dedupResultPanel.classList.add('hidden');
        } else {
            const rows = Array.from(selected).map(cb => cb.closest('tr'));
            const backups = rows.map(tr => ({ parent: tr.parentNode, sibling: tr.nextSibling, node: tr }));
            undoManager.push(`刪除 ${rows.length} 列 (主表)`, () => { backups.forEach(b => b.parent.insertBefore(b.node, b.sibling)); syncCheckboxesInScope(); });
            rows.forEach(tr => tr.remove());
            if(!state.isMergedView) state.originalHtmlString = elements.displayArea.innerHTML;
        }
        syncCheckboxesInScope();
    }

    function deleteSelectedTables(specificTableWrapper = null) {
        if(!elements.displayArea) return;
        const selected = specificTableWrapper ? [specificTableWrapper] : Array.from(elements.displayArea.querySelectorAll('.table-select-checkbox:checked')).map(cb => cb.closest('.table-wrapper'));
        if (!selected.length) return;
        
        const backups = selected.map(w => ({ parent: w.parentNode, sibling: w.nextSibling, node: w }));
        const prevCount = state.loadedTables;
        const prevFiles = [...state.loadedFiles];

        undoManager.push(`刪除 ${selected.length} 個工作表`, () => {
            backups.forEach(b => b.parent.insertBefore(b.node, b.sibling));
            state.loadedTables = prevCount; state.loadedFiles = prevFiles;
            updateDropAreaDisplay(); showControls(detectHiddenElements());
        });

        selected.forEach(w => w.remove());
        updateFileStateAfterDeletion();
    }

    function deleteColumn(headerToDelete) {
        if (!confirm(`確定要刪除「${headerToDelete.replace(/\n/g, ' ')}」欄位？`)) return;
        const backupHeaders = [...state.mergedHeaders];
        const backupData = state.mergedData.map(row => ({...row}));

        undoManager.push(`刪除欄位: ${headerToDelete.replace(/\n/g, ' ')}`, () => {
            state.mergedHeaders = backupHeaders; state.mergedData = backupData;
            if (state.isMergedView) { renderMergedTable(); updateColumnSelects(state.mergedHeaders); }
        });

        state.mergedHeaders = state.mergedHeaders.filter(h => h !== headerToDelete);
        state.mergedData.forEach(row => delete row[headerToDelete]);
        renderMergedTable();
        updateColumnSelects(state.mergedHeaders);
    }

    function syncCheckboxesInScope() {
        setTimeout(() => {
            const scope = getActiveScope();
            if(!scope) return;
            scope.querySelectorAll('table').forEach(t => {
                const rowCbs = t.querySelectorAll('tbody tr:not(.row-hidden-search) .row-checkbox');
                const hcb = t.closest('.table-wrapper')?.querySelector('.table-select-checkbox') ?? t.querySelector('thead input[type="checkbox"]');
                if (!hcb) return;
                const checkedCount = Array.from(rowCbs).filter(cb => cb.checked).length;
                hcb.checked = rowCbs.length > 0 && checkedCount === rowCbs.length;
                hcb.indeterminate = checkedCount > 0 && checkedCount < rowCbs.length;
            });
            if (!state.isMergedView) updateSelectionInfo();
        }, 0);
    }

    function selectAllRows() {
        const scope = getActiveScope();
        if(!scope) return;
        scope.querySelectorAll('tbody tr:not(.row-hidden-search)').forEach(row => { const cb = row.querySelector('.row-checkbox'); if (cb) cb.checked = true; });
        scope.querySelectorAll('thead input[type="checkbox"]').forEach(cb => cb.checked = true);
    }

    function invertSelection() { const scope = getActiveScope(); if(!scope) return; scope.querySelectorAll('tbody tr:not(.row-hidden-search) .row-checkbox').forEach(cb => { cb.checked = !cb.checked; }); }

    function selectEmptyRows() {
        const scope = getActiveScope();
        if(!scope) return;
        let count = 0;
        scope.querySelectorAll('tbody tr:not(.row-hidden-search)').forEach(row => {
            if (Array.from(row.cells).slice(1).every(c => c.textContent.trim() === '')) { row.querySelector('.row-checkbox').checked = true; count++; }
        });
        if (!count) alert('未找到空白列');
    }

    function unselectAllMergedRows() {
        if (!state.isMergedView || !elements.mergeViewContent) return;
        elements.mergeViewContent.querySelectorAll('.row-checkbox:checked').forEach(cb => cb.checked = false);
        const hcb = elements.mergeViewContent.querySelector('thead input[type="checkbox"]');
        if (hcb) { hcb.checked = false; hcb.indeterminate = false; }
    }

    function selectByKeyword() {
        const inputEl = state.isMergedView ? elements.selectKeywordInputMerged : elements.selectKeywordInput;
        const regexEl = state.isMergedView ? elements.selectKeywordRegexMerged : elements.selectKeywordRegex;
        if(!inputEl) return;
        const keyword = inputEl.value.trim();
        if (!keyword) { alert('請輸入關鍵字'); return; }

        let matcher;
        try { matcher = utils.buildKeywordMatcher(keyword, regexEl?.checked || false); } catch (e) { alert('Regex 錯誤：' + e.message); return; }

        let count = 0;
        const scope = getActiveScope();
        if(!scope) return;
        scope.querySelectorAll('tbody tr:not(.row-hidden-search)').forEach(row => {
            const text = Array.from(row.querySelectorAll('td:not(.checkbox-cell)')).map(c => c.textContent).join(' ').toLowerCase();
            if (matcher && matcher(text)) { row.querySelector('.row-checkbox').checked = true; count++; }
        });
        alert(count > 0 ? `已勾選 ${count} 列` : '未找到相符資料');
    }

    function filterTable() {
        const inputEl = state.isMergedView ? elements.searchInputMerged : elements.searchInput;
        if (!inputEl) return;
        const keywords = inputEl.value.toLowerCase().trim().split(/\s+/).filter(Boolean);
        const scope = getActiveScope();
        if(!scope) return;

        scope.querySelectorAll('tbody tr').forEach(row => {
            const text = Array.from(row.querySelectorAll('td:not(.checkbox-cell)')).map(c => c.textContent).join(' ').toLowerCase();
            row.classList.toggle('row-hidden-search', !keywords.every(k => text.includes(k)));
        });

        if (!state.isMergedView && elements.displayArea) {
            elements.displayArea.querySelectorAll('.table-wrapper').forEach(wrapper => {
                wrapper.style.display = wrapper.querySelectorAll('tbody tr:not(.row-hidden-search)').length > 0 ? '' : 'none';
            });
        }
        syncCheckboxesInScope();
    }

    function updateDropAreaDisplay() {
        const hasFiles = state.loadedTables > 0;
        if(elements.dropArea) elements.dropArea.classList.toggle('compact', hasFiles);
        // 支援舊版 HTML 結構
        if(elements.dropAreaInitial) elements.dropAreaInitial.classList.toggle('hidden', hasFiles);
        if(elements.dropAreaLoaded) elements.dropAreaLoaded.classList.toggle('hidden', !hasFiles);
        if(elements.importOptionsContainer) elements.importOptionsContainer.classList.toggle('hidden', hasFiles);
        if (hasFiles && elements.fileCount && elements.fileNames) {
            elements.fileCount.textContent = state.loadedTables;
            elements.fileNames.textContent = state.loadedFiles.slice(0, 3).join(', ') + (state.loadedFiles.length > 3 ? '...' : '');
        }
        // 新版 HTML: 簡化 drop area 內容
        const dropContent = document.getElementById('drop-area-content');
        if (dropContent) {
            if (hasFiles) {
                dropContent.innerHTML = `<h3>📊 已載入 ${state.loadedTables} 張表</h3><p>拖放更多檔案或點擊上傳</p>`;
            }
        }
    }

    function showControls(hiddenCount) {
        if(elements.controlPanel) elements.controlPanel.classList.remove('hidden');
        if(elements.stepData) elements.stepData.classList.remove('hidden');
        if(elements.stepExport) elements.stepExport.classList.remove('hidden');
        if(elements.mergeViewBtn) elements.mergeViewBtn.classList.toggle('hidden', state.loadedTables <= 1);
        if(elements.showHiddenBtn) elements.showHiddenBtn.classList.toggle('hidden', hiddenCount === 0);
    }

    function updateSelectionInfo() {
        if(!elements.displayArea) return;
        const selected = elements.displayArea.querySelectorAll('.table-select-checkbox:checked, .table-select-checkbox:indeterminate');
        if(elements.selectedTablesList) elements.selectedTablesList.textContent = Array.from(selected).map(cb => cb.closest('.table-header').querySelector('h4').textContent).join('; ');
        if(elements.selectedTablesInfo) elements.selectedTablesInfo.classList.toggle('hidden', selected.length === 0);
    }

    function detectHiddenElements() { return elements.displayArea ? elements.displayArea.querySelectorAll('tr[style*="display: none"], td[style*="display: none"], th[style*="display: none"]').length : 0; }

    function updateFileStateAfterDeletion() {
        if(!elements.displayArea) return;
        state.loadedTables = elements.displayArea.querySelectorAll('.table-wrapper').length;
        if (!state.loadedTables) clearAllFiles(true); else { updateDropAreaDisplay(); updateSelectionInfo(); }
    }

    function clearAllFiles(silent = false) {
        if (!silent && !confirm('確定清除所有檔案？')) return;
        if (state.isMergedView) closeMergeView();
        state.originalHtmlString = ''; state.loadedFiles = []; state.loadedTables = 0; state.rawSheetsCache = [];
        state.fileInfos = []; state.aggregatedRows = [];
        if(elements.displayArea) elements.displayArea.innerHTML = '';
        if (elements.fileInput) elements.fileInput.value = '';
        if(elements.stepMapping) elements.stepMapping.classList.add('hidden');
        if(elements.stepData) elements.stepData.classList.add('hidden');
        if(elements.stepExport) elements.stepExport.classList.add('hidden');
        if(elements.aggregatePanel) elements.aggregatePanel.classList.add('hidden');
        if(elements.topBarActions) elements.topBarActions.classList.add('hidden');
        updateDropAreaDisplay(); resetControls(); setViewMode('list');
    }

    function setViewMode(mode) {
        const isGrid = mode === 'grid';
        if(elements.displayArea) { elements.displayArea.classList.toggle('grid-view', isGrid); elements.displayArea.classList.toggle('list-view', !isGrid); }
        if(elements.gridViewBtn) elements.gridViewBtn.classList.toggle('active', isGrid);
        if(elements.listViewBtn) elements.listViewBtn.classList.toggle('active', !isGrid);
        if(elements.gridScaleControl) elements.gridScaleControl.classList.toggle('hidden', !isGrid);
    }

    function updateGridScale() { if(elements.displayArea && elements.gridScaleSlider) elements.displayArea.style.setProperty('--grid-columns', elements.gridScaleSlider.value); }

    function showAllHiddenElements() {
        if(!elements.displayArea) return;
        const hidden = elements.displayArea.querySelectorAll('tr[style*="display: none"], td[style*="display: none"], th[style*="display: none"]');
        if (!hidden.length) return;
        hidden.forEach(el => el.style.display = '');
        if(elements.showHiddenBtn) elements.showHiddenBtn.classList.add('hidden');
        if(elements.loadStatusMessage) elements.loadStatusMessage.classList.add('hidden');
    }

    function toggleToolbar() { if(elements.collapsibleToolbar) { const collapsed = elements.collapsibleToolbar.classList.toggle('collapsed'); if (elements.toggleToolbarBtn) elements.toggleToolbarBtn.textContent = collapsed ? '展開工具列' : '收合工具列'; } }
    function scrollToTop() { window.scrollTo({ top: 0, behavior: 'smooth' }); }
    function handleScroll() { if(elements.backToTopBtn) elements.backToTopBtn.classList.toggle('visible', window.scrollY > window.innerHeight / 2); }

    function openPreview(card) {
        if (state.zoomedCard) return;
        card.classList.add('is-zoomed'); state.zoomedCard = card; document.body.classList.add('no-scroll');
    }

    function closePreview() {
        if (!state.zoomedCard) return;
        state.zoomedCard.classList.remove('is-zoomed'); state.zoomedCard = null; document.body.classList.remove('no-scroll');
    }

    function resetView() {
        if (state.isMergedView) closeMergeView();
        if (!state.originalHtmlString || !elements.displayArea) return;
        elements.displayArea.innerHTML = state.originalHtmlString;
        injectCheckboxes(elements.displayArea);
        ['searchInput', 'selectKeywordInput'].forEach(id => { if(elements[id]) elements[id].value = ''; });
        if(elements.selectKeywordRegex) elements.selectKeywordRegex.checked = false;
        filterTable(); updateSelectionInfo(); setViewMode('list');
    }

    function handleCriteriaChange(e) {
        const radio = e.target;
        if (radio.type !== 'radio') return;
        const group = radio.closest('.radio-group');
        if (!group) return;
        const target = elements[group.dataset.target];
        if (!target) return;
        const needsInput = radio.value === 'exact' || radio.value === 'includes';
        target.disabled = !needsInput;
        if (needsInput) { target.focus(); } else { target.value = ''; }
    }

    function extractTableData(table, { onlySelected = false, includeFilename = false } = {}) {
        const data = [];
        const headerRow = table.querySelector('thead tr');
        if (headerRow) {
            const headers = Array.from(headerRow.querySelectorAll('th:not(.checkbox-cell):not(.column-hidden)')).map(th => th.textContent.replace('×', '').trim());
            if (includeFilename) headers.unshift('Source File');
            data.push(headers);
        }
        const filename = includeFilename ? (table.closest('.table-wrapper')?.querySelector('h4')?.textContent || 'Merged Table') : null;
        const rows = onlySelected ? Array.from(table.querySelectorAll('tbody .row-checkbox:checked')).map(cb => cb.closest('tr')) : Array.from(table.querySelectorAll('tbody tr:not(.row-hidden-search)'));

        rows.forEach(row => {
            const cells = Array.from(row.querySelectorAll('td:not(.checkbox-cell):not(.column-hidden)')).map(td => {
                const val = td.textContent.trim();
                const clean = val.replace(/,/g, '');
                return clean !== '' && !isNaN(clean) ? Number(clean) : val;
            });
            if (includeFilename) cells.unshift(filename);
            data.push(cells);
        });
        return data;
    }

    function exportToXlsx(data, filename, sheetName) {
        if (data.length <= 1) { alert('無資料可匯出。'); return; }
        try {
            const ws = XLSX.utils.aoa_to_sheet(data);
            ws['!cols'] = utils.calculateColumnWidths(data);
            const wb = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(wb, ws, sheetName);
            XLSX.writeFile(wb, filename);
        } catch (err) { alert('匯出錯誤：' + err.message); }
    }

    function exportCurrentMergedXlsx() {
        if (!state.isMergedView || !elements.mergeViewContent) return;
        const table = elements.mergeViewContent.querySelector('table');
        if (!table) return;
        exportToXlsx(extractTableData(table, { includeFilename: state.showSourceColumn }), `merged_view_${new Date().toISOString().slice(0, 10)}.xlsx`, 'Merged Data');
    }

    function exportMergedXlsx() {
        if (!elements.displayArea) return;
        const tables = Array.from(elements.displayArea.querySelectorAll('.table-wrapper:not([style*="display: none"]) table'));
        if (!tables.length) { alert('沒有可匯出的表格。'); return; }
        const allData = [];
        tables.forEach((table, i) => {
            const data = extractTableData(table, { includeFilename: true });
            if (data.length > 1) allData.push(...(i === 0 ? data : data.slice(1)));
        });
        exportToXlsx(allData, `report_${new Date().toISOString().slice(0, 10)}.xlsx`, 'Data');
    }

    function getFundSortPriority(fileName) {
        if (!state.fundSortOrder || state.fundSortOrder.length === 0) return { index: Infinity, name: fileName };
        const alias = state.fundAliasKeys.find(a => fileName.includes(a));
        const canonical = alias ? state.fundAliasMap[alias] : null;
        const index = canonical ? state.fundSortOrder.indexOf(canonical) : -1;
        return { index: index === -1 ? Infinity : index, name: fileName };
    }

    function sortTablesByFundName() {
        if (!state.fundSortOrder || state.fundSortOrder.length === 0) return;
        if (!elements.displayArea) return;
        const wrappers = Array.from(elements.displayArea.querySelectorAll('.table-wrapper'));
        wrappers.sort((a, b) => {
            const h4A = a.querySelector('h4');
            const h4B = b.querySelector('h4');
            const fa = getFundSortPriority(h4A ? h4A.textContent : '');
            const fb = getFundSortPriority(h4B ? h4B.textContent : '');
            return fa.index !== fb.index ? fa.index - fb.index : fa.name.localeCompare(fb.name);
        });
        elements.displayArea.innerHTML = '';
        wrappers.forEach(w => elements.displayArea.appendChild(w));
    }

    function sortMergedTableByFundName() {
        if (state.isEditing) { alert('請先儲存或取消編輯。'); return; }
        if (!state.fundSortOrder || state.fundSortOrder.length === 0) return;
        state.mergedData.sort((a, b) => {
            const fa = getFundSortPriority(a._sourceFile || '');
            const fb = getFundSortPriority(b._sourceFile || '');
            return fa.index !== fb.index ? fa.index - fb.index : (a._sourceFile || '').localeCompare(b._sourceFile || '');
        });
        renderMergedTable();
    }

    function handleMergedHeaderClick(th) {
        if (state.isEditing || th.classList.contains('source-col')) return;
        const table = th.closest('table');
        const header = th.textContent.replace('×', '').trim();
        const isAsc = th.classList.contains('sort-asc');
        table.querySelectorAll('th').forEach(h => h.classList.remove('sort-asc', 'sort-desc'));
        th.classList.add(isAsc ? 'sort-desc' : 'sort-asc');

        state.mergedData.sort((a, b) => {
            const va = a[header] || '', vb = b[header] || '';
            const na = parseFloat(String(va).replace(/,/g, '')), nb = parseFloat(String(vb).replace(/,/g, ''));
            const cmp = !isNaN(na) && !isNaN(nb) ? na - nb : va.localeCompare(vb, undefined, { numeric: true, sensitivity: 'base' });
            return isAsc ? -cmp : cmp;
        });
        renderMergedTable();
    }

    function executeSmartDeduplication() {
        const keyCol = elements.dedupColSelect?.value;
        if (!keyCol) return;

        const groups = {};
        let markedCount = 0;

        if (state.dedupContext === 'main') {
            const rows = Array.from(elements.displayArea.querySelectorAll('tbody tr:not(.row-hidden-search)'));
            rows.forEach(tr => tr.classList.remove('dedup-marked'));

            rows.forEach(tr => {
                const table = tr.closest('table');
                const ths = Array.from(table.querySelectorAll('thead th:not(.checkbox-cell)'));
                const colIdx = ths.findIndex(th => th.textContent.trim().replace(/\n/g, ' ') === keyCol);
                if (colIdx === -1) return;

                const dataTd = tr.children[colIdx + 1]; 
                if (!dataTd) return;
                const key = dataTd.textContent.trim();
                if (!key) return;

                const sourceFile = tr.closest('.table-wrapper').querySelector('h4').textContent;
                if (!groups[key]) groups[key] = [];
                groups[key].push({ tr, sourceFile });
            });

            Object.entries(groups).forEach(([key, items]) => {
                if (items.length <= 1) return;
                const cleanKey = key.replace(/[\s_基金]/g, '');
                let best = items.find(({ sourceFile }) => {
                    const src = String(sourceFile).replace(/[\s_]/g, '');
                    return src.includes(cleanKey) || cleanKey.includes(src);
                }) ?? items[0];

                items.forEach(item => {
                    if (item === best) return;
                    const cb = item.tr.querySelector('.row-checkbox');
                    if (cb && !cb.checked) { 
                        cb.checked = true; 
                        item.tr.classList.add('dedup-marked'); 
                        markedCount++; 
                    }
                });
            });
        } else {
            const rows = Array.from(elements.mergeViewContent.querySelectorAll('tbody tr:not(.row-hidden-search)'));
            rows.forEach(tr => tr.classList.remove('dedup-marked'));

            rows.forEach(tr => {
                const idx = parseInt(tr.dataset.rowIndex, 10);
                const rowData = state.mergedData[idx];
                if (!rowData) return;
                const key = String(rowData[keyCol] || '').trim();
                if (!key) return;
                if (!groups[key]) groups[key] = [];
                groups[key].push({ tr, rowData });
            });

            Object.entries(groups).forEach(([key, items]) => {
                if (items.length <= 1) return;
                const cleanKey = key.replace(/[\s_基金]/g, '');
                let best = items.find(({ rowData }) => {
                    const src = String(rowData._sourceFile || '').replace(/[\s_]/g, '');
                    return src.includes(cleanKey) || cleanKey.includes(src);
                }) ?? items[0];

                items.forEach(item => {
                    if (item === best) return;
                    const cb = item.tr.querySelector('.row-checkbox');
                    if (cb && !cb.checked) { 
                        cb.checked = true; 
                        item.tr.classList.add('dedup-marked'); 
                        markedCount++; 
                    }
                });
            });
        }

        if(elements.dedupModal) elements.dedupModal.classList.add('hidden');
        syncCheckboxesInScope();

        if (markedCount > 0) {
            if (elements.dedupResultText) {
                elements.dedupResultText.innerHTML = `🎯 <b>智慧去重完成：</b> 已為您自動標記並勾選了 <b>${markedCount}</b> 筆不符合來源規則的舊資料。`;
            } else {
                alert(`🎯 智慧去重完成：\n已為您自動標記並勾選了 ${markedCount} 筆不符合來源規則的舊資料。`);
            }
            if(elements.dedupResultPanel) elements.dedupResultPanel.classList.remove('hidden');
        } else {
            undoManager.showToast('未發現需要處理的重複資料');
        }
    }

    function clearDedupMarks() {
        const scope = getActiveScope();
        if(!scope) return;
        scope.querySelectorAll('.dedup-marked').forEach(tr => {
            tr.classList.remove('dedup-marked');
            const cb = tr.querySelector('.row-checkbox');
            if (cb) cb.checked = false;
        });
        if(elements.dedupResultPanel) elements.dedupResultPanel.classList.add('hidden');
        syncCheckboxesInScope();
    }

    function bindEvents() {
        setupDragAndDrop();

        if(elements.clearFilesBtn) elements.clearFilesBtn.addEventListener('click', () => clearAllFiles(false));

        const importSettingsPanel = document.getElementById('import-settings-panel');
        if(importSettingsPanel) {
            importSettingsPanel.addEventListener('change', e => {
                if (e.target.name !== 'import-mode') return;
                const mode = e.target.value;
                if(elements.specificSheetNameGroup) elements.specificSheetNameGroup.classList.toggle('hidden', mode !== 'specific');
                if(elements.specificSheetPositionGroup) elements.specificSheetPositionGroup.classList.toggle('hidden', mode !== 'position');
            });
        }

        ['skipTopRowsCheckbox', 'discardRowsInput', 'headerRowsInput', 'removeEmptyRowsCheckbox', 'removeKeywordRowsCheckbox', 'removeKeywordRowsInput'].forEach(id => {
            if(elements[id]) {
                elements[id].addEventListener('change', markSettingsDirty);
                elements[id].addEventListener('input', utils.debounce(markSettingsDirty, 500));
            }
        });
        
        if(elements.skipTopRowsCheckbox) {
            elements.skipTopRowsCheckbox.addEventListener('change', e => { 
                if (elements.discardRowsInput) elements.discardRowsInput.disabled = !e.target.checked; 
                if (elements.headerRowsInput) elements.headerRowsInput.disabled = !e.target.checked; 
            });
        }

        if(elements.removeKeywordRowsCheckbox) elements.removeKeywordRowsCheckbox.addEventListener('change', e => { if (elements.removeKeywordRowsInput) elements.removeKeywordRowsInput.disabled = !e.target.checked; });
        if(elements.reapplySettingsBtn) elements.reapplySettingsBtn.addEventListener('click', reapplyPreprocessing);

        if(elements.listViewBtn) elements.listViewBtn.addEventListener('click', () => setViewMode('list'));
        if(elements.gridViewBtn) elements.gridViewBtn.addEventListener('click', () => setViewMode('grid'));
        if(elements.gridScaleSlider) elements.gridScaleSlider.addEventListener('input', updateGridScale);

        if(elements.mainColumnOperationsBtn) elements.mainColumnOperationsBtn.addEventListener('click', openMainColumnModal);
        if(elements.mainEditDataBtn) elements.mainEditDataBtn.addEventListener('click', () => toggleMainEditMode(true));
        if(elements.mainSaveEditsBtn) elements.mainSaveEditsBtn.addEventListener('click', saveMainEdits);
        if(elements.mainCopyRowsBtn) elements.mainCopyRowsBtn.addEventListener('click', copyMainSelectedRows);
        if(elements.mainExecuteFilterBtn) elements.mainExecuteFilterBtn.addEventListener('click', executeMainComplexFilter);

        if(elements.toggleMainConditionBtn) {
            elements.toggleMainConditionBtn.addEventListener('click', () => {
                const isHidden = elements.mainCondition2Wrapper.classList.toggle('hidden');
                elements.toggleMainConditionBtn.textContent = isHidden ? '+ 新增第二條件' : '- 移除第二條件';
                if (isHidden) {
                    if (elements.mainColSelect2) elements.mainColSelect2.value = '';
                    if (elements.mainInputCriteria2) {
                        elements.mainInputCriteria2.value = '';
                        elements.mainInputCriteria2.disabled = true;
                    }
                    const emptyRadio = document.querySelector('input[name="main-criteria-2"][value="empty"]');
                    if (emptyRadio) emptyRadio.checked = true;
                }
            });
        }
        
        if(elements.toggleMergeConditionBtn) {
            elements.toggleMergeConditionBtn.addEventListener('click', () => {
                const isHidden = elements.mergeCondition2Wrapper.classList.toggle('hidden');
                elements.toggleMergeConditionBtn.textContent = isHidden ? '+ 新增第二條件' : '- 移除第二條件';
                if (isHidden) {
                    if (elements.colSelect2) elements.colSelect2.value = '';
                    if (elements.inputCriteria2) {
                        elements.inputCriteria2.value = '';
                        elements.inputCriteria2.disabled = true;
                    }
                    const emptyRadio = document.querySelector('input[name="criteria-2"][value="empty"]');
                    if (emptyRadio) emptyRadio.checked = true;
                }
            });
        }

        if(elements.selectAllTablesBtn) elements.selectAllTablesBtn.addEventListener('click', () => { if(elements.displayArea) { elements.displayArea.querySelectorAll('.table-select-checkbox').forEach(cb => cb.checked = true); updateSelectionInfo(); } });
        if(elements.unselectAllTablesBtn) elements.unselectAllTablesBtn.addEventListener('click', () => { if(elements.displayArea) { elements.displayArea.querySelectorAll('.table-select-checkbox').forEach(cb => cb.checked = false); updateSelectionInfo(); } });
        if(elements.deleteSelectedTablesBtn) elements.deleteSelectedTablesBtn.addEventListener('click', () => deleteSelectedTables());
        if(elements.sortByNameBtn) elements.sortByNameBtn.addEventListener('click', sortTablesByFundName);

        if(elements.selectAllBtn) elements.selectAllBtn.addEventListener('click', () => { selectAllRows(); syncCheckboxesInScope(); });
        if(elements.invertSelectionBtn) elements.invertSelectionBtn.addEventListener('click', () => { invertSelection(); syncCheckboxesInScope(); });
        if(elements.selectEmptyBtn) elements.selectEmptyBtn.addEventListener('click', () => { selectEmptyRows(); syncCheckboxesInScope(); });
        if(elements.selectByKeywordBtn) elements.selectByKeywordBtn.addEventListener('click', () => { selectByKeyword(); syncCheckboxesInScope(); });
        if(elements.deleteSelectedBtn) elements.deleteSelectedBtn.addEventListener('click', () => deleteSelectedRows());
        if(elements.searchInput) elements.searchInput.addEventListener('input', utils.debounce(filterTable, 300));

        if(elements.resetViewBtn) elements.resetViewBtn.addEventListener('click', resetView);
        if(elements.showHiddenBtn) elements.showHiddenBtn.addEventListener('click', showAllHiddenElements);
        if(elements.exportMergedXlsxBtn) elements.exportMergedXlsxBtn.addEventListener('click', exportMergedXlsx);

        if(elements.mergeViewBtn) elements.mergeViewBtn.addEventListener('click', () => createMergedView('all'));
        if(elements.viewCheckedCombinedBtn) elements.viewCheckedCombinedBtn.addEventListener('click', () => createMergedView('checked'));
        if(elements.closeMergeViewBtn) elements.closeMergeViewBtn.addEventListener('click', closeMergeView);

        if(elements.searchInputMerged) elements.searchInputMerged.addEventListener('input', utils.debounce(filterTable, 300));
        if(elements.executeFilterSelectionBtn) elements.executeFilterSelectionBtn.addEventListener('click', () => { executeCombinedSelection(); syncCheckboxesInScope(); });
        if(elements.invertSelectionMergedBtn) elements.invertSelectionMergedBtn.addEventListener('click', () => { invertSelection(); syncCheckboxesInScope(); });
        if(elements.unselectMergedRowsBtn) elements.unselectMergedRowsBtn.addEventListener('click', unselectAllMergedRows);
        if(elements.toggleToolbarBtn) elements.toggleToolbarBtn.addEventListener('click', toggleToolbar);

        if(elements.editDataBtn) elements.editDataBtn.addEventListener('click', () => toggleEditMode(true));
        if(elements.saveEditsBtn) elements.saveEditsBtn.addEventListener('click', saveEdits);
        if(elements.cancelEditsBtn) elements.cancelEditsBtn.addEventListener('click', () => toggleEditMode(false));
        if(elements.addNewRowBtn) elements.addNewRowBtn.addEventListener('click', addNewRow);
        if(elements.copySelectedRowsBtn) elements.copySelectedRowsBtn.addEventListener('click', copySelectedRows);
        if(elements.deleteMergedRowsBtn) elements.deleteMergedRowsBtn.addEventListener('click', () => deleteSelectedRows());
        if(elements.toggleSourceColBtn) elements.toggleSourceColBtn.addEventListener('click', toggleSourceColumn);
        if(elements.toggleTotalRowBtn) elements.toggleTotalRowBtn.addEventListener('click', () => { state.showTotalRow = !state.showTotalRow; renderMergedTable(); });
        if(elements.exportCurrentMergedXlsxBtn) elements.exportCurrentMergedXlsxBtn.addEventListener('click', exportCurrentMergedXlsx);
        if(elements.sortMergedByNameBtn) elements.sortMergedByNameBtn.addEventListener('click', sortMergedTableByFundName);

        if(elements.columnOperationsBtn) elements.columnOperationsBtn.addEventListener('click', openMergeColumnModal);
        
        if(elements.closeColumnModalBtn) elements.closeColumnModalBtn.addEventListener('click', () => toggleColumnModal(false));
        if(elements.applyColumnChangesBtn) elements.applyColumnChangesBtn.addEventListener('click', () => { applyColumnChanges(); toggleColumnModal(false); });
        
        if(elements.modalCheckAll) elements.modalCheckAll.addEventListener('click', () => { if(elements.columnChecklist) elements.columnChecklist.querySelectorAll('input').forEach(i => i.checked = true); });
        if(elements.modalUncheckAll) elements.modalUncheckAll.addEventListener('click', () => { if(elements.columnChecklist) elements.columnChecklist.querySelectorAll('input').forEach(i => i.checked = false); });

        const setupDedupModal = (context) => {
            state.dedupContext = context;
            if(!elements.dedupColSelect) return;
            const allHeaders = new Set();
            if (context === 'main') {
                elements.displayArea.querySelectorAll('thead th:not(.checkbox-cell)').forEach(th => allHeaders.add(th.textContent.trim().replace(/\n/g, ' ')));
            } else {
                state.mergedHeaders.forEach(h => allHeaders.add(h.replace(/\n/g, ' ')));
            }
            elements.dedupColSelect.innerHTML = Array.from(allHeaders).map(h => `<option value="${h}">${h}</option>`).join('');
            const defaultMatch = Array.from(allHeaders).find(h => h.includes('名') || h.includes('基金') || h.includes('科目'));
            if (defaultMatch) elements.dedupColSelect.value = defaultMatch;
            if(elements.dedupModal) elements.dedupModal.classList.remove('hidden'); 
        };

        if(elements.mainSmartDedupBtn) elements.mainSmartDedupBtn.addEventListener('click', () => setupDedupModal('main'));
        if(elements.smartDedupBtn) elements.smartDedupBtn.addEventListener('click', () => setupDedupModal('merge'));

        if(elements.closeDedupModalBtn) elements.closeDedupModalBtn.addEventListener('click', () => { if(elements.dedupModal) elements.dedupModal.classList.add('hidden'); });
        if(elements.cancelDedupBtn) elements.cancelDedupBtn.addEventListener('click', () => { if(elements.dedupModal) elements.dedupModal.classList.add('hidden'); });
        if(elements.executeDedupBtn) elements.executeDedupBtn.addEventListener('click', executeSmartDeduplication);
        if(elements.clearDedupMarksBtn) elements.clearDedupMarksBtn.addEventListener('click', clearDedupMarks);
        if(elements.deleteDedupMarksBtn) elements.deleteDedupMarksBtn.addEventListener('click', () => deleteSelectedRows());

        if(elements.mergeViewContent) {
            elements.mergeViewContent.addEventListener('click', e => {
                const th = e.target.closest('th:not(.checkbox-cell)');
                const delBtn = e.target.closest('.delete-col-btn');
                if (delBtn && th) { e.stopPropagation(); deleteColumn(delBtn.dataset.header); }
                else if (th) { handleMergedHeaderClick(th); }
            });
        }

        if(elements.mergeViewModal) {
            elements.mergeViewModal.addEventListener('change', e => {
                if (e.target.name === 'criteria-1' || e.target.name === 'criteria-2') handleCriteriaChange(e);
            });
        }

        if(elements.controlPanel) {
            elements.controlPanel.addEventListener('change', e => {
                if (e.target.name === 'main-criteria-1' || e.target.name === 'main-criteria-2') {
                    const group = e.target.closest('.radio-group');
                    if(!group) return;
                    const targetInput = document.getElementById(group.dataset.target);
                    if(!targetInput) return;
                    const needsInput = e.target.value === 'exact' || e.target.value === 'includes';
                    targetInput.disabled = !needsInput;
                    if (needsInput) targetInput.focus(); else targetInput.value = '';
                }
            });
        }

        if(elements.displayArea) {
            elements.displayArea.addEventListener('change', e => {
                if (e.target.matches('.table-select-checkbox,[id^="select-all-cb"], .row-checkbox')) syncCheckboxesInScope();
            });
            elements.displayArea.addEventListener('click', e => {
                const card = e.target.closest('.table-wrapper');
                if (!card) return;
                if (e.target.classList.contains('close-zoom')) { closePreview(); return; }
                if (e.target.classList.contains('delete-rows-btn')) { deleteSelectedRows(card); return; }
                if (e.target.classList.contains('delete-table-btn')) { deleteSelectedTables(card); return; }
                const isGridView = elements.displayArea.classList.contains('grid-view');
                if (isGridView && !card.classList.contains('is-zoomed') && !e.target.matches('input, a, button, .btn')) { openPreview(card); }
            });
        }

        const onKeywordEnter = e => {
            if (e.key !== 'Enter') return;
            e.preventDefault();
            if(state.isMergedView) elements.executeFilterSelectionBtn?.click();
            else if(e.target.id.includes('main-input-criteria')) elements.mainExecuteFilterBtn?.click();
            else elements.selectByKeywordBtn?.click();
        };
        if(elements.selectKeywordInput) elements.selectKeywordInput.addEventListener('keydown', onKeywordEnter);
        if(elements.selectKeywordInputMerged) elements.selectKeywordInputMerged.addEventListener('keydown', onKeywordEnter);
        if(elements.mainInputCriteria1) elements.mainInputCriteria1.addEventListener('keydown', onKeywordEnter);
        if(elements.mainInputCriteria2) elements.mainInputCriteria2.addEventListener('keydown', onKeywordEnter);

        if(elements.backToTopBtn) elements.backToTopBtn.addEventListener('click', scrollToTop);
        window.addEventListener('scroll', handleScroll);

        document.addEventListener('keydown', e => {
            if ((e.ctrlKey || e.metaKey) && e.key === 'z') { e.preventDefault(); undoManager.undoLast(); }
            if (e.key === 'Escape') {
                if (elements.columnModal && !elements.columnModal.classList.contains('hidden')) { toggleColumnModal(false); }
                else if (elements.dedupModal && !elements.dedupModal.classList.contains('hidden')) { elements.dedupModal.classList.add('hidden'); }
                else if (state.isMergedView) { closeMergeView(); }
                else if (state.zoomedCard) { closePreview(); }
            }
        });
        if(elements.undoBtn) elements.undoBtn.addEventListener('click', () => undoManager.undoLast());

        // --- Step 2: 基金對照 & 彙整 事件 ---
        if(elements.executeMatchBtn) elements.executeMatchBtn.addEventListener('click', handleAutoMatch);
        if(elements.exportMatchedBtn) elements.exportMatchedBtn.addEventListener('click', handleExportMatched);

        if(elements.mapCheckAll) elements.mapCheckAll.addEventListener('change', e => {
            if(elements.mappingTbody) elements.mappingTbody.querySelectorAll('.map-row-check').forEach(cb => cb.checked = e.target.checked);
        });
        if(elements.mapSelectAllBtn) elements.mapSelectAllBtn.addEventListener('click', () => {
            if(elements.mappingTbody) elements.mappingTbody.querySelectorAll('.map-row-check').forEach(cb => cb.checked = true);
        });
        if(elements.mapDeselectAllBtn) elements.mapDeselectAllBtn.addEventListener('click', () => {
            if(elements.mappingTbody) elements.mappingTbody.querySelectorAll('.map-row-check').forEach(cb => cb.checked = false);
        });
        if(elements.mapInvertBtn) elements.mapInvertBtn.addEventListener('click', () => {
            if(elements.mappingTbody) elements.mappingTbody.querySelectorAll('.map-row-check').forEach(cb => cb.checked = !cb.checked);
        });

        // 彙整面板事件
        if(elements.aggCheckReviewBtn) elements.aggCheckReviewBtn.addEventListener('click', () => {
            state.aggregatedRows.forEach(r => { if (r.needsReview) r.checked = true; });
            renderAggregationPanel();
        });
        if(elements.aggUncheckAllBtn) elements.aggUncheckAllBtn.addEventListener('click', () => {
            state.aggregatedRows.forEach(r => r.checked = false);
            renderAggregationPanel();
        });
        if(elements.aggDeleteCheckedBtn) elements.aggDeleteCheckedBtn.addEventListener('click', deleteCheckedAggregatedRows);
        if(elements.aggExportBtn) elements.aggExportBtn.addEventListener('click', exportAggregatedRows);

        if(elements.matchColSelect) elements.matchColSelect.addEventListener('change', () => {
            state.matchColumn = elements.matchColSelect.value;
            const colLabel = state.matchColumn === '1' ? '第一欄' : '第二欄';
            document.querySelectorAll('#match-col-label, #match-col-label2').forEach(el => el.textContent = colLabel);
            rebuildAggregationWithNewMatchColumn();
        });

        if(elements.aggregatePanel) {
            elements.aggregatePanel.addEventListener('change', e => {
                if (e.target.classList.contains('agg-row-check')) {
                    const tr = e.target.closest('tr');
                    const idx = parseInt(tr.dataset.aggIndex, 10);
                    state.aggregatedRows[idx].checked = e.target.checked;
                    tr.classList.toggle('row-checked-delete', e.target.checked);
                    updateAggregationCounts();
                }
                if (e.target.id === 'agg-check-all-cb') {
                    const isChecked = e.target.checked;
                    state.aggregatedRows.forEach(r => r.checked = isChecked);
                    renderAggregationPanel();
                }
            });
        }

        // 進階設定折疊
        if(elements.toggleImportSettings) {
            elements.toggleImportSettings.addEventListener('click', () => {
                if(elements.importSettingsPanel) elements.importSettingsPanel.classList.toggle('collapsed');
            });
        }

        // 清除所有
        if(elements.clearAllBtn) elements.clearAllBtn.addEventListener('click', () => clearAllFiles(false));
    }

    // --- 全域 handler ---
    function handleAutoMatch() {
        console.log('🔍 handleAutoMatch called');
        console.log('  fileInfos:', state.fileInfos.length);
        console.log('  fundSortOrder:', state.fundSortOrder.length);
        console.log('  mappingTbody:', !!elements.mappingTbody);
        if (elements.mappingTbody) {
            const checked = elements.mappingTbody.querySelectorAll('.map-row-check:checked');
            console.log('  checked files:', checked.length);
        }
        buildAggregation();
        console.log('  aggregatedRows:', state.aggregatedRows.length);
        
        // 滾動到彙整資料預覽
        if (elements.aggregatePanel && !elements.aggregatePanel.classList.contains('hidden')) {
            elements.aggregatePanel.scrollIntoView({ behavior: 'smooth', block: 'start' });
        }
    }

    function handleExportMatched() {
        if (state.aggregatedRows.length > 0) {
            exportAggregatedRows();
        } else {
            // 如果還沒有彙整資料，先執行彙整再匯出
            buildMappingTable();
            buildAggregation();
            if (state.aggregatedRows.length > 0) exportAggregatedRows();
            else alert('沒有找到可匯出的資料。');
        }
    }

    async function init() {
        try {
            cacheElements();
            await loadFundConfig();
            bindEvents();
            console.log("✅ ExcelViewer 初始化成功！");
        } catch (error) {
            console.error("❌ 初始化過程中發生錯誤：", error);
        }
    }

    return { init };
})();

if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', ExcelViewer.init);
} else {
    ExcelViewer.init();
}
