/**
 * ExcelViewer — 終極效能黃金版 (結合舊版真實邊界偵測，徹底防堵 160 億格假死)
 */

const ExcelViewer = (() => {
    'use strict';

    const CONSTANTS = { VALID_FILE_EXTENSIONS: ['.xls', '.xlsx'] };

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
            undoToast: 'undo-toast', undoText: 'undo-text', undoBtn: 'undo-btn'
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

    // 🚀 【完美清洗引擎】：確保拿到的是已切斷幽靈的真實矩陣
    function applyPreprocessing(jsonData, sheet) {
        const useHeaderCompression = elements.skipTopRowsCheckbox && elements.skipTopRowsCheckbox.checked;
        const discardCount = useHeaderCompression && elements.discardRowsInput ? parseInt(elements.discardRowsInput.value, 10) || 0 : 0;
        const headerCount = useHeaderCompression && elements.headerRowsInput ? parseInt(elements.headerRowsInput.value, 10) || 1 : 1;
        const totalSkipCount = useHeaderCompression ? (discardCount + headerCount) : 0;

        const removeEmpty = elements.removeEmptyRowsCheckbox ? elements.removeEmptyRowsCheckbox.checked : false;
        const removeKeywords = elements.removeKeywordRowsCheckbox ? elements.removeKeywordRowsCheckbox.checked : false;
        const keywords = removeKeywords && elements.removeKeywordRowsInput ? elements.removeKeywordRowsInput.value.split(',').map(k => k.trim().toLowerCase()).filter(Boolean) : [];

        const colProps = sheet['!cols'] || [];
        const maxCol = jsonData[0] ? jsonData[0].length - 1 : 0;
        
        // 抓出所有沒被隱藏的欄位索引
        const visibleColIndices = [];
        for (let c = 0; c <= maxCol; c++) {
            if (!(colProps[c] && colProps[c].hidden)) visibleColIndices.push(c);
        }

        const result = [];
        const usedNames = new Set(); 

        // 1. 產生換行表頭
        if (useHeaderCompression && jsonData.length > discardCount) {
            const headerRow = visibleColIndices.map((colIdx, index) => { 
                const headerParts = [];
                const scanEnd = Math.min(discardCount + headerCount, jsonData.length);
                for (let r = discardCount; r < scanEnd; r++) {
                    if (jsonData[r] && jsonData[r][colIdx].trim() !== '') {
                        headerParts.push(jsonData[r][colIdx].trim().replace(/\r?\n|\r/g, ''));
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

        // 2. 處理資料列
        for (let r = totalSkipCount; r < jsonData.length; r++) {
            const row = jsonData[r];
            if (!row) continue;

            const visibleRow = visibleColIndices.map(c => row[c]);
            const isEmpty = visibleRow.every(val => val.trim() === '');
            
            if (removeEmpty && isEmpty) continue; 

            if (keywords.length > 0 && !isEmpty) {
                const content = visibleRow.join(' ').toLowerCase();
                if (keywords.some(k => content.includes(k))) continue; 
            }
            
            if (visibleRow.some(val => val.trim() !== '')) {
                result.push(visibleRow);
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
        if(elements.displayArea) elements.displayArea.innerHTML = '<div class="loading">正在掃描真實邊界...</div>';
        resetControls(true);
        state.rawSheetsCache = []; state.loadedFiles = [];
        const tablesToRender = [];

        try {
            for (let i = 0; i < files.length; i++) {
                const file = files[i];
                if(elements.displayArea) elements.displayArea.innerHTML = `<div class="loading">極速讀取中... (${i + 1}/${files.length})</div>`;
                await new Promise(r => setTimeout(r, 10)); // 讓 UI 有時間轉圈圈

                const binary = await utils.readFileAsBinary(file);
                const workbook = XLSX.read(binary, { type: 'binary', cellStyles: true });
                const sheetNames = await getSelectedSheetNames(file.name, workbook, importMode, sheetCriteria);

                for (const sheetName of sheetNames) {
                    const sheet = workbook.Sheets[sheetName];
                    
                    // 🚀🚀🚀 【救命防線】：使用 defval: null 產生稀疏陣列，完美避開 104 萬列崩潰！
                    let rawData = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: null });
                    if (!rawData || rawData.length === 0) continue;

                    // 【邊界掃描 1】: 由下往上掃，找到最後一列真正有字的
                    let lastDataRowIdx = -1;
                    for (let r = rawData.length - 1; r >= 0; r--) {
                        if (rawData[r] && rawData[r].some(c => c !== null && c !== undefined && String(c).trim() !== '')) {
                            lastDataRowIdx = r;
                            break;
                        }
                    }
                    if (lastDataRowIdx === -1) continue; // 整張表全空

                    // 【邊界掃描 2】: 找出真實的最右側欄位
                    let maxCol = -1;
                    for (let r = 0; r <= lastDataRowIdx; r++) {
                        if (rawData[r]) {
                            for(let c = rawData[r].length - 1; c >= 0; c--) {
                                if (rawData[r][c] !== null && rawData[r][c] !== undefined && String(rawData[r][c]).trim() !== '') {
                                    if (c > maxCol) maxCol = c;
                                    break;
                                }
                            }
                        }
                    }

                    // 【建立乾淨矩陣】: 把稀疏陣列轉成完美切割、零幽靈的乾淨陣列
                    let jsonData = new Array(lastDataRowIdx + 1);
                    for(let r = 0; r <= lastDataRowIdx; r++) {
                        jsonData[r] = new Array(maxCol + 1).fill('');
                        if (rawData[r]) {
                            for (let c = 0; c <= maxCol; c++) {
                                if (rawData[r][c] !== null && rawData[r][c] !== undefined) {
                                    jsonData[r][c] = String(rawData[r][c]);
                                }
                            }
                        }
                    }

                    // 安全地填滿合併儲存格 (絕對不會超過真實資料區)
                    if (sheet['!merges']) {
                        sheet['!merges'].forEach(merge => {
                            const sr = merge.s.r, sc = merge.s.c;
                            if (sr <= lastDataRowIdx && sc <= maxCol && jsonData[sr]) {
                                const val = jsonData[sr][sc];
                                if (val.trim() === '') return;
                                
                                const er = Math.min(merge.e.r, lastDataRowIdx);
                                const ec = Math.min(merge.e.c, maxCol);
                                
                                for (let r = sr; r <= er; r++) {
                                    for (let c = sc; c <= ec; c++) {
                                        jsonData[r][c] = val;
                                    }
                                }
                            }
                        });
                    }

                    const label = `${file.name} (${sheetName})`;
                    // 保存這份小巧乾淨的 jsonData 以供重新套用
                    state.rawSheetsCache.push({ label, sheet, jsonData: JSON.parse(JSON.stringify(jsonData)) });

                    const filtered = applyPreprocessing(jsonData, sheet);
                    if (filtered.length > 0) {
                        const cleanSheet = XLSX.utils.aoa_to_sheet(filtered);
                        tablesToRender.push({ html: XLSX.utils.sheet_to_html(cleanSheet), filename: label });
                        state.loadedFiles.push(label);
                    }
                }
            }
            state.loadedTables = tablesToRender.length;
            renderTables(tablesToRender);
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
                // 從快取拿出來的是乾淨小矩陣
                const clonedData = JSON.parse(JSON.stringify(cache.jsonData));
                const filtered = applyPreprocessing(clonedData, cache.sheet);
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
        if(elements.dropAreaInitial) elements.dropAreaInitial.classList.toggle('hidden', hasFiles);
        if(elements.dropAreaLoaded) elements.dropAreaLoaded.classList.toggle('hidden', !hasFiles);
        if(elements.importOptionsContainer) elements.importOptionsContainer.classList.toggle('hidden', hasFiles);
        if (hasFiles && elements.fileCount && elements.fileNames) {
            elements.fileCount.textContent = state.loadedTables;
            elements.fileNames.textContent = state.loadedFiles.slice(0, 3).join(', ') + (state.loadedFiles.length > 3 ? '...' : '');
        }
    }

    function showControls(hiddenCount) { 
        if(elements.controlPanel) elements.controlPanel.classList.remove('hidden'); 
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
        if(elements.displayArea) elements.displayArea.innerHTML = ''; 
        if (elements.fileInput) elements.fileInput.value = '';
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

        if(elements.importOptionsContainer) {
            elements.importOptionsContainer.addEventListener('change', e => {
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
    }

    async function init() {
        try {
            cacheElements();
            await loadFundConfig();
            bindEvents();
            console.log("✅ ExcelViewer 初始化成功！終極效能黃金版啟動。");
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
