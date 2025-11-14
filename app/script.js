// --- 移除了頂部的 FUND_ORDER_LIST ---

const ExcelViewer = (() => {
    'use strict';
    const CONSTANTS = { VALID_FILE_EXTENSIONS: ['.xls', '.xlsx'] };
    const state = { 
        originalHtmlString: '', 
        isProcessing: false, 
        loadedFiles: [], 
        loadedTables: 0, 
        zoomedCard: null,
        isMergedView: false,
        isEditing: false,
        showTotalRow: false,
        showSourceColumn: false, 
        mergedData: [],
        mergedHeaders: [],
        fundSortOrder: [], 
        fundAliasMap: {},   
        fundAliasKeys: []  
    };
    const elements = {};

    async function init() {
        cacheElements();
        await loadFundConfig(); 
        bindEvents();
    }

    async function loadFundConfig() {
        try {
            const response = await fetch(`fund-config.json?v=${Date.now()}`); 
            if (!response.ok) {
                throw new Error(`HTTP error! status: ${response.status}`);
            }
            const config = await response.json();
            
            if (config.sortOrder && config.aliasMap) {
                state.fundSortOrder = config.sortOrder;
                state.fundAliasMap = config.aliasMap;
                state.fundAliasKeys = Object.keys(config.aliasMap).sort((a, b) => b.length - a.length);
                console.log("基金設定檔 (fund-config.json) 載入成功。");
            } else {
                console.error("基金設定檔格式錯誤：缺少 sortOrder 或 aliasMap。");
                alert("錯誤：基金設定檔 (fund-config.json) 格式不正確。");
            }
        } catch (err) {
            console.error("載入基金設定檔 (fund-config.json) 失敗:", err);
            alert("警告：無法載入基金排序設定檔 (fund-config.json)。\n「依基金名稱排序」功能將無法使用。");
        }
    }

    function cacheElements() {
        const mapping = {
            fileInput: 'file-input', displayArea: 'excel-display-area', searchInput: 'search-input',
            dropArea: 'drop-area', deleteSelectedBtn: 'delete-selected-btn', invertSelectionBtn: 'invert-selection-btn',
            resetViewBtn: 'reset-view-btn', selectEmptyBtn: 'select-empty-btn',
            showHiddenBtn: 'show-hidden-btn',
            exportMergedXlsxBtn: 'export-merged-xlsx-btn',
            selectByKeywordGroup: 'select-by-keyword-group', selectKeywordInput: 'select-keyword-input',
            selectByKeywordBtn: 'select-by-keyword-btn', selectKeywordRegex: 'select-keyword-regex',
            loadStatusMessage: 'load-status-message', controlPanel: 'control-panel',
            dropAreaInitial: 'drop-area-initial', dropAreaLoaded: 'drop-area-loaded',
            fileCount: 'file-count', fileNames: 'file-names', clearFilesBtn: 'clear-files-btn',
            selectAllBtn: 'select-all-btn',
            importOptionsContainer: 'import-options-container', specificSheetNameGroup: 'specific-sheet-name-group',
            specificSheetNameInput: 'specific-sheet-name-input', specificSheetPositionGroup: 'specific-sheet-position-group',
            specificSheetPositionInput: 'specific-sheet-position-input', selectAllTablesBtn: 'select-all-tables-btn',
            unselectAllTablesBtn: 'unselect-all-tables-btn', deleteSelectedTablesBtn: 'delete-selected-tables-btn',
            sortByNameBtn: 'sort-by-fund-name-btn', 
            tableLevelControls: 'table-level-controls', selectedTablesInfo: 'selected-tables-info',
            selectedTablesList: 'selected-tables-list', listViewBtn: 'list-view-btn',
            gridViewBtn: 'grid-view-btn', backToTopBtn: 'back-to-top-btn',
            gridScaleControl: 'grid-scale-control', gridScaleSlider: 'grid-scale-slider',
            
            mergeViewModal: 'merge-view-modal',
            closeMergeViewBtn: 'close-merge-view-btn',
            mergeViewContent: 'merge-view-content',
            mergeViewBtn: 'merge-view-btn',
            viewCheckedCombinedBtn: 'view-checked-combined-btn', 
            
            columnOperationsBtn: 'column-operations-btn',
            columnModal: 'column-modal',
            closeColumnModalBtn: 'close-column-modal-btn',
            columnChecklist: 'column-checklist',
            applyColumnChangesBtn: 'apply-column-changes-btn',
            modalCheckAll: 'modal-check-all',
            modalUncheckAll: 'modal-uncheck-all',
            
            editDataBtn: 'edit-data-btn',
            saveEditsBtn: 'save-edits-btn',
            cancelEditsBtn: 'cancel-edits-btn',
            addNewRowBtn: 'add-new-row-btn',
            copySelectedRowsBtn: 'copy-selected-rows-btn',
            deleteMergedRowsBtn: 'delete-merged-rows-btn',
            toggleTotalRowBtn: 'toggle-total-row-btn', 
            toggleSourceColBtn: 'toggle-source-col-btn', 
            invertSelectionMergedBtn: 'invert-selection-merged-btn',
            exportSelectedMergedXlsxBtn: 'export-selected-merged-xlsx-btn', 
            exportCurrentMergedXlsxBtn: 'export-current-merged-xlsx-btn', 
            sortMergedByNameBtn: 'sort-merged-by-fund-name-btn',
            
            colSelect1: 'col-select-1',
            colSelect2: 'col-select-2',
            inputCriteria1: 'input-criteria-1',
            inputCriteria2: 'input-criteria-2',

            searchInputMerged: 'search-input-merged',
            selectKeywordInputMerged: 'select-keyword-input-merged',
            selectKeywordRegexMerged: 'select-keyword-regex-merged',
            
            // --- MODIFIED: Combined Button ---
            executeFilterSelectionBtn: 'execute-filter-selection-btn',

            toggleToolbarBtn: 'toggle-toolbar-btn',
            collapsibleToolbar: 'collapsible-toolbar-area'
        };
        Object.keys(mapping).forEach(key => {
            elements[key] = document.getElementById(mapping[key]);
        });
    }
    
    function handleCriteriaChange(e) {
        const radio = e.target;
        if (radio.type !== 'radio') return;
        const group = radio.closest('.radio-group');
        if (!group) return;
        const targetInputId = group.dataset.target;
        const targetInput = elements[targetInputId];
        if (!targetInput) return;
        const newValue = radio.value;
        if (newValue === 'exact' || newValue === 'includes') {
            targetInput.disabled = false;
            targetInput.focus();
        } else {
            targetInput.disabled = true;
            targetInput.value = '';
        }
    }
    
    function bindEvents() {
        elements.fileInput.addEventListener('change', e => processFiles(e.target.files));
        setupDragAndDrop();
        elements.clearFilesBtn.addEventListener('click', () => clearAllFiles(false));

        elements.listViewBtn.addEventListener('click', () => setViewMode('list'));
        elements.gridViewBtn.addEventListener('click', () => setViewMode('grid'));
        elements.gridScaleSlider.addEventListener('input', updateGridScale);
        elements.selectAllTablesBtn.addEventListener('click', () => { selectAllTables(true); updateSelectionInfo(); });
        elements.unselectAllTablesBtn.addEventListener('click', () => { selectAllTables(false); updateSelectionInfo(); });
        elements.deleteSelectedTablesBtn.addEventListener('click', deleteSelectedTables);
        elements.sortByNameBtn.addEventListener('click', sortTablesByFundName); 
        
        elements.selectByKeywordBtn.addEventListener('click', () => { selectByKeyword(); syncCheckboxesInScope(); });
        elements.selectEmptyBtn.addEventListener('click', () => { selectEmptyRows(); syncCheckboxesInScope(); });
        elements.selectAllBtn.addEventListener('click', () => { selectAllRows(); syncCheckboxesInScope(); });
        elements.invertSelectionBtn.addEventListener('click', () => { invertSelection(); syncCheckboxesInScope(); });
        elements.deleteSelectedBtn.addEventListener('click', deleteSelectedRows);
        
        // --- MODIFIED: Event for combined button ---
        if (elements.executeFilterSelectionBtn) {
            elements.executeFilterSelectionBtn.addEventListener('click', () => {
                executeCombinedSelection(); 
                syncCheckboxesInScope();
            });
        }
        
        elements.resetViewBtn.addEventListener('click', resetView);
        elements.showHiddenBtn.addEventListener('click', showAllHiddenElements);
        elements.exportMergedXlsxBtn.addEventListener('click', exportMergedXlsx);
        
        elements.mergeViewBtn.addEventListener('click', () => createMergedView('all')); 
        elements.viewCheckedCombinedBtn.addEventListener('click', () => createMergedView('checked'));
        elements.closeMergeViewBtn.addEventListener('click', closeMergeView);
        elements.columnOperationsBtn.addEventListener('click', () => toggleColumnModal(true));
        elements.closeColumnModalBtn.addEventListener('click', () => toggleColumnModal(false));
        elements.applyColumnChangesBtn.addEventListener('click', () => { applyColumnChanges(); toggleColumnModal(false); });
        elements.modalCheckAll.addEventListener('click', () => setAllColumnCheckboxes(true));
        elements.modalUncheckAll.addEventListener('click', () => setAllColumnCheckboxes(false));
        elements.editDataBtn.addEventListener('click', () => toggleEditMode(true));
        elements.saveEditsBtn.addEventListener('click', saveEdits);
        elements.cancelEditsBtn.addEventListener('click', () => toggleEditMode(false));
        elements.addNewRowBtn.addEventListener('click', addNewRow);
        elements.copySelectedRowsBtn.addEventListener('click', copySelectedRows);
        elements.deleteMergedRowsBtn.addEventListener('click', deleteSelectedRows);
        elements.toggleTotalRowBtn.addEventListener('click', () => {
            state.showTotalRow = !state.showTotalRow;
            renderMergedTable();
        });
        
        elements.toggleSourceColBtn.addEventListener('click', toggleSourceColumn); 
        elements.invertSelectionMergedBtn.addEventListener('click', () => { invertSelection(); syncCheckboxesInScope(); });
        elements.exportSelectedMergedXlsxBtn.addEventListener('click', exportSelectedMergedXlsx); 
        elements.exportCurrentMergedXlsxBtn.addEventListener('click', exportCurrentMergedXlsx); 
        elements.sortMergedByNameBtn.addEventListener('click', sortMergedTableByFundName); 

        elements.searchInput.addEventListener('input', debounce(filterTable, 300));
        elements.displayArea.addEventListener('change', handleDisplayAreaChange);
        elements.displayArea.addEventListener('click', handleCardClick);
        
        elements.mergeViewContent.addEventListener('click', e => {
            const th = e.target.closest('th:not(.checkbox-cell)');
            const delBtn = e.target.closest('.delete-col-btn');
            if (delBtn && th) {
                e.stopPropagation();
                deleteColumn(delBtn.dataset.header);
            } else if (th) {
                handleMergedHeaderClick(th);
            }
        });
        
        elements.mergeViewModal.addEventListener('change', e => {
            if (e.target.name === 'criteria-1' || e.target.name === 'criteria-2') {
                handleCriteriaChange(e);
            }
        });
        
        elements.importOptionsContainer.addEventListener('change', e => {
            if (e.target.name === 'import-mode') {
                const selectedMode = e.target.value;
                elements.specificSheetNameGroup.classList.toggle('hidden', selectedMode !== 'specific');
                elements.specificSheetPositionGroup.classList.toggle('hidden', selectedMode !== 'position');
            }
        });

        elements.searchInputMerged.addEventListener('input', debounce(filterTable, 300));

        const handleKeywordEnter = (e) => {
            if (e.key === 'Enter') {
                e.preventDefault();
                if (state.isMergedView) {
                    elements.executeFilterSelectionBtn.click(); // Trigger combined button
                } else {
                    elements.selectByKeywordBtn.click();
                }
            }
        };
        elements.selectKeywordInput.addEventListener('keydown', handleKeywordEnter);
        elements.selectKeywordInputMerged.addEventListener('keydown', handleKeywordEnter);

        elements.toggleToolbarBtn.addEventListener('click', toggleToolbar);

        elements.backToTopBtn.addEventListener('click', scrollToTop);
        window.addEventListener('scroll', handleScroll);
        document.addEventListener('keydown', e => { 
            if (e.key === 'Escape') { 
                if (!elements.columnModal.classList.contains('hidden')) {
                    toggleColumnModal(false);
                } else if (state.isMergedView) {
                    closeMergeView();
                } else if (state.zoomedCard) {
                    closePreview();
                }
            } 
        });
        
        elements.mergeViewModal.dispatchEvent(new Event('change', { bubbles: true }));
        elements.mergeViewBtn.addEventListener('click', () => {
            setTimeout(() => {
                elements.mergeViewModal.dispatchEvent(new Event('change', { bubbles: true }));
            }, 50);
        });
    }

    // --- [NEW] Combined selection function ---
    function executeCombinedSelection() {
        if (!state.isMergedView) return;

        // --- Part 1: Get Keyword Filter Settings ---
        const keywordInput = elements.selectKeywordInputMerged.value.trim();
        const isRegex = elements.selectKeywordRegexMerged.checked;
        let keywordMatchLogic = () => false; // Default to no match
        let hasKeywordCriteria = false;

        if (keywordInput) {
            hasKeywordCriteria = true;
            try {
                if (isRegex) {
                    const regex = new RegExp(keywordInput, 'i');
                    keywordMatchLogic = text => regex.test(text);
                } else if (keywordInput.includes(',')) {
                    const keywords = keywordInput.split(',').map(k => k.trim().toLowerCase()).filter(Boolean);
                    keywordMatchLogic = text => keywords.some(k => text.toLowerCase().includes(k));
                } else {
                    const keywords = keywordInput.split(/\s+/).map(k => k.trim().toLowerCase()).filter(Boolean);
                    keywordMatchLogic = text => keywords.every(k => text.toLowerCase().includes(k));
                }
            } catch (e) {
                alert('無效的 Regex 表示式：\n' + e.message);
                return;
            }
        }

        // --- Part 2: Get Complex Filter Settings ---
        const col1 = elements.colSelect1.value;
        const col2 = elements.colSelect2.value;
        const criteria1 = document.querySelector('input[name="criteria-1"]:checked').value;
        const criteria2 = document.querySelector('input[name="criteria-2"]:checked').value;
        const logicOp = document.querySelector('input[name="logic-op"]:checked').value;
        const inputVal1 = elements.inputCriteria1.value;
        const inputVal2 = elements.inputCriteria2.value;
        const hasComplexCriteria = col1 || col2;

        if (!hasKeywordCriteria && !hasComplexCriteria) {
            alert('請至少輸入關鍵字或設定一個欄位篩選條件。');
            return;
        }

        const checkValue = (cellVal, criteria, inputVal) => {
            const strCellVal = String(cellVal).trim();
            const strInputVal = String(inputVal).trim();
            switch (criteria) {
                case 'empty': return strCellVal === '';
                case 'zero': return strCellVal === '0';
                case 'value': return strCellVal !== '';
                case 'exact': return strCellVal === strInputVal;
                case 'includes': return strInputVal !== '' && strCellVal.toLowerCase().includes(strInputVal.toLowerCase());
                default: return false;
            }
        };

        // --- Part 3: Iterate and Apply Logic ---
        let count = 0;
        elements.mergeViewContent.querySelectorAll('tbody tr:not(.row-hidden-search)').forEach(row => {
            const checkbox = row.querySelector('.row-checkbox');
            if (!checkbox) return;

            // Check keyword match
            let keywordMatch = false;
            if (hasKeywordCriteria) {
                const rowText = Array.from(row.querySelectorAll('td:not(.checkbox-cell)')).map(c => c.textContent).join(' ');
                keywordMatch = keywordMatchLogic(rowText);
            }

            // Check complex filter match
            let complexMatch = false;
            if (hasComplexCriteria) {
                let result1 = null, result2 = null;
                if (col1) {
                    const cell = row.querySelector(`td[data-col-header="${col1}"]`);
                    result1 = checkValue(cell ? cell.textContent : '', criteria1, inputVal1);
                }
                if (col2) {
                    const cell = row.querySelector(`td[data-col-header="${col2}"]`);
                    result2 = checkValue(cell ? cell.textContent : '', criteria2, inputVal2);
                }
                if (col1 && col2) {
                    complexMatch = (logicOp === 'and') ? (result1 && result2) : (result1 || result2);
                } else if (col1) {
                    complexMatch = result1;
                } else if (col2) {
                    complexMatch = result2;
                }
            }

            // Combine results (OR logic)
            if (keywordMatch || complexMatch) {
                checkbox.checked = true;
                count++;
            }
        });

        alert(count > 0 ? `已勾選 ${count} 筆符合條件的資料。` : '未找到符合條件的資料。');
    }

    // ... (The rest of the code from processFiles down to the end remains largely the same)
    // ... I will paste the full, correct code below for clarity and to avoid errors.
    
    // --- Core Logic (File Processing, Rendering) ---
    function setupDragAndDrop() {
        elements.dropArea.addEventListener('click', e => {
            if (e.target.id === 'clear-files-btn' || e.target.closest('.btn-clear') || e.target.id === 'file-input') {
                return;
            }
            elements.fileInput.click();
        });
        ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => { elements.dropArea.addEventListener(eventName, e => { e.preventDefault(); e.stopPropagation(); }); });
        ['dragenter', 'dragover'].forEach(eventName => { elements.dropArea.addEventListener(eventName, () => elements.dropArea.classList.add('highlight')); });
        ['dragleave', 'drop'].forEach(eventName => { elements.dropArea.addEventListener(eventName, () => elements.dropArea.classList.remove('highlight')); });
        elements.dropArea.addEventListener('drop', e => processFiles(e.dataTransfer.files));
    }
    
    async function processFiles(fileList) { 
        const validation = validateFiles(fileList); 
        if (!validation.valid) { alert(`錯誤：${validation.error}`); return; } 
        if (state.isProcessing) { alert('正在處理檔案...'); return; } 
        const importMode = document.querySelector('input[name="import-mode"]:checked').value; 
        const specificSheetName = elements.specificSheetNameInput.value.trim(); 
        const specificSheetPosition = elements.specificSheetPositionInput.value.trim(); 
        if (importMode === 'specific' && !specificSheetName) { alert('請輸入工作表名稱！'); return; } 
        if (importMode === 'position' && !specificSheetPosition) { alert('請輸入工作表位置！'); return; } 
        
        state.isProcessing = true; 
        elements.displayArea.innerHTML = '<div class="loading">讀取中...</div>'; 
        resetControls(true); 
        const tablesToRender = []; 
        const missedFiles = []; 
        state.loadedFiles = []; 
        
        try { 
            for (let index = 0; index < validation.files.length; index++) { 
                const file = validation.files[index]; 
                elements.displayArea.innerHTML = `<div class="loading">讀取中... (${index + 1}/${validation.files.length})</div>`; 
                const binaryData = await readFileAsBinary(file); 
                const workbook = XLSX.read(binaryData, { type: 'binary', cellStyles: true }); 
                const sheetNames = await getSelectedSheetNames(file.name, workbook, importMode, { name: specificSheetName, position: specificSheetPosition }); 
                
                if ((importMode === 'specific' || importMode === 'position') && sheetNames.length === 0 && workbook.SheetNames.length > 0) { 
                    missedFiles.push(file.name); 
                } 
                
                for (const sheetName of sheetNames) { 
                    const sheet = workbook.Sheets[sheetName];
                    let startRow = 0, startCol = 0, endCol = 0;
                    if (sheet['!ref']) {
                        const range = XLSX.utils.decode_range(sheet['!ref']);
                        startRow = range.s.r; startCol = range.s.c; endCol = range.e.c;
                    }
                    const jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '', range: sheet['!ref'], raw: false });
                    if (sheet['!merges']) {
                        sheet['!merges'].forEach(merge => {
                            const startR = merge.s.r - startRow, startC = merge.s.c - startCol, endR = merge.e.r - startRow, endC = merge.e.c - startCol;
                            if (startR >= 0 && startC >= 0 && jsonData[startR]) {
                                const primaryValue = jsonData[startR][startC];
                                for (let r = startR; r <= endR; r++) for (let c = startC; c <= endC; c++) if (jsonData[r]) jsonData[r][c] = primaryValue;
                            }
                        });
                    }
                    const rowProps = sheet['!rows'] || [], colProps = sheet['!cols'] || []; 
                    const visibleRelativeIndices = [];
                    for (let c = startCol; c <= endCol; c++) if (!(colProps[c] && colProps[c].hidden)) visibleRelativeIndices.push(c - startCol);
                    const filteredData = [];
                    jsonData.forEach((row, index) => {
                        const absoluteRowIndex = startRow + index;
                        if (rowProps[absoluteRowIndex] && rowProps[absoluteRowIndex].hidden) return; 
                        const safeRow = row || [];
                        const newRow = visibleRelativeIndices.map(i => (safeRow[i] !== undefined ? safeRow[i] : ''));
                        if (newRow.some(cell => String(cell).trim() !== '')) filteredData.push(newRow);
                    });
                    if (filteredData.length > 0) {
                        const cleanedSheet = XLSX.utils.aoa_to_sheet(filteredData);
                        const htmlString = XLSX.utils.sheet_to_html(cleanedSheet); 
                        tablesToRender.push({ html: htmlString, filename: `${file.name} (${sheetName})` }); 
                        state.loadedFiles.push(`${file.name} (${sheetName})`); 
                    }
                } 
            } 
            if (missedFiles.length > 0) { 
                const criteria = importMode === 'specific' ? `名稱包含 "${specificSheetName}"` : `位置符合 "${specificSheetPosition}"`; 
                alert(`以下檔案找不到 ${criteria} 的工作表：\n\n- ${missedFiles.join('\n- ')}`); 
            } 
            state.loadedTables = tablesToRender.length; 
            renderTables(tablesToRender); 
            updateDropAreaDisplay(); 
        } catch (err) { 
            console.error("處理檔案時發生錯誤:", err); 
            elements.displayArea.innerHTML = `<p style="color: red;">處理檔案錯誤：${err.message || '未知錯誤'}</p>`; 
            resetControls(true); 
        } finally { 
            state.isProcessing = false; 
        } 
    }
    
    function renderTables(tablesToRender) { 
        if (tablesToRender.length === 0) { 
            elements.displayArea.innerHTML = `<p>沒有找到符合條件的工作表。</p>`; 
            return; 
        } 
        const fragment = document.createDocumentFragment(); 
        tablesToRender.forEach(({ html, filename }) => { 
            const wrapper = document.createElement('div'); 
            wrapper.className = 'table-wrapper'; 
            const header = document.createElement('div'); 
            header.className = 'table-header'; 
            header.innerHTML = `<input type="checkbox" class="table-select-checkbox" title="選取此表格"><h4>${filename}</h4><div class="header-actions"><button class="btn btn-danger btn-sm delete-rows-btn">刪除選取列</button><button class="btn btn-danger btn-sm delete-table-btn">刪除此表</button></div><button class="close-zoom">&times;</button>`; 
            const tableContent = document.createElement('div'); 
            tableContent.className = 'table-content'; 
            const tempDiv = document.createElement('div'); 
            tempDiv.innerHTML = html; 
            const table = tempDiv.querySelector('table'); 
            if (table) { 
                tableContent.appendChild(table); 
                wrapper.appendChild(header); 
                wrapper.appendChild(tableContent); 
                fragment.appendChild(wrapper); 
            } 
        }); 
        elements.displayArea.innerHTML = ''; 
        elements.displayArea.appendChild(fragment); 
        state.originalHtmlString = elements.displayArea.innerHTML; 
        injectCheckboxes(elements.displayArea); 
        showControls(detectHiddenElements()); 
        sortTablesByFundName();
    }
    
    function injectCheckboxes(scope) { scope.querySelectorAll('thead tr').forEach((headRow, index) => { if(headRow.querySelector('.checkbox-cell')) return; const th = document.createElement('th'); th.innerHTML = `<input type="checkbox" id="select-all-checkbox-${scope.id}-${index}" title="全選/全不選">`; th.classList.add('checkbox-cell'); headRow.prepend(th); }); scope.querySelectorAll('tbody tr').forEach(row => { if(row.querySelector('.checkbox-cell')) return; const td = document.createElement('td'); td.innerHTML = '<input type="checkbox" class="row-checkbox">'; td.classList.add('checkbox-cell'); row.prepend(td); }); }
    
    function createMergedView(mode = 'all') {
        const allVisibleTables = Array.from(elements.displayArea.querySelectorAll('.table-wrapper:not([style*="display: none"]) table'));
        if (allVisibleTables.length === 0) { alert('沒有可合併的表格。'); return; }
        if (mode === 'checked') {
            let checkedRowsInVisibleTables = 0;
            allVisibleTables.forEach(table => { checkedRowsInVisibleTables += table.querySelectorAll('tbody .row-checkbox:checked').length; });
            if (checkedRowsInVisibleTables === 0) { alert('請先在 *可見* 的表格中勾選至少一個資料列。'); return; }
        }
        const allHeaders = new Set(), tableData = [], tableHeaderMap = new Map();
        allVisibleTables.forEach(table => {
            let headers = Array.from(table.querySelectorAll('thead th:not(.checkbox-cell)')).map((th, i) => th.textContent.trim() || `(欄位 ${i + 1})`);
            const allDataRowsInTable = Array.from(table.querySelectorAll('tbody tr'));
            if (headers.length === 0 && allDataRowsInTable.length > 0) {
                let maxCols = 0;
                allDataRowsInTable.slice(0, 10).forEach(r => { const colCount = r.querySelectorAll('td:not(.checkbox-cell)').length; if (colCount > maxCols) maxCols = colCount; });
                headers = Array.from({ length: maxCols }, (_, i) => `(欄位 ${i + 1})`);
            }
            headers.forEach(h => allHeaders.add(h));
            tableHeaderMap.set(table, headers);
        });
        allVisibleTables.forEach(table => {
            const headers = tableHeaderMap.get(table); if (!headers) return; 
            const filename = table.closest('.table-wrapper')?.querySelector('h4')?.textContent || '未知來源';
            let rowsToProcess = (mode === 'all') ? Array.from(table.querySelectorAll('tbody tr')).filter(row => !row.classList.contains('row-hidden-search')) : Array.from(table.querySelectorAll('tbody .row-checkbox:checked')).map(cb => cb.closest('tr'));
            rowsToProcess.forEach(row => {
                const rowData = { _sourceFile: filename };
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
        if (state.isEditing && !confirm("您有未儲存的編輯，確定要關閉並捨棄變更嗎？")) return;
        elements.mergeViewModal.classList.add('hidden');
        document.body.classList.remove('no-scroll');
        state.isMergedView = false; state.isEditing = false; state.showTotalRow = false; state.showSourceColumn = false; 
        elements.toggleSourceColBtn.textContent = '新增來源欄位'; 
        elements.toggleSourceColBtn.classList.remove('active'); 
        state.mergedData = []; state.mergedHeaders = [];
        elements.mergeViewContent.innerHTML = '';
        
        elements.collapsibleToolbar.classList.remove('collapsed');
        elements.toggleToolbarBtn.textContent = '收合工具列';
        elements.toggleToolbarBtn.title = '收合工具列';
        
        elements.colSelect1.innerHTML = '<option value="">-- 選擇欄位 1 --</option>'; 
        elements.colSelect2.innerHTML = '<option value="">-- 選擇欄位 2 (選填) --</option>';
        elements.inputCriteria1.value = ''; elements.inputCriteria2.value = '';
        elements.inputCriteria1.disabled = true; elements.inputCriteria2.disabled = true;
        document.querySelector('input[name="criteria-1"][value="empty"]').checked = true;
        document.querySelector('input[name="criteria-2"][value="empty"]').checked = true;
        document.querySelector('input[name="logic-op"][value="and"]').checked = true;

        elements.searchInputMerged.value = '';
        elements.selectKeywordInputMerged.value = '';
        elements.selectKeywordRegexMerged.checked = false;
        toggleEditMode(false);
    }

    function renderMergedTable() {
        const table = document.createElement('table'), thead = table.createTHead(), headerRow = thead.insertRow();
        if (state.showSourceColumn) {
            const thSource = document.createElement('th');
            thSource.textContent = '來源檔案'; thSource.classList.add('source-col'); headerRow.prepend(thSource);
        }
        state.mergedHeaders.forEach(header => {
            const th = document.createElement('th'); th.textContent = header;
            const deleteBtn = document.createElement('span');
            deleteBtn.className = 'delete-col-btn'; deleteBtn.innerHTML = '&times;';
            deleteBtn.title = `刪除 ${header}`; deleteBtn.dataset.header = header;
            th.appendChild(deleteBtn); headerRow.appendChild(th); 
        });
        const tbody = table.createTBody();
        state.mergedData.forEach((rowData, index) => {
            const tr = tbody.insertRow(); tr.dataset.rowIndex = index;
            if (rowData._sourceFile) tr.title = `來源: ${rowData._sourceFile}`;
            if (state.showSourceColumn) {
                const tdSource = document.createElement('td');
                tdSource.textContent = rowData._sourceFile || ''; tdSource.classList.add('source-col'); tr.prepend(tdSource);
            }
            if (rowData._isNew) tr.classList.add('new-row-highlight');
            state.mergedHeaders.forEach(header => {
                const td = tr.insertCell(); td.contentEditable = state.isEditing; td.dataset.colHeader = header;
                const value = rowData[header] || '';
                const cleanVal = String(value).replace(/,/g, '').trim();
                const isStrictNumber = cleanVal !== '' && !isNaN(cleanVal);
                td.classList.toggle('numeric', isStrictNumber);
                td.textContent = isStrictNumber ? formatNumber(value) : value;
            });
        });
        if (state.showTotalRow) {
            const tfoot = table.createTFoot(), totalRow = tfoot.insertRow();
            totalRow.className = 'total-row'; totalRow.innerHTML = ''; totalRow.insertCell(); 
            if (state.showSourceColumn) totalRow.insertCell().textContent = ''; 
            const totalsCache = calculateTotals(); let totalLabelApplied = false;
            state.mergedHeaders.forEach(header => {
                const td = totalRow.insertCell(), totalVal = totalsCache[header];
                if (totalVal) { td.textContent = formatNumber(totalVal); td.classList.add('numeric'); } 
                else if (!totalLabelApplied) { td.textContent = '合計'; totalLabelApplied = true; } 
                else { td.textContent = ''; }
            });
        }
        elements.mergeViewContent.innerHTML = '';
        elements.mergeViewContent.appendChild(table);
        elements.mergeViewContent.classList.toggle('is-editing', state.isEditing);
        injectCheckboxes(elements.mergeViewContent); 
        const selectAllCheckbox = elements.mergeViewContent.querySelector('thead input[type="checkbox"]');
        if (selectAllCheckbox) selectAllCheckbox.addEventListener('change', (e) => {
            elements.mergeViewContent.querySelectorAll('.row-checkbox').forEach(cb => cb.checked = e.target.checked);
        });
    }
    
    function toggleSourceColumn() {
        if (state.isEditing) { alert('請先儲存或取消編輯。'); return; }
        state.showSourceColumn = !state.showSourceColumn;
        renderMergedTable(); 
        elements.toggleSourceColBtn.textContent = state.showSourceColumn ? '移除來源欄位' : '新增來源欄位';
        elements.toggleSourceColBtn.classList.toggle('active', state.showSourceColumn);
    }

    function updateColumnSelects(headers) { 
        elements.columnChecklist.innerHTML = headers.map(header => `<label><input type="checkbox" value="${header}" checked> ${header}</label>`).join('');
        const createOption = (value, text) => { const opt = document.createElement('option'); opt.value = value; opt.textContent = text; return opt; };
        elements.colSelect1.innerHTML = ''; elements.colSelect1.appendChild(createOption('', '-- 選擇欄位 1 --'));
        headers.forEach(h => elements.colSelect1.appendChild(createOption(h, h)));
        elements.colSelect2.innerHTML = ''; elements.colSelect2.appendChild(createOption('', '-- 選擇欄位 2 (選填) --'));
        headers.forEach(h => elements.colSelect2.appendChild(createOption(h, h)));
    }
    
    function toggleColumnModal(forceShow) { elements.columnModal.classList.toggle('hidden', forceShow === false || !elements.columnModal.classList.contains('hidden')); }
    function setAllColumnCheckboxes(isChecked) { elements.columnChecklist.querySelectorAll('input').forEach(input => input.checked = isChecked); }
    function applyColumnChanges() {
        const mergedTable = elements.mergeViewContent.querySelector('table'); if (!mergedTable) return;
        const visibility = {};
        elements.columnChecklist.querySelectorAll('input').forEach(input => { visibility[input.value] = input.checked; });
        const allHeaders = Array.from(mergedTable.querySelectorAll('thead th'));
        const firstDataColIndex = allHeaders.findIndex(th => !th.classList.contains('checkbox-cell') && !th.classList.contains('source-col'));
        if (firstDataColIndex === -1) return; 
        const dataHeaders = allHeaders.slice(firstDataColIndex);
        dataHeaders.forEach((th, dataIndex) => {
            const colIndex = dataIndex + firstDataColIndex, headerText = th.textContent.replace('×', '').trim();
            mergedTable.querySelectorAll(`tr > *:nth-child(${colIndex + 1})`).forEach(cell => cell.classList.toggle('column-hidden', !visibility[headerText]));
        });
    }
    function handleMergedHeaderClick(th) {
        if (state.isEditing || th.classList.contains('source-col')) return;
        const table = th.closest('table'), headerText = th.textContent.replace('×','').trim(), isAsc = th.classList.contains('sort-asc');
        table.querySelectorAll('th').forEach(h => h.classList.remove('sort-asc', 'sort-desc'));
        th.classList.add(isAsc ? 'sort-desc' : 'sort-asc');
        state.mergedData.sort((a, b) => {
            const valA = a[headerText] || '', valB = b[headerText] || '';
            const numA = parseFloat(String(valA).replace(/,/g, '')), numB = parseFloat(String(valB).replace(/,/g, ''));
            const comparison = (!isNaN(numA) && !isNaN(numB)) ? numA - numB : valA.localeCompare(valB, undefined, { numeric: true, sensitivity: 'base' });
            return isAsc ? -comparison : comparison;
        });
        renderMergedTable();
    }
    function deleteColumn(headerToDelete) {
        if (confirm(`確定要刪除「${headerToDelete}」這個欄位嗎？此操作無法復原。`)) {
            state.mergedHeaders = state.mergedHeaders.filter(h => h !== headerToDelete);
            state.mergedData.forEach(row => { delete row[headerToDelete]; });
            renderMergedTable();
            updateColumnSelects(state.mergedHeaders); 
        }
    }
    function calculateTotals() {
        const totals = {};
        state.mergedHeaders.forEach(header => {
            const sum = state.mergedData.reduce((acc, row) => acc + (isNaN(parseFloat(String(row[header]).replace(/,/g, ''))) ? 0 : parseFloat(String(row[header]).replace(/,/g, ''))), 0);
            if (sum !== 0 || state.mergedData.some(row => !isNaN(parseFloat(String(row[header]).replace(/,/g, ''))))) totals[header] = sum;
        });
        return totals;
    }

    function toggleEditMode(startEditing) {
        state.isEditing = startEditing;
        elements.editDataBtn.classList.toggle('hidden', state.isEditing);
        elements.saveEditsBtn.classList.toggle('hidden', !state.isEditing);
        elements.cancelEditsBtn.classList.toggle('hidden', !state.isEditing);
        const disableOnEdit = ['addNewRowBtn', 'copySelectedRowsBtn', 'deleteMergedRowsBtn', 'columnOperationsBtn', 'toggleTotalRowBtn', 'toggleSourceColBtn', 'invertSelectionMergedBtn', 'exportSelectedMergedXlsxBtn', 'exportCurrentMergedXlsxBtn', 'sortMergedByNameBtn', 'colSelect1', 'colSelect2', 'executeFilterSelectionBtn', 'searchInputMerged', 'selectKeywordInputMerged', 'selectKeywordRegexMerged'];
        disableOnEdit.forEach(elId => { if (elements[elId]) elements[elId].disabled = state.isEditing; });
        elements.inputCriteria1.disabled = true; elements.inputCriteria2.disabled = true;
        document.querySelectorAll('input[name="criteria-1"], input[name="criteria-2"], input[name="logic-op"]').forEach(r => r.disabled = state.isEditing);
        renderMergedTable();
    }
    function saveEdits() {
        const newData = Array.from(elements.mergeViewContent.querySelectorAll('tbody tr')).map(tr => {
            const newRowData = {};
            const sourceCell = tr.querySelector('.source-col');
            if (sourceCell) newRowData._sourceFile = sourceCell.textContent;
            else {
                const originalIndex = parseInt(tr.dataset.rowIndex, 10);
                newRowData._sourceFile = (!isNaN(originalIndex) && state.mergedData[originalIndex]) ? state.mergedData[originalIndex]._sourceFile : ' (新增資料列)';
            }
            tr.querySelectorAll('td[data-col-header]').forEach(cell => newRowData[cell.dataset.colHeader] = cell.textContent);
            return newRowData;
        });
        state.mergedData = newData;
        toggleEditMode(false);
    }
    function addNewRow() {
        const newRow = { _isNew: true, _sourceFile: ' (新增資料列)' };
        state.mergedHeaders.forEach(header => { newRow[header] = ''; });
        state.mergedData.unshift(newRow);
        toggleEditMode(true);
    }
    function copySelectedRows() {
        const selectedCheckboxes = elements.mergeViewContent.querySelectorAll('.row-checkbox:checked');
        if (selectedCheckboxes.length === 0) { alert("請先勾選要複製的資料列。"); return; }
        const rowsToCopy = Array.from(selectedCheckboxes).map(cb => {
            const rowIndex = parseInt(cb.closest('tr').dataset.rowIndex, 10);
            if (!isNaN(rowIndex) && state.mergedData[rowIndex]) {
                const newRow = JSON.parse(JSON.stringify(state.mergedData[rowIndex]));
                newRow._isNew = true; newRow._sourceFile += ' (複製)'; return newRow;
            }
            return null;
        }).filter(Boolean);
        state.mergedData.unshift(...rowsToCopy);
        toggleEditMode(true);
    }

    function syncCheckboxesInScope() { setTimeout(() => { const scope = state.isMergedView ? elements.mergeViewContent : elements.displayArea; scope.querySelectorAll('table').forEach(syncTableCheckboxState); if (!state.isMergedView) updateSelectionInfo(); }, 0); }
    function selectAllRows() { const scope = state.isMergedView ? elements.mergeViewContent : elements.displayArea; const rows = scope.querySelectorAll('tbody tr:not(.row-hidden-search)'); if (rows.length === 0) { alert('沒有可勾選的列'); return; } rows.forEach(row => row.querySelector('.row-checkbox').checked = true); scope.querySelectorAll('thead input[type="checkbox"]').forEach(cb => cb.checked = true); }
    function invertSelection() { const scope = state.isMergedView ? elements.mergeViewContent : elements.displayArea; scope.querySelectorAll('tbody tr:not(.row-hidden-search) .row-checkbox').forEach(cb => cb.checked = !cb.checked); }
    function deleteSelectedRows() {
        const scope = state.isMergedView ? elements.mergeViewContent : elements.displayArea;
        const selected = scope.querySelectorAll('tbody .row-checkbox:checked');
        if (selected.length === 0) { alert('請先勾選要刪除的列'); return; }
        if (confirm(`確定要刪除 ${selected.length} 筆資料列嗎？`)) {
            if (state.isMergedView) {
                const indicesToDelete = new Set(Array.from(selected).map(cb => parseInt(cb.closest('tr').dataset.rowIndex, 10)).filter(i => !isNaN(i)));
                state.mergedData = state.mergedData.filter((_, index) => !indicesToDelete.has(index));
                renderMergedTable();
            } else {
                selected.forEach(cb => cb.closest('tr').remove());
            }
        }
        syncCheckboxesInScope();
    }
    function selectEmptyRows() { let count = 0; const scope = state.isMergedView ? elements.mergeViewContent : elements.displayArea; scope.querySelectorAll('tbody tr:not(.row-hidden-search)').forEach(row => { if (Array.from(row.cells).slice(1).every(c => c.textContent.trim() === '')) { row.querySelector('.row-checkbox').checked = true; count++; } }); if (count === 0) alert('未找到空白列'); }
    
    function selectByKeyword() { 
        const inputEl = state.isMergedView ? elements.selectKeywordInputMerged : elements.selectKeywordInput;
        const regexEl = state.isMergedView ? elements.selectKeywordRegexMerged : elements.selectKeywordRegex;
        const keywordInput = inputEl.value.trim(); 
        if (!keywordInput) { alert('請先輸入關鍵字'); return; } 
        let matchLogic; 
        try { 
            if (regexEl.checked) matchLogic = text => new RegExp(keywordInput, 'i').test(text); 
            else if (keywordInput.includes(',')) { const keywords = keywordInput.split(',').map(k => k.trim().toLowerCase()).filter(Boolean); matchLogic = text => keywords.some(k => text.includes(k)); } 
            else { const keywords = keywordInput.split(/\s+/).map(k => k.trim().toLowerCase()).filter(Boolean); matchLogic = text => keywords.every(k => text.includes(k)); } 
        } catch (e) { alert('無效的 Regex 表示式：\n' + e.message); return; } 
        let count = 0; 
        const scope = state.isMergedView ? elements.mergeViewContent : elements.displayArea; 
        scope.querySelectorAll('tbody tr:not(.row-hidden-search)').forEach(row => { 
            let rowText = Array.from(row.querySelectorAll('td:not(.checkbox-cell)')).map(c => c.textContent).join(' ');
            if (matchLogic(rowText.toLowerCase())) { row.querySelector('.row-checkbox').checked = true; count++; } 
        }); 
        alert(count > 0 ? `已勾選 ${count} 個符合條件的列` : `未找到符合條件的列`); 
    }
    
    function filterTable() { 
        const inputEl = state.isMergedView ? elements.searchInputMerged : elements.searchInput;
        const keywords = inputEl.value.toLowerCase().trim().split(/\s+/).filter(Boolean); 
        const scope = state.isMergedView ? elements.mergeViewContent : elements.displayArea; 
        scope.querySelectorAll('tbody tr').forEach(row => { 
            let rowText = Array.from(row.querySelectorAll('td:not(.checkbox-cell)')).map(c => c.textContent).join(' ').toLowerCase();
            row.classList.toggle('row-hidden-search', !keywords.every(k => rowText.includes(k))); 
        }); 
        if (!state.isMergedView) { 
            elements.displayArea.querySelectorAll('.table-wrapper').forEach(wrapper => { 
                wrapper.style.display = wrapper.querySelectorAll('tbody tr:not(.row-hidden-search)').length > 0 ? '' : 'none'; 
            }); 
        } 
        syncCheckboxesInScope(); 
    }

    function toggleToolbar() {
        const isCollapsed = elements.collapsibleToolbar.classList.toggle('collapsed');
        elements.toggleToolbarBtn.textContent = isCollapsed ? '展開工具列' : '收合工具列';
        elements.toggleToolbarBtn.title = isCollapsed ? '展開工具列' : '收合工具列';
    }
    function formatNumber(value) { try { const num = parseFloat(String(value).replace(/,/g, '')); return isNaN(num) ? value : new Intl.NumberFormat('en-US').format(num); } catch (e) { return value; } }
    function detectHiddenElements() { return elements.displayArea.querySelectorAll('tr[style*="display: none"], td[style*="display: none"], th[style*="display: none"]').length; }
    function showAllHiddenElements() { const hidden = elements.displayArea.querySelectorAll('tr[style*="display: none"], td[style*="display: none"], th[style*="display: none"]'); if (hidden.length === 0) { alert('沒有需要顯示的隱藏行列。'); return; } hidden.forEach(el => el.style.display = ''); alert(`已顯示 ${hidden.length} 個隱藏的行列。`); elements.showHiddenBtn.classList.add('hidden'); elements.loadStatusMessage.classList.add('hidden'); }
    function selectAllTables(isChecked) { elements.displayArea.querySelectorAll('.table-select-checkbox').forEach(cb => { if (cb.checked !== isChecked) cb.click(); }); }
    function readFileAsBinary(file) { return new Promise((resolve, reject) => { const reader = new FileReader(); reader.onload = e => resolve(e.target.result); reader.onerror = reject; reader.readAsBinaryString(file); }); }
    function parsePositionString(str) { const indices = new Set(); str.split(',').map(p => p.trim()).filter(Boolean).forEach(part => { if (part.includes('-')) { const [start, end] = part.split('-').map(Number); if (!isNaN(start) && !isNaN(end) && start <= end) for (let i = start; i <= end; i++) indices.add(i - 1); } else { const num = Number(part); if (!isNaN(num)) indices.add(num - 1); } }); return Array.from(indices).sort((a, b) => a - b); }
    async function getSelectedSheetNames(filename, workbook, mode, criteria) { const sheetNames = workbook.SheetNames; if (sheetNames.length === 0) return []; switch (mode) { case 'all': return sheetNames; case 'first': return sheetNames.length > 0 ? [sheetNames[0]] : []; case 'specific': return sheetNames.filter(name => name.toLowerCase().includes(criteria.name.toLowerCase())); case 'position': return parsePositionString(criteria.position).map(index => sheetNames[index]).filter(Boolean); case 'manual': return await showWorksheetSelectionModal(filename, sheetNames); default: return []; } }
    function showWorksheetSelectionModal(filename, sheetNames) { return new Promise(resolve => { if (sheetNames.length <= 1) { resolve(sheetNames); return; } const overlay = document.createElement('div'); overlay.className = 'modal-overlay'; const dialog = document.createElement('div'); dialog.className = 'modal-dialog'; dialog.innerHTML = `<div class="modal-header"><h3>選擇工作表 (手動模式)</h3><p>檔案 "<strong>${filename}</strong>"</p></div><div class="modal-body"><ul class="sheet-list">${sheetNames.map(name => `<li class="sheet-item"><label><input type="checkbox" class="sheet-checkbox" value="${name}" checked> ${name}</label></li>`).join('')}</ul></div><div class="modal-footer"><button class="btn btn-secondary" id="modal-skip">跳過</button><button class="btn btn-success" id="modal-confirm">確認</button></div>`; overlay.appendChild(dialog); document.body.appendChild(overlay); const closeModal = () => document.body.removeChild(overlay); dialog.querySelector('#modal-confirm').addEventListener('click', () => { resolve(Array.from(dialog.querySelectorAll('.sheet-checkbox:checked')).map(cb => cb.value)); closeModal(); }); dialog.querySelector('#modal-skip').addEventListener('click', () => { resolve([]); closeModal(); }); }); }
    function extractTableData(table, { onlySelected = false, includeFilename = false } = {}) { const data = []; const headerRow = table.querySelector('thead tr'); if (headerRow) { let headerData = Array.from(headerRow.querySelectorAll('th:not(.checkbox-cell):not(.column-hidden)')).map(th => th.textContent.replace('×','').trim()); if (includeFilename) headerData.unshift('Source File'); data.push(headerData); } const filename = includeFilename ? (table.closest('.table-wrapper')?.querySelector('h4')?.textContent || 'Merged Table') : null; let rows = onlySelected ? Array.from(table.querySelectorAll('tbody .row-checkbox:checked')).map(cb => cb.closest('tr')) : table.querySelectorAll('tbody tr:not(.row-hidden-search)'); rows.forEach(row => { let rowData = Array.from(row.querySelectorAll('td:not(.checkbox-cell):not(.column-hidden)')).map(td => td.textContent.trim()); if (includeFilename) rowData.unshift(filename); data.push(rowData); }); return data; }
    function exportToXlsx(data, filename, sheetName) { if (data.length <= 1) { alert('沒有足夠的資料可以匯出。'); return; } try { const ws = XLSX.utils.aoa_to_sheet(data); ws['!cols'] = calculateColumnWidths(data); const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, sheetName); XLSX.writeFile(wb, filename); } catch (err) { console.error('匯出 XLSX 時發生錯誤:', err); alert('匯出時發生錯誤：' + err.message); } }
    function exportCurrentMergedXlsx() { if (!state.isMergedView) return; const table = elements.mergeViewContent.querySelector('table'); if (!table) return; const data = extractTableData(table, { onlySelected: false, includeFilename: state.showSourceColumn }); exportToXlsx(data, `merged_view_export_${new Date().toISOString().slice(0, 10)}.xlsx`, "Merged View Data"); }
    function exportSelectedMergedXlsx() { if (!state.isMergedView) return; const table = elements.mergeViewContent.querySelector('table'); if (!table) return; const data = extractTableData(table, { onlySelected: true, includeFilename: state.showSourceColumn }); exportToXlsx(data, `merged_selected_export_${new Date().toISOString().slice(0, 10)}.xlsx`, "Merged Selected Data"); }
    function exportMergedXlsx() { const tables = Array.from(elements.displayArea.querySelectorAll('.table-wrapper:not([style*="display: none"]) table')); if (tables.length === 0) { alert('沒有可匯出的表格。'); return; } const allData = []; tables.forEach((table, i) => { const data = extractTableData(table, { includeFilename: true }); if (data.length > 1) allData.push(...(i === 0 ? data : data.slice(1))); }); exportToXlsx(allData, `report_merged_${new Date().toISOString().slice(0, 10)}.xlsx`, 'Merged Data'); }
    function calculateColumnWidths(data) { if (data.length === 0) return []; return data[0].map((_, col) => ({ wch: Math.min(50, Math.max(10, ...data.map(row => row[col] ? String(row[col]).length : 0)) + 2) })); }
    
    function resetView() { if(state.isMergedView) closeMergeView(); if (!state.originalHtmlString) return; elements.displayArea.innerHTML = state.originalHtmlString; injectCheckboxes(elements.displayArea); ['searchInput', 'selectKeywordInput'].forEach(id => elements[id].value = ''); elements.selectKeywordRegex.checked = false; filterTable(); elements.loadStatusMessage.classList.add('hidden'); const hiddenCount = detectHiddenElements(); if (hiddenCount > 0) { elements.loadStatusMessage.textContent = `注意：已重設表格，${hiddenCount} 個隱藏的行列已還原。`; elements.loadStatusMessage.classList.remove('hidden'); elements.showHiddenBtn.classList.remove('hidden'); } else { elements.showHiddenBtn.classList.add('hidden'); } updateSelectionInfo(); setViewMode('list'); }
    function resetControls(isNewFile) { if (!isNewFile) return; if(state.isMergedView) closeMergeView(); state.originalHtmlString = ''; ['searchInput', 'selectKeywordInput'].forEach(id => elements[id].value = ''); elements.selectKeywordRegex.checked = false; elements.controlPanel.classList.add('hidden'); updateSelectionInfo(); }
    function clearAllFiles(silent = false) { if (!silent && !confirm('確定要清除所有已載入的檔案嗎？')) return; if(state.isMergedView) closeMergeView(); state.originalHtmlString = ''; state.loadedFiles = []; state.loadedTables = 0; elements.displayArea.innerHTML = ''; elements.fileInput.value = ''; ['specificSheetNameInput', 'specificSheetPositionInput'].forEach(id => elements[id].value = ''); elements.gridScaleSlider.value = 3; elements.colSelect1.innerHTML = '<option value="">-- 選擇欄位 1 --</option>'; elements.colSelect2.innerHTML = '<option value="">-- 選擇欄位 2 (選填) --</option>'; updateGridScale(); updateDropAreaDisplay(); resetControls(true); setViewMode('list'); }
    function updateDropAreaDisplay() { const hasFiles = state.loadedTables > 0; elements.dropArea.classList.toggle('compact', hasFiles); elements.dropAreaInitial.classList.toggle('hidden', hasFiles); elements.dropAreaLoaded.classList.toggle('hidden', !hasFiles); elements.importOptionsContainer.classList.toggle('hidden', hasFiles); if (hasFiles) { elements.fileCount.textContent = state.loadedTables; const names = state.loadedFiles.slice(0, 3).join(', '); const more = state.loadedFiles.length > 3 ? ` 及其他 ${state.loadedFiles.length - 3} 個...` : ''; elements.fileNames.textContent = names + more; } }
    function showControls(hiddenCount) { elements.controlPanel.classList.remove('hidden'); const buttonsToShow = ['selectByKeywordGroup', 'selectByKeywordBtn', 'selectEmptyBtn', 'deleteSelectedBtn', 'invertSelectionBtn', 'selectAllBtn', 'exportMergedXlsxBtn', 'resetViewBtn', 'tableLevelControls', 'listViewBtn', 'gridViewBtn', 'showHiddenBtn', 'viewCheckedCombinedBtn', 'sortByNameBtn']; buttonsToShow.forEach(id => { if(elements[id]) elements[id].classList.remove('hidden'); }); elements.mergeViewBtn.classList.toggle('hidden', state.loadedTables <= 1); const showHiddenStuff = hiddenCount > 0; elements.loadStatusMessage.classList.toggle('hidden', !showHiddenStuff); elements.showHiddenBtn.classList.toggle('hidden', !showHiddenStuff); if (showHiddenStuff) elements.loadStatusMessage.textContent = `注意：檔案中包含 ${hiddenCount} 個被隱藏的行列。`; }
    function setViewMode(mode) { if (mode === 'grid') { elements.displayArea.classList.remove('list-view'); elements.displayArea.classList.add('grid-view'); elements.gridViewBtn.classList.add('active'); elements.listViewBtn.classList.remove('active'); elements.gridScaleControl.classList.remove('hidden'); } else { elements.displayArea.classList.remove('grid-view'); elements.displayArea.classList.add('list-view'); elements.listViewBtn.classList.add('active'); elements.gridViewBtn.classList.remove('active'); elements.gridScaleControl.classList.add('hidden'); } }
    function updateGridScale() { elements.displayArea.style.setProperty('--grid-columns', elements.gridScaleSlider.value); }
    function validateFiles(fileList) { if (!fileList || fileList.length === 0) return { valid: false, error: '沒有選擇檔案' }; const validFiles = Array.from(fileList).filter(file => CONSTANTS.VALID_FILE_EXTENSIONS.some(ext => file.name.toLowerCase().endsWith(ext))); if (validFiles.length === 0) return { valid: false, error: '請上傳 .xls 或 .xlsx 格式的檔案' }; return { valid: true, files: validFiles }; }
    function debounce(func, wait) { let timeout; return (...args) => { clearTimeout(timeout); timeout = setTimeout(() => func(...args), wait); }; }
    function handleScroll() { elements.backToTopBtn.classList.toggle('visible', window.scrollY > window.innerHeight / 2); }
    function scrollToTop() { window.scrollTo({ top: 0, behavior: 'smooth' }); }
    function handleCardClick(e) { const card = e.target.closest('.table-wrapper'); if (!card) return; if (e.target.classList.contains('close-zoom')) { closePreview(); return; } if (e.target.classList.contains('delete-rows-btn')) { deleteSelectedRowsInScope(e.target.closest('.table-wrapper')); return; } if (e.target.classList.contains('delete-table-btn')) { if (confirm(`確定要永久刪除此工作表 (${card.querySelector('h4').textContent}) 嗎？`)) { closePreview(); setTimeout(() => { card.remove(); updateFileStateAfterDeletion(); }, 300); } return; } if (elements.displayArea.classList.contains('grid-view') && !card.classList.contains('is-zoomed') && !e.target.matches('input, a, button, .btn')) openPreview(card); }
    function deleteSelectedRowsInScope(scope) { const selected = scope.querySelectorAll('tbody .row-checkbox:checked'); if (selected.length === 0) { alert('請先勾選要刪除的列'); return; } if (confirm(`確定要在此表格內刪除 ${selected.length} 筆資料列嗎？`)) selected.forEach(cb => cb.closest('tr').remove()); syncCheckboxesInScope(); }
    function openPreview(card) { if (state.zoomedCard) return; card.classList.add('is-zoomed'); state.zoomedCard = card; document.body.classList.add('no-scroll'); }
    function closePreview() { if (!state.zoomedCard) return; state.zoomedCard.classList.remove('is-zoomed'); state.zoomedCard = null; document.body.classList.remove('no-scroll'); }
    function syncTableCheckboxState(table) { const headerCheckbox = table.closest('.table-wrapper')?.querySelector('.table-select-checkbox') || table.querySelector('thead input[type="checkbox"]'); if (!headerCheckbox) return; const rowCheckboxes = table.querySelectorAll('tbody tr:not(.row-hidden-search) .row-checkbox'); if (rowCheckboxes.length === 0) { headerCheckbox.checked = false; headerCheckbox.indeterminate = false; return; } const checkedCount = Array.from(rowCheckboxes).filter(cb => cb.checked).length; if (checkedCount === 0) { headerCheckbox.checked = false; headerCheckbox.indeterminate = false; } else if (checkedCount === rowCheckboxes.length) { headerCheckbox.checked = true; headerCheckbox.indeterminate = false; } else { headerCheckbox.checked = false; headerCheckbox.indeterminate = true; } }
    function updateSelectionInfo() { const selectedCheckboxes = elements.displayArea.querySelectorAll('.table-select-checkbox:checked, .table-select-checkbox:indeterminate'); if (selectedCheckboxes.length > 0) { elements.selectedTablesList.textContent = Array.from(selectedCheckboxes).map(cb => cb.closest('.table-header').querySelector('h4').textContent).join('; '); elements.selectedTablesInfo.classList.remove('hidden'); } else { elements.selectedTablesInfo.classList.add('hidden'); } }
    function deleteSelectedTables() { const selectedWrappers = Array.from(elements.displayArea.querySelectorAll('.table-select-checkbox:checked')).map(cb => cb.closest('.table-wrapper')); if (selectedWrappers.length === 0) { alert('請先勾選要刪除的表格。'); return; } if (confirm(`確定要永久刪除 ${selectedWrappers.length} 個選定的表格嗎？`)) { selectedWrappers.forEach(wrapper => wrapper.remove()); updateFileStateAfterDeletion(); } }
    function updateFileStateAfterDeletion() { const remainingWrappers = elements.displayArea.querySelectorAll('.table-wrapper'); state.loadedTables = remainingWrappers.length; state.loadedFiles = Array.from(remainingWrappers).map(w => w.querySelector('h4').textContent); if (state.loadedTables === 0) clearAllFiles(true); else { updateDropAreaDisplay(); showControls(detectHiddenElements()); } }
    function handleDisplayAreaChange(e) { const target = e.target; if (!target.matches('.table-select-checkbox, [id^="select-all-checkbox"], .row-checkbox')) return; let table; if (target.matches('.table-select-checkbox')) { table = target.closest('.table-wrapper')?.querySelector('table'); if (table) toggleSelectAll(target.checked, table); } else { table = target.closest('table'); if (target.matches('[id^="select-all-checkbox"]')) toggleSelectAll(target.checked, table); } if (table) syncTableCheckboxState(table); if (!state.isMergedView) updateSelectionInfo(); }
    function toggleSelectAll(isChecked, table) { if (!table) return; table.querySelectorAll('tbody tr:not(.row-hidden-search) .row-checkbox').forEach(cb => cb.checked = isChecked); const headerCheckbox = table.querySelector('thead input[type="checkbox"]'); if (headerCheckbox) { headerCheckbox.checked = isChecked; headerCheckbox.indeterminate = false; } }
    
    function getFundSortPriority(fileName) {
        if (state.fundSortOrder.length === 0) return { index: Infinity, name: fileName };
        const foundAlias = state.fundAliasKeys.find(alias => fileName.includes(alias));
        const canonicalName = foundAlias ? state.fundAliasMap[foundAlias] : null;
        const index = canonicalName ? state.fundSortOrder.indexOf(canonicalName) : -1;
        return { index: (index === -1) ? Infinity : index, name: fileName };
    }

    function sortTablesByFundName() {
        if (state.fundSortOrder.length === 0 || Object.keys(state.fundAliasMap).length === 0) { console.warn('基金順序列表尚未載入，暫不執行自動排序。'); return; }
        const wrappers = Array.from(elements.displayArea.querySelectorAll('.table-wrapper'));
        const getFileName = (wrapper) => { const text = wrapper.querySelector('h4').textContent; const match = text.match(/(.*)\s\(.*\)$/); return match ? match[1].trim() : text.trim(); };
        wrappers.sort((a, b) => {
            const fileA = getFundSortPriority(getFileName(a)), fileB = getFundSortPriority(getFileName(b));
            return (fileA.index === fileB.index) ? fileA.name.localeCompare(fileB.name) : fileA.index - fileB.index;
        });
        elements.displayArea.innerHTML = '';
        wrappers.forEach(wrapper => elements.displayArea.appendChild(wrapper));
    }

    function sortMergedTableByFundName() {
        if (state.isEditing) { alert('請先儲存或取消編輯。'); return; }
        if (state.fundSortOrder.length === 0 || Object.keys(state.fundAliasMap).length === 0) { alert('錯誤：基金順序列表尚未載入或為空。\n請檢查 fund-config.json 檔案。'); return; }
        state.mergedData.sort((a, b) => {
            const fileA = getFundSortPriority(a._sourceFile || ''), fileB = getFundSortPriority(b._sourceFile || '');
            return (fileA.index === fileB.index) ? (a._sourceFile || '').localeCompare(b._sourceFile || '') : fileA.index - fileB.index;
        });
        renderMergedTable();
    }

    return { init };
})();

ExcelViewer.init();
