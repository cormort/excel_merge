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
        fundSortOrder: [], // <-- NEW: Will hold the sortOrder array
        fundAliasMap: {},   // <-- NEW: Will hold the aliasMap object
        fundAliasKeys: []  // <-- NEW: Cached keys for searching
    };
    const elements = {};

    // ▼▼▼ MODIFIED: Made init async ▼▼▼
    async function init() {
        cacheElements();
        await loadFundConfig(); // <-- NEW: Wait for config before binding events
        bindEvents();
    }
    // ▲▲▲ MODIFIED ▲▲▲

    // --- ▼▼▼ NEW FUNCTION ▼▼▼ ---
    /**
     * Fetches and loads the fund sorting configuration file.
     */
    async function loadFundConfig() {
        try {
            // 為了避免瀏覽器快取，加入一個隨機參數
            const response = await fetch(`fund-config.json?v=${Date.now()}`); 
            if (!response.ok) {
                throw new Error(`HTTP error! status: ${response.status}`);
            }
            const config = await response.json();
            
            if (config.sortOrder && config.aliasMap) {
                state.fundSortOrder = config.sortOrder;
                state.fundAliasMap = config.aliasMap;
                // 將 aliasMap 的 key 儲存起來，並依照長度排序，優先比對長的關鍵字
                state.fundAliasKeys = Object.keys(config.aliasMap).sort((a, b) => b.length - a.length);
                console.log("基金設定檔 (fund-config.json) 載入成功。");
            } else {
                console.error("基金設定檔格式錯誤：缺少 sortOrder 或 aliasMap。");
                alert("錯誤：基金設定檔 (fund-config.json) 格式不正確。");
            }
        } catch (err) {
            console.error("載入基金設定檔 (fund-config.json) 失敗:", err);
            // 允許程式繼續運行，但排序功能會失效
            alert("警告：無法載入基金排序設定檔 (fund-config.json)。\n「依基金名稱排序」功能將無法使用。");
        }
    }
    // --- ▲▲▲ NEW FUNCTION ▲▲▲ ---


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
            // Merge view Modal
            mergeViewModal: 'merge-view-modal',
            closeMergeViewBtn: 'close-merge-view-btn',
            mergeViewContent: 'merge-view-content',
            mergeViewBtn: 'merge-view-btn',
            // Column operations
            columnOperationsBtn: 'column-operations-btn',
            columnModal: 'column-modal',
            closeColumnModalBtn: 'close-column-modal-btn',
            columnChecklist: 'column-checklist',
            applyColumnChangesBtn: 'apply-column-changes-btn',
            modalCheckAll: 'modal-check-all',
            modalUncheckAll: 'modal-uncheck-all',
            // Edit operations
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
            sortMergedByNameBtn: 'sort-merged-by-fund-name-btn', // <-- NEW ELEMENT
            // --- MOVED ELEMENTS (IDs are the same, now inside merge modal) ---
            viewCheckedCombinedBtn: 'view-checked-combined-btn',
            columnSelectOps: 'column-select-ops',
            selectByColZeroBtn: 'select-by-col-zero-btn',
            selectByColEmptyBtn: 'select-by-col-empty-btn',
            selectByColExistsBtn: 'select-by-col-exists-btn',

            // --- NEW MERGED VIEW CONTROLS ---
            searchInputMerged: 'search-input-merged',
            selectKeywordInputMerged: 'select-keyword-input-merged',
            selectKeywordRegexMerged: 'select-keyword-regex-merged',
            selectByKeywordBtnMerged: 'select-by-keyword-btn-merged'
        };
        Object.keys(mapping).forEach(key => {
            elements[key] = document.getElementById(mapping[key]);
        });
    }

    function bindEvents() {
        // --- Core File Handling ---
        elements.fileInput.addEventListener('change', e => processFiles(e.target.files));
        setupDragAndDrop();
        elements.clearFilesBtn.addEventListener('click', () => clearAllFiles(false));

        // --- View and Table Level Controls ---
        elements.listViewBtn.addEventListener('click', () => setViewMode('list'));
        elements.gridViewBtn.addEventListener('click', () => setViewMode('grid'));
        elements.gridScaleSlider.addEventListener('input', updateGridScale);
        elements.selectAllTablesBtn.addEventListener('click', () => { selectAllTables(true); updateSelectionInfo(); });
        elements.unselectAllTablesBtn.addEventListener('click', () => { selectAllTables(false); updateSelectionInfo(); });
        elements.deleteSelectedTablesBtn.addEventListener('click', deleteSelectedTables);
        elements.sortByNameBtn.addEventListener('click', sortTablesByFundName); 
        
        // --- Row Operations (Main View) ---
        elements.selectByKeywordBtn.addEventListener('click', () => { selectByKeyword(); syncCheckboxesInScope(); });
        elements.selectEmptyBtn.addEventListener('click', () => { selectEmptyRows(); syncCheckboxesInScope(); });
        elements.selectAllBtn.addEventListener('click', () => { selectAllRows(); syncCheckboxesInScope(); });
        elements.invertSelectionBtn.addEventListener('click', () => { invertSelection(); syncCheckboxesInScope(); });
        elements.deleteSelectedBtn.addEventListener('click', deleteSelectedRows);

        // --- ROW OPERATIONS (Now inside merge view, but events are bound globally) ---
        elements.selectByColZeroBtn.addEventListener('click', () => { selectByColumn('zero'); syncCheckboxesInScope(); });
        elements.selectByColEmptyBtn.addEventListener('click', () => { selectByColumn('empty'); syncCheckboxesInScope(); });
        elements.selectByColExistsBtn.addEventListener('click', () => { selectByColumn('exists'); syncCheckboxesInScope(); });
        
        // --- Global, Export, and Merge Operations ---
        elements.resetViewBtn.addEventListener('click', resetView);
        elements.showHiddenBtn.addEventListener('click', showAllHiddenElements);
        elements.exportMergedXlsxBtn.addEventListener('click', exportMergedXlsx);
        
        // --- Merge View Modal Events ---
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


        // --- Input and Dynamic Content Handling ---
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
        elements.importOptionsContainer.addEventListener('change', e => {
            if (e.target.name === 'import-mode') {
                const selectedMode = e.target.value;
                elements.specificSheetNameGroup.classList.toggle('hidden', selectedMode !== 'specific');
                elements.specificSheetPositionGroup.classList.toggle('hidden', selectedMode !== 'position');
            }
        });

        // --- ▼▼▼ NEW BINDINGS FOR MERGE VIEW & ENTER KEY ---
        elements.searchInputMerged.addEventListener('input', debounce(filterTable, 300));
        elements.selectByKeywordBtnMerged.addEventListener('click', () => { selectByKeyword(); syncCheckboxesInScope(); });

        const handleKeywordEnter = (e) => {
            if (e.key === 'Enter') {
                e.preventDefault();
                // 觸發目前 scope 對應的按鈕點擊
                if (state.isMergedView) {
                    elements.selectByKeywordBtnMerged.click();
                } else {
                    elements.selectByKeywordBtn.click();
                }
            }
        };
        elements.selectKeywordInput.addEventListener('keydown', handleKeywordEnter);
        elements.selectKeywordInputMerged.addEventListener('keydown', handleKeywordEnter);
        // --- ▲▲▲ NEW BINDINGS ---


        // --- Window/Document Level Events ---
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
    }

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
    
    // ▼▼▼ MODIFIED: Added blank row removal ▼▼▼
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
                const workbook = XLSX.read(binaryData, { type: 'binary' }); 
                const sheetNames = await getSelectedSheetNames(file.name, workbook, importMode, { name: specificSheetName, position: specificSheetPosition }); 
                
                if ((importMode === 'specific' || importMode === 'position') && sheetNames.length === 0 && workbook.SheetNames.length > 0) { 
                    missedFiles.push(file.name); 
                } 
                
                for (const sheetName of sheetNames) { 
                    const sheet = workbook.Sheets[sheetName];
                    
                    // --- NEW: Remove blank rows ---
                    const jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
                    const filteredData = jsonData.filter(row => 
                        Array.isArray(row) && row.some(cell => String(cell).trim() !== '')
                    );
                    const cleanedSheet = XLSX.utils.aoa_to_sheet(filteredData);
                    // --- End: Remove blank rows ---

                    const htmlString = XLSX.utils.sheet_to_html(cleanedSheet); // Use cleanedSheet
                    
                    tablesToRender.push({ html: htmlString, filename: `${file.name} (${sheetName})` }); 
                    state.loadedFiles.push(`${file.name} (${sheetName})`); 
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
    // ▲▲▲ MODIFIED ▲▲▲
    
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
        
        // ▼▼▼ NEW: Auto-sort on render ▼▼▼
        sortTablesByFundName();
        // ▲▲▲ NEW ▲▲▲
    }
    
    function injectCheckboxes(scope) { scope.querySelectorAll('thead tr').forEach((headRow, index) => { if(headRow.querySelector('.checkbox-cell')) return; const th = document.createElement('th'); th.innerHTML = `<input type="checkbox" id="select-all-checkbox-${scope.id}-${index}" title="全選/全不選">`; th.classList.add('checkbox-cell'); headRow.prepend(th); }); scope.querySelectorAll('tbody tr').forEach(row => { if(row.querySelector('.checkbox-cell')) return; const td = document.createElement('td'); td.innerHTML = '<input type="checkbox" class="row-checkbox">'; td.classList.add('checkbox-cell'); row.prepend(td); }); }
    
    // --- Merged View and Column Operations ---

    /**
     * Creates the merged view modal based on selected tables and mode.
     * @param {string} mode - 'all' (all visible rows) or 'checked' (only checked rows)
     */
    function createMergedView(mode = 'all') {
        
        // 1. Get ALL visible tables to build a complete header map.
        const allVisibleTables = Array.from(elements.displayArea.querySelectorAll('.table-wrapper:not([style*="display: none"]) table'));
        if (allVisibleTables.length === 0) {
            alert('沒有可合併的表格。');
            return;
        }

        // 2. If mode is 'checked', do a pre-check.
        if (mode === 'checked') {
            // We must check *within* the visible tables.
            let checkedRowsInVisibleTables = 0;
            allVisibleTables.forEach(table => {
                checkedRowsInVisibleTables += table.querySelectorAll('tbody .row-checkbox:checked').length;
            });
            
            if (checkedRowsInVisibleTables === 0) {
                alert('請先在 *可見* 的表格中勾選至少一個資料列。');
                return;
            }
        }

        const allHeaders = new Set();
        const tableData = [];
        const tableHeaderMap = new Map(); // Cache headers per table

        // 3. First pass: Iterate over *all visible tables* just to gather ALL headers
        allVisibleTables.forEach(table => {
            // Get headers for *this* table using the robust logic
            let headers = Array.from(table.querySelectorAll('thead th:not(.checkbox-cell)'))
                             .map((th, i) => th.textContent.trim() || `(欄位 ${i + 1})`);
            
            const allDataRowsInTable = Array.from(table.querySelectorAll('tbody tr')); // Used for 'no header' check
            
            if (headers.length === 0 && allDataRowsInTable.length > 0) {
                let maxCols = 0;
                allDataRowsInTable.slice(0, 10).forEach(r => { // Sample rows
                    const colCount = r.querySelectorAll('td:not(.checkbox-cell)').length;
                    if (colCount > maxCols) maxCols = colCount;
                });
                headers = [];
                for (let i = 0; i < maxCols; i++) {
                    headers.push(`(欄位 ${i + 1})`);
                }
            }
            
            // Add this table's headers to the global set
            headers.forEach(h => allHeaders.add(h));
            // Cache the headers for this specific table
            tableHeaderMap.set(table, headers);
        });


        // 4. Second pass: Iterate over tables again to gather data based on mode
        allVisibleTables.forEach(table => {
            const headers = tableHeaderMap.get(table); // Get cached headers
            if (!headers) return; // Should not happen

            const wrapper = table.closest('.table-wrapper');
            const filename = wrapper?.querySelector('h4')?.textContent || '未知來源';

            // Select which rows to process based on 'mode'
            let rowsToProcess;
            if (mode === 'all') {
                // Get all rows that are NOT hidden by the search filter
                rowsToProcess = Array.from(table.querySelectorAll('tbody tr')).filter(row => !row.classList.contains('row-hidden-search'));
            } else { // mode === 'checked'
                // Get only rows that are checked
                rowsToProcess = Array.from(table.querySelectorAll('tbody .row-checkbox:checked')).map(cb => cb.closest('tr'));
            }

            // Map the data from *only* the selected rows
            rowsToProcess.forEach(row => {
                const rowData = {};
                rowData._sourceFile = filename; // Add the source file property
                Array.from(row.querySelectorAll('td:not(.checkbox-cell)')).forEach((td, i) => {
                    if (headers[i]) { // Use the correct header for this column index
                        rowData[headers[i]] = td.textContent;
                    }
                });
                tableData.push(rowData); // Push directly to the final flat array
            });
        });

        // 5. Set state and render
        state.mergedHeaders = Array.from(allHeaders);
        state.mergedData = tableData;
        
        renderMergedTable();
        updateColumnSelects(state.mergedHeaders); 
        
        state.isMergedView = true;
        elements.mergeViewModal.classList.remove('hidden');
        document.body.classList.add('no-scroll');
    }

    function closeMergeView() {
        if (state.isEditing && !confirm("您有未儲存的編輯，確定要關閉並捨棄變更嗎？")) {
            return;
        }
        elements.mergeViewModal.classList.add('hidden');
        document.body.classList.remove('no-scroll');
        state.isMergedView = false;
        state.isEditing = false;
        state.showTotalRow = false;
        state.showSourceColumn = false; 
        elements.toggleSourceColBtn.textContent = '新增來源欄位'; 
        elements.toggleSourceColBtn.classList.remove('active'); 
        state.mergedData = [];
        state.mergedHeaders = [];
        elements.mergeViewContent.innerHTML = '';
        elements.columnSelectOps.innerHTML = '<option value="">-- 請選擇欄位 --</option>'; 
        
        elements.searchInputMerged.value = '';
        elements.selectKeywordInputMerged.value = '';
        elements.selectKeywordRegexMerged.checked = false;

        toggleEditMode(false);
    }

    function renderMergedTable() {
        const table = document.createElement('table');
        const thead = table.createTHead();
        const headerRow = thead.insertRow();
        
        if (state.showSourceColumn) {
            const thSource = document.createElement('th');
            thSource.textContent = '來源檔案';
            thSource.classList.add('source-col');
            headerRow.prepend(thSource);
        }
        
        state.mergedHeaders.forEach(header => {
            const th = document.createElement('th');
            th.textContent = header;
            const deleteBtn = document.createElement('span');
            deleteBtn.className = 'delete-col-btn';
            deleteBtn.innerHTML = '&times;';
            deleteBtn.title = `刪除 ${header}`;
            deleteBtn.dataset.header = header;
            th.appendChild(deleteBtn);
            headerRow.appendChild(th); 
        });

        const tbody = table.createTBody();
        state.mergedData.forEach((rowData, index) => {
            const tr = tbody.insertRow();
            tr.dataset.rowIndex = index;
            
            if (rowData._sourceFile) {
                tr.title = `來源: ${rowData._sourceFile}`;
            }

            if (state.showSourceColumn) {
                const tdSource = document.createElement('td');
                tdSource.textContent = rowData._sourceFile || '';
                tdSource.classList.add('source-col');
                tr.prepend(tdSource);
            }

            if (rowData._isNew) tr.classList.add('new-row-highlight');
            
            state.mergedHeaders.forEach(header => {
                const td = tr.insertCell(); 
                td.contentEditable = state.isEditing;
                td.dataset.colHeader = header;
                const value = rowData[header] || '';
                td.textContent = value;
                if (value && !isNaN(parseFloat(value.replace(/,/g, '')))) {
                    td.classList.add('numeric');
                    td.textContent = formatNumber(value);
                } else if (!value) {
                    td.textContent = '';
                }
            });
        });

        if (state.showTotalRow) {
            const tfoot = table.createTFoot();
            const totalRow = tfoot.insertRow();
            totalRow.className = 'total-row';

            totalRow.innerHTML = ''; 
            
            totalRow.insertCell(); 
            
            if (state.showSourceColumn) {
                totalRow.insertCell().textContent = ''; // Empty cell for source
            }

            const totalsCache = calculateTotals();
            let totalLabelApplied = false;

            state.mergedHeaders.forEach(header => {
                const td = totalRow.insertCell();
                const totalVal = totalsCache[header];
                
                if (totalVal) {
                    td.textContent = formatNumber(totalVal);
                    td.classList.add('numeric');
                } else if (!totalLabelApplied) {
                    td.textContent = '合計';
                    totalLabelApplied = true;
                } else {
                    td.textContent = '';
                }
            });
        }
        
        elements.mergeViewContent.innerHTML = '';
        elements.mergeViewContent.appendChild(table);
        elements.mergeViewContent.classList.toggle('is-editing', state.isEditing);
        
        injectCheckboxes(elements.mergeViewContent); 

        const selectAllCheckbox = elements.mergeViewContent.querySelector('thead input[type="checkbox"]');
        if (selectAllCheckbox) {
            selectAllCheckbox.addEventListener('change', (e) => {
                const isChecked = e.target.checked;
                elements.mergeViewContent.querySelectorAll('.row-checkbox').forEach(cb => cb.checked = isChecked);
            });
        }
    }
    
    function toggleSourceColumn() {
        if (state.isEditing) {
            alert('請先儲存或取消編輯。');
            return;
        }
        state.showSourceColumn = !state.showSourceColumn;
        renderMergedTable(); // Re-render the table with/without the column
        
        // Update button text and style
        elements.toggleSourceColBtn.textContent = state.showSourceColumn ? '移除來源欄位' : '新增來源欄位';
        elements.toggleSourceColBtn.classList.toggle('active', state.showSourceColumn);
    }

    function updateColumnSelects(headers) { 
        // 1. Populate modal checklist
        elements.columnChecklist.innerHTML = headers.map(header => `<label><input type="checkbox" value="${header}" checked> ${header}</label>`).join('');
        
        // 2. Populate new operations dropdown
        elements.columnSelectOps.innerHTML = '<option value="">-- 請選擇欄位 --</option>';
        headers.forEach(header => {
            const option = document.createElement('option');
            option.value = header;
            option.textContent = header;
            elements.columnSelectOps.appendChild(option);
        });
    }

    function toggleColumnModal(forceShow) { elements.columnModal.classList.toggle('hidden', forceShow === false || !elements.columnModal.classList.contains('hidden')); }
    function setAllColumnCheckboxes(isChecked) { elements.columnChecklist.querySelectorAll('input').forEach(input => input.checked = isChecked); }
    function applyColumnChanges() {
        const mergedTable = elements.mergeViewContent.querySelector('table');
        if (!mergedTable) return;
        
        // 1. 從勾選清單中獲取可見性設定
        const visibility = {};
        elements.columnChecklist.querySelectorAll('input').forEach(input => {
            visibility[input.value] = input.checked;
        });

        // 2. 遍歷表格的標頭 (th)
        const allHeaders = Array.from(mergedTable.querySelectorAll('thead th'));
        const firstDataColIndex = allHeaders.findIndex(th => !th.classList.contains('checkbox-cell') && !th.classList.contains('source-col'));
        
        if (firstDataColIndex === -1) return; // No data headers found
        
        const dataHeaders = allHeaders.slice(firstDataColIndex);

        dataHeaders.forEach((th, dataIndex) => {
            const colIndex = dataIndex + firstDataColIndex;
            const headerText = th.textContent.replace('×', '').trim();
            const isVisible = visibility[headerText];
            
            // nth-child is 1-based, so +1
            mergedTable.querySelectorAll(`tr > *:nth-child(${colIndex + 1})`).forEach(cell => {
                cell.classList.toggle('column-hidden', !isVisible);
            });
        });
    }
    function handleMergedHeaderClick(th) {
        if (state.isEditing) return;
        if (th.classList.contains('source-col')) return; // Don't sort by source column
        
        const table = th.closest('table');
        const headerText = th.textContent.replace('×','').trim();
        const isAsc = th.classList.contains('sort-asc');
        table.querySelectorAll('th').forEach(h => h.classList.remove('sort-asc', 'sort-desc'));
        th.classList.add(isAsc ? 'sort-desc' : 'sort-asc');
        
        state.mergedData.sort((a, b) => {
            const valA = a[headerText] || '';
            const valB = b[headerText] || '';
            const numA = parseFloat(String(valA).replace(/,/g, ''));
            const numB = parseFloat(String(valB).replace(/,/g, ''));
            const comparison = (!isNaN(numA) && !isNaN(numB)) ? numA - numB : valA.localeCompare(valB, undefined, { numeric: true, sensitivity: 'base' });
            return isAsc ? -comparison : comparison;
        });
        renderMergedTable();
    }
    function deleteColumn(headerToDelete) {
        if (confirm(`確定要刪除「${headerToDelete}」這個欄位嗎？此操作無法復原。`)) {
            state.mergedHeaders = state.mergedHeaders.filter(h => h !== headerToDelete);
            state.mergedData.forEach(row => {
                delete row[headerToDelete];
            });
            renderMergedTable();
            updateColumnSelects(state.mergedHeaders); 
        }
    }
    function calculateTotals() {
        const totals = {};
        state.mergedHeaders.forEach(header => {
            const sum = state.mergedData.reduce((acc, row) => {
                const value = parseFloat(String(row[header]).replace(/,/g, ''));
                return acc + (isNaN(value) ? 0 : value);
            }, 0);
            if (sum !== 0 || state.mergedData.some(row => !isNaN(parseFloat(String(row[header]).replace(/,/g, ''))))) {
                totals[header] = sum;
            }
        });
        return totals;
    }

    // --- Edit Mode Functions ---
    function toggleEditMode(startEditing) {
        state.isEditing = startEditing;
        elements.editDataBtn.classList.toggle('hidden', state.isEditing);
        elements.saveEditsBtn.classList.toggle('hidden', !state.isEditing);
        elements.cancelEditsBtn.classList.toggle('hidden', !state.isEditing);
        elements.addNewRowBtn.disabled = state.isEditing;
        elements.copySelectedRowsBtn.disabled = state.isEditing;
        elements.deleteMergedRowsBtn.disabled = state.isEditing;
        elements.columnOperationsBtn.disabled = state.isEditing;
        elements.toggleTotalRowBtn.disabled = state.isEditing;
        elements.toggleSourceColBtn.disabled = state.isEditing; 
        elements.invertSelectionMergedBtn.disabled = state.isEditing; 
        elements.exportSelectedMergedXlsxBtn.disabled = state.isEditing; 
        elements.exportCurrentMergedXlsxBtn.disabled = state.isEditing; 
        elements.sortMergedByNameBtn.disabled = state.isEditing; // <-- DISABLE BUTTON
        elements.columnSelectOps.disabled = state.isEditing;
        elements.selectByColZeroBtn.disabled = state.isEditing;
        elements.selectByColEmptyBtn.disabled = state.isEditing;
        elements.selectByColExistsBtn.disabled = state.isEditing;
        
        elements.searchInputMerged.disabled = state.isEditing;
        elements.selectKeywordInputMerged.disabled = state.isEditing;
        elements.selectKeywordRegexMerged.disabled = state.isEditing;
        elements.selectByKeywordBtnMerged.disabled = state.isEditing;

        renderMergedTable();
    }
    function saveEdits() {
        const tableRows = elements.mergeViewContent.querySelectorAll('tbody tr');
        const newData = [];
        tableRows.forEach((tr, index) => {
            const newRowData = {};
            const sourceCell = tr.querySelector('.source-col');
            if (sourceCell) {
                newRowData._sourceFile = sourceCell.textContent;
            } else {
                const originalIndex = parseInt(tr.dataset.rowIndex, 10);
                if(!isNaN(originalIndex) && state.mergedData[originalIndex]) {
                     newRowData._sourceFile = state.mergedData[originalIndex]._sourceFile;
                } else {
                    newRowData._sourceFile = ' (新增資料列)';
                }
            }
            
            tr.querySelectorAll('td[data-col-header]').forEach(cell => {
                const header = cell.dataset.colHeader;
                newRowData[header] = cell.textContent;
            });
            newData.push(newRowData);
        });
        state.mergedData = newData;
        toggleEditMode(false);
    }
    function addNewRow() {
        const newRow = {};
        state.mergedHeaders.forEach(header => { newRow[header] = ''; });
        newRow._isNew = true;
        newRow._sourceFile = ' (新增資料列)'; // <-- Give it a source
        state.mergedData.unshift(newRow);
        if (!state.isEditing) {
            toggleEditMode(true);
        } else {
            renderMergedTable();
        }
    }
    function copySelectedRows() {
        const selectedCheckboxes = elements.mergeViewContent.querySelectorAll('.row-checkbox:checked');
        if (selectedCheckboxes.length === 0) {
            alert("請先勾選要複製的資料列。");
            return;
        }
        const rowsToCopy = [];
        selectedCheckboxes.forEach(cb => {
            const rowIndex = parseInt(cb.closest('tr').dataset.rowIndex, 10);
            if (!isNaN(rowIndex) && state.mergedData[rowIndex]) {
                const newRow = JSON.parse(JSON.stringify(state.mergedData[rowIndex]));
                newRow._isNew = true;
                newRow._sourceFile += ' (複製)'; // <-- Tag as copied
                rowsToCopy.push(newRow);
            }
        });
        state.mergedData.unshift(...rowsToCopy);
        if (!state.isEditing) {
            toggleEditMode(true);
        } else {
            renderMergedTable();
        }
    }

    // --- Row and Table Operations (Scope-aware) ---
    function syncCheckboxesInScope() {
        setTimeout(() => {
            const scope = state.isMergedView ? elements.mergeViewContent : elements.displayArea;
            scope.querySelectorAll('table').forEach(syncTableCheckboxState);
            if (!state.isMergedView) updateSelectionInfo();
        }, 0);
    }
    function selectAllRows() { const scope = state.isMergedView ? elements.mergeViewContent : elements.displayArea; const rows = scope.querySelectorAll('tbody tr:not(.row-hidden-search)'); if (rows.length === 0) { alert('沒有可勾選的列'); return; } rows.forEach(row => row.querySelector('.row-checkbox').checked = true); scope.querySelectorAll('thead input[type="checkbox"]').forEach(cb => cb.checked = true); }
    function invertSelection() { const scope = state.isMergedView ? elements.mergeViewContent : elements.displayArea; scope.querySelectorAll('tbody tr:not(.row-hidden-search) .row-checkbox').forEach(cb => cb.checked = !cb.checked); }
    function deleteSelectedRows() {
        const scope = state.isMergedView ? elements.mergeViewContent : elements.displayArea;
        const selected = scope.querySelectorAll('tbody .row-checkbox:checked');
        if (selected.length === 0) { alert('請先勾選要刪除的列'); return; }
        if (confirm(`確定要刪除 ${selected.length} 筆資料列嗎？`)) {
            if (state.isMergedView) {
                const indicesToDelete = new Set();
                selected.forEach(cb => {
                    const rowIndex = parseInt(cb.closest('tr').dataset.rowIndex, 10);
                    if (!isNaN(rowIndex)) indicesToDelete.add(rowIndex);
                });
                state.mergedData = state.mergedData.filter((_, index) => !indicesToDelete.has(index));
                renderMergedTable();
            } else {
                selected.forEach(cb => cb.closest('tr').remove());
            }
        }
        syncCheckboxesInScope();
    }
    function selectEmptyRows() { let count = 0; const scope = state.isMergedView ? elements.mergeViewContent : elements.displayArea; scope.querySelectorAll('tbody tr:not(.row-hidden-search)').forEach(row => { if (Array.from(row.cells).slice(1).every(c => c.textContent.trim() === '')) { row.querySelector('.row-checkbox').checked = true; count++; } }); if (count === 0) alert('未找到空白列'); }
    
    // ▼▼▼ MODIFIED: To read from correct input based on scope ▼▼▼
    function selectByKeyword() { 
        const inputEl = state.isMergedView ? elements.selectKeywordInputMerged : elements.selectKeywordInput;
        const regexEl = state.isMergedView ? elements.selectKeywordRegexMerged : elements.selectKeywordRegex;

        const keywordInput = inputEl.value.trim(); 
        const isRegex = regexEl.checked; 

        if (!keywordInput) { alert('請先輸入關鍵字'); return; } 
        
        let matchLogic; 
        try { 
            if (isRegex) { 
                const regex = new RegExp(keywordInput, 'i'); 
                matchLogic = text => regex.test(text); 
            } else if (keywordInput.includes(',')) { 
                const keywords = keywordInput.split(',').map(k => k.trim().toLowerCase()).filter(Boolean); 
                matchLogic = text => keywords.some(k => text.includes(k)); 
            } else { 
                const keywords = keywordInput.split(/\s+/).map(k => k.trim().toLowerCase()).filter(Boolean); 
                matchLogic = text => keywords.every(k => text.includes(k)); 
            } 
        } catch (e) { 
            alert('無效的 Regex 表示式：\n' + e.message); 
            return; 
        } 
        
        let count = 0; 
        const scope = state.isMergedView ? elements.mergeViewContent : elements.displayArea; 
        scope.querySelectorAll('tbody tr:not(.row-hidden-search)').forEach(row => { 
            // Also check sourceFile if it's visible
            let rowText = Array.from(row.querySelectorAll('td:not(.checkbox-cell)'))
                             .map(c => c.textContent)
                             .join(' ');
            
            if (matchLogic(rowText)) { 
                row.querySelector('.row-checkbox').checked = true; 
                count++; 
            } 
        }); 
        alert(count > 0 ? `已勾選 ${count} 個符合條件的列` : `未找到符合條件的列`); 
    }
    // ▲▲▲ MODIFIED ▲▲▲
    
    function selectByColumn(mode) {
        if (!state.isMergedView) {
            alert('此功能僅適用於「合併檢視」模式。\n請先點擊「合併檢視」按鈕。');
            return;
        }

        const colName = elements.columnSelectOps.value;
        if (!colName) {
            alert('請先從下拉選單中指定一個欄位。');
            return;
        }

        const scope = elements.mergeViewContent;
        let count = 0;

        let checkFunction;
        switch (mode) {
            case 'zero':   checkFunction = (val) => val === '0'; break;
            case 'empty':  checkFunction = (val) => val === ''; break;
            case 'exists': checkFunction = (val) => val !== ''; break;
            default: return;
        }

        scope.querySelectorAll('tbody tr:not(.row-hidden-search)').forEach(row => {
            const cell = row.querySelector(`td[data-col-header="${colName}"]`);
            const cellValue = cell ? cell.textContent.trim() : '';
            const checkbox = row.querySelector('.row-checkbox');
            
            if (checkbox && checkFunction(cellValue)) {
                checkbox.checked = true;
                count++;
            }
        });

        if (count > 0) {
            alert(`已勾選 ${count} 個符合條件的資料列。`);
        } else {
            alert('未找到符合條件的資料列。');
        }
    }

    // ▼▼▼ MODIFIED: To read from correct input based on scope ▼▼▼
    function filterTable() { 
        const inputEl = state.isMergedView ? elements.searchInputMerged : elements.searchInput;
        const keywords = inputEl.value.toLowerCase().trim().split(/\s+/).filter(Boolean); 
        
        const scope = state.isMergedView ? elements.mergeViewContent : elements.displayArea; 
        scope.querySelectorAll('tbody tr').forEach(row => { 
            // Also check sourceFile if it's visible
            let rowText = Array.from(row.querySelectorAll('td:not(.checkbox-cell)'))
                             .map(c => c.textContent)
                             .join(' ')
                             .toLowerCase();

            const isVisible = keywords.every(k => rowText.includes(k)); 
            row.classList.toggle('row-hidden-search', !isVisible); 
        }); 
        
        if (!state.isMergedView) { 
            elements.displayArea.querySelectorAll('.table-wrapper').forEach(wrapper => { 
                const visibleRowCount = wrapper.querySelectorAll('tbody tr:not(.row-hidden-search)').length; 
                wrapper.style.display = visibleRowCount > 0 ? '' : 'none'; 
            }); 
        } 
        syncCheckboxesInScope(); 
    }
    // ▲▲▲ MODIFIED ▲▲▲

    // --- Utility and Helper Functions ---
    function formatNumber(value) {
        const num = parseFloat(String(value).replace(/,/g, ''));
        if (isNaN(num)) return value;
        try {
            return new Intl.NumberFormat('en-US').format(num);
        } catch (e) {
            return num;
        }
    }
    function detectHiddenElements() { return elements.displayArea.querySelectorAll('tr[style*="display: none"], td[style*="display: none"], th[style*="display: none"]').length; }
    function showAllHiddenElements() { const hidden = elements.displayArea.querySelectorAll('tr[style*="display: none"], td[style*="display: none"], th[style*="display: none"]'); if (hidden.length === 0) { alert('沒有需要顯示的隱藏行列。'); return; } hidden.forEach(el => el.style.display = ''); alert(`已顯示 ${hidden.length} 個隱藏的行列。`); elements.showHiddenBtn.classList.add('hidden'); elements.loadStatusMessage.classList.add('hidden'); }
    function selectAllTables(isChecked) { elements.displayArea.querySelectorAll('.table-select-checkbox').forEach(cb => { if (cb.checked !== isChecked) { cb.click(); } }); }
    
    // ▼▼▼ CORRECTED: Use readAsBinaryString for this XLSX library version ▼▼▼
    function readFileAsBinary(file) { 
        return new Promise((resolve, reject) => { 
            const reader = new FileReader(); 
            reader.onload = e => resolve(e.target.result); 
            reader.onerror = reject; 
            reader.readAsBinaryString(file); // <-- This is the correct method
        }); 
    }
    // ▲▲▲ CORRECTED ▲▲▲

    function parsePositionString(str) { const indices = new Set(); const parts = str.split(',').map(p => p.trim()).filter(Boolean); for (const part of parts) { if (part.includes('-')) { const [start, end] = part.split('-').map(Number); if (!isNaN(start) && !isNaN(end) && start <= end) { for (let i = start; i <= end; i++) indices.add(i - 1); } } else { const num = Number(part); if (!isNaN(num)) indices.add(num - 1); } } return Array.from(indices).sort((a, b) => a - b); }
    async function getSelectedSheetNames(filename, workbook, mode, criteria) { const sheetNames = workbook.SheetNames; if (sheetNames.length === 0) return []; switch (mode) { case 'all': return sheetNames; case 'first': return sheetNames.length > 0 ? [sheetNames[0]] : []; case 'specific': return sheetNames.filter(name => name.toLowerCase().includes(criteria.name.toLowerCase())); case 'position': return parsePositionString(criteria.position).map(index => sheetNames[index]).filter(Boolean); case 'manual': return await showWorksheetSelectionModal(filename, sheetNames); default: return []; } }
    function showWorksheetSelectionModal(filename, sheetNames) { return new Promise(resolve => { if (sheetNames.length <= 1) { resolve(sheetNames); return; } const overlay = document.createElement('div'); overlay.className = 'modal-overlay'; const dialog = document.createElement('div'); dialog.className = 'modal-dialog'; dialog.innerHTML = `<div class="modal-header"><h3>選擇工作表 (手動模式)</h3><p>檔案 "<strong>${filename}</strong>"</p></div><div class="modal-body"><ul class="sheet-list">${sheetNames.map(name => `<li class="sheet-item"><label><input type="checkbox" class="sheet-checkbox" value="${name}" checked> ${name}</label></li>`).join('')}</ul></div><div class="modal-footer"><button class="btn btn-secondary" id="modal-skip">跳過</button><button class="btn btn-success" id="modal-confirm">確認</button></div>`; overlay.appendChild(dialog); document.body.appendChild(overlay); const closeModal = () => document.body.removeChild(overlay); dialog.querySelector('#modal-confirm').addEventListener('click', () => { resolve(Array.from(dialog.querySelectorAll('.sheet-checkbox')).filter(cb => cb.checked).map(cb => cb.value)); closeModal(); }); dialog.querySelector('#modal-skip').addEventListener('click', () => { resolve([]); closeModal(); }); }); }
    
    function extractTableData(table, { onlySelected = false, includeFilename = false, includeSourceCol = false } = {}) { 
        const data = []; 
        const headerRow = table.querySelector('thead tr'); 
        if (headerRow) { 
            // Get *visible* headers
            let headerData = Array.from(headerRow.querySelectorAll('th:not(.checkbox-cell):not(.column-hidden)'))
                                .map(th => th.textContent.replace('×','').trim()); 
            
            // This is for MAIN view export
            if (includeFilename) { 
                headerData.unshift('Source File'); 
            }
            data.push(headerData); 
        } 
        
        const filename = includeFilename ? (table.closest('.table-wrapper')?.querySelector('h4')?.textContent || 'Merged Table') : null; 
        
        // Determine which rows to process
        let rows;
        if (onlySelected === true) { // Explicitly true: only checked
            rows = Array.from(table.querySelectorAll('tbody .row-checkbox:checked')).map(cb => cb.closest('tr'));
        } else if (onlySelected === false) { // Explicitly false: only UNCHECKED (and not hidden by search)
             rows = Array.from(table.querySelectorAll('tbody tr:not(.row-hidden-search) .row-checkbox:not(:checked)')).map(cb => cb.closest('tr'));
        } else { // null or undefined: ALL (and not hidden by search)
            rows = table.querySelectorAll('tbody tr:not(.row-hidden-search)');
        }
            
        rows.forEach(row => { 
            let rowData = Array.from(row.querySelectorAll('td:not(.checkbox-cell):not(.column-hidden)'))
                             .map(td => td.textContent.trim()); 
            
            if (includeFilename) { 
                rowData.unshift(filename); 
            }
            data.push(rowData); 
        }); 
        return data; 
    }
    
    // --- ▼▼▼ NEW FUNCTION ▼▼▼ ---
    /**
     * Exports ALL visible data from the MERGED view.
     */
    function exportCurrentMergedXlsx() {
        if (!state.isMergedView) {
            alert('此功能僅限合併檢視模式使用。');
            return;
        }

        const table = elements.mergeViewContent.querySelector('table');
        if (!table) {
            alert('找不到合併表格。');
            return;
        }
        
        // 1. Get visible headers
        const headerData = extractTableData(table, { 
            onlySelected: null, // We pass null to get ALL headers
            includeSourceCol: state.showSourceColumn 
        })[0]; // Get only the header row
        
        if (!headerData) {
            alert('無法讀取表頭。');
            return;
        }

        // 2. Get data for *all* rows that are *not hidden by search*
        const data = extractTableData(table, {
            onlySelected: null, // null means get ALL visible
            includeSourceCol: state.showSourceColumn
        }).slice(1); // Get data rows (slice(1) skips header)
        
        if (data.length === 0) {
            alert('沒有可見的資料列可匯出。');
            return;
        }

        // 3. Create and download workbook
        try {
            const ws_data = [headerData].concat(data);
            const ws = XLSX.utils.aoa_to_sheet(ws_data);
            ws['!cols'] = calculateColumnWidths(ws_data);
            const wb = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(wb, ws, "Merged View Data");
            XLSX.writeFile(wb, `merged_view_export_${new Date().toISOString().slice(0, 10)}.xlsx`);
        } catch (err) {
            console.error('匯出目前合併檢視時發生錯誤:', err);
            alert('匯出時發生錯誤：' + err.message);
        }
    }
    // --- ▲▲▲ NEW FUNCTION ▲▲▲ ---

    function exportSelectedMergedXlsx() {
        if (!state.isMergedView) {
            alert('此功能僅限合併檢視模式使用。');
            return;
        }

        const table = elements.mergeViewContent.querySelector('table');
        if (!table) {
            alert('找不到合併表格。');
            return;
        }
        
        // 1. Get visible headers
        const headerData = extractTableData(table, { 
            onlySelected: null, // We pass null to get ALL headers
            includeSourceCol: state.showSourceColumn 
        })[0]; // Get only the header row
        
        if (!headerData) {
            alert('無法讀取表頭。');
            return;
        }

        // 2. Get data for *unselected* rows that are *not hidden by search*
        const data = extractTableData(table, {
            onlySelected: false, // false means get UNCHECKED
            includeSourceCol: state.showSourceColumn
        }).slice(1); // Get data rows (slice(1) skips header)
        
        if (data.length === 0) {
            alert('沒有剩餘的(未勾選)資料列可匯出。\n(注意：搜尋結果外的資料列不會被匯出)');
            return;
        }

        // 3. Create and download workbook
        try {
            const ws_data = [headerData].concat(data);
            const ws = XLSX.utils.aoa_to_sheet(ws_data);
            ws['!cols'] = calculateColumnWidths(ws_data);
            const wb = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(wb, ws, "Merged Remaining Data");
            XLSX.writeFile(wb, `merged_remaining_export_${new Date().toISOString().slice(0, 10)}.xlsx`);
        } catch (err) {
            console.error('匯出剩餘的合併資料時發生錯誤:', err);
            alert('匯出時發生錯誤：' + err.message);
        }
    }

    function downloadHtml(content, filename) { const blob = new Blob([content], { type: 'text/html;charset=utf-8;' }); const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = filename; a.click(); URL.revokeObjectURL(a.href); }
    
    function exportMergedXlsx() { const tables = Array.from(elements.displayArea.querySelectorAll('.table-wrapper:not([style*="display: none"]) table')); if (tables.length === 0) { alert('沒有可匯出的表格。'); return; } try { const allData = []; tables.forEach((table, i) => { const data = extractTableData(table, { includeFilename: true, onlySelected: null }); if (data.length > 1) allData.push(...(i === 0 ? data : data.slice(1))); }); if (allData.length <= 1) { alert('沒有足夠的資料可以匯出。'); return; } const ws = XLSX.utils.aoa_to_sheet(allData); ws['!cols'] = calculateColumnWidths(allData); const workbook = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(workbook, ws, 'Merged Data'); XLSX.writeFile(workbook, `report_merged_${new Date().toISOString().slice(0, 10)}.xlsx`); alert(`成功合併 ${tables.length} 個表格，共 ${allData.length - 1} 筆資料。`); } catch (err) { console.error('合併匯出 XLSX 錯誤:', err); alert('合併匯出 XLSX 時發生錯誤：' + err.message); } }
    
    function calculateColumnWidths(data) { if (data.length === 0) return []; return data[0].map((_, col) => ({ wch: Math.min(50, Math.max(10, ...data.map(row => row[col] ? String(row[col]).length : 0)) + 2) })); }
    
    // --- State Management and UI Updates ---
    function resetView() { 
        if(state.isMergedView) closeMergeView(); 
        if (!state.originalHtmlString) return; 
        elements.displayArea.innerHTML = state.originalHtmlString; 
        injectCheckboxes(elements.displayArea); 
        ['searchInput', 'selectKeywordInput'].forEach(id => elements[id].value = ''); 
        elements.selectKeywordRegex.checked = false; 
        filterTable(); 
        elements.loadStatusMessage.classList.add('hidden'); 
        const hiddenCount = detectHiddenElements(); 
        if (hiddenCount > 0) { 
            elements.loadStatusMessage.textContent = `注意：已重設表格，${hiddenCount} 個隱藏的行列已還原。`; 
            elements.loadStatusMessage.classList.remove('hidden'); 
            elements.showHiddenBtn.classList.remove('hidden'); 
        } else { 
            elements.showHiddenBtn.classList.add('hidden'); 
        } 
        updateSelectionInfo(); 
        setViewMode('list'); 
    }
    
    function resetControls(isNewFile) { if (!isNewFile) return; if(state.isMergedView) closeMergeView(); state.originalHtmlString = ''; ['searchInput', 'selectKeywordInput'].forEach(id => elements[id].value = ''); elements.selectKeywordRegex.checked = false; elements.controlPanel.classList.add('hidden'); updateSelectionInfo(); }
    
    function clearAllFiles(silent = false) { 
        if (!silent && !confirm('確定要清除所有已載入的檔案嗎？')) return; 
        if(state.isMergedView) closeMergeView(); 
        state.originalHtmlString = ''; 
        state.loadedFiles = []; 
        state.loadedTables = 0; 
        elements.displayArea.innerHTML = ''; 
        elements.fileInput.value = ''; 
        ['specificSheetNameInput', 'specificSheetPositionInput'].forEach(id => elements[id].value = ''); 
        elements.gridScaleSlider.value = 3; 
        elements.columnSelectOps.innerHTML = '<option value="">-- 請選擇欄位 --</option>'; 
        updateGridScale(); 
        updateDropAreaDisplay(); 
        resetControls(true); 
        setViewMode('list'); 
    }
    
    function updateDropAreaDisplay() { const hasFiles = state.loadedTables > 0; elements.dropArea.classList.toggle('compact', hasFiles); elements.dropAreaInitial.classList.toggle('hidden', hasFiles); elements.dropAreaLoaded.classList.toggle('hidden', !hasFiles); elements.importOptionsContainer.classList.toggle('hidden', hasFiles); if (hasFiles) { elements.fileCount.textContent = state.loadedTables; const names = state.loadedFiles.slice(0, 3).join(', '); const more = state.loadedFiles.length > 3 ? ` 及其他 ${state.loadedFiles.length - 3} 個...` : ''; elements.fileNames.textContent = names + more; } }
    
    function showControls(hiddenCount) {
        elements.controlPanel.classList.remove('hidden');
        const buttonsToShow = [
            'selectByKeywordGroup', 'selectByKeywordBtn', 'selectEmptyBtn', 'deleteSelectedBtn', 
            'invertSelectionBtn', 
            'selectAllBtn', 
            'exportMergedXlsxBtn', 'resetViewBtn', 
            'tableLevelControls', 'listViewBtn', 'gridViewBtn', 'showHiddenBtn',
            'viewCheckedCombinedBtn', // New
            'sortByNameBtn' // New
        ];
        buttonsToShow.forEach(id => {
            if(elements[id]) elements[id].classList.remove('hidden');
        });

        elements.mergeViewBtn.classList.toggle('hidden', state.loadedTables <= 1);
        const showHiddenStuff = hiddenCount > 0;
        elements.loadStatusMessage.classList.toggle('hidden', !showHiddenStuff);
        elements.showHiddenBtn.classList.toggle('hidden', !showHiddenStuff);
        if (showHiddenStuff) {
            elements.loadStatusMessage.textContent = `注意：檔案中包含 ${hiddenCount} 個被隱藏的行列。`;
        }
    }
    function setViewMode(mode) { if (mode === 'grid') { elements.displayArea.classList.remove('list-view'); elements.displayArea.classList.add('grid-view'); elements.gridViewBtn.classList.add('active'); elements.listViewBtn.classList.remove('active'); elements.gridScaleControl.classList.remove('hidden'); } else { elements.displayArea.classList.remove('grid-view'); elements.displayArea.classList.add('list-view'); elements.listViewBtn.classList.add('active'); elements.gridViewBtn.classList.remove('active'); elements.gridScaleControl.classList.add('hidden'); } }
    function updateGridScale() { const columns = elements.gridScaleSlider.value; elements.displayArea.style.setProperty('--grid-columns', columns); }
    function validateFiles(fileList) { if (!fileList || fileList.length === 0) return { valid: false, error: '沒有選擇檔案' }; const validFiles = Array.from(fileList).filter(file => CONSTANTS.VALID_FILE_EXTENSIONS.some(ext => file.name.toLowerCase().endsWith(ext))); if (validFiles.length === 0) return { valid: false, error: '請上傳 .xls 或 .xlsx 格式的檔案' }; return { valid: true, files: validFiles }; }
    function debounce(func, wait) { let timeout; return (...args) => { clearTimeout(timeout); timeout = setTimeout(() => func(...args), wait); }; }
    function handleScroll() { if (window.scrollY > window.innerHeight / 2) { elements.backToTopBtn.classList.add('visible'); } else { elements.backToTopBtn.classList.remove('visible'); } }
    function scrollToTop() { window.scrollTo({ top: 0, behavior: 'smooth' }); }
    function handleCardClick(e) { const card = e.target.closest('.table-wrapper'); if (!card) return; if (e.target.classList.contains('close-zoom')) { closePreview(); return; } if (e.target.classList.contains('delete-rows-btn')) { const scope = e.target.closest('.table-wrapper'); deleteSelectedRowsInScope(scope); return; } if (e.target.classList.contains('delete-table-btn')) { if (confirm(`確定要永久刪除此工作表 (${card.querySelector('h4').textContent}) 嗎？`)) { closePreview(); setTimeout(() => { card.remove(); updateFileStateAfterDeletion(); }, 300); } return; } if (elements.displayArea.classList.contains('grid-view') && !card.classList.contains('is-zoomed')) { if (!e.target.matches('input, a, button, .btn')) { openPreview(card); } } }
    function deleteSelectedRowsInScope(scope) { const selected = scope.querySelectorAll('tbody .row-checkbox:checked'); if (selected.length === 0) { alert('請先勾選要刪除的列'); return; } if (confirm(`確定要在此表格內刪除 ${selected.length} 筆資料列嗎？`)) { selected.forEach(cb => cb.closest('tr').remove()); } syncCheckboxesInScope(); }
    function openPreview(card) { if (state.zoomedCard) return; card.classList.add('is-zoomed'); state.zoomedCard = card; document.body.classList.add('no-scroll'); }
    function closePreview() { if (!state.zoomedCard) return; state.zoomedCard.classList.remove('is-zoomed'); state.zoomedCard = null; document.body.classList.remove('no-scroll'); }
    function syncTableCheckboxState(table) { const wrapper = table.closest('.table-wrapper'); const headerCheckbox = wrapper ? wrapper.querySelector('.table-select-checkbox') : table.querySelector('thead input[type="checkbox"]'); if (!headerCheckbox) return; const rowCheckboxes = table.querySelectorAll('tbody tr:not(.row-hidden-search) .row-checkbox'); if (rowCheckboxes.length === 0) { headerCheckbox.checked = false; headerCheckbox.indeterminate = false; return; } const checkedCount = Array.from(rowCheckboxes).filter(cb => cb.checked).length; if (checkedCount === 0) { headerCheckbox.checked = false; headerCheckbox.indeterminate = false; } else if (checkedCount === rowCheckboxes.length) { headerCheckbox.checked = true; headerCheckbox.indeterminate = false; } else { headerCheckbox.checked = false; headerCheckbox.indeterminate = true; } }
    function updateSelectionInfo() { const selectedCheckboxes = elements.displayArea.querySelectorAll('.table-select-checkbox:checked, .table-select-checkbox:indeterminate'); if (selectedCheckboxes.length > 0) { const names = Array.from(selectedCheckboxes).map(cb => cb.closest('.table-header').querySelector('h4').textContent); elements.selectedTablesList.textContent = names.join('; '); elements.selectedTablesInfo.classList.remove('hidden'); } else { elements.selectedTablesInfo.classList.add('hidden'); } }
    function deleteSelectedTables() { const selectedWrappers = Array.from(elements.displayArea.querySelectorAll('.table-select-checkbox:checked')).map(cb => cb.closest('.table-wrapper')); if (selectedWrappers.length === 0) { alert('請先勾選要刪除的表格。'); return; } if (confirm(`確定要永久刪除 ${selectedWrappers.length} 個選定的表格嗎？`)) { selectedWrappers.forEach(wrapper => wrapper.remove()); updateFileStateAfterDeletion(); } }
    function updateFileStateAfterDeletion() { const remainingWrappers = elements.displayArea.querySelectorAll('.table-wrapper'); state.loadedTables = remainingWrappers.length; state.loadedFiles = Array.from(remainingWrappers).map(w => w.querySelector('h4').textContent); if (state.loadedTables === 0) { clearAllFiles(true); } else { updateDropAreaDisplay(); showControls(detectHiddenElements()); } }
    function handleDisplayAreaChange(e) { const target = e.target; if (!target.matches('.table-select-checkbox, [id^="select-all-checkbox"], .row-checkbox')) return; let table; if (target.matches('.table-select-checkbox')) { const wrapper = target.closest('.table-wrapper'); table = wrapper ? wrapper.querySelector('table') : null; if (table) { toggleSelectAll(target.checked, table); } } else if (target.matches('[id^="select-all-checkbox"]')) { table = target.closest('table'); toggleSelectAll(target.checked, table); } else { table = target.closest('table'); } if (table) { syncTableCheckboxState(table); } if (!state.isMergedView) { updateSelectionInfo(); } }
    function toggleSelectAll(isChecked, table) { if (!table) return; table.querySelectorAll('tbody tr:not(.row-hidden-search) .row-checkbox').forEach(cb => cb.checked = isChecked); }
    
    // ▼▼▼ HELPER FUNCTION (MOVED) ▼▼▼
    function getFundSortPriority(fileName) {
        if (state.fundSortOrder.length === 0) return { index: Infinity, name: fileName };
        
        // 1. 根據檔名 (別名) 找到標準名稱
        // state.fundAliasKeys 已經依長度排序 (長的優先)
        const foundAlias = state.fundAliasKeys.find(alias => fileName.includes(alias));
        const canonicalName = foundAlias ? state.fundAliasMap[foundAlias] : null;

        // 2. 從標準名稱找到排序順序
        const index = canonicalName ? state.fundSortOrder.indexOf(canonicalName) : -1;

        // 3. 決定優先級 (找不到的排最後)
        const priority = (index === -1) ? Infinity : index;

        return { index: priority, name: fileName };
    }
    // ▲▲▲ HELPER FUNCTION (MOVED) ▲▲▲


    // ▼▼▼ MODIFIED FUNCTION (Using new state properties) ▼▼▼
    function sortTablesByFundName() {
        if (state.fundSortOrder.length === 0 || Object.keys(state.fundAliasMap).length === 0) {
            // 由於 init() 是 async，在 loadFundConfig() 完成前，renderTables() 可能會先執行
            // 所以這裡不要 alert，而是靜默失敗
            console.warn('基金順序列表尚未載入，暫不執行自動排序。');
            return;
        }

        const wrappers = Array.from(elements.displayArea.querySelectorAll('.table-wrapper'));

        // Helper: 從卡片標題 "檔名 (工作表名稱)" 中提取 "檔名"
        const getFileName = (wrapper) => {
            const text = wrapper.querySelector('h4').textContent;
            const match = text.match(/(.*)\s\(.*\)$/); // 抓取開頭到第一個 ( 之前的內容
            return match ? match[1].trim() : text.trim(); // 找不到括號就回傳完整名稱
        };

        wrappers.sort((a, b) => {
            const fileA = getFundSortPriority(getFileName(a));
            const fileB = getFundSortPriority(getFileName(b));

            if (fileA.index === fileB.index) {
                // 如果順序相同 (例如兩個都找不到)，則維持字母/筆劃順序
                return fileA.name.localeCompare(fileB.name);
            }
            
            return fileA.index - fileB.index;
        });

        // Re-append sorted wrappers to the display area
        elements.displayArea.innerHTML = '';
        wrappers.forEach(wrapper => {
            elements.displayArea.appendChild(wrapper);
        });
    }
    // ▲▲▲ MODIFIED FUNCTION ▲▲▲

    // --- ▼▼▼ NEW FUNCTION ▼▼▼ ---
    /**
     * Sorts the data currently in the MERGED VIEW by fund name
     */
    function sortMergedTableByFundName() {
        if (state.isEditing) {
            alert('請先儲存或取消編輯。');
            return;
        }
        if (state.fundSortOrder.length === 0 || Object.keys(state.fundAliasMap).length === 0) {
            alert('錯誤：基金順序列表尚未載入或為空。\n請檢查 fund-config.json 檔案。');
            return;
        }

        state.mergedData.sort((a, b) => {
            // 1. Get the source filename from the row data
            const fileNameA = a._sourceFile || '';
            const fileNameB = b._sourceFile || '';

            // 2. Get sort priority
            const fileA = getFundSortPriority(fileNameA);
            const fileB = getFundSortPriority(fileNameB);
            
            if (fileA.index === fileB.index) {
                // 如果順序相同 (例如兩個都找不到)，則維持字母/筆劃順序
                return fileNameA.localeCompare(fileNameB);
            }
            
            return fileA.index - fileB.index;
        });

        // Re-render the merged table
        renderMergedTable();
    }
    // --- ▲▲▲ NEW FUNCTION ▲▲▲ ---


    return { init };
})();

ExcelViewer.init();
