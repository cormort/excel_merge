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
        isEditing: false, // For merge view edit mode
        mergedData: [], // To store data for the merged table
        mergedHeaders: []
    };
    const elements = {};

    function init() {
        cacheElements();
        bindEvents();
    }

    function cacheElements() {
        const mapping = {
            fileInput: 'file-input', displayArea: 'excel-display-area', searchInput: 'search-input',
            dropArea: 'drop-area', deleteSelectedBtn: 'delete-selected-btn', invertSelectionBtn: 'invert-selection-btn',
            resetViewBtn: 'reset-view-btn', selectEmptyBtn: 'select-empty-btn', exportHtmlBtn: 'export-html-btn',
            showHiddenBtn: 'show-hidden-btn', exportSelectedBtn: 'export-selected-btn', exportXlsxBtn: 'export-xlsx-btn',
            exportSelectedXlsxBtn: 'export-selected-xlsx-btn', exportMergedXlsxBtn: 'export-merged-xlsx-btn',
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
            tableLevelControls: 'table-level-controls', selectedTablesInfo: 'selected-tables-info',
            selectedTablesList: 'selected-tables-list', listViewBtn: 'list-view-btn',
            gridViewBtn: 'grid-view-btn', backToTopBtn: 'back-to-top-btn',
            gridScaleControl: 'grid-scale-control', gridScaleSlider: 'grid-scale-slider',
            // Merge view
            mergeViewBtn: 'merge-view-btn',
            backToMultiViewBtn: 'back-to-multi-view-btn',
            mergedDisplayArea: 'merged-display-area',
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
            copySelectedRowsBtn: 'copy-selected-rows-btn'
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
        
        // --- Row Operations ---
        elements.selectByKeywordBtn.addEventListener('click', () => { selectByKeyword(); syncCheckboxesInScope(); });
        elements.selectEmptyBtn.addEventListener('click', () => { selectEmptyRows(); syncCheckboxesInScope(); });
        elements.selectAllBtn.addEventListener('click', () => { selectAllRows(); syncCheckboxesInScope(); });
        elements.invertSelectionBtn.addEventListener('click', () => { invertSelection(); syncCheckboxesInScope(); });
        elements.deleteSelectedBtn.addEventListener('click', deleteSelectedRows);

        // --- Global, Export, and Merge Operations ---
        elements.resetViewBtn.addEventListener('click', resetView);
        elements.showHiddenBtn.addEventListener('click', showAllHiddenElements);
        elements.exportHtmlBtn.addEventListener('click', () => exportHtml('all'));
        elements.exportSelectedBtn.addEventListener('click', () => exportHtml('selected'));
        elements.exportXlsxBtn.addEventListener('click', () => exportXlsx('all'));
        elements.exportSelectedXlsxBtn.addEventListener('click', () => exportXlsx('selected'));
        elements.exportMergedXlsxBtn.addEventListener('click', exportMergedXlsx);
        elements.mergeViewBtn.addEventListener('click', createMergedView);
        elements.backToMultiViewBtn.addEventListener('click', showMultiTableView);
        elements.columnOperationsBtn.addEventListener('click', () => toggleColumnModal(true));
        elements.closeColumnModalBtn.addEventListener('click', () => toggleColumnModal(false));
        elements.applyColumnChangesBtn.addEventListener('click', () => { applyColumnChanges(); toggleColumnModal(false); });
        elements.modalCheckAll.addEventListener('click', () => setAllColumnCheckboxes(true));
        elements.modalUncheckAll.addEventListener('click', () => setAllColumnCheckboxes(false));
        
        // --- Edit Operations ---
        elements.editDataBtn.addEventListener('click', () => toggleEditMode(true));
        elements.saveEditsBtn.addEventListener('click', saveEdits);
        elements.cancelEditsBtn.addEventListener('click', () => toggleEditMode(false));
        elements.addNewRowBtn.addEventListener('click', addNewRow);
        elements.copySelectedRowsBtn.addEventListener('click', copySelectedRows);

        // --- Input and Dynamic Content Handling ---
        elements.searchInput.addEventListener('input', debounce(filterTable, 300));
        elements.displayArea.addEventListener('change', handleDisplayAreaChange);
        elements.displayArea.addEventListener('click', handleCardClick);
        elements.mergedDisplayArea.addEventListener('click', e => {
            if (e.target.matches('th:not(.checkbox-cell)')) {
                handleMergedHeaderClick(e.target);
            }
        });
        elements.importOptionsContainer.addEventListener('change', e => {
            if (e.target.name === 'import-mode') {
                const selectedMode = e.target.value;
                elements.specificSheetNameGroup.classList.toggle('hidden', selectedMode !== 'specific');
                elements.specificSheetPositionGroup.classList.toggle('hidden', selectedMode !== 'position');
            }
        });

        // --- Window/Document Level Events ---
        elements.backToTopBtn.addEventListener('click', scrollToTop);
        window.addEventListener('scroll', handleScroll);
        document.addEventListener('keydown', e => { if (e.key === 'Escape') { if(state.zoomedCard) closePreview(); if(!elements.columnModal.classList.contains('hidden')) toggleColumnModal(false); } });
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
    async function processFiles(fileList) { const validation = validateFiles(fileList); if (!validation.valid) { alert(`錯誤：${validation.error}`); return; } if (state.isProcessing) { alert('正在處理檔案...'); return; } const importMode = document.querySelector('input[name="import-mode"]:checked').value; const specificSheetName = elements.specificSheetNameInput.value.trim(); const specificSheetPosition = elements.specificSheetPositionInput.value.trim(); if (importMode === 'specific' && !specificSheetName) { alert('請輸入工作表名稱！'); return; } if (importMode === 'position' && !specificSheetPosition) { alert('請輸入工作表位置！'); return; } state.isProcessing = true; elements.displayArea.innerHTML = '<div class="loading">讀取中...</div>'; resetControls(true); const tablesToRender = []; const missedFiles = []; state.loadedFiles = []; try { for (let index = 0; index < validation.files.length; index++) { const file = validation.files[index]; elements.displayArea.innerHTML = `<div class="loading">讀取中... (${index + 1}/${validation.files.length})</div>`; const binaryData = await readFileAsBinary(file); const workbook = XLSX.read(binaryData, { type: 'binary' }); const sheetNames = await getSelectedSheetNames(file.name, workbook, importMode, { name: specificSheetName, position: specificSheetPosition }); if ((importMode === 'specific' || importMode === 'position') && sheetNames.length === 0 && workbook.SheetNames.length > 0) { missedFiles.push(file.name); } for (const sheetName of sheetNames) { const sheet = workbook.Sheets[sheetName]; const htmlString = XLSX.utils.sheet_to_html(sheet); tablesToRender.push({ html: htmlString, filename: `${file.name} (${sheetName})` }); state.loadedFiles.push(`${file.name} (${sheetName})`); } } if (missedFiles.length > 0) { const criteria = importMode === 'specific' ? `名稱包含 "${specificSheetName}"` : `位置符合 "${specificSheetPosition}"`; alert(`以下檔案找不到 ${criteria} 的工作表：\n\n- ${missedFiles.join('\n- ')}`); } state.loadedTables = tablesToRender.length; renderTables(tablesToRender); updateDropAreaDisplay(); } catch (err) { console.error("處理檔案時發生錯誤:", err); elements.displayArea.innerHTML = `<p style="color: red;">處理檔案錯誤：${err.message || '未知錯誤'}</p>`; resetControls(true); } finally { state.isProcessing = false; } }
    function renderTables(tablesToRender) { if (tablesToRender.length === 0) { elements.displayArea.innerHTML = `<p>沒有找到符合條件的工作表。</p>`; return; } const fragment = document.createDocumentFragment(); tablesToRender.forEach(({ html, filename }) => { const wrapper = document.createElement('div'); wrapper.className = 'table-wrapper'; const header = document.createElement('div'); header.className = 'table-header'; header.innerHTML = `<input type="checkbox" class="table-select-checkbox" title="選取此表格"><h4>${filename}</h4><div class="header-actions"><button class="btn btn-danger btn-sm delete-rows-btn">刪除選取列</button><button class="btn btn-danger btn-sm delete-table-btn">刪除此表</button></div><button class="close-zoom">&times;</button>`; const tableContent = document.createElement('div'); tableContent.className = 'table-content'; const tempDiv = document.createElement('div'); tempDiv.innerHTML = html; const table = tempDiv.querySelector('table'); if (table) { tableContent.appendChild(table); wrapper.appendChild(header); wrapper.appendChild(tableContent); fragment.appendChild(wrapper); } }); elements.displayArea.innerHTML = ''; elements.displayArea.appendChild(fragment); state.originalHtmlString = elements.displayArea.innerHTML; injectCheckboxes(elements.displayArea); showControls(detectHiddenElements()); }
    function injectCheckboxes(scope) { scope.querySelectorAll('thead tr').forEach((headRow, index) => { if(headRow.querySelector('.checkbox-cell')) return; const th = document.createElement('th'); th.innerHTML = `<input type="checkbox" id="select-all-checkbox-${scope.id}-${index}" title="全選/全不選">`; th.classList.add('checkbox-cell'); headRow.prepend(th); }); scope.querySelectorAll('tbody tr').forEach(row => { if(row.querySelector('.checkbox-cell')) return; const td = document.createElement('td'); td.innerHTML = '<input type="checkbox" class="row-checkbox">'; td.classList.add('checkbox-cell'); row.prepend(td); }); }
    
    // --- Merged View and Column Operations ---
    function createMergedView() {
        const tables = Array.from(elements.displayArea.querySelectorAll('.table-wrapper:not([style*="display: none"]) table'));
        if (tables.length === 0) { alert('沒有可合併的表格。'); return; }

        const allHeaders = new Set();
        const tableData = tables.map(table => {
            const headers = Array.from(table.querySelectorAll('thead th:not(.checkbox-cell)')).map(th => th.textContent.trim());
            headers.forEach(h => allHeaders.add(h));
            return Array.from(table.querySelectorAll('tbody tr')).map(row => {
                const rowData = {};
                Array.from(row.querySelectorAll('td:not(.checkbox-cell)')).forEach((td, i) => {
                    rowData[headers[i]] = td.textContent; // Store text content
                });
                return rowData;
            });
        }).flat();

        state.mergedHeaders = Array.from(allHeaders);
        state.mergedData = tableData;
        
        renderMergedTable();
        toggleViewState(true);
        populateColumnModal(state.mergedHeaders);
    }

    function renderMergedTable() {
        let html = '<table><thead><tr>';
        state.mergedHeaders.forEach(header => html += `<th>${header}</th>`);
        html += '</tr></thead><tbody>';

        state.mergedData.forEach((rowData, index) => {
            html += `<tr data-row-index="${index}" class="${rowData._isNew ? 'new-row-highlight' : ''}">`;
            state.mergedHeaders.forEach(header => {
                html += `<td contenteditable="${state.isEditing}" data-col-header="${header}">${rowData[header] || ''}</td>`;
            });
            html += '</tr>';
        });
        html += '</tbody></table>';
        
        elements.mergedDisplayArea.innerHTML = html;
        elements.mergedDisplayArea.classList.toggle('is-editing', state.isEditing);
        injectCheckboxes(elements.mergedDisplayArea);
    }

    function toggleViewState(isMerged) {
        state.isMergedView = isMerged;
        elements.displayArea.classList.toggle('hidden', isMerged);
        elements.mergedDisplayArea.classList.toggle('hidden', !isMerged);
        
        // Toggle main merge buttons
        elements.mergeViewBtn.classList.toggle('hidden', isMerged);
        elements.backToMultiViewBtn.classList.toggle('hidden', !isMerged);
        elements.columnOperationsBtn.classList.toggle('hidden', !isMerged);
        
        // Toggle standard row operation buttons
        const standardRowButtons = ['select-by-keyword-btn', 'select-empty-btn', 'select-all-btn', 'invert-selection-btn'];
        standardRowButtons.forEach(id => elements[id].classList.toggle('hidden', isMerged && state.isEditing));

        // Toggle edit buttons
        const editButtons = ['edit-data-btn', 'add-new-row-btn', 'copy-selected-rows-btn'];
        editButtons.forEach(id => elements[id].classList.toggle('hidden', !isMerged));

        // Hide controls not applicable in merged view
        [elements.listViewBtn, elements.gridViewBtn, elements.tableLevelControls, elements.exportMergedXlsxBtn].forEach(el => el.classList.toggle('hidden', isMerged));
        
        if (!isMerged) {
            toggleEditMode(false); // Ensure edit mode is off when leaving
        }
    }
    function showMultiTableView() {
        if (state.isEditing) {
            if (!confirm("您有未儲存的編輯，確定要返回並捨棄變更嗎？")) {
                return;
            }
        }
        toggleViewState(false);
        state.mergedData = [];
        state.mergedHeaders = [];
        elements.mergedDisplayArea.innerHTML = '';
        filterTable(); // Re-apply filter to multi-view
    }
    function populateColumnModal(headers) { elements.columnChecklist.innerHTML = headers.map(header => `<label><input type="checkbox" value="${header}" checked> ${header}</label>`).join(''); }
    function toggleColumnModal(forceShow) { elements.columnModal.classList.toggle('hidden', forceShow === false || !elements.columnModal.classList.contains('hidden')); }
    function setAllColumnCheckboxes(isChecked) { elements.columnChecklist.querySelectorAll('input').forEach(input => input.checked = isChecked); }
    function applyColumnChanges() {
        const mergedTable = elements.mergedDisplayArea.querySelector('table');
        if (!mergedTable) return;
        const visibility = {};
        elements.columnChecklist.querySelectorAll('input').forEach(input => { visibility[input.value] = input.checked; });
        const headers = Array.from(mergedTable.querySelectorAll('thead th:not(.checkbox-cell)'));
        headers.forEach((th, index) => {
            const isVisible = visibility[th.textContent.trim()];
            const colIndex = index + 1; // +1 for checkbox cell
            mergedTable.querySelectorAll(`tr > *:nth-child(${colIndex})`).forEach(cell => { cell.classList.toggle('column-hidden', !isVisible); });
        });
    }
    function handleMergedHeaderClick(th) {
        if (state.isEditing) return; // Disable sorting in edit mode
        const table = th.closest('table'), tbody = table.querySelector('tbody'), rows = Array.from(tbody.querySelectorAll('tr'));
        const colIndex = Array.from(th.parentNode.children).indexOf(th);
        const isAsc = th.classList.contains('sort-asc');
        table.querySelectorAll('th').forEach(h => h.classList.remove('sort-asc', 'sort-desc'));
        th.classList.add(isAsc ? 'sort-desc' : 'sort-asc');
        rows.sort((a, b) => {
            const valA = a.cells[colIndex]?.textContent.trim() || '';
            const valB = b.cells[colIndex]?.textContent.trim() || '';
            const numA = parseFloat(valA), numB = parseFloat(valB);
            const comparison = (!isNaN(numA) && !isNaN(numB)) ? numA - numB : valA.localeCompare(valB, undefined, { numeric: true, sensitivity: 'base' });
            return isAsc ? -comparison : comparison;
        });
        tbody.append(...rows);
    }

    // --- Edit Mode Functions ---
    function toggleEditMode(startEditing) {
        if (startEditing && !state.isMergedView) return;
        state.isEditing = startEditing;

        // Toggle button visibility
        elements.editDataBtn.classList.toggle('hidden', state.isEditing);
        elements.saveEditsBtn.classList.toggle('hidden', !state.isEditing);
        elements.cancelEditsBtn.classList.toggle('hidden', !state.isEditing);
        elements.addNewRowBtn.classList.toggle('hidden', state.isEditing);
        elements.copySelectedRowsBtn.classList.toggle('hidden', state.isEditing);
        
        const standardRowButtons = ['select-by-keyword-btn', 'select-empty-btn', 'select-all-btn', 'invert-selection-btn'];
        standardRowButtons.forEach(id => elements[id].classList.toggle('hidden', state.isEditing));

        renderMergedTable(); // Re-render to apply/remove contenteditable
    }
    function saveEdits() {
        const tableRows = elements.mergedDisplayArea.querySelectorAll('tbody tr');
        const newData = [];
        tableRows.forEach(tr => {
            const newRowData = {};
            tr.querySelectorAll('td[data-col-header]').forEach(cell => {
                const header = cell.dataset.colHeader;
                newRowData[header] = cell.textContent;
            });
            newData.push(newRowData);
        });
        state.mergedData = newData; // Update the data store
        toggleEditMode(false); // Exit edit mode
    }
    function addNewRow() {
        const newRow = {};
        state.mergedHeaders.forEach(header => { newRow[header] = ''; });
        newRow._isNew = true;
        state.mergedData.unshift(newRow);
        if (!state.isEditing) {
            toggleEditMode(true);
        } else {
            renderMergedTable();
        }
    }
    function copySelectedRows() {
        const selectedCheckboxes = elements.mergedDisplayArea.querySelectorAll('.row-checkbox:checked');
        if (selectedCheckboxes.length === 0) {
            alert("請先勾選要複製的資料列。");
            return;
        }
        const rowsToCopy = [];
        selectedCheckboxes.forEach(cb => {
            const rowIndex = parseInt(cb.closest('tr').dataset.rowIndex, 10);
            if (!isNaN(rowIndex) && state.mergedData[rowIndex]) {
                const newRow = JSON.parse(JSON.stringify(state.mergedData[rowIndex])); // Deep copy
                newRow._isNew = true;
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
            const scope = state.isMergedView ? elements.mergedDisplayArea : elements.displayArea;
            scope.querySelectorAll('table').forEach(syncTableCheckboxState);
            if (!state.isMergedView) updateSelectionInfo();
        }, 0);
    }
    function selectAllRows() { const scope = state.isMergedView ? elements.mergedDisplayArea : elements.displayArea; const rows = scope.querySelectorAll('tbody tr:not(.row-hidden-search)'); if (rows.length === 0) { alert('沒有可勾選的列'); return; } rows.forEach(row => row.querySelector('.row-checkbox').checked = true); scope.querySelectorAll('thead input[type="checkbox"]').forEach(cb => cb.checked = true); }
    function invertSelection() { const scope = state.isMergedView ? elements.mergedDisplayArea : elements.displayArea; scope.querySelectorAll('tbody tr:not(.row-hidden-search) .row-checkbox').forEach(cb => cb.checked = !cb.checked); }
    function deleteSelectedRows() {
        const scope = state.isMergedView ? elements.mergedDisplayArea : elements.displayArea;
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
    function selectEmptyRows() { let count = 0; const scope = state.isMergedView ? elements.mergedDisplayArea : elements.displayArea; scope.querySelectorAll('tbody tr:not(.row-hidden-search)').forEach(row => { if (Array.from(row.cells).slice(1).every(c => c.textContent.trim() === '')) { row.querySelector('.row-checkbox').checked = true; count++; } }); if (count === 0) alert('未找到空白列'); }
    function selectByKeyword() { const keywordInput = elements.selectKeywordInput.value.trim(); const isRegex = elements.selectKeywordRegex.checked; if (!keywordInput) { alert('請先輸入關鍵字'); return; } let matchLogic; try { if (isRegex) { const regex = new RegExp(keywordInput, 'i'); matchLogic = text => regex.test(text); } else if (keywordInput.includes(',')) { const keywords = keywordInput.split(',').map(k => k.trim().toLowerCase()).filter(Boolean); matchLogic = text => keywords.some(k => text.includes(k)); } else { const keywords = keywordInput.split(/\s+/).map(k => k.trim().toLowerCase()).filter(Boolean); matchLogic = text => keywords.every(k => text.includes(k)); } } catch (e) { alert('無效的 Regex 表示式：\n' + e.message); return; } let count = 0; const scope = state.isMergedView ? elements.mergedDisplayArea : elements.displayArea; scope.querySelectorAll('tbody tr:not(.row-hidden-search)').forEach(row => { if (matchLogic(Array.from(row.cells).slice(1).map(c => c.textContent).join(' '))) { row.querySelector('.row-checkbox').checked = true; count++; } }); alert(count > 0 ? `已勾選 ${count} 個符合條件的列` : `未找到符合條件的列`); }
    function filterTable() { const keywords = elements.searchInput.value.toLowerCase().trim().split(/\s+/).filter(Boolean); const scope = state.isMergedView ? elements.mergedDisplayArea : elements.displayArea; scope.querySelectorAll('tbody tr').forEach(row => { const text = Array.from(row.cells).slice(1).map(c => c.textContent).join(' ').toLowerCase(); const isVisible = keywords.every(k => text.includes(k)); row.classList.toggle('row-hidden-search', !isVisible); }); if (!state.isMergedView) { elements.displayArea.querySelectorAll('.table-wrapper').forEach(wrapper => { const visibleRowCount = wrapper.querySelectorAll('tbody tr:not(.row-hidden-search)').length; wrapper.style.display = visibleRowCount > 0 ? '' : 'none'; }); } syncCheckboxesInScope(); }

    // --- Utility and Helper Functions ---
    function detectHiddenElements() { return elements.displayArea.querySelectorAll('tr[style*="display: none"], td[style*="display: none"], th[style*="display: none"]').length; }
    function showAllHiddenElements() { const hidden = elements.displayArea.querySelectorAll('tr[style*="display: none"], td[style*="display: none"], th[style*="display: none"]'); if (hidden.length === 0) { alert('沒有需要顯示的隱藏行列。'); return; } hidden.forEach(el => el.style.display = ''); alert(`已顯示 ${hidden.length} 個隱藏的行列。`); elements.showHiddenBtn.classList.add('hidden'); elements.loadStatusMessage.classList.add('hidden'); }
    function selectAllTables(isChecked) { elements.displayArea.querySelectorAll('.table-select-checkbox').forEach(cb => { if (cb.checked !== isChecked) { cb.click(); } }); }
    function readFileAsBinary(file) { return new Promise((resolve, reject) => { const reader = new FileReader(); reader.onload = e => resolve(e.target.result); reader.onerror = reject; reader.readAsBinaryString(file); }); }
    function parsePositionString(str) { const indices = new Set(); const parts = str.split(',').map(p => p.trim()).filter(Boolean); for (const part of parts) { if (part.includes('-')) { const [start, end] = part.split('-').map(Number); if (!isNaN(start) && !isNaN(end) && start <= end) { for (let i = start; i <= end; i++) indices.add(i - 1); } } else { const num = Number(part); if (!isNaN(num)) indices.add(num - 1); } } return Array.from(indices).sort((a, b) => a - b); }
    async function getSelectedSheetNames(filename, workbook, mode, criteria) { const sheetNames = workbook.SheetNames; if (sheetNames.length === 0) return []; switch (mode) { case 'all': return sheetNames; case 'first': return sheetNames.length > 0 ? [sheetNames[0]] : []; case 'specific': return sheetNames.filter(name => name.toLowerCase().includes(criteria.name.toLowerCase())); case 'position': return parsePositionString(criteria.position).map(index => sheetNames[index]).filter(Boolean); case 'manual': return await showWorksheetSelectionModal(filename, sheetNames); default: return []; } }
    function showWorksheetSelectionModal(filename, sheetNames) { return new Promise(resolve => { if (sheetNames.length <= 1) { resolve(sheetNames); return; } const overlay = document.createElement('div'); overlay.className = 'modal-overlay'; const dialog = document.createElement('div'); dialog.className = 'modal-dialog'; dialog.innerHTML = `<div class="modal-header"><h3>選擇工作表 (手動模式)</h3><p>檔案 "<strong>${filename}</strong>"</p></div><div class="modal-body"><ul class="sheet-list">${sheetNames.map(name => `<li class="sheet-item"><label><input type="checkbox" class="sheet-checkbox" value="${name}" checked> ${name}</label></li>`).join('')}</ul></div><div class="modal-footer"><button class="btn btn-secondary" id="modal-skip">跳過</button><button class="btn btn-success" id="modal-confirm">確認</button></div>`; overlay.appendChild(dialog); document.body.appendChild(overlay); const closeModal = () => document.body.removeChild(overlay); dialog.querySelector('#modal-confirm').addEventListener('click', () => { resolve(Array.from(dialog.querySelectorAll('.sheet-checkbox')).filter(cb => cb.checked).map(cb => cb.value)); closeModal(); }); dialog.querySelector('#modal-skip').addEventListener('click', () => { resolve([]); closeModal(); }); }); }
    function extractTableData(table, { onlySelected = false, includeFilename = false } = {}) { const data = []; const headerRow = table.querySelector('thead tr'); if (headerRow) { let headerData = Array.from(headerRow.querySelectorAll('th:not(.checkbox-cell):not(.column-hidden)')).map(th => th.textContent.trim()); if (includeFilename) { headerData.unshift('Source File'); } data.push(headerData); } const filename = includeFilename ? (table.closest('.table-wrapper')?.querySelector('h4')?.textContent || 'Merged Table') : null; const rows = onlySelected ? Array.from(table.querySelectorAll('tbody .row-checkbox:checked')).map(cb => cb.closest('tr')) : table.querySelectorAll('tbody tr:not(.row-hidden-search)'); rows.forEach(row => { let rowData = Array.from(row.querySelectorAll('td:not(.checkbox-cell):not(.column-hidden)')).map(td => td.textContent.trim()); if (includeFilename) { rowData.unshift(filename); } data.push(rowData); }); return data; }
    function exportHtml(mode) { const tables = state.isMergedView ? [elements.mergedDisplayArea.querySelector('table')] : Array.from(elements.displayArea.querySelectorAll('.table-wrapper:not([style*="display: none"]) table')); if (tables.length === 0 || !tables[0]) { alert('沒有可匯出的表格。'); return; } let content = ''; if (mode === 'all') { content = tables.map(table => { const clone = table.cloneNode(true); clone.querySelectorAll('.checkbox-cell, .column-hidden').forEach(el => el.remove()); return `<h4>${table.closest('.table-wrapper')?.querySelector('h4')?.textContent || '合併檢視'}</h4>` + clone.outerHTML; }).join('<br><hr><br>'); } else { const scope = state.isMergedView ? elements.mergedDisplayArea : elements.displayArea; if (scope.querySelectorAll('.row-checkbox:checked').length === 0) { alert('請先勾選要匯出的資料列。'); return; } content = tables.map(table => { const selectedRows = table.querySelectorAll('tbody .row-checkbox:checked'); if (selectedRows.length === 0) return ''; const headerClone = table.querySelector('thead').cloneNode(true); headerClone.querySelectorAll('.checkbox-cell, .column-hidden')?.forEach(el => el.remove()); const rowsHtml = Array.from(selectedRows).map(cb => { const rowClone = cb.closest('tr').cloneNode(true); rowClone.querySelectorAll('.checkbox-cell, .column-hidden')?.forEach(el => el.remove()); return rowClone.outerHTML; }).join(''); return `<h4>${table.closest('.table-wrapper')?.querySelector('h4')?.textContent || '合併檢視 (選取)'}</h4><table>${headerClone.outerHTML}<tbody>${rowsHtml}</tbody></table>`; }).join('<br><hr><br>'); } if (!content) { alert('沒有找到可匯出的內容。'); return; } const title = `匯出報表 (${mode === 'all' ? '全部' : '選取項目'})`; const html = `<!DOCTYPE html><html lang="zh-Hant"><head><meta charset="UTF-8"><title>${title}</title><style>body{font-family:sans-serif;margin:20px}table{border-collapse:collapse;width:100%;border:1px solid #ccc;margin-bottom:20px}th,td{border:1px solid #ddd;padding:8px 12px}th{background-color:#f2f2f2}</style></head><body><h1>${title}</h1><p>產生時間: ${new Date().toLocaleString()}</p>${content}</body></html>`; downloadHtml(html, `report_${mode}_${new Date().toISOString().slice(0, 10)}.html`); }
    function downloadHtml(content, filename) { const blob = new Blob([content], { type: 'text/html;charset=utf-8;' }); const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = filename; a.click(); URL.revokeObjectURL(a.href); }
    function exportXlsx(mode) { const tables = state.isMergedView ? [elements.mergedDisplayArea.querySelector('table')] : Array.from(elements.displayArea.querySelectorAll('.table-wrapper:not([style*="display: none"]) table')); if (tables.length === 0 || !tables[0]) { alert('沒有可匯出的表格。'); return; } const scope = state.isMergedView ? elements.mergedDisplayArea : elements.displayArea; if (mode === 'selected' && scope.querySelectorAll('.row-checkbox:checked').length === 0) { alert('請先勾選要匯出的資料列。'); return; } try { const workbook = XLSX.utils.book_new(); tables.forEach((table, i) => { const data = extractTableData(table, { onlySelected: mode === 'selected', includeFilename: !state.isMergedView }); if (data.length > 1) { const ws = XLSX.utils.aoa_to_sheet(data); ws['!cols'] = calculateColumnWidths(data); const sheetName = state.isMergedView ? 'Merged Data' : (table.closest('.table-wrapper')?.querySelector('h4')?.textContent.replace(/[*?:/\\\[\]]/g, '').substring(0, 31) || `Sheet${i + 1}`); XLSX.utils.book_append_sheet(workbook, ws, sheetName); } }); if (workbook.SheetNames.length === 0) { alert('沒有資料可以匯出。'); return; } XLSX.writeFile(workbook, `report_${mode}_${new Date().toISOString().slice(0, 10)}.xlsx`); } catch (err) { console.error('匯出 XLSX 錯誤:', err); alert('匯出 XLSX 時發生錯誤：' + err.message); } }
    function exportMergedXlsx() { const tables = Array.from(elements.displayArea.querySelectorAll('.table-wrapper:not([style*="display: none"]) table')); if (tables.length === 0) { alert('沒有可匯出的表格。'); return; } try { const allData = []; tables.forEach((table, i) => { const data = extractTableData(table, { includeFilename: true }); if (data.length > 1) allData.push(...(i === 0 ? data : data.slice(1))); }); if (allData.length <= 1) { alert('沒有足夠的資料可以匯出。'); return; } const ws = XLSX.utils.aoa_to_sheet(allData); ws['!cols'] = calculateColumnWidths(allData); const workbook = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(workbook, ws, 'Merged Data'); XLSX.writeFile(workbook, `report_merged_${new Date().toISOString().slice(0, 10)}.xlsx`); alert(`成功合併 ${tables.length} 個表格，共 ${allData.length - 1} 筆資料。`); } catch (err) { console.error('合併匯出 XLSX 錯誤:', err); alert('合併匯出 XLSX 時發生錯誤：' + err.message); } }
    function calculateColumnWidths(data) { if (data.length === 0) return []; return data[0].map((_, col) => ({ wch: Math.min(50, Math.max(10, ...data.map(row => row[col] ? String(row[col]).length : 0)) + 2) })); }
    
    // --- State Management and UI Updates ---
    function resetView() { if(state.isMergedView) showMultiTableView(); if (!state.originalHtmlString) return; elements.displayArea.innerHTML = state.originalHtmlString; injectCheckboxes(elements.displayArea); ['searchInput', 'selectKeywordInput'].forEach(id => elements[id].value = ''); elements.selectKeywordRegex.checked = false; filterTable(); elements.loadStatusMessage.classList.add('hidden'); const hiddenCount = detectHiddenElements(); if (hiddenCount > 0) { elements.loadStatusMessage.textContent = `注意：已重設表格，${hiddenCount} 個隱藏的行列已還原。`; elements.loadStatusMessage.classList.remove('hidden'); elements.showHiddenBtn.classList.remove('hidden'); } else { elements.showHiddenBtn.classList.add('hidden'); } updateSelectionInfo(); setViewMode('list'); }
    function resetControls(isNewFile) { if (!isNewFile) return; if(state.isMergedView) showMultiTableView(); state.originalHtmlString = ''; ['searchInput', 'selectKeywordInput'].forEach(id => elements[id].value = ''); elements.selectKeywordRegex.checked = false; elements.controlPanel.classList.add('hidden'); updateSelectionInfo(); }
    function clearAllFiles(silent = false) { if (!silent && !confirm('確定要清除所有已載入的檔案嗎？')) return; if(state.isMergedView) showMultiTableView(); state.originalHtmlString = ''; state.loadedFiles = []; state.loadedTables = 0; elements.displayArea.innerHTML = ''; elements.fileInput.value = ''; ['specificSheetNameInput', 'specificSheetPositionInput'].forEach(id => elements[id].value = ''); elements.gridScaleSlider.value = 3; updateGridScale(); updateDropAreaDisplay(); resetControls(true); setViewMode('list'); }
    function updateDropAreaDisplay() { const hasFiles = state.loadedTables > 0; elements.dropArea.classList.toggle('compact', hasFiles); elements.dropAreaInitial.classList.toggle('hidden', hasFiles); elements.dropAreaLoaded.classList.toggle('hidden', !hasFiles); elements.importOptionsContainer.classList.toggle('hidden', hasFiles); if (hasFiles) { elements.fileCount.textContent = state.loadedTables; const names = state.loadedFiles.slice(0, 3).join(', '); const more = state.loadedFiles.length > 3 ? ` 及其他 ${state.loadedFiles.length - 3} 個...` : ''; elements.fileNames.textContent = names + more; } }
    function showControls(hiddenCount) {
        elements.controlPanel.classList.remove('hidden');
        const buttonsToShow = [
            'selectByKeywordGroup', 'selectByKeywordBtn', 'selectEmptyBtn', 'deleteSelectedBtn', 
            'invertSelectionBtn', 'exportHtmlBtn', 'selectAllBtn', 'exportSelectedBtn', 
            'exportXlsxBtn', 'exportSelectedXlsxBtn', 'exportMergedXlsxBtn', 'resetViewBtn', 
            'tableLevelControls', 'listViewBtn', 'gridViewBtn', 'showHiddenBtn'
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
    
    return { init };
})();

ExcelViewer.init();
