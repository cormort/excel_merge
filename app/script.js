const ExcelViewer = (() => {
    'use strict';
    const CONSTANTS = { VALID_FILE_EXTENSIONS: ['.xls', '.xlsx'] };
    const state = { 
        originalHtmlString: '', 
        isProcessing: false, 
        loadedFiles: [], 
        loadedTables: 0, 
        zoomedCard: null 
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
            gridScaleControl: 'grid-scale-control', gridScaleSlider: 'grid-scale-slider'
        };
        Object.keys(mapping).forEach(key => {
            elements[key] = document.getElementById(mapping[key]);
        });
    }

    function bindEvents() {
        // --- Core File Handling ---
        elements.fileInput.addEventListener('change', e => processFiles(e.target.files));
        setupDragAndDrop(); // This sets up the crucial click listener for the drop area
        elements.clearFilesBtn.addEventListener('click', () => clearAllFiles(false));

        // --- View and Table Level Controls ---
        elements.listViewBtn.addEventListener('click', () => setViewMode('list'));
        elements.gridViewBtn.addEventListener('click', () => setViewMode('grid'));
        elements.gridScaleSlider.addEventListener('input', updateGridScale);
        elements.selectAllTablesBtn.addEventListener('click', () => { selectAllTables(true); updateSelectionInfo(); });
        elements.unselectAllTablesBtn.addEventListener('click', () => { selectAllTables(false); updateSelectionInfo(); });
        elements.deleteSelectedTablesBtn.addEventListener('click', deleteSelectedTables);
        
        // --- Row Operations ---
        elements.selectByKeywordBtn.addEventListener('click', selectByKeyword);
        elements.selectEmptyBtn.addEventListener('click', selectEmptyRows);
        elements.selectAllBtn.addEventListener('click', selectAllRows);
        elements.invertSelectionBtn.addEventListener('click', invertSelection);
        elements.deleteSelectedBtn.addEventListener('click', deleteSelectedRows);

        // --- Global and Export Operations ---
        elements.resetViewBtn.addEventListener('click', resetView);
        elements.showHiddenBtn.addEventListener('click', showAllHiddenElements);
        elements.exportHtmlBtn.addEventListener('click', () => exportHtml('all'));
        elements.exportSelectedBtn.addEventListener('click', () => exportHtml('selected'));
        elements.exportXlsxBtn.addEventListener('click', () => exportXlsx('all'));
        elements.exportSelectedXlsxBtn.addEventListener('click', () => exportXlsx('selected'));
        elements.exportMergedXlsxBtn.addEventListener('click', exportMergedXlsx);

        // --- Input and Dynamic Content Handling ---
        elements.searchInput.addEventListener('input', debounce(filterTable, 300));
        elements.displayArea.addEventListener('change', handleDisplayAreaChange);
        elements.displayArea.addEventListener('click', handleCardClick);
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
        document.addEventListener('keydown', e => { if (e.key === 'Escape' && state.zoomedCard) closePreview(); });
    }

    function handleDisplayAreaChange(e) {
        const target = e.target;
        if (!target.matches('.table-select-checkbox, [id^="select-all-checkbox"], .row-checkbox')) return;
        
        let table;
        if (target.matches('.table-select-checkbox')) {
            const wrapper = target.closest('.table-wrapper');
            table = wrapper ? wrapper.querySelector('table') : null;
            if (table) {
                toggleSelectAll(target.checked, table);
            }
        } else if (target.matches('[id^="select-all-checkbox"]')) {
            table = target.closest('table');
            toggleSelectAll(target.checked, table);
        } else {
            table = target.closest('table');
        }
        
        if (table) {
            syncTableCheckboxState(table);
        }
        updateSelectionInfo();
    }
    
    function setupDragAndDrop() {
        // CRITICAL FIX: This listener ensures clicking the area triggers the hidden file input.
        elements.dropArea.addEventListener('click', e => {
            if (e.target.id === 'clear-files-btn' || e.target.closest('.btn-clear') || e.target.id === 'file-input') {
                return;
            }
            elements.fileInput.click();
        });
    
        ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
            elements.dropArea.addEventListener(eventName, e => {
                e.preventDefault();
                e.stopPropagation();
            });
        });
        ['dragenter', 'dragover'].forEach(eventName => {
            elements.dropArea.addEventListener(eventName, () => elements.dropArea.classList.add('highlight'));
        });
        ['dragleave', 'drop'].forEach(eventName => {
            elements.dropArea.addEventListener(eventName, () => elements.dropArea.classList.remove('highlight'));
        });
        elements.dropArea.addEventListener('drop', e => processFiles(e.dataTransfer.files));
    }

    async function processFiles(fileList) { const validation = validateFiles(fileList); if (!validation.valid) { alert(`錯誤：${validation.error}`); return; } if (state.isProcessing) { alert('正在處理檔案...'); return; } const importMode = document.querySelector('input[name="import-mode"]:checked').value; const specificSheetName = elements.specificSheetNameInput.value.trim(); const specificSheetPosition = elements.specificSheetPositionInput.value.trim(); if (importMode === 'specific' && !specificSheetName) { alert('請輸入工作表名稱！'); return; } if (importMode === 'position' && !specificSheetPosition) { alert('請輸入工作表位置！'); return; } state.isProcessing = true; elements.displayArea.innerHTML = '<div class="loading">讀取中...</div>'; resetControls(true); const tablesToRender = []; const missedFiles = []; state.loadedFiles = []; try { for (let index = 0; index < validation.files.length; index++) { const file = validation.files[index]; elements.displayArea.innerHTML = `<div class="loading">讀取中... (${index + 1}/${validation.files.length})</div>`; const binaryData = await readFileAsBinary(file); const workbook = XLSX.read(binaryData, { type: 'binary' }); const sheetNames = await getSelectedSheetNames(file.name, workbook, importMode, { name: specificSheetName, position: specificSheetPosition }); if ((importMode === 'specific' || importMode === 'position') && sheetNames.length === 0 && workbook.SheetNames.length > 0) { missedFiles.push(file.name); } for (const sheetName of sheetNames) { const sheet = workbook.Sheets[sheetName]; const htmlString = XLSX.utils.sheet_to_html(sheet); tablesToRender.push({ html: htmlString, filename: `${file.name} (${sheetName})` }); state.loadedFiles.push(`${file.name} (${sheetName})`); } } if (missedFiles.length > 0) { const criteria = importMode === 'specific' ? `名稱包含 "${specificSheetName}"` : `位置符合 "${specificSheetPosition}"`; alert(`以下檔案找不到 ${criteria} 的工作表：\n\n- ${missedFiles.join('\n- ')}`); } state.loadedTables = tablesToRender.length; renderTables(tablesToRender); updateDropAreaDisplay(); } catch (err) { console.error("處理檔案時發生錯誤:", err); elements.displayArea.innerHTML = `<p style="color: red;">處理檔案錯誤：${err.message || '未知錯誤'}</p>`; resetControls(true); } finally { state.isProcessing = false; } }
    function readFileAsBinary(file) { return new Promise((resolve, reject) => { const reader = new FileReader(); reader.onload = e => resolve(e.target.result); reader.onerror = reject; reader.readAsBinaryString(file); }); }
    function parsePositionString(str) { const indices = new Set(); const parts = str.split(',').map(p => p.trim()).filter(Boolean); for (const part of parts) { if (part.includes('-')) { const [start, end] = part.split('-').map(Number); if (!isNaN(start) && !isNaN(end) && start <= end) { for (let i = start; i <= end; i++) indices.add(i - 1); } } else { const num = Number(part); if (!isNaN(num)) indices.add(num - 1); } } return Array.from(indices).sort((a, b) => a - b); }
    async function getSelectedSheetNames(filename, workbook, mode, criteria) { const sheetNames = workbook.SheetNames; if (sheetNames.length === 0) return []; switch (mode) { case 'all': return sheetNames; case 'first': return sheetNames.length > 0 ? [sheetNames[0]] : []; case 'specific': return sheetNames.filter(name => name.toLowerCase().includes(criteria.name.toLowerCase())); case 'position': return parsePositionString(criteria.position).map(index => sheetNames[index]).filter(Boolean); case 'manual': return await showWorksheetSelectionModal(filename, sheetNames); default: return []; } }
    function showWorksheetSelectionModal(filename, sheetNames) { return new Promise(resolve => { if (sheetNames.length <= 1) { resolve(sheetNames); return; } const overlay = document.createElement('div'); overlay.className = 'modal-overlay'; const dialog = document.createElement('div'); dialog.className = 'modal-dialog'; dialog.innerHTML = `<div class="modal-header"><h3>選擇工作表 (手動模式)</h3><p>檔案 "<strong>${filename}</strong>"</p></div><div class="modal-quick-actions"><button class="btn btn-primary btn-sm" id="modal-select-all">全選</button><button class="btn btn-secondary btn-sm" id="modal-select-none">全不選</button></div><div class="modal-body"><ul class="sheet-list">${sheetNames.map(name => `<li class="sheet-item"><label><input type="checkbox" class="sheet-checkbox" value="${name}" checked> ${name}</label></li>`).join('')}</ul></div><div class="modal-footer"><button class="btn btn-secondary" id="modal-skip">跳過</button><button class="btn btn-success" id="modal-confirm">確認</button></div>`; overlay.appendChild(dialog); document.body.appendChild(overlay); const checkboxes = dialog.querySelectorAll('.sheet-checkbox'); const closeModal = () => document.body.removeChild(overlay); dialog.querySelector('#modal-confirm').addEventListener('click', () => { resolve(Array.from(checkboxes).filter(cb => cb.checked).map(cb => cb.value)); closeModal(); }); dialog.querySelector('#modal-skip').addEventListener('click', () => { resolve([]); closeModal(); }); dialog.querySelector('#modal-select-all').addEventListener('click', () => checkboxes.forEach(cb => cb.checked = true)); dialog.querySelector('#modal-select-none').addEventListener('click', () => checkboxes.forEach(cb => cb.checked = false)); }); }
    function renderTables(tablesToRender) { if (tablesToRender.length === 0) { elements.displayArea.innerHTML = `<p>沒有找到符合條件的工作表。</p>`; return; } const fragment = document.createDocumentFragment(); tablesToRender.forEach(({ html, filename }) => { const wrapper = document.createElement('div'); wrapper.className = 'table-wrapper'; const header = document.createElement('div'); header.className = 'table-header'; header.innerHTML = `<input type="checkbox" class="table-select-checkbox" title="選取此表格"><h4>${filename}</h4><div class="header-actions"><button class="btn btn-danger btn-sm delete-rows-btn">刪除選取列</button><button class="btn btn-danger btn-sm delete-table-btn">刪除此表</button></div><button class="close-zoom">&times;</button>`; const tableContent = document.createElement('div'); tableContent.className = 'table-content'; const tempDiv = document.createElement('div'); tempDiv.innerHTML = html; const table = tempDiv.querySelector('table'); if (table) { tableContent.appendChild(table); wrapper.appendChild(header); wrapper.appendChild(tableContent); fragment.appendChild(wrapper); } }); elements.displayArea.innerHTML = ''; elements.displayArea.appendChild(fragment); state.originalHtmlString = elements.displayArea.innerHTML; injectCheckboxes(); showControls(detectHiddenElements()); }
    function injectCheckboxes() { elements.displayArea.querySelectorAll('thead tr').forEach((headRow, index) => { const th = document.createElement('th'); th.innerHTML = `<input type="checkbox" id="select-all-checkbox-${index}" title="全選/全不選">`; th.classList.add('checkbox-cell'); headRow.prepend(th); }); elements.displayArea.querySelectorAll('tbody tr').forEach(row => { const td = document.createElement('td'); td.innerHTML = '<input type="checkbox" class="row-checkbox">'; td.classList.add('checkbox-cell'); row.prepend(td); }); }
    function toggleSelectAll(isChecked, table) { if (!table) return; table.querySelectorAll('tbody tr:not(.row-hidden-search) .row-checkbox').forEach(cb => cb.checked = isChecked); }
    function detectHiddenElements() { return elements.displayArea.querySelectorAll('tr[style*="display: none"], td[style*="display: none"], th[style*="display: none"]').length; }
    function showAllHiddenElements() { const hidden = elements.displayArea.querySelectorAll('tr[style*="display: none"], td[style*="display: none"], th[style*="display: none"]'); if (hidden.length === 0) { alert('沒有需要顯示的隱藏行列。'); return; } hidden.forEach(el => el.style.display = ''); alert(`已顯示 ${hidden.length} 個隱藏的行列。`); elements.showHiddenBtn.classList.add('hidden'); elements.loadStatusMessage.classList.add('hidden'); }
    function selectAllTables(isChecked) { elements.displayArea.querySelectorAll('.table-select-checkbox').forEach(cb => { if (cb.checked !== isChecked) { cb.click(); } }); }
    function deleteSelectedTables() { const selectedWrappers = Array.from(elements.displayArea.querySelectorAll('.table-select-checkbox:checked')).map(cb => cb.closest('.table-wrapper')); if (selectedWrappers.length === 0) { alert('請先勾選要刪除的表格。'); return; } if (confirm(`確定要永久刪除 ${selectedWrappers.length} 個選定的表格嗎？`)) { selectedWrappers.forEach(wrapper => wrapper.remove()); updateFileStateAfterDeletion(); } }
    function updateFileStateAfterDeletion() { const remainingWrappers = elements.displayArea.querySelectorAll('.table-wrapper'); state.loadedTables = remainingWrappers.length; state.loadedFiles = Array.from(remainingWrappers).map(w => w.querySelector('h4').textContent); if (state.loadedTables === 0) { clearAllFiles(true); } else { updateDropAreaDisplay(); } }
    function selectByKeyword() { const keywordInput = elements.selectKeywordInput.value.trim(); const isRegex = elements.selectKeywordRegex.checked; if (!keywordInput) { alert('請先輸入關鍵字'); return; } let matchLogic; try { if (isRegex) { const regex = new RegExp(keywordInput, 'i'); matchLogic = text => regex.test(text); } else if (keywordInput.includes(',')) { const keywords = keywordInput.split(',').map(k => k.trim().toLowerCase()).filter(Boolean); matchLogic = text => keywords.some(k => text.includes(k)); } else { const keywords = keywordInput.split(/\s+/).map(k => k.trim().toLowerCase()).filter(Boolean); matchLogic = text => keywords.every(k => text.includes(k)); } } catch (e) { alert('無效的 Regex 表示式：\n' + e.message); return; } let count = 0; elements.displayArea.querySelectorAll('tbody tr:not(.row-hidden-search)').forEach(row => { if (matchLogic(Array.from(row.cells).slice(1).map(c => c.textContent).join(' '))) { row.querySelector('.row-checkbox').checked = true; count++; } }); alert(count > 0 ? `已勾選 ${count} 個符合條件的列` : `未找到符合條件的列`); syncAllTableCheckboxes(); }
    function selectEmptyRows() { let count = 0; elements.displayArea.querySelectorAll('tbody tr:not(.row-hidden-search)').forEach(row => { if (Array.from(row.cells).slice(1).every(c => c.textContent.trim() === '')) { row.querySelector('.row-checkbox').checked = true; count++; } }); if (count === 0) alert('未找到空白列'); syncAllTableCheckboxes(); }
    function selectAllRows() { const rows = elements.displayArea.querySelectorAll('tbody tr:not(.row-hidden-search)'); if (rows.length === 0) { alert('沒有可勾選的列'); return; } rows.forEach(row => row.querySelector('.row-checkbox').checked = true); elements.displayArea.querySelectorAll('[id^="select-all-checkbox"]').forEach(cb => cb.checked = true); syncAllTableCheckboxes(); }
    function invertSelection() { elements.displayArea.querySelectorAll('tbody tr:not(.row-hidden-search) .row-checkbox').forEach(cb => cb.checked = !cb.checked); syncAllTableCheckboxes(); }
    function deleteSelectedRows() { const selected = elements.displayArea.querySelectorAll('.row-checkbox:checked'); if (selected.length === 0) { alert('請先勾選要刪除的列'); return; } if (confirm(`確定要刪除 ${selected.length} 筆資料列嗎？`)) { selected.forEach(cb => cb.closest('tr').remove()); } }
    function filterTable() { const keywords = elements.searchInput.value.toLowerCase().trim().split(/\s+/).filter(Boolean); elements.displayArea.querySelectorAll('.table-wrapper').forEach(wrapper => { let visibleRowCount = 0; wrapper.querySelectorAll('tbody tr').forEach(row => { const text = Array.from(row.cells).slice(1).map(c => c.textContent).join(' ').toLowerCase(); const isVisible = keywords.every(k => text.includes(k)); row.classList.toggle('row-hidden-search', !isVisible); if (isVisible) visibleRowCount++; }); wrapper.style.display = visibleRowCount > 0 ? '' : 'none'; }); syncAllTableCheckboxes(); }
    function extractTableData(table, { onlySelected = false, includeFilename = false } = {}) { const data = []; const headerRow = table.querySelector('thead tr'); if (headerRow) { let headerData = Array.from(headerRow.querySelectorAll('th:not(.checkbox-cell)')).map(th => th.textContent.trim()); if (includeFilename) { headerData.unshift('Source File'); } data.push(headerData); } const filename = includeFilename ? table.closest('.table-wrapper').querySelector('h4').textContent : null; const rows = onlySelected ? Array.from(table.querySelectorAll('tbody .row-checkbox:checked')).map(cb => cb.closest('tr')) : table.querySelectorAll('tbody tr:not(.row-hidden-search)'); rows.forEach(row => { let rowData = Array.from(row.querySelectorAll('td:not(.checkbox-cell)')).map(td => td.textContent.trim()); if (includeFilename) { rowData.unshift(filename); } data.push(rowData); }); return data; }
    function exportHtml(mode) { const tables = Array.from(elements.displayArea.querySelectorAll('.table-wrapper:not([style*="display: none"]) table')); if (tables.length === 0) { alert('沒有可匯出的表格。'); return; } let content = ''; if (mode === 'all') { content = tables.map(table => { const clone = table.cloneNode(true); clone.querySelectorAll('.checkbox-cell').forEach(el => el.remove()); return clone.outerHTML; }).join('<br><hr><br>'); } else { if (elements.displayArea.querySelectorAll('.row-checkbox:checked').length === 0) { alert('請先勾選要匯出的資料列。'); return; } content = tables.map(table => { const selectedRows = table.querySelectorAll('tbody .row-checkbox:checked'); if (selectedRows.length === 0) return ''; const headerClone = table.querySelector('thead').cloneNode(true); headerClone.querySelector('.checkbox-cell')?.remove(); const rowsHtml = Array.from(selectedRows).map(cb => { const rowClone = cb.closest('tr').cloneNode(true); rowClone.querySelector('.checkbox-cell')?.remove(); return rowClone.outerHTML; }).join(''); return `<table>${headerClone.outerHTML}<tbody>${rowsHtml}</tbody></table>`; }).join('<br><hr><br>'); } if (!content) { alert('沒有找到可匯出的內容。'); return; } const title = `匯出報表 (${mode === 'all' ? '全部' : '選取項目'})`; const html = `<!DOCTYPE html><html lang="zh-Hant"><head><meta charset="UTF-8"><title>${title}</title><style>body{font-family:sans-serif;margin:20px}table{border-collapse:collapse;width:100%;border:1px solid #ccc;margin-bottom:20px}th,td{border:1px solid #ddd;padding:8px 12px}th{background-color:#f2f2f2}</style></head><body><h1>${title}</h1><p>產生時間: ${new Date().toLocaleString()}</p>${content}</body></html>`; downloadHtml(html, `report_${mode}_${new Date().toISOString().slice(0, 10)}.html`); }
    function downloadHtml(content, filename) { const blob = new Blob([content], { type: 'text/html;charset=utf-8;' }); const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = filename; a.click(); URL.revokeObjectURL(a.href); }
    function exportXlsx(mode) { const tables = Array.from(elements.displayArea.querySelectorAll('.table-wrapper:not([style*="display: none"]) table')); if (tables.length === 0) { alert('沒有可匯出的表格。'); return; } if (mode === 'selected' && elements.displayArea.querySelectorAll('.row-checkbox:checked').length === 0) { alert('請先勾選要匯出的資料列。'); return; } try { const workbook = XLSX.utils.book_new(); tables.forEach((table, i) => { const data = extractTableData(table, { onlySelected: mode === 'selected', includeFilename: true }); if (data.length > 1) { const ws = XLSX.utils.aoa_to_sheet(data); ws['!cols'] = calculateColumnWidths(data); XLSX.utils.book_append_sheet(workbook, ws, `Sheet${i + 1}`); } }); if (workbook.SheetNames.length === 0) { alert('沒有資料可以匯出。'); return; } XLSX.writeFile(workbook, `report_${mode}_${new Date().toISOString().slice(0, 10)}.xlsx`); } catch (err) { console.error('匯出 XLSX 錯誤:', err); alert('匯出 XLSX 時發生錯誤：' + err.message); } }
    function exportMergedXlsx() { const tables = Array.from(elements.displayArea.querySelectorAll('.table-wrapper:not([style*="display: none"]) table')); if (tables.length === 0) { alert('沒有可匯出的表格。'); return; } try { const allData = []; tables.forEach((table, i) => { const data = extractTableData(table, { includeFilename: true }); if (data.length > 1) allData.push(...(i === 0 ? data : data.slice(1))); }); if (allData.length <= 1) { alert('沒有足夠的資料可以匯出。'); return; } const ws = XLSX.utils.aoa_to_sheet(allData); ws['!cols'] = calculateColumnWidths(allData); const workbook = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(workbook, ws, 'Merged Data'); XLSX.writeFile(workbook, `report_merged_${new Date().toISOString().slice(0, 10)}.xlsx`); alert(`成功合併 ${tables.length} 個表格，共 ${allData.length - 1} 筆資料。`); } catch (err) { console.error('合併匯出 XLSX 錯誤:', err); alert('合併匯出 XLSX 時發生錯誤：' + err.message); } }
    function calculateColumnWidths(data) { if (data.length === 0) return []; return data[0].map((_, col) => ({ wch: Math.min(50, Math.max(10, ...data.map(row => row[col] ? String(row[col]).length : 0)) + 2) })); }
    function resetView() { if (!state.originalHtmlString) return; elements.displayArea.innerHTML = state.originalHtmlString; injectCheckboxes(); ['searchInput', 'selectKeywordInput'].forEach(id => elements[id].value = ''); elements.selectKeywordRegex.checked = false; filterTable(); elements.loadStatusMessage.classList.add('hidden'); const hiddenCount = detectHiddenElements(); if (hiddenCount > 0) { elements.loadStatusMessage.textContent = `注意：已重設表格，${hiddenCount} 個隱藏的行列已還原。`; elements.loadStatusMessage.classList.remove('hidden'); elements.showHiddenBtn.classList.remove('hidden'); } updateSelectionInfo(); setViewMode('list'); }
    function resetControls(isNewFile) { if (!isNewFile) return; state.originalHtmlString = ''; ['searchInput', 'selectKeywordInput'].forEach(id => elements[id].value = ''); elements.selectKeywordRegex.checked = false; elements.controlPanel.classList.add('hidden'); updateSelectionInfo(); }
    function updateDropAreaDisplay() { const hasFiles = state.loadedTables > 0; elements.dropArea.classList.toggle('compact', hasFiles); elements.dropAreaInitial.classList.toggle('hidden', hasFiles); elements.dropAreaLoaded.classList.toggle('hidden', !hasFiles); elements.importOptionsContainer.classList.toggle('hidden', hasFiles); if (hasFiles) { elements.fileCount.textContent = state.loadedTables; const names = state.loadedFiles.slice(0, 3).join(', '); const more = state.loadedFiles.length > 3 ? ` 及其他 ${state.loadedFiles.length - 3} 個...` : ''; elements.fileNames.textContent = names + more; } }
    function clearAllFiles(silent = false) { if (!silent && !confirm('確定要清除所有已載入的檔案嗎？')) return; state.originalHtmlString = ''; state.loadedFiles = []; state.loadedTables = 0; elements.displayArea.innerHTML = ''; elements.fileInput.value = ''; ['specificSheetNameInput', 'specificSheetPositionInput'].forEach(id => elements[id].value = ''); elements.gridScaleSlider.value = 3; updateGridScale(); updateDropAreaDisplay(); resetControls(true); setViewMode('list'); }
    function showControls(hiddenCount) { elements.controlPanel.classList.remove('hidden'); ['selectByKeywordGroup', 'selectByKeywordBtn', 'selectEmptyBtn', 'deleteSelectedBtn', 'invertSelectionBtn', 'exportHtmlBtn', 'selectAllBtn', 'exportSelectedBtn', 'exportXlsxBtn', 'exportSelectedXlsxBtn', 'exportMergedXlsxBtn', 'resetViewBtn', 'tableLevelControls', 'listViewBtn', 'gridViewBtn'].forEach(id => { const el = elements[id]; if(el) el.classList.remove('hidden'); }); if (hiddenCount > 0) { elements.loadStatusMessage.textContent = `注意：檔案中包含 ${hiddenCount} 個被隱藏的行列。`; elements.loadStatusMessage.classList.remove('hidden'); elements.showHiddenBtn.classList.remove('hidden'); } }
    function setViewMode(mode) { if (mode === 'grid') { elements.displayArea.classList.remove('list-view'); elements.displayArea.classList.add('grid-view'); elements.gridViewBtn.classList.add('active'); elements.listViewBtn.classList.remove('active'); elements.gridScaleControl.classList.remove('hidden'); } else { elements.displayArea.classList.remove('grid-view'); elements.displayArea.classList.add('list-view'); elements.listViewBtn.classList.add('active'); elements.gridViewBtn.classList.remove('active'); elements.gridScaleControl.classList.add('hidden'); } }
    function updateGridScale() { const columns = elements.gridScaleSlider.value; elements.displayArea.style.setProperty('--grid-columns', columns); }
    function validateFiles(fileList) { if (!fileList || fileList.length === 0) return { valid: false, error: '沒有選擇檔案' }; const validFiles = Array.from(fileList).filter(file => CONSTANTS.VALID_FILE_EXTENSIONS.some(ext => file.name.toLowerCase().endsWith(ext))); if (validFiles.length === 0) return { valid: false, error: '請上傳 .xls 或 .xlsx 格式的檔案' }; return { valid: true, files: validFiles }; }
    function debounce(func, wait) { let timeout; return (...args) => { clearTimeout(timeout); timeout = setTimeout(() => func(...args), wait); }; }
    function handleScroll() { if (window.scrollY > window.innerHeight / 2) { elements.backToTopBtn.classList.add('visible'); } else { elements.backToTopBtn.classList.remove('visible'); } }
    function scrollToTop() { window.scrollTo({ top: 0, behavior: 'smooth' }); }
    function handleCardClick(e) { const card = e.target.closest('.table-wrapper'); if (!card) return; if (e.target.classList.contains('close-zoom')) { closePreview(); return; } if (e.target.classList.contains('delete-rows-btn')) { deleteSelectedRows(); return; } if (e.target.classList.contains('delete-table-btn')) { if (confirm(`確定要永久刪除此工作表 (${card.querySelector('h4').textContent}) 嗎？`)) { closePreview(); setTimeout(() => { card.remove(); updateFileStateAfterDeletion(); }, 300); } return; } if (elements.displayArea.classList.contains('grid-view') && !card.classList.contains('is-zoomed')) { if (!e.target.matches('input, a, button, .btn')) { openPreview(card); } } }
    function openPreview(card) { if (state.zoomedCard) return; card.classList.add('is-zoomed'); state.zoomedCard = card; document.body.classList.add('no-scroll'); }
    function closePreview() { if (!state.zoomedCard) return; state.zoomedCard.classList.remove('is-zoomed'); state.zoomedCard = null; document.body.classList.remove('no-scroll'); }
    function syncAllTableCheckboxes() { setTimeout(() => { elements.displayArea.querySelectorAll('table').forEach(syncTableCheckboxState); updateSelectionInfo(); }, 0); }
    function syncTableCheckboxState(table) { const wrapper = table.closest('.table-wrapper'); if (!wrapper) return; const headerCheckbox = wrapper.querySelector('.table-select-checkbox'); const rowCheckboxes = table.querySelectorAll('tbody tr:not(.row-hidden-search) .row-checkbox'); if (rowCheckboxes.length === 0) { headerCheckbox.checked = false; headerCheckbox.indeterminate = false; return; } const checkedCount = Array.from(rowCheckboxes).filter(cb => cb.checked).length; if (checkedCount === 0) { headerCheckbox.checked = false; headerCheckbox.indeterminate = false; } else if (checkedCount === rowCheckboxes.length) { headerCheckbox.checked = true; headerCheckbox.indeterminate = false; } else { headerCheckbox.checked = false; headerCheckbox.indeterminate = true; } }
    function updateSelectionInfo() { const selectedCheckboxes = elements.displayArea.querySelectorAll('.table-select-checkbox:checked, .table-select-checkbox:indeterminate'); if (selectedCheckboxes.length > 0) { const names = Array.from(selectedCheckboxes).map(cb => cb.closest('.table-header').querySelector('h4').textContent); elements.selectedTablesList.textContent = names.join('; '); elements.selectedTablesInfo.classList.remove('hidden'); } else { elements.selectedTablesInfo.classList.add('hidden'); } }

    return { init };
})();

ExcelViewer.init();
