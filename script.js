// Excel Viewer - v13 - UI Enhancement with Upload Summary
const ExcelViewer = (() => {
    'use strict';

    // --- 狀態與常數定義 ---
    const CONSTANTS = {
        DATA_START_INDEX: 2, // 資料欄位的起始索引 (跳過勾選框和檔名)
        VALID_FILE_EXTENSIONS: ['.xls', '.xlsx']
    };

    const state = {
        originalHtmlString: '', // 用於「重設視圖」的原始 HTML 快照
        isProcessing: false,    // 防止在處理檔案時重複觸發
        loadedFileCount: 0      // 已載入的檔案數量
    };

    // 集中管理所有 DOM 元素
    const elements = {};

    // --- 初始化模組 ---
    function init() {
        cacheElements();
        bindEvents();
    }

    /**
     * 快取所有需要操作的 DOM 元素，提高效能。
     */
    function cacheElements() {
        const ids = {
            fileInput: 'file-input', displayArea: 'excel-display-area',
            searchInput: 'search-input', dropArea: 'drop-area',
            deleteSelectedBtn: 'delete-selected-btn', invertSelectionBtn: 'invert-selection-btn',
            resetViewBtn: 'reset-view-btn', selectEmptyBtn: 'select-empty-btn',
            exportHtmlBtn: 'export-html-btn', showHiddenBtn: 'show-hidden-btn',
            exportSelectedBtn: 'export-selected-btn', exportXlsxBtn: 'export-xlsx-btn',
            exportSelectedXlsxBtn: 'export-selected-xlsx-btn', exportMergedXlsxBtn: 'export-merged-xlsx-btn',
            selectByKeywordGroup: 'select-by-keyword-group', selectKeywordInput: 'select-keyword-input',
            selectByKeywordBtn: 'select-by-keyword-btn', selectKeywordRegex: 'select-keyword-regex',
            loadStatusMessage: 'load-status-message', controlPanel: 'control-panel',
            // 新增的 UI 元素
            uploadSummary: 'upload-summary', fileCount: 'file-count', clearFilesBtn: 'clear-files-btn'
        };
        for (const key in ids) {
            elements[key] = document.getElementById(ids[key]);
        }
    }

    /**
     * 綁定所有事件監聽器。
     */
    function bindEvents() {
        elements.fileInput.addEventListener('change', e => processFiles(e.target.files));
        setupDragAndDrop();
        elements.displayArea.addEventListener('change', handleDisplayAreaChange);
        
        // 按鈕事件綁定
        elements.clearFilesBtn.addEventListener('click', fullReset); // 新增：完全重設按鈕
        elements.selectEmptyBtn.addEventListener('click', selectEmptyRows);
        elements.deleteSelectedBtn.addEventListener('click', deleteSelectedRows);
        elements.invertSelectionBtn.addEventListener('click', invertSelection);
        elements.resetViewBtn.addEventListener('click', resetView);
        elements.selectByKeywordBtn.addEventListener('click', selectByKeyword);
        elements.exportHtmlBtn.addEventListener('click', () => exportHtml('all'));
        elements.exportSelectedBtn.addEventListener('click', () => exportHtml('selected'));
        elements.exportXlsxBtn.addEventListener('click', () => exportXlsx('all'));
        elements.exportSelectedXlsxBtn.addEventListener('click', () => exportXlsx('selected'));
        elements.exportMergedXlsxBtn.addEventListener('click', exportMergedXlsx);
        elements.showHiddenBtn.addEventListener('click', showAllHiddenElements);
        elements.searchInput.addEventListener('input', debounce(filterTable, 300));
    }

    // --- 事件處理與設定 ---
    function setupDragAndDrop() {
        const preventDefaults = e => { e.preventDefault(); e.stopPropagation(); };
        ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
            elements.dropArea.addEventListener(eventName, preventDefaults);
        });
        ['dragenter', 'dragover'].forEach(eventName => {
            elements.dropArea.addEventListener(eventName, () => elements.dropArea.classList.add('highlight'));
        });
        ['dragleave', 'drop'].forEach(eventName => {
            elements.dropArea.addEventListener(eventName, () => elements.dropArea.classList.remove('highlight'));
        });
        elements.dropArea.addEventListener('drop', e => processFiles(e.dataTransfer.files));
    }

    function handleDisplayAreaChange(e) {
        if (e.target.matches('[id^="select-all-checkbox"]')) {
            toggleSelectAll(e.target.checked, e.target.closest('table'));
        }
    }

    // --- 核心檔案處理邏輯 ---
    async function processFiles(fileList) {
        const validation = validateFiles(fileList);
        if (!validation.valid) {
            alert(`錯誤：${validation.error}`);
            return;
        }
        if (state.isProcessing) return alert('正在處理檔案，請稍候...');

        state.isProcessing = true;
        fullReset(); // 每次新上傳都從乾淨的狀態開始
        elements.displayArea.innerHTML = '<div class="loading">讀取中，請稍候</div>';

        const tablesToRender = [];
        try {
            for (let i = 0; i < validation.files.length; i++) {
                const file = validation.files[i];
                elements.displayArea.innerHTML = `<div class="loading">讀取中 (${i + 1}/${validation.files.length}): ${file.name}</div>`;
                const binaryData = await readFileAsBinary(file);
                const workbook = XLSX.read(binaryData, { type: 'binary' });
                const selectedSheets = await selectWorksheets(file.name, workbook.SheetNames);
                for (const sheetName of selectedSheets) {
                    const htmlString = XLSX.utils.sheet_to_html(workbook.Sheets[sheetName]);
                    tablesToRender.push({ html: htmlString, filename: `${file.name} (${sheetName})` });
                }
            }
            state.loadedFileCount = validation.files.length;
            renderTables(tablesToRender);
        } catch (err) {
            console.error("處理檔案時發生錯誤:", err);
            elements.displayArea.innerHTML = `<p style="color: red;">處理檔案時發生錯誤：${err.message || '未知錯誤'}</p>`;
            fullReset(); // 發生錯誤時也重設回初始狀態
        } finally {
            state.isProcessing = false;
        }
    }

    function renderTables(tablesToRender) {
        if (tablesToRender.length === 0) {
            fullReset();
            elements.displayArea.innerHTML = '<p>沒有選擇任何工作表，請重新上傳檔案。</p>';
            return;
        }

        const fragment = document.createDocumentFragment();
        tablesToRender.forEach(tableData => {
            const tempDiv = document.createElement('div');
            tempDiv.innerHTML = tableData.html;
            const table = tempDiv.querySelector('table');
            if (table) {
                injectCustomColumns(table, tableData.filename);
                fragment.appendChild(table);
            }
        });

        elements.displayArea.innerHTML = '';
        elements.displayArea.appendChild(fragment);
        state.originalHtmlString = elements.displayArea.innerHTML;
        
        const hiddenCount = detectHiddenElements();
        updateUIVisibility(hiddenCount, state.loadedFileCount);
    }
    
    // --- 表格 DOM 操作 ---
    function injectCustomColumns(tableElement, filename) {
        const headRow = tableElement.querySelector('thead tr');
        if (!headRow) return;

        // 注入勾選框表頭
        const selectAllTh = document.createElement('th');
        selectAllTh.innerHTML = `<input type="checkbox" id="select-all-checkbox-${Date.now()}" title="全選/全不選">`;
        selectAllTh.classList.add('checkbox-cell');
        headRow.prepend(selectAllTh);

        // 注入檔名表頭
        headRow.children[1].insertAdjacentElement('beforebegin', createElementWithText('th', 'Source File', 'filename-cell'));

        // 注入表格內容的勾選框與檔名欄位
        tableElement.querySelectorAll('tbody tr').forEach(row => {
            const checkCell = document.createElement('td');
            checkCell.innerHTML = '<input type="checkbox" class="row-checkbox">';
            checkCell.classList.add('checkbox-cell');
            row.prepend(checkCell);
            row.children[1].insertAdjacentElement('beforebegin', createElementWithText('td', filename, 'filename-cell'));
        });
    }
    
    // --- UI 狀態管理 ---
    function fullReset() {
        state.originalHtmlString = '';
        state.isProcessing = false;
        state.loadedFileCount = 0;
        
        elements.fileInput.value = ''; // 允許重新上傳相同檔案
        elements.displayArea.innerHTML = '';
        elements.searchInput.value = '';
        elements.selectKeywordInput.value = '';
        elements.selectKeywordRegex.checked = false;

        // 隱藏所有控制項，只顯示初始上傳區
        elements.controlPanel.classList.add('hidden');
        elements.loadStatusMessage.classList.add('hidden');
        elements.uploadSummary.classList.add('hidden');
        elements.dropArea.classList.remove('hidden');
    }
    
    function resetView() {
        if (!state.originalHtmlString) return;
        elements.displayArea.innerHTML = state.originalHtmlString; // 重設表格內容和勾選框
        elements.searchInput.value = '';
        elements.selectKeywordInput.value = '';
        elements.selectKeywordRegex.checked = false;
        filterTable(); // 套用空篩選以顯示所有列
        updateUIVisibility(detectHiddenElements(), state.loadedFileCount);
    }

    function updateUIVisibility(hiddenCount, fileCount) {
        // 切換主佈局：隱藏大上傳區，顯示小總結面板和控制區
        elements.dropArea.classList.add('hidden');
        elements.uploadSummary.classList.remove('hidden');
        elements.fileCount.textContent = fileCount;
        elements.controlPanel.classList.remove('hidden');
        
        // 根據是否存在隱藏元素，決定是否顯示提示訊息和按鈕
        const hasHidden = hiddenCount > 0;
        elements.loadStatusMessage.classList.toggle('hidden', !hasHidden);
        elements.showHiddenBtn.classList.toggle('hidden', !hasHidden);
        if (hasHidden) {
            elements.loadStatusMessage.textContent = `注意：載入的檔案中包含 ${hiddenCount} 個被 Excel 隱藏的行列。`;
        }
    }
    
    // --- 選取功能 ---
    function toggleSelectAll(isChecked, table) {
        if (!table) return;
        table.querySelectorAll('tbody tr:not(.row-hidden-search) .row-checkbox')
             .forEach(checkbox => checkbox.checked = isChecked);
    }

    function selectByKeyword() {
        const keywordInput = elements.selectKeywordInput.value.trim();
        const isRegex = elements.selectKeywordRegex.checked;
        if (!keywordInput) return alert('請先輸入要勾選的關鍵字。');

        let matchLogic;
        try {
            if (isRegex) {
                const regex = new RegExp(keywordInput, 'i');
                matchLogic = text => regex.test(text);
            } else if (keywordInput.includes(',')) {
                const keywords = keywordInput.split(',').map(k => k.trim().toLowerCase()).filter(Boolean);
                matchLogic = text => keywords.some(k => text.toLowerCase().includes(k));
            } else {
                const keywords = keywordInput.split(/\s+/).map(k => k.trim().toLowerCase()).filter(Boolean);
                matchLogic = text => keywords.every(k => text.toLowerCase().includes(k));
            }
        } catch (e) { return alert('無效的 Regex 表示式：\n' + e.message); }

        let count = 0;
        elements.displayArea.querySelectorAll('tbody tr:not(.row-hidden-search)').forEach(row => {
            if (matchLogic(getDataCellsText(row, false))) {
                const cb = row.querySelector('.row-checkbox');
                if (cb) { cb.checked = true; count++; }
            }
        });

        alert(count > 0 ? `已勾選 ${count} 個符合條件的資料列。` : `在可見的資料列中，沒有找到符合條件的列。`);
        if (count > 0) uncheckAllSelectAllCheckboxes();
    }

    function selectEmptyRows() {
        let count = 0;
        elements.displayArea.querySelectorAll('tbody tr:not(.row-hidden-search)').forEach(row => {
            if (getDataCellsText(row).trim() === '') {
                const cb = row.querySelector('.row-checkbox');
                if (cb) { cb.checked = true; count++; }
            }
        });
        alert(count > 0 ? `已勾選 ${count} 個空白列。` : '在可見的資料列中，沒有找到完全空白的列。');
        if (count > 0) uncheckAllSelectAllCheckboxes();
    }

    function invertSelection() {
        elements.displayArea.querySelectorAll('tbody tr:not(.row-hidden-search) .row-checkbox')
             .forEach(cb => cb.checked = !cb.checked);
        uncheckAllSelectAllCheckboxes();
    }

    // --- 資料操作與檢視 ---
    function deleteSelectedRows() {
        const toDelete = elements.displayArea.querySelectorAll('.row-checkbox:checked');
        if (toDelete.length === 0) return alert('請先勾選要刪除的資料列。');
        if (confirm(`確定要永久刪除 ${toDelete.length} 筆勾選的資料列嗎？`)) {
            toDelete.forEach(cb => cb.closest('tr').remove());
            uncheckAllSelectAllCheckboxes();
        }
    }

    function filterTable() {
        const keywords = elements.searchInput.value.toLowerCase().trim().split(/\s+/).filter(Boolean);
        elements.displayArea.querySelectorAll('tbody tr').forEach(row => {
            const cellText = getDataCellsText(row, true);
            const isVisible = keywords.every(k => cellText.includes(k));
            row.classList.toggle('row-hidden-search', !isVisible);
        });
        uncheckAllSelectAllCheckboxes();
    }
    
    function detectHiddenElements() {
        return elements.displayArea.querySelectorAll('tr[style*="display: none"], td[style*="display: none"], th[style*="display: none"]').length;
    }

    function showAllHiddenElements() {
        const hidden = elements.displayArea.querySelectorAll('[style*="display: none"]');
        if (hidden.length === 0) return alert('沒有需要顯示的隱藏行列。');
        hidden.forEach(el => el.style.display = '');
        alert(`已顯示 ${hidden.length} 個隱藏的行列。`);
        elements.showHiddenBtn.classList.add('hidden');
        elements.loadStatusMessage.classList.add('hidden');
    }

    // --- 匯出功能 ---
    function exportXlsx(mode) {
        if (elements.displayArea.querySelectorAll('table').length === 0) return alert('沒有可匯出的表格。');
        if (mode === 'selected' && elements.displayArea.querySelectorAll('.row-checkbox:checked').length === 0) {
            return alert('請先勾選要匯出的資料列。');
        }

        try {
            const workbook = XLSX.utils.book_new();
            elements.displayArea.querySelectorAll('table').forEach((table, index) => {
                const data = extractTableData(table, mode === 'selected');
                if (data.length > 1) { // 必須有表頭 + 至少一筆資料
                    const ws = XLSX.utils.aoa_to_sheet(data);
                    ws['!cols'] = calculateColumnWidths(data);
                    XLSX.utils.book_append_sheet(workbook, ws, `Sheet${index + 1}`);
                }
            });
            if (workbook.SheetNames.length === 0) return alert('沒有資料可以匯出。');
            XLSX.writeFile(workbook, `report_${mode}_${getCurrentDateString()}.xlsx`);
        } catch (err) {
            console.error('匯出 XLSX 時發生錯誤:', err);
            alert('匯出 XLSX 時發生錯誤：' + (err.message || '未知錯誤'));
        }
    }

    function exportMergedXlsx() {
        const tables = elements.displayArea.querySelectorAll('table');
        if (tables.length === 0) return alert('沒有可匯出的表格。');
        try {
            const allData = [];
            let masterHeader = null;
            let processedTableCount = 0;

            tables.forEach((table, index) => {
                const tableData = extractTableData(table, false); // 合併模式永遠使用所有可見資料
                if (tableData.length > 1) {
                    processedTableCount++;
                    if (!masterHeader) {
                        masterHeader = tableData[0];
                        allData.push(...tableData);
                    } else {
                        if (JSON.stringify(masterHeader) !== JSON.stringify(tableData[0])) {
                            console.warn(`警告: 表格 ${index + 1} 的欄位與第一個表格不符。`);
                        }
                        allData.push(...tableData.slice(1)); // 只加入資料列
                    }
                }
            });

            if (allData.length <= 1) return alert('沒有足夠的資料可以合併匯出。');
            const ws = XLSX.utils.aoa_to_sheet(allData);
            ws['!cols'] = calculateColumnWidths(allData);
            const wb = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(wb, ws, 'Merged_Data');
            XLSX.writeFile(wb, `report_merged_${getCurrentDateString()}.xlsx`);
            alert(`成功合併 ${processedTableCount} 個表格，共 ${allData.length - 1} 筆資料。`);
        } catch (err) {
            console.error('匯出合併 XLSX 時發生錯誤:', err);
            alert('匯出合併 XLSX 時發生錯誤：' + (err.message || '未知錯誤'));
        }
    }
    
    // --- 輔助函數 ---
    function extractTableData(table, onlySelected) {
        const data = [];
        const headerCells = table.querySelectorAll('thead th:not(.checkbox-cell)');
        data.push(Array.from(headerCells).map(th => th.textContent.trim()));

        const rows = onlySelected
            ? Array.from(table.querySelectorAll('tbody .row-checkbox:checked')).map(cb => cb.closest('tr'))
            : table.querySelectorAll('tbody tr:not(.row-hidden-search)');

        rows.forEach(row => {
            const rowData = [];
            row.querySelectorAll('td:not(.checkbox-cell)').forEach(td => {
                const text = td.textContent.trim();
                const num = Number(text.replace(/,/g, ''));
                rowData.push(text !== '' && !isNaN(num) && text.search(/[^0-9.,-]/) === -1 ? num : text);
            });
            data.push(rowData);
        });
        return data;
    }

    function calculateColumnWidths(data) {
        if (data.length === 0) return [];
        return data[0].map((_, colIndex) => {
            let maxWidth = 10;
            data.forEach(row => {
                if (row[colIndex]) {
                    const cellLength = String(row[colIndex]).split('').reduce((acc, char) => acc + (char.match(/[^\x00-\xff]/) ? 2 : 1), 0);
                    maxWidth = Math.max(maxWidth, cellLength);
                }
            });
            return { wch: Math.min(maxWidth + 2, 50) };
        });
    }

    function debounce(func, wait) {
        let timeout;
        return (...args) => {
            clearTimeout(timeout);
            timeout = setTimeout(() => func.apply(this, args), wait);
        };
    }

    function readFileAsBinary(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = e => resolve(e.target.result);
            reader.onerror = reject;
            reader.readAsBinaryString(file);
        });
    }
    
    function validateFiles(fileList) {
        if (!fileList || fileList.length === 0) return { valid: false, error: '沒有選擇檔案' };
        const validFiles = Array.from(fileList).filter(f => CONSTANTS.VALID_FILE_EXTENSIONS.some(ext => f.name.toLowerCase().endsWith(ext)));
        if (validFiles.length === 0) return { valid: false, error: `請上傳 ${CONSTANTS.VALID_FILE_EXTENSIONS.join(' 或 ')} 格式的檔案！` };
        return { valid: true, files: validFiles };
    }

    async function selectWorksheets(filename, sheetNames) {
        if (sheetNames.length === 1) return [sheetNames[0]];
        const promptMsg = `檔案 "${filename}" 包含多個工作表：\n\n${sheetNames.map((n, i) => `${i + 1}: ${n}`).join('\n')}\n\n請輸入要匯入的工作表編號(用逗號分隔)，留白則跳過。`;
        const choices = prompt(promptMsg);
        return choices ? choices.split(',').map(s => sheetNames[parseInt(s.trim()) - 1]).filter(Boolean) : [];
    }
    
    function createElementWithText(tag, text, className) {
        const el = document.createElement(tag);
        el.textContent = text;
        el.className = className;
        return el;
    }

    // --- 匯出公開 API ---
    return { init };
})();

// 啟動應用程式
ExcelViewer.init();
