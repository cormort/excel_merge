// Excel Viewer - 重構優化版本 + XLSX 匯出功能 (包含合併所有表格)
const ExcelViewer = (() => {
    'use strict';

    // ==================== 常數定義 ====================
    const CONSTANTS = {
        CHECKBOX_COLUMN_INDEX: 0,
        FILENAME_COLUMN_INDEX: 1,
        DATA_START_INDEX: 2, // 勾選框和檔名欄之後
        VALID_FILE_EXTENSIONS: ['.xls', '.xlsx']
    };

    // ==================== 狀態管理 ====================
    const state = {
        originalHtmlString: '',
        isProcessing: false
    };

    // ==================== DOM 元素快取 ====================
    const elements = {
        fileInput: null,
        displayArea: null,
        searchInput: null,
        dropArea: null,
        deleteSelectedBtn: null,
        invertSelectionBtn: null,
        resetViewBtn: null,
        selectEmptyBtn: null,
        exportHtmlBtn: null,
        showHiddenBtn: null,
        exportSelectedBtn: null,
        exportXlsxBtn: null,
        exportSelectedXlsxBtn: null,
        exportMergedXlsxBtn: null,
        selectByKeywordGroup: null,
        selectKeywordInput: null,
        selectByKeywordBtn: null,
        selectKeywordRegex: null,
        loadStatusMessage: null,
        controlPanel: null
    };

    // ==================== 初始化 ====================
    function init() {
        cacheElements();
        bindEvents();
    }

    function cacheElements() {
        elements.fileInput = document.getElementById('file-input');
        elements.displayArea = document.getElementById('excel-display-area');
        elements.searchInput = document.getElementById('search-input');
        elements.dropArea = document.getElementById('drop-area');
        elements.deleteSelectedBtn = document.getElementById('delete-selected-btn');
        elements.invertSelectionBtn = document.getElementById('invert-selection-btn');
        elements.resetViewBtn = document.getElementById('reset-view-btn');
        elements.selectEmptyBtn = document.getElementById('select-empty-btn');
        elements.exportHtmlBtn = document.getElementById('export-html-btn');
        elements.showHiddenBtn = document.getElementById('show-hidden-btn');
        elements.exportSelectedBtn = document.getElementById('export-selected-btn');
        elements.exportXlsxBtn = document.getElementById('export-xlsx-btn');
        elements.exportSelectedXlsxBtn = document.getElementById('export-selected-xlsx-btn');
        elements.exportMergedXlsxBtn = document.getElementById('export-merged-xlsx-btn');
        elements.selectByKeywordGroup = document.getElementById('select-by-keyword-group');
        elements.selectKeywordInput = document.getElementById('select-keyword-input');
        elements.selectByKeywordBtn = document.getElementById('select-by-keyword-btn');
        elements.selectKeywordRegex = document.getElementById('select-keyword-regex');
        elements.loadStatusMessage = document.getElementById('load-status-message');
        elements.controlPanel = document.getElementById('control-panel');
    }

    // ==================== 事件綁定 ====================
    function bindEvents() {
        elements.fileInput.addEventListener('change', (e) => processFiles(e.target.files));
        setupDragAndDrop();
        elements.displayArea.addEventListener('change', handleDisplayAreaChange);
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

    function setupDragAndDrop() {
        elements.dropArea.addEventListener('click', (e) => {
            if (e.target.tagName !== 'BUTTON' && e.target.id !== 'file-input') {
                elements.fileInput.click();
            }
        });
        ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
            elements.dropArea.addEventListener(eventName, preventDefaults);
        });
        ['dragenter', 'dragover'].forEach(eventName => {
            elements.dropArea.addEventListener(eventName, () => elements.dropArea.classList.add('highlight'));
        });
        ['dragleave', 'drop'].forEach(eventName => {
            elements.dropArea.addEventListener(eventName, () => elements.dropArea.classList.remove('highlight'));
        });
        elements.dropArea.addEventListener('drop', (e) => {
            processFiles(e.dataTransfer.files);
        });
    }

    function handleDisplayAreaChange(e) {
        if (e.target.matches('[id^="select-all-checkbox"]')) {
            const table = e.target.closest('table');
            toggleSelectAll(e.target.checked, table);
        }
    }

    // ==================== 工具函數 ====================
    function preventDefaults(e) {
        e.preventDefault();
        e.stopPropagation();
    }

    function debounce(func, wait) {
        let timeout;
        return function executedFunction(...args) {
            const later = () => {
                clearTimeout(timeout);
                func(...args);
            };
            clearTimeout(timeout);
            timeout = setTimeout(later, wait);
        };
    }

    function readFileAsBinary(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = (e) => resolve(e.target.result);
            reader.onerror = reject;
            reader.readAsBinaryString(file);
        });
    }

    function createElementWithText(tag, text, className = null) {
        const el = document.createElement(tag);
        el.textContent = text;
        if (className) el.classList.add(className);
        return el;
    }

    function uncheckAllSelectAllCheckboxes() {
        elements.displayArea.querySelectorAll('[id^="select-all-checkbox"]')
            .forEach(cb => cb.checked = false);
    }

    function getDataCellsText(row, toLowerCase = true) {
        let text = '';
        for (let i = CONSTANTS.DATA_START_INDEX; i < row.cells.length; i++) {
            const cellContent = row.cells[i].textContent;
            text += toLowerCase ? cellContent.toLowerCase() : cellContent;
        }
        return text;
    }

    function getCurrentDateString() {
        return new Date().toISOString().slice(0, 10);
    }
    
    function generateExportHtml(htmlContent, title) {
        return `<!DOCTYPE html>
<html lang="zh-Hant">
<head>
    <meta charset="UTF-8">
    <title>匯出報表 - ${new Date().toLocaleString()}</title>
    <style>
        body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif; margin: 20px; line-height: 1.6; }
        table { border-collapse: collapse; width: 100%; border: 1px solid #ccc; margin-bottom: 20px; }
        th, td { border: 1px solid #ddd; padding: 8px 12px; text-align: left; }
        th { background-color: #f2f2f2; }
        .filename-cell { background-color: #f8f9f9; font-size: 0.9em; color: #555; }
    </style>
</head>
<body>
    <h1>${title}</h1>
    <p>產生時間: ${new Date().toLocaleString()}</p>
    ${htmlContent}
</body>
</html>`;
    }

    function downloadHtml(content, filename) {
        const blob = new Blob([content], { type: 'text/html;charset=utf-8;' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.download = filename;
        a.href = url;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
    }
    
    // ==================== 檔案處理 ====================
    function validateFiles(fileList) {
        if (!fileList || fileList.length === 0) {
            return { valid: false, error: '沒有選擇檔案' };
        }
        const validFiles = Array.from(fileList).filter(file => {
            const fileName = file.name.toLowerCase();
            return CONSTANTS.VALID_FILE_EXTENSIONS.some(ext => fileName.endsWith(ext));
        });
        if (validFiles.length === 0) {
            return { 
                valid: false, 
                error: `請上傳 ${CONSTANTS.VALID_FILE_EXTENSIONS.join(' 或 ')} 格式的 Excel 檔案！` 
            };
        }
        return { valid: true, files: validFiles };
    }

    async function processFiles(fileList) {
        const validation = validateFiles(fileList);
        if (!validation.valid) {
            alert(`錯誤：${validation.error}`);
            return;
        }
        if (state.isProcessing) {
            alert('正在處理檔案，請稍候...');
            return;
        }
        state.isProcessing = true;
        elements.displayArea.innerHTML = '<div class="loading">讀取中，請稍候</div>';
        resetControls(true);
        const tablesToRender = [];
        try {
            for (let index = 0; index < validation.files.length; index++) {
                const file = validation.files[index];
                elements.displayArea.innerHTML = 
                    `<div class="loading">讀取中，請稍候 (${index + 1}/${validation.files.length}): ${file.name}</div>`;
                const binaryData = await readFileAsBinary(file);
                const workbook = XLSX.read(binaryData, { type: 'binary' });
                const selectedSheets = await selectWorksheets(file.name, workbook.SheetNames);
                for (const sheetName of selectedSheets) {
                    const sheet = workbook.Sheets[sheetName];
                    const htmlString = XLSX.utils.sheet_to_html(sheet);
                    tablesToRender.push({
                        html: htmlString,
                        filename: `${file.name} (${sheetName})`
                    });
                }
            }
            renderTables(tablesToRender);
        } catch (err) {
            console.error("處理檔案時發生錯誤:", err);
            elements.displayArea.innerHTML = 
                `<p style="color: red;">處理檔案時發生錯誤：${err.message || '未知錯誤'}</p>`;
            resetControls(true);
        } finally {
            state.isProcessing = false;
        }
    }

    async function selectWorksheets(filename, sheetNames) {
        if (sheetNames.length === 1) {
            return [sheetNames[0]];
        }
        const promptMessage = 
            `檔案 "${filename}" 包含多個工作表：\n\n` +
            sheetNames.map((name, i) => `${i + 1}: ${name}`).join('\n') +
            '\n\n請輸入您要匯入的工作表編號(可多選，用逗號分隔)，例如: 1,3\n(留白則跳過此檔案)';
        
        const choices = prompt(promptMessage);
        if (!choices) return [];
        return choices.split(',')
            .map(s => sheetNames[parseInt(s.trim()) - 1])
            .filter(Boolean);
    }

    function renderTables(tablesToRender) {
        if (tablesToRender.length === 0) {
            elements.displayArea.innerHTML = '<p>沒有選擇任何工作表，請重新上傳檔案。</p>';
            return;
        }
        const fragment = document.createDocumentFragment();
        for (const tableData of tablesToRender) {
            const tempDiv = document.createElement('div');
            tempDiv.innerHTML = tableData.html;
            const table = tempDiv.querySelector('table');
            if (table) {
                injectFilenameColumn(table, tableData.filename);
                fragment.appendChild(table);
            }
        }
        elements.displayArea.innerHTML = '';
        elements.displayArea.appendChild(fragment);
        state.originalHtmlString = elements.displayArea.innerHTML;
        injectCheckboxes();
        const hiddenCount = detectHiddenElements();
        showControls(hiddenCount);
    }

    // ==================== 表格操作 ====================
    function injectFilenameColumn(tableElement, filename) {
        const th = createElementWithText('th', 'Source File', 'filename-cell');
        tableElement.querySelector('thead tr')?.prepend(th);
        const rows = tableElement.querySelectorAll('tbody tr');
        rows.forEach(row => {
            const td = createElementWithText('td', filename, 'filename-cell');
            row.prepend(td);
        });
    }

    function injectCheckboxes() {
        const headRows = elements.displayArea.querySelectorAll('thead tr');
        headRows.forEach((headRow, index) => {
            const selectAllTh = document.createElement('th');
            selectAllTh.innerHTML = 
                `<input type="checkbox" id="select-all-checkbox-${index}" title="全選/全不選 (此表格)">`;
            selectAllTh.classList.add('checkbox-cell');
            headRow.prepend(selectAllTh);
        });
        const bodyRows = elements.displayArea.querySelectorAll('tbody tr');
        bodyRows.forEach(row => {
            const checkCell = document.createElement('td');
            checkCell.innerHTML = '<input type="checkbox" class="row-checkbox">';
            checkCell.classList.add('checkbox-cell');
            row.prepend(checkCell);
        });
    }

    function toggleSelectAll(isChecked, table) {
        if (!table) return;
        const visibleRows = table.querySelectorAll('tbody tr:not(.row-hidden-search)');
        visibleRows.forEach(row => {
            const checkbox = row.querySelector('.row-checkbox');
            if (checkbox) checkbox.checked = isChecked;
        });
    }

    // ==================== 隱藏元素處理 ====================
    function detectHiddenElements() {
        const selector = 'tr[style*="display: none"], td[style*="display: none"], th[style*="display: none"]';
        return elements.displayArea.querySelectorAll(selector).length;
    }

    function showAllHiddenElements() {
        const selector = 'tr[style*="display: none"], td[style*="display: none"], th[style*="display: none"]';
        const hiddenElements = elements.displayArea.querySelectorAll(selector);
        
        if (hiddenElements.length === 0) {
            alert('沒有需要顯示的隱藏行列。');
            return;
        }
        hiddenElements.forEach(el => el.style.display = '');
        alert(`已顯示 ${hiddenElements.length} 個隱藏的行列。`);
        elements.showHiddenBtn.classList.add('hidden');
        elements.loadStatusMessage.classList.add('hidden');
    }

    // ==================== 選取功能 ====================
    function selectByKeyword() {
        const keywordInput = elements.selectKeywordInput.value.trim();
        const isRegex = elements.selectKeywordRegex.checked;
        if (keywordInput === '') {
            alert('請先輸入要勾選的關鍵字。');
            return;
        }
        let matchLogic;
        try {
            if (isRegex) {
                const regex = new RegExp(keywordInput, 'i');
                matchLogic = (cellText) => regex.test(cellText);
            } else if (keywordInput.includes(',')) {
                const keywords = keywordInput.split(',')
                    .map(k => k.trim().toLowerCase()).filter(Boolean);
                matchLogic = (cellText) => keywords.some(k => cellText.includes(k));
            } else {
                const keywords = keywordInput.split(/\s+/)
                    .map(k => k.trim().toLowerCase()).filter(Boolean);
                matchLogic = (cellText) => keywords.every(k => cellText.includes(k));
            }
        } catch (e) {
            alert('無效的 Regex 表示式：\n' + e.message);
            return;
        }

        const rows = elements.displayArea.querySelectorAll('tbody tr:not(.row-hidden-search)');
        let count = 0;
        rows.forEach(row => {
            const cellText = getDataCellsText(row, !isRegex);
            if (matchLogic(cellText)) {
                const cb = row.querySelector('.row-checkbox');
                if (cb) {
                    cb.checked = true;
                    count++;
                }
            }
        });
        if (count === 0) {
            alert(`在可見的資料列中，沒有找到符合 "${keywordInput}" 的列。`);
        } else {
            alert(`已勾選 ${count} 個符合 "${keywordInput}" 的資料列。`);
            uncheckAllSelectAllCheckboxes();
        }
    }

    function selectEmptyRows() {
        const rows = elements.displayArea.querySelectorAll('tbody tr:not(.row-hidden-search)');
        let count = 0;
        rows.forEach(row => {
            if (getDataCellsText(row).trim() === '') {
                const cb = row.querySelector('.row-checkbox');
                if (cb) {
                    cb.checked = true;
                    count++;
                }
            }
        });
        if (count === 0) {
            alert('在可見的資料列中，沒有找到完全空白的列。');
        } else {
            uncheckAllSelectAllCheckboxes();
        }
    }

    function invertSelection() {
        const rows = elements.displayArea.querySelectorAll('tbody tr:not(.row-hidden-search)');
        rows.forEach(row => {
            const cb = row.querySelector('.row-checkbox');
            if (cb) cb.checked = !cb.checked;
        });
        uncheckAllSelectAllCheckboxes();
    }

    // ==================== 刪除/篩選 ====================
    function deleteSelectedRows() {
        const selectedCheckboxes = elements.displayArea.querySelectorAll('.row-checkbox:checked');
        if (selectedCheckboxes.length === 0) {
            alert('請先勾選要刪除的資料列。');
            return;
        }
        if (confirm(`確定要永久刪除 ${selectedCheckboxes.length} 筆勾選的資料列嗎？`)) {
            selectedCheckboxes.forEach(cb => cb.closest('tr').remove());
            uncheckAllSelectAllCheckboxes();
        }
    }

    function filterTable() {
        const searchTerm = elements.searchInput.value.toLowerCase().trim();
        const keywords = searchTerm.split(/\s+/).filter(Boolean);
        const rows = elements.displayArea.querySelectorAll('tbody tr');
        rows.forEach(row => {
            const cellText = getDataCellsText(row, true);
            const matchesAll = keywords.every(k => cellText.includes(k));
            row.classList.toggle('row-hidden-search', !matchesAll);
        });
        uncheckAllSelectAllCheckboxes();
    }

    // ==================== 匯出 HTML ====================
    function exportHtml(mode) {
        const tables = elements.displayArea.querySelectorAll('table');
        if (tables.length === 0) {
            alert('沒有可匯出的表格。');
            return;
        }
        let combinedCleanedHtml = '';
        if (mode === 'all') {
            combinedCleanedHtml = exportAllTablesHtml(tables);
        } else if (mode === 'selected') {
            if (elements.displayArea.querySelectorAll('.row-checkbox:checked').length === 0) {
                alert('請先勾選要匯出的資料列。');
                return;
            }
            combinedCleanedHtml = exportSelectedTablesHtml(tables);
        }
        if (combinedCleanedHtml === '') {
            alert('沒有找到可匯出的內容。');
            return;
        }
        const title = mode === 'all' ? '匯出報表 (全部)' : '匯出報表 (選取項目)';
        const filename = `report_${mode}_${getCurrentDateString()}.html`;
        downloadHtml(generateExportHtml(combinedCleanedHtml, title), filename);
    }

    function exportAllTablesHtml(tables) {
        let html = '';
        tables.forEach(table => {
            const tableClone = table.cloneNode(true);
            tableClone.querySelectorAll('.checkbox-cell').forEach(cell => cell.remove());
            tableClone.querySelectorAll('.row-hidden-search').forEach(row => row.classList.remove('row-hidden-search'));
            html += tableClone.outerHTML + '<br><hr><br>';
        });
        return html;
    }

    function exportSelectedTablesHtml(tables) {
        let html = '';
        tables.forEach(table => {
            const selectedRows = table.querySelectorAll('tbody .row-checkbox:checked');
            if (selectedRows.length > 0) {
                const headerClone = table.querySelector('thead').cloneNode(true);
                headerClone.querySelector('.checkbox-cell')?.remove();
                let selectedRowsHtml = '';
                selectedRows.forEach(cb => {
                    const rowClone = cb.closest('tr').cloneNode(true);
                    rowClone.querySelector('.checkbox-cell')?.remove();
                    rowClone.classList.remove('row-hidden-search');
                    selectedRowsHtml += rowClone.outerHTML;
                });
                html += '<table>' + headerClone.outerHTML + '<tbody>' + selectedRowsHtml + '</tbody></table><br><hr><br>';
            }
        });
        return html;
    }

    // ==================== 匯出 XLSX ====================
    function exportXlsx(mode) {
        const tables = elements.displayArea.querySelectorAll('table');
        if (tables.length === 0) {
            alert('沒有可匯出的表格。');
            return;
        }
        if (mode === 'selected' && elements.displayArea.querySelectorAll('.row-checkbox:checked').length === 0) {
            alert('請先勾選要匯出的資料列。');
            return;
        }
        try {
            const workbook = XLSX.utils.book_new();
            let sheetIndex = 1;
            tables.forEach(table => {
                const dataToExport = extractTableData(table, mode === 'selected');
                if (dataToExport.length > 1) { // 至少要有表頭+一筆資料
                    const worksheet = XLSX.utils.aoa_to_sheet(dataToExport);
                    worksheet['!cols'] = calculateColumnWidths(dataToExport);
                    let sheetName = `Sheet${sheetIndex}`;
                    try {
                        const firstRow = table.querySelector('tbody tr .filename-cell');
                        if (firstRow) {
                            const match = firstRow.textContent.match(/\(([^)]+)\)$/);
                            if (match && match[1]) {
                                sheetName = match[1].substring(0, 27) + `_${sheetIndex}`;
                            }
                        }
                    } catch (e) { /* 忽略錯誤 */ }
                    XLSX.utils.book_append_sheet(workbook, worksheet, sheetName);
                    sheetIndex++;
                }
            });
            if (workbook.SheetNames.length === 0) {
                alert('沒有資料可以匯出。');
                return;
            }
            XLSX.writeFile(workbook, `report_${mode}_${getCurrentDateString()}.xlsx`);
        } catch (err) {
            console.error('匯出 XLSX 時發生錯誤:', err);
            alert('匯出 XLSX 時發生錯誤：' + (err.message || '未知錯誤'));
        }
    }

    // ==================== 匯出合併 XLSX (新功能) ====================
    function exportMergedXlsx() {
        const tables = elements.displayArea.querySelectorAll('table');
        if (tables.length === 0) {
            alert('沒有可匯出的表格。');
            return;
        }
        try {
            const allData = [];
            let isFirstTable = true;
            let headerRow = null;

            tables.forEach((table, tableIndex) => {
                // 合併時，總是匯出所有可見的資料 (非選取)
                const tableData = extractTableData(table, false);
                if (tableData.length > 1) { // 至少有表頭+一筆資料
                    if (isFirstTable) {
                        headerRow = tableData[0];
                        allData.push(...tableData);
                        isFirstTable = false;
                    } else {
                        // 檢查表頭是否相符，不符則警告
                        if (JSON.stringify(headerRow) !== JSON.stringify(tableData[0])) {
                            console.warn(`警告: 表格 ${tableIndex + 1} 的欄位與第一個表格不符，可能導致合併結果錯位。`);
                        }
                        allData.push(...tableData.slice(1)); // 只加入資料列
                    }
                }
            });

            if (allData.length <= 1) {
                alert('沒有足夠的資料可以合併匯出。');
                return;
            }

            const worksheet = XLSX.utils.aoa_to_sheet(allData);
            worksheet['!cols'] = calculateColumnWidths(allData);
            const workbook = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(workbook, worksheet, 'Merged_Data');
            XLSX.writeFile(workbook, `report_merged_${getCurrentDateString()}.xlsx`);
            alert(`成功合併 ${tables.length} 個表格，共 ${allData.length - 1} 筆資料（不含表頭）`);
        } catch (err) {
            console.error('匯出合併 XLSX 時發生錯誤:', err);
            alert('匯出合併 XLSX 時發生錯誤：' + (err.message || '未知錯誤'));
        }
    }

    function extractTableData(table, onlySelected) {
        const data = [];
        const headerRow = table.querySelector('thead tr');
        if (headerRow) {
            const headerData = [];
            headerRow.querySelectorAll('th:not(.checkbox-cell)').forEach(th => {
                headerData.push(th.textContent.trim());
            });
            data.push(headerData);
        }

        const rows = onlySelected 
            ? Array.from(table.querySelectorAll('tbody .row-checkbox:checked')).map(cb => cb.closest('tr'))
            : table.querySelectorAll('tbody tr:not(.row-hidden-search)');

        rows.forEach(row => {
            const rowData = [];
            row.querySelectorAll('td:not(.checkbox-cell)').forEach(td => {
                const text = td.textContent.trim();
                const num = Number(text.replace(/,/g, ''));
                if (text !== '' && !isNaN(num) && text.search(/[^0-9.,-]/) === -1) {
                    rowData.push(num); // 作為數字儲存
                } else {
                    rowData.push(text); // 作為文字儲存
                }
            });
            data.push(rowData);
        });
        return data;
    }

    function calculateColumnWidths(data) {
        if (data.length === 0) return [];
        const widths = [];
        data[0].forEach((_, colIndex) => {
            let maxWidth = 10; // 最小寬度
            data.forEach(row => {
                if (row[colIndex]) {
                    let cellLength = 0;
                    String(row[colIndex]).split('').forEach(char => {
                        cellLength += char.match(/[^\x00-\xff]/) ? 2 : 1; // 全形算2, 半形算1
                    });
                    maxWidth = Math.max(maxWidth, cellLength);
                }
            });
            widths.push({ wch: Math.min(maxWidth + 2, 50) }); // 限制最大寬度
        });
        return widths;
    }

    // ==================== 重設/UI控制 ====================
    function resetView() {
        if (!state.originalHtmlString) return;
        elements.displayArea.innerHTML = state.originalHtmlString;
        injectCheckboxes();
        elements.searchInput.value = '';
        elements.selectKeywordInput.value = '';
        elements.selectKeywordRegex.checked = false;
        filterTable();
        elements.loadStatusMessage.classList.add('hidden');
        elements.showHiddenBtn.classList.add('hidden');
        const hiddenCount = detectHiddenElements();
        if (hiddenCount > 0) {
            elements.loadStatusMessage.textContent = 
                `注意：已重設表格，${hiddenCount} 個隱藏的行列已還原為隱藏狀態。`;
            elements.loadStatusMessage.classList.remove('hidden');
            elements.showHiddenBtn.classList.remove('hidden');
        }
    }

    function resetControls(isNewFile = false) {
        if (isNewFile) {
            state.originalHtmlString = '';
            elements.searchInput.value = '';
            elements.selectKeywordInput.value = '';
            elements.selectKeywordRegex.checked = false;
            elements.controlPanel.classList.add('hidden');
            Object.values(elements).forEach(el => {
                if (el && el.id !== 'control-panel' && el.id !== 'file-input' && el.id !== 'drop-area') {
                    el.classList?.add('hidden');
                }
            });
        }
    }

    function showControls(hiddenCount) {
        elements.controlPanel.classList.remove('hidden');
        [
            elements.selectByKeywordGroup, elements.selectByKeywordBtn,
            elements.selectEmptyBtn, elements.deleteSelectedBtn,
            elements.invertSelectionBtn, elements.exportHtmlBtn,
            elements.exportSelectedBtn, elements.exportXlsxBtn,
            elements.exportSelectedXlsxBtn, elements.exportMergedXlsxBtn,
            elements.resetViewBtn
        ].forEach(el => el?.classList.remove('hidden'));

        if (hiddenCount > 0) {
            elements.loadStatusMessage.textContent = 
                `注意：載入的檔案中包含 ${hiddenCount} 個被 Excel 隱藏的行列。您可以使用「顯示隱藏的行列」按鈕來查看它們。`;
            elements.loadStatusMessage.classList.remove('hidden');
            elements.showHiddenBtn.classList.remove('hidden');
        }
    }

    // ==================== 公開 API ====================
    return {
        init
    };
})();

// 初始化應用程式
document.addEventListener('DOMContentLoaded', ExcelViewer.init);
