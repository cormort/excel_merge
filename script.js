// Excel Viewer - v11 - Refactored + XLSX Export
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
    const elements = {};

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
        elements.selectByKeywordGroup = document.getElementById('select-by-keyword-group');
        elements.selectKeywordInput = document.getElementById('select-keyword-input');
        elements.selectByKeywordBtn = document.getElementById('select-by-keyword-btn');
        elements.selectKeywordRegex = document.getElementById('select-keyword-regex');
        elements.loadStatusMessage = document.getElementById('load-status-message');
        elements.controlPanel = document.getElementById('control-panel');
    }

    // ==================== 事件綁定 ====================
    function bindEvents() {
        // 檔案上傳相關
        elements.fileInput.addEventListener('change', (e) => processFiles(e.target.files));
        
        // 拖放區域
        setupDragAndDrop();
        
        // 事件委派 - 處理動態生成的全選 checkbox
        elements.displayArea.addEventListener('change', handleDisplayAreaChange);
        
        // 按鈕事件
        elements.selectEmptyBtn.addEventListener('click', selectEmptyRows);
        elements.deleteSelectedBtn.addEventListener('click', deleteSelectedRows);
        elements.invertSelectionBtn.addEventListener('click', invertSelection);
        elements.resetViewBtn.addEventListener('click', resetView);
        elements.selectByKeywordBtn.addEventListener('click', selectByKeyword);
        elements.exportHtmlBtn.addEventListener('click', () => exportHtml('all'));
        elements.exportSelectedBtn.addEventListener('click', () => exportHtml('selected'));
        elements.exportXlsxBtn.addEventListener('click', () => exportXlsx('all'));
        elements.exportSelectedXlsxBtn.addEventListener('click', () => exportXlsx('selected'));
        elements.showHiddenBtn.addEventListener('click', showAllHiddenElements);
        
        // 搜尋輸入 (使用 Debounce)
        elements.searchInput.addEventListener('input', debounce(filterTable, 300));
    }

    function setupDragAndDrop() {
        elements.dropArea.addEventListener('click', (e) => {
            // 確保點擊按鈕時不會觸發 dropArea 的點擊
            if (e.target.tagName !== 'BUTTON' && e.target.id !== 'file-input') {
                elements.fileInput.click();
            }
        });

        ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
            elements.dropArea.addEventListener(eventName, preventDefaults);
        });

        ['dragenter', 'dragover'].forEach(eventName => {
            elements.dropArea.addEventListener(eventName, () => {
                elements.dropArea.classList.add('highlight');
            });
        });

        ['dragleave', 'drop'].forEach(eventName => {
            elements.dropArea.addEventListener(eventName, () => {
                elements.dropArea.classList.remove('highlight');
            });
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
        // 從 DATA_START_INDEX (預設為 2) 開始，跳過 勾選框 和 檔名
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

    // ==================== 檔案處理 (v10 邏輯) ====================
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

    // ==================== 表格操作 (v10 邏輯) ====================
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
            const selectAllId = `select-all-checkbox-${index}`;
            selectAllTh.innerHTML = 
                `<input type="checkbox" id="${selectAllId}" title="全選/全不選 (此表格)">`;
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

    // ==================== 隱藏元素處理 (v10 邏輯) ====================
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

    // ==================== 選取功能 (v10 邏輯) ====================
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
                    .map(k => k.trim().toLowerCase())
                    .filter(k => k.length > 0);
                matchLogic = (cellText) => keywords.some(k => cellText.includes(k));
            } else {
                const keywords = keywordInput.split(/\s+/)
                    .map(k => k.trim().toLowerCase())
                    .filter(k => k.length > 0);
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
            const dataCellText = getDataCellsText(row).trim();
            if (dataCellText === '') {
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

    // ==================== 刪除/篩選 (User's v11 函式) ====================
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
        const keywords = searchTerm.split(/\s+/).filter(k => k.length > 0);
        const rows = elements.displayArea.querySelectorAll('tbody tr');

        rows.forEach(row => {
            const cellText = getDataCellsText(row, true);
            const matchesAll = keywords.every(k => cellText.includes(k));
            
            row.classList.toggle('row-hidden-search', !matchesAll);
        });

        uncheckAllSelectAllCheckboxes();
    }

    // ==================== 匯出 HTML (User's v11 函式) ====================
    function exportHtml(mode) {
        const tables = elements.displayArea.querySelectorAll('table');
        
        if (tables.length === 0) {
            alert('沒有可匯出的表格。');
            return;
        }

        let combinedCleanedHtml = '';

        if (mode === 'all') {
            combinedCleanedHtml = exportAllTables(tables);
        } else if (mode === 'selected') {
            const selectedCheckboxes = elements.displayArea.querySelectorAll('.row-checkbox:checked');
            
            if (selectedCheckboxes.length === 0) {
                alert('請先勾選要匯出的資料列。');
                return;
            }
            
            combinedCleanedHtml = exportSelectedTables(tables);
        }

        if (combinedCleanedHtml === '') {
            alert('沒有找到可匯出的內容。');
            return;
        }

        const title = mode === 'all' ? '匯出報表 (全部)' : '匯出報表 (選取項目)';
        const filename = `report_${mode}_${getCurrentDateString()}.html`;
        const htmlContent = generateExportHtml(combinedCleanedHtml, title);
        
        downloadHtml(htmlContent, filename);
    }

    function exportAllTables(tables) {
        let html = '';
        
        tables.forEach(table => {
            const tableClone = table.cloneNode(true);
            tableClone.querySelectorAll('.checkbox-cell').forEach(cell => cell.remove());
            tableClone.querySelectorAll('tr.row-hidden-search')
                .forEach(row => row.classList.remove('row-hidden-search'));
            html += tableClone.outerHTML + '<br><hr><br>';
        });
        
        return html;
    }

    function exportSelectedTables(tables) {
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
                
                html += '<table>' + headerClone.outerHTML + 
                        '<tbody>' + selectedRowsHtml + '</tbody></table><br><hr><br>';
            }
        });
        
        return html;
    }


    // ==================== 匯出 XLSX (User's v11 函式) ====================
    function exportXlsx(mode) {
        const tables = elements.displayArea.querySelectorAll('table');
        
        if (tables.length === 0) {
            alert('沒有可匯出的表格。');
            return;
        }

        if (mode === 'selected') {
            const selectedCheckboxes = elements.displayArea.querySelectorAll('.row-checkbox:checked');
            if (selectedCheckboxes.length === 0) {
                alert('請先勾選要匯出的資料列。');
                return;
            }
        }

        try {
            const workbook = XLSX.utils.book_new();
            let sheetIndex = 1;

            tables.forEach((table, tableIndex) => {
                let dataToExport = [];

                if (mode === 'all') {
                    dataToExport = extractTableData(table, false);
                } else if (mode === 'selected') {
                    dataToExport = extractTableData(table, true);
                }

                if (dataToExport.length > 0) {
                    const worksheet = XLSX.utils.aoa_to_sheet(dataToExport);
                    
                    // 自動調整欄寬
                    const colWidths = calculateColumnWidths(dataToExport);
                    worksheet['!cols'] = colWidths;
                    
                    // 試著從原始檔名中獲取工作表名
                    let sheetName = `Sheet${sheetIndex}`;
                    try {
                        const firstRow = table.querySelector('tbody tr .filename-cell');
                        if (firstRow) {
                            // 從 "filename.xlsx (SheetName)" 中提取 "SheetName"
                            const match = firstRow.textContent.match(/\(([^)]+)\)$/);
                            if(match && match[1]) {
                                // 確保工作表名稱不重複且在31字元內
                                sheetName = match[1].substring(0, 27) + `_${sheetIndex}`;
                            }
                        }
                    } catch (e) { /* 忽略錯誤，使用預設工作表名 */ }

                    XLSX.utils.book_append_sheet(workbook, worksheet, sheetName);
                    sheetIndex++;
                }
            });

            if (workbook.SheetNames.length === 0) {
                alert('沒有資料可以匯出。');
                return;
            }

            const filename = `report_${mode}_${getCurrentDateString()}.xlsx`;
            XLSX.writeFile(workbook, filename);
            
        } catch (err) {
            console.error('匯出 XLSX 時發生錯誤:', err);
            alert('匯出 XLSX 時發生錯誤：' + (err.message || '未知錯誤'));
        }
    }

    function extractTableData(table, onlySelected) {
        const data = [];
        
        // 提取表頭
        const headerRow = table.querySelector('thead tr');
        if (headerRow) {
            const headerData = [];
            headerRow.querySelectorAll('th').forEach(th => {
                if (!th.classList.contains('checkbox-cell')) {
                    headerData.push(th.textContent.trim());
                }
            });
            data.push(headerData);
        }

        // 提取資料列
        let rows;
        if (onlySelected) {
            rows = table.querySelectorAll('tbody .row-checkbox:checked');
            rows = Array.from(rows).map(cb => cb.closest('tr'));
        } else {
            rows = table.querySelectorAll('tbody tr:not(.row-hidden-search)');
        }

        rows.forEach(row => {
            const rowData = [];
            row.querySelectorAll('td').forEach(td => {
                if (!td.classList.contains('checkbox-cell')) {
                    // 試著將純數字字串轉回數字，以便 Excel 格式化
                    const text = td.textContent.trim();
                    const num = Number(text.replace(/,/g, ''));
                    if (text !== '' && !isNaN(num) && text.search(/[^0-9.,-]/) === -1) {
                        rowData.push(num);
                    } else {
                        rowData.push(text);
                    }
                }
            });
            data.push(rowData);
        });

        return data;
    }

    function calculateColumnWidths(data) {
        if (data.length === 0) return [];
        
        const colCount = data[0].length;
        const widths = [];

        for (let col = 0; col < colCount; col++) {
            let maxWidth = 10; // 最小寬度
            
            data.forEach(row => {
                if (row[col]) {
                    // 計算字元寬度 (全形字算2, 半形算1)
                    let cellLength = 0;
                    String(row[col]).split('').forEach(char => {
                        cellLength += char.match(/[^\x00-\xff]/) ? 2 : 1;
                    });
                    maxWidth = Math.max(maxWidth, cellLength);
                }
            });
            
            // 限制最大寬度為 50
            widths.push({ wch: Math.min(maxWidth + 2, 50) });
        }

        return widths;
    }

    // ==================== 重設功能 (User's v11 函式) ====================
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

            const controlElements = [
                elements.selectByKeywordGroup,
                elements.selectByKeywordBtn,
                elements.selectEmptyBtn,
                elements.deleteSelectedBtn,
                elements.invertSelectionBtn,
                elements.exportHtmlBtn,
                elements.exportSelectedBtn,
                elements.exportXlsxBtn,
                elements.exportSelectedXlsxBtn,
                elements.resetViewBtn,
                elements.showHiddenBtn,
                elements.loadStatusMessage
            ];

            controlElements.forEach(el => el?.classList.add('hidden'));
        }
    }

    function showControls(hiddenCount) {
        elements.controlPanel.classList.remove('hidden');

        const controlsToShow = [
            elements.selectByKeywordGroup,
            elements.selectByKeywordBtn,
            elements.selectEmptyBtn,
            elements.deleteSelectedBtn,
            elements.invertSelectionBtn,
            elements.exportHtmlBtn,
            elements.exportSelectedBtn,
            elements.exportXlsxBtn,
            elements.exportSelectedXlsxBtn,
            elements.resetViewBtn
        ];

        controlsToShow.forEach(el => el?.classList.remove('hidden'));

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
ExcelViewer.init();
