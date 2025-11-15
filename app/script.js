<!DOCTYPE html>
<html lang="zh-Hant">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>基金資料彙總報告產生器 (v11.0-final)</title>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
    <style>
        body {
            font-family: -apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,"Helvetica Neue",Arial,"Noto Sans",sans-serif;
            line-height: 1.6; background: #f4f7f6; color: #333; padding:20px; max-width:1400px; margin:0 auto;
        }
        h1,h2,h3{color:#0056b3;}
        h2{border-bottom:2px solid #e0e0e0; padding-bottom:5px;}
        #app-container{background:#fff;border-radius:8px;box-shadow:0 4px 12px #0001;padding:25px;}
        .config-section{margin-bottom:30px;}
        #drop-area{border:2px dashed #007bff;border-radius:8px;padding:40px;text-align:center;background:#f8faff;color:#0056b3;font-size:1.1em;cursor:pointer;transition:.3s;}
        #drop-area:hover{background:#e6f0ff;}
        #file-input{display:none;} #file-list-display{list-style:none;padding:0;margin-top:10px;font-size:.9em;color:#555;}
        #preview-area{max-height:400px;overflow:auto;border:1px solid #ddd;border-radius:5px;background:#fafafa;margin-top:15px;}
        #preview-area table,.report-table{width:100%;border-collapse:collapse;font-size:.9em;}
        #preview-area th,#preview-area td,.report-table th,.report-table td{border:1px solid #ccc;padding:6px 8px;min-width:80px;text-align:left;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;}
        #preview-area thead th{position:-webkit-sticky;position:sticky;top:0;background:#e9ecef;z-index:10;}
        .report-table td.number{text-align:right;font-family:monospace;}
        .input-group{margin-top:15px; display: flex; flex-wrap: wrap; align-items: center; gap: 10px;}
        .input-group label{font-weight:bold;color:#555; flex-shrink: 0;}
        .input-group input[type="text"], .input-group input[type="number"] {padding:8px 10px;border:1px solid #ccc;border-radius:4px;box-sizing:border-box;}
        .input-group input[type="text"] { flex-grow: 1; min-width: 200px; }
        .input-group input[type="number"] { width: 100px; }
        
        button{background:#007bff;color:#fff;border:none;padding:10px 18px;border-radius:5px;cursor:pointer;font-size:1em;font-weight:500;transition:.2s;margin-right:10px;}
        button:hover{background:#0056b3;}
        button:disabled{background:#c0c0c0;cursor:not-allowed;}
        #load-headers-btn{background:#28a745;} #load-headers-btn:hover{background:#218838;}
        #mapping-fields{display:none;padding-top:15px;}
        
        .mapping-table-container{margin-top:20px;overflow-x:auto;}
        .mapping-table{width:100%;border-collapse:collapse;background:#fff;}
        .mapping-table th{background:#007bff;color:#fff;padding:12px;text-align:left;font-weight:600;}
        .mapping-table td{padding:10px;border-bottom:1px solid #e0e0e0;}
        .mapping-table input[type="text"], .mapping-table select {width:100%;padding:8px;border:1px solid #ccc;border-radius:4px;}
        
        .output-section{display:none;}
        .view-tabs{border-bottom:2px solid #dee2e6;margin-bottom:15px;}
        .tab-btn{background:transparent;color:#0056b3;border:none;padding:10px 15px;cursor:pointer;font-size:1.1em;border-radius:5px 5px 0 0;}
        .tab-btn.active{background:#0056b3;color:#fff;}
        .view-pane{display:none;padding:10px;}
        .view-pane.active{display:block;}
        #item-view select{font-size:1.1em;padding:8px;margin-bottom:15px;}
        #individual-view .individual-report{margin-bottom:30px;border-bottom:2px solid #0056b3;padding-bottom:15px;}
        .report-table tbody tr:first-child { font-weight: bold; background-color: #e9ecef; }
    </style>
</head>
<body>
<div id="app-container">
    <h1>基金資料彙總報告產生器 (v11.0-final)</h1>
    
    <div class="config-section">
        <h2>1. 上傳檔案</h2>
        <div id="drop-area">拖曳 <b>多個</b> Excel 檔案至此，或點擊此處選擇檔案
            <input type="file" id="file-input" accept=".xlsx,.xls" multiple>
        </div>
        <ul id="file-list-display"></ul>
    </div>
    
    <div class="config-section">
        <h2>2. 檔案預覽 (第一個檔案)</h2>
        <div id="preview-area"><p style="padding:20px;text-align:center;color:#777;">尚未載入檔案</p></div>
    </div>
    
    <div class="config-section">
        <h2>3. 設定資料範圍</h2>
        <div class="input-group">
            <button id="auto-detect-btn" disabled>1. 自動偵測範圍</button>
        </div>
        <div class="input-group">
            <label for="data-range-input">總資料範圍</label>
            <input type="text" id="data-range-input" placeholder="偵測結果將顯示於此">
            <label for="header-rows-input">標頭佔用列數</label>
            <input type="number" id="header-rows-input" value="1" min="1">
        </div>
        <div class="input-group">
             <button id="load-headers-btn" disabled>2. 讀取欄位</button>
        </div>
    </div>
    
    <div class="mapping-section">
        <h2>4. 欄位設定與對應</h2>
        <div id="mapping-fields"></div>
        <p id="mapping-placeholder" style="color:#777;">請先讀取欄位</p>
    </div>
    
    <div class="config-section">
        <button id="process-btn" disabled>開始彙總處理</button>
    </div>
    
    <div class="output-section" id="output-area">
        <h2>5. 彙總報告結果</h2>
        <div class="view-tabs">
            <button class="tab-btn active" data-view="summary-view">加總總表</button>
            <button class="tab-btn" data-view="individual-view">個別檔案</button>
            <button class="tab-btn" data-view="item-view">項目查詢</button>
        </div>
        <div id="view-content">
            <div id="summary-view" class="view-pane active"></div>
            <div id="individual-view" class="view-pane"></div>
            <div id="item-view" class="view-pane">
                <select id="item-dropdown"></select>
                <div id="item-detail-table"></div>
            </div>
        </div>
    </div>
</div>

<script>
const state = { workbooks: [], columnMappings: [], allFileData:[], summaryData:new Map() };

// DOM Elements
const dropArea = document.getElementById('drop-area');
const fileInput = document.getElementById('file-input');
const fileListDisplay = document.getElementById('file-list-display');
const previewArea = document.getElementById('preview-area');
const mappingFields = document.getElementById('mapping-fields');
const mappingPlaceholder = document.getElementById('mapping-placeholder');
const processBtn = document.getElementById('process-btn');
const outputArea = document.getElementById('output-area');
const itemDropdown = document.getElementById('item-dropdown');
const autoDetectBtn = document.getElementById('auto-detect-btn');
const dataRangeInput = document.getElementById('data-range-input');
const headerRowsInput = document.getElementById('header-rows-input');
const loadHeadersBtn = document.getElementById('load-headers-btn');

// --- Event Listeners ---
dropArea.addEventListener('click', () => fileInput.click());
['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eName => dropArea.addEventListener(eName, e => {
    e.preventDefault();
    dropArea.classList.toggle('highlight', eName === 'dragenter' || eName === 'dragover');
}));
dropArea.addEventListener('drop', e => { if(e.dataTransfer.files.length) handleFiles(e.dataTransfer.files); });
fileInput.addEventListener('change', e => { if(e.target.files.length) handleFiles(e.target.files); });
autoDetectBtn.addEventListener('click', autoDetectBestRange);
loadHeadersBtn.addEventListener('click', loadHeadersAndMapping);
processBtn.addEventListener('click', processData);
itemDropdown.addEventListener('change', renderItemDetailView);
[dataRangeInput, headerRowsInput].forEach(input => input.addEventListener('input', resetMappings));

function resetUI() {
    resetMappings();
    dataRangeInput.value = '';
    headerRowsInput.value = '1';
    loadHeadersBtn.disabled = true;
}

function resetMappings() {
    mappingFields.style.display = 'none';
    mappingPlaceholder.style.display = 'block';
    state.columnMappings = [];
    processBtn.disabled = true;
    outputArea.style.display = 'none';
}

async function handleFiles(fileList) {
    resetUI();
    previewArea.innerHTML = '<p style="text-align:center;">讀取中...</p>';
    try {
        state.workbooks = await Promise.all(Array.from(fileList).map(readFile));
        fileListDisplay.innerHTML = '已載入檔案：' + state.workbooks.map(wb => `<li>- ${wb.file.name}</li>`).join('');
        generatePreview(state.workbooks[0].workbook.Sheets[state.workbooks[0].workbook.SheetNames[0]]);
        autoDetectBtn.disabled = false;
    } catch(err) {
        previewArea.innerHTML = `<p style="color:red;text-align:center;">檔案解析失敗：${err.message}</p>`;
        autoDetectBtn.disabled = true;
    }
}

function readFile(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = e => resolve({file, workbook: XLSX.read(e.target.result, {type: 'array'})});
        reader.onerror = err => reject(err);
        reader.readAsArrayBuffer(file);
    });
}

function generatePreview(sheet) {
    if (!sheet || !sheet['!ref']) return previewArea.innerHTML = '<p>工作表為空</p>';
    const range = XLSX.utils.decode_range(sheet['!ref']);
    range.e.r = Math.min(range.e.r, range.s.r + 100);
    const data = XLSX.utils.sheet_to_json(sheet, { header: 1, range: range, defval: '' });
    let html = '<table><thead><tr><th></th>';
    for (let C = range.s.c; C <= range.e.c; ++C) html += `<th>${XLSX.utils.encode_col(C)}</th>`;
    html += '</tr></thead><tbody>';
    data.forEach((row, i) => {
        html += `<tr><th>${range.s.r + i + 1}</th>`;
        (row).forEach(cell => html += `<td>${cell ?? ''}</td>`);
        html += '</tr>';
    });
    previewArea.innerHTML = html + '</tbody></table>';
}

const isRowEmpty = (row) => !row || row.every(cell => cell == null || String(cell).trim() === '');

function autoDetectBestRange() {
    if (state.workbooks.length === 0) return alert('請先上傳檔案');
    resetMappings();
    
    const sheet = state.workbooks[0].workbook.Sheets[state.workbooks[0].workbook.SheetNames[0]];
    if (!sheet || !sheet['!ref']) return alert("工作表為空或無法讀取。");
    
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: null });
    if (rows.length === 0) return alert("工作表內沒有資料。");

    let headerRowIdx = -1;
    for (let i = 0; i < Math.min(rows.length, 30); i++) {
        const nonEmptyCells = rows[i]?.filter(cell => cell != null && String(cell).trim() !== '').length || 0;
        if (nonEmptyCells > 2) {
            headerRowIdx = i;
            break;
        }
    }
    if (headerRowIdx === -1) return alert("找不到有效的標頭列，請手動輸入範圍。");

    let lastDataRowIdx = headerRowIdx;
    for (let i = rows.length - 1; i > headerRowIdx; i--) {
        if (!isRowEmpty(rows[i])) { lastDataRowIdx = i; break; }
    }

    let firstCol = Infinity, lastCol = -1;
    for (let r = headerRowIdx; r <= lastDataRowIdx; r++) {
        if (!rows[r]) continue;
        rows[r].forEach((cell, c) => {
            if (cell != null && String(cell).trim() !== '') {
                firstCol = Math.min(firstCol, c);
                lastCol = Math.max(lastCol, c);
            }
        });
    }

    if (firstCol > lastCol) return alert("在找到的標頭下找不到有效的資料欄。");

    const rangeStr = XLSX.utils.encode_range({ s: { r: headerRowIdx, c: firstCol }, e: { r: lastDataRowIdx, c: lastCol } });
    dataRangeInput.value = rangeStr;
    loadHeadersBtn.disabled = false;
    alert(`已偵測到範圍：${rangeStr}\n請確認「標頭佔用列數」是否正確，然後點擊「2. 讀取欄位」。`);
}

function unmergeAndFill(data, sheet, range) {
    (sheet['!merges'] || []).forEach(merge => {
        if (merge.s.c > range.e.c || merge.e.c < range.s.c || merge.s.r > range.e.r || merge.e.r < range.s.r) return;
        const s = { r: merge.s.r - range.s.r, c: merge.s.c - range.s.c };
        if (s.r < 0 || s.c < 0 || s.r >= data.length || !data[s.r]) return;
        const val = data[s.r][s.c];
        for (let r = s.r; r <= merge.e.r - range.s.r; r++) {
            if (r >= data.length) break;
            if (!data[r]) data[r] = [];
            for (let c = s.c; c <= merge.e.c - range.s.c; c++) data[r][c] = val;
        }
    });
    return data;
}

function loadHeadersAndMapping() {
    const rangeStr = dataRangeInput.value.trim().toUpperCase();
    const headerRowCount = parseInt(headerRowsInput.value, 10);
    if (!rangeStr) return alert('請先偵測或手動輸入總資料範圍');
    if (isNaN(headerRowCount) || headerRowCount < 1) return alert('標頭佔用列數必須是至少為 1 的數字');

    try {
        const sheet = state.workbooks[0].workbook.Sheets[state.workbooks[0].workbook.SheetNames[0]];
        const range = XLSX.utils.decode_range(rangeStr);
        if (headerRowCount > (range.e.r - range.s.r + 1)) return alert("標頭列數不能大於總範圍的列數。");
        
        const headerBlockRange = { s: range.s, e: { c: range.e.c, r: range.s.r + headerRowCount - 1 } };
        let headerData = XLSX.utils.sheet_to_json(sheet, { header: 1, range: headerBlockRange, defval: null });
        headerData = unmergeAndFill(headerData, sheet, headerBlockRange);
        
        const finalHeaders = Array.from({ length: range.e.c - range.s.c + 1 }, (_, c) => {
            const headerParts = [];
            for (let r = 0; r < headerRowCount; r++) {
                const cellValue = headerData[r]?.[c] || '';
                if (cellValue && !headerParts.includes(String(cellValue).trim())) {
                    headerParts.push(String(cellValue).trim());
                }
            }
            return headerParts.join(' ').trim();
        });

        state.columnMappings = finalHeaders.map((header, i) => {
            const col = range.s.c + i, excelCol = XLSX.utils.encode_col(col);
            let autoRole = /科目|項目|名稱|類別|品項/.test(header) ? 'key' : 'value';
            if (!header) autoRole = 'ignore';
            return { excelCol, autoHeader: header || `(空白欄 ${excelCol})`, customName: header || '', role: autoRole, include: autoRole !== 'ignore' };
        });
        
        renderMappingTable();
    } catch(err) {
        alert(`讀取欄位失敗：${err.message}`);
        resetMappings();
    }
}

function renderMappingTable() {
    let html = `<div class="mapping-table-container"><table class="mapping-table"><thead><tr><th>Excel 欄位</th><th>合併後標頭</th><th>報表欄位名稱</th><th>欄位角色</th><th>使用</th></tr></thead><tbody>`;
    state.columnMappings.forEach((col, idx) => {
        html += `
            <tr>
                <td>${col.excelCol}</td><td>${col.autoHeader}</td>
                <td><input type="text" data-idx="${idx}" class="custom-name-input" value="${col.customName}"></td>
                <td>
                    <select data-idx="${idx}" class="role-select">
                        <option value="ignore" ${col.role === 'ignore' ? 'selected' : ''}>忽略</option>
                        <option value="key" ${col.role === 'key' ? 'selected' : ''}>主鍵欄位</option>
                        <option value="value" ${col.role === 'value' ? 'selected' : ''}>加總欄位</option>
                    </select>
                </td>
                <td><input type="checkbox" data-idx="${idx}" class="include-checkbox" ${col.include ? 'checked' : ''}></td>
            </tr>`;
    });
    mappingFields.innerHTML = html + `</tbody></table></div>`;
    
    mappingFields.querySelectorAll('.custom-name-input').forEach(i=>i.addEventListener('input', e=>state.columnMappings[e.target.dataset.idx].customName=e.target.value.trim()));
    mappingFields.querySelectorAll('.include-checkbox').forEach(c=>c.addEventListener('change', e=>state.columnMappings[e.target.dataset.idx].include=e.target.checked));
    mappingFields.querySelectorAll('.role-select').forEach(s=>s.addEventListener('change', e=>{
        const idx = e.target.dataset.idx, isIgnored = e.target.value === 'ignore';
        state.columnMappings[idx].role = e.target.value;
        state.columnMappings[idx].include = !isIgnored;
        mappingFields.querySelector(`.include-checkbox[data-idx="${idx}"]`).checked = !isIgnored;
    }));
    
    mappingFields.style.display = 'block';
    mappingPlaceholder.style.display = 'none';
    processBtn.disabled = false;
}

function processData() {
    const keyColumns = state.columnMappings.filter(c => c.role === 'key' && c.include);
    const valueColumns = state.columnMappings.filter(c => c.role === 'value' && c.include);
    if (keyColumns.length !== 1) return alert('必須且只能選擇一個主鍵欄位！');
    if (valueColumns.length === 0) return alert('請至少選擇一個加總欄位！');
    
    const keyCol = keyColumns[0], keyName = keyCol.customName || keyCol.autoHeader;
    const rangeStr = dataRangeInput.value.trim();
    if (!rangeStr) return alert('請輸入總資料範圍！');
    
    state.allFileData = [], state.summaryData = new Map();
    try {
        const range = XLSX.utils.decode_range(rangeStr);
        const headerRowCount = parseInt(headerRowsInput.value, 10);
        const dataRange = { s: { r: range.s.r + headerRowCount, c: range.s.c }, e: range.e };
        
        state.workbooks.forEach(wb => {
            const sheet = wb.workbook.Sheets[wb.workbook.SheetNames[0]];
            if (!sheet || !sheet['!ref']) return;
            
            let dataRows = XLSX.utils.sheet_to_json(sheet, { header: 1, range: dataRange, defval: null });
            dataRows = unmergeAndFill(dataRows, sheet, dataRange);

            const keyColIndex = XLSX.utils.decode_col(keyCol.excelCol) - range.s.c;
            const transformedData = dataRows.map(row => {
                if (isRowEmpty(row)) return null;
                const keyValue = row[keyColIndex];
                if (keyValue == null || String(keyValue).trim() === '') return null;
                
                const dataRow = { [keyName]: String(keyValue).trim() };
                valueColumns.forEach(valCol => {
                    const colIdx = XLSX.utils.decode_col(valCol.excelCol) - range.s.c;
                    dataRow[valCol.customName || valCol.autoHeader] = toNumber(row[colIdx]);
                });
                return dataRow;
            }).filter(Boolean);

            state.allFileData.push({ fileName: wb.file.name, data: transformedData });
        });

        // Aggregation
        state.allFileData.forEach(file => file.data.forEach(row => {
            const key = row[keyName];
            let summaryRow = state.summaryData.get(key);
            if (!summaryRow) {
                summaryRow = { [keyName]: key };
                valueColumns.forEach(c => summaryRow[c.customName || c.autoHeader] = 0);
                state.summaryData.set(key, summaryRow);
            }
            valueColumns.forEach(c => summaryRow[c.customName || c.autoHeader] += row[c.customName || c.autoHeader] || 0);
        }));
        
        renderSummaryView(keyName, valueColumns);
        renderIndividualViews(keyName, valueColumns);
        renderItemView(keyName, valueColumns);
        outputArea.style.display = 'block';
        alert(`處理完成！\n共處理 ${state.workbooks.length} 個檔案\n彙總了 ${state.summaryData.size} 個項目`);
    } catch(err) {
        alert(`處理資料錯誤：${err.message}`);
        console.error(err);
    }
}


function toNumber(val) {
    if (val == null || val === '') return 0;
    if (typeof val === 'number') return val;
    return parseFloat(String(val).replace(/,/g, '')) || 0;
}

function generateHtmlTable(data, headers, formatNumbers = false) {
    let html = '<table class="report-table"><thead><tr>' + headers.map(h => `<th>${h}</th>`).join('') + '</tr></thead><tbody>';
    data.forEach(row => {
        html += '<tr>' + headers.map(h => {
            const val = row[h];
            // Added check for the '檔案名稱' column to avoid formatting it as a number
            if (formatNumbers && h !== '檔案名稱' && typeof val === 'number') {
                return `<td class="number">${val.toLocaleString('en-US', {minimumFractionDigits: 0, maximumFractionDigits: 0})}</td>`;
            }
            return `<td>${val ?? ''}</td>`;
        }).join('') + '</tr>';
    });
    return html + '</tbody></table>';
}

function renderSummaryView(keyName, valueColumns) {
    const headers = [keyName, ...valueColumns.map(c => c.customName || c.autoHeader)];
    document.getElementById('summary-view').innerHTML = '<h3>所有檔案加總結果</h3>' + generateHtmlTable(Array.from(state.summaryData.values()), headers, true);
}

function renderIndividualViews(keyName, valueColumns) {
    const headers = [keyName, ...valueColumns.map(c => c.customName || c.autoHeader)];
    document.getElementById('individual-view').innerHTML = state.allFileData.map(file => 
        `<div class="individual-report"><h3>${file.fileName}</h3>${generateHtmlTable(file.data, headers, true)}</div>`
    ).join('');
}

// --- UPDATED: This function now sets up the ITEM query dropdown ---
function renderItemView(keyName, valueColumns) {
    itemDropdown.innerHTML = '<option value="">--- 請選擇要查詢的項目 ---</option>' + 
        Array.from(state.summaryData.keys()).sort().map(item => `<option value="${item}">${item}</option>`).join('');
    
    document.getElementById('item-detail-table').innerHTML = '';
    itemDropdown.dataset.keyName = keyName;
    itemDropdown.dataset.valueColumns = JSON.stringify(valueColumns.map(c => c.customName || c.autoHeader));
}

// --- UPDATED: This function now renders a comparison table for the selected ITEM ---
function renderItemDetailView() {
    const selectedItem = itemDropdown.value;
    const container = document.getElementById('item-detail-table');
    if (!selectedItem) {
        container.innerHTML = '';
        return;
    }

    const keyName = itemDropdown.dataset.keyName;
    const valueColNames = JSON.parse(itemDropdown.dataset.valueColumns);
    const headers = ['檔案名稱', ...valueColNames];
    
    // 1. Gather data from each file for the selected item
    const data = state.allFileData.map(file => {
        const row = file.data.find(r => r[keyName] === selectedItem);
        // Create a row for the table even if the item is not found or values are zero
        const dataRow = { '檔案名稱': file.fileName };
        valueColNames.forEach(colName => {
            dataRow[colName] = row ? (row[colName] || 0) : 0;
        });
        return dataRow;
    });

    // 2. Get the pre-calculated summary (total) row
    const summaryRow = state.summaryData.get(selectedItem);
    if (summaryRow) {
        const totalRow = { '檔案名稱': '<strong>合計</strong>' };
        valueColNames.forEach(colName => {
            totalRow[colName] = summaryRow[colName] || 0;
        });
        // 3. Add the total row to the beginning of the data array
        data.unshift(totalRow);
    }
    
    container.innerHTML = `<h3>項目：${selectedItem}</h3>` + generateHtmlTable(data, headers, true);
}


function setupTabs() {
    document.querySelector('.view-tabs').addEventListener('click', e => {
        if (e.target.classList.contains('tab-btn')) {
            const targetView = e.target.dataset.view;
            document.querySelectorAll('.tab-btn').forEach(btn => btn.classList.toggle('active', btn.dataset.view === targetView));
            document.querySelectorAll('.view-pane').forEach(pane => pane.classList.toggle('active', pane.id === targetView));
        }
    });
}

setupTabs();
</script>
</body>
</html>
