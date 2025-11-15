<!DOCTYPE html>
<html lang="zh-Hant">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>基金資料彙總報告產生器 v13.0</title>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
    <style>
        * { box-sizing: border-box; margin: 0; padding: 0; }
        
        body {
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif;
            line-height: 1.6;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            min-height: 100vh;
            padding: 20px;
        }
        
        #app-container {
            max-width: 1400px;
            margin: 0 auto;
            background: #fff;
            border-radius: 16px;
            box-shadow: 0 20px 60px rgba(0,0,0,0.3);
            overflow: hidden;
        }
        
        /* 頂部標題區 */
        .app-header {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 30px 40px;
            text-align: center;
        }
        
        .app-header h1 {
            font-size: 2em;
            font-weight: 700;
            margin-bottom: 8px;
            text-shadow: 0 2px 4px rgba(0,0,0,0.2);
        }
        
        .app-header .version {
            font-size: 0.9em;
            opacity: 0.9;
        }
        
        /* 進度指示器 */
        .progress-steps {
            display: flex;
            justify-content: space-between;
            padding: 30px 40px;
            background: #f8f9fa;
            border-bottom: 1px solid #e0e0e0;
        }
        
        .step {
            flex: 1;
            text-align: center;
            position: relative;
            padding: 0 10px;
        }
        
        .step:not(:last-child)::after {
            content: '';
            position: absolute;
            top: 20px;
            right: -50%;
            width: 100%;
            height: 2px;
            background: #ddd;
            z-index: 0;
        }
        
        .step.active:not(:last-child)::after {
            background: #667eea;
        }
        
        .step-number {
            width: 40px;
            height: 40px;
            border-radius: 50%;
            background: #ddd;
            color: #999;
            display: inline-flex;
            align-items: center;
            justify-content: center;
            font-weight: 700;
            margin-bottom: 8px;
            position: relative;
            z-index: 1;
            transition: all 0.3s;
        }
        
        .step.active .step-number {
            background: #667eea;
            color: white;
            transform: scale(1.1);
        }
        
        .step.completed .step-number {
            background: #28a745;
            color: white;
        }
        
        .step-label {
            font-size: 0.85em;
            color: #666;
            font-weight: 500;
        }
        
        .step.active .step-label {
            color: #667eea;
            font-weight: 600;
        }
        
        /* 內容區域 */
        .app-content {
            padding: 40px;
        }
        
        .section {
            margin-bottom: 40px;
            animation: fadeIn 0.5s;
        }
        
        @keyframes fadeIn {
            from { opacity: 0; transform: translateY(10px); }
            to { opacity: 1; transform: translateY(0); }
        }
        
        .section-title {
            font-size: 1.4em;
            color: #333;
            margin-bottom: 20px;
            padding-bottom: 10px;
            border-bottom: 3px solid #667eea;
            display: inline-block;
        }
        
        /* 檔案上傳區 */
        #drop-area {
            border: 3px dashed #667eea;
            border-radius: 12px;
            padding: 60px 40px;
            text-align: center;
            background: linear-gradient(135deg, #f8f9ff 0%, #f0f2ff 100%);
            cursor: pointer;
            transition: all 0.3s;
            position: relative;
            overflow: hidden;
        }
        
        #drop-area::before {
            content: '📁';
            font-size: 4em;
            display: block;
            margin-bottom: 15px;
            animation: bounce 2s infinite;
        }
        
        @keyframes bounce {
            0%, 100% { transform: translateY(0); }
            50% { transform: translateY(-10px); }
        }
        
        #drop-area:hover {
            background: linear-gradient(135deg, #e8ebff 0%, #dce0ff 100%);
            border-color: #5568d3;
            transform: translateY(-2px);
            box-shadow: 0 5px 15px rgba(102, 126, 234, 0.3);
        }
        
        #drop-area.drag-over {
            background: linear-gradient(135deg, #d0d8ff 0%, #c0cbff 100%);
            border-color: #4050bf;
            transform: scale(1.02);
        }
        
        .drop-text {
            font-size: 1.2em;
            color: #667eea;
            font-weight: 600;
            margin-bottom: 10px;
        }
        
        .drop-hint {
            color: #888;
            font-size: 0.9em;
        }
        
        #file-input { display: none; }
        
        /* 檔案列表 */
        .file-list {
            margin-top: 20px;
            background: #f8f9fa;
            border-radius: 8px;
            padding: 15px;
        }
        
        .file-item {
            display: flex;
            align-items: center;
            padding: 10px 15px;
            background: white;
            border-radius: 6px;
            margin-bottom: 8px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.05);
        }
        
        .file-item::before {
            content: '📄';
            margin-right: 10px;
            font-size: 1.2em;
        }
        
        .file-item span {
            flex: 1;
            color: #333;
            font-size: 0.9em;
        }
        
        /* 預覽區 */
        #preview-area {
            background: #f8f9fa;
            border: 1px solid #e0e0e0;
            border-radius: 8px;
            padding: 20px;
            max-height: 450px;
            overflow: auto;
        }
        
        #preview-area::-webkit-scrollbar {
            width: 8px;
            height: 8px;
        }
        
        #preview-area::-webkit-scrollbar-track {
            background: #f1f1f1;
            border-radius: 4px;
        }
        
        #preview-area::-webkit-scrollbar-thumb {
            background: #667eea;
            border-radius: 4px;
        }
        
        #preview-area table {
            width: 100%;
            border-collapse: collapse;
            background: white;
            font-size: 0.85em;
        }
        
        #preview-area th, #preview-area td {
            border: 1px solid #ddd;
            padding: 8px 12px;
            text-align: left;
        }
        
        #preview-area thead th {
            background: #667eea;
            color: white;
            position: sticky;
            top: 0;
            z-index: 10;
            font-weight: 600;
        }
        
        #preview-area tbody th {
            background: #f1f3f5;
            font-weight: 600;
        }
        
        /* 表單控制項 */
        .form-row {
            display: flex;
            gap: 20px;
            margin-bottom: 20px;
            align-items: flex-end;
            flex-wrap: wrap;
        }
        
        .form-group {
            flex: 1;
            min-width: 200px;
        }
        
        .form-label {
            display: block;
            font-weight: 600;
            color: #555;
            margin-bottom: 8px;
            font-size: 0.9em;
        }
        
        input[type="text"],
        input[type="number"],
        select {
            width: 100%;
            padding: 12px 15px;
            border: 2px solid #e0e0e0;
            border-radius: 8px;
            font-size: 1em;
            transition: all 0.3s;
            background: white;
        }
        
        input[type="text"]:focus,
        input[type="number"]:focus,
        select:focus {
            outline: none;
            border-color: #667eea;
            box-shadow: 0 0 0 3px rgba(102, 126, 234, 0.1);
        }
        
        /* 按鈕樣式 */
        .btn {
            padding: 12px 30px;
            border: none;
            border-radius: 8px;
            font-size: 1em;
            font-weight: 600;
            cursor: pointer;
            transition: all 0.3s;
            display: inline-flex;
            align-items: center;
            gap: 8px;
        }
        
        .btn-primary {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            box-shadow: 0 4px 15px rgba(102, 126, 234, 0.4);
        }
        
        .btn-primary:hover:not(:disabled) {
            transform: translateY(-2px);
            box-shadow: 0 6px 20px rgba(102, 126, 234, 0.5);
        }
        
        .btn-success {
            background: linear-gradient(135deg, #28a745 0%, #20c997 100%);
            color: white;
            box-shadow: 0 4px 15px rgba(40, 167, 69, 0.4);
        }
        
        .btn-success:hover:not(:disabled) {
            transform: translateY(-2px);
            box-shadow: 0 6px 20px rgba(40, 167, 69, 0.5);
        }
        
        .btn:disabled {
            background: #ccc;
            cursor: not-allowed;
            transform: none;
            box-shadow: none;
        }
        
        .btn::before {
            content: attr(data-icon);
            font-size: 1.2em;
        }
        
        /* 映射表格 */
        .mapping-table-wrapper {
            background: #f8f9fa;
            border-radius: 8px;
            padding: 20px;
            overflow-x: auto;
        }
        
        .mapping-table {
            width: 100%;
            border-collapse: collapse;
            background: white;
            border-radius: 8px;
            overflow: hidden;
            box-shadow: 0 2px 8px rgba(0,0,0,0.1);
        }
        
        .mapping-table thead {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
        }
        
        .mapping-table th {
            padding: 15px;
            text-align: left;
            font-weight: 600;
            font-size: 0.9em;
        }
        
        .mapping-table td {
            padding: 12px 15px;
            border-bottom: 1px solid #e0e0e0;
        }
        
        .mapping-table tr:hover {
            background: #f8f9ff;
        }
        
        .mapping-table input[type="text"],
        .mapping-table select {
            padding: 8px 12px;
            border: 1px solid #ddd;
            border-radius: 6px;
        }
        
        .mapping-table input[type="checkbox"] {
            width: 20px;
            height: 20px;
            cursor: pointer;
            accent-color: #667eea;
        }
        
        .excel-col {
            font-family: 'Courier New', monospace;
            font-weight: 700;
            color: #667eea;
            font-size: 1.1em;
        }
        
        /* 輸出區域 */
        .output-section {
            display: none;
            animation: fadeIn 0.5s;
        }
        
        .view-tabs {
            display: flex;
            gap: 10px;
            margin-bottom: 20px;
            border-bottom: 2px solid #e0e0e0;
            padding-bottom: 0;
        }
        
        .tab-btn {
            padding: 12px 24px;
            border: none;
            background: transparent;
            color: #666;
            font-size: 1em;
            font-weight: 600;
            cursor: pointer;
            border-bottom: 3px solid transparent;
            transition: all 0.3s;
            margin-bottom: -2px;
        }
        
        .tab-btn:hover {
            color: #667eea;
            background: #f8f9ff;
        }
        
        .tab-btn.active {
            color: #667eea;
            border-bottom-color: #667eea;
            background: #f8f9ff;
        }
        
        .view-pane {
            display: none;
            animation: fadeIn 0.3s;
        }
        
        .view-pane.active {
            display: block;
        }
        
        /* 報表表格 */
        .report-table {
            width: 100%;
            border-collapse: collapse;
            background: white;
            box-shadow: 0 2px 8px rgba(0,0,0,0.1);
            border-radius: 8px;
            overflow: hidden;
        }
        
        .report-table thead {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
        }
        
        .report-table th {
            padding: 15px;
            text-align: left;
            font-weight: 600;
        }
        
        .report-table td {
            padding: 12px 15px;
            border-bottom: 1px solid #e0e0e0;
        }
        
        .report-table tbody tr:hover {
            background: #f8f9ff;
        }
        
        .report-table td.number {
            text-align: right;
            font-family: 'Courier New', monospace;
            font-weight: 600;
            color: #667eea;
        }
        
        .report-table .total-row {
            background: #fff3cd !important;
            font-weight: 700;
        }
        
        .report-table .total-row td {
            border-top: 3px solid #ffc107;
            border-bottom: 3px solid #ffc107;
        }
        
        /* 提示訊息 */
        .alert {
            padding: 15px 20px;
            border-radius: 8px;
            margin-bottom: 20px;
            display: flex;
            align-items: center;
            gap: 12px;
        }
        
        .alert-info {
            background: #d1ecf1;
            border-left: 4px solid #17a2b8;
            color: #0c5460;
        }
        
        .alert-success {
            background: #d4edda;
            border-left: 4px solid #28a745;
            color: #155724;
        }
        
        .alert::before {
            content: 'ℹ️';
            font-size: 1.5em;
        }
        
        .alert-success::before {
            content: '✅';
        }
        
        /* 空狀態 */
        .empty-state {
            text-align: center;
            padding: 60px 20px;
            color: #999;
        }
        
        .empty-state::before {
            content: '📊';
            font-size: 4em;
            display: block;
            margin-bottom: 15px;
            opacity: 0.5;
        }
        
        /* 查詢下拉選單 */
        .query-select {
            width: 100%;
            max-width: 500px;
            padding: 12px 15px;
            border: 2px solid #e0e0e0;
            border-radius: 8px;
            font-size: 1.1em;
            margin-bottom: 20px;
            background: white;
            cursor: pointer;
        }
        
        .query-select:focus {
            border-color: #667eea;
            box-shadow: 0 0 0 3px rgba(102, 126, 234, 0.1);
        }
        
        /* 響應式 */
        @media (max-width: 768px) {
            .app-content {
                padding: 20px;
            }
            
            .progress-steps {
                flex-wrap: wrap;
            }
            
            .step {
                flex-basis: 50%;
                margin-bottom: 20px;
            }
            
            .form-row {
                flex-direction: column;
            }
            
            .app-header h1 {
                font-size: 1.5em;
            }
        }
    </style>
</head>
<body>
    <div id="app-container">
        <!-- 頂部標題 -->
        <div class="app-header">
            <h1>📊 基金資料彙總報告產生器</h1>
            <div class="version">Version 13.0 - UX Enhanced</div>
        </div>
        
        <!-- 進度指示器 -->
        <div class="progress-steps">
            <div class="step" id="step-1">
                <div class="step-number">1</div>
                <div class="step-label">上傳檔案</div>
            </div>
            <div class="step" id="step-2">
                <div class="step-number">2</div>
                <div class="step-label">偵測範圍</div>
            </div>
            <div class="step" id="step-3">
                <div class="step-number">3</div>
                <div class="step-label">設定欄位</div>
            </div>
            <div class="step" id="step-4">
                <div class="step-number">4</div>
                <div class="step-label">查看報告</div>
            </div>
        </div>
        
        <!-- 主要內容 -->
        <div class="app-content">
            <!-- 步驟 1: 上傳檔案 -->
            <div class="section" id="section-upload">
                <h2 class="section-title">上傳 Excel 檔案</h2>
                <div id="drop-area">
                    <input type="file" id="file-input" accept=".xlsx,.xls" multiple>
                    <div class="drop-text">拖曳多個 Excel 檔案至此</div>
                    <div class="drop-hint">或點擊選擇檔案 (支援 .xlsx, .xls)</div>
                </div>
                <div id="file-list-container"></div>
            </div>
            
            <!-- 步驟 2: 預覽與偵測 -->
            <div class="section" id="section-preview" style="display:none;">
                <h2 class="section-title">檔案預覽與範圍偵測</h2>
                <div id="preview-area">
                    <div class="empty-state">尚未載入檔案</div>
                </div>
                <div class="form-row" style="margin-top:20px;">
                    <button class="btn btn-success" id="auto-detect-btn" data-icon="🔍">自動偵測範圍</button>
                </div>
            </div>
            
            <!-- 步驟 3: 設定範圍 -->
            <div class="section" id="section-range" style="display:none;">
                <h2 class="section-title">確認資料範圍</h2>
                <div id="detect-result"></div>
                <div class="form-row">
                    <div class="form-group">
                        <label class="form-label">資料範圍 (含標頭)</label>
                        <input type="text" id="data-range-input" placeholder="例如: A5:G50">
                    </div>
                    <div class="form-group">
                        <label class="form-label">標頭佔用列數</label>
                        <input type="number" id="header-rows-input" value="1" min="1">
                    </div>
                    <button class="btn btn-primary" id="load-headers-btn" data-icon="📋">讀取欄位</button>
                </div>
            </div>
            
            <!-- 步驟 4: 欄位映射 -->
            <div class="section" id="section-mapping" style="display:none;">
                <h2 class="section-title">欄位設定與對應</h2>
                <div class="alert alert-info">
                    💡 請選擇一個<strong>主鍵欄位</strong>（項目/科目），並選擇要<strong>加總的數值欄位</strong>。可自訂報表中的欄位名稱。
                </div>
                <div id="mapping-fields"></div>
                <div class="form-row" style="margin-top:20px;">
                    <button class="btn btn-primary" id="process-btn" data-icon="⚡">開始彙總處理</button>
                </div>
            </div>
            
            <!-- 步驟 5: 輸出結果 -->
            <div class="output-section" id="output-area">
                <h2 class="section-title">彙總報告結果</h2>
                <div class="view-tabs">
                    <button class="tab-btn active" data-view="summary-view">📊 加總總表</button>
                    <button class="tab-btn" data-view="file-query-view">📁 檔案查詢</button>
                    <button class="tab-btn" data-view="item-view">🔍 項目查詢</button>
                </div>
                <div id="view-content">
                    <div id="summary-view" class="view-pane active"></div>
                    <div id="file-query-view" class="view-pane">
                        <select class="query-select" id="file-dropdown">
                            <option value="">--- 請選擇要查詢的檔案 ---</option>
                        </select>
                        <div id="file-detail-table"></div>
                    </div>
                    <div id="item-view" class="view-pane">
                        <select class="query-select" id="item-dropdown">
                            <option value="">--- 請選擇要查詢的項目 ---</option>
                        </select>
                        <div id="item-detail-table"></div>
                    </div>
                </div>
            </div>
        </div>
    </div>

<script>
const state = { workbooks: [], columnMappings: [], allFileData:[], summaryData:new Map(), orderedItemKeys: [] };

// DOM Elements
const dropArea = document.getElementById('drop-area');
const fileInput = document.getElementById('file-input');
const fileListContainer = document.getElementById('file-list-container');
const previewArea = document.getElementById('preview-area');
const mappingFields = document.getElementById('mapping-fields');
const processBtn = document.getElementById('process-btn');
const outputArea = document.getElementById('output-area');
const itemDropdown = document.getElementById('item-dropdown');
const autoDetectBtn = document.getElementById('auto-detect-btn');
const dataRangeInput = document.getElementById('data-range-input');
const headerRowsInput = document.getElementById('header-rows-input');
const loadHeadersBtn = document.getElementById('load-headers-btn');
const fileDropdown = document.getElementById('file-dropdown');
const fileDetailTable = document.getElementById('file-detail-table');
const detectResult = document.getElementById('detect-result');

// 進度步驟更新
function updateStep(stepNum, status = 'active') {
    document.querySelectorAll('.step').forEach((step, i) => {
        step.classList.remove('active', 'completed');
        if (i + 1 < stepNum) step.classList.add('completed');
        if (i + 1 === stepNum) step.classList.add(status);
    });
}

// 事件監聽
dropArea.addEventListener('click', () => fileInput.click());
['dragenter', 'dragover'].forEach(e => dropArea.addEventListener(e, evt => {
    evt.preventDefault();
    dropArea.classList.add('drag-over');
}));
['dragleave', 'drop'].forEach(e => dropArea.addEventListener(e, evt => {
    evt.preventDefault();
    dropArea.classList.remove('drag-over');
}));
dropArea.addEventListener('drop', e => { 
    if(e.dataTransfer.files.length) handleFiles(e.dataTransfer.files); 
});
fileInput.addEventListener('change', e => { 
    if(e.target.files.length) handleFiles(e.target.files); 
});
autoDetectBtn.addEventListener('click', autoDetectBestRange);
loadHeadersBtn.addEventListener('click', loadHeadersAndMapping);
processBtn.addEventListener('click', processData);
itemDropdown.addEventListener('change', renderItemDetailView);
fileDropdown.addEventListener('change', renderFileDetailView);
[dataRangeInput, headerRowsInput].forEach(input => input.addEventListener('input', resetMappings));

function resetUI() {
    resetMappings();
    dataRangeInput.value = '';
    headerRowsInput.value = '1';
    detectResult.innerHTML = '';
    document.getElementById('section-preview').style.display = 'none';
    document.getElementById('section-range').style.display = 'none';
    document.getElementById('section-mapping').style.display = 'none';
    updateStep(1);
}

function resetMappings() {
    document.getElementById('section-mapping').style.display = 'none';
    state.columnMappings = [];
    processBtn.disabled = true;
    outputArea.style.display = 'none';
}

async function handleFiles(fileList) {
    resetUI();
    previewArea.innerHTML = '<div class="empty-state">正在讀取檔案...</div>';
    
    try {
        state.workbooks = await Promise.all(Array.from(fileList).map(readFile));
        
        const fileListHtml = `
            <div class="file-list">
                <div style="font-weight:600; margin-bottom:10px; color:#667eea;">
                    ✓ 已載入 ${state.workbooks.length} 個檔案
                </div>
                ${state.workbooks.map(wb => 
                    `<div class="file-item"><span>${wb.file.name}</span></div>`
                ).join('')}
            </div>
        `;
        fileListContainer.innerHTML = fileListHtml;
        
        generatePreview(state.workbooks[0].workbook.Sheets[state.workbooks[0].workbook.SheetNames[0]]);
        document.getElementById('section-preview').style.display = 'block';
        updateStep(2);
        
    } catch(err) {
        previewArea.innerHTML = `<div class="empty-state" style="color:#dc3545;">檔案解析失敗：${err.message}</div>`;
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
    if (!sheet || !sheet['!ref']) {
        previewArea.innerHTML = '<div class="empty-state">工作表為空</div>';
        return;
    }
    
    const range = XLSX.utils.decode_range(sheet['!ref']);
    range.e.r = Math.min(range.e.r, range.s.r + 100);
    const data = XLSX.utils.sheet_to_json(sheet, { header: 1, range: range, defval: '' });
    
    let html = '<table><thead><tr><th></th>';
    for (let C = range.s.c; C <= range.e.c; ++C) {
        html += `<th>${XLSX.utils.encode_col(C)}</th>`;
    }
    html += '</tr></thead><tbody>';
    
    data.forEach((row, i) => {
        html += `<tr><th>${range.s.r + i + 1}</th>`;
        (row || []).forEach(cell => html += `<td>${cell ?? ''}</td>`);
        html += '</tr>';
    });
    
    previewArea.innerHTML = html + '</tbody></table>';
}

const isRowEmpty = (row) => !row || row.every(cell => cell == null || String(cell).trim() === '');

function autoDetectBestRange() {
    if (state.workbooks.length === 0) return;
    resetMappings();
    
    const sheet = state.workbooks[0].workbook.Sheets[state.workbooks[0].workbook.SheetNames[0]];
    if (!sheet || !sheet['!ref']) {
        showDetectResult('工作表為空', 'error');
        return;
    }
    
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: null });
    if (rows.length === 0) {
        showDetectResult('工作表內沒有資料', 'error');
        return;
    }

    let headerRowIdx = -1;
    for (let i = 0; i < Math.min(rows.length, 30); i++) {
        if ((rows[i]?.filter(c => c != null && String(c).trim() !== '').length || 0) > 2) {
            headerRowIdx = i;
            break;
        }
    }
    
    if (headerRowIdx === -1) {
        showDetectResult('找不到有效的標頭列', 'error');
        return;
    }

    let lastDataRowIdx = headerRowIdx;
    for (let i = rows.length - 1; i > headerRowIdx; i--) {
        if (!isRowEmpty(rows[i])) {
            lastDataRowIdx = i;
            break;
        }
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
    
    if (firstCol > lastCol) {
        showDetectResult('找不到有效的資料欄', 'error');
        return;
    }
    
    const detectedRange = { s: { r: headerRowIdx, c: firstCol }, e: { r: lastDataRowIdx, c: lastCol } };
    const rangeStr = XLSX.utils.encode_range(detectedRange);
    dataRangeInput.value = rangeStr;

    const blockRows = XLSX.utils.sheet_to_json(sheet, { header: 1, range: detectedRange, defval: null });
    let firstDataRowIndex = -1;
    for (let i = 1; i < blockRows.length; i++) {
        if (blockRows[i] && blockRows[i].slice(1).some(cell => cell != null && !isNaN(Number(cell)))) {
            firstDataRowIndex = i;
            break;
        }
    }
    const headerCount = (firstDataRowIndex > 0) ? firstDataRowIndex : 1;
    headerRowsInput.value = headerCount;

    showDetectResult(`成功偵測到範圍：${rangeStr}，標頭佔 ${headerCount} 列`, 'success');
    document.getElementById('section-range').style.display = 'block';
    updateStep(2, 'completed');
    updateStep(3);
}

function showDetectResult(message, type) {
    const className = type === 'success' ? 'alert-success' : 'alert-info';
    detectResult.innerHTML = `<div class="alert ${className}">${message}</div>`;
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
            for (let c = s.c; c <= merge.e.c - range.s.c; c++) {
                data[r][c] = val;
            }
        }
    });
    return data;
}

function loadHeadersAndMapping() {
    const rangeStr = dataRangeInput.value.trim().toUpperCase();
    const headerRowCount = parseInt(headerRowsInput.value, 10);
    
    if (!rangeStr || isNaN(headerRowCount) || headerRowCount < 1) {
        showDetectResult('請輸入有效的範圍和標頭列數', 'error');
        return;
    }
    
    try {
        const sheet = state.workbooks[0].workbook.Sheets[state.workbooks[0].workbook.SheetNames[0]];
        const range = XLSX.utils.decode_range(rangeStr);
        
        if (headerRowCount > (range.e.r - range.s.r + 1)) {
            showDetectResult('標頭列數過大', 'error');
            return;
        }
        
        const headerBlockRange = { s: range.s, e: { c: range.e.c, r: range.s.r + headerRowCount - 1 } };
        let headerData = XLSX.utils.sheet_to_json(sheet, { header: 1, range: headerBlockRange, defval: null });
        headerData = unmergeAndFill(headerData, sheet, headerBlockRange);
        
        const finalHeaders = Array.from({ length: range.e.c - range.s.c + 1 }, (_, c) => 
            Array.from({ length: headerRowCount }, (_, r) => headerData[r]?.[c] || '')
            .map(s => String(s).trim())
            .filter((v, i, a) => v && a.indexOf(v) === i)
            .join(' ')
        );
        
        state.columnMappings = finalHeaders.map((header, i) => {
            const excelCol = XLSX.utils.encode_col(range.s.c + i);
            const autoRole = (i === 0) ? 'key' : (!header ? 'ignore' : 'value');
            return { 
                excelCol, 
                autoHeader: header || `(空白欄 ${excelCol})`, 
                customName: header || '', 
                role: autoRole, 
                include: autoRole !== 'ignore' 
            };
        });
        
        renderMappingTable();
        document.getElementById('section-mapping').style.display = 'block';
        updateStep(3, 'completed');
        
    } catch(err) {
        showDetectResult(`讀取欄位失敗：${err.message}`, 'error');
        resetMappings();
    }
}

function renderMappingTable() {
    let html = `
    <div class="mapping-table-wrapper">
        <table class="mapping-table">
            <thead>
                <tr>
                    <th style="width:100px;">Excel 欄位</th>
                    <th>原始標頭</th>
                    <th>報表欄位名稱</th>
                    <th style="width:150px;">欄位角色</th>
                    <th style="width:80px;">使用</th>
                </tr>
            </thead>
            <tbody>
    `;
    
    state.columnMappings.forEach((col, idx) => {
        html += `
            <tr>
                <td><span class="excel-col">${col.excelCol}</span></td>
                <td>${col.autoHeader}</td>
                <td>
                    <input type="text" 
                           data-idx="${idx}" 
                           class="custom-name-input"
                           value="${col.customName}"
                           placeholder="輸入自訂名稱">
                </td>
                <td>
                    <select data-idx="${idx}" class="role-select">
                        <option value="ignore" ${col.role === 'ignore' ? 'selected' : ''}>忽略</option>
                        <option value="key" ${col.role === 'key' ? 'selected' : ''}>主鍵欄位</option>
                        <option value="value" ${col.role === 'value' ? 'selected' : ''}>加總欄位</option>
                    </select>
                </td>
                <td style="text-align:center;">
                    <input type="checkbox" 
                           data-idx="${idx}" 
                           class="include-checkbox"
                           ${col.include ? 'checked' : ''}>
                </td>
            </tr>
        `;
    });
    
    mappingFields.innerHTML = html + `</tbody></table></div>`;
    
    mappingFields.querySelectorAll('.custom-name-input').forEach(i => 
        i.addEventListener('input', e => state.columnMappings[e.target.dataset.idx].customName = e.target.value.trim())
    );
    
    mappingFields.querySelectorAll('.include-checkbox').forEach(c => 
        c.addEventListener('change', e => state.columnMappings[e.target.dataset.idx].include = e.target.checked)
    );
    
    mappingFields.querySelectorAll('.role-select').forEach(s => 
        s.addEventListener('change', e => {
            const idx = e.target.dataset.idx;
            const isIgnored = e.target.value === 'ignore';
            state.columnMappings[idx].role = e.target.value;
            state.columnMappings[idx].include = !isIgnored;
            mappingFields.querySelector(`.include-checkbox[data-idx="${idx}"]`).checked = !isIgnored;
        })
    );
    
    processBtn.disabled = false;
}

function processData() {
    const keyColumns = state.columnMappings.filter(c => c.role === 'key' && c.include);
    const valueColumns = state.columnMappings.filter(c => c.role === 'value' && c.include);
    
    if (keyColumns.length !== 1) {
        alert('⚠️ 必須選擇一個主鍵欄位！');
        return;
    }
    
    if (valueColumns.length === 0) {
        alert('⚠️ 請至少選擇一個加總欄位！');
        return;
    }
    
    const keyCol = keyColumns[0];
    const keyName = keyCol.customName || keyCol.autoHeader;
    const rangeStr = dataRangeInput.value.trim();
    
    if (!rangeStr) {
        alert('⚠️ 請輸入資料範圍！');
        return;
    }
    
    state.allFileData = [];
    state.summaryData = new Map();
    state.orderedItemKeys = [];
    
    try {
        const range = XLSX.utils.decode_range(rangeStr);
        const dataRange = { 
            s: { r: range.s.r + parseInt(headerRowsInput.value, 10), c: range.s.c }, 
            e: range.e 
        };
        
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
        
        const seenKeys = new Set();
        state.allFileData.forEach(file => {
            file.data.forEach(row => {
                const key = row[keyName];
                if (!seenKeys.has(key)) {
                    seenKeys.add(key);
                    state.orderedItemKeys.push(key);
                }
            });
        });

        state.allFileData.forEach(file => {
            file.data.forEach(row => {
                const key = row[keyName];
                let summaryRow = state.summaryData.get(key) || { 
                    [keyName]: key, 
                    ...Object.fromEntries(valueColumns.map(c => [c.customName || c.autoHeader, 0])) 
                };
                valueColumns.forEach(c => {
                    summaryRow[c.customName || c.autoHeader] += row[c.customName || c.autoHeader] || 0;
                });
                state.summaryData.set(key, summaryRow);
            });
        });
        
        renderSummaryView(keyName, valueColumns);
        renderFileQueryView(keyName, valueColumns);
        renderItemView(keyName, valueColumns);
        
        outputArea.style.display = 'block';
        updateStep(4, 'completed');
        
        // 滾動到結果區域
        outputArea.scrollIntoView({ behavior: 'smooth' });
        
        alert(`✅ 處理完成！\n\n共處理 ${state.workbooks.length} 個檔案\n彙總了 ${state.summaryData.size} 個項目`);
        
    } catch(err) {
        alert(`❌ 處理資料錯誤：${err.message}`);
        console.error(err);
    }
}

function toNumber(val) {
    if (val == null || val === '') return 0;
    if (typeof val === 'number') return val;
    return parseFloat(String(val).replace(/,/g, '')) || 0;
}

function generateHtmlTable(data, headers, formatNumbers = false, firstRowClass = '') {
    let html = '<table class="report-table"><thead><tr>';
    html += headers.map(h => `<th>${h}</th>`).join('');
    html += '</tr></thead><tbody>';
    
    data.forEach((row, index) => {
        const rowClass = (index === 0 && firstRowClass) ? ` class="${firstRowClass}"` : '';
        html += `<tr${rowClass}>`;
        html += headers.map(h => {
            const val = row[h];
            if (formatNumbers && h !== '檔案名稱' && typeof val === 'number') {
                return `<td class="number">${val.toLocaleString('en-US', {minimumFractionDigits: 0, maximumFractionDigits: 0})}</td>`;
            }
            return `<td>${val ?? ''}</td>`;
        }).join('');
        html += '</tr>';
    });
    
    return html + '</tbody></table>';
}

function renderSummaryView(keyName, valueColumns) {
    const headers = [keyName, ...valueColumns.map(c => c.customName || c.autoHeader)];
    const data = Array.from(state.summaryData.values());
    document.getElementById('summary-view').innerHTML = 
        `<h3 style="margin-bottom:20px; color:#667eea;">📊 所有檔案加總結果 (共 ${data.length} 項)</h3>` + 
        generateHtmlTable(data, headers, true);
}

function renderFileQueryView(keyName, valueColumns) {
    fileDropdown.innerHTML = '<option value="">--- 請選擇要查詢的檔案 ---</option>' + 
        state.workbooks.map(wb => `<option value="${wb.file.name}">${wb.file.name}</option>`).join('');
    fileDetailTable.innerHTML = '';
    fileDropdown.dataset.keyName = keyName;
    fileDropdown.dataset.valueColumns = JSON.stringify(valueColumns.map(c => c.customName || c.autoHeader));
}

function renderFileDetailView() {
    const selectedFilename = fileDropdown.value;
    const container = fileDetailTable;
    
    if (!selectedFilename) {
        container.innerHTML = '';
        return;
    }
    
    const keyName = fileDropdown.dataset.keyName;
    const valueColNames = JSON.parse(fileDropdown.dataset.valueColumns);
    const headers = [keyName, ...valueColNames];
    const fileData = state.allFileData.find(file => file.fileName === selectedFilename);
    
    if (!fileData) {
        container.innerHTML = '<div class="empty-state">找不到該檔案的處理後資料</div>';
        return;
    }
    
    container.innerHTML = 
        `<h3 style="margin-bottom:20px; color:#667eea;">📁 ${selectedFilename} (共 ${fileData.data.length} 項)</h3>` + 
        generateHtmlTable(fileData.data, headers, true);
}

function renderItemView(keyName, valueColumns) {
    itemDropdown.innerHTML = '<option value="">--- 請選擇要查詢的項目 ---</option>' + 
        state.orderedItemKeys.map(item => `<option value="${item}">${item}</option>`).join('');
    document.getElementById('item-detail-table').innerHTML = '';
    itemDropdown.dataset.keyName = keyName;
    itemDropdown.dataset.valueColumns = JSON.stringify(valueColumns.map(c => c.customName || c.autoHeader));
}

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
    
    const data = state.allFileData.map(file => {
        const row = file.data.find(r => r[keyName] === selectedItem);
        const dataRow = { '檔案名稱': file.fileName };
        valueColNames.forEach(colName => {
            dataRow[colName] = row ? (row[colName] || 0) : 0;
        });
        return dataRow;
    });
    
    const summaryRow = state.summaryData.get(selectedItem);
    if (summaryRow) {
        const totalRow = { '檔案名稱': '<strong>📊 合計</strong>' };
        valueColNames.forEach(colName => {
            totalRow[colName] = summaryRow[colName] || 0;
        });
        data.unshift(totalRow);
    }
    
    container.innerHTML = 
        `<h3 style="margin-bottom:20px; color:#667eea;">🔍 項目查詢：${selectedItem}</h3>` + 
        generateHtmlTable(data, headers, true, 'total-row');
}

function setupTabs() {
    document.querySelector('.view-tabs').addEventListener('click', e => {
        if (e.target.classList.contains('tab-btn')) {
            const targetView = e.target.dataset.view;
            document.querySelectorAll('.tab-btn').forEach(btn => {
                btn.classList.toggle('active', btn.dataset.view === targetView);
            });
            document.querySelectorAll('.view-pane').forEach(pane => {
                pane.classList.toggle('active', pane.id === targetView);
            });
        }
    });
}

setupTabs();
</script>
</body>
</html>
