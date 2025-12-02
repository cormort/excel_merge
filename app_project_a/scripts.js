/**
 * 基金資料彙總報告產生器 (Project A) - 核心邏輯腳本
 *包含了：多維度矩陣運算、檔案快取、進度條、檢誤版型排序、復原/重做功能
 */

// --- 1. 全域變數與設定 ---

// 基金資料庫 (可選，用於對照)
let fundDetailsMap = {
    '行政院國家發展基金': { '業別': '投融資', '主管別': '行政院' },
};

// 版型設定
let TEMPLATE_CONFIG = {
    'custom': { 
        name: '🛠️ 自訂/通用模式 (Project A)', 
        range: '', 
        headerRows: 1 
    },
    'op_income': { 
        name: '作業基金 - 收支餘絀表', 
        range: 'A4:I38', 
        headerRows: 2, 
        sortType: 'op_income', 
        nameCell: 'A1' 
    },
    'special_cash': { 
        name: '特別收入基金 - 現金流量表', 
        range: 'A4:E48', 
        headerRows: 2, 
        sortType: 'special_cash', 
        nameCell: 'A1' 
    },
    'op_cash': { 
        name: '作業基金 - 現金流量表', 
        range: 'A4:E49', 
        headerRows: 2, 
        sortType: 'op_cash', 
        nameCell: 'A1' 
    },
    'op_surplus': { 
        name: '作業基金 - 餘絀撥補表', 
        range: 'A4:G29', 
        headerRows: 2, 
        sortType: 'op_surplus', 
        nameCell: 'A1' 
    }
};

// 排序清單 (用於檢誤版型)
let ORDER_LISTS = {
    'op_income': ["業務收入","勞務收入","銷貨收入","教學收入","租金及權利金收入","投融資業務收入","醫療收入","徵收及依法分配收入","保險收入","規費收入","其他業務收入","業務成本與費用","勞務成本","銷貨成本","教學成本","出租資產成本","投融資業務成本","醫療成本","保險成本","其他業務成本","業務費用","管理及總務費用","研究發展及訓練費用","其他業務費用","業務賸餘(短絀)","業務外收入","財務收入","其他業務外收入","業務外費用","財務費用","其他業務外費用","業務外賸餘(短絀)","本期賸餘(短絀)"],
    'special_cash': ["本期賸餘","折舊","攤銷","出售資產利益","應收帳款","存貨","預付款項","應付帳款","預收款項","應計退休金負債","其他","業務活動之淨現金流入","減少（增加）短期投資","出售長期投資","出售資產","存出保證金","投資活動之淨現金流入","增加（減少）短期債務","長期債務舉借","長期債務償還","基金（資本）之撥入","基金（資本）之撥出","融資活動之淨現金流入","現金及約當現金之淨增（減）數","期初現金及約當現金餘額","期末現金及約當現金餘額"],
    'op_surplus': ["賸餘之部","本期賸餘","前期未分配賸餘","追溯適用及追溯重編之影響數","公積轉列數","其他轉入數","分配之部","填補累積短絀","提存公積","賸餘撥充基金數","解繳公庫淨額","其他依法分配數","未分配賸餘","短絀之部","本期短絀","前期待填補之短絀","追溯適用及追溯重編之影響數","其他轉入數","填補之部","撥用賸餘","撥用公積","折減基金","公庫撥款","待填補之短絀"]
};

// 應用程式狀態
const state = { 
    workbooks: [], 
    columnMappings: [], 
    matrixData: new Map(), // 核心資料結構
    allFileNames: [], 
    allValueCols: [], 
    keyColName: '', 
    sortedKeys: [],
    
    originalData: null, 
    isTransposed: false, 
    transposeKeyIndex: null, 
    currentTemplate: 'custom',
    
    fileCache: new Map(), // 快取
    historyStack: [],     // 撤銷堆疊
    redoStack: [],        // 重做堆疊
    maxHistory: 20
};

// DOM 元素參考
const els = {
    dropArea: document.getElementById('drop-area'),
    fileInput: document.getElementById('file-input'),
    fileListContainer: document.getElementById('file-list-container'),
    previewArea: document.getElementById('preview-area'),
    mappingFields: document.getElementById('mapping-fields'),
    processBtn: document.getElementById('process-btn'),
    outputArea: document.getElementById('output-area'),
    
    dataRangeInput: document.getElementById('data-range-input'),
    headerRowsInput: document.getElementById('header-rows-input'),
    loadHeadersBtn: document.getElementById('load-headers-btn'),
    templateSelect: document.getElementById('template-select'),
    autoDetectBtn: document.getElementById('auto-detect-btn'),
    
    transposeBtn: document.getElementById('transpose-btn'),
    transposeKeySelect: document.getElementById('transpose-key-select'),
    transposeControls: document.getElementById('transpose-controls'),
    clearBtn: document.getElementById('clear-btn'),
    
    viewTabs: document.querySelector('.view-tabs'),
    fileDropdown: document.getElementById('file-dropdown'),
    itemDropdown: document.getElementById('item-dropdown'),
    fileDetailTable: document.getElementById('file-detail-table'),
    itemDetailTable: document.getElementById('item-detail-table'),
    
    progressContainer: document.getElementById('progress-container'),
    progressBar: document.getElementById('progress-bar'),
    progressText: document.getElementById('progress-text'),
    progressPercent: document.getElementById('progress-percent'),
    
    undoBtn: document.getElementById('undo-btn'),
    redoBtn: document.getElementById('redo-btn'),
    
    sourceNameMode: document.getElementById('source-name-mode'),
    sourceNameCell: document.getElementById('source-name-cell'),
    sourceCellGroup: document.getElementById('source-cell-group'),
    
    matrixValueSelect: null 
};

// --- 2. 初始化與事件監聽 ---

function init() {
    populateTemplateDropdown();
    setupEventListeners();
    updateStep(1);
    updateHistoryButtons();
}

function populateTemplateDropdown() {
    const select = els.templateSelect;
    if (!select) return;
    select.innerHTML = '';
    for (const [key, config] of Object.entries(TEMPLATE_CONFIG)) {
        const opt = document.createElement('option');
        opt.value = key;
        opt.textContent = config.name;
        select.appendChild(opt);
    }
}

function setupEventListeners() {
    // 檔案拖放與選擇
    els.dropArea.addEventListener('click', () => els.fileInput.click());
    ['dragenter', 'dragover'].forEach(e => els.dropArea.addEventListener(e, evt => { evt.preventDefault(); els.dropArea.classList.add('drag-over'); }));
    ['dragleave', 'drop'].forEach(e => els.dropArea.addEventListener(e, evt => { evt.preventDefault(); els.dropArea.classList.remove('drag-over'); }));
    els.dropArea.addEventListener('drop', e => { if(e.dataTransfer.files.length) handleFiles(e.dataTransfer.files); });
    els.fileInput.addEventListener('change', e => { if(e.target.files.length) handleFiles(e.target.files); });
    
    els.clearBtn.addEventListener('click', (e) => {
        e.stopPropagation();
        if(confirm('確定要清除所有已上傳的檔案嗎？(快取將保留)')) resetUI();
    });

    // 設定與操作
    els.templateSelect.addEventListener('change', handleTemplateChange);
    els.autoDetectBtn.addEventListener('click', autoDetectBestRange);
    els.loadHeadersBtn.addEventListener('click', () => { 
        saveState("讀取欄位"); 
        loadHeadersAndMapping(); 
    });
    els.processBtn.addEventListener('click', () => { 
        saveState("執行彙總"); 
        processProjectAData(); 
    });
    
    // 轉置與匯出
    els.transposeBtn.addEventListener('click', transposeData);
    els.transposeKeySelect.addEventListener('change', applyTranspose);
    document.getElementById('export-csv-btn').addEventListener('click', () => exportMatrix('csv'));
    document.getElementById('export-xlsx-btn').addEventListener('click', () => exportMatrix('xlsx'));
    document.getElementById('export-html-btn').addEventListener('click', () => exportMatrix('html'));
    document.getElementById('export-json-btn').addEventListener('click', () => exportMatrix('json'));
    
    // 分頁切換
    els.viewTabs.addEventListener('click', e => {
        if (e.target.classList.contains('tab-btn')) {
            const targetId = e.target.dataset.view;
            document.querySelectorAll('.tab-btn').forEach(b => b.classList.toggle('active', b === e.target));
            document.querySelectorAll('.view-pane').forEach(p => p.classList.toggle('active', p.id === targetId));
        }
    });

    // 詳情查詢
    els.fileDropdown.addEventListener('change', renderFileDetailView);
    els.itemDropdown.addEventListener('change', renderItemDetailView);
    
    // 歷史記錄與名稱模式
    els.undoBtn.addEventListener('click', undo);
    els.redoBtn.addEventListener('click', redo);
    els.sourceNameMode.addEventListener('change', (e) => {
        els.sourceCellGroup.style.display = e.target.value === 'cell' ? 'block' : 'none';
    });
}

// --- 3. 檔案處理與介面更新 ---

function showProgress(msg) {
    els.progressContainer.style.display = 'block';
    updateProgress(0, msg);
}

function updateProgress(percent, msg) {
    const p = Math.round(percent);
    els.progressBar.style.width = `${p}%`;
    els.progressPercent.textContent = `${p}%`;
    if(msg) els.progressText.textContent = msg;
}

function hideProgress() {
    setTimeout(() => {
        els.progressContainer.style.display = 'none';
        updateProgress(0, '');
    }, 500);
}

function updateStep(stepNum, status = 'active') {
    document.querySelectorAll('.step').forEach((step, i) => {
        step.classList.remove('active', 'completed');
        if (i + 1 < stepNum) step.classList.add('completed');
        if (i + 1 === stepNum) step.classList.add(status);
    });
}

function resetUI() {
    state.workbooks = [];
    state.matrixData.clear();
    state.columnMappings = [];
    state.originalData = null;
    
    els.fileInput.value = '';
    els.fileListContainer.innerHTML = '';
    
    document.getElementById('section-preview').style.display = 'none';
    document.getElementById('section-range').style.display = 'none';
    document.getElementById('section-mapping').style.display = 'none';
    els.outputArea.style.display = 'none';
    els.clearBtn.style.display = 'none';
    
    updateStep(1);
}

// 處理檔案上傳
async function handleFiles(fileList) {
    showProgress('讀取檔案中...');
    const files = Array.from(fileList);
    const total = files.length;
    let loadedCount = 0;
    
    const successItems = [];
    const failedItems = [];
    
    state.workbooks = []; 

    for (let i = 0; i < total; i++) {
        const file = files[i];
        try {
            if (state.fileCache.has(file.name)) {
                state.workbooks.push({ 
                    file, 
                    workbook: state.fileCache.get(file.name).workbook, 
                    fromCache: true 
                });
                successItems.push({ name: file.name, cached: true });
            } else {
                const workbook = await readFileAsync(file);
                state.fileCache.set(file.name, { workbook, timestamp: Date.now() });
                state.workbooks.push({ file, workbook, fromCache: false });
                successItems.push({ name: file.name, cached: false });
            }
        } catch (err) {
            console.error(err);
            failedItems.push({ name: file.name, error: err.message });
        }
        
        loadedCount++;
        updateProgress((loadedCount / total) * 100, `讀取中 (${loadedCount}/${total})`);
        await new Promise(r => setTimeout(r, 0));
    }
    
    hideProgress();
    
    // 生成清單 HTML (預設收合)
    let listHtml = `<details class="file-list-details">
        <summary class="file-list-summary">
            <span>📂 匯入結果：成功 ${successItems.length} / 失敗 ${failedItems.length} (點擊展開)</span>
        </summary>
        <div class="file-list">`;
    
    if (successItems.length > 0) {
        listHtml += `<div style="padding:5px 10px; background:#f0f9eb; color:#28a745; font-weight:bold; font-size:0.9em;">✅ 成功列表</div>`;
        listHtml += successItems.map(item => 
            `<div class="file-item"><span>📄 ${item.name} ${item.cached ? '<small style="color:green">(快取)</small>' : ''}</span></div>`
        ).join('');
    }
    
    if (failedItems.length > 0) {
        listHtml += `<div style="padding:5px 10px; background:#fef0f0; color:#dc3545; font-weight:bold; font-size:0.9em; margin-top:10px;">❌ 失敗列表</div>`;
        listHtml += failedItems.map(item => 
            `<div class="file-item" style="color:#dc3545;"><span>⚠️ ${item.name}</span><small>${item.error || '讀取錯誤'}</small></div>`
        ).join('');
    }
    
    listHtml += `</div></details>`;
    els.fileListContainer.innerHTML = listHtml;
    
    els.clearBtn.style.display = 'inline-flex';
    if(state.workbooks.length > 0) {
        generatePreview(state.workbooks[0].workbook.Sheets[state.workbooks[0].workbook.SheetNames[0]]);
        document.getElementById('section-preview').style.display = 'block';
        updateStep(2);
        if (els.templateSelect.value !== 'custom') handleTemplateChange();
    }
}

function readFileAsync(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = e => resolve(XLSX.read(e.target.result, {type: 'array'}));
        reader.onerror = reject;
        reader.readAsArrayBuffer(file);
    });
}

function generatePreview(sheet) {
    const range = XLSX.utils.decode_range(sheet['!ref'] || 'A1:A1');
    range.e.r = Math.min(range.e.r, range.s.r + 20); // 只預覽前20列
    els.previewArea.innerHTML = XLSX.utils.sheet_to_html(sheet, { range: range, id: 'preview-table', editable: false });
}

// 處理合併儲存格填值
function unmergeAndFill(data, sheet, range) {
    const merges = sheet['!merges'] || [];
    merges.forEach(merge => {
        if (merge.s.c > range.e.c || merge.e.c < range.s.c || merge.s.r > range.e.r || merge.e.r < range.s.r) return;
        
        const startRow = Math.max(0, merge.s.r - range.s.r);
        const startCol = Math.max(0, merge.s.c - range.s.c);
        const endRow = Math.min(data.length - 1, merge.e.r - range.s.r);
        const endCol = Math.min((data[0]?.length || 1) - 1, merge.e.c - range.s.c);
        
        const val = data[startRow] ? data[startRow][startCol] : null;
        
        for (let r = startRow; r <= endRow; r++) {
            if (!data[r]) data[r] = [];
            for (let c = startCol; c <= endCol; c++) {
                if (val != null) data[r][c] = val;
            }
        }
    });
    return data;
}

// --- 4. 歷史記錄管理 (Undo/Redo) ---

function saveState(actionName) {
    const snapshot = {
        columnMappings: JSON.parse(JSON.stringify(state.columnMappings)),
        matrixData: Array.from(state.matrixData.entries()).map(([k, v]) => [k, Array.from(v.entries())]),
        allFileNames: [...state.allFileNames],
        keyColName: state.keyColName,
        allValueCols: [...state.allValueCols],
        action: actionName
    };
    
    state.historyStack.push(snapshot);
    if (state.historyStack.length > state.maxHistory) state.historyStack.shift();
    state.redoStack = []; // 清空重做堆疊
    updateHistoryButtons();
}

function undo() {
    if (state.historyStack.length === 0) return;
    
    // 儲存當前狀態到 Redo
    state.redoStack.push({
         columnMappings: JSON.parse(JSON.stringify(state.columnMappings)),
         matrixData: Array.from(state.matrixData.entries()).map(([k, v]) => [k, Array.from(v.entries())]),
         allFileNames: [...state.allFileNames],
         keyColName: state.keyColName,
         allValueCols: [...state.allValueCols]
    });
    
    restoreState(state.historyStack.pop());
    updateHistoryButtons();
}

function redo() {
    if (state.redoStack.length === 0) return;
    
    const snapshot = state.redoStack.pop();
    saveState("redo"); // 這裡 saveState 會 push 到 history
    // 但因為 redo 動作本身就是從 redoStack 移回 history，上面的 saveState 邏輯可能會造成兩次 push
    // 簡單處理：pop 掉剛才 saveState 產生的一筆，再 push 正確的 snapshot
    state.historyStack.pop(); 
    state.historyStack.push(snapshot);
    
    restoreState(snapshot);
    updateHistoryButtons();
}

function restoreState(snapshot) {
    state.columnMappings = snapshot.columnMappings;
    state.matrixData = new Map(snapshot.matrixData.map(([k, v]) => [k, new Map(v)]));
    state.allFileNames = snapshot.allFileNames;
    state.keyColName = snapshot.keyColName;
    state.allValueCols = snapshot.allValueCols;
    
    if (state.matrixData.size > 0) {
        renderMatrixView();
        updateDetailDropdowns();
    }
    // 若在欄位設定頁面，刷新 Mapping 表
    if (document.getElementById('section-mapping').style.display !== 'none') {
        renderMappingTableDOM();
    }
    updateHistoryButtons();
}

function updateHistoryButtons() {
    els.undoBtn.disabled = state.historyStack.length === 0;
    els.redoBtn.disabled = state.redoStack.length === 0;
}

// --- 5. 欄位讀取與範圍設定 ---

function handleTemplateChange() {
    const key = els.templateSelect.value;
    state.currentTemplate = key;
    if (key !== 'custom') {
        const conf = TEMPLATE_CONFIG[key];
        els.dataRangeInput.value = conf.range;
        els.headerRowsInput.value = conf.headerRows;
        
        // 自動帶入預設名稱儲存格
        if (conf.nameCell) {
            els.sourceNameMode.value = 'cell';
            els.sourceNameCell.value = conf.nameCell;
            els.sourceCellGroup.style.display = 'block';
        }
        
        document.getElementById('section-range').style.display = 'block';
        updateStep(3);
    }
}

function autoDetectBestRange() {
    const sheet = state.workbooks[0].workbook.Sheets[state.workbooks[0].workbook.SheetNames[0]];
    if(sheet['!ref']) {
        els.dataRangeInput.value = sheet['!ref'];
        els.headerRowsInput.value = 1;
        document.getElementById('section-range').style.display = 'block';
        updateStep(3);
    } else {
        alert('無法自動偵測範圍');
    }
}

function loadHeadersAndMapping() {
    try {
        const rangeStr = els.dataRangeInput.value.trim();
        const headerRows = parseInt(els.headerRowsInput.value);
        
        if (!rangeStr) return alert('請輸入資料範圍');

        const sheet = state.workbooks[0].workbook.Sheets[state.workbooks[0].workbook.SheetNames[0]];
        const range = XLSX.utils.decode_range(rangeStr);
        
        // 讀取並處理標頭
        const headerRange = { s: range.s, e: { r: range.s.r + headerRows - 1, c: range.e.c } };
        let headerData = XLSX.utils.sheet_to_json(sheet, { header: 1, range: headerRange, defval: null });
        headerData = unmergeAndFill(headerData, sheet, headerRange);

        const headers = [];
        const usedNames = new Set();
        
        // 組合多列標頭名稱
        for(let c = 0; c <= range.e.c - range.s.c; c++) {
            let parts = [];
            for(let r = 0; r < headerRows; r++) {
                if(headerData[r] && headerData[r][c]) {
                    parts.push(String(headerData[r][c]).trim());
                }
            }
            
            let baseName = parts.filter((v,i,a)=>a.indexOf(v)===i).join('_') || `欄位${c+1}`;
            let finalName = baseName;
            let counter = 2;
            
            // 處理重複名稱
            while (usedNames.has(finalName)) {
                finalName = `${baseName}_${counter++}`;
            }
            usedNames.add(finalName);
            headers.push(finalName);
        }

        // 預讀取數據區以備轉置 (暫不實作複雜轉置，僅保留結構)
        const dataRange = { s: { r: range.s.r + headerRows, c: range.s.c }, e: range.e };
        let bodyData = XLSX.utils.sheet_to_json(sheet, { header: 1, range: dataRange, defval: null });
        bodyData = unmergeAndFill(bodyData, sheet, dataRange);

        state.originalData = { headers, bodyData, range, headerRows };
        state.isTransposed = false; 
        els.transposeControls.style.display = 'none';
        els.transposeBtn.textContent = '🔄 欄列轉置';

        generateMappingTable(headers, range.s.c);
        
    } catch(e) { 
        alert('讀取失敗: ' + e.message); 
    }
}

function generateMappingTable(headers, startColIdx) {
    state.columnMappings = headers.map((h, i) => ({
        excelCol: XLSX.utils.encode_col(startColIdx + i),
        name: h,
        role: i === 0 ? 'key' : 'value', // 預設第一欄為主鍵
        include: true
    }));
    renderMappingTableDOM();
    document.getElementById('section-mapping').style.display = 'block';
    updateStep(3, 'completed');
}

function renderMappingTableDOM() {
    let html = `<table class="mapping-table"><thead><tr><th>Excel</th><th>欄位名稱 (X軸)</th><th>角色</th><th>納入</th></tr></thead><tbody>`;
    state.columnMappings.forEach((col, i) => {
        html += `<tr>
            <td>${col.excelCol}</td>
            <td><input type="text" value="${col.name}" onchange="updateMapName(${i},this.value)" style="width:100%"></td>
            <td>
                <select onchange="updateMapRole(${i},this.value)">
                    <option value="key" ${col.role==='key'?'selected':''}>🔑 主鍵 (Y軸)</option>
                    <option value="value" ${col.role==='value'?'selected':''}>📊 數值 (X軸)</option>
                    <option value="ignore" ${col.role==='ignore'?'selected':''}>🚫 忽略</option>
                </select>
            </td>
            <td><input type="checkbox" ${col.include?'checked':''} onchange="updateMapInclude(${i},this.checked)"></td>
        </tr>`;
    });
    els.mappingFields.innerHTML = html + '</tbody></table>';
}

// 綁定到 window 以供 HTML 中的 onchange 呼叫
window.updateMapName = (i, v) => state.columnMappings[i].name = v;
window.updateMapRole = (i, v) => { 
    state.columnMappings[i].role = v; 
    state.columnMappings[i].include = (v !== 'ignore'); 
    renderMappingTableDOM(); 
};
window.updateMapInclude = (i, v) => state.columnMappings[i].include = v;

function transposeData() {
    alert('Project A 矩陣模式建議直接使用標準檢視。若需轉置請手動調整 Excel。');
}
function applyTranspose() {} 


// --- 6. 核心處理邏輯 (Project A) ---

async function processProjectAData() {
    try {
        const keyCol = state.columnMappings.find(c => c.role === 'key');
        if (!keyCol) return alert('請設定一個主鍵欄位');

        const range = XLSX.utils.decode_range(els.dataRangeInput.value);
        const headerRows = parseInt(els.headerRowsInput.value);
        const startRow = range.s.r + headerRows;

        // 檢查名稱來源設定
        const nameMode = els.sourceNameMode.value;
        const nameCellAddr = els.sourceNameCell.value.trim().toUpperCase();
        if (nameMode === 'cell' && !nameCellAddr) return alert('請輸入名稱來源的儲存格座標 (如 A1)');

        showProgress("正在彙總資料...");
        
        state.keyColName = keyCol.name;
        state.allValueCols = state.columnMappings.filter(c => c.role === 'value' && c.include).map(c => c.name);
        state.allFileNames = [];
        state.matrixData.clear();
        
        const totalFiles = state.workbooks.length;

        for(let i=0; i<totalFiles; i++) {
            const wb = state.workbooks[i];
            const sheet = wb.workbook.Sheets[wb.workbook.SheetNames[0]]; 
            
            // --- 基金名稱處理 ---
            let fundName = wb.file.name.replace(/\.(xlsx|xls)$/i, ''); 
            
            if (nameMode === 'cell') {
                const cell = sheet[nameCellAddr];
                if (cell && cell.v) {
                    fundName = String(cell.v).trim().replace(/\s+/g, '');
                }
            }
            
            // 截斷「基金」之後的文字 (移除表名)
            const idx = fundName.lastIndexOf('基金');
            if (idx > -1) {
                fundName = fundName.substring(0, idx + 2); // 保留 "基金" 兩字
            }
            
            // 防止名稱重複
            let uniqueName = fundName;
            let counter = 2;
            while (state.allFileNames.includes(uniqueName)) {
                uniqueName = `${fundName}_${counter++}`;
            }
            state.allFileNames.push(uniqueName);
            
            // --- 數據讀取與過濾 ---
            const rawData = XLSX.utils.sheet_to_json(sheet, {header:1, range: {s:{r:startRow, c:range.s.c}, e:range.e}, defval:null});
            
            rawData.forEach(row => {
                const keyMapIdx = state.columnMappings.findIndex(c => c.role === 'key');
                const relKeyIdx = XLSX.utils.decode_col(state.columnMappings[keyMapIdx].excelCol) - range.s.c;
                const keyVal = row[relKeyIdx];
                if (!keyVal) return;
                const keyStr = String(keyVal).trim();

                if (!state.matrixData.has(keyStr)) state.matrixData.set(keyStr, new Map());
                const fileMap = state.matrixData.get(keyStr);
                
                const rowData = {};
                state.columnMappings.forEach(map => {
                    if (map.role === 'value' && map.include) {
                        const cIdx = XLSX.utils.decode_col(map.excelCol) - range.s.c;
                        let val = row[cIdx];
                        
                        // 嚴格數值轉換：排除文字干擾
                        if (val == null || val === '') {
                            val = 0;
                        } else if (typeof val !== 'number') {
                            // 僅保留數字、小數點與負號
                            const cleanStr = String(val).replace(/[^0-9.-]/g, '');
                            val = parseFloat(cleanStr) || 0; 
                        }
                        
                        rowData[map.name] = val;
                    }
                });
                fileMap.set(uniqueName, rowData); 
            });
            
            updateProgress((i / totalFiles) * 100, `處理中: ${uniqueName}`);
            if (i % 5 === 0) await new Promise(r => setTimeout(r, 0));
        }

        renderMatrixView();
        updateDetailDropdowns();
        updateStep(4);
        hideProgress();
        alert(`✅ 彙總完成！共 ${state.allFileNames.length} 個檔案 (名稱已清洗)。`);
        
    } catch (err) { 
        hideProgress(); 
        console.error(err); 
        alert('處理錯誤: ' + err.message); 
    }
}

// --- 7. 視圖渲染與匯出 ---

function renderMatrixView() {
    els.outputArea.style.display = 'block';
    
    // 排序邏輯
    let sortedKeys = Array.from(state.matrixData.keys());
    const tmpl = TEMPLATE_CONFIG[state.currentTemplate];
    if (tmpl && tmpl.sortType && ORDER_LISTS[tmpl.sortType]) {
        const orderMap = new Map(ORDER_LISTS[tmpl.sortType].map((k, i) => [k, i]));
        sortedKeys.sort((a, b) => {
            const idxA = orderMap.has(a) ? orderMap.get(a) : 9999;
            const idxB = orderMap.has(b) ? orderMap.get(b) : 9999;
            return idxA - idxB;
        });
    } else {
        sortedKeys.sort();
    }
    state.sortedKeys = sortedKeys;

    const html = `
        <div class="alert alert-success"><strong>矩陣視圖</strong>：${tmpl.name || '自訂模式'}</div>
        <div style="margin-bottom:15px; background:#f8f9fa; padding:10px; border-radius:5px;">
            <label>👁️ 選擇顯示數值 (X軸)：</label>
            <select id="matrix-value-select" onchange="updateMatrixTable()">
                ${state.allValueCols.map(c => `<option value="${c}">${c}</option>`).join('')}
            </select>
        </div>
        <div id="matrix-table-container" style="overflow-x:auto; max-height:600px;"></div>
    `;
    document.getElementById('summary-view').innerHTML = html;
    
    els.matrixValueSelect = document.getElementById('matrix-value-select');
    updateMatrixTable();
}

window.updateMatrixTable = function() {
    const targetCol = document.getElementById('matrix-value-select').value;
    let html = `<table class="report-table"><thead><tr>
        <th style="position:sticky;left:0;z-index:10;min-width:150px;">${state.keyColName}</th>
        ${state.allFileNames.map(f=>`<th>${f}</th>`).join('')}
        <th style="background:#444;color:#fff">合計</th>
    </tr></thead><tbody>`;
    
    state.sortedKeys.forEach(key => {
        const fileMap = state.matrixData.get(key);
        let sum = 0;
        html += `<tr><td style="position:sticky;left:0;background:#fff;font-weight:bold">${key}</td>`;
        
        state.allFileNames.forEach(f => {
            const val = fileMap.get(f) ? (fileMap.get(f)[targetCol]||0) : 0;
            sum += val;
            html += `<td class="number">${val===0?'-':val.toLocaleString()}</td>`;
        });
        
        html += `<td class="number total-col">${sum.toLocaleString()}</td></tr>`;
    });
    
    document.getElementById('matrix-table-container').innerHTML = html + '</tbody></table>';
};

function updateDetailDropdowns() {
    els.fileDropdown.innerHTML = '<option value="">-- 請選擇 --</option>' + state.allFileNames.map(f => `<option value="${f}">${f}</option>`).join('');
    els.itemDropdown.innerHTML = '<option value="">-- 請選擇 --</option>' + state.sortedKeys.map(k => `<option value="${k}">${k}</option>`).join('');
}

function renderFileDetailView() {
    const fname = els.fileDropdown.value;
    if(!fname) return;
    
    let html = `<h3>${fname}</h3><table class="report-table"><thead><tr>
        <th>${state.keyColName}</th>
        ${state.allValueCols.map(c=>`<th>${c}</th>`).join('')}
    </tr></thead><tbody>`;
    
    state.sortedKeys.forEach(key => {
        const fileMap = state.matrixData.get(key);
        if(fileMap.has(fname)) {
            const d = fileMap.get(fname);
            html += `<tr><td>${key}</td>${state.allValueCols.map(c=>`<td class="number">${(d[c]||0).toLocaleString()}</td>`).join('')}</tr>`;
        }
    });
    els.fileDetailTable.innerHTML = html + '</tbody></table>';
}

function renderItemDetailView() {
    const key = els.itemDropdown.value;
    if(!key) return;
    
    const fileMap = state.matrixData.get(key);
    let html = `<h3>${key}</h3><table class="report-table"><thead><tr>
        <th>檔案</th>
        ${state.allValueCols.map(c=>`<th>${c}</th>`).join('')}
    </tr></thead><tbody>`;
    
    state.allFileNames.forEach(f => {
        const d = fileMap.get(f);
        html += `<tr><td>${f}</td>${state.allValueCols.map(c=>`<td class="number">${d?(d[c]||0).toLocaleString():'-'}</td>`).join('')}</tr>`;
    });
    els.itemDetailTable.innerHTML = html + '</tbody></table>';
}

function exportMatrix(type) {
    if(!state.matrixData.size) return alert('無資料');
    const targetCol = els.matrixValueSelect ? els.matrixValueSelect.value : state.allValueCols[0];
    
    const data = state.sortedKeys.map(key => {
        const row = { [state.keyColName]: key };
        const fileMap = state.matrixData.get(key);
        let sum = 0;
        state.allFileNames.forEach(f => {
            const val = fileMap.get(f)?.[targetCol]||0; 
            row[f] = val; 
            sum += val;
        });
        row['總計'] = sum; 
        return row;
    });
    
    if(type==='csv') {
        const wb = XLSX.utils.book_new(); 
        XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(data), "Sheet1");
        XLSX.writeFile(wb, 'report.csv');
    } else if(type==='xlsx') {
        const wb = XLSX.utils.book_new(); 
        XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(data), "Sheet1");
        XLSX.writeFile(wb, 'report.xlsx');
    } else if(type==='json') {
        const a=document.createElement('a'); 
        a.href=URL.createObjectURL(new Blob([JSON.stringify(data,null,2)],{type:'application/json'})); 
        a.download='report.json'; 
        a.click();
    } else if (type === 'html') {
        const tbl = document.getElementById('matrix-table-container').innerHTML;
        const blob = new Blob([`<html><head><meta charset="utf-8"><style>table{border-collapse:collapse;width:100%}td,th{border:1px solid #999;padding:4px}</style></head><body>${tbl}</body></html>`], { type: 'text/html' });
        const a = document.createElement('a'); 
        a.href = URL.createObjectURL(blob); 
        a.download = 'report.html'; 
        a.click();
    }
}

// 啟動程式
init();