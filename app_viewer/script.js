/**
 * ExcelViewer — 第一階段優化版 (完整無刪減)
 * 實作項目：
 * [A1] 清洗設定快取與免重傳重新套用
 * [A3] 合併視圖「最小化」保留狀態
 * [A4] 刪除操作支援 Undo (復原) Stack
 * [B3] 智慧去重面板與視覺高亮標示
 */

const ExcelViewer = (() => {
  'use strict';

  // ─────────────────────────────────────────────
  // 1. 常數與初始狀態
  // ─────────────────────────────────────────────

  const VALID_EXTENSIONS = ['.xls', '.xlsx'];

  const state = {
    originalHtmlString: '',
    isProcessing: false,
    loadedFiles:[],
    loadedTables: 0,
    zoomedCard: null,

    // [A1] 儲存解析後的原始 JSON，用於「重新套用清洗設定」而不必重新讀檔
    rawSheetsCache:[],
    isSettingsDirty: false,

    // [A4] Undo 復原堆疊
    undoStack: [],

    //[A3] 合併視圖狀態 (現在關閉時不會被清空)
    isMergedView: false,
    isEditing: false,
    showTotalRow: false,
    showSourceColumn: false,
    mergedData: [],
    mergedHeaders: [],

    fundSortOrder:[],
    fundAliasMap: {},
    fundAliasKeys:[],
  };

  // ─────────────────────────────────────────────
  // 2. 工具函數 & Undo 機制 [A4]
  // ─────────────────────────────────────────────

  const utils = {
    debounce(fn, ms) {
      let timer;
      return (...args) => { clearTimeout(timer); timer = setTimeout(() => fn(...args), ms); };
    },
    formatNumber(value) {
      const num = parseFloat(String(value).replace(/,/g, ''));
      return isNaN(num) ? value : new Intl.NumberFormat('en-US').format(num);
    },
    isStrictNumber(value) {
      const clean = String(value).replace(/,/g, '').trim();
      return clean !== '' && !isNaN(clean);
    },
    parsePositionString(str) {
      const indices = new Set();
      str.split(',').map(p => p.trim()).filter(Boolean).forEach(part => {
        if (part.includes('-')) {
          const [s, e] = part.split('-').map(Number);
          if (!isNaN(s) && !isNaN(e) && s <= e) {
            for (let i = s; i <= e; i++) indices.add(i - 1);
          }
        } else {
          const n = Number(part);
          if (!isNaN(n)) indices.add(n - 1);
        }
      });
      return Array.from(indices).sort((a, b) => a - b);
    },
    readFileAsBinary(file) {
      return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = e => resolve(e.target.result);
        reader.onerror = reject;
        reader.readAsBinaryString(file);
      });
    },
    calculateColumnWidths(data) {
      if (!data.length) return[];
      return data[0].map((_, col) => ({
        wch: Math.min(50, Math.max(10, ...data.map(r => (r[col] ? String(r[col]).length : 0))) + 2),
      }));
    }
  };

  /**[A4] Undo 管理器 */
  const undoManager = {
    push(description, restoreFn) {
      state.undoStack.push({ description, restoreFn });
      if (state.undoStack.length > 5) state.undoStack.shift(); // 保留最近 5 步
      this.showToast(description);
    },
    undoLast() {
      if (state.undoStack.length === 0) return;
      const action = state.undoStack.pop();
      action.restoreFn();
      this.hideToast();
      console.log(`已復原: ${action.description}`);
    },
    showToast(desc) {
      const toast = el.get('undoToast');
      if (toast) {
        el.get('undoText').textContent = `已${desc}`;
        toast.classList.add('show');
        clearTimeout(this.timer);
        this.timer = setTimeout(() => this.hideToast(), 8000); // 8 秒後消失
      }
    },
    hideToast() {
      const toast = el.get('undoToast');
      if (toast) toast.classList.remove('show');
    }
  };

  // ─────────────────────────────────────────────
  // 3. DOM 快取
  // ─────────────────────────────────────────────

  const el = (() => {
    const ids = {
      fileInput: 'file-input', dropArea: 'drop-area', dropAreaInitial: 'drop-area-initial',
      dropAreaLoaded: 'drop-area-loaded', fileCount: 'file-count', fileNames: 'file-names',
      clearFilesBtn: 'clear-files-btn', importOptionsContainer: 'import-options-container',
      specificSheetNameGroup: 'specific-sheet-name-group', specificSheetNameInput: 'specific-sheet-name-input',
      specificSheetPositionGroup: 'specific-sheet-position-group', specificSheetPositionInput: 'specific-sheet-position-input',
      
      // 清洗設定相關
      skipTopRowsCheckbox: 'skip-top-rows-checkbox', skipTopRowsInput: 'skip-top-rows-input',
      removeEmptyRowsCheckbox: 'remove-empty-rows-checkbox',
      removeKeywordRowsCheckbox: 'remove-keyword-rows-checkbox', removeKeywordRowsInput: 'remove-keyword-rows-input',
      reapplyBanner: 'reapply-banner', reapplySettingsBtn: 'reapply-settings-btn',

      displayArea: 'excel-display-area', controlPanel: 'control-panel', loadStatusMessage: 'load-status-message',
      listViewBtn: 'list-view-btn', gridViewBtn: 'grid-view-btn', gridScaleControl: 'grid-scale-control', gridScaleSlider: 'grid-scale-slider',
      tableLevelControls: 'table-level-controls', selectAllTablesBtn: 'select-all-tables-btn',
      unselectAllTablesBtn: 'unselect-all-tables-btn', deleteSelectedTablesBtn: 'delete-selected-tables-btn',
      sortByNameBtn: 'sort-by-fund-name-btn', selectedTablesInfo: 'selected-tables-info', selectedTablesList: 'selected-tables-list',
      searchInput: 'search-input', selectAllBtn: 'select-all-btn', invertSelectionBtn: 'invert-selection-btn',
      deleteSelectedBtn: 'delete-selected-btn', selectEmptyBtn: 'select-empty-btn', selectByKeywordGroup: 'select-by-keyword-group',
      selectKeywordInput: 'select-keyword-input', selectByKeywordBtn: 'select-by-keyword-btn', selectKeywordRegex: 'select-keyword-regex',
      resetViewBtn: 'reset-view-btn', showHiddenBtn: 'show-hidden-btn', exportMergedXlsxBtn: 'export-merged-xlsx-btn',
      mergeViewBtn: 'merge-view-btn', viewCheckedCombinedBtn: 'view-checked-combined-btn', backToTopBtn: 'back-to-top-btn',
      mergeViewModal: 'merge-view-modal', closeMergeViewBtn: 'close-merge-view-btn', mergeViewContent: 'merge-view-content',
      toggleToolbarBtn: 'toggle-toolbar-btn', collapsibleToolbar: 'collapsible-toolbar-area',
      searchInputMerged: 'search-input-merged', selectKeywordInputMerged: 'select-keyword-input-merged',
      selectKeywordRegexMerged: 'select-keyword-regex-merged', executeFilterSelectionBtn: 'execute-filter-selection-btn',
      unselectMergedRowsBtn: 'unselect-merged-rows-btn', invertSelectionMergedBtn: 'invert-selection-merged-btn',
      colSelect1: 'col-select-1', colSelect2: 'col-select-2', inputCriteria1: 'input-criteria-1', inputCriteria2: 'input-criteria-2',
      editDataBtn: 'edit-data-btn', saveEditsBtn: 'save-edits-btn', cancelEditsBtn: 'cancel-edits-btn', addNewRowBtn: 'add-new-row-btn',
      copySelectedRowsBtn: 'copy-selected-rows-btn', deleteMergedRowsBtn: 'delete-merged-rows-btn', toggleTotalRowBtn: 'toggle-total-row-btn',
      toggleSourceColBtn: 'toggle-source-col-btn', exportCurrentMergedXlsxBtn: 'export-current-merged-xlsx-btn',
      sortMergedByNameBtn: 'sort-merged-by-fund-name-btn', columnOperationsBtn: 'column-operations-btn', columnModal: 'column-modal',
      closeColumnModalBtn: 'close-column-modal-btn', columnChecklist: 'column-checklist', applyColumnChangesBtn: 'apply-column-changes-btn',
      modalCheckAll: 'modal-check-all', modalUncheckAll: 'modal-uncheck-all', smartDedupBtn: 'smart-dedup-btn',
      dedupModal: 'dedup-modal', closeDedupModalBtn: 'close-dedup-modal-btn', cancelDedupBtn: 'cancel-dedup-btn', executeDedupBtn: 'execute-dedup-btn',
      dedupColSelect: 'dedup-col-select',
      
      // 去重面板與 Undo
      dedupResultPanel: 'dedup-result-panel', dedupResultText: 'dedup-result-text', 
      clearDedupMarksBtn: 'clear-dedup-marks-btn', deleteDedupMarksBtn: 'delete-dedup-marks-btn',
      undoToast: 'undo-toast', undoText: 'undo-text', undoBtn: 'undo-btn'
    };

    const cache = {};
    return {
      init() { Object.keys(ids).forEach(k => { cache[k] = document.getElementById(ids[k]); }); },
      get(key) { return cache[key]; }
    };
  })();

  const elements = new Proxy(el, { get: (t, p) => p in t ? t[p] : t.get(p) });

  function getActiveScope() {
    return state.isMergedView ? el.get('mergeViewContent') : el.get('displayArea');
  }

  function resetControls() {
    state.originalHtmlString = '';
    ['searchInput', 'selectKeywordInput'].forEach(id => { const e = el.get(id); if (e) e.value = ''; });
    el.get('selectKeywordRegex').checked = false;
    el.get('controlPanel').classList.add('hidden');
    updateSelectionInfo();
  }

  // ─────────────────────────────────────────────
  // 5. 檔案匯入與預處理 [A1]
  // ─────────────────────────────────────────────

  async function loadFundConfig() {
    try {
      const response = await fetch(`fund-config.json?v=${Date.now()}`);
      if (!response.ok) return;
      const config = await response.json();
      if (config.sortOrder && config.aliasMap) {
        state.fundSortOrder = config.sortOrder;
        state.fundAliasMap = config.aliasMap;
        state.fundAliasKeys = Object.keys(config.aliasMap).sort((a, b) => b.length - a.length);
      }
    } catch (err) { console.error('設定檔讀取失敗', err); }
  }

  function applyPreprocessing(jsonData, sheet, startRow, startCol, endCol) {
    const skipCount = el.get('skipTopRowsCheckbox')?.checked ? parseInt(el.get('skipTopRowsInput').value, 10) || 0 : 0;
    const removeEmpty = el.get('removeEmptyRowsCheckbox')?.checked ?? false;
    const removeKeywords = el.get('removeKeywordRowsCheckbox')?.checked ?? false;
    const keywords = removeKeywords
      ? el.get('removeKeywordRowsInput').value.split(',').map(k => k.trim().toLowerCase()).filter(Boolean)
      :[];

    const colProps = sheet['!cols'] || [];
    const rowProps = sheet['!rows'] ||[];
    const visibleCols =[];
    for (let c = startCol; c <= endCol; c++) {
      if (!(colProps[c] && colProps[c].hidden)) visibleCols.push(c - startCol);
    }

    const result =[];
    jsonData.forEach((row, idx) => {
      if (idx < skipCount) return; // 防線一
      const absRow = startRow + idx;
      if (rowProps[absRow]?.hidden) return; 

      const newRow = visibleCols.map(i => (row?.[i] ?? ''));
      const isEmpty = newRow.every(c => String(c).trim() === '');
      if (removeEmpty && isEmpty) return; // 防線二

      if (keywords.length > 0 && !isEmpty) {
        const content = newRow.join(' ').toLowerCase();
        if (keywords.some(k => content.includes(k))) return; // 防線三
      }
      result.push(newRow);
    });
    return result;
  }

  async function processFiles(fileList) {
    if (!fileList || fileList.length === 0) return;
    const files = Array.from(fileList).filter(f => VALID_EXTENSIONS.some(ext => f.name.toLowerCase().endsWith(ext)));
    if (!files.length) { alert('請上傳 Excel 檔案'); return; }
    if (state.isProcessing) return;

    const importMode = document.querySelector('input[name="import-mode"]:checked').value;
    const sheetCriteria = { name: el.get('specificSheetNameInput').value.trim(), position: el.get('specificSheetPositionInput').value.trim() };

    state.isProcessing = true;
    el.get('displayArea').innerHTML = '<div class="loading">讀取中...</div>';
    resetControls();
    
    state.rawSheetsCache = []; 
    state.loadedFiles =[];
    const tablesToRender =[];

    try {
      for (let i = 0; i < files.length; i++) {
        const file = files[i];
        el.get('displayArea').innerHTML = `<div class="loading">讀取中... (${i + 1}/${files.length})</div>`;

        const binary = await utils.readFileAsBinary(file);
        const workbook = XLSX.read(binary, { type: 'binary', cellStyles: true });
        const sheetNames = await getSelectedSheetNames(file.name, workbook, importMode, sheetCriteria);

        for (const sheetName of sheetNames) {
          const sheet = workbook.Sheets[sheetName];
          const ref = sheet['!ref'];
          const range = ref ? XLSX.utils.decode_range(ref) : { s: { r: 0, c: 0 }, e: { c: 0 } };
          const { r: startRow, c: startCol } = range.s;
          const endCol = range.e.c;

          let jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '', raw: false });

          if (sheet['!merges']) {
            sheet['!merges'].forEach(merge => {
              const sr = merge.s.r - startRow, sc = merge.s.c - startCol;
              const er = merge.e.r - startRow, ec = merge.e.c - startCol;
              if (sr >= 0 && sc >= 0 && jsonData[sr]) {
                const val = jsonData[sr][sc];
                for (let r = sr; r <= er; r++) {
                  for (let c = sc; c <= ec; c++) {
                    if (jsonData[r]) jsonData[r][c] = val;
                  }
                }
              }
            });
          }

          const label = `${file.name} (${sheetName})`;
          
          state.rawSheetsCache.push({
            label, startRow, startCol, endCol, sheet,
            jsonData: JSON.parse(JSON.stringify(jsonData))
          });

          const filtered = applyPreprocessing(jsonData, sheet, startRow, startCol, endCol);
          if (filtered.length > 0) {
            const cleanSheet = XLSX.utils.aoa_to_sheet(filtered);
            tablesToRender.push({ html: XLSX.utils.sheet_to_html(cleanSheet), filename: label });
            state.loadedFiles.push(label);
          }
        }
      }

      state.loadedTables = tablesToRender.length;
      renderTables(tablesToRender);
      updateDropAreaDisplay();
      markSettingsClean();
    } catch (err) {
      console.error(err);
      el.get('displayArea').innerHTML = `<p style="color:red;">處理錯誤：${err.message}</p>`;
    } finally {
      state.isProcessing = false;
    }
  }

  function reapplyPreprocessing() {
    if (!state.rawSheetsCache.length) return;
    state.isProcessing = true;
    el.get('displayArea').innerHTML = '<div class="loading">重新清洗中...</div>';
    
    setTimeout(() => {
      const tablesToRender = [];
      state.loadedFiles =[];
      
      state.rawSheetsCache.forEach(cache => {
        const filtered = applyPreprocessing(
          JSON.parse(JSON.stringify(cache.jsonData)), 
          cache.sheet, cache.startRow, cache.startCol, cache.endCol
        );
        if (filtered.length > 0) {
          const cleanSheet = XLSX.utils.aoa_to_sheet(filtered);
          tablesToRender.push({ html: XLSX.utils.sheet_to_html(cleanSheet), filename: cache.label });
          state.loadedFiles.push(cache.label);
        }
      });

      state.loadedTables = tablesToRender.length;
      renderTables(tablesToRender);
      updateDropAreaDisplay();
      markSettingsClean();
      state.isProcessing = false;
      
      undoManager.showToast('重新套用清洗設定');
    }, 50);
  }

  function markSettingsDirty() {
    if (state.loadedTables > 0) {
      state.isSettingsDirty = true;
      el.get('reapplyBanner').classList.remove('hidden');
    }
  }

  function markSettingsClean() {
    state.isSettingsDirty = false;
    el.get('reapplyBanner').classList.add('hidden');
  }

  async function getSelectedSheetNames(filename, workbook, mode, criteria) {
    const names = workbook.SheetNames;
    if (!names.length) return[];
    if (mode === 'first') return [names[0]];
    if (mode === 'specific') return names.filter(n => n.toLowerCase().includes(criteria.name.toLowerCase()));
    if (mode === 'position') return utils.parsePositionString(criteria.position).map(i => names[i]).filter(Boolean);
    return names;
  }

  // ─────────────────────────────────────────────
  // 6. 表格渲染 (主表)
  // ─────────────────────────────────────────────

  function renderTables(tables) {
    if (!tables.length) {
      el.get('displayArea').innerHTML = '<p>沒有找到符合條件的工作表。</p>';
      return;
    }

    const fragment = document.createDocumentFragment();
    tables.forEach(({ html, filename }) => {
      const wrapper = document.createElement('div');
      wrapper.className = 'table-wrapper';
      wrapper.innerHTML = `
        <div class="table-header">
          <input type="checkbox" class="table-select-checkbox" title="選取此表格">
          <h4>${filename}</h4>
          <div class="header-actions">
            <button class="btn btn-danger btn-sm delete-rows-btn">刪除選取列</button>
            <button class="btn btn-danger btn-sm delete-table-btn">刪除此表</button>
          </div>
          <button class="close-zoom">&times;</button>
        </div>
        <div class="table-content">${html}</div>`;
      fragment.appendChild(wrapper);
    });

    el.get('displayArea').innerHTML = '';
    el.get('displayArea').appendChild(fragment);
    state.originalHtmlString = el.get('displayArea').innerHTML;

    injectCheckboxes(el.get('displayArea'));
    showControls(detectHiddenElements());
    sortTablesByFundName();
  }

  function injectCheckboxes(scope) {
    scope.querySelectorAll('thead tr').forEach((row, idx) => {
      if (row.querySelector('.checkbox-cell')) return;
      row.insertAdjacentHTML('afterbegin', `<th class="checkbox-cell"><input type="checkbox" id="select-all-cb-${scope.id}-${idx}"></th>`);
    });
    scope.querySelectorAll('tbody tr').forEach(row => {
      if (row.querySelector('.checkbox-cell')) return;
      row.insertAdjacentHTML('afterbegin', '<td class="checkbox-cell"><input type="checkbox" class="row-checkbox"></td>');
    });
  }

  // ─────────────────────────────────────────────
  // 6.5 合併視圖 [A3]
  // ─────────────────────────────────────────────

  function createMergedView(mode = 'all') {
    const tables = Array.from(el.get('displayArea').querySelectorAll('.table-wrapper:not([style*="display: none"]) table'));
    if (!tables.length) { alert('沒有可合併的表格。'); return; }

    if (state.mergedData.length > 0 && state.isMergedView === false) {
      if (confirm('偵測到上次的合併紀錄，是否接續上次編輯狀態？\n(按「取消」將放棄舊狀態，重新合併最新主表)')) {
        el.get('mergeViewModal').classList.remove('hidden');
        document.body.classList.add('no-scroll');
        state.isMergedView = true;
        renderMergedTable();
        return;
      }
    }

    if (mode === 'checked') {
      const checkedCount = tables.reduce((n, t) => n + t.querySelectorAll('tbody .row-checkbox:checked').length, 0);
      if (!checkedCount) { alert('請先勾選至少一個資料列。'); return; }
    }

    const allHeaders = new Set();
    const tableHeaderMap = new Map();

    tables.forEach(table => {
      let headers = Array.from(table.querySelectorAll('thead th:not(.checkbox-cell)'))
        .map((th, i) => th.textContent.trim() || `(欄位 ${i + 1})`);
      if (!headers.length) headers = Array.from({ length: 10 }, (_, i) => `(欄位 ${i + 1})`);
      headers.forEach(h => allHeaders.add(h));
      tableHeaderMap.set(table, headers);
    });

    const tableData =[];
    tables.forEach(table => {
      const headers = tableHeaderMap.get(table);
      if (!headers) return;
      const filename = table.closest('.table-wrapper')?.querySelector('h4')?.textContent || '未知來源';
      const rows = mode === 'all'
        ? Array.from(table.querySelectorAll('tbody tr')).filter(r => !r.classList.contains('row-hidden-search'))
        : Array.from(table.querySelectorAll('tbody .row-checkbox:checked')).map(cb => cb.closest('tr'));

      rows.forEach(row => {
        const rowData = { _sourceFile: filename, _id: Date.now() + Math.random() };
        Array.from(row.querySelectorAll('td:not(.checkbox-cell)')).forEach((td, i) => {
          if (headers[i]) rowData[headers[i]] = td.textContent;
        });
        tableData.push(rowData);
      });
    });

    state.mergedHeaders = Array.from(allHeaders);
    state.mergedData = tableData;
    renderMergedTable();
    updateColumnSelects(state.mergedHeaders);
    
    state.isMergedView = true;
    el.get('mergeViewModal').classList.remove('hidden');
    document.body.classList.add('no-scroll');
  }

  function closeMergeView() {
    if (state.isEditing && !confirm('確定要放棄目前的編輯狀態？')) return;
    el.get('mergeViewModal').classList.add('hidden');
    document.body.classList.remove('no-scroll');
    state.isMergedView = false;
    toggleEditMode(false);
  }

  function renderMergedTable() {
    const table = document.createElement('table');
    const thead = table.createTHead();
    const headerRow = thead.insertRow();

    if (state.showSourceColumn) {
      headerRow.insertAdjacentHTML('beforeend', '<th class="source-col">來源檔案</th>');
    }

    state.mergedHeaders.forEach(header => {
      headerRow.insertAdjacentHTML('beforeend', `<th>${header}<span class="delete-col-btn" data-header="${header}">&times;</span></th>`);
    });

    const tbody = table.createTBody();
    state.mergedData.forEach((rowData, index) => {
      const tr = tbody.insertRow();
      tr.dataset.rowIndex = index;
      if (rowData._sourceFile) tr.title = `來源: ${rowData._sourceFile}`;
      if (rowData._isNew) tr.classList.add('new-row-highlight');

      if (state.showSourceColumn) {
        const td = document.createElement('td');
        td.textContent = rowData._sourceFile || '';
        td.classList.add('source-col');
        tr.prepend(td);
      }

      state.mergedHeaders.forEach(header => {
        const td = tr.insertCell();
        td.contentEditable = state.isEditing;
        td.dataset.colHeader = header;
        const value = rowData[header] || '';
        if (utils.isStrictNumber(value)) {
          td.textContent = utils.formatNumber(value);
          td.classList.add('numeric');
        } else {
          td.textContent = value;
        }
      });
    });

    if (state.showTotalRow) {
      const tfoot = table.createTFoot();
      const totalRow = tfoot.insertRow();
      totalRow.className = 'total-row';
      totalRow.insertCell(); // checkbox 佔位
      if (state.showSourceColumn) totalRow.insertCell().textContent = '';

      const totals = calculateTotals();
      let labelApplied = false;
      state.mergedHeaders.forEach(header => {
        const td = totalRow.insertCell();
        if (totals[header] !== undefined) {
          td.textContent = utils.formatNumber(totals[header]);
          td.classList.add('numeric');
        } else if (!labelApplied) {
          td.textContent = '合計';
          labelApplied = true;
        }
      });
    }

    el.get('mergeViewContent').innerHTML = '';
    el.get('mergeViewContent').appendChild(table);
    el.get('mergeViewContent').classList.toggle('is-editing', state.isEditing);
    injectCheckboxes(el.get('mergeViewContent'));

    const selectAllCb = el.get('mergeViewContent').querySelector('thead input[type="checkbox"]');
    if (selectAllCb) selectAllCb.addEventListener('change', e => {
      el.get('mergeViewContent').querySelectorAll('.row-checkbox').forEach(cb => cb.checked = e.target.checked);
    });
  }

  // ─────────────────────────────────────────────
  // 7. 選取 / 刪除 / 撤銷操作 [A4]
  // ─────────────────────────────────────────────

  function deleteSelectedRows(specificScope = null) {
    const scope = specificScope || getActiveScope();
    const selected = scope.querySelectorAll('tbody .row-checkbox:checked');
    if (!selected.length) { alert('請先勾選列'); return; }

    if (state.isMergedView) {
      const backupData = [...state.mergedData];
      const toDelete = new Set(Array.from(selected).map(cb => parseInt(cb.closest('tr').dataset.rowIndex, 10)));
      
      undoManager.push(`刪除 ${toDelete.size} 列 (合併視圖)`, () => {
        state.mergedData = backupData;
        if (state.isMergedView) renderMergedTable();
      });

      state.mergedData = state.mergedData.filter((_, i) => !toDelete.has(i));
      renderMergedTable();
      el.get('dedupResultPanel').classList.add('hidden');

    } else {
      const rows = Array.from(selected).map(cb => cb.closest('tr'));
      const backups = rows.map(tr => ({ parent: tr.parentNode, sibling: tr.nextSibling, node: tr }));
      
      undoManager.push(`刪除 ${rows.length} 列 (主表)`, () => {
        backups.forEach(b => b.parent.insertBefore(b.node, b.sibling));
        syncCheckboxesInScope();
      });

      rows.forEach(tr => tr.remove());
    }
    syncCheckboxesInScope();
  }

  function deleteSelectedTables(specificTableWrapper = null) {
    const selected = specificTableWrapper 
      ? [specificTableWrapper] 
      : Array.from(el.get('displayArea').querySelectorAll('.table-select-checkbox:checked')).map(cb => cb.closest('.table-wrapper'));
    
    if (!selected.length) return;
    
    const backups = selected.map(w => ({ parent: w.parentNode, sibling: w.nextSibling, node: w }));
    const previousLoadedCount = state.loadedTables;
    const previousFiles = [...state.loadedFiles];

    undoManager.push(`刪除 ${selected.length} 個工作表`, () => {
      backups.forEach(b => b.parent.insertBefore(b.node, b.sibling));
      state.loadedTables = previousLoadedCount;
      state.loadedFiles = previousFiles;
      updateDropAreaDisplay();
      showControls(detectHiddenElements());
    });

    selected.forEach(w => w.remove());
    updateFileStateAfterDeletion();
  }

  function deleteColumn(headerToDelete) {
    if (!confirm(`確定要刪除「${headerToDelete}」欄位？`)) return;
    
    const backupHeaders = [...state.mergedHeaders];
    const backupData = state.mergedData.map(row => ({...row}));

    undoManager.push(`刪除欄位: ${headerToDelete}`, () => {
      state.mergedHeaders = backupHeaders;
      state.mergedData = backupData;
      if (state.isMergedView) {
        renderMergedTable();
        updateColumnSelects(state.mergedHeaders);
      }
    });

    state.mergedHeaders = state.mergedHeaders.filter(h => h !== headerToDelete);
    state.mergedData.forEach(row => delete row[headerToDelete]);
    renderMergedTable();
    updateColumnSelects(state.mergedHeaders);
  }

  function syncCheckboxesInScope() {
    setTimeout(() => {
      getActiveScope().querySelectorAll('table').forEach(t => {
        const rowCbs = t.querySelectorAll('tbody tr:not(.row-hidden-search) .row-checkbox');
        const hcb = t.closest('.table-wrapper')?.querySelector('.table-select-checkbox') ?? t.querySelector('thead input[type="checkbox"]');
        if (!hcb) return;
        const checkedCount = Array.from(rowCbs).filter(cb => cb.checked).length;
        hcb.checked = rowCbs.length > 0 && checkedCount === rowCbs.length;
        hcb.indeterminate = checkedCount > 0 && checkedCount < rowCbs.length;
      });
      if (!state.isMergedView) updateSelectionInfo();
    }, 0);
  }

  // ─────────────────────────────────────────────
  // 8. 篩選、選取與其他操作
  // ─────────────────────────────────────────────

  function selectAllRows() {
    const scope = getActiveScope();
    scope.querySelectorAll('tbody tr:not(.row-hidden-search)').forEach(row => {
      const cb = row.querySelector('.row-checkbox');
      if (cb) cb.checked = true;
    });
    scope.querySelectorAll('thead input[type="checkbox"]').forEach(cb => cb.checked = true);
  }

  function invertSelection() {
    getActiveScope().querySelectorAll('tbody tr:not(.row-hidden-search) .row-checkbox').forEach(cb => {
      cb.checked = !cb.checked;
    });
  }

  function selectEmptyRows() {
    let count = 0;
    getActiveScope().querySelectorAll('tbody tr:not(.row-hidden-search)').forEach(row => {
      const isBlank = Array.from(row.cells).slice(1).every(c => c.textContent.trim() === '');
      if (isBlank) {
        row.querySelector('.row-checkbox').checked = true;
        count++;
      }
    });
    if (!count) alert('未找到空白列');
  }

  function unselectAllMergedRows() {
    if (!state.isMergedView) return;
    el.get('mergeViewContent').querySelectorAll('.row-checkbox:checked').forEach(cb => cb.checked = false);
    const hcb = el.get('mergeViewContent').querySelector('thead input[type="checkbox"]');
    if (hcb) { hcb.checked = false; hcb.indeterminate = false; }
  }

  function buildKeywordMatcher(keyword, isRegex) {
    if (!keyword) return null;
    if (isRegex) return text => new RegExp(keyword, 'i').test(text);
    if (keyword.includes(',')) {
      const kws = keyword.split(',').map(k => k.trim().toLowerCase()).filter(Boolean);
      return text => kws.some(k => text.includes(k));
    }
    const kws = keyword.split(/\s+/).map(k => k.trim().toLowerCase()).filter(Boolean);
    return text => kws.every(k => text.includes(k));
  }

  function selectByKeyword() {
    const inputEl = state.isMergedView ? el.get('selectKeywordInputMerged') : el.get('selectKeywordInput');
    const regexEl = state.isMergedView ? el.get('selectKeywordRegexMerged') : el.get('selectKeywordRegex');
    const keyword = inputEl.value.trim();
    if (!keyword) { alert('請輸入關鍵字'); return; }

    let matcher;
    try {
      matcher = buildKeywordMatcher(keyword, regexEl.checked);
    } catch (e) {
      alert('Regex 錯誤：' + e.message);
      return;
    }

    let count = 0;
    getActiveScope().querySelectorAll('tbody tr:not(.row-hidden-search)').forEach(row => {
      const text = Array.from(row.querySelectorAll('td:not(.checkbox-cell)')).map(c => c.textContent).join(' ').toLowerCase();
      if (matcher(text)) {
        row.querySelector('.row-checkbox').checked = true;
        count++;
      }
    });
    alert(count > 0 ? `已勾選 ${count} 列` : '未找到相符資料');
  }

  function filterTable() {
    const inputEl = state.isMergedView ? el.get('searchInputMerged') : el.get('searchInput');
    const keywords = inputEl.value.toLowerCase().trim().split(/\s+/).filter(Boolean);
    const scope = getActiveScope();

    scope.querySelectorAll('tbody tr').forEach(row => {
      const text = Array.from(row.querySelectorAll('td:not(.checkbox-cell)')).map(c => c.textContent).join(' ').toLowerCase();
      row.classList.toggle('row-hidden-search', !keywords.every(k => text.includes(k)));
    });

    if (!state.isMergedView) {
      el.get('displayArea').querySelectorAll('.table-wrapper').forEach(wrapper => {
        const hasVisible = wrapper.querySelectorAll('tbody tr:not(.row-hidden-search)').length > 0;
        wrapper.style.display = hasVisible ? '' : 'none';
      });
    }
    syncCheckboxesInScope();
  }

  function executeCombinedSelection() {
    if (!state.isMergedView) return;

    const keyword = el.get('selectKeywordInputMerged').value.trim();
    const isRegex = el.get('selectKeywordRegexMerged').checked;
    let keywordMatcher = null;
    if (keyword) {
      try { keywordMatcher = buildKeywordMatcher(keyword, isRegex); }
      catch (e) { alert('Regex 錯誤'); return; }
    }

    const col1 = el.get('colSelect1').value;
    const col2 = el.get('colSelect2').value;
    const criteria1 = document.querySelector('input[name="criteria-1"]:checked')?.value;
    const criteria2 = document.querySelector('input[name="criteria-2"]:checked')?.value;
    const logicOp = document.querySelector('input[name="logic-op"]:checked')?.value ?? 'and';
    const inputVal1 = el.get('inputCriteria1').value;
    const inputVal2 = el.get('inputCriteria2').value;

    if (!keyword && !col1 && !col2) { alert('請輸入至少一個條件'); return; }

    const checkValue = (cellVal, cr, val) => {
      const s = String(cellVal).trim(), v = String(val).trim();
      return cr === 'empty'    ? s === ''
           : cr === 'zero'     ? s === '0'
           : cr === 'value'    ? s !== ''
           : cr === 'exact'    ? s === v
           : cr === 'includes' ? v !== '' && s.toLowerCase().includes(v.toLowerCase())
           : false;
    };

    let count = 0;
    el.get('mergeViewContent').querySelectorAll('tbody tr:not(.row-hidden-search)').forEach(row => {
      const checkbox = row.querySelector('.row-checkbox');
      if (!checkbox) return;

      const rowText = Array.from(row.querySelectorAll('td:not(.checkbox-cell)')).map(c => c.textContent).join(' ');
      const kMatch = keywordMatcher ? keywordMatcher(rowText) : false;

      let cMatch = false;
      if (col1 || col2) {
        const r1 = col1 ? checkValue(row.querySelector(`td[data-col-header="${col1}"]`)?.textContent ?? '', criteria1, inputVal1) : null;
        const r2 = col2 ? checkValue(row.querySelector(`td[data-col-header="${col2}"]`)?.textContent ?? '', criteria2, inputVal2) : null;
        cMatch = col1 && col2
          ? (logicOp === 'and' ? r1 && r2 : r1 || r2)
          : (col1 ? r1 : r2);
      }

      if (kMatch || cMatch) { checkbox.checked = true; count++; }
    });

    alert(count > 0 ? `已勾選 ${count} 筆` : '未找到符合條件的資料');
  }

  // ─────────────────────────────────────────────
  // 9. 編輯與欄位操作
  // ─────────────────────────────────────────────

  function toggleEditMode(startEditing) {
    state.isEditing = startEditing;
    el.get('editDataBtn').classList.toggle('hidden', startEditing);
    el.get('saveEditsBtn').classList.toggle('hidden', !startEditing);
    el.get('cancelEditsBtn').classList.toggle('hidden', !startEditing);
    
    const toggleIds =[
      'addNewRowBtn', 'copySelectedRowsBtn', 'deleteMergedRowsBtn', 'columnOperationsBtn',
      'toggleTotalRowBtn', 'toggleSourceColBtn', 'invertSelectionMergedBtn',
      'exportCurrentMergedXlsxBtn', 'sortMergedByNameBtn', 'colSelect1', 'colSelect2',
      'executeFilterSelectionBtn', 'searchInputMerged', 'selectKeywordInputMerged',
      'selectKeywordRegexMerged', 'unselectMergedRowsBtn', 'smartDedupBtn',
    ];
    toggleIds.forEach(id => {
      const elem = el.get(id);
      if (elem) elem.disabled = startEditing;
    });

    ['inputCriteria1', 'inputCriteria2'].forEach(id => { el.get(id).disabled = true; });
    document.querySelectorAll('input[name="criteria-1"], input[name="criteria-2"], input[name="logic-op"]')
      .forEach(r => r.disabled = startEditing);

    renderMergedTable();
  }
  
  function saveEdits() {
    const backupData =[...state.mergedData];
    const newData = Array.from(el.get('mergeViewContent').querySelectorAll('tbody tr')).map(tr => {
      const row = {};
      const origIdx = parseInt(tr.dataset.rowIndex, 10);
      row._sourceFile = state.showSourceColumn ? tr.querySelector('.source-col').textContent : (state.mergedData[origIdx]?._sourceFile || '(修改)');
      tr.querySelectorAll('td[data-col-header]').forEach(cell => row[cell.dataset.colHeader] = cell.textContent);
      return row;
    });
    
    undoManager.push('儲存編輯', () => {
      state.mergedData = backupData;
      if (state.isMergedView) renderMergedTable();
    });

    state.mergedData = newData;
    toggleEditMode(false);
  }

  function addNewRow() {
    const newRow = { _isNew: true, _sourceFile: '(新增資料列)' };
    state.mergedHeaders.forEach(h => newRow[h] = '');
    state.mergedData.unshift(newRow);
    toggleEditMode(true);
  }

  function copySelectedRows() {
    const selected = el.get('mergeViewContent').querySelectorAll('.row-checkbox:checked');
    if (!selected.length) { alert('請先勾選要複製的資料列。'); return; }

    const copies = Array.from(selected).map(cb => {
      const idx = parseInt(cb.closest('tr').dataset.rowIndex, 10);
      if (isNaN(idx) || !state.mergedData[idx]) return null;
      const copy = JSON.parse(JSON.stringify(state.mergedData[idx]));
      copy._isNew = true;
      copy._sourceFile += ' (複製)';
      return copy;
    }).filter(Boolean);

    state.mergedData.unshift(...copies);
    toggleEditMode(true);
  }

  function toggleSourceColumn() {
    if (state.isEditing) { alert('請先儲存或取消編輯。'); return; }
    state.showSourceColumn = !state.showSourceColumn;
    renderMergedTable();
    el.get('toggleSourceColBtn').textContent = state.showSourceColumn ? '移除來源欄位' : '新增來源欄位';
    el.get('toggleSourceColBtn').classList.toggle('active', state.showSourceColumn);
  }

  function calculateTotals() {
    const totals = {};
    state.mergedHeaders.forEach(header => {
      const sum = state.mergedData.reduce((acc, row) => {
        const n = parseFloat(String(row[header] || '').replace(/,/g, ''));
        return acc + (isNaN(n) ? 0 : n);
      }, 0);
      const hasAnyNumber = state.mergedData.some(row => {
        return !isNaN(parseFloat(String(row[header] || '').replace(/,/g, '')));
      });
      if (sum !== 0 || hasAnyNumber) totals[header] = sum;
    });
    return totals;
  }

  function updateColumnSelects(headers) {
    const checklist = el.get('columnChecklist');
    if(checklist) {
        checklist.innerHTML = headers.map(h => `<label><input type="checkbox" value="${h}" checked> ${h}</label>`).join('');
    }

    const makeOption = (value, text) => {
      const opt = document.createElement('option');
      opt.value = value; opt.textContent = text;
      return opt;
    };
    
    const col1 = el.get('colSelect1');
    if(col1) {
        col1.innerHTML = '';
        col1.appendChild(makeOption('', '-- 選擇欄位 1 --'));
        headers.forEach(h => col1.appendChild(makeOption(h, h)));
    }

    const col2 = el.get('colSelect2');
    if(col2) {
        col2.innerHTML = '';
        col2.appendChild(makeOption('', '-- 選擇欄位 2 (選填) --'));
        headers.forEach(h => col2.appendChild(makeOption(h, h)));
    }
  }

  function toggleColumnModal(show) {
    el.get('columnModal').classList.toggle('hidden', !show);
  }

  function applyColumnChanges() {
    const mergedTable = el.get('mergeViewContent').querySelector('table');
    if (!mergedTable) return;

    const visibility = {};
    el.get('columnChecklist').querySelectorAll('input').forEach(input => {
      visibility[input.value] = input.checked;
    });

    const allThs = Array.from(mergedTable.querySelectorAll('thead th'));
    const firstDataIdx = allThs.findIndex(th => !th.classList.contains('checkbox-cell') && !th.classList.contains('source-col'));
    if (firstDataIdx === -1) return;

    allThs.slice(firstDataIdx).forEach((th, i) => {
      const colIdx = i + firstDataIdx;
      const headerText = th.textContent.replace('×', '').trim();
      mergedTable.querySelectorAll(`tr > *:nth-child(${colIdx + 1})`).forEach(cell => {
        cell.classList.toggle('column-hidden', !visibility[headerText]);
      });
    });
  }

  // ─────────────────────────────────────────────
  // 10. 匯出與排序
  // ─────────────────────────────────────────────

  function extractTableData(table, { onlySelected = false, includeFilename = false } = {}) {
    const data =[];
    const headerRow = table.querySelector('thead tr');
    if (headerRow) {
      const headers = Array.from(headerRow.querySelectorAll('th:not(.checkbox-cell):not(.column-hidden)'))
        .map(th => th.textContent.replace('×', '').trim());
      if (includeFilename) headers.unshift('Source File');
      data.push(headers);
    }

    const filename = includeFilename
      ? (table.closest('.table-wrapper')?.querySelector('h4')?.textContent || 'Merged Table')
      : null;

    const rows = onlySelected
      ? Array.from(table.querySelectorAll('tbody .row-checkbox:checked')).map(cb => cb.closest('tr'))
      : Array.from(table.querySelectorAll('tbody tr:not(.row-hidden-search)'));

    rows.forEach(row => {
      const cells = Array.from(row.querySelectorAll('td:not(.checkbox-cell):not(.column-hidden)'))
        .map(td => {
          const val = td.textContent.trim();
          const clean = val.replace(/,/g, '');
          return clean !== '' && !isNaN(clean) ? Number(clean) : val;
        });
      if (includeFilename) cells.unshift(filename);
      data.push(cells);
    });
    return data;
  }

  function exportToXlsx(data, filename, sheetName) {
    if (data.length <= 1) { alert('無資料可匯出。'); return; }
    try {
      const ws = XLSX.utils.aoa_to_sheet(data);
      ws['!cols'] = utils.calculateColumnWidths(data);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, sheetName);
      XLSX.writeFile(wb, filename);
    } catch (err) {
      alert('匯出錯誤：' + err.message);
    }
  }

  function exportCurrentMergedXlsx() {
    if (!state.isMergedView) return;
    const table = el.get('mergeViewContent').querySelector('table');
    if (!table) return;
    exportToXlsx(
      extractTableData(table, { includeFilename: state.showSourceColumn }),
      `merged_view_${new Date().toISOString().slice(0, 10)}.xlsx`,
      'Merged Data'
    );
  }

  function exportMergedXlsx() {
    const tables = Array.from(
      el.get('displayArea').querySelectorAll('.table-wrapper:not([style*="display: none"]) table')
    );
    if (!tables.length) { alert('沒有可匯出的表格。'); return; }

    const allData =[];
    tables.forEach((table, i) => {
      const data = extractTableData(table, { includeFilename: true });
      if (data.length > 1) allData.push(...(i === 0 ? data : data.slice(1)));
    });
    exportToXlsx(allData, `report_${new Date().toISOString().slice(0, 10)}.xlsx`, 'Data');
  }

  function getFundSortPriority(fileName) {
    if (!state.fundSortOrder.length) return { index: Infinity, name: fileName };
    const alias = state.fundAliasKeys.find(a => fileName.includes(a));
    const canonical = alias ? state.fundAliasMap[alias] : null;
    const index = canonical ? state.fundSortOrder.indexOf(canonical) : -1;
    return { index: index === -1 ? Infinity : index, name: fileName };
  }

  function sortTablesByFundName() {
    if (!state.fundSortOrder.length) return;
    const wrappers = Array.from(el.get('displayArea').querySelectorAll('.table-wrapper'));
    wrappers.sort((a, b) => {
      const fa = getFundSortPriority(a.querySelector('h4').textContent);
      const fb = getFundSortPriority(b.querySelector('h4').textContent);
      return fa.index !== fb.index ? fa.index - fb.index : fa.name.localeCompare(fb.name);
    });
    el.get('displayArea').innerHTML = '';
    wrappers.forEach(w => el.get('displayArea').appendChild(w));
  }

  function sortMergedTableByFundName() {
    if (state.isEditing) { alert('請先儲存或取消編輯。'); return; }
    if (!state.fundSortOrder.length) return;
    state.mergedData.sort((a, b) => {
      const fa = getFundSortPriority(a._sourceFile || '');
      const fb = getFundSortPriority(b._sourceFile || '');
      return fa.index !== fb.index ? fa.index - fb.index : (a._sourceFile || '').localeCompare(b._sourceFile || '');
    });
    renderMergedTable();
  }

  function handleMergedHeaderClick(th) {
    if (state.isEditing || th.classList.contains('source-col')) return;
    const table = th.closest('table');
    const header = th.textContent.replace('×', '').trim();
    const isAsc = th.classList.contains('sort-asc');
    table.querySelectorAll('th').forEach(h => h.classList.remove('sort-asc', 'sort-desc'));
    th.classList.add(isAsc ? 'sort-desc' : 'sort-asc');

    state.mergedData.sort((a, b) => {
      const va = a[header] || '', vb = b[header] || '';
      const na = parseFloat(String(va).replace(/,/g, '')), nb = parseFloat(String(vb).replace(/,/g, ''));
      const cmp = !isNaN(na) && !isNaN(nb) ? na - nb : va.localeCompare(vb, undefined, { numeric: true, sensitivity: 'base' });
      return isAsc ? -cmp : cmp;
    });
    renderMergedTable();
  }

  // ─────────────────────────────────────────────
  // 11. 智慧去重 [B3]
  // ─────────────────────────────────────────────

  function executeSmartDeduplication() {
    const keyCol = el.get('dedupColSelect').value;
    if (!keyCol) return;

    const groups = {};
    const rows = Array.from(el.get('mergeViewContent').querySelectorAll('tbody tr:not(.row-hidden-search)'));

    rows.forEach(tr => tr.classList.remove('dedup-marked'));

    rows.forEach(tr => {
      const idx = parseInt(tr.dataset.rowIndex, 10);
      const rowData = state.mergedData[idx];
      if (!rowData) return;
      const key = String(rowData[keyCol] || '').trim();
      if (!key) return;
      if (!groups[key]) groups[key] = [];
      groups[key].push({ tr, rowData });
    });

    let markedCount = 0;
    Object.entries(groups).forEach(([key, items]) => {
      if (items.length <= 1) return;

      const cleanKey = key.replace(/[\s_基金]/g, '');
      let best = items.find(({ rowData }) => {
        const src = String(rowData._sourceFile || '').replace(/[\s_]/g, '');
        return src.includes(cleanKey) || cleanKey.includes(src);
      }) ?? items[0];

      items.forEach(item => {
        if (item === best) return;
        const cb = item.tr.querySelector('.row-checkbox');
        if (cb && !cb.checked) { 
          cb.checked = true; 
          item.tr.classList.add('dedup-marked'); 
          markedCount++; 
        }
      });
    });

    el.get('dedupModal').classList.add('hidden');
    syncCheckboxesInScope();

    if (markedCount > 0) {
      el.get('dedupResultText').innerHTML = `🎯 <b>智慧去重完成：</b> 已為您自動標記並勾選了 <b>${markedCount}</b> 筆不符合來源規則的舊資料。`;
      el.get('dedupResultPanel').classList.remove('hidden');
    } else {
      undoManager.showToast('未發現需要處理的重複資料');
    }
  }

  function clearDedupMarks() {
    el.get('mergeViewContent').querySelectorAll('.dedup-marked').forEach(tr => {
      tr.classList.remove('dedup-marked');
      tr.querySelector('.row-checkbox').checked = false;
    });
    el.get('dedupResultPanel').classList.add('hidden');
    syncCheckboxesInScope();
  }

  // ─────────────────────────────────────────────
  // 12. UI 狀態與其他
  // ─────────────────────────────────────────────

  function updateDropAreaDisplay() {
    const hasFiles = state.loadedTables > 0;
    el.get('dropArea').classList.toggle('compact', hasFiles);
    el.get('dropAreaInitial').classList.toggle('hidden', hasFiles);
    el.get('dropAreaLoaded').classList.toggle('hidden', !hasFiles);
    el.get('importOptionsContainer').classList.toggle('hidden', hasFiles);
    if (hasFiles) {
      el.get('fileCount').textContent = state.loadedTables;
      el.get('fileNames').textContent = state.loadedFiles.slice(0, 3).join(', ') + (state.loadedFiles.length > 3 ? '...' : '');
    }
  }

  function showControls(hiddenCount) { 
    el.get('controlPanel').classList.remove('hidden'); 
    el.get('mergeViewBtn').classList.toggle('hidden', state.loadedTables <= 1); 
    el.get('showHiddenBtn').classList.toggle('hidden', hiddenCount === 0);
  }

  function updateSelectionInfo() {
    const selected = el.get('displayArea').querySelectorAll('.table-select-checkbox:checked, .table-select-checkbox:indeterminate');
    el.get('selectedTablesList').textContent = Array.from(selected).map(cb => cb.closest('.table-header').querySelector('h4').textContent).join('; ');
    el.get('selectedTablesInfo').classList.toggle('hidden', selected.length === 0);
  }

  function detectHiddenElements() {
    return el.get('displayArea').querySelectorAll(
      'tr[style*="display: none"], td[style*="display: none"], th[style*="display: none"]'
    ).length;
  }

  function updateFileStateAfterDeletion() {
    state.loadedTables = el.get('displayArea').querySelectorAll('.table-wrapper').length;
    if (!state.loadedTables) clearAllFiles(true);
    else { updateDropAreaDisplay(); updateSelectionInfo(); }
  }

  function clearAllFiles(silent = false) {
    if (!silent && !confirm('確定清除所有檔案？')) return;
    if (state.isMergedView) closeMergeView();
    state.originalHtmlString = ''; state.loadedFiles =[]; state.loadedTables = 0; state.rawSheetsCache =[];
    el.get('displayArea').innerHTML = ''; el.get('fileInput').value = '';
    updateDropAreaDisplay(); resetControls(); setViewMode('list');
  }

  function setViewMode(mode) {
    const isGrid = mode === 'grid';
    el.get('displayArea').classList.toggle('grid-view', isGrid);
    el.get('displayArea').classList.toggle('list-view', !isGrid);
    el.get('gridViewBtn').classList.toggle('active', isGrid);
    el.get('listViewBtn').classList.toggle('active', !isGrid);
    el.get('gridScaleControl').classList.toggle('hidden', !isGrid);
  }

  function updateGridScale() {
    el.get('displayArea').style.setProperty('--grid-columns', el.get('gridScaleSlider').value);
  }

  function showAllHiddenElements() {
    const hidden = el.get('displayArea').querySelectorAll(
      'tr[style*="display: none"], td[style*="display: none"], th[style*="display: none"]'
    );
    if (!hidden.length) return;
    hidden.forEach(el => el.style.display = '');
    el.get('showHiddenBtn').classList.add('hidden');
    el.get('loadStatusMessage').classList.add('hidden');
  }

  function toggleToolbar() {
    const collapsed = el.get('collapsibleToolbar').classList.toggle('collapsed');
    el.get('toggleToolbarBtn').textContent = collapsed ? '展開工具列' : '收合工具列';
  }

  function scrollToTop() {
    window.scrollTo({ top: 0, behavior: 'smooth' });
  }

  function handleScroll() {
    el.get('backToTopBtn').classList.toggle('visible', window.scrollY > window.innerHeight / 2);
  }

  function openPreview(card) {
    if (state.zoomedCard) return;
    card.classList.add('is-zoomed');
    state.zoomedCard = card;
    document.body.classList.add('no-scroll');
  }

  function closePreview() {
    if (!state.zoomedCard) return;
    state.zoomedCard.classList.remove('is-zoomed');
    state.zoomedCard = null;
    document.body.classList.remove('no-scroll');
  }

  function resetView() {
    if (state.isMergedView) closeMergeView();
    if (!state.originalHtmlString) return;
    el.get('displayArea').innerHTML = state.originalHtmlString;
    injectCheckboxes(el.get('displayArea'));
    ['searchInput', 'selectKeywordInput'].forEach(id => { el.get(id).value = ''; });
    el.get('selectKeywordRegex').checked = false;
    filterTable();
    updateSelectionInfo();
    setViewMode('list');
  }

  function handleCriteriaChange(e) {
    const radio = e.target;
    if (radio.type !== 'radio') return;
    const group = radio.closest('.radio-group');
    if (!group) return;
    const target = el.get(group.dataset.target);
    if (!target) return;
    const needsInput = radio.value === 'exact' || radio.value === 'includes';
    target.disabled = !needsInput;
    if (needsInput) { target.focus(); } else { target.value = ''; }
  }

  // ─────────────────────────────────────────────
  // 13. 事件綁定
  // ─────────────────────────────────────────────

  function bindEvents() {
    const dropArea = el.get('dropArea');
    const fileInput = el.get('fileInput');

    // ── 1. 上傳與拖曳事件 (增加防呆與錯誤檢查) ──
    if (dropArea && fileInput) {
      ['dragenter', 'dragover'].forEach(e => dropArea.addEventListener(e, ev => { 
        ev.preventDefault(); 
        dropArea.classList.add('highlight'); 
      }));
      
      ['dragleave', 'drop'].forEach(e => dropArea.addEventListener(e, ev => { 
        ev.preventDefault(); 
        dropArea.classList.remove('highlight'); 
      }));
      
      dropArea.addEventListener('drop', e => processFiles(e.dataTransfer.files));
      
      // 確保點擊外框或按鈕時能正確觸發 input，並防止無限迴圈
      dropArea.addEventListener('click', (e) => {
        if (e.target !== fileInput) {
          fileInput.click();
        }
      });

      fileInput.addEventListener('change', e => {
        processFiles(e.target.files);
        e.target.value = ''; // 清空 value，確保重複上傳同一個檔案也能觸發
      });
    } else {
      console.error("❌ 找不到上傳區塊 (dropArea) 或檔案輸入框 (fileInput)！");
    }

    el.get('clearFilesBtn')?.addEventListener('click', () => clearAllFiles(false));

    // ── 2. 匯入設定 ──
    el.get('importOptionsContainer')?.addEventListener('change', e => {
      if (e.target.name !== 'import-mode') return;
      const mode = e.target.value;
      el.get('specificSheetNameGroup').classList.toggle('hidden', mode !== 'specific');
      el.get('specificSheetPositionGroup').classList.toggle('hidden', mode !== 'position');
    });

    // ── 3. 預處理開關與重新套用 [A1] ──['skipTopRowsCheckbox', 'skipTopRowsInput', 'removeEmptyRowsCheckbox', 'removeKeywordRowsCheckbox', 'removeKeywordRowsInput'].forEach(id => {
      el.get(id)?.addEventListener('change', markSettingsDirty);
      el.get(id)?.addEventListener('input', utils.debounce(markSettingsDirty, 500));
    });
    el.get('skipTopRowsCheckbox')?.addEventListener('change', e => el.get('skipTopRowsInput').disabled = !e.target.checked);
    el.get('removeKeywordRowsCheckbox')?.addEventListener('change', e => el.get('removeKeywordRowsInput').disabled = !e.target.checked);
    el.get('reapplySettingsBtn')?.addEventListener('click', reapplyPreprocessing);

    // ── 4. 視圖切換 ──
    el.get('listViewBtn')?.addEventListener('click', () => setViewMode('list'));
    el.get('gridViewBtn')?.addEventListener('click', () => setViewMode('grid'));
    el.get('gridScaleSlider')?.addEventListener('input', updateGridScale);

    // ── 5. 表格層級 ──
    el.get('selectAllTablesBtn')?.addEventListener('click', () => { el.get('displayArea').querySelectorAll('.table-select-checkbox').forEach(cb => cb.checked = true); updateSelectionInfo(); });
    el.get('unselectAllTablesBtn')?.addEventListener('click', () => { el.get('displayArea').querySelectorAll('.table-select-checkbox').forEach(cb => cb.checked = false); updateSelectionInfo(); });
    el.get('deleteSelectedTablesBtn')?.addEventListener('click', () => deleteSelectedTables());
    el.get('sortByNameBtn')?.addEventListener('click', sortTablesByFundName);

    // ── 6. 列層級操作（主表） ──
    el.get('selectAllBtn')?.addEventListener('click', () => { selectAllRows(); syncCheckboxesInScope(); });
    el.get('invertSelectionBtn')?.addEventListener('click', () => { invertSelection(); syncCheckboxesInScope(); });
    el.get('selectEmptyBtn')?.addEventListener('click', () => { selectEmptyRows(); syncCheckboxesInScope(); });
    el.get('selectByKeywordBtn')?.addEventListener('click', () => { selectByKeyword(); syncCheckboxesInScope(); });
    el.get('deleteSelectedBtn')?.addEventListener('click', () => deleteSelectedRows());
    el.get('searchInput')?.addEventListener('input', utils.debounce(filterTable, 300));

    // ── 7. 全域工具 ──
    el.get('resetViewBtn')?.addEventListener('click', resetView);
    el.get('showHiddenBtn')?.addEventListener('click', showAllHiddenElements);
    el.get('exportMergedXlsxBtn')?.addEventListener('click', exportMergedXlsx);

    // ── 8. 合併視圖開啟 ──
    el.get('mergeViewBtn')?.addEventListener('click', () => createMergedView('all'));
    el.get('viewCheckedCombinedBtn')?.addEventListener('click', () => createMergedView('checked'));
    el.get('closeMergeViewBtn')?.addEventListener('click', closeMergeView);

    // ── 9. 合併視圖工具列 ──
    el.get('searchInputMerged')?.addEventListener('input', utils.debounce(filterTable, 300));
    el.get('executeFilterSelectionBtn')?.addEventListener('click', () => { executeCombinedSelection(); syncCheckboxesInScope(); });
    el.get('invertSelectionMergedBtn')?.addEventListener('click', () => { invertSelection(); syncCheckboxesInScope(); });
    el.get('unselectMergedRowsBtn')?.addEventListener('click', unselectAllMergedRows);
    el.get('toggleToolbarBtn')?.addEventListener('click', toggleToolbar);

    // ── 10. 合併視圖操作 ──
    el.get('editDataBtn')?.addEventListener('click', () => toggleEditMode(true));
    el.get('saveEditsBtn')?.addEventListener('click', saveEdits);
    el.get('cancelEditsBtn')?.addEventListener('click', () => toggleEditMode(false));
    el.get('addNewRowBtn')?.addEventListener('click', addNewRow);
    el.get('copySelectedRowsBtn')?.addEventListener('click', copySelectedRows);
    el.get('deleteMergedRowsBtn')?.addEventListener('click', () => deleteSelectedRows());
    el.get('toggleTotalRowBtn')?.addEventListener('click', () => { state.showTotalRow = !state.showTotalRow; renderMergedTable(); });
    el.get('toggleSourceColBtn')?.addEventListener('click', toggleSourceColumn);
    el.get('exportCurrentMergedXlsxBtn')?.addEventListener('click', exportCurrentMergedXlsx);
    el.get('sortMergedByNameBtn')?.addEventListener('click', sortMergedTableByFundName);

    // ── 11. 欄位 Modal ──
    el.get('columnOperationsBtn')?.addEventListener('click', () => toggleColumnModal(true));
    el.get('closeColumnModalBtn')?.addEventListener('click', () => toggleColumnModal(false));
    el.get('applyColumnChangesBtn')?.addEventListener('click', () => { applyColumnChanges(); toggleColumnModal(false); });
    el.get('modalCheckAll')?.addEventListener('click', () => { el.get('columnChecklist').querySelectorAll('input').forEach(i => i.checked = true); });
    el.get('modalUncheckAll')?.addEventListener('click', () => { el.get('columnChecklist').querySelectorAll('input').forEach(i => i.checked = false); });

    // ── 12. 智慧去重 [B3] ──
    el.get('smartDedupBtn')?.addEventListener('click', () => { el.get('dedupColSelect').innerHTML = state.mergedHeaders.map(h=>`<option>${h}</option>`).join(''); el.get('dedupModal').classList.remove('hidden'); });
    el.get('closeDedupModalBtn')?.addEventListener('click', () => el.get('dedupModal').classList.add('hidden'));
    el.get('cancelDedupBtn')?.addEventListener('click', () => el.get('dedupModal').classList.add('hidden'));
    el.get('executeDedupBtn')?.addEventListener('click', executeSmartDeduplication);
    el.get('clearDedupMarksBtn')?.addEventListener('click', clearDedupMarks);
    el.get('deleteDedupMarksBtn')?.addEventListener('click', () => deleteSelectedRows());

    // ── 13. 合併視圖表頭點擊（排序 / 刪欄） ──
    el.get('mergeViewContent')?.addEventListener('click', e => {
      const th = e.target.closest('th:not(.checkbox-cell)');
      const delBtn = e.target.closest('.delete-col-btn');
      if (delBtn && th) { e.stopPropagation(); deleteColumn(delBtn.dataset.header); }
      else if (th) { handleMergedHeaderClick(th); }
    });

    // ── 14. 條件篩選 radio 連動 ──
    el.get('mergeViewModal')?.addEventListener('change', e => {
      if (e.target.name === 'criteria-1' || e.target.name === 'criteria-2') handleCriteriaChange(e);
    });

    // ── 15. 主顯示區委派事件 ──
    el.get('displayArea')?.addEventListener('change', e => {
      if (e.target.matches('.table-select-checkbox,[id^="select-all-cb"], .row-checkbox')) syncCheckboxesInScope();
    });
    el.get('displayArea')?.addEventListener('click', e => {
      const card = e.target.closest('.table-wrapper');
      if (!card) return;

      if (e.target.classList.contains('close-zoom')) { closePreview(); return; }
      if (e.target.classList.contains('delete-rows-btn')) { deleteSelectedRows(card); return; }
      if (e.target.classList.contains('delete-table-btn')) { deleteSelectedTables(card); return; }

      const isGridView = el.get('displayArea').classList.contains('grid-view');
      if (isGridView && !card.classList.contains('is-zoomed') && !e.target.matches('input, a, button, .btn')) {
        openPreview(card);
      }
    });

    // ── 16. Enter 快捷鍵（關鍵字輸入框） ──
    const onKeywordEnter = e => {
      if (e.key !== 'Enter') return;
      e.preventDefault();
      state.isMergedView ? el.get('executeFilterSelectionBtn').click() : el.get('selectByKeywordBtn').click();
    };
    el.get('selectKeywordInput')?.addEventListener('keydown', onKeywordEnter);
    el.get('selectKeywordInputMerged')?.addEventListener('keydown', onKeywordEnter);

    // ── 17. 捲動 / 返回頂端 ──
    el.get('backToTopBtn')?.addEventListener('click', scrollToTop);
    window.addEventListener('scroll', handleScroll);

    // ── 18. Undo 快捷鍵與 ESC 關閉 Modal [A4] ──
    document.addEventListener('keydown', e => {
      if ((e.ctrlKey || e.metaKey) && e.key === 'z') { e.preventDefault(); undoManager.undoLast(); }
      if (e.key === 'Escape') {
        if (!el.get('columnModal').classList.contains('hidden')) { toggleColumnModal(false); }
        else if (!el.get('dedupModal').classList.contains('hidden')) { el.get('dedupModal').classList.add('hidden'); }
        else if (state.isMergedView) { closeMergeView(); }
        else if (state.zoomedCard) { closePreview(); }
      }
    });
    el.get('undoBtn')?.addEventListener('click', () => undoManager.undoLast());
  }

  // ─────────────────────────────────────────────
  // 14. 初始化
  // ─────────────────────────────────────────────

  async function init() {
    try {
      el.init();
      await loadFundConfig();
      bindEvents();
      console.log("✅ ExcelViewer 初始化成功，事件已綁定！");
    } catch (error) {
      console.error("❌ 初始化過程中發生錯誤：", error);
    }
  }

  return { init };
})();

// 確保在各種載入情況下（包含 GitHub Pages 延遲載入）都能正確執行
if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', ExcelViewer.init);
} else {
  ExcelViewer.init();
}
