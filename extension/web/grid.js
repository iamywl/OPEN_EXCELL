// Open Excel - VS Code Webview Grid Controller
(function () {
  // @ts-ignore
  const vscode = acquireVsCodeApi();

  // State
  let currentFileName = 'Workbook.xlsx';
  let activeSheet = 'Sheet1';
  let sheets = {
    Sheet1: createInitialGrid(35, 16)
  };
  let selectedCell = { r: 0, c: 0 };
  let isDirty = false;

  // DOM Elements
  const docTitleEl = document.getElementById('doc-title');
  const btnSave = document.getElementById('btn-save');
  const btnOpenNative = document.getElementById('btn-open-native');
  const cellNameEl = document.getElementById('cell-name');
  const formulaInput = document.getElementById('formula-input');
  const theadEl = document.getElementById('grid-thead');
  const tbodyEl = document.getElementById('grid-tbody');
  const sheetTabsEl = document.getElementById('sheet-tabs');
  const btnAddSheet = document.getElementById('btn-add-sheet');
  const statCount = document.getElementById('stat-count');
  const statSum = document.getElementById('stat-sum');
  const statAvg = document.getElementById('stat-avg');

  // Tool buttons
  const toolBold = document.getElementById('tool-bold');
  const toolItalic = document.getElementById('tool-italic');
  const toolAlignLeft = document.getElementById('tool-align-left');
  const toolAlignCenter = document.getElementById('tool-align-center');
  const toolAlignRight = document.getElementById('tool-align-right');
  const toolAddRow = document.getElementById('tool-add-row');
  const toolAddCol = document.getElementById('tool-add-col');
  const toolSortAsc = document.getElementById('tool-sort-asc');
  const toolSortDesc = document.getElementById('tool-sort-desc');
  const toolExportCsv = document.getElementById('tool-export-csv');

  function createInitialGrid(rows, cols) {
    const grid = [];
    for (let r = 0; r < rows; r++) {
      const row = [];
      for (let c = 0; c < cols; c++) {
        row.push({ raw: '', val: '', bold: false, italic: false, align: 'left' });
      }
      grid.push(row);
    }
    return grid;
  }

  function colIndexToName(index) {
    let name = '';
    let i = index;
    while (i >= 0) {
      name = String.fromCharCode((i % 26) + 65) + name;
      i = Math.floor(i / 26) - 1;
    }
    return name;
  }

  function colNameToIndex(name) {
    let index = 0;
    const str = name.toUpperCase();
    for (let i = 0; i < str.length; i++) {
      index = index * 26 + (str.charCodeAt(i) - 64);
    }
    return index - 1;
  }

  function cellCoordsToName(r, c) {
    return `${colIndexToName(c)}${r + 1}`;
  }

  function parseCellCoords(ref) {
    const match = ref.toUpperCase().match(/^([A-Z]+)(\d+)$/);
    if (!match) return null;
    return {
      c: colNameToIndex(match[1]),
      r: parseInt(match[2], 10) - 1
    };
  }

  // Formula Evaluator
  function evaluateFormula(input, grid) {
    if (!input || !input.startsWith('=')) {
      return input;
    }
    const expr = input.substring(1).trim();

    // Check SUM(range)
    const sumMatch = expr.match(/^SUM\(([A-Z0-9:]+)\)$/i);
    if (sumMatch) {
      const nums = getRangeValues(sumMatch[1], grid);
      return nums.reduce((acc, v) => acc + (parseFloat(v) || 0), 0);
    }

    // Check AVERAGE(range)
    const avgMatch = expr.match(/^AVERAGE\(([A-Z0-9:]+)\)$/i);
    if (avgMatch) {
      const nums = getRangeValues(avgMatch[1], grid);
      if (nums.length === 0) return 0;
      const sum = nums.reduce((acc, v) => acc + (parseFloat(v) || 0), 0);
      return Math.round((sum / nums.length) * 100) / 100;
    }

    // Check COUNT(range)
    const countMatch = expr.match(/^COUNT\(([A-Z0-9:]+)\)$/i);
    if (countMatch) {
      const nums = getRangeValues(countMatch[1], grid);
      return nums.filter(v => v !== '' && !isNaN(Number(v))).length;
    }

    // Direct cell reference like =A1
    const cellCoords = parseCellCoords(expr);
    if (cellCoords && grid[cellCoords.r] && grid[cellCoords.r][cellCoords.c]) {
      return grid[cellCoords.r][cellCoords.c].val;
    }

    return input;
  }

  function getRangeValues(rangeStr, grid) {
    const parts = rangeStr.split(':');
    if (parts.length === 1) {
      const c = parseCellCoords(parts[0]);
      if (c && grid[c.r] && grid[c.r][c.c]) {
        return [grid[c.r][c.c].val];
      }
      return [];
    }
    const start = parseCellCoords(parts[0]);
    const end = parseCellCoords(parts[1]);
    if (!start || !end) return [];

    const minR = Math.min(start.r, end.r);
    const maxR = Math.max(start.r, end.r);
    const minC = Math.min(start.c, end.c);
    const maxC = Math.max(start.c, end.c);

    const values = [];
    for (let r = minR; r <= maxR; r++) {
      for (let c = minC; c <= maxC; c++) {
        if (grid[r] && grid[r][c]) {
          values.push(grid[r][c].val);
        }
      }
    }
    return values;
  }

  function recalculateAll(grid) {
    for (let r = 0; r < grid.length; r++) {
      for (let c = 0; c < grid[r].length; c++) {
        const cell = grid[r][c];
        if (cell.raw.startsWith('=')) {
          cell.val = evaluateFormula(cell.raw, grid);
        } else {
          cell.val = cell.raw;
        }
      }
    }
  }

  // Render Table
  function renderGrid() {
    const grid = sheets[activeSheet];
    if (!grid) return;

    recalculateAll(grid);

    const numRows = grid.length;
    const numCols = grid[0] ? grid[0].length : 0;

    // Build Header
    let theadHtml = '<tr><th class="corner-header"></th>';
    for (let c = 0; c < numCols; c++) {
      theadHtml += `<th class="col-header">${colIndexToName(c)}</th>`;
    }
    theadHtml += '</tr>';
    theadEl.innerHTML = theadHtml;

    // Build Body
    let tbodyHtml = '';
    for (let r = 0; r < numRows; r++) {
      tbodyHtml += `<tr><th class="row-header">${r + 1}</th>`;
      for (let c = 0; c < numCols; c++) {
        const cell = grid[r][c];
        const isSel = (r === selectedCell.r && c === selectedCell.c);
        const style = [
          cell.bold ? 'font-weight:bold;' : '',
          cell.italic ? 'font-style:italic;' : '',
          cell.align ? `text-align:${cell.align};` : ''
        ].join(' ');

        tbodyHtml += `<td class="grid-cell ${isSel ? 'selected' : ''}" data-r="${r}" data-c="${c}" style="${style}">${escapeHtml(String(cell.val || ''))}</td>`;
      }
      tbodyHtml += '</tr>';
    }
    tbodyEl.innerHTML = tbodyHtml;

    updateSelectionUI();
  }

  function updateSelectionUI() {
    const grid = sheets[activeSheet];
    if (!grid) return;

    const cell = (grid[selectedCell.r] && grid[selectedCell.r][selectedCell.c]) || { raw: '', val: '' };
    cellNameEl.textContent = cellCoordsToName(selectedCell.r, selectedCell.c);
    formulaInput.value = cell.raw || '';

    // Update style buttons state
    toolBold.classList.toggle('active', !!cell.bold);
    toolItalic.classList.toggle('active', !!cell.italic);

    // Update stats
    const val = parseFloat(cell.val);
    statCount.textContent = `Count: ${cell.val !== '' ? 1 : 0}`;
    statSum.textContent = `Sum: ${!isNaN(val) ? val : 0}`;
    statAvg.textContent = `Average: ${!isNaN(val) ? val : 0}`;
  }

  function selectCell(r, c) {
    const grid = sheets[activeSheet];
    if (!grid || r < 0 || r >= grid.length || c < 0 || c >= grid[0].length) return;

    // Remove old selected class
    const oldSel = tbodyEl.querySelector('.grid-cell.selected');
    if (oldSel) oldSel.classList.remove('selected');

    selectedCell = { r, c };

    // Add new selected class
    const newSel = tbodyEl.querySelector(`td[data-r="${r}"][data-c="${c}"]`);
    if (newSel) {
      newSel.classList.add('selected');
      newSel.scrollIntoView({ block: 'nearest', inline: 'nearest' });
    }

    updateSelectionUI();
  }

  function setCellValue(r, c, rawValue) {
    const grid = sheets[activeSheet];
    if (!grid || !grid[r] || !grid[r][c]) return;

    grid[r][c].raw = rawValue;
    markDirty(true);
    renderGrid();
  }

  function markDirty(dirty) {
    isDirty = dirty;
    vscode.postMessage({ type: 'gridcraft:dirty', dirty });
  }

  // Parse CSV / TSV text
  function parseDelimited(text, delimiter) {
    const lines = text.split(/\r?\n/).filter(line => line.trim().length > 0);
    const data = [];
    for (const line of lines) {
      const row = line.split(delimiter).map(field => {
        // Strip quotes if present
        let val = field.trim();
        if (val.startsWith('"') && val.endsWith('"')) {
          val = val.substring(1, val.length - 1).replace(/""/g, '"');
        }
        return { raw: val, val: val, bold: false, italic: false, align: 'left' };
      });
      data.push(row);
    }
    // Pad to minimum 30 rows and 12 cols
    const maxCols = Math.max(12, ...data.map(r => r.length));
    const maxRows = Math.max(30, data.length);
    for (let r = 0; r < maxRows; r++) {
      if (!data[r]) data[r] = [];
      while (data[r].length < maxCols) {
        data[r].push({ raw: '', val: '', bold: false, italic: false, align: 'left' });
      }
    }
    return data;
  }

  // Convert grid to CSV text
  function gridToCsv(grid) {
    return grid.map(row => {
      return row.map(cell => {
        const text = String(cell.val || '');
        if (text.includes(',') || text.includes('"') || text.includes('\n')) {
          return `"${text.replace(/"/g, '""')}"`;
        }
        return text;
      }).join(',');
    }).join('\n');
  }

  // Parse binary file data
  function loadFileBytes(filename, bytes) {
    currentFileName = filename;
    docTitleEl.textContent = filename;

    const lower = filename.toLowerCase();
    if (lower.endsWith('.csv')) {
      const text = new TextDecoder('utf-8').decode(new Uint8Array(bytes));
      sheets['Sheet1'] = parseDelimited(text, ',');
    } else if (lower.endsWith('.tsv')) {
      const text = new TextDecoder('utf-8').decode(new Uint8Array(bytes));
      sheets['Sheet1'] = parseDelimited(text, '\t');
    } else {
      // For XLSX or other binary formats:
      // Try text fallback first; if binary, initialize clean grid with filename
      try {
        const text = new TextDecoder('utf-8').decode(new Uint8Array(bytes));
        if (text.includes(',') && !text.includes('\x00')) {
          sheets['Sheet1'] = parseDelimited(text, ',');
        } else {
          // Initialize sample structured data if file was empty
          if (bytes.length === 0) {
            sheets['Sheet1'] = createInitialGrid(35, 16);
            // Default sample headers
            sheets['Sheet1'][0][0] = { raw: 'Product', val: 'Product', bold: true, italic: false, align: 'left' };
            sheets['Sheet1'][0][1] = { raw: 'Q1 Sales', val: 'Q1 Sales', bold: true, italic: false, align: 'right' };
            sheets['Sheet1'][0][2] = { raw: 'Q2 Sales', val: 'Q2 Sales', bold: true, italic: false, align: 'right' };
            sheets['Sheet1'][0][3] = { raw: 'Total', val: 'Total', bold: true, italic: false, align: 'right' };

            sheets['Sheet1'][1][0] = { raw: 'Widgets', val: 'Widgets', bold: false, italic: false, align: 'left' };
            sheets['Sheet1'][1][1] = { raw: '1200', val: 1200, bold: false, italic: false, align: 'right' };
            sheets['Sheet1'][1][2] = { raw: '1450', val: 1450, bold: false, italic: false, align: 'right' };
            sheets['Sheet1'][1][3] = { raw: '=SUM(B2:C2)', val: 2650, bold: false, italic: false, align: 'right' };

            sheets['Sheet1'][2][0] = { raw: 'Gadgets', val: 'Gadgets', bold: false, italic: false, align: 'left' };
            sheets['Sheet1'][2][1] = { raw: '850', val: 850, bold: false, italic: false, align: 'right' };
            sheets['Sheet1'][2][2] = { raw: '920', val: 920, bold: false, italic: false, align: 'right' };
            sheets['Sheet1'][2][3] = { raw: '=SUM(B3:C3)', val: 1770, bold: false, italic: false, align: 'right' };
          } else {
            // Binary loaded: initialize grid
            sheets['Sheet1'] = createInitialGrid(35, 16);
          }
        }
      } catch (e) {
        sheets['Sheet1'] = createInitialGrid(35, 16);
      }
    }

    renderGrid();
  }

  function saveCurrentWorkbook() {
    const grid = sheets[activeSheet];
    const csvContent = gridToCsv(grid);
    const bytes = Array.from(new TextEncoder().encode(csvContent));
    vscode.postMessage({
      type: 'gridcraft:save',
      bytes: bytes
    });
    markDirty(false);
  }

  // Event Listeners
  btnSave.addEventListener('click', saveCurrentWorkbook);

  btnOpenNative.addEventListener('click', () => {
    vscode.postMessage({ type: 'gridcraft:openNative' });
  });

  formulaInput.addEventListener('keydown', (e) => {
    if (e.key === 'Enter') {
      setCellValue(selectedCell.r, selectedCell.c, formulaInput.value);
      selectCell(selectedCell.r + 1, selectedCell.c);
    } else if (e.key === 'Escape') {
      updateSelectionUI();
    }
  });

  // Table Cell Interaction
  tbodyEl.addEventListener('click', (e) => {
    const cell = e.target.closest('.grid-cell');
    if (!cell) return;
    const r = parseInt(cell.getAttribute('data-r'), 10);
    const c = parseInt(cell.getAttribute('data-c'), 10);
    selectCell(r, c);
  });

  tbodyEl.addEventListener('dblclick', (e) => {
    const cell = e.target.closest('.grid-cell');
    if (!cell) return;
    const r = parseInt(cell.getAttribute('data-r'), 10);
    const c = parseInt(cell.getAttribute('data-c'), 10);
    selectCell(r, c);

    // Prompt inline or focus formula bar
    formulaInput.focus();
    formulaInput.select();
  });

  // Toolbar Actions
  toolBold.addEventListener('click', () => {
    const grid = sheets[activeSheet];
    const cell = grid[selectedCell.r][selectedCell.c];
    cell.bold = !cell.bold;
    markDirty(true);
    renderGrid();
  });

  toolItalic.addEventListener('click', () => {
    const grid = sheets[activeSheet];
    const cell = grid[selectedCell.r][selectedCell.c];
    cell.italic = !cell.italic;
    markDirty(true);
    renderGrid();
  });

  toolAlignLeft.addEventListener('click', () => {
    const grid = sheets[activeSheet];
    grid[selectedCell.r][selectedCell.c].align = 'left';
    markDirty(true);
    renderGrid();
  });

  toolAlignCenter.addEventListener('click', () => {
    const grid = sheets[activeSheet];
    grid[selectedCell.r][selectedCell.c].align = 'center';
    markDirty(true);
    renderGrid();
  });

  toolAlignRight.addEventListener('click', () => {
    const grid = sheets[activeSheet];
    grid[selectedCell.r][selectedCell.c].align = 'right';
    markDirty(true);
    renderGrid();
  });

  toolAddRow.addEventListener('click', () => {
    const grid = sheets[activeSheet];
    const newRow = [];
    const cols = grid[0] ? grid[0].length : 10;
    for (let c = 0; c < cols; c++) {
      newRow.push({ raw: '', val: '', bold: false, italic: false, align: 'left' });
    }
    grid.splice(selectedCell.r + 1, 0, newRow);
    markDirty(true);
    renderGrid();
  });

  toolAddCol.addEventListener('click', () => {
    const grid = sheets[activeSheet];
    for (let r = 0; r < grid.length; r++) {
      grid[r].splice(selectedCell.c + 1, 0, { raw: '', val: '', bold: false, italic: false, align: 'left' });
    }
    markDirty(true);
    renderGrid();
  });

  toolSortAsc.addEventListener('click', () => {
    const grid = sheets[activeSheet];
    const col = selectedCell.c;
    // Keep header row intact
    const header = grid[0];
    const rows = grid.slice(1);
    rows.sort((a, b) => {
      const valA = a[col] ? a[col].val : '';
      const valB = b[col] ? b[col].val : '';
      return String(valA).localeCompare(String(valB), 'ko-KR', { numeric: true });
    });
    sheets[activeSheet] = [header, ...rows];
    markDirty(true);
    renderGrid();
  });

  toolSortDesc.addEventListener('click', () => {
    const grid = sheets[activeSheet];
    const col = selectedCell.c;
    const header = grid[0];
    const rows = grid.slice(1);
    rows.sort((a, b) => {
      const valA = a[col] ? a[col].val : '';
      const valB = b[col] ? b[col].val : '';
      return String(valB).localeCompare(String(valA), 'ko-KR', { numeric: true });
    });
    sheets[activeSheet] = [header, ...rows];
    markDirty(true);
    renderGrid();
  });

  toolExportCsv.addEventListener('click', () => {
    const grid = sheets[activeSheet];
    const csv = gridToCsv(grid);
    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = `${activeSheet}.csv`;
    link.click();
  });

  // Keyboard Navigation
  window.addEventListener('keydown', (e) => {
    if (e.target === formulaInput) return;

    if ((e.ctrlKey || e.metaKey) && e.key === 's') {
      e.preventDefault();
      saveCurrentWorkbook();
      return;
    }

    if (e.key === 'ArrowUp') {
      e.preventDefault();
      selectCell(selectedCell.r - 1, selectedCell.c);
    } else if (e.key === 'ArrowDown' || e.key === 'Enter') {
      e.preventDefault();
      selectCell(selectedCell.r + 1, selectedCell.c);
    } else if (e.key === 'ArrowLeft') {
      e.preventDefault();
      selectCell(selectedCell.r, selectedCell.c - 1);
    } else if (e.key === 'ArrowRight' || e.key === 'Tab') {
      e.preventDefault();
      selectCell(selectedCell.r, selectedCell.c + 1);
    } else if (e.key === 'Delete' || e.key === 'Backspace') {
      setCellValue(selectedCell.r, selectedCell.c, '');
    } else if (e.key.length === 1 && !e.ctrlKey && !e.metaKey && !e.altKey) {
      formulaInput.value = e.key;
      formulaInput.focus();
    }
  });

  // Sheet Add Button
  btnAddSheet.addEventListener('click', () => {
    const name = `Sheet${Object.keys(sheets).length + 1}`;
    sheets[name] = createInitialGrid(35, 16);
    const tab = document.createElement('button');
    tab.className = 'sheet-tab';
    tab.textContent = name;
    tab.setAttribute('data-sheet', name);
    tab.addEventListener('click', () => switchSheet(name));
    sheetTabsEl.insertBefore(tab, btnAddSheet);
    switchSheet(name);
  });

  function switchSheet(name) {
    activeSheet = name;
    document.querySelectorAll('.sheet-tab').forEach(t => {
      t.classList.toggle('active', t.getAttribute('data-sheet') === name);
    });
    renderGrid();
  }

  function escapeHtml(text) {
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
  }

  // Listen to messages from VS Code Host
  window.addEventListener('message', (event) => {
    const message = event.data;
    switch (message.type) {
      case 'gridcraft:open':
        loadFileBytes(message.filename, message.bytes);
        break;
    }
  });

  // Signal Ready to VS Code Host
  vscode.postMessage({ type: 'gridcraft:ready' });
})();
