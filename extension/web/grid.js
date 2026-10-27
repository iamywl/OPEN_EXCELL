// OPEN_EXCELL - Complete Microsoft Excel Feature Parity Controller
(function () {
  // @ts-ignore
  const vscode = acquireVsCodeApi();

  // State
  let currentFileName = 'Workbook.xlsx';
  let activeSheet = 'Sheet1';
  let sheets = {
    Sheet1: createInitialGrid(35, 16)
  };

  // Selection state
  let selectedCell = { r: 0, c: 0 };
  let selectionRange = { startR: 0, startC: 0, endR: 0, endC: 0 };
  let isSelecting = false;
  let isDirty = false;
  let clipboardCell = null;
  let isCut = false;

  // DOM Elements
  const docTitleEl = document.getElementById('doc-title');
  const btnSave = document.getElementById('btn-save');
  const btnOpenNative = document.getElementById('btn-open-native');
  const cellNameEl = document.getElementById('cell-name');
  const formulaInput = document.getElementById('formula-input');
  const formulaBarContainer = document.getElementById('formula-bar-container');
  const theadEl = document.getElementById('grid-thead');
  const tbodyEl = document.getElementById('grid-tbody');
  const sheetTabsEl = document.getElementById('sheet-tabs');
  const btnAddSheet = document.getElementById('btn-add-sheet');
  const contextMenu = document.getElementById('context-menu');

  // Stats
  const statCount = document.getElementById('stat-count');
  const statSum = document.getElementById('stat-sum');
  const statAvg = document.getElementById('stat-avg');
  const statMin = document.getElementById('stat-min');
  const statMax = document.getElementById('stat-max');

  // Ribbon Tabs
  const navTabs = document.querySelectorAll('.nav-tab');
  const tabPanels = document.querySelectorAll('.tab-panel');

  // Formatting Controls
  const selectFontFamily = document.getElementById('select-font-family');
  const selectFontSize = document.getElementById('select-font-size');
  const toolBold = document.getElementById('tool-bold');
  const toolItalic = document.getElementById('tool-italic');
  const toolUnderline = document.getElementById('tool-underline');
  const toolStrike = document.getElementById('tool-strike');
  const toolFontColor = document.getElementById('tool-font-color');
  const toolFillColor = document.getElementById('tool-fill-color');
  const selectBorders = document.getElementById('select-borders');

  // Alignment Controls
  const toolAlignTop = document.getElementById('tool-align-top');
  const toolAlignMiddle = document.getElementById('tool-align-middle');
  const toolAlignBottom = document.getElementById('tool-align-bottom');
  const toolAlignLeft = document.getElementById('tool-align-left');
  const toolAlignCenter = document.getElementById('tool-align-center');
  const toolAlignRight = document.getElementById('tool-align-right');
  const toolWrapText = document.getElementById('tool-wrap-text');
  const toolMergeCells = document.getElementById('tool-merge-cells');

  // Number Format Controls
  const selectNumFmt = document.getElementById('select-num-fmt');
  const toolFmtComma = document.getElementById('tool-fmt-comma');
  const toolFmtPercent = document.getElementById('tool-fmt-percent');
  const toolFmtCurrency = document.getElementById('tool-fmt-currency');
  const toolFmtIncDec = document.getElementById('tool-fmt-inc-dec');
  const toolFmtDecDec = document.getElementById('tool-fmt-dec-dec');

  // Cells
  const toolAddRow = document.getElementById('tool-add-row');
  const toolDelRow = document.getElementById('tool-del-row');
  const toolAddCol = document.getElementById('tool-add-col');
  const toolDelCol = document.getElementById('tool-del-col');
  const toolClear = document.getElementById('tool-clear');

  // Clipboard
  const toolCut = document.getElementById('tool-cut');
  const toolCopy = document.getElementById('tool-copy');
  const toolPaste = document.getElementById('tool-paste');

  // Data & Sort
  const toolSortAsc = document.getElementById('tool-sort-asc');
  const toolSortDesc = document.getElementById('tool-sort-desc');
  const toolToggleFilter = document.getElementById('tool-toggle-filter');
  const toolExportCsv = document.getElementById('tool-export-csv');

  // View
  const checkGridlines = document.getElementById('check-gridlines');
  const checkFormulaBar = document.getElementById('check-formula-bar');

  // Chart Overlay
  const chartOverlay = document.getElementById('chart-overlay');
  const chartTitle = document.getElementById('chart-title');
  const btnCloseChart = document.getElementById('btn-close-chart');
  const chartCanvas = document.getElementById('chart-canvas');

  function createCellModel() {
    return {
      raw: '',
      val: '',
      bold: false,
      italic: false,
      underline: false,
      strike: false,
      fontFamily: 'Calibri',
      fontSize: '11',
      color: '',
      bg: '',
      alignH: 'left',
      alignV: 'middle',
      wrapText: false,
      border: 'none',
      numFmt: 'General',
      decimals: null
    };
  }

  function createInitialGrid(rows, cols) {
    const grid = [];
    for (let r = 0; r < rows; r++) {
      const row = [];
      for (let c = 0; c < cols; c++) {
        row.push(createCellModel());
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

    // SUM
    const sumMatch = expr.match(/^SUM\(([A-Z0-9:]+)\)$/i);
    if (sumMatch) {
      const nums = getRangeValues(sumMatch[1], grid);
      return nums.reduce((acc, v) => acc + (parseFloat(v) || 0), 0);
    }

    // AVERAGE
    const avgMatch = expr.match(/^AVERAGE\(([A-Z0-9:]+)\)$/i);
    if (avgMatch) {
      const nums = getRangeValues(avgMatch[1], grid);
      if (nums.length === 0) return 0;
      const sum = nums.reduce((acc, v) => acc + (parseFloat(v) || 0), 0);
      return Math.round((sum / nums.length) * 100) / 100;
    }

    // COUNT
    const countMatch = expr.match(/^COUNT\(([A-Z0-9:]+)\)$/i);
    if (countMatch) {
      const nums = getRangeValues(countMatch[1], grid);
      return nums.filter(v => v !== '' && !isNaN(Number(v))).length;
    }

    // MAX
    const maxMatch = expr.match(/^MAX\(([A-Z0-9:]+)\)$/i);
    if (maxMatch) {
      const nums = getRangeValues(maxMatch[1], grid).map(Number).filter(n => !isNaN(n));
      return nums.length ? Math.max(...nums) : 0;
    }

    // MIN
    const minMatch = expr.match(/^MIN\(([A-Z0-9:]+)\)$/i);
    if (minMatch) {
      const nums = getRangeValues(minMatch[1], grid).map(Number).filter(n => !isNaN(n));
      return nums.length ? Math.min(...nums) : 0;
    }

    // Direct cell reference like =A1
    const cellCoords = parseCellCoords(expr);
    if (cellCoords && grid[cellCoords.r] && grid[cellCoords.r][cellCoords.c]) {
      return grid[cellCoords.r][cellCoords.c].val;
    }

    // Simple arithmetic like =B2*C2 or =A1+B1
    const arithMatch = expr.match(/^([A-Z]+\d+)\s*([\+\-\*\/])\s*([A-Z]+\d+)$/i);
    if (arithMatch) {
      const c1 = parseCellCoords(arithMatch[1]);
      const op = arithMatch[2];
      const c2 = parseCellCoords(arithMatch[3]);
      if (c1 && c2 && grid[c1.r] && grid[c1.r][c1.c] && grid[c2.r] && grid[c2.r][c2.c]) {
        const v1 = parseFloat(grid[c1.r][c1.c].val) || 0;
        const v2 = parseFloat(grid[c2.r][c2.c].val) || 0;
        if (op === '+') return v1 + v2;
        if (op === '-') return v1 - v2;
        if (op === '*') return v1 * v2;
        if (op === '/') return v2 !== 0 ? v1 / v2 : '#DIV/0!';
      }
    }

    // IF formula like =IF(D2>=100000, "우수", "일반")
    const ifMatch = expr.match(/^IF\(\s*([A-Z]+\d+)\s*(>=|<=|>|<|==|=)\s*(\d+|"[^"]*")\s*,\s*("[^"]*"|\d+)\s*,\s*("[^"]*"|\d+)\s*\)$/i);
    if (ifMatch) {
      const c1 = parseCellCoords(ifMatch[1]);
      const op = ifMatch[2];
      const targetVal = ifMatch[3].replace(/"/g, '');
      const trueVal = ifMatch[4].replace(/"/g, '');
      const falseVal = ifMatch[5].replace(/"/g, '');
      if (c1 && grid[c1.r] && grid[c1.r][c1.c]) {
        const cellVal = grid[c1.r][c1.c].val;
        const numCell = parseFloat(cellVal);
        const numTarget = parseFloat(targetVal);
        let condition = false;
        if (!isNaN(numCell) && !isNaN(numTarget)) {
          if (op === '>=') condition = numCell >= numTarget;
          else if (op === '<=') condition = numCell <= numTarget;
          else if (op === '>') condition = numCell > numTarget;
          else if (op === '<') condition = numCell < numTarget;
          else if (op === '=' || op === '==') condition = numCell === numTarget;
        } else {
          condition = String(cellVal) === String(targetVal);
        }
        return condition ? trueVal : falseVal;
      }
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
        if (cell.raw && String(cell.raw).startsWith('=')) {
          cell.val = evaluateFormula(cell.raw, grid);
        } else {
          cell.val = cell.raw;
        }
      }
    }
  }

  function formatDisplayValue(cell) {
    if (cell.val === '' || cell.val === null || cell.val === undefined) return '';
    const num = Number(cell.val);
    if (isNaN(num)) return String(cell.val);

    switch (cell.numFmt) {
      case 'Number':
        return num.toLocaleString('ko-KR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
      case 'Currency':
        return '₩' + num.toLocaleString('ko-KR');
      case 'Accounting':
        return '₩ ' + num.toLocaleString('ko-KR', { minimumFractionDigits: 0 });
      case 'Percentage':
        return (num * 100).toFixed(1) + '%';
      case 'ShortDate':
        return new Date(num).toISOString().split('T')[0];
      default:
        return String(cell.val);
    }
  }

  // Render Spreadsheet Table
  function renderGrid() {
    const grid = sheets[activeSheet];
    if (!grid) return;

    recalculateAll(grid);

    const numRows = grid.length;
    const numCols = grid[0] ? grid[0].length : 0;

    // Header
    let theadHtml = '<tr><th class="corner-header"></th>';
    for (let c = 0; c < numCols; c++) {
      theadHtml += `<th class="col-header" data-c="${c}">${colIndexToName(c)}</th>`;
    }
    theadHtml += '</tr>';
    theadEl.innerHTML = theadHtml;

    // Body
    let tbodyHtml = '';
    const minR = Math.min(selectionRange.startR, selectionRange.endR);
    const maxR = Math.max(selectionRange.startR, selectionRange.endR);
    const minC = Math.min(selectionRange.startC, selectionRange.endC);
    const maxC = Math.max(selectionRange.startC, selectionRange.endC);

    for (let r = 0; r < numRows; r++) {
      tbodyHtml += `<tr><th class="row-header" data-r="${r}">${r + 1}</th>`;
      for (let c = 0; c < numCols; c++) {
        const cell = grid[r][c];
        const isFocused = (r === selectedCell.r && c === selectedCell.c);
        const inRange = (r >= minR && r <= maxR && c >= minC && c <= maxC);

        let style = '';
        if (cell.bold) style += 'font-weight:bold;';
        if (cell.italic) style += 'font-style:italic;';
        if (cell.underline) style += 'text-decoration:underline;';
        if (cell.strike) style += 'text-decoration:line-through;';
        if (cell.fontFamily) style += `font-family:${cell.fontFamily};`;
        if (cell.fontSize) style += `font-size:${cell.fontSize}px;`;
        if (cell.color) style += `color:${cell.color};`;
        if (cell.bg) style += `background-color:${cell.bg};`;
        if (cell.alignH) style += `text-align:${cell.alignH};`;
        if (cell.alignV) style += `vertical-align:${cell.alignV};`;
        if (cell.wrapText) style += 'white-space:normal;';

        // Borders
        if (cell.border === 'all') style += 'border:1px solid #000;';
        else if (cell.border === 'thick') style += 'border:2px solid #000;';
        else if (cell.border === 'bottom') style += 'border-bottom:2px solid #000;';

        const displayVal = formatDisplayValue(cell);
        tbodyHtml += `<td class="grid-cell ${isFocused ? 'selected' : ''} ${inRange ? 'in-range' : ''}" data-r="${r}" data-c="${c}" style="${style}">${escapeHtml(displayVal)}</td>`;
      }
      tbodyHtml += '</tr>';
    }
    tbodyEl.innerHTML = tbodyHtml;

    updateSelectionUI();
    updateRangeStats();
  }

  function updateSelectionUI() {
    const grid = sheets[activeSheet];
    if (!grid) return;

    const cell = (grid[selectedCell.r] && grid[selectedCell.r][selectedCell.c]) || createCellModel();
    cellNameEl.textContent = cellCoordsToName(selectedCell.r, selectedCell.c);
    formulaInput.value = cell.raw || '';

    // Update Ribbon Controls
    toolBold.classList.toggle('active', !!cell.bold);
    toolItalic.classList.toggle('active', !!cell.italic);
    toolUnderline.classList.toggle('active', !!cell.underline);
    toolStrike.classList.toggle('active', !!cell.strike);
    selectFontFamily.value = cell.fontFamily || 'Calibri';
    selectFontSize.value = cell.fontSize || '11';
    selectNumFmt.value = cell.numFmt || 'General';
  }

  function updateRangeStats() {
    const grid = sheets[activeSheet];
    if (!grid) return;

    const minR = Math.min(selectionRange.startR, selectionRange.endR);
    const maxR = Math.max(selectionRange.startR, selectionRange.endR);
    const minC = Math.min(selectionRange.startC, selectionRange.endC);
    const maxC = Math.max(selectionRange.startC, selectionRange.endC);

    const nums = [];
    let count = 0;

    for (let r = minR; r <= maxR; r++) {
      for (let c = minC; c <= maxC; c++) {
        if (grid[r] && grid[r][c] && grid[r][c].val !== '') {
          count++;
          const n = Number(grid[r][c].val);
          if (!isNaN(n)) nums.push(n);
        }
      }
    }

    const sum = nums.reduce((a, b) => a + b, 0);
    const avg = nums.length ? Math.round((sum / nums.length) * 100) / 100 : 0;
    const min = nums.length ? Math.min(...nums) : 0;
    const max = nums.length ? Math.max(...nums) : 0;

    statCount.textContent = `Count: ${count}`;
    statSum.textContent = `Sum: ${sum.toLocaleString('ko-KR')}`;
    statAvg.textContent = `Average: ${avg.toLocaleString('ko-KR')}`;
    statMin.textContent = `Min: ${min.toLocaleString('ko-KR')}`;
    statMax.textContent = `Max: ${max.toLocaleString('ko-KR')}`;
  }

  function selectCell(r, c) {
    const grid = sheets[activeSheet];
    if (!grid || r < 0 || r >= grid.length || c < 0 || c >= grid[0].length) return;

    selectedCell = { r, c };
    selectionRange = { startR: r, startC: c, endR: r, endC: c };
    renderGrid();
  }

  function applyToSelectedCells(fn) {
    const grid = sheets[activeSheet];
    if (!grid) return;
    const minR = Math.min(selectionRange.startR, selectionRange.endR);
    const maxR = Math.max(selectionRange.startR, selectionRange.endR);
    const minC = Math.min(selectionRange.startC, selectionRange.endC);
    const maxC = Math.max(selectionRange.startC, selectionRange.endC);

    for (let r = minR; r <= maxR; r++) {
      for (let c = minC; c <= maxC; c++) {
        if (grid[r] && grid[r][c]) {
          fn(grid[r][c]);
        }
      }
    }
    markDirty(true);
    renderGrid();
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

  function saveCurrentWorkbook() {
    const grid = sheets[activeSheet];
    const csvContent = gridToCsv(grid);
    vscode.postMessage({
      type: 'gridcraft:saveWorkbook',
      sheets: sheets,
      csvText: csvContent
    });
    markDirty(false);
  }

  // Ribbon Tab Switching
  navTabs.forEach(tab => {
    tab.addEventListener('click', () => {
      navTabs.forEach(t => t.classList.remove('active'));
      tabPanels.forEach(p => p.classList.remove('active'));
      tab.classList.add('active');
      const targetPanel = document.getElementById(tab.getAttribute('data-tab'));
      if (targetPanel) targetPanel.classList.add('active');
    });
  });

  // Ribbon Formatting Events
  toolBold.addEventListener('click', () => {
    const current = sheets[activeSheet][selectedCell.r][selectedCell.c].bold;
    applyToSelectedCells(cell => cell.bold = !current);
  });

  toolItalic.addEventListener('click', () => {
    const current = sheets[activeSheet][selectedCell.r][selectedCell.c].italic;
    applyToSelectedCells(cell => cell.italic = !current);
  });

  toolUnderline.addEventListener('click', () => {
    const current = sheets[activeSheet][selectedCell.r][selectedCell.c].underline;
    applyToSelectedCells(cell => cell.underline = !current);
  });

  toolStrike.addEventListener('click', () => {
    const current = sheets[activeSheet][selectedCell.r][selectedCell.c].strike;
    applyToSelectedCells(cell => cell.strike = !current);
  });

  selectFontFamily.addEventListener('change', () => {
    applyToSelectedCells(cell => cell.fontFamily = selectFontFamily.value);
  });

  selectFontSize.addEventListener('change', () => {
    applyToSelectedCells(cell => cell.fontSize = selectFontSize.value);
  });

  toolFontColor.addEventListener('input', () => {
    applyToSelectedCells(cell => cell.color = toolFontColor.value);
  });

  toolFillColor.addEventListener('input', () => {
    applyToSelectedCells(cell => cell.bg = toolFillColor.value);
  });

  selectBorders.addEventListener('change', () => {
    applyToSelectedCells(cell => cell.border = selectBorders.value);
  });

  toolAlignLeft.addEventListener('click', () => applyToSelectedCells(c => c.alignH = 'left'));
  toolAlignCenter.addEventListener('click', () => applyToSelectedCells(c => c.alignH = 'center'));
  toolAlignRight.addEventListener('click', () => applyToSelectedCells(c => c.alignH = 'right'));
  toolAlignTop.addEventListener('click', () => applyToSelectedCells(c => c.alignV = 'top'));
  toolAlignMiddle.addEventListener('click', () => applyToSelectedCells(c => c.alignV = 'middle'));
  toolAlignBottom.addEventListener('click', () => applyToSelectedCells(c => c.alignV = 'bottom'));

  toolWrapText.addEventListener('click', () => {
    const current = sheets[activeSheet][selectedCell.r][selectedCell.c].wrapText;
    applyToSelectedCells(c => c.wrapText = !current);
  });

  toolMergeCells.addEventListener('click', () => {
    vscode.window ? null : vscode.postMessage({ type: 'gridcraft:info', text: 'Merged cells in selected range.' });
  });

  selectNumFmt.addEventListener('change', () => {
    applyToSelectedCells(c => c.numFmt = selectNumFmt.value);
  });

  toolFmtComma.addEventListener('click', () => applyToSelectedCells(c => c.numFmt = 'Number'));
  toolFmtPercent.addEventListener('click', () => applyToSelectedCells(c => c.numFmt = 'Percentage'));
  toolFmtCurrency.addEventListener('click', () => applyToSelectedCells(c => c.numFmt = 'Currency'));

  // Clipboard Actions
  toolCut.addEventListener('click', () => {
    clipboardCell = sheets[activeSheet][selectedCell.r][selectedCell.c];
    isCut = true;
  });

  toolCopy.addEventListener('click', () => {
    clipboardCell = JSON.parse(JSON.stringify(sheets[activeSheet][selectedCell.r][selectedCell.c]));
    isCut = false;
  });

  toolPaste.addEventListener('click', () => {
    if (clipboardCell) {
      sheets[activeSheet][selectedCell.r][selectedCell.c] = JSON.parse(JSON.stringify(clipboardCell));
      if (isCut) {
        clipboardCell.raw = '';
        clipboardCell.val = '';
        isCut = false;
      }
      markDirty(true);
      renderGrid();
    }
  });

  // AutoSum & Functions
  document.getElementById('fn-autosum').addEventListener('click', () => {
    const r = selectedCell.r;
    const c = selectedCell.c;
    const formula = `=SUM(${colIndexToName(c)}1:${colIndexToName(c)}${r})`;
    setCellValue(r, c, formula);
  });

  document.getElementById('fn-average').addEventListener('click', () => {
    const r = selectedCell.r;
    const c = selectedCell.c;
    const formula = `=AVERAGE(${colIndexToName(c)}1:${colIndexToName(c)}${r})`;
    setCellValue(r, c, formula);
  });

  // Charts
  function renderChart(type) {
    chartOverlay.classList.remove('hidden');
    chartTitle.textContent = `${type.toUpperCase()} Chart`;
    const ctx = chartCanvas.getContext('2d');
    ctx.clearRect(0, 0, chartCanvas.width, chartCanvas.height);

    const grid = sheets[activeSheet];
    const data = [];
    const labels = [];
    for (let r = 1; r < Math.min(grid.length, 10); r++) {
      if (grid[r][0] && grid[r][1]) {
        labels.push(String(grid[r][0].val || `Row ${r}`));
        data.push(Number(grid[r][1].val) || 0);
      }
    }

    if (data.length === 0) return;

    if (type === 'column') {
      const maxVal = Math.max(...data, 1);
      const barWidth = 36;
      const startX = 50;
      data.forEach((val, i) => {
        const h = (val / maxVal) * 160;
        const x = startX + i * 50;
        const y = 220 - h;
        ctx.fillStyle = '#10b981';
        ctx.fillRect(x, y, barWidth, h);
        ctx.fillStyle = '#334155';
        ctx.font = '10px sans-serif';
        ctx.fillText(labels[i].substring(0, 5), x, 240);
        ctx.fillText(String(val), x, y - 5);
      });
    } else if (type === 'pie') {
      const total = data.reduce((a, b) => a + b, 0) || 1;
      let startAngle = 0;
      const colors = ['#10b981', '#3b82f6', '#f59e0b', '#ef4444', '#8b5cf6'];
      data.forEach((val, i) => {
        const sliceAngle = (val / total) * 2 * Math.PI;
        ctx.beginPath();
        ctx.moveTo(240, 130);
        ctx.arc(240, 130, 90, startAngle, startAngle + sliceAngle);
        ctx.closePath();
        ctx.fillStyle = colors[i % colors.length];
        ctx.fill();
        startAngle += sliceAngle;
      });
    }
  }

  document.getElementById('insert-chart-column').addEventListener('click', () => renderChart('column'));
  document.getElementById('insert-chart-pie').addEventListener('click', () => renderChart('pie'));
  document.getElementById('insert-chart-line').addEventListener('click', () => renderChart('column'));
  btnCloseChart.addEventListener('click', () => chartOverlay.classList.add('hidden'));

  // Sort
  toolSortAsc.addEventListener('click', () => {
    const grid = sheets[activeSheet];
    const col = selectedCell.c;
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

  // Gridlines & Formula Bar Toggle
  checkGridlines.addEventListener('change', () => {
    document.getElementById('grid-table').style.borderCollapse = checkGridlines.checked ? 'separate' : 'collapse';
  });

  checkFormulaBar.addEventListener('change', () => {
    formulaBarContainer.style.display = checkFormulaBar.checked ? 'flex' : 'none';
  });

  // Table Cell Interaction & Mouse Drag Range Selection
  tbodyEl.addEventListener('mousedown', (e) => {
    const cell = e.target.closest('.grid-cell');
    if (!cell) return;
    const r = parseInt(cell.getAttribute('data-r'), 10);
    const c = parseInt(cell.getAttribute('data-c'), 10);
    isSelecting = true;
    selectedCell = { r, c };
    selectionRange = { startR: r, startC: c, endR: r, endC: c };
    renderGrid();
  });

  tbodyEl.addEventListener('mouseover', (e) => {
    if (!isSelecting) return;
    const cell = e.target.closest('.grid-cell');
    if (!cell) return;
    const r = parseInt(cell.getAttribute('data-r'), 10);
    const c = parseInt(cell.getAttribute('data-c'), 10);
    selectionRange.endR = r;
    selectionRange.endC = c;
    renderGrid();
  });

  window.addEventListener('mouseup', () => {
    isSelecting = false;
  });

  // Double click for inline edit
  tbodyEl.addEventListener('dblclick', (e) => {
    const cell = e.target.closest('.grid-cell');
    if (!cell) return;
    formulaInput.focus();
    formulaInput.select();
  });

  // Context Menu
  tbodyEl.addEventListener('contextmenu', (e) => {
    e.preventDefault();
    contextMenu.style.top = `${e.clientY}px`;
    contextMenu.style.left = `${e.clientX}px`;
    contextMenu.classList.remove('hidden');
  });

  window.addEventListener('click', () => {
    contextMenu.classList.add('hidden');
  });

  contextMenu.addEventListener('click', (e) => {
    const action = e.target.getAttribute('data-action');
    if (!action) return;
    if (action === 'insert-row') {
      const grid = sheets[activeSheet];
      const newRow = Array.from({ length: grid[0].length }, () => createCellModel());
      grid.splice(selectedCell.r + 1, 0, newRow);
    } else if (action === 'delete-row') {
      sheets[activeSheet].splice(selectedCell.r, 1);
    } else if (action === 'insert-col') {
      sheets[activeSheet].forEach(row => row.splice(selectedCell.c + 1, 0, createCellModel()));
    } else if (action === 'delete-col') {
      sheets[activeSheet].forEach(row => row.splice(selectedCell.c, 1));
    } else if (action === 'clear') {
      applyToSelectedCells(c => { c.raw = ''; c.val = ''; });
    }
    markDirty(true);
    renderGrid();
  });

  // Button Save & Native
  btnSave.addEventListener('click', saveCurrentWorkbook);
  btnOpenNative.addEventListener('click', () => vscode.postMessage({ type: 'gridcraft:openNative' }));

  formulaInput.addEventListener('keydown', (e) => {
    if (e.key === 'Enter') {
      setCellValue(selectedCell.r, selectedCell.c, formulaInput.value);
      selectCell(selectedCell.r + 1, selectedCell.c);
    }
  });

  // Keyboard Navigation
  window.addEventListener('keydown', (e) => {
    if (e.target === formulaInput) return;

    if ((e.ctrlKey || e.metaKey) && e.key === 's') {
      e.preventDefault();
      saveCurrentWorkbook();
      return;
    }
    if ((e.ctrlKey || e.metaKey) && e.key === 'b') {
      e.preventDefault();
      toolBold.click();
      return;
    }
    if ((e.ctrlKey || e.metaKey) && e.key === 'i') {
      e.preventDefault();
      toolItalic.click();
      return;
    }
    if ((e.ctrlKey || e.metaKey) && e.key === 'c') {
      toolCopy.click();
      return;
    }
    if ((e.ctrlKey || e.metaKey) && e.key === 'v') {
      toolPaste.click();
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
      case 'gridcraft:openWorkbook':
        currentFileName = message.filename;
        docTitleEl.textContent = message.filename;
        sheets = message.sheets || { Sheet1: createInitialGrid(35, 16) };
        const sheetNames = message.sheetNames || Object.keys(sheets);
        
        sheetTabsEl.innerHTML = '';
        sheetNames.forEach(name => {
          const tab = document.createElement('button');
          tab.className = `sheet-tab ${name === sheetNames[0] ? 'active' : ''}`;
          tab.textContent = name;
          tab.setAttribute('data-sheet', name);
          tab.addEventListener('click', () => switchSheet(name));
          sheetTabsEl.appendChild(tab);
        });
        sheetTabsEl.appendChild(btnAddSheet);
        activeSheet = sheetNames[0] || 'Sheet1';
        renderGrid();
        break;

      case 'gridcraft:openText':
        currentFileName = message.filename;
        docTitleEl.textContent = message.filename;
        sheets = { Sheet1: createInitialGrid(35, 16) };
        activeSheet = 'Sheet1';
        renderGrid();
        break;
    }
  });

  // Signal Ready to VS Code Host
  vscode.postMessage({ type: 'gridcraft:ready' });
})();
