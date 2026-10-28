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
  let colWidths = {};
  let rowHeights = {};
  let isResizingCol = false;
  let resizeColIndex = -1;
  let resizeStartX = 0;
  let resizeStartWidth = 0;
  let isResizingRow = false;
  let resizeRowIndex = -1;
  let resizeStartY = 0;
  let resizeStartHeight = 0;
  let isDraggingFill = false;
  let fillTargetR = -1;
  let fillTargetC = -1;

  // Advanced Enterprise State
  let isFilterActive = false;
  let activeFilters = {}; // { [colIdx]: Set of allowed strings }
  let currentFilterCol = -1;
  let findMatches = [];
  let findMatchIndex = -1;
  let isReplaceMode = false;
  let targetContextMenuSheet = '';

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

  // Advanced Enterprise Controls
  const findReplaceBar = document.getElementById('find-replace-bar');
  const findInput = document.getElementById('find-input');
  const replaceInput = document.getElementById('replace-input');
  const btnToggleReplace = document.getElementById('btn-toggle-replace');
  const btnFindPrev = document.getElementById('btn-find-prev');
  const btnFindNext = document.getElementById('btn-find-next');
  const btnReplaceOne = document.getElementById('btn-replace-one');
  const btnReplaceAll = document.getElementById('btn-replace-all');
  const findMatchBadge = document.getElementById('find-match-badge');
  const btnCloseFind = document.getElementById('btn-close-find');

  const filterDropdownMenu = document.getElementById('filter-dropdown-menu');
  const filterColTitle = document.getElementById('filter-col-title');
  const filterSearchInput = document.getElementById('filter-search-input');
  const filterItemsList = document.getElementById('filter-items-list');
  const btnFilterApply = document.getElementById('btn-filter-apply');
  const btnFilterClear = document.getElementById('btn-filter-clear');
  const btnFilterSortAsc = document.getElementById('btn-filter-sort-asc');
  const btnFilterSortDesc = document.getElementById('btn-filter-sort-desc');

  const formulaAutocomplete = document.getElementById('formula-autocomplete');
  const formulaSignatureHint = document.getElementById('formula-signature-hint');
  const sheetTabContextMenu = document.getElementById('sheet-tab-context-menu');

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
      alignH: '',
      alignV: 'middle',
      wrapText: false,
      border: 'none',
      numFmt: 'General',
      decimals: null,
      merge: null,
      mergedInto: null
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

  // --- FUNCTION CATALOG & METADATA ---
  const FUNCTION_CATALOG = [
    { name: 'SUM', sig: 'SUM(number1, [number2], ...)', desc: 'Adds all the numbers in a range of cells.' },
    { name: 'AVERAGE', sig: 'AVERAGE(number1, [number2], ...)', desc: 'Returns the average (arithmetic mean) of arguments.' },
    { name: 'COUNT', sig: 'COUNT(value1, [value2], ...)', desc: 'Counts the number of cells that contain numbers.' },
    { name: 'COUNTA', sig: 'COUNTA(value1, [value2], ...)', desc: 'Counts the number of cells that are not empty.' },
    { name: 'COUNTBLANK', sig: 'COUNTBLANK(range)', desc: 'Counts the number of empty cells in a range.' },
    { name: 'MAX', sig: 'MAX(number1, [number2], ...)', desc: 'Returns the largest value in a set of values.' },
    { name: 'MIN', sig: 'MIN(number1, [number2], ...)', desc: 'Returns the smallest value in a set of values.' },
    { name: 'MEDIAN', sig: 'MEDIAN(number1, [number2], ...)', desc: 'Returns the median of the given numbers.' },
    { name: 'PRODUCT', sig: 'PRODUCT(number1, [number2], ...)', desc: 'Multiplies all the numbers given as arguments.' },
    { name: 'IF', sig: 'IF(logical_test, value_if_true, [value_if_false])', desc: 'Returns one value if TRUE and another if FALSE.' },
    { name: 'IFS', sig: 'IFS(logical_test1, value1, [logical_test2, value2], ...)', desc: 'Checks whether one or more conditions are met and returns corresponding value.' },
    { name: 'IFERROR', sig: 'IFERROR(value, value_if_error)', desc: 'Returns value_if_error if expression is an error, otherwise returns value.' },
    { name: 'AND', sig: 'AND(logical1, [logical2], ...)', desc: 'Returns TRUE if all its arguments are TRUE.' },
    { name: 'OR', sig: 'OR(logical1, [logical2], ...)', desc: 'Returns TRUE if any argument is TRUE.' },
    { name: 'NOT', sig: 'NOT(logical)', desc: 'Reverses the logic of its argument.' },
    { name: 'ISBLANK', sig: 'ISBLANK(value)', desc: 'Returns TRUE if the value is blank.' },
    { name: 'ISNUMBER', sig: 'ISNUMBER(value)', desc: 'Returns TRUE if the value is a number.' },
    { name: 'ISTEXT', sig: 'ISTEXT(value)', desc: 'Returns TRUE if the value is text.' },
    { name: 'VLOOKUP', sig: 'VLOOKUP(lookup_value, table_array, col_index, [range_lookup])', desc: 'Looks for a value in the leftmost column of a table and returns a value in the same row.' },
    { name: 'HLOOKUP', sig: 'HLOOKUP(lookup_value, table_array, row_index, [range_lookup])', desc: 'Looks for a value in the top row of a table and returns a value in the same column.' },
    { name: 'XLOOKUP', sig: 'XLOOKUP(lookup_value, lookup_array, return_array, [if_not_found], [match_mode])', desc: 'Searches a range/array and returns the matching item from a second range/array.' },
    { name: 'INDEX', sig: 'INDEX(array, row_num, [col_num])', desc: 'Returns a value or reference of the cell at the intersection of row and column.' },
    { name: 'MATCH', sig: 'MATCH(lookup_value, lookup_array, [match_type])', desc: 'Returns the relative position of an item in an array that matches a specified value.' },
    { name: 'CHOOSE', sig: 'CHOOSE(index_num, value1, [value2], ...)', desc: 'Chooses a value from a list of values based on an index number.' },
    { name: 'SUMIF', sig: 'SUMIF(range, criteria, [sum_range])', desc: 'Adds the cells specified by a given condition or criteria.' },
    { name: 'COUNTIF', sig: 'COUNTIF(range, criteria)', desc: 'Counts the number of cells within a range that meet the given criteria.' },
    { name: 'AVERAGEIF', sig: 'AVERAGEIF(range, criteria, [average_range])', desc: 'Finds average for the cells specified by a given condition or criteria.' },
    { name: 'ROUND', sig: 'ROUND(number, num_digits)', desc: 'Rounds a number to a specified number of digits.' },
    { name: 'ROUNDUP', sig: 'ROUNDUP(number, num_digits)', desc: 'Rounds a number up, away from zero.' },
    { name: 'ROUNDDOWN', sig: 'ROUNDDOWN(number, num_digits)', desc: 'Rounds a number down, toward zero.' },
    { name: 'ABS', sig: 'ABS(number)', desc: 'Returns the absolute value of a number.' },
    { name: 'MOD', sig: 'MOD(number, divisor)', desc: 'Returns the remainder after number is divided by divisor.' },
    { name: 'POWER', sig: 'POWER(number, power)', desc: 'Returns the result of a number raised to a power.' },
    { name: 'SQRT', sig: 'SQRT(number)', desc: 'Returns a positive square root.' },
    { name: 'INT', sig: 'INT(number)', desc: 'Rounds a number down to the nearest integer.' },
    { name: 'CONCAT', sig: 'CONCAT(text1, [text2], ...)', desc: 'Combines the text from multiple ranges and/or strings.' },
    { name: 'CONCATENATE', sig: 'CONCATENATE(text1, [text2], ...)', desc: 'Joins several text items into one text item.' },
    { name: 'LEFT', sig: 'LEFT(text, [num_chars])', desc: 'Returns the specified number of characters from the start of a text string.' },
    { name: 'RIGHT', sig: 'RIGHT(text, [num_chars])', desc: 'Returns the specified number of characters from the end of a text string.' },
    { name: 'MID', sig: 'MID(text, start_num, num_chars)', desc: 'Returns a specific number of characters from a text string.' },
    { name: 'LEN', sig: 'LEN(text)', desc: 'Returns the number of characters in a text string.' },
    { name: 'TRIM', sig: 'TRIM(text)', desc: 'Removes all spaces from text except for single spaces between words.' },
    { name: 'UPPER', sig: 'UPPER(text)', desc: 'Converts text to uppercase.' },
    { name: 'LOWER', sig: 'LOWER(text)', desc: 'Converts all letters in a text string to lowercase.' },
    { name: 'PROPER', sig: 'PROPER(text)', desc: 'Capitalizes the first letter in each word of a text value.' },
    { name: 'TODAY', sig: 'TODAY()', desc: 'Returns the current date formatted as YYYY-MM-DD.' },
    { name: 'NOW', sig: 'NOW()', desc: 'Returns the current date and time.' },
    { name: 'DATE', sig: 'DATE(year, month, day)', desc: 'Returns the serial number that represents a particular date.' },
    { name: 'YEAR', sig: 'YEAR(serial_number)', desc: 'Returns the year of a date.' },
    { name: 'MONTH', sig: 'MONTH(serial_number)', desc: 'Returns the month of a date.' },
    { name: 'DAY', sig: 'DAY(serial_number)', desc: 'Returns the day of the month.' }
  ];

  // Split arguments at depth 0
  function splitArguments(argString) {
    const args = [];
    let current = '';
    let depth = 0;
    let inQuote = false;

    for (let i = 0; i < argString.length; i++) {
      const ch = argString[i];
      if (ch === '"') {
        inQuote = !inQuote;
        current += ch;
      } else if (ch === '(' && !inQuote) {
        depth++;
        current += ch;
      } else if (ch === ')' && !inQuote) {
        depth--;
        current += ch;
      } else if (ch === ',' && depth === 0 && !inQuote) {
        args.push(current.trim());
        current = '';
      } else {
        current += ch;
      }
    }
    if (current.trim()) {
      args.push(current.trim());
    }
    return args;
  }

  function matchesCriteria(val, criteria) {
    if (criteria === undefined || criteria === null) return false;
    const critStr = String(criteria).trim();
    const match = critStr.match(/^(>=|<=|<>|!=|>|<|=)(.*)$/);
    if (match) {
      const op = match[1];
      const target = match[2].trim();
      const numVal = parseFloat(val);
      const numTarget = parseFloat(target);
      if (!isNaN(numVal) && !isNaN(numTarget)) {
        if (op === '>') return numVal > numTarget;
        if (op === '<') return numVal < numTarget;
        if (op === '>=') return numVal >= numTarget;
        if (op === '<=') return numVal <= numTarget;
        if (op === '=' || op === '==') return numVal === numTarget;
        if (op === '<>' || op === '!=') return numVal !== numTarget;
      }
      const strVal = String(val).toLowerCase();
      const strTarget = target.toLowerCase();
      if (op === '=' || op === '==') return strVal === strTarget;
      if (op === '<>' || op === '!=') return strVal !== strTarget;
    }
    return String(val).toLowerCase() === critStr.toLowerCase();
  }

  function get2DRange(rangeStr, grid) {
    const parts = rangeStr.split(':');
    if (parts.length === 1) {
      const c = parseCellCoords(parts[0].trim());
      if (c && grid[c.r] && grid[c.r][c.c]) {
        return [[grid[c.r][c.c].val]];
      }
      return [[]];
    }
    const start = parseCellCoords(parts[0].trim());
    const end = parseCellCoords(parts[1].trim());
    if (!start || !end) return [[]];

    const minR = Math.min(start.r, end.r);
    const maxR = Math.max(start.r, end.r);
    const minC = Math.min(start.c, end.c);
    const maxC = Math.max(start.c, end.c);

    const table = [];
    for (let r = minR; r <= maxR; r++) {
      const row = [];
      for (let c = minC; c <= maxC; c++) {
        row.push(grid[r] && grid[r][c] ? grid[r][c].val : '');
      }
      table.push(row);
    }
    return table;
  }

  // --- RECURSIVE ENTERPRISE FORMULA EVALUATOR ---
  function evaluateFormula(input, grid, depth = 0) {
    if (input === null || input === undefined) return '';
    if (depth > 25) return '#RECURSION!';

    let expr = String(input).trim();
    if (expr.startsWith('=')) {
      expr = expr.substring(1).trim();
    }

    if (!expr) return '';

    // Literal Quoted String: "Hello"
    if (expr.startsWith('"') && expr.endsWith('"') && expr.length >= 2) {
      return expr.substring(1, expr.length - 1);
    }

    // Pure Number
    const rawNum = Number(expr);
    if (!isNaN(rawNum) && !expr.includes(':') && !/[A-Z]/i.test(expr)) {
      return rawNum;
    }

    // Direct Cell Reference: A1, B12
    const singleCell = parseCellCoords(expr);
    if (singleCell && grid[singleCell.r] && grid[singleCell.r][singleCell.c]) {
      return grid[singleCell.r][singleCell.c].val;
    }

    // String Concatenation: expr & expr
    let quote = false, paren = 0, concatIdx = -1;
    for (let i = 0; i < expr.length; i++) {
      if (expr[i] === '"') quote = !quote;
      else if (expr[i] === '(' && !quote) paren++;
      else if (expr[i] === ')' && !quote) paren--;
      else if (expr[i] === '&' && !quote && paren === 0) {
        concatIdx = i;
        break;
      }
    }
    if (concatIdx !== -1) {
      const left = evaluateFormula(expr.substring(0, concatIdx), grid, depth + 1);
      const right = evaluateFormula(expr.substring(concatIdx + 1), grid, depth + 1);
      return String(left) + String(right);
    }

    // Binary Comparisons: >=, <=, <>, !=, ==, =, >, <
    const compOps = ['>=', '<=', '<>', '!=', '==', '=', '>', '<'];
    for (const op of compOps) {
      quote = false; paren = 0;
      let opIdx = -1;
      for (let i = 0; i < expr.length; i++) {
        if (expr[i] === '"') quote = !quote;
        else if (expr[i] === '(' && !quote) paren++;
        else if (expr[i] === ')' && !quote) paren--;
        else if (!quote && paren === 0 && expr.startsWith(op, i)) {
          opIdx = i;
          break;
        }
      }
      if (opIdx !== -1) {
        const left = evaluateFormula(expr.substring(0, opIdx), grid, depth + 1);
        const right = evaluateFormula(expr.substring(opIdx + op.length), grid, depth + 1);
        const numL = parseFloat(left);
        const numR = parseFloat(right);
        if (!isNaN(numL) && !isNaN(numR)) {
          if (op === '>') return numL > numR;
          if (op === '<') return numL < numR;
          if (op === '>=') return numL >= numR;
          if (op === '<=') return numL <= numR;
          if (op === '=' || op === '==') return numL === numR;
          if (op === '<>' || op === '!=') return numL !== numR;
        }
        const sL = String(left).toLowerCase();
        const sR = String(right).toLowerCase();
        if (op === '=' || op === '==') return sL === sR;
        if (op === '<>' || op === '!=') return sL !== sR;
      }
    }

    // Binary Arithmetic (+, - outside parens)
    quote = false; paren = 0;
    for (let i = expr.length - 1; i >= 0; i--) {
      if (expr[i] === '"') quote = !quote;
      else if (expr[i] === ')' && !quote) paren++;
      else if (expr[i] === '(' && !quote) paren--;
      else if ((expr[i] === '+' || expr[i] === '-') && !quote && paren === 0 && i > 0 && !/[+\-*\/^]/.test(expr[i - 1])) {
        const left = evaluateFormula(expr.substring(0, i), grid, depth + 1);
        const right = evaluateFormula(expr.substring(i + 1), grid, depth + 1);
        return expr[i] === '+' ? (parseFloat(left) || 0) + (parseFloat(right) || 0) : (parseFloat(left) || 0) - (parseFloat(right) || 0);
      }
    }

    // Binary Arithmetic (*, /, ^ outside parens)
    quote = false; paren = 0;
    for (let i = expr.length - 1; i >= 0; i--) {
      if (expr[i] === '"') quote = !quote;
      else if (expr[i] === ')' && !quote) paren++;
      else if (expr[i] === '(' && !quote) paren--;
      else if ((expr[i] === '*' || expr[i] === '/' || expr[i] === '^') && !quote && paren === 0) {
        const left = parseFloat(evaluateFormula(expr.substring(0, i), grid, depth + 1)) || 0;
        const right = parseFloat(evaluateFormula(expr.substring(i + 1), grid, depth + 1)) || 0;
        if (expr[i] === '*') return left * right;
        if (expr[i] === '/') return right !== 0 ? left / right : '#DIV/0!';
        if (expr[i] === '^') return Math.pow(left, right);
      }
    }

    // Function Calls: NAME(...)
    const fnMatch = expr.match(/^([A-Z0-9_\.]+)\s*\((.*)\)$/is);
    if (fnMatch) {
      const fnName = fnMatch[1].toUpperCase();
      const rawArgs = splitArguments(fnMatch[2]);

      // --- LOOKUP FUNCTIONS ---
      if (fnName === 'VLOOKUP') {
        const lookupVal = evaluateFormula(rawArgs[0], grid, depth + 1);
        const tableRange = rawArgs[1] ? rawArgs[1].trim() : '';
        const colIdx = parseInt(evaluateFormula(rawArgs[2], grid, depth + 1), 10);
        const table = get2DRange(tableRange, grid);
        for (let r = 0; r < table.length; r++) {
          if (String(table[r][0]).toLowerCase() === String(lookupVal).toLowerCase()) {
            return table[r][colIdx - 1] !== undefined ? table[r][colIdx - 1] : '';
          }
        }
        return '#N/A';
      }

      if (fnName === 'HLOOKUP') {
        const lookupVal = evaluateFormula(rawArgs[0], grid, depth + 1);
        const tableRange = rawArgs[1] ? rawArgs[1].trim() : '';
        const rowIdx = parseInt(evaluateFormula(rawArgs[2], grid, depth + 1), 10);
        const table = get2DRange(tableRange, grid);
        if (table.length > 0) {
          for (let c = 0; c < table[0].length; c++) {
            if (String(table[0][c]).toLowerCase() === String(lookupVal).toLowerCase()) {
              return table[rowIdx - 1] && table[rowIdx - 1][c] !== undefined ? table[rowIdx - 1][c] : '';
            }
          }
        }
        return '#N/A';
      }

      if (fnName === 'XLOOKUP') {
        const lookupVal = evaluateFormula(rawArgs[0], grid, depth + 1);
        const lookupList = getRangeValues(rawArgs[1], grid);
        const returnList = getRangeValues(rawArgs[2], grid);
        const ifNotFound = rawArgs[3] !== undefined ? evaluateFormula(rawArgs[3], grid, depth + 1) : '#N/A';
        for (let i = 0; i < lookupList.length; i++) {
          if (String(lookupList[i]).toLowerCase() === String(lookupVal).toLowerCase()) {
            return returnList[i] !== undefined ? returnList[i] : '';
          }
        }
        return ifNotFound;
      }

      if (fnName === 'INDEX') {
        const table = get2DRange(rawArgs[0], grid);
        const rNum = parseInt(evaluateFormula(rawArgs[1], grid, depth + 1), 10);
        const cNum = rawArgs[2] !== undefined ? parseInt(evaluateFormula(rawArgs[2], grid, depth + 1), 10) : 1;
        if (table[rNum - 1] && table[rNum - 1][cNum - 1] !== undefined) {
          return table[rNum - 1][cNum - 1];
        }
        return '#REF!';
      }

      if (fnName === 'MATCH') {
        const lookupVal = evaluateFormula(rawArgs[0], grid, depth + 1);
        const list = getRangeValues(rawArgs[1], grid);
        for (let i = 0; i < list.length; i++) {
          if (String(list[i]).toLowerCase() === String(lookupVal).toLowerCase()) {
            return i + 1;
          }
        }
        return '#N/A';
      }

      if (fnName === 'CHOOSE') {
        const idx = parseInt(evaluateFormula(rawArgs[0], grid, depth + 1), 10);
        if (idx >= 1 && idx < rawArgs.length) {
          return evaluateFormula(rawArgs[idx], grid, depth + 1);
        }
        return '#VALUE!';
      }

      // --- LOGICAL & ERROR FUNCTIONS ---
      if (fnName === 'IF') {
        const cond = evaluateFormula(rawArgs[0], grid, depth + 1);
        const isTrue = cond === true || (typeof cond === 'number' && cond !== 0) || (typeof cond === 'string' && cond.toLowerCase() === 'true');
        if (isTrue) {
          return evaluateFormula(rawArgs[1], grid, depth + 1);
        } else {
          return rawArgs[2] !== undefined ? evaluateFormula(rawArgs[2], grid, depth + 1) : false;
        }
      }

      if (fnName === 'IFS') {
        for (let i = 0; i < rawArgs.length; i += 2) {
          const cond = evaluateFormula(rawArgs[i], grid, depth + 1);
          if (cond === true || (typeof cond === 'number' && cond !== 0) || (typeof cond === 'string' && cond.toLowerCase() === 'true')) {
            return evaluateFormula(rawArgs[i + 1], grid, depth + 1);
          }
        }
        return '#N/A';
      }

      if (fnName === 'IFERROR') {
        try {
          const val = evaluateFormula(rawArgs[0], grid, depth + 1);
          if (val === '#N/A' || val === '#DIV/0!' || val === '#VALUE!' || val === '#REF!' || val === '#ERROR!' || (typeof val === 'number' && isNaN(val))) {
            return evaluateFormula(rawArgs[1], grid, depth + 1);
          }
          return val;
        } catch {
          return evaluateFormula(rawArgs[1], grid, depth + 1);
        }
      }

      if (fnName === 'AND') {
        return rawArgs.every(a => {
          const v = evaluateFormula(a, grid, depth + 1);
          return v === true || (typeof v === 'number' && v !== 0) || String(v).toLowerCase() === 'true';
        });
      }

      if (fnName === 'OR') {
        return rawArgs.some(a => {
          const v = evaluateFormula(a, grid, depth + 1);
          return v === true || (typeof v === 'number' && v !== 0) || String(v).toLowerCase() === 'true';
        });
      }

      if (fnName === 'NOT') {
        const v = evaluateFormula(rawArgs[0], grid, depth + 1);
        return !(v === true || (typeof v === 'number' && v !== 0) || String(v).toLowerCase() === 'true');
      }

      if (fnName === 'ISBLANK') {
        const v = evaluateFormula(rawArgs[0], grid, depth + 1);
        return v === '' || v === null || v === undefined;
      }

      if (fnName === 'ISNUMBER') {
        const v = evaluateFormula(rawArgs[0], grid, depth + 1);
        return typeof v === 'number' && !isNaN(v);
      }

      if (fnName === 'ISTEXT') {
        const v = evaluateFormula(rawArgs[0], grid, depth + 1);
        return typeof v === 'string' && isNaN(Number(v));
      }

      // --- CONDITIONAL AGGREGATIONS ---
      if (fnName === 'SUMIF') {
        const checkRange = getRangeValues(rawArgs[0], grid);
        const crit = evaluateFormula(rawArgs[1], grid, depth + 1);
        const sumRange = rawArgs[2] ? getRangeValues(rawArgs[2], grid) : checkRange;
        let total = 0;
        for (let i = 0; i < checkRange.length; i++) {
          if (matchesCriteria(checkRange[i], crit)) {
            total += parseFloat(sumRange[i]) || 0;
          }
        }
        return total;
      }

      if (fnName === 'COUNTIF') {
        const checkRange = getRangeValues(rawArgs[0], grid);
        const crit = evaluateFormula(rawArgs[1], grid, depth + 1);
        let count = 0;
        for (let i = 0; i < checkRange.length; i++) {
          if (matchesCriteria(checkRange[i], crit)) count++;
        }
        return count;
      }

      if (fnName === 'AVERAGEIF') {
        const checkRange = getRangeValues(rawArgs[0], grid);
        const crit = evaluateFormula(rawArgs[1], grid, depth + 1);
        const avgRange = rawArgs[2] ? getRangeValues(rawArgs[2], grid) : checkRange;
        let sum = 0, count = 0;
        for (let i = 0; i < checkRange.length; i++) {
          if (matchesCriteria(checkRange[i], crit)) {
            sum += parseFloat(avgRange[i]) || 0;
            count++;
          }
        }
        return count ? Math.round((sum / count) * 100) / 100 : 0;
      }

      // --- MATH & STATS ---
      if (fnName === 'ROUND' || fnName === 'ROUNDUP' || fnName === 'ROUNDDOWN') {
        const num = parseFloat(evaluateFormula(rawArgs[0], grid, depth + 1)) || 0;
        const digits = parseInt(evaluateFormula(rawArgs[1] || '0', grid, depth + 1), 10);
        const factor = Math.pow(10, digits);
        if (fnName === 'ROUND') return Math.round(num * factor) / factor;
        if (fnName === 'ROUNDUP') return Math.ceil(num * factor) / factor;
        if (fnName === 'ROUNDDOWN') return Math.floor(num * factor) / factor;
      }

      if (fnName === 'ABS') return Math.abs(parseFloat(evaluateFormula(rawArgs[0], grid, depth + 1)) || 0);
      if (fnName === 'MOD') {
        const n = parseFloat(evaluateFormula(rawArgs[0], grid, depth + 1)) || 0;
        const d = parseFloat(evaluateFormula(rawArgs[1], grid, depth + 1)) || 1;
        return n % d;
      }
      if (fnName === 'POWER') {
        return Math.pow(parseFloat(evaluateFormula(rawArgs[0], grid, depth + 1)) || 0, parseFloat(evaluateFormula(rawArgs[1], grid, depth + 1)) || 0);
      }
      if (fnName === 'SQRT') return Math.sqrt(parseFloat(evaluateFormula(rawArgs[0], grid, depth + 1)) || 0);
      if (fnName === 'INT') return Math.floor(parseFloat(evaluateFormula(rawArgs[0], grid, depth + 1)) || 0);

      // --- TEXT FUNCTIONS ---
      if (fnName === 'CONCAT' || fnName === 'CONCATENATE') {
        return rawArgs.map(a => {
          if (/^[A-Z]+\d+:[A-Z]+\d+$/i.test(a.trim())) return getRangeValues(a, grid).join('');
          return String(evaluateFormula(a, grid, depth + 1) || '');
        }).join('');
      }

      if (fnName === 'LEFT') {
        const str = String(evaluateFormula(rawArgs[0], grid, depth + 1) || '');
        const len = rawArgs[1] ? parseInt(evaluateFormula(rawArgs[1], grid, depth + 1), 10) : 1;
        return str.substring(0, len);
      }

      if (fnName === 'RIGHT') {
        const str = String(evaluateFormula(rawArgs[0], grid, depth + 1) || '');
        const len = rawArgs[1] ? parseInt(evaluateFormula(rawArgs[1], grid, depth + 1), 10) : 1;
        return str.slice(-len);
      }

      if (fnName === 'MID') {
        const str = String(evaluateFormula(rawArgs[0], grid, depth + 1) || '');
        const start = parseInt(evaluateFormula(rawArgs[1], grid, depth + 1), 10);
        const len = parseInt(evaluateFormula(rawArgs[2], grid, depth + 1), 10);
        return str.substring(start - 1, start - 1 + len);
      }

      if (fnName === 'LEN') return String(evaluateFormula(rawArgs[0], grid, depth + 1) || '').length;
      if (fnName === 'TRIM') return String(evaluateFormula(rawArgs[0], grid, depth + 1) || '').trim().replace(/\s+/g, ' ');
      if (fnName === 'UPPER') return String(evaluateFormula(rawArgs[0], grid, depth + 1) || '').toUpperCase();
      if (fnName === 'LOWER') return String(evaluateFormula(rawArgs[0], grid, depth + 1) || '').toLowerCase();
      if (fnName === 'PROPER') {
        return String(evaluateFormula(rawArgs[0], grid, depth + 1) || '').replace(/\b\w/g, l => l.toUpperCase());
      }

      // --- DATE & TIME ---
      if (fnName === 'TODAY') return new Date().toISOString().split('T')[0];
      if (fnName === 'NOW') return new Date().toLocaleString('ko-KR');
      if (fnName === 'DATE') {
        const y = parseInt(evaluateFormula(rawArgs[0], grid, depth + 1), 10);
        const m = parseInt(evaluateFormula(rawArgs[1], grid, depth + 1), 10) - 1;
        const d = parseInt(evaluateFormula(rawArgs[2], grid, depth + 1), 10);
        return new Date(Date.UTC(y, m, d)).toISOString().split('T')[0];
      }
      if (fnName === 'YEAR' || fnName === 'MONTH' || fnName === 'DAY') {
        const dVal = new Date(evaluateFormula(rawArgs[0], grid, depth + 1));
        if (isNaN(dVal.getTime())) return '#VALUE!';
        if (fnName === 'YEAR') return dVal.getFullYear();
        if (fnName === 'MONTH') return dVal.getMonth() + 1;
        if (fnName === 'DAY') return dVal.getDate();
      }

      // --- STANDARD AGGREGATIONS ---
      const allVals = [];
      for (const a of rawArgs) {
        if (/^[A-Z]+\d+:[A-Z]+\d+$/i.test(a.trim())) {
          allVals.push(...getRangeValues(a, grid));
        } else {
          allVals.push(evaluateFormula(a, grid, depth + 1));
        }
      }

      const numList = allVals.map(Number).filter(n => !isNaN(n));
      if (fnName === 'SUM') return numList.reduce((acc, v) => acc + v, 0);
      if (fnName === 'AVERAGE') {
        return numList.length ? Math.round((numList.reduce((acc, v) => acc + v, 0) / numList.length) * 100) / 100 : 0;
      }
      if (fnName === 'COUNT') return numList.length;
      if (fnName === 'COUNTA') return allVals.filter(v => v !== '' && v !== null && v !== undefined).length;
      if (fnName === 'COUNTBLANK') return allVals.filter(v => v === '' || v === null || v === undefined).length;
      if (fnName === 'MAX') return numList.length ? Math.max(...numList) : 0;
      if (fnName === 'MIN') return numList.length ? Math.min(...numList) : 0;
      if (fnName === 'PRODUCT') return numList.length ? numList.reduce((acc, v) => acc * v, 1) : 0;
      if (fnName === 'MEDIAN') {
        if (!numList.length) return 0;
        numList.sort((a, b) => a - b);
        const mid = Math.floor(numList.length / 2);
        return numList.length % 2 !== 0 ? numList[mid] : (numList[mid - 1] + numList[mid]) / 2;
      }
    }

    return expr;
  }

  function getRangeValues(rangeStr, grid) {
    if (!rangeStr) return [];
    if (rangeStr.includes(',')) {
      const parts = rangeStr.split(',');
      const allVals = [];
      for (const p of parts) {
        allVals.push(...getRangeValues(p.trim(), grid));
      }
      return allVals;
    }
    const litNum = Number(rangeStr);
    if (!isNaN(litNum) && rangeStr.trim() !== '') {
      return [litNum];
    }
    const parts = rangeStr.split(':');
    if (parts.length === 1) {
      const c = parseCellCoords(parts[0].trim());
      if (c && grid[c.r] && grid[c.r][c.c]) {
        return [grid[c.r][c.c].val];
      }
      return [];
    }
    const start = parseCellCoords(parts[0].trim());
    const end = parseCellCoords(parts[1].trim());
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

    const minR = Math.min(selectionRange.startR, selectionRange.endR);
    const maxR = Math.max(selectionRange.startR, selectionRange.endR);
    const minC = Math.min(selectionRange.startC, selectionRange.endC);
    const maxC = Math.max(selectionRange.startC, selectionRange.endC);

    // Header
    let theadHtml = '<tr><th class="corner-header"></th>';
    for (let c = 0; c < numCols; c++) {
      const w = colWidths[c] ? `style="width:${colWidths[c]}px;min-width:${colWidths[c]}px;"` : '';
      const isActive = (c >= minC && c <= maxC);
      let filterBtn = '';
      if (isFilterActive) {
        const isFiltered = activeFilters[c] !== undefined;
        filterBtn = `<span class="col-filter-btn ${isFiltered ? 'filtered' : ''}" data-c="${c}">▾</span>`;
      }
      theadHtml += `<th class="col-header ${isActive ? 'active' : ''}" data-c="${c}" ${w}>
        <div class="col-header-content">
          <span>${colIndexToName(c)}</span>
          ${filterBtn}
        </div>
        <div class="col-resizer" data-c="${c}"></div>
      </th>`;
    }
    theadHtml += '</tr>';
    theadEl.innerHTML = theadHtml;

    // Body
    let tbodyHtml = '';

    for (let r = 0; r < numRows; r++) {
      let isRowVisible = true;
      if (isFilterActive && Object.keys(activeFilters).length > 0 && r > 0) {
        for (const [colIdx, allowedSet] of Object.entries(activeFilters)) {
          const cellVal = String(grid[r][colIdx] ? grid[r][colIdx].val : '');
          if (!allowedSet.has(cellVal)) {
            isRowVisible = false;
            break;
          }
        }
      }

      const rowStyle = (rowHeights[r] ? `height:${rowHeights[r]}px;` : '') + (isRowVisible ? '' : 'display:none;');
      const h = rowStyle ? `style="${rowStyle}"` : '';
      const isActive = (r >= minR && r <= maxR);
      tbodyHtml += `<tr ${h}><th class="row-header ${isActive ? 'active' : ''}" data-r="${r}">${r + 1}<div class="row-resizer" data-r="${r}"></div></th>`;
      for (let c = 0; c < numCols; c++) {
        const cell = grid[r][c];
        if (cell.mergedInto) {
          continue; // Spanned by top-left cell
        }

        const isFocused = (r === selectedCell.r && c === selectedCell.c);
        const inRange = (r >= minR && r <= maxR && c >= minC && c <= maxC);

        let spanAttrs = '';
        if (cell.merge) {
          if (cell.merge.rowspan > 1) spanAttrs += ` rowspan="${cell.merge.rowspan}"`;
          if (cell.merge.colspan > 1) spanAttrs += ` colspan="${cell.merge.colspan}"`;
        }

        let style = '';
        if (cell.bold) style += 'font-weight:bold;';
        if (cell.italic) style += 'font-style:italic;';
        if (cell.underline) style += 'text-decoration:underline;';
        if (cell.strike) style += 'text-decoration:line-through;';
        if (cell.fontFamily) style += `font-family:${cell.fontFamily};`;
        if (cell.fontSize) style += `font-size:${cell.fontSize}px;`;
        if (cell.color) style += `color:${cell.color};`;
        if (cell.bg) style += `background-color:${cell.bg};`;

        // Auto alignment: numbers right-aligned, text left-aligned
        let textAlign = cell.alignH;
        if (!textAlign) {
          const isNum = !isNaN(Number(cell.val)) && cell.val !== '' && typeof cell.val !== 'boolean';
          textAlign = isNum ? 'right' : 'left';
        }
        style += `text-align:${textAlign};`;

        if (cell.alignV) style += `vertical-align:${cell.alignV};`;
        if (cell.wrapText) style += 'white-space:normal;';

        // Borders
        if (cell.border === 'all') style += 'border:1px solid #000;';
        else if (cell.border === 'thick') style += 'border:2px solid #000;';
        else if (cell.border === 'bottom') style += 'border-bottom:2px solid #000;';

        const displayVal = formatDisplayValue(cell);
        const fillHandle = isFocused ? '<div class="auto-fill-handle" id="fill-handle"></div>' : '';
        tbodyHtml += `<td class="grid-cell ${isFocused ? 'selected' : ''} ${inRange ? 'in-range' : ''}" data-r="${r}" data-c="${c}" ${spanAttrs} style="${style}">${escapeHtml(displayVal)}${fillHandle}</td>`;
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
    const grid = sheets[activeSheet];
    if (!grid) return;
    const minR = Math.min(selectionRange.startR, selectionRange.endR);
    const maxR = Math.max(selectionRange.startR, selectionRange.endR);
    const minC = Math.min(selectionRange.startC, selectionRange.endC);
    const maxC = Math.max(selectionRange.startC, selectionRange.endC);

    const topLeft = grid[minR][minC];
    if (topLeft.merge) {
      // Unmerge
      for (let r = minR; r <= maxR; r++) {
        for (let c = minC; c <= maxC; c++) {
          if (grid[r] && grid[r][c]) {
            grid[r][c].merge = null;
            grid[r][c].mergedInto = null;
          }
        }
      }
    } else {
      // Merge
      const rowSpan = maxR - minR + 1;
      const colSpan = maxC - minC + 1;
      if (rowSpan === 1 && colSpan === 1) return;

      topLeft.merge = { rowspan: rowSpan, colspan: colSpan };
      topLeft.alignH = 'center';

      for (let r = minR; r <= maxR; r++) {
        for (let c = minC; c <= maxC; c++) {
          if (r === minR && c === minC) continue;
          if (grid[r] && grid[r][c]) {
            grid[r][c].mergedInto = { r: minR, c: minC };
          }
        }
      }
    }
    markDirty(true);
    renderGrid();
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

  // --- AUTOFILTER CONTROLLER ---
  toolToggleFilter.addEventListener('click', () => {
    isFilterActive = !isFilterActive;
    toolToggleFilter.classList.toggle('active', isFilterActive);
    if (!isFilterActive) {
      activeFilters = {};
    }
    renderGrid();
  });

  theadEl.addEventListener('click', (e) => {
    const filterBtn = e.target.closest('.col-filter-btn');
    if (!filterBtn) return;
    e.stopPropagation();
    const c = parseInt(filterBtn.getAttribute('data-c'), 10);
    currentFilterCol = c;
    const rect = filterBtn.getBoundingClientRect();
    filterDropdownMenu.style.top = `${rect.bottom + 4}px`;
    filterDropdownMenu.style.left = `${Math.min(window.innerWidth - 240, rect.left)}px`;
    filterColTitle.textContent = `Filter: ${colIndexToName(c)}`;

    const grid = sheets[activeSheet];
    const uniqueVals = new Set();
    for (let r = 1; r < grid.length; r++) {
      if (grid[r] && grid[r][c]) {
        uniqueVals.add(String(grid[r][c].val !== undefined ? grid[r][c].val : ''));
      }
    }
    const sortedVals = Array.from(uniqueVals).sort((a, b) => a.localeCompare(b, 'ko-KR', { numeric: true }));
    const allowed = activeFilters[c] || new Set(sortedVals);

    let listHtml = `<label class="filter-item-label" style="font-weight:700;"><input type="checkbox" id="filter-select-all" ${allowed.size === sortedVals.length ? 'checked' : ''}> (Select All)</label>`;
    sortedVals.forEach(val => {
      const isChecked = allowed.has(val);
      listHtml += `<label class="filter-item-label"><input type="checkbox" class="filter-val-cb" value="${escapeHtml(val)}" ${isChecked ? 'checked' : ''}> <span>${escapeHtml(val || '(Blanks)')}</span></label>`;
    });
    filterItemsList.innerHTML = listHtml;
    filterSearchInput.value = '';
    filterDropdownMenu.classList.remove('hidden');

    const selectAllCb = document.getElementById('filter-select-all');
    if (selectAllCb) {
      selectAllCb.addEventListener('change', () => {
        const cbs = filterItemsList.querySelectorAll('.filter-val-cb');
        cbs.forEach(cb => cb.checked = selectAllCb.checked);
      });
    }
  });

  filterSearchInput.addEventListener('input', () => {
    const q = filterSearchInput.value.toLowerCase();
    const labels = filterItemsList.querySelectorAll('.filter-item-label');
    labels.forEach((lbl, idx) => {
      if (idx === 0) return;
      const txt = lbl.textContent.toLowerCase();
      lbl.style.display = txt.includes(q) ? 'flex' : 'none';
    });
  });

  btnFilterApply.addEventListener('click', () => {
    const selectedVals = new Set();
    filterItemsList.querySelectorAll('.filter-val-cb:checked').forEach(cb => {
      selectedVals.add(cb.value);
    });
    activeFilters[currentFilterCol] = selectedVals;
    filterDropdownMenu.classList.add('hidden');
    renderGrid();
  });

  btnFilterClear.addEventListener('click', () => {
    delete activeFilters[currentFilterCol];
    filterDropdownMenu.classList.add('hidden');
    renderGrid();
  });

  btnFilterSortAsc.addEventListener('click', () => {
    const grid = sheets[activeSheet];
    const col = currentFilterCol;
    const header = grid[0];
    const rows = grid.slice(1);
    rows.sort((a, b) => {
      const valA = a[col] ? a[col].val : '';
      const valB = b[col] ? b[col].val : '';
      return String(valA).localeCompare(String(valB), 'ko-KR', { numeric: true });
    });
    sheets[activeSheet] = [header, ...rows];
    filterDropdownMenu.classList.add('hidden');
    markDirty(true);
    renderGrid();
  });

  btnFilterSortDesc.addEventListener('click', () => {
    const grid = sheets[activeSheet];
    const col = currentFilterCol;
    const header = grid[0];
    const rows = grid.slice(1);
    rows.sort((a, b) => {
      const valA = a[col] ? a[col].val : '';
      const valB = b[col] ? b[col].val : '';
      return String(valB).localeCompare(String(valA), 'ko-KR', { numeric: true });
    });
    sheets[activeSheet] = [header, ...rows];
    filterDropdownMenu.classList.add('hidden');
    markDirty(true);
    renderGrid();
  });

  // --- FORMULAS RIBBON EXTENSIONS ---
  document.getElementById('fn-vlookup')?.addEventListener('click', () => {
    setCellValue(selectedCell.r, selectedCell.c, `=VLOOKUP(A${selectedCell.r + 1}, A1:C10, 2, FALSE)`);
  });
  document.getElementById('fn-xlookup')?.addEventListener('click', () => {
    setCellValue(selectedCell.r, selectedCell.c, `=XLOOKUP(A${selectedCell.r + 1}, A1:A10, B1:B10)`);
  });
  document.getElementById('fn-concat')?.addEventListener('click', () => {
    setCellValue(selectedCell.r, selectedCell.c, `=CONCAT(A${selectedCell.r + 1}, " ", B${selectedCell.r + 1})`);
  });
  document.getElementById('fn-today')?.addEventListener('click', () => {
    setCellValue(selectedCell.r, selectedCell.c, `=TODAY()`);
  });
  document.getElementById('insert-fn-sum')?.addEventListener('click', () => {
    document.getElementById('fn-autosum').click();
  });
  document.getElementById('insert-fn-avg')?.addEventListener('click', () => {
    document.getElementById('fn-average').click();
  });
  document.getElementById('insert-fn-count')?.addEventListener('click', () => {
    setCellValue(selectedCell.r, selectedCell.c, `=COUNT(${colIndexToName(selectedCell.c)}1:${colIndexToName(selectedCell.c)}${selectedCell.r})`);
  });
  document.getElementById('insert-fn-if')?.addEventListener('click', () => {
    setCellValue(selectedCell.r, selectedCell.c, `=IF(${colIndexToName(selectedCell.c)}1>100, "Pass", "Fail")`);
  });

  // --- FIND & REPLACE CONTROLLER ---
  function runFind() {
    const q = findInput.value.trim().toLowerCase();
    findMatches = [];
    findMatchIndex = -1;
    if (!q) {
      findMatchBadge.textContent = '0 / 0';
      return;
    }
    const grid = sheets[activeSheet];
    for (let r = 0; r < grid.length; r++) {
      for (let c = 0; c < grid[r].length; c++) {
        const val = String(grid[r][c].val || '').toLowerCase();
        const raw = String(grid[r][c].raw || '').toLowerCase();
        if (val.includes(q) || raw.includes(q)) {
          findMatches.push({ r, c });
        }
      }
    }
    if (findMatches.length > 0) {
      findMatchIndex = 0;
      findMatchBadge.textContent = `1 / ${findMatches.length}`;
      selectCell(findMatches[0].r, findMatches[0].c);
    } else {
      findMatchBadge.textContent = '0 / 0';
    }
  }

  btnFindNext.addEventListener('click', () => {
    if (findMatches.length === 0) return;
    findMatchIndex = (findMatchIndex + 1) % findMatches.length;
    findMatchBadge.textContent = `${findMatchIndex + 1} / ${findMatches.length}`;
    selectCell(findMatches[findMatchIndex].r, findMatches[findMatchIndex].c);
  });

  btnFindPrev.addEventListener('click', () => {
    if (findMatches.length === 0) return;
    findMatchIndex = (findMatchIndex - 1 + findMatches.length) % findMatches.length;
    findMatchBadge.textContent = `${findMatchIndex + 1} / ${findMatches.length}`;
    selectCell(findMatches[findMatchIndex].r, findMatches[findMatchIndex].c);
  });

  btnToggleReplace.addEventListener('click', () => {
    isReplaceMode = !isReplaceMode;
    replaceInput.style.display = isReplaceMode ? 'inline-block' : 'none';
    btnReplaceOne.style.display = isReplaceMode ? 'inline-block' : 'none';
    btnReplaceAll.style.display = isReplaceMode ? 'inline-block' : 'none';
  });

  btnReplaceOne.addEventListener('click', () => {
    if (findMatches.length === 0 || findMatchIndex < 0) return;
    const match = findMatches[findMatchIndex];
    const grid = sheets[activeSheet];
    const targetCell = grid[match.r][match.c];
    const q = findInput.value.trim();
    const rep = replaceInput.value;
    const regex = new RegExp(q.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'gi');
    targetCell.raw = String(targetCell.raw || targetCell.val || '').replace(regex, rep);
    markDirty(true);
    renderGrid();
    runFind();
  });

  btnReplaceAll.addEventListener('click', () => {
    if (findMatches.length === 0) return;
    const q = findInput.value.trim();
    const rep = replaceInput.value;
    const regex = new RegExp(q.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'gi');
    const grid = sheets[activeSheet];
    findMatches.forEach(m => {
      const cell = grid[m.r][m.c];
      cell.raw = String(cell.raw || cell.val || '').replace(regex, rep);
    });
    markDirty(true);
    renderGrid();
    runFind();
  });

  findInput.addEventListener('input', runFind);
  btnCloseFind.addEventListener('click', () => findReplaceBar.classList.add('hidden'));

  // --- FORMULA AUTOCOMPLETE & SIGNATURE HELPER ---
  formulaInput.addEventListener('input', () => {
    const val = formulaInput.value;
    if (!val.startsWith('=')) {
      formulaAutocomplete.classList.add('hidden');
      formulaSignatureHint.classList.add('hidden');
      return;
    }

    const openParenMatch = val.match(/([A-Z0-9_\.]+)\s*\([^\)]*$/i);
    if (openParenMatch) {
      const fnName = openParenMatch[1].toUpperCase();
      const fnInfo = FUNCTION_CATALOG.find(f => f.name === fnName);
      if (fnInfo) {
        formulaSignatureHint.textContent = `${fnInfo.sig} — ${fnInfo.desc}`;
        formulaSignatureHint.classList.remove('hidden');
        formulaAutocomplete.classList.add('hidden');
        return;
      }
    }

    if (!val.includes('(')) {
      formulaSignatureHint.classList.add('hidden');
      const searchWord = val.substring(1).trim().toUpperCase();
      if (searchWord.length > 0) {
        const matches = FUNCTION_CATALOG.filter(f => f.name.startsWith(searchWord)).slice(0, 8);
        if (matches.length > 0) {
          let html = '';
          matches.forEach((m, idx) => {
            html += `<div class="formula-suggest-item ${idx === 0 ? 'selected' : ''}" data-name="${m.name}">
              <span class="formula-suggest-name">${m.name}</span>
              <span class="formula-suggest-desc">${m.desc}</span>
            </div>`;
          });
          formulaAutocomplete.innerHTML = html;
          formulaAutocomplete.classList.remove('hidden');
          return;
        }
      }
    }
    formulaAutocomplete.classList.add('hidden');
    if (!val.includes('(')) {
      formulaSignatureHint.classList.add('hidden');
    }
  });

  formulaAutocomplete.addEventListener('click', (e) => {
    const item = e.target.closest('.formula-suggest-item');
    if (!item) return;
    const fn = item.getAttribute('data-name');
    formulaInput.value = `=${fn}(`;
    formulaAutocomplete.classList.add('hidden');
    formulaInput.focus();
  });

  // Gridlines & Formula Bar Toggle
  checkGridlines.addEventListener('change', () => {
    document.getElementById('grid-table').style.borderCollapse = checkGridlines.checked ? 'separate' : 'collapse';
  });

  checkFormulaBar.addEventListener('change', () => {
    formulaBarContainer.style.display = checkFormulaBar.checked ? 'flex' : 'none';
  });

  // Header Resizing Logic (Column & Row)
  theadEl.addEventListener('mousedown', (e) => {
    const resizer = e.target.closest('.col-resizer');
    if (resizer) {
      isResizingCol = true;
      resizeColIndex = parseInt(resizer.getAttribute('data-c'), 10);
      resizeStartX = e.clientX;
      const th = resizer.closest('th');
      resizeStartWidth = th ? th.offsetWidth : 100;
      e.stopPropagation();
      e.preventDefault();
    }
  });

  window.addEventListener('mousemove', (e) => {
    if (isResizingCol && resizeColIndex >= 0) {
      const diff = e.clientX - resizeStartX;
      const newWidth = Math.max(40, resizeStartWidth + diff);
      colWidths[resizeColIndex] = newWidth;
      const th = theadEl.querySelector(`th[data-c="${resizeColIndex}"]`);
      if (th) {
        th.style.width = `${newWidth}px`;
        th.style.minWidth = `${newWidth}px`;
      }
    }
    if (isResizingRow && resizeRowIndex >= 0) {
      const diff = e.clientY - resizeStartY;
      const newHeight = Math.max(20, resizeStartHeight + diff);
      rowHeights[resizeRowIndex] = newHeight;
      const th = tbodyEl.querySelector(`th[data-r="${resizeRowIndex}"]`);
      if (th) {
        th.style.height = `${newHeight}px`;
      }
    }
  });

  window.addEventListener('mouseup', () => {
    if (isResizingCol) {
      isResizingCol = false;
      resizeColIndex = -1;
      renderGrid();
    }
    if (isResizingRow) {
      isResizingRow = false;
      resizeRowIndex = -1;
      renderGrid();
    }
    if (isDraggingFill) {
      isDraggingFill = false;
      const grid = sheets[activeSheet];
      if (grid && fillTargetR > selectedCell.r) {
        const startR = selectedCell.r;
        const endR = fillTargetR;
        const col = selectedCell.c;
        const baseCell = grid[startR][col];
        const baseNum = Number(baseCell.val);
        const isNumeric = !isNaN(baseNum) && String(baseCell.raw).trim() === String(baseNum);

        for (let r = startR + 1; r <= endR; r++) {
          if (!grid[r]) continue;
          grid[r][col] = JSON.parse(JSON.stringify(baseCell));
          if (isNumeric) {
            const increment = r - startR;
            const newVal = baseNum + increment;
            grid[r][col].raw = String(newVal);
            grid[r][col].val = newVal;
          } else if (String(baseCell.raw).startsWith('=')) {
            const offset = r - startR;
            const adjusted = baseCell.raw.replace(/([A-Z]+)(\d+)/g, (_m, c, rowNum) => {
              return `${c}${parseInt(rowNum, 10) + offset}`;
            });
            grid[r][col].raw = adjusted;
          }
        }
        selectionRange.endR = endR;
        markDirty(true);
        renderGrid();
      }
    }
    isSelecting = false;
  });

  // Table Cell Interaction & Mouse Drag Range Selection
  tbodyEl.addEventListener('mousedown', (e) => {
    // Check if row resizer was clicked
    const rowResizer = e.target.closest('.row-resizer');
    if (rowResizer) {
      isResizingRow = true;
      resizeRowIndex = parseInt(rowResizer.getAttribute('data-r'), 10);
      resizeStartY = e.clientY;
      const th = rowResizer.closest('th');
      resizeStartHeight = th ? th.offsetHeight : 26;
      e.stopPropagation();
      e.preventDefault();
      return;
    }

    // Check if auto-fill handle was clicked
    if (e.target.id === 'fill-handle' || e.target.classList.contains('auto-fill-handle')) {
      isDraggingFill = true;
      fillTargetR = selectedCell.r;
      fillTargetC = selectedCell.c;
      e.stopPropagation();
      e.preventDefault();
      return;
    }

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
    if (isDraggingFill) {
      const cell = e.target.closest('.grid-cell');
      if (cell) {
        const r = parseInt(cell.getAttribute('data-r'), 10);
        if (r >= selectedCell.r) {
          fillTargetR = r;
          selectionRange.endR = r;
          selectionRange.endC = selectedCell.c;
          renderGrid();
        }
      }
      return;
    }

    if (!isSelecting) return;
    const cell = e.target.closest('.grid-cell');
    if (!cell) return;
    const r = parseInt(cell.getAttribute('data-r'), 10);
    const c = parseInt(cell.getAttribute('data-c'), 10);
    selectionRange.endR = r;
    selectionRange.endC = c;
    renderGrid();
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
    } else if (e.key === 'Escape') {
      const cell = sheets[activeSheet][selectedCell.r][selectedCell.c];
      formulaInput.value = cell.raw || '';
      formulaInput.blur();
    }
  });

  // Keyboard Navigation & Range Expansion
  window.addEventListener('keydown', (e) => {
    if (e.target === formulaInput) return;

    if ((e.ctrlKey || e.metaKey) && e.key === 's') {
      e.preventDefault();
      saveCurrentWorkbook();
      return;
    }
    if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'f') {
      e.preventDefault();
      findReplaceBar.classList.remove('hidden');
      findInput.focus();
      findInput.select();
      return;
    }
    if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'h') {
      e.preventDefault();
      findReplaceBar.classList.remove('hidden');
      isReplaceMode = true;
      replaceInput.style.display = 'inline-block';
      btnReplaceOne.style.display = 'inline-block';
      btnReplaceAll.style.display = 'inline-block';
      findInput.focus();
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

    // Shift + Arrow Range Selection Expansion
    if (e.shiftKey && (e.key === 'ArrowUp' || e.key === 'ArrowDown' || e.key === 'ArrowLeft' || e.key === 'ArrowRight')) {
      e.preventDefault();
      const grid = sheets[activeSheet];
      if (!grid) return;
      if (e.key === 'ArrowUp') {
        selectionRange.endR = Math.max(0, selectionRange.endR - 1);
      } else if (e.key === 'ArrowDown') {
        selectionRange.endR = Math.min(grid.length - 1, selectionRange.endR + 1);
      } else if (e.key === 'ArrowLeft') {
        selectionRange.endC = Math.max(0, selectionRange.endC - 1);
      } else if (e.key === 'ArrowRight') {
        selectionRange.endC = Math.min((grid[0] ? grid[0].length : 16) - 1, selectionRange.endC + 1);
      }
      renderGrid();
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
    } else if (e.key === 'F2') {
      e.preventDefault();
      formulaInput.focus();
      formulaInput.select();
    } else if (e.key.length === 1 && !e.ctrlKey && !e.metaKey && !e.altKey) {
      formulaInput.value = e.key;
      formulaInput.focus();
    }
  });

  // --- SHEET TAB MANAGEMENT (Add, Inline Rename, Context Menu) ---
  function renderSheetTabs() {
    sheetTabsEl.innerHTML = '';
    Object.keys(sheets).forEach(name => {
      const tab = document.createElement('button');
      tab.className = `sheet-tab ${name === activeSheet ? 'active' : ''}`;
      tab.textContent = name;
      tab.setAttribute('data-sheet', name);
      tab.addEventListener('click', () => switchSheet(name));
      sheetTabsEl.appendChild(tab);
    });
    sheetTabsEl.appendChild(btnAddSheet);
  }

  sheetTabsEl.addEventListener('dblclick', (e) => {
    const tab = e.target.closest('.sheet-tab');
    if (!tab) return;
    const oldName = tab.getAttribute('data-sheet');
    const input = document.createElement('input');
    input.type = 'text';
    input.className = 'sheet-tab-input';
    input.value = oldName;
    tab.innerHTML = '';
    tab.appendChild(input);
    input.focus();
    input.select();

    const finishRename = () => {
      const newName = input.value.trim();
      if (newName && newName !== oldName && !sheets[newName]) {
        sheets[newName] = sheets[oldName];
        delete sheets[oldName];
        if (activeSheet === oldName) activeSheet = newName;
        markDirty(true);
      }
      renderSheetTabs();
    };

    input.addEventListener('keydown', (ke) => {
      if (ke.key === 'Enter') finishRename();
      if (ke.key === 'Escape') renderSheetTabs();
    });
    input.addEventListener('blur', finishRename);
  });

  sheetTabsEl.addEventListener('contextmenu', (e) => {
    const tab = e.target.closest('.sheet-tab');
    if (!tab) return;
    e.preventDefault();
    targetContextMenuSheet = tab.getAttribute('data-sheet');
    sheetTabContextMenu.style.top = `${e.clientY - 90}px`;
    sheetTabContextMenu.style.left = `${e.clientX}px`;
    sheetTabContextMenu.classList.remove('hidden');
  });

  sheetTabContextMenu.addEventListener('click', (e) => {
    const action = e.target.getAttribute('data-action');
    if (!action || !targetContextMenuSheet) return;
    sheetTabContextMenu.classList.add('hidden');

    if (action === 'rename-sheet') {
      const tab = sheetTabsEl.querySelector(`.sheet-tab[data-sheet="${targetContextMenuSheet}"]`);
      if (tab) {
        tab.dispatchEvent(new MouseEvent('dblclick'));
      }
    } else if (action === 'duplicate-sheet') {
      const newName = `${targetContextMenuSheet} (Copy)`;
      sheets[newName] = JSON.parse(JSON.stringify(sheets[targetContextMenuSheet]));
      markDirty(true);
      switchSheet(newName);
    } else if (action === 'delete-sheet') {
      if (Object.keys(sheets).length <= 1) {
        vscode.postMessage({ type: 'gridcraft:error', text: 'A workbook must contain at least one visible worksheet.' });
        return;
      }
      delete sheets[targetContextMenuSheet];
      const remaining = Object.keys(sheets);
      activeSheet = remaining[0];
      markDirty(true);
      renderSheetTabs();
      renderGrid();
    }
  });

  window.addEventListener('click', (e) => {
    if (!sheetTabContextMenu.contains(e.target) && !e.target.closest('.sheet-tab')) {
      sheetTabContextMenu.classList.add('hidden');
    }
  });

  btnAddSheet.addEventListener('click', () => {
    const name = `Sheet${Object.keys(sheets).length + 1}`;
    sheets[name] = createInitialGrid(35, 16);
    markDirty(true);
    switchSheet(name);
  });

  function switchSheet(name) {
    activeSheet = name;
    renderSheetTabs();
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
        activeSheet = sheetNames[0] || 'Sheet1';
        renderSheetTabs();
        renderGrid();
        break;

      case 'gridcraft:openText':
        currentFileName = message.filename;
        docTitleEl.textContent = message.filename;
        const delimiter = message.delimiter || (message.filename.toLowerCase().endsWith('.tsv') ? '\t' : ',');
        const lines = (message.text || '').split(/\r?\n/);
        const parsedRows = [];
        lines.forEach(line => {
          if (!line.trim() && lines.length > 1) return;
          const row = [];
          let insideQuote = false;
          let entry = '';
          for (let i = 0; i < line.length; i++) {
            const ch = line[i];
            if (ch === '"') {
              if (insideQuote && line[i + 1] === '"') {
                entry += '"';
                i++;
              } else {
                insideQuote = !insideQuote;
              }
            } else if (ch === delimiter && !insideQuote) {
              row.push(entry);
              entry = '';
            } else {
              entry += ch;
            }
          }
          row.push(entry);
          parsedRows.push(row);
        });

        const totalRows = Math.max(parsedRows.length, 35);
        let maxCols = 16;
        parsedRows.forEach(r => maxCols = Math.max(maxCols, r.length));

        const textGrid = [];
        for (let r = 0; r < totalRows; r++) {
          const rowArr = [];
          for (let c = 0; c < maxCols; c++) {
            const model = createCellModel();
            if (parsedRows[r] && parsedRows[r][c] !== undefined) {
              const val = parsedRows[r][c];
              model.raw = val;
              model.val = val;
            }
            rowArr.push(model);
          }
          textGrid.push(rowArr);
        }

        sheets = { Sheet1: textGrid };
        activeSheet = 'Sheet1';
        renderSheetTabs();
        renderGrid();
        break;
    }
  });

  // Signal Ready to VS Code Host
  vscode.postMessage({ type: 'gridcraft:ready' });
})();
