import * as vscode from 'vscode';
import * as path from 'path';
import * as XLSX from 'xlsx';

/**
 * Custom Document for Open Excel
 */
export class ExcelDocument implements vscode.CustomDocument {
  private readonly _uri: vscode.Uri;
  private _isDirty = false;

  constructor(uri: vscode.Uri) {
    this._uri = uri;
  }

  get uri(): vscode.Uri {
    return this._uri;
  }

  get isDirty(): boolean {
    return this._isDirty;
  }

  setDirty(dirty: boolean) {
    this._isDirty = dirty;
  }

  dispose(): void {
    // cleanup
  }
}

/**
 * Custom Editor Provider for Open Excel (.xlsx, .xls, .csv, .tsv)
 */
export class ExcelEditorProvider implements vscode.CustomEditorProvider<ExcelDocument> {
  public static readonly viewType = 'gridcraft.excelEditor';

  private readonly _onDidChangeCustomDocument = new vscode.EventEmitter<vscode.CustomDocumentEditEvent<ExcelDocument>>();
  public readonly onDidChangeCustomDocument = this._onDidChangeCustomDocument.event;

  constructor(private readonly context: vscode.ExtensionContext) {}

  public static register(context: vscode.ExtensionContext): vscode.Disposable {
    const provider = new ExcelEditorProvider(context);
    const providerRegistration = vscode.window.registerCustomEditorProvider(
      ExcelEditorProvider.viewType,
      provider,
      {
        webviewOptions: {
          retainContextWhenHidden: true,
          enableFindWidget: true,
        },
        supportsMultipleEditorsPerDocument: false,
      }
    );
    return providerRegistration;
  }

  async openCustomDocument(
    uri: vscode.Uri,
    _openContext: vscode.CustomDocumentOpenContext,
    _token: vscode.CancellationToken
  ): Promise<ExcelDocument> {
    return new ExcelDocument(uri);
  }

  async resolveCustomEditor(
    document: ExcelDocument,
    webviewPanel: vscode.WebviewPanel,
    _token: vscode.CancellationToken
  ): Promise<void> {
    webviewPanel.webview.options = {
      enableScripts: true,
      localResourceRoots: [
        vscode.Uri.file(path.join(this.context.extensionPath, 'extension', 'web')),
        vscode.Uri.file(path.join(this.context.extensionPath, 'assets')),
        vscode.Uri.file(path.join(this.context.extensionPath, 'dist')),
      ],
    };

    webviewPanel.webview.html = this.getHtmlForWebview(webviewPanel.webview);

    // Initial load file data
    const loadFileData = async () => {
      try {
        const fileData = await vscode.workspace.fs.readFile(document.uri);
        const filename = path.basename(document.uri.fsPath);
        const lower = filename.toLowerCase();

        if (lower.endsWith('.csv') || lower.endsWith('.tsv')) {
          const text = new TextDecoder('utf-8').decode(fileData);
          webviewPanel.webview.postMessage({
            type: 'gridcraft:openText',
            filename,
            text,
            delimiter: lower.endsWith('.tsv') ? '\t' : ',',
          });
        } else {
          // Parse binary XLSX/XLS workbook
          const wb = XLSX.read(fileData, { type: 'buffer', cellFormula: true });
          const sheetsData: Record<string, any[][]> = {};

          for (const sheetName of wb.SheetNames) {
            const ws = wb.Sheets[sheetName];
            const range = XLSX.utils.decode_range(ws['!ref'] || 'A1:M30');
            const maxR = Math.max(range.e.r + 1, 30);
            const maxC = Math.max(range.e.c + 1, 15);

            const grid: any[][] = [];
            for (let r = 0; r < maxR; r++) {
              const row: any[] = [];
              for (let c = 0; c < maxC; c++) {
                const cellAddr = XLSX.utils.encode_cell({ r, c });
                const cell = ws[cellAddr];
                let raw = '';
                let val: any = '';
                if (cell) {
                  if (cell.f) {
                    raw = '=' + cell.f;
                    val = cell.v !== undefined ? cell.v : cell.w || '';
                  } else {
                    raw = cell.v !== undefined ? String(cell.v) : '';
                    val = cell.v !== undefined ? cell.v : '';
                  }
                }
                row.push({
                  raw,
                  val,
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
                  numFmt: typeof val === 'number' ? 'Number' : 'General',
                  decimals: null,
                  merge: null,
                  mergedInto: null,
                });
              }
              grid.push(row);
            }

            // Decode XLSX merges
            if (ws['!merges']) {
              for (const m of ws['!merges']) {
                const startR = m.s.r;
                const startC = m.s.c;
                const rowSpan = m.e.r - m.s.r + 1;
                const colSpan = m.e.c - m.s.c + 1;
                if (grid[startR] && grid[startR][startC]) {
                  grid[startR][startC].merge = { rowspan: rowSpan, colspan: colSpan };
                }
                for (let r = m.s.r; r <= m.e.r; r++) {
                  for (let c = m.s.c; c <= m.e.c; c++) {
                    if (r === startR && c === startC) continue;
                    if (grid[r] && grid[r][c]) {
                      grid[r][c].mergedInto = { r: startR, c: startC };
                    }
                  }
                }
              }
            }

            sheetsData[sheetName] = grid;
          }

          webviewPanel.webview.postMessage({
            type: 'gridcraft:openWorkbook',
            filename,
            sheets: sheetsData,
            sheetNames: wb.SheetNames,
          });
        }
      } catch (err: any) {
        vscode.window.showErrorMessage(`Failed to read spreadsheet file: ${err?.message || err}`);
      }
    };

    // Handle messages from Webview
    webviewPanel.webview.onDidReceiveMessage(async (message) => {
      switch (message.type) {
        case 'gridcraft:ready':
          await loadFileData();
          break;

        case 'gridcraft:dirty':
          document.setDirty(message.dirty);
          this._onDidChangeCustomDocument.fire({
            document,
            undo: () => {},
            redo: () => {},
          });
          break;

        case 'gridcraft:saveWorkbook':
          try {
            const filename = path.basename(document.uri.fsPath);
            const lower = filename.toLowerCase();

            if (lower.endsWith('.csv') || lower.endsWith('.tsv')) {
              const csvText = message.csvText;
              const uint8 = new TextEncoder().encode(csvText);
              await vscode.workspace.fs.writeFile(document.uri, uint8);
            } else {
              // Build standard binary XLSX
              const wb = XLSX.utils.book_new();
              const sheets = message.sheets || {};

              for (const [sheetName, grid] of Object.entries(sheets)) {
                const ws: XLSX.WorkSheet = {};
                const merges: XLSX.Range[] = [];
                let maxR = 0;
                let maxC = 0;

                (grid as any[][]).forEach((row, r) => {
                  row.forEach((cell, c) => {
                    if (cell) {
                      if (cell.merge) {
                        merges.push({
                          s: { r, c },
                          e: { r: r + cell.merge.rowspan - 1, c: c + cell.merge.colspan - 1 }
                        });
                      }
                      if (cell.raw || cell.val) {
                        maxR = Math.max(maxR, r);
                        maxC = Math.max(maxC, c);
                        const addr = XLSX.utils.encode_cell({ r, c });
                        if (String(cell.raw).startsWith('=')) {
                          ws[addr] = {
                            t: typeof cell.val === 'number' ? 'n' : 's',
                            f: String(cell.raw).substring(1),
                            v: cell.val,
                          };
                        } else {
                          const num = Number(cell.raw);
                          if (!isNaN(num) && cell.raw !== '') {
                            ws[addr] = { t: 'n', v: num };
                          } else {
                            ws[addr] = { t: 's', v: String(cell.raw) };
                          }
                        }
                      }
                    }
                  });
                });

                if (merges.length > 0) {
                  ws['!merges'] = merges;
                }

                ws['!ref'] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: Math.max(maxR, 5), c: Math.max(maxC, 5) } });
                XLSX.utils.book_append_sheet(wb, ws, sheetName);
              }

              const buffer = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
              await vscode.workspace.fs.writeFile(document.uri, buffer);
            }

            document.setDirty(false);
            vscode.window.setStatusBarMessage(`Spreadsheet saved: ${filename}`, 3000);
          } catch (err: any) {
            vscode.window.showErrorMessage(`Failed to save spreadsheet: ${err?.message || err}`);
          }
          break;

        case 'gridcraft:openNative':
          vscode.commands.executeCommand('openExcell.openInNative', document.uri);
          break;

        case 'gridcraft:info':
          vscode.window.showInformationMessage(message.text);
          break;

        case 'gridcraft:error':
          vscode.window.showErrorMessage(message.text);
          break;
      }
    });

    // Auto reload when file changes externally
    const fileWatcher = vscode.workspace.createFileSystemWatcher(
      new vscode.RelativePattern(path.dirname(document.uri.fsPath), path.basename(document.uri.fsPath))
    );
    fileWatcher.onDidChange(async () => {
      if (!document.isDirty) {
        await loadFileData();
      }
    });

    webviewPanel.onDidDispose(() => {
      fileWatcher.dispose();
      document.dispose();
    });
  }

  async saveCustomDocument(
    document: ExcelDocument,
    cancellation: vscode.CancellationToken
  ): Promise<void> {
    // Trigger webview to save
    // Webview will respond with gridcraft:save
  }

  async saveCustomDocumentAs(
    document: ExcelDocument,
    destination: vscode.Uri,
    cancellation: vscode.CancellationToken
  ): Promise<void> {
    const data = await vscode.workspace.fs.readFile(document.uri);
    await vscode.workspace.fs.writeFile(destination, data);
  }

  async revertCustomDocument(
    document: ExcelDocument,
    cancellation: vscode.CancellationToken
  ): Promise<void> {
    // Revert logic
  }

  async backupCustomDocument(
    document: ExcelDocument,
    context: vscode.CustomDocumentBackupContext,
    cancellation: vscode.CancellationToken
  ): Promise<vscode.CustomDocumentBackup> {
    return {
      id: context.destination.toString(),
      delete: () => {
        vscode.workspace.fs.delete(context.destination);
      },
    };
  }

  private getHtmlForWebview(webview: vscode.Webview): string {
    const webPath = path.join(this.context.extensionPath, 'extension', 'web');
    const scriptUri = webview.asWebviewUri(vscode.Uri.file(path.join(webPath, 'grid.js')));
    const styleUri = webview.asWebviewUri(vscode.Uri.file(path.join(webPath, 'grid.css')));
    const iconUri = webview.asWebviewUri(vscode.Uri.file(path.join(this.context.extensionPath, 'assets', 'app-icon', 'gridcraft.ico')));

    const nonce = getNonce();

    return `<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <meta http-equiv="Content-Security-Policy" content="default-src 'none'; img-src ${webview.cspSource} data: blob:; style-src ${webview.cspSource} 'unsafe-inline'; script-src 'nonce-${nonce}' ${webview.cspSource};">
  <title>Open Excel</title>
  <link rel="stylesheet" href="${styleUri}">
</head>
<body class="gridcraft-root">
  <!-- Ribbon Toolbar Container -->
  <header class="ribbon-bar">
    <!-- Top Branding & Tabs Bar -->
    <div class="ribbon-tab-bar">
      <div class="ribbon-brand">
        <span class="brand-badge">EXCEL</span>
        <span class="file-name" id="doc-title">Workbook.xlsx</span>
      </div>
      <nav class="ribbon-nav-tabs">
        <button class="nav-tab active" data-tab="tab-home">Home</button>
        <button class="nav-tab" data-tab="tab-insert">Insert</button>
        <button class="nav-tab" data-tab="tab-formulas">Formulas</button>
        <button class="nav-tab" data-tab="tab-data">Data</button>
        <button class="nav-tab" data-tab="tab-view">View</button>
      </nav>
      <div class="ribbon-actions">
        <button id="btn-save" class="ribbon-action-btn primary" title="Save file (Ctrl+S / Cmd+S)">
          <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M19 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11l5 5v11a2 2 0 0 1-2 2z"></path><polyline points="17 21 17 13 7 13 7 21"></polyline><polyline points="7 3 7 8 15 8"></polyline></svg>
          Save
        </button>
        <button id="btn-open-native" class="ribbon-action-btn secondary" title="Open in Native GridCraft App">
          <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M18 13v6a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h6"></path><polyline points="15 3 21 3 21 9"></polyline><line x1="10" y1="14" x2="21" y2="3"></line></svg>
          Native App
        </button>
      </div>
    </div>
    
    <!-- Tab Contents Panels -->
    <div class="ribbon-panels">
      <!-- HOME PANEL -->
      <div class="tab-panel active" id="tab-home">
        <!-- Clipboard -->
        <div class="tool-section">
          <div class="tool-group">
            <button id="tool-cut" class="tool-btn icon-btn" title="Cut (Ctrl+X)"><svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="6" cy="6" r="3"></circle><circle cx="6" cy="18" r="3"></circle><line x1="20" y1="4" x2="8.12" y2="15.88"></line><line x1="14.47" y1="14.48" x2="20" y2="20"></line><line x1="8.12" y1="8.12" x2="12" y2="12"></line></svg></button>
            <button id="tool-copy" class="tool-btn icon-btn" title="Copy (Ctrl+C)"><svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="9" y="9" width="13" height="13" rx="2" ry="2"></rect><path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1"></path></svg></button>
            <button id="tool-paste" class="tool-btn icon-btn" title="Paste (Ctrl+V)"><svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M16 4h2a2 2 0 0 1 2 2v14a2 2 0 0 1-2 2H6a2 2 0 0 1-2-2V6a2 2 0 0 1 2-2h2"></path><rect x="8" y="2" width="8" height="4" rx="1" ry="1"></rect></svg></button>
          </div>
          <span class="section-label">Clipboard</span>
        </div>
        <div class="tool-divider"></div>

        <!-- Font Controls -->
        <div class="tool-section">
          <div class="tool-group flex-wrap">
            <select id="select-font-family" class="tool-select font-family" title="Font Family">
              <option value="Arial">Arial</option>
              <option value="Calibri" selected>Calibri</option>
              <option value="Segoe UI">Segoe UI</option>
              <option value="Apple SD Gothic Neo">맑은 고딕 / Apple SD</option>
              <option value="Times New Roman">Times New Roman</option>
              <option value="monospace">Consolas / Monospace</option>
            </select>
            <select id="select-font-size" class="tool-select font-size" title="Font Size">
              <option value="9">9</option>
              <option value="10">10</option>
              <option value="11" selected>11</option>
              <option value="12">12</option>
              <option value="14">14</option>
              <option value="16">16</option>
              <option value="18">18</option>
              <option value="24">24</option>
            </select>
            <div class="button-row">
              <button id="tool-bold" class="tool-btn" title="Bold (Ctrl+B)"><b>B</b></button>
              <button id="tool-italic" class="tool-btn" title="Italic (Ctrl+I)"><i>I</i></button>
              <button id="tool-underline" class="tool-btn" title="Underline (Ctrl+U)"><u>U</u></button>
              <button id="tool-strike" class="tool-btn" title="Strikethrough"><s>S</s></button>
              <label class="color-picker-label" title="Font Color">
                <span class="color-indicator">A</span>
                <input type="color" id="tool-font-color" value="#000000">
              </label>
              <label class="color-picker-label" title="Fill Color">
                <span class="fill-indicator">🪣</span>
                <input type="color" id="tool-fill-color" value="#ffffff">
              </label>
              <select id="select-borders" class="tool-select border-select" title="Borders">
                <option value="none">Border: None</option>
                <option value="all">All Borders</option>
                <option value="outside">Outside Border</option>
                <option value="thick">Thick Box</option>
                <option value="bottom">Bottom Border</option>
              </select>
            </div>
          </div>
          <span class="section-label">Font</span>
        </div>
        <div class="tool-divider"></div>

        <!-- Alignment Controls -->
        <div class="tool-section">
          <div class="tool-group flex-wrap">
            <div class="button-row">
              <button id="tool-align-top" class="tool-btn" title="Top Align">⇪</button>
              <button id="tool-align-middle" class="tool-btn" title="Middle Align">⬌</button>
              <button id="tool-align-bottom" class="tool-btn" title="Bottom Align">⇩</button>
              <div class="tool-mini-divider"></div>
              <button id="tool-align-left" class="tool-btn" title="Align Left">⇤</button>
              <button id="tool-align-center" class="tool-btn" title="Align Center">≡</button>
              <button id="tool-align-right" class="tool-btn" title="Align Right">⇥</button>
            </div>
            <div class="button-row">
              <button id="tool-wrap-text" class="tool-btn text-btn" title="Wrap Text">Wrap Text</button>
              <button id="tool-merge-cells" class="tool-btn text-btn" title="Merge & Center">Merge & Center</button>
            </div>
          </div>
          <span class="section-label">Alignment</span>
        </div>
        <div class="tool-divider"></div>

        <!-- Number Format Controls -->
        <div class="tool-section">
          <div class="tool-group flex-wrap">
            <select id="select-num-fmt" class="tool-select format-select" title="Number Format">
              <option value="General">General</option>
              <option value="Number">Number (1,234.00)</option>
              <option value="Currency">Currency ($ / ₩)</option>
              <option value="Accounting">Accounting</option>
              <option value="ShortDate">Short Date (YYYY-MM-DD)</option>
              <option value="Percentage">Percentage (%)</option>
            </select>
            <div class="button-row">
              <button id="tool-fmt-comma" class="tool-btn" title="Comma Style">,</button>
              <button id="tool-fmt-percent" class="tool-btn" title="Percent Style">%</button>
              <button id="tool-fmt-currency" class="tool-btn" title="Currency Style">₩</button>
              <button id="tool-fmt-inc-dec" class="tool-btn" title="Increase Decimal">.00</button>
              <button id="tool-fmt-dec-dec" class="tool-btn" title="Decrease Decimal">.0</button>
            </div>
          </div>
          <span class="section-label">Number</span>
        </div>
        <div class="tool-divider"></div>

        <!-- Cells & Editing -->
        <div class="tool-section">
          <div class="tool-group">
            <button id="tool-add-row" class="tool-btn text-btn">+ Row</button>
            <button id="tool-del-row" class="tool-btn text-btn">- Row</button>
            <button id="tool-add-col" class="tool-btn text-btn">+ Col</button>
            <button id="tool-del-col" class="tool-btn text-btn">- Col</button>
            <button id="tool-clear" class="tool-btn text-btn">Clear</button>
          </div>
          <span class="section-label">Cells</span>
        </div>
      </div>

      <!-- INSERT PANEL -->
      <div class="tab-panel" id="tab-insert">
        <div class="tool-section">
          <div class="tool-group">
            <button id="insert-chart-column" class="tool-btn text-btn">📊 Column Chart</button>
            <button id="insert-chart-pie" class="tool-btn text-btn">🥧 Pie Chart</button>
            <button id="insert-chart-line" class="tool-btn text-btn">📈 Line Chart</button>
          </div>
          <span class="section-label">Charts</span>
        </div>
        <div class="tool-divider"></div>
        <div class="tool-section">
          <div class="tool-group">
            <button id="insert-fn-sum" class="tool-btn text-btn">=SUM()</button>
            <button id="insert-fn-avg" class="tool-btn text-btn">=AVERAGE()</button>
            <button id="insert-fn-count" class="tool-btn text-btn">=COUNT()</button>
            <button id="insert-fn-if" class="tool-btn text-btn">=IF()</button>
          </div>
          <span class="section-label">Functions</span>
        </div>
      </div>

      <!-- FORMULAS PANEL -->
      <div class="tab-panel" id="tab-formulas">
        <div class="tool-section">
          <div class="tool-group">
            <button id="fn-autosum" class="tool-btn text-btn font-bold">∑ AutoSum</button>
            <button id="fn-average" class="tool-btn text-btn">Average</button>
            <button id="fn-count" class="tool-btn text-btn">Count Numbers</button>
            <button id="fn-max" class="tool-btn text-btn">Max</button>
            <button id="fn-min" class="tool-btn text-btn">Min</button>
          </div>
          <span class="section-label">AutoSum</span>
        </div>
        <div class="tool-divider"></div>
        <div class="tool-section">
          <div class="tool-group">
            <button id="fn-vlookup" class="tool-btn text-btn">VLOOKUP</button>
            <button id="fn-xlookup" class="tool-btn text-btn">XLOOKUP</button>
            <button id="fn-concat" class="tool-btn text-btn">CONCAT</button>
            <button id="fn-today" class="tool-btn text-btn">TODAY</button>
          </div>
          <span class="section-label">Library</span>
        </div>
      </div>

      <!-- DATA PANEL -->
      <div class="tab-panel" id="tab-data">
        <div class="tool-section">
          <div class="tool-group">
            <button id="tool-sort-asc" class="tool-btn text-btn">Sort A→Z</button>
            <button id="tool-sort-desc" class="tool-btn text-btn">Sort Z→A</button>
            <button id="tool-toggle-filter" class="tool-btn text-btn">Filter</button>
          </div>
          <span class="section-label">Sort & Filter</span>
        </div>
        <div class="tool-divider"></div>
        <div class="tool-section">
          <div class="tool-group">
            <button id="tool-export-csv" class="tool-btn text-btn">Export CSV</button>
          </div>
          <span class="section-label">Export</span>
        </div>
      </div>

      <!-- VIEW PANEL -->
      <div class="tab-panel" id="tab-view">
        <div class="tool-section">
          <div class="tool-group">
            <label class="view-checkbox-label">
              <input type="checkbox" id="check-gridlines" checked> Gridlines
            </label>
            <label class="view-checkbox-label">
              <input type="checkbox" id="check-formula-bar" checked> Formula Bar
            </label>
          </div>
          <span class="section-label">Show</span>
        </div>
      </div>
    </div>

    <!-- Excel Formula Bar -->
    <div class="formula-bar" id="formula-bar-container">
      <div class="cell-name-box" id="cell-name">A1</div>
      <div class="formula-fx">fx</div>
      <input type="text" class="formula-input" id="formula-input" placeholder="Enter value or formula like =SUM(A1:A5)" spellcheck="false" autocomplete="off" />
    </div>
  </header>

  <!-- Spreadsheet Canvas Area -->
  <main class="grid-container" id="grid-container">
    <div class="grid-scroll" id="grid-scroll">
      <table class="grid-table" id="grid-table">
        <thead id="grid-thead"></thead>
        <tbody id="grid-tbody"></tbody>
      </table>
    </div>

    <!-- Floating Chart Container (Optional popup) -->
    <div class="chart-overlay hidden" id="chart-overlay">
      <div class="chart-header">
        <span class="chart-title" id="chart-title">Data Chart</span>
        <button id="btn-close-chart" class="chart-close">✕</button>
      </div>
      <canvas id="chart-canvas" width="480" height="260"></canvas>
    </div>
  </main>

  <!-- Sheet Tabs & Status Bar -->
  <footer class="sheet-footer">
    <div class="sheet-tabs" id="sheet-tabs">
      <button class="sheet-tab active" data-sheet="Sheet1">Sheet1</button>
      <button class="sheet-add-btn" id="btn-add-sheet" title="Add Sheet">+</button>
    </div>
    <div class="status-stats" id="status-stats">
      <span id="stat-count">Count: 0</span>
      <span id="stat-sum">Sum: 0</span>
      <span id="stat-avg">Average: 0</span>
      <span id="stat-min">Min: 0</span>
      <span id="stat-max">Max: 0</span>
    </div>
  </footer>

  <!-- Context Menu -->
  <div class="context-menu hidden" id="context-menu">
    <div class="context-item" data-action="cut">Cut (Ctrl+X)</div>
    <div class="context-item" data-action="copy">Copy (Ctrl+C)</div>
    <div class="context-item" data-action="paste">Paste (Ctrl+V)</div>
    <div class="context-divider"></div>
    <div class="context-item" data-action="insert-row">Insert Row</div>
    <div class="context-item" data-action="delete-row">Delete Row</div>
    <div class="context-item" data-action="insert-col">Insert Column</div>
    <div class="context-item" data-action="delete-col">Delete Column</div>
    <div class="context-divider"></div>
    <div class="context-item" data-action="clear">Clear Contents</div>
  </div>

  <script nonce="${nonce}" src="${scriptUri}"></script>
</body>
</html>`;
  }
}

function getNonce(): string {
  let text = '';
  const possible = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';
  for (let i = 0; i < 32; i++) {
    text += possible.charAt(Math.floor(Math.random() * possible.length));
  }
  return text;
}
