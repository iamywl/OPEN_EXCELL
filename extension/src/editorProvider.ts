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
                let val = '';
                if (cell) {
                  if (cell.f) {
                    raw = '=' + cell.f;
                    val = cell.v !== undefined ? cell.v : cell.w || '';
                  } else {
                    raw = cell.v !== undefined ? String(cell.v) : '';
                    val = raw;
                  }
                }
                row.push({ raw, val, bold: false, italic: false, align: 'left' });
              }
              grid.push(row);
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
                let maxR = 0;
                let maxC = 0;

                (grid as any[][]).forEach((row, r) => {
                  row.forEach((cell, c) => {
                    if (cell && (cell.raw || cell.val)) {
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
                  });
                });

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
  <!-- Ribbon Toolbar -->
  <header class="ribbon-bar">
    <div class="ribbon-tab-bar">
      <div class="ribbon-brand">
        <span class="brand-badge">GLASS EXCEL</span>
        <span class="file-name" id="doc-title">Workbook.xlsx</span>
      </div>
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
    
    <div class="ribbon-tools">
      <div class="tool-group">
        <button id="tool-bold" class="tool-btn" title="Bold (Ctrl+B)"><b>B</b></button>
        <button id="tool-italic" class="tool-btn" title="Italic (Ctrl+I)"><i>I</i></button>
        <button id="tool-align-left" class="tool-btn" title="Align Left">⇤</button>
        <button id="tool-align-center" class="tool-btn" title="Align Center">≡</button>
        <button id="tool-align-right" class="tool-btn" title="Align Right">⇥</button>
      </div>
      <div class="tool-divider"></div>
      <div class="tool-group">
        <button id="tool-add-row" class="tool-btn text-btn">+ Row</button>
        <button id="tool-add-col" class="tool-btn text-btn">+ Col</button>
        <button id="tool-sort-asc" class="tool-btn text-btn">A→Z</button>
        <button id="tool-sort-desc" class="tool-btn text-btn">Z→A</button>
      </div>
      <div class="tool-divider"></div>
      <div class="tool-group">
        <button id="tool-export-csv" class="tool-btn text-btn">Export CSV</button>
      </div>
    </div>

    <!-- Formula Bar -->
    <div class="formula-bar">
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
    </div>
  </footer>

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
