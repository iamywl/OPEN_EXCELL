import * as vscode from 'vscode';
import * as path from 'path';
import * as fs from 'fs';
import * as cp from 'child_process';
import { ExcelEditorProvider } from './editorProvider';

export function activate(context: vscode.ExtensionContext) {
  // Register the custom editor provider for .xlsx, .xls, .csv, .tsv
  context.subscriptions.push(ExcelEditorProvider.register(context));

  // Command: Open file in Native GridCraft App (Cross-platform: Windows & macOS)
  context.subscriptions.push(
    vscode.commands.registerCommand('openExcell.openInNative', async (uri?: vscode.Uri) => {
      let targetUri = uri;
      if (!targetUri) {
        const activeTab = vscode.window.tabGroups.activeTabGroup?.activeTab;
        if (activeTab?.input instanceof vscode.TabInputCustom) {
          targetUri = activeTab.input.uri;
        } else if (vscode.window.activeTextEditor) {
          targetUri = vscode.window.activeTextEditor.document.uri;
        }
      }

      const filePath = targetUri?.fsPath;

      // Locate Native Binary (macOS / Windows cross-platform)
      const config = vscode.workspace.getConfiguration('openExcell');
      let binaryPath = config.get<string>('nativeBinaryPath');

      const isWindows = process.platform === 'win32';
      const binaryName = isWindows ? 'gridcraft.exe' : 'gridcraft';

      const candidatePaths = [
        binaryPath,
        path.join(context.extensionPath, 'target', 'release', binaryName),
        path.join(context.extensionPath, 'target', 'debug', binaryName),
        path.join(context.extensionPath, 'bin', binaryName),
      ].filter((p): p is string => Boolean(p && fs.existsSync(p)));

      const nativeBin = candidatePaths[0];

      if (nativeBin) {
        const args = filePath ? [filePath] : [];
        try {
          const child = cp.spawn(nativeBin, args, {
            detached: true,
            stdio: 'ignore',
            windowsHide: false,
          });
          child.unref();
          vscode.window.showInformationMessage(`Launched Native GridCraft${filePath ? ` with ${path.basename(filePath)}` : ''}`);
        } catch (err: any) {
          vscode.window.showErrorMessage(`Failed to launch native app: ${err?.message || err}`);
        }
      } else {
        // Fallback: If binary not built yet, offer to build or run with cargo
        const selection = await vscode.window.showInformationMessage(
          'Native GridCraft binary not found. Would you like to build and launch it now with Cargo?',
          'Run via Cargo',
          'Build Native App'
        );

        const terminal = vscode.window.createTerminal('GridCraft Native');
        terminal.show();

        const fileArg = filePath ? ` -- "${filePath}"` : '';
        if (selection === 'Run via Cargo') {
          terminal.sendText(`cargo run --release -p gridcraft${fileArg}`);
        } else if (selection === 'Build Native App') {
          terminal.sendText('cargo build --release -p gridcraft');
        }
      }
    })
  );

  // Command: Create New Spreadsheet
  context.subscriptions.push(
    vscode.commands.registerCommand('openExcell.newWorkbook', async () => {
      const workspaceFolder = vscode.workspace.workspaceFolders?.[0];
      const defaultUri = workspaceFolder ? workspaceFolder.uri : vscode.Uri.file(context.extensionPath);
      const saveUri = await vscode.window.showSaveDialog({
        defaultUri: vscode.Uri.joinPath(defaultUri, 'NewWorkbook.xlsx'),
        filters: {
          'Excel Workbook (*.xlsx)': ['xlsx'],
          'CSV File (*.csv)': ['csv'],
        },
      });

      if (saveUri) {
        // Write initial empty workbook bytes or basic CSV header
        const isCsv = saveUri.fsPath.endsWith('.csv');
        const emptyContent = isCsv
          ? Buffer.from('Column 1,Column 2,Column 3\n\n')
          : Buffer.from([]); // Empty will be initialized with a blank grid
        await vscode.workspace.fs.writeFile(saveUri, emptyContent);
        await vscode.commands.executeCommand('vscode.openWith', saveUri, ExcelEditorProvider.viewType);
      }
    })
  );
}

export function deactivate() {}
