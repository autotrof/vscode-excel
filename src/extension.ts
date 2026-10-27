import * as vscode from 'vscode';
import { randomBytes } from 'node:crypto';
import { parseWorkbook, WorkbookView } from './workbook';

class ExcelDocument implements vscode.CustomDocument {
  constructor(readonly uri: vscode.Uri, readonly workbook: WorkbookView) {}
  dispose(): void { /* no file handles */ }
}

class ExcelProvider implements vscode.CustomReadonlyEditorProvider<ExcelDocument> {
  async openCustomDocument(uri: vscode.Uri): Promise<ExcelDocument> {
    const data = await vscode.workspace.fs.readFile(uri);
    return new ExcelDocument(uri, await parseWorkbook(data, uri.path));
  }

  async resolveCustomEditor(document: ExcelDocument, panel: vscode.WebviewPanel): Promise<void> {
    const webview = panel.webview;
    webview.options = { enableScripts: true };
    const nonce = randomBytes(16).toString('hex');
    webview.html = html(nonce);
    const send = (sheet: number, start: number) => {
      try {
        webview.postMessage({ type: 'page', names: document.workbook.names, data: document.workbook.page(sheet, start) });
      } catch (error) {
        webview.postMessage({ type: 'error', message: error instanceof Error ? error.message : String(error) });
      }
    };
    const listener = webview.onDidReceiveMessage(message => {
      if (message?.type === 'ready') send(0, 1);
      if (message?.type === 'page' && Number.isInteger(message.sheet) && Number.isInteger(message.start)) send(message.sheet, message.start);
    });
    panel.onDidDispose(() => listener.dispose());
  }
}

export function activate(context: vscode.ExtensionContext): void {
  context.subscriptions.push(vscode.window.registerCustomEditorProvider('excelViewer.viewer', new ExcelProvider(), {
    webviewOptions: { retainContextWhenHidden: true }
  }));
}

function html(nonce: string): string {
  return `<!DOCTYPE html>
<html lang="en"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<meta http-equiv="Content-Security-Policy" content="default-src 'none'; style-src 'nonce-${nonce}'; script-src 'nonce-${nonce}';">
<style nonce="${nonce}">
  body { margin: 0; color: var(--vscode-editor-foreground); background: var(--vscode-editor-background); font-family: var(--vscode-font-family); }
  header { position: sticky; top: 0; z-index: 2; background: var(--vscode-editor-background); border-bottom: 1px solid var(--vscode-panel-border); padding: 8px 12px; }
  .bar { display: flex; align-items: center; gap: 8px; flex-wrap: wrap; }
  button { color: var(--vscode-button-foreground); background: var(--vscode-button-background); border: 0; border-radius: 3px; padding: 5px 9px; cursor: pointer; }
  button:hover { background: var(--vscode-button-hoverBackground); }
  button:disabled { opacity: .5; cursor: default; }
  button.active { outline: 2px solid var(--vscode-focusBorder); }
  #tabs { display: flex; gap: 5px; overflow-x: auto; padding-bottom: 4px; }
  #status { opacity: .8; font-size: 12px; }
  .viewport { overflow: auto; height: calc(100vh - 88px); }
  table { border-collapse: separate; border-spacing: 0; font-size: 13px; }
  th, td { border-right: 1px solid var(--vscode-panel-border); border-bottom: 1px solid var(--vscode-panel-border); min-width: 95px; max-width: 320px; height: 24px; padding: 3px 7px; white-space: pre; overflow: hidden; text-overflow: ellipsis; }
  th { background: var(--vscode-editorGroupHeader-tabsBackground); position: sticky; top: 0; z-index: 1; font-weight: 600; }
  th:first-child { left: 0; z-index: 2; min-width: 45px; text-align: right; }
  td:first-child { position: sticky; left: 0; z-index: 1; min-width: 45px; text-align: right; background: var(--vscode-editorGroupHeader-tabsBackground); }
  #message { padding: 20px; }
</style></head><body><header><div id="tabs"></div><div class="bar">
<button id="prev" aria-label="Previous page">Previous</button><button id="next" aria-label="Next page">Next</button><span id="status"></span></div></header>
<div class="viewport"><table id="grid" aria-label="Spreadsheet data"></table><div id="message">Loading workbook…</div></div>
<script nonce="${nonce}">
  const vscode = acquireVsCodeApi();
  const tabs = document.getElementById('tabs'), grid = document.getElementById('grid');
  const status = document.getElementById('status'), message = document.getElementById('message');
  const prev = document.getElementById('prev'), next = document.getElementById('next');
  let current = null;
  const request = (sheet, start) => vscode.postMessage({ type: 'page', sheet, start });
  const columnName = index => { let result = ''; for (index++; index; index = Math.floor((index - 1) / 26)) result = String.fromCharCode(65 + (index - 1) % 26) + result; return result; };
  prev.addEventListener('click', () => current && request(current.sheet, Math.max(1, current.start - 100)));
  next.addEventListener('click', () => current && request(current.sheet, current.end + 1));
  window.addEventListener('message', event => {
    const payload = event.data;
    if (payload.type === 'error') { message.textContent = payload.message; message.hidden = false; return; }
    if (payload.type !== 'page') return;
    const data = current = payload.data;
    tabs.replaceChildren();
    payload.names.forEach((name, index) => {
      const button = document.createElement('button'); button.textContent = name; button.title = name;
      button.classList.toggle('active', index === data.sheet);
      button.addEventListener('click', () => request(index, 1)); tabs.append(button);
    });
    grid.replaceChildren();
    const heading = document.createElement('tr'); heading.append(document.createElement('th'));
    for (let col = 0; col < data.columns; col++) { const th = document.createElement('th'); th.textContent = columnName(col); heading.append(th); }
    grid.append(heading);
    data.rows.forEach((row, index) => {
      const tr = document.createElement('tr'), label = document.createElement('td'); label.textContent = String(data.start + index); tr.append(label);
      row.forEach(cell => {
        const td = document.createElement('td'); td.textContent = cell.text; td.title = cell.text;
        if (cell.background) td.style.backgroundColor = cell.background;
        if (cell.foreground) td.style.color = cell.foreground;
        if (cell.bold) td.style.fontWeight = 'bold';
        if (cell.italic) td.style.fontStyle = 'italic';
        if (cell.align) td.style.textAlign = cell.align;
        tr.append(td);
      }); grid.append(tr);
    });
    prev.disabled = data.start <= 1; next.disabled = data.end >= data.totalRows;
    status.textContent = data.totalRows ? 'Rows ' + data.start + '–' + data.end + ' of ' + data.totalRows + ' · ' + data.columns + ' columns' : 'Empty sheet';
    message.hidden = data.totalRows > 0;
    if (!data.totalRows) message.textContent = 'This sheet is empty.';
  });
  vscode.postMessage({ type: 'ready' });
</script></body></html>`;
}
