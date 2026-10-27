import ExcelJS from 'exceljs';
import * as XLSX from 'xlsx';

export interface CellView {
  text: string;
  background?: string;
  foreground?: string;
  bold?: boolean;
  italic?: boolean;
  align?: string;
}
export interface Page { sheet: number; start: number; end: number; totalRows: number; columns: number; rows: CellView[][] }
export interface WorkbookView { names: string[]; page(sheet: number, start: number): Page }
const PAGE_SIZE = 100;
const MAX_COLUMNS = 200;

function color(argb?: string): string | undefined {
  return argb && /^(?:[0-9a-f]{8}|[0-9a-f]{6})$/i.test(argb) ? `#${argb.slice(-6)}` : undefined;
}

function excelValue(value: ExcelJS.CellValue): string {
  if (value === null || value === undefined) return '';
  if (value instanceof Date) return value.toISOString().replace('T', ' ').replace(/\.\d{3}Z$/, '');
  if (typeof value === 'object') {
    if ('result' in value) return excelValue(value.result as ExcelJS.CellValue);
    if ('richText' in value) return value.richText.map(part => part.text).join('');
    if ('text' in value) return String(value.text);
    if ('error' in value) return String(value.error);
    return '';
  }
  return String(value);
}

function styledCell(cell: ExcelJS.Cell): CellView {
  const value = cell.value;
  let text = excelValue(value);
  const numeric = typeof value === 'number' ? value :
    value && typeof value === 'object' && 'result' in value && typeof value.result === 'number' ? value.result : undefined;
  if (numeric !== undefined && cell.numFmt && cell.numFmt !== 'General') {
    try { text = XLSX.SSF.format(cell.numFmt, numeric); } catch { /* use raw value */ }
  }
  const font = cell.font;
  const fill = cell.fill;
  const alignment = cell.alignment?.horizontal;
  return {
    text,
    foreground: color(font?.color?.argb),
    background: fill?.type === 'pattern' && fill.pattern === 'solid' ? color(fill.fgColor?.argb) : undefined,
    bold: font?.bold || undefined,
    italic: font?.italic || undefined,
    align: alignment === 'left' || alignment === 'center' || alignment === 'right' ? alignment : undefined
  };
}

export async function parseWorkbook(bytes: Uint8Array, filename: string): Promise<WorkbookView> {
  if (/\.xlsx$/i.test(filename)) {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(Buffer.from(bytes) as never);
    const sheets = workbook.worksheets;
    return {
      names: sheets.map(sheet => sheet.name),
      page(sheetIndex, requestedStart) {
        const sheet = sheets[sheetIndex];
        if (!sheet) throw new Error('Sheet tidak ditemukan');
        const totalRows = sheet.rowCount;
        const columns = Math.min(sheet.columnCount, MAX_COLUMNS);
        const start = Math.max(1, Math.min(Math.floor(requestedStart) || 1, Math.max(1, totalRows)));
        const end = Math.min(totalRows, start + PAGE_SIZE - 1);
        const rows: CellView[][] = [];
        for (let row = start; row <= end; row++) {
          const source = sheet.getRow(row);
          rows.push(Array.from({ length: columns }, (_, col) => styledCell(source.getCell(col + 1))));
        }
        return { sheet: sheetIndex, start, end, totalRows, columns, rows };
      }
    };
  }
  if (!/\.xls$/i.test(filename)) throw new Error('Format file tidak didukung');
  const workbook = XLSX.read(bytes, { type: 'array', cellNF: true, cellStyles: true });
  return {
    names: workbook.SheetNames,
    page(sheetIndex, requestedStart) {
      const sheet = workbook.Sheets[workbook.SheetNames[sheetIndex]];
      if (!sheet) throw new Error('Sheet tidak ditemukan');
      const range = XLSX.utils.decode_range(sheet['!ref'] || 'A1');
      const totalRows = sheet['!ref'] ? range.e.r + 1 : 0;
      const columns = sheet['!ref'] ? Math.min(range.e.c + 1, MAX_COLUMNS) : 0;
      const start = Math.max(1, Math.min(Math.floor(requestedStart) || 1, Math.max(1, totalRows)));
      const end = Math.min(totalRows, start + PAGE_SIZE - 1);
      const rows: CellView[][] = [];
      for (let row = start; row <= end; row++) {
        rows.push(Array.from({ length: columns }, (_, col) => {
          const cell = sheet[XLSX.utils.encode_cell({ r: row - 1, c: col })];
          return { text: cell ? String(cell.w ?? cell.v ?? '') : '' };
        }));
      }
      return { sheet: sheetIndex, start, end, totalRows, columns, rows };
    }
  };
}
