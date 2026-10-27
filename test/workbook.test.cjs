const { test } = require('node:test');
const assert = require('node:assert/strict');
const ExcelJS = require('exceljs');
const XLSX = require('xlsx');
const { parseWorkbook } = require('../dist/workbook.js');

test('xlsx sheets, styles, formula result and paging', async () => {
  const book = new ExcelJS.Workbook();
  const sheet = book.addWorksheet('Sales');
  sheet.getCell('A1').value = 1234.5;
  sheet.getCell('A1').numFmt = '#,##0.00';
  sheet.getCell('A1').font = { bold: true, color: { argb: 'FFFF0000' } };
  sheet.getCell('A1').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFF00' } };
  sheet.getCell('B1').value = { formula: '1+2', result: 3 };
  sheet.getCell('A101').value = 'last';
  book.addWorksheet('Other').getCell('A1').value = 'another sheet';
  const parsed = await parseWorkbook(await book.xlsx.writeBuffer(), 'sample.xlsx');
  assert.deepEqual(parsed.names, ['Sales', 'Other']);
  const first = parsed.page(0, 1);
  assert.equal(first.rows[0][0].text, '1,234.50');
  assert.equal(first.rows[0][0].foreground, '#FF0000');
  assert.equal(first.rows[0][0].background, '#FFFF00');
  assert.equal(first.rows[0][0].bold, true);
  assert.equal(first.rows[0][1].text, '3');
  assert.equal(first.end, 100);
  assert.equal(parsed.page(0, 101).rows[0][0].text, 'last');
  assert.equal(parsed.page(1, 1).rows[0][0].text, 'another sheet');
});

test('legacy xls formatted values', async () => {
  const book = XLSX.utils.book_new();
  const sheet = XLSX.utils.aoa_to_sheet([[12.5, 'legacy']]);
  sheet.A1.z = '0.00';
  XLSX.utils.book_append_sheet(book, sheet, 'Old');
  const bytes = XLSX.write(book, { type: 'buffer', bookType: 'biff8' });
  const parsed = await parseWorkbook(bytes, 'sample.xls');
  assert.deepEqual(parsed.names, ['Old']);
  assert.equal(parsed.page(0, 1).rows[0][0].text, '12.50');
  assert.equal(parsed.page(0, 1).rows[0][1].text, 'legacy');
});
