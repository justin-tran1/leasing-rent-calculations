'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const ExcelJS = require('exceljs');
const C = require('../js/calc.js');
const X = require('../js/excel.js');

const optionA = Object.assign({}, C.DEFAULT_OPTION, {
  name: 'Option A: 100 Main St',
  size: 10000,
  commencement: '2026-11-01',
  expiration: '2031-10-31',
  baseRate: 35,
  rateType: 'NNN',
  opex: 12,
  freeMonths: 3,
  parkingSpaces: 20,
  parkingRate: 150,
  depositBasis: 'first',
});
const optionB = Object.assign({}, optionA, {
  name: "Option B: King's Tower",
  rateType: 'FS',
  baseRate: 48,
  expiration: '2033-01-15',
  freeMonths: 6,
  freeAppliesTo: 'gross',
  depositBasis: 'last',
});

function payload() {
  const settings = { currency: '$', areaUnit: 'SF', discountRate: 8 };
  return {
    settings,
    generatedAt: new Date(Date.UTC(2026, 9, 5)),
    options: [optionA, optionB].map((o, i) => ({
      input: o,
      result: C.calculate(o, settings),
      color: ['#80BBAD', '#435254'][i],
    })),
  };
}

async function roundTrip(wb) {
  const buf = await wb.xlsx.writeBuffer();
  const back = new ExcelJS.Workbook();
  await back.xlsx.load(buf);
  return back;
}

test('workbook has comparison, monthly comparison and one sheet per option', async () => {
  const wb = await roundTrip(X.buildWorkbook(ExcelJS, payload()));
  assert.deepEqual(wb.worksheets.map((w) => w.name), [
    'Comparison', 'Monthly comparison', 'Option A 100 Main St', 'Option B King\u2019s Tower',
  ]);
});

test('option sheet lists every month with live totals', async () => {
  const p = payload();
  const wb = await roundTrip(X.buildWorkbook(ExcelJS, p));
  const ws = wb.getWorksheet('Option A 100 Main St');
  const r = p.options[0].result;
  const first = X.SCHEDULE_HEADER_ROW + 1;
  assert.equal(ws.getCell(first, 1).value, 1);
  assert.equal(ws.getCell(first + 59, 1).value, 60);
  const totalRow = first + 60;
  assert.equal(ws.getCell(totalRow, 1).value, 'Total');
  const total = ws.getCell(totalRow, 10).value;
  assert.equal(total.formula, `SUM(J${first}:J${first + 59})`);
  assert.ok(Math.abs(total.result - r.summary.totalCost) < 0.01);
  const rowTotal = ws.getCell(first, 10).value;
  assert.equal(rowTotal.formula, `F${first}+G${first}+H${first}+I${first}`);
});

test('comparison sheet links to option sheets and ranks total cost', async () => {
  const p = payload();
  const wb = await roundTrip(X.buildWorkbook(ExcelJS, p));
  const ws = wb.getWorksheet('Comparison');
  let costRow = null, rankRow = null;
  ws.eachRow((row, n) => {
    if (row.getCell(1).value === 'Total lease cost') costRow = n;
    if (String(row.getCell(1).value).startsWith('Rank by total')) rankRow = n;
  });
  assert.ok(costRow && rankRow);
  const b = ws.getCell(costRow, 3).value;
  assert.match(b.formula, /^'Option B King\u2019s Tower'!K\d+$/);
  assert.ok(Math.abs(b.result - p.options[1].result.summary.totalCost) < 0.01);
  assert.equal(ws.getCell('B4').value, 0.08);
  const ranks = [ws.getCell(rankRow, 2).value.result, ws.getCell(rankRow, 3).value.result];
  assert.deepEqual([...ranks].sort(), [1, 2]);
});

test('charts sheet is added when chart images are supplied', async () => {
  const p = payload();
  const png = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==';
  p.charts = [{ title: 'Total monthly cost', dataUrl: png, width: 900, height: 450 }];
  p.logo = png;
  const wb = await roundTrip(X.buildWorkbook(ExcelJS, p));
  assert.ok(wb.getWorksheet('Charts'));
  assert.equal(wb.getWorksheet('Charts').getImages().length, 2); // logo + chart
});

test('sheet names are sanitized and unique', () => {
  const used = new Set(['comparison']);
  assert.equal(X.uniqueSheetName('Comparison', used), 'Comparison (2)');
  assert.equal(X.uniqueSheetName('A/B: [test]?', used), 'A B test');
  assert.equal(X.uniqueSheetName('A/B: [test]?', used), 'A B test (2)');
  assert.equal(X.uniqueSheetName('x'.repeat(40), used).length, 31);
  // No straight apostrophes, so none can land at the start or end after truncation.
  const longApostrophe = X.uniqueSheetName("Option C, 10 Main St, Children's Hospital", used);
  assert.ok(!longApostrophe.includes("'") && longApostrophe.length <= 31);
});

test('options named History or with apostrophes export and reload cleanly', async () => {
  const p = payload();
  p.options[0].input = Object.assign({}, optionA, { name: 'History' });
  p.options[1].input = Object.assign({}, optionB, { name: "Option C, 10 Main St, Children's Hospital" });
  p.options.forEach((e) => { e.result = C.calculate(e.input, p.settings); });
  const wb = await roundTrip(X.buildWorkbook(ExcelJS, p));
  const names = wb.worksheets.map((w) => w.name);
  assert.ok(names.includes('History (2)'));
  const opt = wb.worksheets[3];
  assert.equal(opt.pageSetup.printTitlesRow, '5:5', 'print titles survive a reload');
});

test('invalid options are skipped; no valid options throws', () => {
  const bad = { input: Object.assign({}, optionA, { size: 0 }), result: C.calculate(Object.assign({}, optionA, { size: 0 })) };
  assert.throws(() => X.buildWorkbook(ExcelJS, { options: [bad] }), /No valid lease options/);
});
