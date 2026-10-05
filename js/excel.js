/*
 * Excel export. Builds a CBRE-branded workbook with ExcelJS:
 *   Comparison          every option side by side: terms and cost summary
 *   Monthly comparison  month-by-month rate, net base rent and total cost
 *   Charts              PNG snapshots of the charts on screen
 *   <one per option>    monthly schedule, annual summary, inputs and results
 *
 * Totals, annual rows and results are live Excel formulas (with cached
 * values) so the workbook stays consistent if someone edits a month.
 *
 * buildWorkbook() takes the ExcelJS module as an argument so the same code
 * runs in the browser (window.ExcelJS) and in Node tests (require('exceljs')).
 */
(function (root, factory) {
  const fmt = typeof module === 'object' && module.exports ? require('./format.js') : root.LeaseFormat;
  const api = factory(fmt);
  if (typeof module === 'object' && module.exports) module.exports = api;
  else root.LeaseExcel = api;
})(typeof self !== 'undefined' ? self : this, function (F) {
  'use strict';

  const C = {
    green: 'FF003F2D',
    darkGreen: 'FF012A2D',
    accent: 'FF17E88F',
    darkGrey: 'FF435254',
    lightGrey: 'FFCAD1D3',
    cement: 'FF7F8480',
    tint: 'FFE6F4EC',
    band: 'FFF6F6F6',
    white: 'FFFFFFFF',
  };
  const BODY_FONT = 'Arial';
  const HEAD_FONT = 'Georgia';
  // Our own sheet names, plus 'History', which Excel reserves.
  const RESERVED = ['Comparison', 'Monthly comparison', 'Charts', 'History'];

  const SCHEDULE_HEADER_ROW = 5;
  const SCHEDULE_COLS = [
    { key: 'month', title: 'Month', width: 8 },
    { key: 'leaseYear', title: 'Lease year', width: 10 },
    { key: 'start', title: 'Period start', width: 14 },
    { key: 'end', title: 'Period end', width: 14 },
    { key: 'rate', title: 'Base rent rate', width: 15 },
    { key: 'baseRent', title: 'Base rent', width: 16 },
    { key: 'freeRent', title: 'Free rent', width: 15 },
    { key: 'opex', title: 'OpEx', width: 15 },
    { key: 'parking', title: 'Parking', width: 13 },
    { key: 'total', title: 'Total monthly cost', width: 18 },
    { key: 'cumulative', title: 'Cumulative cost', width: 18 },
  ];

  // ------------------------------------------------------------- helpers

  function argb(hex) {
    const h = String(hex || '').replace('#', '');
    return /^[0-9a-fA-F]{6}$/.test(h) ? 'FF' + h.toUpperCase() : C.green;
  }

  function moneyFmt(cur, decimals) {
    const d = decimals === 0 ? '' : '.' + '0'.repeat(decimals === undefined ? 2 : decimals);
    const c = String(cur || '$').replace(/"/g, '');
    const pos = `"${c}"#,##0${d}`;
    return `${pos}_);(${pos});"-"_)`;
  }

  const DATE_FMT = 'mmm d, yyyy';
  const PCT_FMT = '0.0%';

  function colLetter(n) {
    let s = '';
    while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
    return s;
  }

  function sheetRef(sheet, addr) {
    return `'${sheet.replace(/'/g, "''")}'!${addr}`;
  }

  // Excel sheet names: at most 31 characters, none of []:*?/\, and no
  // leading or trailing apostrophe. Straight apostrophes become typographic
  // ones because ExcelJS does not escape them in print-title names, which
  // makes Excel report the file as damaged.
  function uniqueSheetName(name, used) {
    let n = String(name || 'Option').replace(/[\[\]:*?/\\]/g, ' ').replace(/'/g, '\u2019').replace(/\s+/g, ' ').trim();
    n = n.slice(0, 31).trim() || 'Option';
    let candidate = n, k = 2;
    while (used.has(candidate.toLowerCase())) {
      const suffix = ` (${k++})`;
      candidate = n.slice(0, 31 - suffix.length).trim() + suffix;
    }
    used.add(candidate.toLowerCase());
    return candidate;
  }

  function isoToDate(iso) {
    const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(iso || '');
    return m ? new Date(Date.UTC(+m[1], +m[2] - 1, +m[3])) : null;
  }

  function font(extra) {
    return Object.assign({ name: BODY_FONT, size: 10, color: { argb: C.darkGrey } }, extra);
  }

  function fill(color) {
    return { type: 'pattern', pattern: 'solid', fgColor: { argb: color } };
  }

  const thin = (color) => ({ style: 'thin', color: { argb: color || C.lightGrey } });

  function headerCell(cell, text, opts) {
    cell.value = text;
    cell.font = font({ bold: true, color: { argb: C.white } });
    cell.fill = fill(C.green);
    cell.alignment = Object.assign({ vertical: 'middle', horizontal: 'center', wrapText: true }, opts && opts.alignment);
    cell.border = { bottom: thin(C.green) };
  }

  function sectionCell(ws, row, fromCol, toCol, text) {
    const cell = ws.getCell(row, fromCol);
    cell.value = text;
    cell.font = font({ bold: true, size: 11, color: { argb: C.green } });
    for (let c = fromCol; c <= toCol; c++) {
      const cc = ws.getCell(row, c);
      cc.fill = fill(C.tint);
      cc.border = { bottom: thin(C.accent) };
    }
  }

  function formula(f, result) {
    return { formula: f, result: Number.isFinite(result) ? result : 0 };
  }

  function addBrandHeader(wb, ws, logoId, title, subtitle, widthCols) {
    ws.getRow(1).height = 42;
    if (logoId !== null && logoId !== undefined) {
      ws.addImage(logoId, { tl: { col: 0.15, row: 0.25 }, ext: { width: 112, height: 28 } });
    }
    const t = ws.getCell(2, 1);
    t.value = title;
    t.font = { name: HEAD_FONT, size: 18, bold: true, color: { argb: C.green } };
    ws.getRow(2).height = 26;
    const s = ws.getCell(3, 1);
    s.value = subtitle;
    s.font = font({ size: 10, color: { argb: C.cement } });
    // Accent rule under the header block.
    for (let c = 1; c <= widthCols; c++) ws.getCell(3, c).border = { bottom: { style: 'medium', color: { argb: C.green } } };
  }

  // -------------------------------------------------------- option sheets

  function buildOptionSheet(wb, entry, ctx) {
    const { input: o, result: r, sheetName, color } = entry;
    const s = r.summary;
    const cur = ctx.currency, unit = ctx.areaUnit;
    const money = moneyFmt(cur);
    const ws = wb.addWorksheet(sheetName, {
      properties: { tabColor: { argb: argb(color) } },
      views: [{ state: 'frozen', xSplit: 0, ySplit: SCHEDULE_HEADER_ROW, showGridLines: false }],
      pageSetup: { orientation: 'landscape', fitToPage: true, fitToWidth: 1, fitToHeight: 0, printTitlesRow: `${SCHEDULE_HEADER_ROW}:${SCHEDULE_HEADER_ROW}` },
    });
    SCHEDULE_COLS.forEach((c, i) => { ws.getColumn(i + 1).width = c.width; });

    const subtitle = [
      F.RATE_TYPE_LABELS[o.rateType],
      `${F.number(s.size)} ${unit}`,
      `${F.date(o.commencement)} - ${F.date(o.expiration)}`,
      F.termLabel(s.termMonths),
    ].join('  |  ');
    addBrandHeader(wb, ws, ctx.logoId, o.name, subtitle, SCHEDULE_COLS.length);

    // Key totals in row 4, filled in once the totals row address is known.
    ws.getRow(4).height = 20;

    // Schedule header.
    const h = SCHEDULE_HEADER_ROW;
    SCHEDULE_COLS.forEach((c, i) => {
      const title = c.key === 'rate' ? `Base rent (${cur}/${unit}/yr)` : c.title;
      headerCell(ws.getCell(h, i + 1), title);
    });
    ws.getRow(h).height = 30;

    const first = h + 1;
    const lastRow = first + r.months.length - 1;
    r.months.forEach((m, idx) => {
      const row = first + idx;
      const values = [
        m.month,
        m.leaseYear,
        isoToDate(m.start),
        isoToDate(m.end),
        m.baseRatePsfYr,
        m.baseRent,
        m.freeRent,
        m.opex,
        m.parking,
        formula(`F${row}+G${row}+H${row}+I${row}`, m.total),
        formula(idx === 0 ? `J${row}` : `K${row - 1}+J${row}`, m.cumulative),
      ];
      values.forEach((v, i) => {
        const cell = ws.getCell(row, i + 1);
        cell.value = v;
        cell.font = font();
        cell.border = { bottom: thin('FFE6EAEB') };
        if (idx % 2 === 1) cell.fill = fill(C.band);
        if (i <= 1) cell.alignment = { horizontal: 'center' };
        else if (i <= 3) cell.numFmt = DATE_FMT;
        else cell.numFmt = money;
      });
      if (m.freeFraction > 0) ws.getCell(row, 7).font = font({ color: { argb: C.green }, bold: true });
    });

    // Totals row.
    const tRow = lastRow + 1;
    const totals = {
      F: s.totalBaseRent, G: s.totalFreeRent, H: s.totalOpex, I: s.totalParking, J: s.totalCost,
    };
    for (let c = 1; c <= SCHEDULE_COLS.length; c++) {
      const cell = ws.getCell(tRow, c);
      cell.fill = fill(C.tint);
      cell.font = font({ bold: true, color: { argb: C.green } });
      cell.border = { top: { style: 'medium', color: { argb: C.green } }, bottom: { style: 'medium', color: { argb: C.green } } };
    }
    ws.getCell(tRow, 1).value = 'Total';
    Object.keys(totals).forEach((col) => {
      const cell = ws.getCell(`${col}${tRow}`);
      cell.value = formula(`SUM(${col}${first}:${col}${lastRow})`, totals[col]);
      cell.numFmt = money;
    });
    const totalCostAddr = `J${tRow}`;

    // Annual summary.
    let row = tRow + 3;
    sectionCell(ws, row, 1, SCHEDULE_COLS.length, 'Annual summary');
    row++;
    const annualHeads = ['Lease year', 'Months', 'Period start', 'Period end', `Base rent (${cur}/${unit}/yr)`,
      'Base rent', 'Free rent', 'OpEx', 'Parking', 'Total annual cost', 'Cumulative cost'];
    annualHeads.forEach((t, i) => headerCell(ws.getCell(row, i + 1), t));
    ws.getRow(row).height = 30;
    const annualFirst = row + 1;
    let cumulative = 0;
    r.years.forEach((y, idx) => {
      const rr = annualFirst + idx;
      cumulative += y.total;
      const sumif = (col, val) => formula(`SUMIF($B$${first}:$B$${lastRow},$A${rr},${col}$${first}:${col}$${lastRow})`, val);
      const values = [
        y.leaseYear,
        Math.round(y.months * 100) / 100,
        isoToDate(y.start),
        isoToDate(y.end),
        y.baseRatePsfYr,
        sumif('F', y.baseRent),
        sumif('G', y.freeRent),
        sumif('H', y.opex),
        sumif('I', y.parking),
        sumif('J', y.total),
        formula(idx === 0 ? `J${rr}` : `K${rr - 1}+J${rr}`, cumulative),
      ];
      values.forEach((v, i) => {
        const cell = ws.getCell(rr, i + 1);
        cell.value = v;
        cell.font = font();
        cell.border = { bottom: thin('FFE6EAEB') };
        if (idx % 2 === 1) cell.fill = fill(C.band);
        if (i <= 1) cell.alignment = { horizontal: 'center' };
        else if (i <= 3) cell.numFmt = DATE_FMT;
        else cell.numFmt = money;
      });
    });
    row = annualFirst + r.years.length + 2;

    // Inputs (left) and results (right).
    const inputsStart = row;
    sectionCell(ws, row, 1, 5, 'Lease inputs');
    sectionCell(ws, row, 7, 11, 'Results');
    const inputs = [
      ['Rate type', F.RATE_TYPE_LABELS[o.rateType]],
      [`Size (${unit})`, s.size, '#,##0'],
      ['Commencement', isoToDate(o.commencement), DATE_FMT],
      ['Expiration', isoToDate(o.expiration), DATE_FMT],
      ['Term (months)', s.termMonths, Number.isInteger(s.termMonths) ? '0' : '0.00'],
      [`Base rent (${F.rateBasisLabel(o.rateBasis, cur, unit)})`, Number(o.baseRate), money],
      ['Annual escalation', F.escalationLabel(o, cur, unit)],
      [`OpEx (${cur}/${unit}/yr)`, Number(o.opex), money],
      ['Annual OpEx increase', Number(o.opexIncrease) / 100, PCT_FMT],
      ['Tenant share of OpEx', s.tenantOpexShare, '0%'],
      ['Pays OpEx increases over base year', o.rateType === 'NNN' ? 'n/a (NNN)' : (s.baseYear ? 'Yes' : 'No')],
      ['Free rent', F.freeRentLabel(o, s.freeMonths)],
      ['Free rent applies to', F.APPLIES_LABELS[o.freeAppliesTo]],
      ['Parking spaces', Number(o.parkingSpaces), '#,##0'],
      [`Parking rate (${cur}/space/month)`, Number(o.parkingRate), money],
      ['Annual parking increase', Number(o.parkingIncrease) / 100, PCT_FMT],
      ['Security deposit basis', F.depositLabel(o.depositBasis, s.depositMonths)],
    ];
    const termCell = `E${inputsStart + 5}`;
    const sizeCell = `E${inputsStart + 2}`;
    inputs.forEach(([label, value, numFmt], i) => {
      const rr = inputsStart + 1 + i;
      ws.mergeCells(rr, 1, rr, 4);
      const l = ws.getCell(rr, 1);
      l.value = label;
      l.font = font();
      const v = ws.getCell(rr, 5);
      v.value = value;
      v.font = font({ bold: true, color: { argb: C.darkGreen } });
      v.alignment = { horizontal: 'right' };
      if (numFmt) v.numFmt = numFmt;
      [l, v].forEach((c) => { c.border = { bottom: thin('FFE6EAEB') }; });
    });

    const npvRate = ctx.discountCell;
    const rm = `((1+${npvRate})^(1/12)-1)`;
    const results = [
      ['Total base rent', formula(`F${tRow}`, s.totalBaseRent), money],
      ['Free rent', formula(`G${tRow}`, s.totalFreeRent), money],
      // Free rent column G holds every abatement; it equals the base-rent
      // abatement only when free rent covers base rent alone.
      ['Net base rent', o.freeAppliesTo === 'base' ? formula(`F${tRow}+G${tRow}`, s.netBaseRent) : s.netBaseRent, money],
      ['OpEx', formula(`H${tRow}`, s.totalOpex), money],
      ['Parking', formula(`I${tRow}`, s.totalParking), money],
      ['Total lease cost', formula(totalCostAddr, s.totalCost), money, true],
      ['Average monthly cost', formula(`${totalCostAddr}/${termCell}`, s.avgMonthlyCost), money],
      ['Average annual cost', formula(`${totalCostAddr}/(${termCell}/12)`, s.avgAnnualCost), money],
      [`Effective rent (${cur}/${unit}/yr)`, formula(`${totalCostAddr}/${sizeCell}/(${termCell}/12)`, s.effectiveRentPsfYr), money],
      ['NPV of total cost', formula(`NPV(${rm},J${first}:J${lastRow})*(1+${rm})`, s.npv), money],
      ['Security deposit', s.securityDeposit, money, true],
    ];
    const resultCells = {};
    results.forEach(([label, value, numFmt, strong], i) => {
      const rr = inputsStart + 1 + i;
      ws.mergeCells(rr, 7, rr, 10);
      const l = ws.getCell(rr, 7);
      l.value = label;
      l.font = font({ bold: !!strong, color: { argb: strong ? C.green : C.darkGrey } });
      const v = ws.getCell(rr, 11);
      v.value = value;
      v.numFmt = numFmt;
      v.font = font({ bold: true, color: { argb: strong ? C.green : C.darkGreen } });
      if (strong) [l, v].forEach((c) => { c.fill = fill(C.tint); });
      [l, v].forEach((c) => { c.border = { bottom: thin('FFE6EAEB') }; });
      resultCells[label] = `K${rr}`;
    });
    const noteRow = inputsStart + Math.max(inputs.length, results.length) + 2;
    const notes = [
      `NPV discounts monthly payments, paid in advance, at the discount rate on the Comparison sheet.`,
      s.finalMonthProration < 1
        ? `Final month prorated at ${F.pct(s.finalMonthProration * 100)} of a full month.`
        : null,
      'Estimates for comparison only. Confirm all figures against the lease documents.',
    ].filter(Boolean);
    notes.forEach((n, i) => {
      const cell = ws.getCell(noteRow + i, 1);
      cell.value = n;
      cell.font = font({ italic: true, size: 9, color: { argb: C.cement } });
    });

    // Headline figures in row 4.
    const headline = [
      ['Total lease cost', formula(totalCostAddr, s.totalCost), 1, 3, 4],
      ['Average monthly cost', formula(resultCells['Average monthly cost'], s.avgMonthlyCost), 5, 7, 8],
      ['Security deposit', formula(resultCells['Security deposit'], s.securityDeposit), 9, 11, 11],
    ];
    headline.forEach(([label, value, col, vFrom, vTo]) => {
      ws.mergeCells(4, col, 4, vFrom - 1);
      const l = ws.getCell(4, col);
      l.value = label;
      l.font = font({ size: 9, color: { argb: C.cement } });
      l.alignment = { vertical: 'middle', horizontal: 'right' };
      if (vTo > vFrom) ws.mergeCells(4, vFrom, 4, vTo);
      const v = ws.getCell(4, vFrom);
      v.value = value;
      v.numFmt = money;
      v.font = font({ bold: true, size: 11, color: { argb: C.green } });
      v.alignment = { horizontal: 'left', vertical: 'middle' };
    });

    return {
      ws,
      first,
      lastRow,
      totalsRow: tRow,
      cells: Object.assign({ termMonths: termCell, size: sizeCell }, resultCells),
    };
  }

  // ------------------------------------------------------ comparison sheet

  function buildComparisonSheet(wb, entries, ctx) {
    const ws = wb.getWorksheet('Comparison');
    const cur = ctx.currency, unit = ctx.areaUnit;
    const money = moneyFmt(cur);
    const n = entries.length;
    ws.getColumn(1).width = 38;
    for (let i = 0; i < n; i++) ws.getColumn(i + 2).width = 22;

    const generated = ctx.generatedAt;
    const pad = (x) => String(x).padStart(2, '0');
    const localDay = `${generated.getFullYear()}-${pad(generated.getMonth() + 1)}-${pad(generated.getDate())}`;
    const subtitle = `${n} lease ${n === 1 ? 'option' : 'options'}  |  Prepared ${F.date(localDay)}`;
    addBrandHeader(wb, ws, ctx.logoId, 'Lease comparison', subtitle, Math.max(4, n + 1));

    // Discount rate input (used by every NPV formula).
    const dl = ws.getCell('A4');
    dl.value = 'Discount rate for NPV (edit to recalculate)';
    dl.font = font({ size: 9, color: { argb: C.cement } });
    const dv = ws.getCell('B4');
    dv.value = ctx.discountRate;
    dv.numFmt = '0.00%';
    dv.font = font({ bold: true, color: { argb: 'FF0000FF' } }); // blue = editable input
    dv.border = { bottom: thin(C.accent) };

    // Option header with a colour key matching the charts.
    const hRow = 6;
    headerCell(ws.getCell(hRow, 1), '', { alignment: { horizontal: 'left' } });
    entries.forEach((e, i) => {
      const key = ws.getCell(hRow - 1, i + 2);
      key.fill = fill(argb(e.color));
      headerCell(ws.getCell(hRow, i + 2), e.input.name);
    });
    ws.getRow(hRow - 1).height = 6;
    ws.getRow(hRow).height = 30;
    ws.views = [{ state: 'frozen', xSplit: 1, ySplit: hRow, showGridLines: false }];

    let row = hRow + 1;
    const lowest = entries.reduce((best, e, i) => (e.result.summary.totalCost < entries[best].result.summary.totalCost ? i : best), 0);

    const section = (title) => { sectionCell(ws, row, 1, n + 1, title); row++; };
    const line = (label, getter, numFmt, opts) => {
      const l = ws.getCell(row, 1);
      l.value = label;
      l.font = font({ bold: !!(opts && opts.strong), color: { argb: opts && opts.strong ? C.green : C.darkGrey } });
      l.border = { bottom: thin('FFE6EAEB') };
      entries.forEach((e, i) => {
        const cell = ws.getCell(row, i + 2);
        cell.value = getter(e, i);
        if (numFmt) cell.numFmt = typeof numFmt === 'function' ? numFmt(e) : numFmt;
        cell.alignment = { horizontal: 'right', wrapText: true, vertical: 'top' };
        cell.font = font({ bold: !!(opts && opts.strong), color: { argb: opts && opts.strong ? C.green : C.darkGrey } });
        cell.border = { bottom: thin('FFE6EAEB') };
        if (opts && opts.strong) cell.fill = fill(C.tint);
      });
      if (opts && opts.strong) l.fill = fill(C.tint);
      row++;
    };
    const ref = (e, key, value) => formula(sheetRef(e.sheetName, e.cells[key]), value);

    section('Lease terms');
    line('Rate type', (e) => F.RATE_TYPE_LABELS[e.input.rateType]);
    line(`Size (${unit})`, (e) => e.result.summary.size, '#,##0');
    line('Commencement', (e) => isoToDate(e.input.commencement), DATE_FMT);
    line('Expiration', (e) => isoToDate(e.input.expiration), DATE_FMT);
    line('Term (months)', (e) => e.result.summary.termMonths,
      (e) => (Number.isInteger(e.result.summary.termMonths) ? '0' : '0.00'));
    line(`Starting base rent (${cur}/${unit}/yr)`, (e) => e.result.summary.startingRatePsfYr, money);
    line(`Ending base rent (${cur}/${unit}/yr)`, (e) => e.result.summary.endingRatePsfYr, money);
    line('Starting monthly base rent', (e) => e.result.summary.startingMonthlyBase, money);
    line('Annual escalation', (e) => F.escalationLabel(e.input, cur, unit));
    line(`OpEx (${cur}/${unit}/yr)`, (e) => Number(e.input.opex), money);
    line('Annual OpEx increase', (e) => Number(e.input.opexIncrease) / 100, PCT_FMT);
    line('Tenant share of OpEx', (e) => e.result.summary.tenantOpexShare, '0%');
    line('Pays OpEx increases over base year', (e) => (e.input.rateType === 'NNN' ? 'n/a' : e.result.summary.baseYear ? 'Yes' : 'No'));
    line('Free rent', (e) => F.freeRentLabel(e.input, e.result.summary.freeMonths));
    line('Free rent applies to', (e) => F.APPLIES_LABELS[e.input.freeAppliesTo]);
    line('Parking', (e) => (Number(e.input.parkingSpaces) > 0
      ? `${F.number(e.input.parkingSpaces)} spaces at ${F.money(e.input.parkingRate, cur)}/month`
      : 'None'));
    line('Annual parking increase', (e) => Number(e.input.parkingIncrease) / 100, PCT_FMT);
    row++;

    section('Cost summary');
    line('Total base rent', (e) => ref(e, 'Total base rent', e.result.summary.totalBaseRent), money);
    line('Free rent', (e) => ref(e, 'Free rent', e.result.summary.totalFreeRent), money);
    line('Net base rent', (e) => ref(e, 'Net base rent', e.result.summary.netBaseRent), money);
    line('OpEx', (e) => ref(e, 'OpEx', e.result.summary.totalOpex), money);
    line('Parking', (e) => ref(e, 'Parking', e.result.summary.totalParking), money);
    const costRow = row;
    line('Total lease cost', (e) => ref(e, 'Total lease cost', e.result.summary.totalCost), money, { strong: true });
    line('Average monthly cost', (e) => ref(e, 'Average monthly cost', e.result.summary.avgMonthlyCost), money);
    line('Average annual cost', (e) => ref(e, 'Average annual cost', e.result.summary.avgAnnualCost), money);
    line(`Effective rent (${cur}/${unit}/yr)`, (e) => ref(e, `Effective rent (${cur}/${unit}/yr)`, e.result.summary.effectiveRentPsfYr), money);
    line(`NPV of total cost`, (e) => ref(e, 'NPV of total cost', e.result.summary.npv), money);
    if (n > 1) {
      const range = `$B$${costRow}:$${colLetter(n + 1)}$${costRow}`;
      const ranks = entries.map((e) => e.result.summary.totalCost).map((v, _, all) => 1 + all.filter((x) => x < v).length);
      line('Rank by total lease cost (1 = lowest)',
        (e, i) => formula(`COUNTIF(${range},"<"&${colLetter(i + 2)}${costRow})+1`, ranks[i]), '0');
    }
    row++;

    section('Security deposit');
    line('Basis', (e) => F.depositLabel(e.input.depositBasis, e.result.summary.depositMonths));
    line('Security deposit', (e) => ref(e, 'Security deposit', e.result.summary.securityDeposit), money, { strong: true });
    row++;

    if (n > 1) {
      const c = ws.getCell(row, 1);
      c.value = `Lowest total lease cost: ${entries[lowest].input.name}`;
      c.font = font({ bold: true, color: { argb: C.green } });
      row++;
    }
    const notes = [
      'Rate types: NNN passes all OpEx through. Full Service includes OpEx in rent; where selected, the tenant pays increases over the base year.',
      'Modified Gross: the tenant pays its share of OpEx directly, plus base-year increases on the landlord share where selected.',
      'Escalations and OpEx increases apply on each lease anniversary. Estimates for comparison only; confirm against the lease documents.',
    ];
    notes.forEach((t) => {
      const cell = ws.getCell(row++, 1);
      cell.value = t;
      cell.font = font({ italic: true, size: 9, color: { argb: C.cement } });
    });
    ws.pageSetup = { orientation: 'landscape', fitToPage: true, fitToWidth: 1, fitToHeight: 0 };
  }

  // ---------------------------------------------- monthly comparison sheet

  function buildMonthlySheet(wb, entries, ctx) {
    const cur = ctx.currency, unit = ctx.areaUnit;
    const money = moneyFmt(cur);
    const ws = wb.addWorksheet('Monthly comparison', {
      properties: { tabColor: { argb: C.darkGrey } },
      views: [{ state: 'frozen', xSplit: 1, ySplit: 6, showGridLines: false }],
      pageSetup: { orientation: 'landscape', fitToPage: true, fitToWidth: 1, fitToHeight: 0 },
    });
    const per = 3;
    const width = 1 + entries.length * per;
    ws.getColumn(1).width = 9;
    for (let c = 2; c <= width; c++) ws.getColumn(c).width = 16;
    addBrandHeader(wb, ws, ctx.logoId, 'Monthly comparison', 'Aligned by lease month. Month 1 is each option\'s commencement month.', width);

    const top = 5, sub = 6;
    headerCell(ws.getCell(top, 1), '');
    headerCell(ws.getCell(sub, 1), 'Month');
    ws.mergeCells(top, 1, sub, 1);
    entries.forEach((e, i) => {
      const c0 = 2 + i * per;
      ws.mergeCells(top, c0, top, c0 + per - 1);
      headerCell(ws.getCell(top, c0), e.input.name);
      ws.getCell(top, c0).border = { bottom: { style: 'thick', color: { argb: argb(e.color) } } };
      [`Rate (${cur}/${unit}/yr)`, 'Net base rent', 'Total monthly cost'].forEach((t, k) => headerCell(ws.getCell(sub, c0 + k), t));
    });
    ws.getRow(sub).height = 30;

    const maxMonths = Math.max(...entries.map((e) => e.result.months.length));
    for (let i = 0; i < maxMonths; i++) {
      const rr = sub + 1 + i;
      const mc = ws.getCell(rr, 1);
      mc.value = i + 1;
      mc.alignment = { horizontal: 'center' };
      mc.font = font();
      entries.forEach((e, k) => {
        const c0 = 2 + k * per;
        const m = e.result.months[i];
        const src = e.first + i;
        const net = e.input.freeAppliesTo === 'base'
          ? formula(`${sheetRef(e.sheetName, `F${src}`)}+${sheetRef(e.sheetName, `G${src}`)}`, m && m.netBaseRent)
          : m && m.netBaseRent;
        const vals = m ? [
          formula(sheetRef(e.sheetName, `E${src}`), m.baseRatePsfYr),
          net,
          formula(sheetRef(e.sheetName, `J${src}`), m.total),
        ] : [null, null, null];
        vals.forEach((v, j) => {
          const cell = ws.getCell(rr, c0 + j);
          cell.value = v;
          cell.numFmt = money;
          cell.font = font();
        });
      });
      for (let c = 1; c <= width; c++) {
        const cell = ws.getCell(rr, c);
        cell.border = { bottom: thin('FFE6EAEB') };
        if (i % 2 === 1) cell.fill = fill(C.band);
      }
    }
    const tRow = sub + 1 + maxMonths;
    for (let c = 1; c <= width; c++) {
      const cell = ws.getCell(tRow, c);
      cell.fill = fill(C.tint);
      cell.font = font({ bold: true, color: { argb: C.green } });
      cell.border = { top: { style: 'medium', color: { argb: C.green } }, bottom: { style: 'medium', color: { argb: C.green } } };
    }
    ws.getCell(tRow, 1).value = 'Total';
    entries.forEach((e, k) => {
      const c0 = 2 + k * per;
      [[1, e.result.summary.netBaseRent], [2, e.result.summary.totalCost]].forEach(([j, v]) => {
        const L = colLetter(c0 + j);
        const cell = ws.getCell(tRow, c0 + j);
        cell.value = formula(`SUM(${L}${sub + 1}:${L}${tRow - 1})`, v);
        cell.numFmt = money;
      });
    });
  }

  // -------------------------------------------------------- charts sheet

  function buildChartsSheet(wb, charts, ctx) {
    const ws = wb.addWorksheet('Charts', {
      properties: { tabColor: { argb: C.accent } },
      views: [{ showGridLines: false }],
    });
    for (let c = 1; c <= 20; c++) ws.getColumn(c).width = 10;
    addBrandHeader(wb, ws, ctx.logoId, 'Charts', 'Snapshot of the charts as configured in the calculator.', 14);
    // Default column width 10 chars ~ 75px; default row height 15pt = 20px.
    const colPx = 75, rowPx = 20;
    const perRow = charts.length > 1 ? 2 : 1;
    const maxW = perRow === 1 ? 960 : 620;
    let rowCursor = 5;
    for (let i = 0; i < charts.length; i += perRow) {
      let tallest = 0;
      for (let j = 0; j < perRow && i + j < charts.length; j++) {
        const ch = charts[i + j];
        const scale = Math.min(1, maxW / ch.width);
        const w = Math.round(ch.width * scale), hgt = Math.round(ch.height * scale);
        const id = wb.addImage({ base64: ch.dataUrl, extension: 'png' });
        const col = j * (Math.ceil(maxW / colPx) + 1);
        const t = ws.getCell(rowCursor, col + 1);
        t.value = ch.title;
        t.font = font({ bold: true, size: 11, color: { argb: C.darkGrey } });
        ws.addImage(id, { tl: { col, row: rowCursor }, ext: { width: w, height: hgt } });
        tallest = Math.max(tallest, hgt);
      }
      rowCursor += Math.ceil(tallest / rowPx) + 3;
    }
  }

  // ------------------------------------------------------------- public

  /*
   * payload = {
   *   options:  [{ input, result, color }]   (result from LeaseCalc.calculate, ok only)
   *   settings: { currency, areaUnit, discountRate }
   *   charts:   [{ title, dataUrl, width, height }]   optional
   *   logo:     PNG data URL                          optional
   *   generatedAt: Date                               optional
   * }
   */
  function buildWorkbook(ExcelJS, payload) {
    const entries = payload.options.filter((e) => e.result && e.result.ok).map((e) => Object.assign({}, e));
    if (!entries.length) throw new Error('No valid lease options to export.');
    const settings = payload.settings || {};
    const wb = new ExcelJS.Workbook();
    wb.creator = 'CBRE lease rent calculator';
    wb.created = payload.generatedAt || new Date();
    wb.calcProperties.fullCalcOnLoad = true;

    const ctx = {
      currency: settings.currency || '$',
      areaUnit: settings.areaUnit || 'SF',
      discountRate: (Number(settings.discountRate) || 0) / 100,
      discountCell: "Comparison!$B$4",
      generatedAt: payload.generatedAt || new Date(),
      logoId: payload.logo ? wb.addImage({ base64: payload.logo, extension: 'png' }) : null,
    };

    // Comparison must be the first tab; fill it last once option cells exist.
    wb.addWorksheet('Comparison', { properties: { tabColor: { argb: C.green } } });

    const used = new Set(RESERVED.map((s) => s.toLowerCase()));
    entries.forEach((e) => {
      e.sheetName = uniqueSheetName(e.input.name, used);
      e.first = SCHEDULE_HEADER_ROW + 1;
    });

    buildMonthlySheet(wb, entries, ctx);
    if (payload.charts && payload.charts.length) buildChartsSheet(wb, payload.charts, ctx);
    entries.forEach((e) => { e.cells = buildOptionSheet(wb, e, ctx).cells; });
    buildComparisonSheet(wb, entries, ctx);
    return wb;
  }

  function fileName(date) {
    const d = date || new Date();
    const pad = (x) => String(x).padStart(2, '0');
    return `CBRE-lease-comparison-${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}.xlsx`;
  }

  return { buildWorkbook, fileName, uniqueSheetName, moneyFmt, SCHEDULE_HEADER_ROW };
});
