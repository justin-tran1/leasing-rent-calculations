'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const C = require('../js/calc.js');

const close = (actual, expected, msg) =>
  assert.ok(Math.abs(actual - expected) < 0.005, `${msg || ''} expected ${expected}, got ${actual}`);

const base = (over) => Object.assign({}, C.DEFAULT_OPTION, {
  name: 'Test',
  size: 10000,
  commencement: '2026-11-01',
  expiration: '2028-10-31',
  baseRate: 30,
  rateBasis: 'psf_yr',
  rateType: 'NNN',
  escalation: 3,
  opex: 10,
  opexIncrease: 3.5,
  freeMonths: 0,
  parkingSpaces: 0,
  parkingRate: 0,
}, over);

test('defaults: escalation 3.0% and OpEx increase 3.5%', () => {
  assert.equal(C.DEFAULT_OPTION.escalation, 3.0);
  assert.equal(C.DEFAULT_OPTION.opexIncrease, 3.5);
});

test('expiration from term length', () => {
  assert.equal(C.expirationFromTerm('2026-11-01', 60), '2031-10-31');
  assert.equal(C.expirationFromTerm('2026-11-15', 12), '2027-11-14');
  assert.equal(C.expirationFromTerm('2026-01-31', 1), '2026-02-27');
  assert.equal(C.termMonthsFromInput(5, 'years'), 60);
  assert.equal(C.termMonthsFromInput(5.5, 'years'), 66);
  assert.equal(C.termMonthsFromInput(18, 'months'), 18);
});

test('term from dates handles whole and partial months', () => {
  assert.deepEqual(C.termFromDates('2026-11-01', '2031-10-31'), { whole: 60, fraction: 0, months: 60 });
  assert.equal(C.termFromDates('2026-11-15', '2027-11-14').months, 12);
  const t = C.termFromDates('2026-01-01', '2026-03-15');
  assert.equal(t.whole, 2);
  close(t.fraction, 15 / 31);
});

test('NNN: base rent escalates on the anniversary and OpEx passes through in full', () => {
  const r = C.calculate(base());
  assert.ok(r.ok);
  assert.equal(r.months.length, 24);
  close(r.months[0].baseRent, 25000);
  close(r.months[11].baseRent, 25000);
  close(r.months[12].baseRent, 25750);
  close(r.months[12].baseRatePsfYr, 30.9);
  close(r.months[0].opex, 10 * 10000 / 12);
  close(r.months[12].opex, 10.35 * 10000 / 12);
  close(r.summary.totalBaseRent, 12 * 25000 + 12 * 25750);
  close(r.summary.totalCost, r.summary.totalBaseRent + r.summary.totalOpex);
  assert.equal(r.years.length, 2);
  close(r.years[1].baseRent, 12 * 25750);
});

test('rate bases convert to the same rent', () => {
  const ref = C.calculate(base()).summary.totalBaseRent;
  close(C.calculate(base({ rateBasis: 'psf_mo', baseRate: 2.5 })).summary.totalBaseRent, ref, 'psf_mo');
  close(C.calculate(base({ rateBasis: 'monthly', baseRate: 25000 })).summary.totalBaseRent, ref, 'monthly');
  close(C.calculate(base({ rateBasis: 'annual', baseRate: 300000 })).summary.totalBaseRent, ref, 'annual');
});

test('fixed escalation adds a flat amount each lease year', () => {
  const r = C.calculate(base({ escalationType: 'fixed', escalation: 1 }));
  close(r.months[12].baseRatePsfYr, 31);
});

test('Full Service: tenant pays only OpEx increases over the base year', () => {
  const r = C.calculate(base({ rateType: 'FS' }));
  close(r.months[0].opex, 0);
  close(r.months[12].opex, (10.35 - 10) * 10000 / 12);
  const noBaseYear = C.calculate(base({ rateType: 'FS', baseYear: false }));
  close(noBaseYear.summary.totalOpex, 0);
});

test('Modified Gross: tenant share plus base-year increases on the landlord portion', () => {
  const r = C.calculate(base({ rateType: 'MG', mgTenantShare: 50 }));
  close(r.months[0].opex, 0.5 * 10 * 10000 / 12);
  close(r.months[12].opex, (0.5 * 10.35 + 0.5 * 0.35) * 10000 / 12);
  const noBaseYear = C.calculate(base({ rateType: 'MG', mgTenantShare: 50, baseYear: false }));
  close(noBaseYear.months[12].opex, 0.5 * 10.35 * 10000 / 12);
});

test('free rent at the start, including a partial month', () => {
  const r = C.calculate(base({ freeMonths: 2.5 }));
  close(r.months[0].freeRent, -25000);
  close(r.months[1].freeRent, -25000);
  close(r.months[2].freeRent, -12500);
  close(r.months[3].freeRent, 0);
  close(r.summary.totalFreeRent, -62500);
  close(r.summary.freeMonths, 2.5);
  close(r.months[0].opex, 10 * 10000 / 12, 'OpEx still due on base-only abatement');
});

test('free rent at the end of the term', () => {
  const r = C.calculate(base({ freeMonths: 3, freePlacement: 'end' }));
  close(r.months[20].freeRent, 0);
  close(r.months[21].freeRent, -25750);
  close(r.months[23].freeRent, -25750);
});

test('free rent in custom months', () => {
  const r = C.calculate(base({ freePlacement: 'custom', freeCustom: '1-2, 13' }));
  assert.deepEqual(r.months.filter((m) => m.freeFraction > 0).map((m) => m.month), [1, 2, 13]);
  assert.deepEqual(C.parseMonthList('3, 1-2, x').errors, ['x']);
  assert.ok(C.calculate(base({ freePlacement: 'custom', freeCustom: 'abc' })).errors.freeCustom);
});

test('gross abatement covers OpEx, and "all" covers parking too', () => {
  const opts = { freeMonths: 1, parkingSpaces: 10, parkingRate: 150 };
  const gross = C.calculate(base(Object.assign({ freeAppliesTo: 'gross' }, opts)));
  close(gross.months[0].total, 1500);
  const all = C.calculate(base(Object.assign({ freeAppliesTo: 'all' }, opts)));
  close(all.months[0].total, 0);
  close(all.summary.parkingAbatement, -1500);
});

test('parking cost and escalation', () => {
  const r = C.calculate(base({ parkingSpaces: 10, parkingRate: 150, parkingIncrease: 5 }));
  close(r.months[0].parking, 1500);
  close(r.months[12].parking, 1575);
  close(r.summary.totalParking, 12 * 1500 + 12 * 1575);
});

test('security deposit from first or last month base rent', () => {
  close(C.calculate(base({ depositBasis: 'first' })).summary.securityDeposit, 25000);
  close(C.calculate(base({ depositBasis: 'last' })).summary.securityDeposit, 25750);
  close(C.calculate(base({ depositBasis: 'first', depositMonths: 2 })).summary.securityDeposit, 50000);
  close(C.calculate(base({ depositBasis: 'none' })).summary.securityDeposit, 0);
  // Free rent never reduces the deposit.
  close(C.calculate(base({ depositBasis: 'first', freeMonths: 3 })).summary.securityDeposit, 25000);
});

test('partial final month is prorated by days', () => {
  const r = C.calculate(base({ expiration: '2027-01-15' }));
  assert.equal(r.months.length, 3);
  const last = r.months[2];
  close(last.proration, 15 / 31);
  close(last.baseRent, 25000 * 15 / 31);
  close(r.summary.termMonths, 2 + 15 / 31);
});

test('NPV equals total cost at a 0% discount rate and is lower at a positive rate', () => {
  const r0 = C.calculate(base(), { discountRate: 0 });
  close(r0.summary.npv, r0.summary.totalCost);
  const r8 = C.calculate(base(), { discountRate: 8 });
  assert.ok(r8.summary.npv < r8.summary.totalCost);
  // First payment is not discounted (rent paid in advance).
  const oneMonth = C.calculate(base({ expiration: '2026-11-30' }), { discountRate: 8 });
  close(oneMonth.summary.npv, oneMonth.summary.totalCost);
});

test('effective rent per SF per year', () => {
  const r = C.calculate(base());
  close(r.summary.effectiveRentPsfYr, r.summary.totalCost / 10000 / 2);
});

test('validation catches missing and inconsistent inputs', () => {
  const r = C.calculate(base({ size: 0, commencement: '', expiration: '2026-01-01' }));
  assert.equal(r.ok, false);
  assert.ok(r.errors.size);
  assert.ok(r.errors.commencement);
  const backwards = C.calculate(base({ expiration: '2026-10-01' }));
  assert.ok(backwards.errors.expiration);
});

test('month periods anchor to the commencement day', () => {
  const r = C.calculate(base({ commencement: '2026-01-31', expiration: '2026-04-29' }));
  assert.deepEqual(r.months.map((m) => [m.start, m.end]), [
    ['2026-01-31', '2026-02-27'],
    ['2026-02-28', '2026-03-30'],
    ['2026-03-31', '2026-04-29'],
  ]);
});

// ------------------------------------------------- regressions from review

test('free rent at the end counts full months of rent when the last month is partial', () => {
  const opts = { baseRate: 12, escalation: 0, opex: 0, commencement: '2026-01-01', expiration: '2031-01-15', freeMonths: 1 };
  const end = C.calculate(base(Object.assign({ freePlacement: 'end' }, opts)));
  const start = C.calculate(base(Object.assign({ freePlacement: 'start' }, opts)));
  close(end.summary.totalFreeRent, -10000);
  close(start.summary.totalFreeRent, -10000);
  assert.equal(end.summary.freeMonths, 1);
  close(end.months[end.months.length - 1].freeFraction, 1);
  close(end.months[end.months.length - 2].freeFraction, 1 - end.summary.finalMonthProration);
});

test('fractional terms keep their partial month', () => {
  assert.equal(C.termMonthsFromInput(24.5, 'months'), 24.5);
  assert.equal(C.expirationFromTerm('2026-11-01', 24.5), '2028-11-15');
  close(C.termFromDates('2026-11-01', '2028-11-15').months, 24.5);
  // Round trip across awkward commencement days and terms.
  ['2024-01-31', '2024-02-29', '2025-03-30', '2026-11-15', '2026-12-31'].forEach((c) => {
    [1, 12, 24.5, 60, 66.25, 600].forEach((t) => {
      const back = C.termFromDates(c, C.expirationFromTerm(c, t)).months;
      assert.ok(Math.abs(back - t) < 0.04, `${c} + ${t} months came back as ${back}`);
    });
  });
});

test('terms past 50 years are rejected without throwing', () => {
  assert.equal(C.expirationFromTerm('2026-11-01', C.termMonthsFromInput(300000, 'years')), '');
  const r = C.calculate(base({ termValue: 300000, termUnit: 'years' }));
  assert.equal(r.ok, false);
  assert.match(r.errors.termValue, /50 years/);
});

test('custom free months accept spaced ranges and must fall inside the term', () => {
  assert.deepEqual(C.parseMonthList('1 - 3, 13').months, [1, 2, 3, 13]);
  const r = C.calculate(base({ freePlacement: 'custom', freeCustom: '1-2, 30' }));
  assert.match(r.errors.freeCustom, /Month 30 falls after the term ends \(24 months\)/);
  assert.match(C.calculate(base({ freeMonths: 30 })).errors.freeMonths, /longer than the term/);
});

test('rates and shares outside sensible bounds are flagged', () => {
  const e = C.validateOption(base({ rateType: 'MG', mgTenantShare: 150, parkingRate: -5, parkingIncrease: -250, escalation: -100, opexIncrease: -100 }));
  assert.ok(e.mgTenantShare && e.parkingRate && e.parkingIncrease && e.escalation && e.opexIncrease);
  assert.ok(C.validateOption(base({ rateType: 'MG', mgTenantShare: '' })).mgTenantShare);
  assert.equal(C.validateOption(base({ escalation: -2 })).escalation, undefined, 'small negative escalations are allowed');
});

test('net components add up to the total lease cost', () => {
  const r = C.calculate(base({ freeMonths: 3, freeAppliesTo: 'all', parkingSpaces: 10, parkingRate: 100 }));
  const s = r.summary;
  close(s.netBaseRent + s.netOpex + s.netParking, s.totalCost);
});

test('a negative discount rate is treated as zero', () => {
  const r = C.calculate(base(), { discountRate: -100 });
  close(r.summary.npv, r.summary.totalCost);
});
