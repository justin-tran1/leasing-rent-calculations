/*
 * Lease rent calculation engine.
 *
 * Pure functions only: no DOM access, so the same file runs in the browser
 * (as window.LeaseCalc) and in Node for tests (module.exports).
 *
 * Dates are ISO strings (YYYY-MM-DD) handled in UTC so daylight saving never
 * shifts a day. A "lease month" runs from the commencement day-of-month to
 * the day before it in the next month (Nov 15 - Dec 14), and lease years are
 * counted in 12-month blocks from commencement. Escalations apply on each
 * lease anniversary.
 */
(function (root, factory) {
  const api = factory();
  if (typeof module === 'object' && module.exports) module.exports = api;
  else root.LeaseCalc = api;
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  const MS_PER_DAY = 86400000;
  const MAX_TERM_MONTHS = 600;

  const RATE_TYPES = {
    NNN: { label: 'NNN', long: 'Triple net (NNN)' },
    MG: { label: 'Modified Gross', long: 'Modified Gross' },
    FS: { label: 'Full Service', long: 'Full Service' },
  };

  const RATE_BASES = {
    psf_yr: { label: '/SF/yr', perArea: true, monthsPerUnit: 12 },
    psf_mo: { label: '/SF/mo', perArea: true, monthsPerUnit: 1 },
    monthly: { label: '/month', perArea: false, monthsPerUnit: 1 },
    annual: { label: '/year', perArea: false, monthsPerUnit: 12 },
  };

  const DEFAULT_OPTION = {
    name: 'Option A',
    size: 10000,
    commencement: '',
    termValue: 5,
    termUnit: 'years',
    expiration: '',
    baseRate: 0,
    rateBasis: 'psf_yr',
    rateType: 'NNN',
    escalation: 3.0,
    escalationType: 'pct',
    opex: 0,
    opexIncrease: 3.5,
    mgTenantShare: 50,
    baseYear: true,
    freeMonths: 0,
    freePlacement: 'start',
    freeCustom: '',
    freeAppliesTo: 'base',
    parkingSpaces: 0,
    parkingRate: 0,
    parkingIncrease: 0,
    depositBasis: 'first',
    depositMonths: 1,
  };

  // ---------------------------------------------------------------- dates

  function parseDate(iso) {
    if (typeof iso !== 'string') return null;
    const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(iso.trim());
    if (!m) return null;
    const y = +m[1], mo = +m[2] - 1, d = +m[3];
    const dt = new Date(Date.UTC(y, mo, d));
    if (dt.getUTCFullYear() !== y || dt.getUTCMonth() !== mo || dt.getUTCDate() !== d) return null;
    return dt;
  }

  function toISO(dt) {
    return dt.toISOString().slice(0, 10);
  }

  function daysInMonth(y, m) {
    return new Date(Date.UTC(y, m + 1, 0)).getUTCDate();
  }

  // Adds n months to `base`, pinning the day to `anchorDay` (clamped to the
  // month's length) so Jan 31 -> Feb 28 -> Mar 31 instead of drifting.
  function addMonths(base, n, anchorDay) {
    const day = anchorDay || base.getUTCDate();
    const total = base.getUTCFullYear() * 12 + base.getUTCMonth() + n;
    const y = Math.floor(total / 12), m = total - y * 12;
    return new Date(Date.UTC(y, m, Math.min(day, daysInMonth(y, m))));
  }

  function addDays(dt, n) {
    return new Date(dt.getTime() + n * MS_PER_DAY);
  }

  function daysBetween(a, b) {
    return Math.round((b.getTime() - a.getTime()) / MS_PER_DAY);
  }

  // Term input in months. Fractions are kept (5.5 years = 66 months, 24.5
  // months = 24 months plus half of the 25th lease month).
  function termMonthsFromInput(value, unit) {
    const v = Number(value);
    if (value === '' || value === null || !isFinite(v) || v <= 0) return 0;
    return Math.round((unit === 'years' ? v * 12 : v) * 1e4) / 1e4;
  }

  // Expiration is the day before the commencement day, `months` later. A
  // fractional month adds that share of the next lease month's days, so this
  // is the inverse of termFromDates. Returns '' for terms past the 50-year cap.
  function expirationFromTerm(commencementISO, months) {
    const c = parseDate(commencementISO);
    if (!c || !(months > 0) || months > MAX_TERM_MONTHS) return '';
    const anchor = c.getUTCDate();
    const whole = Math.floor(months + 1e-9);
    const fraction = months - whole;
    const start = addMonths(c, whole, anchor);
    let days = 0;
    if (fraction > 1e-6) days = Math.round(fraction * daysBetween(start, addMonths(c, whole + 1, anchor)));
    if (whole === 0 && days === 0) days = 1;
    const end = addDays(start, days - 1);
    return end.getUTCFullYear() > 9999 ? '' : toISO(end);
  }

  // Term between two dates as whole lease months plus a fraction for a
  // partial final month (prorated by days).
  function termFromDates(commencementISO, expirationISO) {
    const c = parseDate(commencementISO), e = parseDate(expirationISO);
    if (!c || !e || e < c) return { whole: 0, fraction: 0, months: 0 };
    const anchor = c.getUTCDate();
    const dayAfter = addDays(e, 1);
    let whole = (e.getUTCFullYear() - c.getUTCFullYear()) * 12 + (e.getUTCMonth() - c.getUTCMonth()) + 1;
    while (whole > 0 && addMonths(c, whole, anchor) > dayAfter) whole--;
    const start = addMonths(c, whole, anchor);
    let fraction = 0;
    if (start < dayAfter) {
      const next = addMonths(c, whole + 1, anchor);
      fraction = daysBetween(start, dayAfter) / daysBetween(start, next);
    }
    return { whole, fraction, months: whole + fraction };
  }

  // ---------------------------------------------------------------- inputs

  function num(v, fallback) {
    const n = typeof v === 'string' ? parseFloat(v.replace(/[, ]/g, '')) : Number(v);
    return isFinite(n) ? n : (fallback === undefined ? 0 : fallback);
  }

  // Parses "1-3, 13, 25-26" into a sorted list of unique month numbers.
  function parseMonthList(text) {
    const months = new Set();
    const errors = [];
    // "1 - 3" and "1-3" are the same range.
    String(text || '').replace(/\s*-\s*/g, '-').split(/[,;\s]+/).filter(Boolean).forEach((tok) => {
      const range = /^(\d+)\s*-\s*(\d+)$/.exec(tok);
      if (range) {
        const a = +range[1], b = +range[2];
        if (a < 1 || b < a || b - a > MAX_TERM_MONTHS) { errors.push(tok); return; }
        for (let i = a; i <= b; i++) months.add(i);
      } else if (/^\d+$/.test(tok) && +tok >= 1) {
        months.add(+tok);
      } else {
        errors.push(tok);
      }
    });
    return { months: [...months].sort((a, b) => a - b), errors };
  }

  function tenantOpexShare(opt) {
    if (opt.rateType === 'NNN') return 1;
    if (opt.rateType === 'FS') return 0;
    return Math.min(100, Math.max(0, num(opt.mgTenantShare))) / 100;
  }

  function blank(v) {
    return v === '' || v === null || v === undefined || !isFinite(Number(v));
  }

  // Number of lease months (a partial final month counts as one).
  function monthCountFromDates(commencementISO, expirationISO) {
    const t = termFromDates(commencementISO, expirationISO);
    return t.whole + (t.fraction > 1e-9 ? 1 : 0);
  }

  function monthsText(n) {
    const r = Math.round(n * 100) / 100;
    const shown = Number.isInteger(r) ? String(r) : r.toFixed(2);
    return `${shown} ${r === 1 ? 'month' : 'months'}`;
  }

  function validateOption(opt) {
    const errors = {};
    const c = parseDate(opt.commencement), e = parseDate(opt.expiration);
    if (!(num(opt.size) > 0)) errors.size = 'Enter the size of the space.';
    if (!c) errors.commencement = 'Enter a commencement date.';
    const termInput = termMonthsFromInput(opt.termValue, opt.termUnit);
    if (termInput > MAX_TERM_MONTHS) {
      errors.termValue = 'Term cannot exceed 50 years (600 months).';
    } else if (c && termInput > 0 && !expirationFromTerm(opt.commencement, termInput)) {
      errors.termValue = 'The ending date would fall after the year 9999.';
    }
    if (!e) errors.expiration = 'Enter an ending date, or a term length to set one.';
    else if (c && e < c) errors.expiration = 'The ending date must fall after the commencement date.';
    else if (c && termFromDates(opt.commencement, opt.expiration).months > MAX_TERM_MONTHS) {
      errors.expiration = 'Term cannot exceed 50 years (600 months).';
    }
    const datesOk = c && e && !errors.expiration;
    // Lease months (a partial final month is a month) and the exact term.
    const months = datesOk ? monthCountFromDates(opt.commencement, opt.expiration) : 0;
    const termLength = datesOk ? Math.round(termFromDates(opt.commencement, opt.expiration).months * 1e6) / 1e6 : 0;

    if (blank(opt.baseRate)) errors.baseRate = 'Enter the base rent.';
    else if (num(opt.baseRate) < 0) errors.baseRate = 'Base rent cannot be negative.';
    if (opt.escalationType !== 'fixed' && num(opt.escalation) <= -100) errors.escalation = 'Escalation must be above -100%.';
    if (num(opt.opex) < 0) errors.opex = 'OpEx cannot be negative.';
    if (num(opt.opexIncrease) <= -100) errors.opexIncrease = 'The increase must be above -100%.';
    if (opt.rateType === 'MG' && (blank(opt.mgTenantShare) || num(opt.mgTenantShare) < 0 || num(opt.mgTenantShare) > 100)) {
      errors.mgTenantShare = 'Enter a share from 0% to 100%.';
    }
    if (opt.freePlacement === 'custom') {
      const parsed = parseMonthList(opt.freeCustom);
      const after = months ? parsed.months.filter((m) => m > months) : [];
      if (parsed.errors.length) errors.freeCustom = `Not a month number: ${parsed.errors.join(', ')}`;
      else if (after.length) {
        const list = after.length > 4 ? `${after.slice(0, 3).join(', ')} and ${after.length - 3} more` : after.join(', ');
        errors.freeCustom = `${after.length === 1 ? 'Month' : 'Months'} ${list} ${after.length === 1 ? 'falls' : 'fall'} after the last lease month (month ${months}).`;
      }
    } else if (num(opt.freeMonths) < 0) errors.freeMonths = 'Free rent cannot be negative.';
    else if (termLength && num(opt.freeMonths) > termLength + 1e-9) {
      errors.freeMonths = `Free rent is longer than the term (${monthsText(termLength)}).`;
    }
    if (num(opt.parkingSpaces) < 0) errors.parkingSpaces = 'Spaces cannot be negative.';
    if (num(opt.parkingRate) < 0) errors.parkingRate = 'Cost cannot be negative.';
    if (num(opt.parkingIncrease) <= -100) errors.parkingIncrease = 'The increase must be above -100%.';
    if (num(opt.depositMonths) < 0) errors.depositMonths = 'Months cannot be negative.';
    return errors;
  }

  // ---------------------------------------------------------------- engine

  // Annual base rent for the whole premises in a given lease year (1-based).
  function annualBaseRent(opt, leaseYear) {
    const basis = RATE_BASES[opt.rateBasis] || RATE_BASES.psf_yr;
    const size = num(opt.size);
    const start = num(opt.baseRate);
    const esc = num(opt.escalation);
    const rate = opt.escalationType === 'fixed'
      ? start + esc * (leaseYear - 1)
      : start * Math.pow(1 + esc / 100, leaseYear - 1);
    const annualPerUnit = Math.max(0, rate) * (12 / basis.monthsPerUnit);
    return basis.perArea ? annualPerUnit * size : annualPerUnit;
  }

  // Fraction of each lease month (index 0 = month 1) that is abated.
  // `prorations` holds each month's share of a full month (1 except a partial
  // final month). Free months are counted in full months of rent, so a free
  // month placed on a half-length final month also abates half of the month
  // before it.
  function freeRentFractions(opt, prorations) {
    const monthCount = prorations.length;
    const fr = new Array(monthCount).fill(0);
    if (opt.freePlacement === 'custom') {
      parseMonthList(opt.freeCustom).months.forEach((m) => { if (m <= monthCount) fr[m - 1] = 1; });
      return fr;
    }
    let remaining = Math.max(0, num(opt.freeMonths));
    const fromEnd = opt.freePlacement === 'end';
    for (let k = 0; k < monthCount && remaining > 1e-9; k++) {
      const i = fromEnd ? monthCount - 1 - k : k;
      const p = prorations[i] || 1;
      const f = Math.min(1, remaining / p);
      fr[i] = f;
      remaining -= f * p;
    }
    return fr;
  }

  function round2(x) {
    return Math.round((x + Number.EPSILON) * 100) / 100;
  }

  /*
   * Builds the month-by-month schedule and summary for one lease option.
   * Returns { ok, errors, months, years, summary }.
   *
   * OpEx handling by rate type (s = tenant share of OpEx):
   *   NNN            s = 1: tenant pays all OpEx on top of base rent.
   *   Full Service   s = 0: OpEx sits inside the rent; with base year on, the
   *                  tenant pays increases over lease year 1.
   *   Modified Gross s = user share: tenant pays s of OpEx directly; the
   *                  landlord's (1 - s) portion sits inside the rent, with
   *                  increases over the base year passed through if enabled.
   */
  function calculate(rawOpt, settings) {
    const opt = Object.assign({}, DEFAULT_OPTION, rawOpt);
    const errors = validateOption(opt);
    if (Object.keys(errors).length) return { ok: false, errors, months: [], years: [], summary: null };

    const discountRate = Math.max(0, num(settings && settings.discountRate, 0)) / 100;
    const c = parseDate(opt.commencement);
    const e = parseDate(opt.expiration);
    const anchor = c.getUTCDate();
    const size = num(opt.size);
    const share = tenantOpexShare(opt);
    const baseYear = opt.rateType !== 'NNN' && !!opt.baseYear;
    const opex1 = num(opt.opex);
    const opexGrowth = num(opt.opexIncrease) / 100;
    const spaces = num(opt.parkingSpaces);
    const parkingRate = num(opt.parkingRate);
    const parkingGrowth = num(opt.parkingIncrease) / 100;
    const abateOpex = opt.freeAppliesTo === 'gross' || opt.freeAppliesTo === 'all';
    const abateParking = opt.freeAppliesTo === 'all';

    const periods = [];
    for (let i = 0; i < MAX_TERM_MONTHS + 1; i++) {
      const start = addMonths(c, i, anchor);
      if (start > e) break;
      const next = addMonths(c, i + 1, anchor);
      let end = addDays(next, -1);
      let proration = 1;
      if (end > e) {
        proration = (daysBetween(start, e) + 1) / daysBetween(start, next);
        end = e;
      }
      periods.push({ start, end, proration });
    }
    const free = freeRentFractions(opt, periods.map((p) => p.proration));
    const monthlyDiscount = Math.pow(1 + discountRate, 1 / 12);

    let cumulative = 0, npv = 0;
    const months = periods.map((p, idx) => {
      const month = idx + 1;
      const leaseYear = Math.floor(idx / 12) + 1;
      const annualBase = annualBaseRent(opt, leaseYear);
      const baseRent = annualBase / 12 * p.proration;

      const opexPsf = opex1 * Math.pow(1 + opexGrowth, leaseYear - 1);
      const tenantOpexPsf = share * opexPsf + (baseYear ? (1 - share) * Math.max(0, opexPsf - opex1) : 0);
      const opex = tenantOpexPsf * size / 12 * p.proration;

      const parking = spaces * parkingRate * Math.pow(1 + parkingGrowth, leaseYear - 1) * p.proration;

      const f = free[idx];
      const baseAbatement = -baseRent * f;
      const opexAbatement = abateOpex ? -opex * f : 0;
      const parkingAbatement = abateParking ? -parking * f : 0;
      const freeRent = baseAbatement + opexAbatement + parkingAbatement;

      const total = baseRent + opex + parking + freeRent;
      cumulative += total;
      npv += total / Math.pow(monthlyDiscount, idx);

      return {
        month,
        leaseYear,
        start: toISO(p.start),
        end: toISO(p.end),
        proration: p.proration,
        freeFraction: f,
        baseRatePsfYr: size > 0 ? annualBase / size : 0,
        opexPsfYr: opexPsf,
        tenantOpexPsfYr: tenantOpexPsf,
        baseRent,
        baseAbatement,
        opexAbatement,
        parkingAbatement,
        freeRent,
        netBaseRent: baseRent + baseAbatement,
        opex,
        parking,
        total,
        cumulative,
      };
    });

    const years = [];
    months.forEach((m) => {
      let y = years[m.leaseYear - 1];
      if (!y) {
        y = years[m.leaseYear - 1] = {
          leaseYear: m.leaseYear, start: m.start, end: m.end, months: 0,
          baseRatePsfYr: m.baseRatePsfYr, baseRent: 0, freeRent: 0, netBaseRent: 0,
          opex: 0, parking: 0, total: 0,
        };
      }
      y.end = m.end;
      y.months += m.proration;
      y.baseRent += m.baseRent;
      y.freeRent += m.freeRent;
      y.netBaseRent += m.netBaseRent;
      y.opex += m.opex;
      y.parking += m.parking;
      y.total += m.total;
    });

    const sum = (key) => months.reduce((acc, m) => acc + m[key], 0);
    const term = termFromDates(opt.commencement, opt.expiration);
    const termMonths = term.months;
    const termYears = termMonths / 12;
    const totalCost = sum('total');
    const last = months[months.length - 1];

    let securityDeposit = 0;
    const depositMonths = Math.max(0, num(opt.depositMonths, 1));
    if (opt.depositBasis === 'first') securityDeposit = annualBaseRent(opt, 1) / 12 * depositMonths;
    else if (opt.depositBasis === 'last') securityDeposit = annualBaseRent(opt, last.leaseYear) / 12 * depositMonths;

    const summary = {
      termMonths,
      termYears,
      monthCount: months.length,
      finalMonthProration: last.proration,
      size,
      tenantOpexShare: share,
      baseYear,
      startingRatePsfYr: months[0].baseRatePsfYr,
      endingRatePsfYr: last.baseRatePsfYr,
      startingMonthlyBase: annualBaseRent(opt, 1) / 12,
      // Free rent in full months of rent (a partial month counts by its share).
      freeMonths: Math.round(free.reduce((a, f, i) => a + f * periods[i].proration, 0) * 1e6) / 1e6,
      totalBaseRent: sum('baseRent'),
      totalFreeRent: sum('freeRent'),
      baseAbatement: sum('baseAbatement'),
      opexAbatement: sum('opexAbatement'),
      parkingAbatement: sum('parkingAbatement'),
      netBaseRent: sum('netBaseRent'),
      totalOpex: sum('opex'),
      totalParking: sum('parking'),
      netOpex: sum('opex') + sum('opexAbatement'),
      netParking: sum('parking') + sum('parkingAbatement'),
      totalCost,
      avgMonthlyCost: termMonths > 0 ? totalCost / termMonths : 0,
      avgAnnualCost: termYears > 0 ? totalCost / termYears : 0,
      effectiveRentPsfYr: termYears > 0 && size > 0 ? totalCost / size / termYears : 0,
      effectiveBasePsfYr: termYears > 0 && size > 0 ? sum('netBaseRent') / size / termYears : 0,
      npv,
      discountRate,
      securityDeposit,
      depositBasis: opt.depositBasis,
      depositMonths,
    };

    return { ok: true, errors: {}, months, years, summary };
  }

  return {
    MAX_TERM_MONTHS,
    RATE_TYPES,
    RATE_BASES,
    DEFAULT_OPTION,
    parseDate,
    toISO,
    addMonths,
    addDays,
    termMonthsFromInput,
    expirationFromTerm,
    termFromDates,
    parseMonthList,
    tenantOpexShare,
    validateOption,
    monthCountFromDates,
    annualBaseRent,
    freeRentFractions,
    calculate,
    round2,
  };
});
