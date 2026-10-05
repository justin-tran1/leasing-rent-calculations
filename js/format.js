/*
 * Display formatting and labels shared by the UI, the charts and the Excel
 * export. Browser global: window.LeaseFormat. Node: module.exports.
 */
(function (root, factory) {
  const api = factory();
  if (typeof module === 'object' && module.exports) module.exports = api;
  else root.LeaseFormat = api;
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  const MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];

  const RATE_TYPE_LABELS = { NNN: 'NNN', MG: 'Modified Gross', FS: 'Full Service' };
  const PLACEMENT_LABELS = { start: 'Start of term', end: 'End of term', custom: 'Custom months' };
  const APPLIES_LABELS = {
    base: 'Base rent only',
    gross: 'Base rent and OpEx',
    all: 'Base rent, OpEx and parking',
  };
  const DEPOSIT_LABELS = { first: "First month's base rent", last: "Last month's base rent", none: 'None' };

  function grouped(abs, decimals) {
    return abs.toLocaleString('en-US', { minimumFractionDigits: decimals, maximumFractionDigits: decimals });
  }

  // Negative amounts in parentheses, the finance convention also used in Excel.
  function money(v, cur, decimals) {
    const d = decimals === undefined ? 2 : decimals;
    const n = Number(v) || 0;
    const rounded = Math.abs(n) < Math.pow(10, -d) / 2 ? 0 : n;
    const s = (cur || '$') + grouped(Math.abs(rounded), d);
    return rounded < 0 ? `(${s})` : s;
  }

  function moneyCompact(v, cur) {
    const n = Number(v) || 0, a = Math.abs(n), c = cur || '$';
    let s;
    if (a >= 1e9) s = c + trim(a / 1e9) + 'B';
    else if (a >= 1e6) s = c + trim(a / 1e6) + 'M';
    else if (a >= 1e3) s = c + trim(a / 1e3) + 'K';
    else s = c + trim(a);
    return n < 0 ? '-' + s : s;
  }

  function trim(x) {
    return (Math.round(x * 10) / 10).toLocaleString('en-US', { maximumFractionDigits: 1 });
  }

  function number(v, decimals) {
    const d = decimals || 0;
    const n = Number(v) || 0;
    return (n < 0 ? '-' : '') + grouped(Math.abs(n), d);
  }

  function pct(v, decimals) {
    return number(v, decimals === undefined ? 1 : decimals) + '%';
  }

  function parts(iso) {
    const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(iso || '');
    return m ? { y: +m[1], m: +m[2] - 1, d: +m[3] } : null;
  }

  // Table style per CBRE writing guidance: three-letter month, no ordinal.
  function date(iso) {
    const p = parts(iso);
    return p ? `${MONTHS[p.m]} ${p.d}, ${p.y}` : '';
  }

  function monthYear(iso) {
    const p = parts(iso);
    return p ? `${MONTHS[p.m]} ${p.y}` : '';
  }

  function termLabel(months) {
    const m = Math.round(months * 100) / 100;
    if (m <= 0) return '';
    const years = m / 12;
    const monthsText = `${number(m, Number.isInteger(m) ? 0 : 2)} months`;
    if (Number.isInteger(m) && m % 12 === 0) return `${years} ${years === 1 ? 'year' : 'years'} (${monthsText})`;
    return monthsText;
  }

  function rateBasisLabel(basis, cur, unit) {
    const c = cur || '$', u = unit || 'SF';
    switch (basis) {
      case 'psf_mo': return `${c}/${u}/mo`;
      case 'monthly': return `${c}/month`;
      case 'annual': return `${c}/year`;
      default: return `${c}/${u}/yr`;
    }
  }

  function escalationLabel(opt, cur, unit) {
    if (opt.escalationType === 'fixed') {
      return `+${money(opt.escalation, cur)}${rateBasisLabel(opt.rateBasis, '', unit)} each year`;
    }
    return `${pct(opt.escalation)} per year`;
  }

  function depositLabel(basis, months) {
    const base = DEPOSIT_LABELS[basis] || DEPOSIT_LABELS.none;
    if (basis === 'none') return base;
    const m = Number(months);
    return m === 1 ? base : `${base} x ${number(m, Number.isInteger(m) ? 0 : 2)}`;
  }

  function freeRentLabel(opt, freeMonths) {
    if (!(freeMonths > 0)) return 'None';
    const n = number(freeMonths, Number.isInteger(freeMonths) ? 0 : 2);
    const count = `${n} ${freeMonths === 1 ? 'month' : 'months'}`;
    if (opt.freePlacement === 'custom') return `${count}: lease months ${String(opt.freeCustom).trim()}`;
    return `${count}, ${PLACEMENT_LABELS[opt.freePlacement].toLowerCase()}`;
  }

  return {
    MONTHS,
    RATE_TYPE_LABELS,
    PLACEMENT_LABELS,
    APPLIES_LABELS,
    DEPOSIT_LABELS,
    money,
    moneyCompact,
    number,
    pct,
    date,
    monthYear,
    termLabel,
    rateBasisLabel,
    escalationLabel,
    depositLabel,
    freeRentLabel,
  };
});
