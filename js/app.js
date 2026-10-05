/*
 * UI controller: state, the option form, comparison table, charts, schedule,
 * scenario save/open and the Excel export.
 */
(function () {
  'use strict';

  const Calc = window.LeaseCalc;
  const F = window.LeaseFormat;
  const Charts = window.LeaseCharts;

  const STORAGE_KEY = 'cbre-lease-calculator.v1';
  const MAX_OPTIONS = 6;
  const PALETTE = Charts.PRESETS.cbre.colors;
  const HEX = /^#[0-9a-fA-F]{6}$/;

  // Allowed values for every enumerated setting. Saved or imported data is
  // checked against these so a bad file can never break rendering.
  const OPTION_ENUMS = {
    rateType: Object.keys(F.RATE_TYPE_LABELS),
    rateBasis: Object.keys(Calc.RATE_BASES),
    termUnit: ['months', 'years'],
    escalationType: ['pct', 'fixed'],
    freePlacement: Object.keys(F.PLACEMENT_LABELS),
    freeAppliesTo: Object.keys(F.APPLIES_LABELS),
    depositBasis: Object.keys(F.DEPOSIT_LABELS),
  };
  const OPTION_NUMBERS = ['size', 'termValue', 'termMonths', 'baseRate', 'escalation', 'opex', 'opexIncrease',
    'mgTenantShare', 'freeMonths', 'parkingSpaces', 'parkingRate', 'parkingIncrease', 'depositMonths'];
  const OPTION_STRINGS = ['commencement', 'expiration', 'freeCustom'];
  const CHART_ENUMS = {
    metric: Object.keys(Charts.METRICS),
    layout: ['combined', 'separate'],
    type: Object.keys(Charts.TYPES),
    align: ['leaseMonth', 'calendar'],
    legend: ['bottom', 'top', 'right', 'none'],
    height: ['compact', 'standard', 'tall'],
    preset: Object.keys(Charts.PRESETS).concat('custom'),
  };
  const CHART_FLAGS = ['dataLabels', 'gridlines', 'beginAtZero', 'markers', 'sameScale'];
  const CURRENCIES = ['$', '£', '€', '¥'];
  const AREA_UNITS = ['SF', 'SM'];

  const $ = (sel, root) => (root || document).querySelector(sel);
  const $$ = (sel, root) => Array.from((root || document).querySelectorAll(sel));

  // ------------------------------------------------------------ helpers

  function el(tag, props) {
    const node = document.createElement(tag);
    if (props) {
      Object.keys(props).forEach((k) => {
        const v = props[k];
        if (v === null || v === undefined || v === false) return;
        if (k === 'class') node.className = v;
        else if (k === 'text') node.textContent = v;
        else if (k === 'style') Object.assign(node.style, v);
        else if (k.slice(0, 2) === 'on') node.addEventListener(k.slice(2), v);
        else node.setAttribute(k, v === true ? '' : v);
      });
    }
    for (let i = 2; i < arguments.length; i++) {
      [].concat(arguments[i]).forEach((c) => {
        if (c === null || c === undefined || c === false) return;
        node.appendChild(typeof c === 'string' ? document.createTextNode(c) : c);
      });
    }
    return node;
  }

  function uid() {
    return 'o' + Math.random().toString(36).slice(2, 9);
  }

  const round2 = (x) => Math.round(x * 100) / 100;
  const round4 = (x) => Math.round(x * 1e4) / 1e4;
  const pick = (v, allowed, dflt) => (allowed.includes(v) ? v : dflt);

  function localISO(d) {
    const pad = (x) => String(x).padStart(2, '0');
    return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
  }

  function firstOfNextMonth() {
    const now = new Date();
    return localISO(new Date(now.getFullYear(), now.getMonth() + 1, 1));
  }

  function displayName(o) {
    return String((o && o.name) || '').trim() || 'Untitled option';
  }

  function storageGet() {
    try { return window.localStorage.getItem(STORAGE_KEY); } catch (e) { return null; }
  }

  function storageSet(value) {
    try { window.localStorage.setItem(STORAGE_KEY, value); } catch (e) { /* private mode or blocked */ }
  }

  let toastTimer = null;
  function toast(message, isError) {
    const t = $('#toast');
    t.textContent = message;
    t.classList.toggle('toast--error', !!isError);
    t.hidden = false;
    clearTimeout(toastTimer);
    toastTimer = setTimeout(() => { t.hidden = true; }, isError ? 6000 : 4000);
  }

  function download(blob, name) {
    const url = URL.createObjectURL(blob);
    const a = el('a', { href: url, download: name });
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 2000);
  }

  function slug(text) {
    return String(text).toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/^-|-$/g, '').slice(0, 60) || 'chart';
  }

  // Re-rendering replaces buttons (tabs, swatches). Controls that are rebuilt
  // carry a data-focus-key so keyboard focus can return to the same control.
  function currentFocusKey() {
    const a = document.activeElement;
    return a && a.dataset ? a.dataset.focusKey || null : null;
  }

  function restoreFocus(key) {
    if (!key) return;
    const node = $$('[data-focus-key]').find((n) => n.dataset.focusKey === key);
    if (node && node !== document.activeElement) node.focus();
  }

  // ------------------------------------------------------ term syncing

  function termMonthsOf(opt) {
    return opt.termMonths > 0 ? opt.termMonths : Calc.termMonthsFromInput(opt.termValue, opt.termUnit);
  }

  // Ending date from commencement plus the exact term (fractions included).
  // Terms past the 50-year cap leave the date alone; validation flags them.
  function syncExpiration(opt) {
    const months = termMonthsOf(opt);
    if (!Calc.parseDate(opt.commencement) || !(months > 0)) return;
    const exp = Calc.expirationFromTerm(opt.commencement, months);
    if (exp) opt.expiration = exp;
  }

  function syncTermFromExpiration(opt) {
    const t = Calc.termFromDates(opt.commencement, opt.expiration);
    if (!(t.months > 0)) return;
    opt.termMonths = round4(t.months);
    if (t.fraction <= 1e-9 && opt.termUnit === 'years' && t.whole % 12 === 0) {
      opt.termValue = t.whole / 12;
    } else {
      opt.termUnit = 'months';
      opt.termValue = round2(t.months);
    }
  }

  // Switching units only changes how the term is shown; the exact month
  // count (and so the ending date) stays as it was.
  function setTermUnit(opt, unit) {
    const months = termMonthsOf(opt);
    opt.termUnit = unit;
    if (months > 0) opt.termValue = unit === 'years' ? round2(months / 12) : round2(months);
  }

  // ------------------------------------------------------------- state

  function nextColor(options) {
    const used = new Set(options.map((o) => String(o.color).toUpperCase()));
    return PALETTE.find((c) => !used.has(c.toUpperCase())) || PALETTE[options.length % PALETTE.length];
  }

  function nextName(options) {
    const names = new Set(options.map((o) => o.name));
    for (let i = 0; i < 26; i++) {
      const n = `Option ${String.fromCharCode(65 + i)}`;
      if (!names.has(n)) return n;
    }
    return `Option ${options.length + 1}`;
  }

  function makeOption(over) {
    const opt = Object.assign({}, Calc.DEFAULT_OPTION, { id: uid(), commencement: firstOfNextMonth() }, over);
    if (!(opt.termMonths > 0)) opt.termMonths = Calc.termMonthsFromInput(opt.termValue, opt.termUnit);
    if (!over || !over.expiration) syncExpiration(opt);
    return opt;
  }

  function defaultChart() {
    return {
      metric: 'monthlyTotal', layout: 'combined', type: 'line', align: 'leaseMonth', legend: 'bottom',
      height: 'standard', dataLabels: false, gridlines: true, beginAtZero: true, markers: false,
      sameScale: true, lineWidth: 2, title: '', preset: 'cbre', componentColors: Charts.COMPONENT_COLORS.slice(),
    };
  }

  function defaultState() {
    // Placeholder example terms so a first-time visitor sees a comparison.
    const a = makeOption({
      name: 'Option A', color: PALETTE[0], size: 10000, baseRate: 38, rateType: 'NNN', opex: 14.5,
      freeMonths: 3, parkingSpaces: 20, parkingRate: 150, depositBasis: 'first', termValue: 5, termUnit: 'years',
    });
    const b = makeOption({
      name: 'Option B', color: PALETTE[1], size: 10000, baseRate: 54, rateType: 'FS', opex: 14.5,
      freeMonths: 6, parkingSpaces: 20, parkingRate: 175, depositBasis: 'last', termValue: 7, termUnit: 'years',
    });
    return {
      settings: { currency: '$', areaUnit: 'SF', discountRate: 8 },
      options: [a, b],
      activeId: a.id,
      chart: defaultChart(),
      schedule: { optionId: a.id, view: 'monthly' },
    };
  }

  function normalizeNumber(v) {
    if (v === '' || v === null || v === undefined) return '';
    const n = typeof v === 'number' ? v : Number(String(v).replace(/[, ]/g, ''));
    return isFinite(n) ? n : '';
  }

  function normalizeOption(o, i) {
    const d = Calc.DEFAULT_OPTION;
    const src = o && typeof o === 'object' ? o : {};
    const opt = Object.assign({}, d);
    opt.id = typeof src.id === 'string' && src.id ? src.id.slice(0, 40) : uid();
    opt.name = String(src.name === undefined || src.name === null ? `Option ${i + 1}` : src.name).slice(0, 60);
    opt.color = HEX.test(src.color || '') ? src.color : PALETTE[i % PALETTE.length];
    Object.keys(OPTION_ENUMS).forEach((k) => { opt[k] = pick(src[k], OPTION_ENUMS[k], d[k]); });
    OPTION_NUMBERS.forEach((k) => { if (k in src) opt[k] = normalizeNumber(src[k]); });
    OPTION_STRINGS.forEach((k) => { if (typeof src[k] === 'string') opt[k] = src[k].slice(0, 200); });
    opt.baseYear = 'baseYear' in src ? !!src.baseYear : d.baseYear;
    if (!(opt.termMonths > 0)) {
      // Files from before termMonths existed: the saved dates are the truth.
      const t = Calc.termFromDates(opt.commencement, opt.expiration);
      opt.termMonths = t.months > 0 ? round4(t.months) : Calc.termMonthsFromInput(opt.termValue, opt.termUnit);
    }
    if (!Calc.parseDate(opt.expiration)) syncExpiration(opt);
    return opt;
  }

  // Accepts saved or imported data and replaces anything unexpected with the
  // defaults. Returns null when the data holds no lease options at all.
  function normalizeState(raw) {
    if (!raw || typeof raw !== 'object' || !Array.isArray(raw.options) || !raw.options.length) return null;
    const options = raw.options.slice(0, MAX_OPTIONS).map(normalizeOption);
    const ids = new Set();
    options.forEach((o) => { if (ids.has(o.id)) o.id = uid(); ids.add(o.id); });

    const rs = raw.settings && typeof raw.settings === 'object' ? raw.settings : {};
    const rate = Number(rs.discountRate);
    const settings = {
      currency: pick(rs.currency, CURRENCIES, '$'),
      areaUnit: pick(rs.areaUnit, AREA_UNITS, 'SF'),
      discountRate: rs.discountRate !== '' && isFinite(rate) && rate >= 0 && rate <= 50 ? rate : 8,
    };

    const rc = raw.chart && typeof raw.chart === 'object' ? raw.chart : {};
    const chart = defaultChart();
    Object.keys(CHART_ENUMS).forEach((k) => { chart[k] = pick(rc[k], CHART_ENUMS[k], chart[k]); });
    CHART_FLAGS.forEach((k) => { if (k in rc) chart[k] = !!rc[k]; });
    const lw = Number(rc.lineWidth);
    if (isFinite(lw) && lw >= 1 && lw <= 5) chart.lineWidth = lw;
    if (typeof rc.title === 'string') chart.title = rc.title.slice(0, 80);
    if (Array.isArray(rc.componentColors) && rc.componentColors.length === 3 && rc.componentColors.every((c) => HEX.test(c))) {
      chart.componentColors = rc.componentColors.slice();
    }

    const activeId = options.some((o) => o.id === raw.activeId) ? raw.activeId : options[0].id;
    const rsch = raw.schedule && typeof raw.schedule === 'object' ? raw.schedule : {};
    return {
      settings,
      options,
      activeId,
      chart,
      schedule: {
        optionId: options.some((o) => o.id === rsch.optionId) ? rsch.optionId : activeId,
        view: pick(rsch.view, ['monthly', 'annual'], 'monthly'),
      },
    };
  }

  function loadState() {
    const raw = storageGet();
    if (!raw) return null;
    try { return normalizeState(JSON.parse(raw)); } catch (e) { return null; }
  }

  let firstVisit = false;
  let state = loadState();
  if (!state) { state = defaultState(); firstVisit = true; }
  let results = new Map();

  function activeOption() {
    return state.options.find((o) => o.id === state.activeId) || state.options[0];
  }

  let saveTimer = null;
  function persist() {
    clearTimeout(saveTimer);
    saveTimer = setTimeout(() => storageSet(JSON.stringify(state)), 250);
  }

  // ------------------------------------------------------------- form

  const form = $('#option-form');

  // Pads stored numbers to the field's decimals (3 -> 3.0, 38 -> 38.00) but
  // never rounds away precision the user typed.
  function displayValue(input, v, opt) {
    if (v === null || v === undefined || v === '') return '';
    let d = input.dataset.decimals;
    if (input.name === 'escalation' && opt.escalationType === 'fixed') d = 2;
    if (d === undefined || !isFinite(Number(v))) return v;
    const places = (String(v).split('.')[1] || '').length;
    return places < Number(d) ? Number(v).toFixed(Number(d)) : String(v);
  }

  function fillForm(opt) {
    Array.from(form.elements).forEach((input) => {
      const name = input.name;
      if (!name || !(name in opt)) return;
      const v = opt[name];
      if (input.type === 'radio') input.checked = String(v) === input.value;
      else if (input.type === 'checkbox') input.checked = !!v;
      else input.value = displayValue(input, v, opt);
    });
    $('#option-panel').setAttribute('aria-labelledby', `tab-${opt.id}`);
    applyVisibility(opt);
    updateOpexNote(opt);
    $('#btn-remove').disabled = state.options.length <= 1;
    $('#btn-duplicate').disabled = state.options.length >= MAX_OPTIONS;
  }

  function applyVisibility(opt) {
    $$('[data-show]', form).forEach((node) => {
      const [key, values] = node.dataset.show.split(':');
      const show = values.split(',').includes(String(opt[key]));
      node.toggleAttribute('data-hidden', !show);
    });
  }

  function updateOpexNote(opt) {
    let text;
    if (opt.rateType === 'NNN') {
      text = 'NNN: the tenant pays all OpEx on top of base rent.';
    } else if (opt.rateType === 'MG') {
      text = 'Modified Gross: the tenant pays its share of OpEx directly and the landlord\'s share sits inside the rent.';
      if (opt.baseYear) text += ' Increases on the landlord\'s share over the base year (lease year 1) pass through.';
    } else {
      text = opt.baseYear
        ? 'Full Service: OpEx sits inside the rent. The tenant pays only increases over the base year (lease year 1).'
        : 'Full Service: OpEx sits inside the rent and the tenant pays no OpEx.';
    }
    $('#opex-note').textContent = text;
  }

  function onFormInput(e) {
    const input = e.target;
    const name = input.name;
    if (!name) return;
    const opt = activeOption();
    let value;
    if (input.type === 'checkbox') value = input.checked;
    else if (input.type === 'radio') { if (!input.checked) return; value = input.value; }
    else if (input.dataset.kind === 'number') value = input.value === '' ? '' : Number(input.value);
    else value = input.value;
    if (opt[name] === value) return;

    if (name === 'termUnit') {
      setTermUnit(opt, value);
      form.elements.termValue.value = opt.termValue;
    } else {
      opt[name] = value;
    }

    if (name === 'termValue') {
      opt.termMonths = Calc.termMonthsFromInput(value, opt.termUnit);
      syncExpiration(opt);
      form.elements.expiration.value = opt.expiration;
    } else if (name === 'commencement') {
      syncExpiration(opt);
      form.elements.expiration.value = opt.expiration;
    } else if (name === 'expiration') {
      syncTermFromExpiration(opt);
      form.elements.termValue.value = opt.termValue;
      form.elements.termUnit.value = opt.termUnit;
    }
    if (['rateType', 'freePlacement', 'depositBasis', 'baseYear'].includes(name)) {
      applyVisibility(opt);
      updateOpexNote(opt);
    }
    scheduleUpdate();
  }

  form.addEventListener('input', onFormInput);
  form.addEventListener('change', onFormInput);
  form.addEventListener('submit', (e) => e.preventDefault());

  // A cleared ending date is rebuilt from the term once the field loses
  // focus (not while typing, which would fight the date picker).
  form.elements.expiration.addEventListener('blur', () => {
    const opt = activeOption();
    if (Calc.parseDate(opt.expiration) || !(termMonthsOf(opt) > 0)) return;
    syncExpiration(opt);
    if (opt.expiration) {
      form.elements.expiration.value = opt.expiration;
      scheduleUpdate();
    }
  });

  function renderFieldErrors(result) {
    const errors = (result && result.errors) || {};
    $$('[data-error-for]', form).forEach((node) => {
      const key = node.dataset.errorFor;
      const msg = errors[key] || '';
      if (node.textContent !== msg) node.textContent = msg;
      const input = form.elements[key];
      if (!input || !input.setAttribute) return;
      const ids = [input.dataset.hint, msg ? node.id : null].filter(Boolean);
      if (ids.length) input.setAttribute('aria-describedby', ids.join(' '));
      else input.removeAttribute('aria-describedby');
      if (msg) input.setAttribute('aria-invalid', 'true');
      else input.removeAttribute('aria-invalid');
    });
  }

  function stat(label, value, sub) {
    return el('div', { class: 'stat' },
      el('div', { class: 'stat__label', text: label }),
      el('div', { class: 'stat__value', text: value }),
      sub ? el('div', { class: 'stat__sub', text: sub }) : null);
  }

  let lastSummary = '';
  function renderOptionSummary(opt, result) {
    const box = $('#option-summary');
    const cur = state.settings.currency, unit = state.settings.areaUnit;
    const deposit = $('#deposit-value');
    let depositText = '-';
    let stats = null;
    if (result && result.ok) {
      const s = result.summary;
      stats = [
        ['Total lease cost', F.money(s.totalCost, cur, 0), F.termLabel(s.termMonths)],
        ['Average monthly cost', F.money(s.avgMonthlyCost, cur, 0), `${F.money(s.startingMonthlyBase, cur, 0)} base rent in month 1`],
        [`Effective rent per ${unit}/yr`, F.money(s.effectiveRentPsfYr, cur, 2), 'All-in, net of free rent'],
        ['Free rent value', F.money(-s.totalFreeRent, cur, 0), F.freeRentLabel(opt, s.freeMonths)],
      ];
      depositText = opt.depositBasis === 'none' ? 'None' : F.money(s.securityDeposit, cur, 2);
    }
    if (deposit.textContent !== depositText) deposit.textContent = depositText;
    // Skip the rebuild when nothing changed (typing elsewhere re-renders often).
    const signature = JSON.stringify([opt.id, stats]);
    if (signature === lastSummary) return;
    lastSummary = signature;
    box.textContent = '';
    if (!stats) {
      box.className = 'option-summary option-summary--error';
      box.textContent = 'Complete the highlighted fields to see results for this option.';
      return;
    }
    box.className = 'option-summary';
    stats.forEach(([l, v, sub]) => box.appendChild(stat(l, v, sub)));
  }

  // ------------------------------------------------------------- tabs

  function renderTabs() {
    const bar = $('#option-tabs');
    bar.textContent = '';
    state.options.forEach((o) => {
      const r = results.get(o.id);
      const selected = o.id === state.activeId;
      bar.appendChild(el('button', {
        type: 'button', class: 'tab', role: 'tab', id: `tab-${o.id}`,
        'aria-selected': selected ? 'true' : 'false', 'aria-controls': 'option-panel',
        tabindex: selected ? '0' : '-1', 'data-focus-key': `tab:${o.id}`,
        onclick: () => setActive(o.id),
      },
      el('span', { class: 'tab__dot', style: { background: o.color }, 'aria-hidden': 'true' }),
      el('span', { class: 'tab__name', text: displayName(o) }),
      r && !r.ok ? el('span', { class: 'tab__warn', text: '!', title: 'Incomplete inputs', 'aria-hidden': 'true' }) : null,
      r && !r.ok ? el('span', { class: 'visually-hidden', text: ' (incomplete inputs)' }) : null));
    });
    const selectedTab = bar.querySelector('[aria-selected="true"]');
    if (selectedTab) {
      const b = bar.getBoundingClientRect(), t = selectedTab.getBoundingClientRect();
      if (t.left < b.left) bar.scrollLeft -= b.left - t.left + 8;
      else if (t.right > b.right) bar.scrollLeft += t.right - b.right + 8;
    }
    const add = $('#btn-add');
    add.disabled = state.options.length >= MAX_OPTIONS;
    add.title = add.disabled ? `Up to ${MAX_OPTIONS} options` : 'Add a lease option';
  }

  $('#option-tabs').addEventListener('keydown', (e) => {
    const keys = ['ArrowRight', 'ArrowLeft', 'Home', 'End'];
    if (!keys.includes(e.key)) return;
    const n = state.options.length;
    const idx = state.options.findIndex((o) => o.id === state.activeId);
    let next = idx;
    if (e.key === 'ArrowRight') next = (idx + 1) % n;
    else if (e.key === 'ArrowLeft') next = (idx - 1 + n) % n;
    else if (e.key === 'Home') next = 0;
    else next = n - 1;
    e.preventDefault();
    setActive(state.options[next].id);
    const tab = $(`#tab-${state.options[next].id}`);
    if (tab) tab.focus();
  });

  function setActive(id) {
    state.activeId = id;
    state.schedule.optionId = id;
    fillForm(activeOption());
    update();
  }

  function addOption() {
    if (state.options.length >= MAX_OPTIONS) return;
    const from = activeOption();
    const opt = makeOption({
      name: nextName(state.options),
      color: nextColor(state.options),
      size: from.size,
      commencement: from.commencement,
      termValue: from.termValue,
      termUnit: from.termUnit,
      termMonths: termMonthsOf(from),
      baseRate: '',
      opex: from.opex,
    });
    state.options.push(opt);
    setActive(opt.id);
    form.elements.baseRate.focus();
  }

  $('#btn-add').addEventListener('click', addOption);

  $('#btn-duplicate').addEventListener('click', () => {
    if (state.options.length >= MAX_OPTIONS) return;
    const from = activeOption();
    const copy = Object.assign({}, from, {
      id: uid(),
      name: `${displayName(from)} (copy)`.slice(0, 60),
      color: nextColor(state.options),
    });
    state.options.splice(state.options.indexOf(from) + 1, 0, copy);
    setActive(copy.id);
    toast(`Duplicated ${displayName(from)}.`);
  });

  $('#btn-remove').addEventListener('click', () => {
    if (state.options.length <= 1) return;
    const opt = activeOption();
    if (!window.confirm(`Remove ${displayName(opt)}? This cannot be undone.`)) return;
    const idx = state.options.indexOf(opt);
    state.options.splice(idx, 1);
    setActive(state.options[Math.max(0, idx - 1)].id);
    toast(`Removed ${displayName(opt)}.`);
  });

  // ---------------------------------------------------------- settings

  function renderUnitLabels() {
    const cur = state.settings.currency, unit = state.settings.areaUnit;
    $$('[data-cur]').forEach((n) => { n.textContent = cur; });
    $$('[data-area]').forEach((n) => { n.textContent = unit; });
    $$('[data-rate-unit]').forEach((n) => { n.textContent = `/${unit}/yr`; });
    const basis = form.elements.rateBasis;
    basis.textContent = '';
    [['psf_yr', `/${unit}/yr`], ['psf_mo', `/${unit}/mo`], ['monthly', '/month total'], ['annual', '/year total']]
      .forEach(([v, t]) => basis.appendChild(el('option', { value: v, text: t })));
    basis.value = activeOption().rateBasis;
  }

  function syncSettingsInputs() {
    $('#set-currency').value = state.settings.currency;
    $('#set-area').value = state.settings.areaUnit;
    const disc = $('#set-discount');
    disc.value = state.settings.discountRate;
    disc.removeAttribute('aria-invalid');
    $('#err-discount').textContent = '';
  }

  function bindSettings() {
    const cur = $('#set-currency'), area = $('#set-area'), disc = $('#set-discount');
    cur.addEventListener('change', () => { state.settings.currency = pick(cur.value, CURRENCIES, '$'); renderUnitLabels(); update(); });
    area.addEventListener('change', () => { state.settings.areaUnit = pick(area.value, AREA_UNITS, 'SF'); renderUnitLabels(); update(); });
    disc.addEventListener('input', () => {
      const v = Number(disc.value);
      const ok = disc.value !== '' && isFinite(v) && v >= 0 && v <= 50;
      $('#err-discount').textContent = ok ? '' : 'Enter a rate from 0% to 50%.';
      if (ok) disc.removeAttribute('aria-invalid');
      else disc.setAttribute('aria-invalid', 'true');
      if (!ok) return;
      state.settings.discountRate = v;
      scheduleUpdate();
    });
  }

  // -------------------------------------------------------- comparison

  function renderComparison() {
    const table = $('#compare-table');
    table.textContent = '';
    const cur = state.settings.currency, unit = state.settings.areaUnit;
    const entries = state.options.map((o) => ({ opt: o, r: results.get(o.id) }));
    const valid = entries.filter((e) => e.r && e.r.ok);
    const lowest = valid.length > 1
      ? valid.reduce((best, e) => (e.r.summary.totalCost < best.r.summary.totalCost ? e : best))
      : null;

    const head = el('tr', null, el('th', { scope: 'col', text: 'Lease option' }));
    entries.forEach((e) => {
      head.appendChild(el('th', { scope: 'col' },
        el('span', { class: 'col-key', style: { background: e.opt.color }, 'aria-hidden': 'true' }),
        displayName(e.opt),
        lowest && lowest === e ? [el('br'), el('span', { class: 'badge' }, '✓ Lowest total cost')] : null,
        e.r && !e.r.ok ? [el('br'), el('span', { class: 'badge badge--warn' }, '! Incomplete inputs')] : null));
    });
    table.appendChild(el('thead', null, head));

    const body = el('tbody');
    const cols = entries.length + 1;
    const section = (title) => body.appendChild(el('tr', { class: 'row-section' }, el('th', { colspan: cols, scope: 'colgroup', text: title })));
    const row = (label, get, opts) => {
      const tr = el('tr', { class: opts && opts.strong ? 'row-strong' : null }, el('th', { scope: 'row', text: label }));
      entries.forEach((e) => {
        if (!e.r || !e.r.ok) { tr.appendChild(el('td', { class: 'cell-muted', text: '-' })); return; }
        tr.appendChild(el('td', { text: get(e.r.summary, e.opt, e) }));
      });
      body.appendChild(tr);
    };

    section('Lease terms');
    row('Rate type', (s, o) => F.RATE_TYPE_LABELS[o.rateType]);
    row('Size', (s) => `${F.number(s.size)} ${unit}`);
    row('Term', (s) => F.termLabel(s.termMonths));
    row('Dates', (s, o) => `${F.date(o.commencement)} - ${F.date(o.expiration)}`);
    row(`Starting base rent (${cur}/${unit}/yr)`, (s) => F.money(s.startingRatePsfYr, cur));
    row(`Ending base rent (${cur}/${unit}/yr)`, (s) => F.money(s.endingRatePsfYr, cur));
    row('Annual escalation', (s, o) => F.escalationLabel(o, cur, unit));
    row(`OpEx (${cur}/${unit}/yr)`, (s, o) => `${F.money(o.opex, cur)}, +${F.pct(o.opexIncrease)} per year`);
    row('Free rent', (s, o) => F.freeRentLabel(o, s.freeMonths));
    row('Parking', (s, o) => (Number(o.parkingSpaces) > 0
      ? `${F.number(o.parkingSpaces)} spaces at ${F.money(o.parkingRate, cur)}/month` : 'None'));

    section('Cost over the term');
    row('Total base rent', (s) => F.money(s.totalBaseRent, cur, 0));
    row('Free rent', (s) => F.money(s.totalFreeRent, cur, 0));
    row('Net base rent', (s) => F.money(s.netBaseRent, cur, 0));
    row('OpEx', (s) => F.money(s.totalOpex, cur, 0));
    row('Parking', (s) => F.money(s.totalParking, cur, 0));
    row('Total lease cost', (s) => F.money(s.totalCost, cur, 0), { strong: true });
    if (lowest) {
      row('Difference from lowest', (s, o, e) => {
        if (e === lowest) return 'Lowest';
        const d = s.totalCost - lowest.r.summary.totalCost;
        return `+${F.money(d, cur, 0)} (+${F.pct(d / lowest.r.summary.totalCost * 100)})`;
      });
    }
    row('Average monthly cost', (s) => F.money(s.avgMonthlyCost, cur, 0));
    row('Average annual cost', (s) => F.money(s.avgAnnualCost, cur, 0));
    row(`Effective rent (${cur}/${unit}/yr)`, (s) => F.money(s.effectiveRentPsfYr, cur));
    row(`NPV at ${F.pct(state.settings.discountRate, 2).replace(/\.?0+%$/, '%')}`, (s) => F.money(s.npv, cur, 0));

    section('Security deposit');
    row('Based on', (s, o) => F.depositLabel(o.depositBasis, s.depositMonths));
    row('Security deposit', (s, o) => (o.depositBasis === 'none' ? 'None' : F.money(s.securityDeposit, cur, 2)), { strong: true });
    table.appendChild(body);

    const meta = $('#compare-meta');
    if (lowest) meta.textContent = `${displayName(lowest.opt)} has the lowest total lease cost.`;
    else if (valid.length === 1 && entries.length === 1) meta.textContent = 'Add an option to compare proposals side by side.';
    else meta.textContent = '';
  }

  // ------------------------------------------------------------ charts

  const chartArea = $('#chart-area');

  function chartSeries() {
    return state.options
      .map((o) => ({ id: o.id, name: displayName(o), color: o.color, result: results.get(o.id) }))
      .filter((s) => s.result && s.result.ok);
  }

  function isLineType(t) {
    return t === 'line' || t === 'stepped' || t === 'area';
  }

  // Only write when different so typing in a field never loses the caret.
  function setVal(node, v) {
    if (node.value !== String(v)) node.value = v;
  }

  function renderChartControls() {
    const c = state.chart;
    const metricSel = $('#c-metric');
    if (!metricSel.options.length) {
      Object.keys(Charts.METRICS).forEach((k) => metricSel.appendChild(el('option', { value: k, text: Charts.METRICS[k].label })));
      Object.keys(Charts.PRESETS).forEach((k) => $('#c-preset').appendChild(el('option', { value: k, text: Charts.PRESETS[k].label })));
      $('#c-preset').appendChild(el('option', { value: 'custom', text: 'Custom' }));
    }
    setVal(metricSel, c.metric);
    const types = Charts.typesFor(c.metric);
    if (!types.includes(c.type)) c.type = types[0];
    const typeSel = $('#c-type');
    if (typeSel.dataset.types !== types.join()) {
      typeSel.textContent = '';
      types.forEach((t) => typeSel.appendChild(el('option', { value: t, text: Charts.TYPES[t] })));
      typeSel.dataset.types = types.join();
    }
    setVal(typeSel, c.type);
    $$('input[name="c-layout"]').forEach((r) => { r.checked = r.value === c.layout; });
    setVal($('#c-align'), c.align);
    setVal($('#c-title'), c.title);
    $('#c-title').placeholder = Charts.defaultTitle(c.metric, state.settings.areaUnit);
    setVal($('#c-preset'), c.preset);
    setVal($('#c-legend'), c.legend);
    setVal($('#c-height'), c.height);
    setVal($('#c-linewidth'), c.lineWidth);
    $('#c-linewidth-out').textContent = `${c.lineWidth}px`;
    $('#c-labels').checked = !!c.dataLabels;
    $('#c-grid').checked = c.gridlines !== false;
    $('#c-zero').checked = !!c.beginAtZero;
    $('#c-markers').checked = !!c.markers;
    $('#c-samescale').checked = c.sameScale !== false;

    const kind = Charts.METRICS[c.metric].kind;
    const show = {
      monthly: kind === 'monthly',
      line: isLineType(c.type),
      separate: c.layout === 'separate' && kind !== 'totals',
    };
    $$('[data-chart-show]').forEach((n) => n.toggleAttribute('data-hidden', !show[n.dataset.chartShow]));
    renderColourList();
  }

  function swatchRow(current, onPick, rowKey) {
    const wrap = el('div', { class: 'swatches', role: 'group', 'aria-label': 'CBRE colours' });
    Charts.SWATCHES.forEach((hex) => {
      wrap.appendChild(el('button', {
        type: 'button', class: 'swatch', style: { background: hex }, title: hex,
        'aria-label': `Use ${hex}`, 'aria-pressed': hex.toUpperCase() === String(current).toUpperCase() ? 'true' : 'false',
        'data-focus-key': `sw:${rowKey}:${hex}`,
        onclick: () => onPick(hex),
      }));
    });
    return wrap;
  }

  function colourRow(name, color, onPick, rowKey) {
    const input = el('input', { type: 'color', value: color, 'aria-label': `Colour for ${name}`, 'data-focus-key': `ci:${rowKey}` });
    input.addEventListener('input', () => onPick(input.value.toUpperCase(), true));
    input.addEventListener('change', () => onPick(input.value.toUpperCase()));
    return el('div', { class: 'colour-row' },
      input,
      el('span', { class: 'colour-row__name', text: name, title: name }),
      swatchRow(color, (hex) => onPick(hex), rowKey));
  }

  function renderColourList() {
    const list = $('#colour-list');
    list.textContent = '';
    if (state.chart.metric === 'breakdown') {
      Charts.COMPONENTS.forEach((comp, i) => {
        list.appendChild(colourRow(comp.label, state.chart.componentColors[i], (hex, live) => {
          state.chart.componentColors[i] = hex;
          state.chart.preset = 'custom';
          afterColourChange(live);
        }, `comp${i}`));
      });
      return;
    }
    state.options.forEach((o) => {
      list.appendChild(colourRow(displayName(o), o.color, (hex, live) => {
        o.color = hex;
        state.chart.preset = 'custom';
        afterColourChange(live);
      }, o.id));
    });
  }

  // While dragging the native colour picker only redraw the charts so the
  // picker keeps focus; rebuild the colour list once the choice settles.
  function afterColourChange(live) {
    if (live) { renderCharts(); renderTabs(); persist(); return; }
    update();
  }

  function renderCharts() {
    const series = chartSeries();
    const note = $('#chart-note');
    const notes = [];
    const invalid = state.options.filter((o) => { const r = results.get(o.id); return r && !r.ok; });
    if (invalid.length) notes.push(`${invalid.map(displayName).join(', ')} ${invalid.length === 1 ? 'is' : 'are'} left out until the inputs are complete.`);
    if (series.length === 1) notes.push('Add another option to compare.');
    if (state.chart.layout === 'separate' && Charts.METRICS[state.chart.metric].kind === 'totals') {
      notes.push('Totals compare every option on one chart.');
    }
    note.textContent = notes.join(' ');
    note.hidden = !notes.length;

    if (!window.Chart) {
      chartArea.textContent = '';
      chartArea.appendChild(el('div', { class: 'chart-empty', text: 'Charts could not load. Check that vendor/chart.umd.min.js is present.' }));
      return;
    }
    if (!series.length) {
      Charts.destroy();
      chartArea.textContent = '';
      chartArea.classList.remove('chart-grid');
      chartArea.appendChild(el('div', { class: 'chart-empty', text: 'Complete at least one lease option to see charts.' }));
      return;
    }
    Charts.render(chartArea, series, state.chart, state.settings);
  }

  // Handlers read state.chart at event time: opening a scenario replaces it.
  function bindChartControls() {
    const c = () => state.chart;
    const on = (sel, evt, fn) => $(sel).addEventListener(evt, fn);
    on('#c-metric', 'change', (e) => {
      c().metric = pick(e.target.value, CHART_ENUMS.metric, 'monthlyTotal');
      const types = Charts.typesFor(c().metric);
      if (!types.includes(c().type)) c().type = types[0];
      update();
    });
    on('#c-type', 'change', (e) => { c().type = e.target.value; update(); });
    $$('input[name="c-layout"]').forEach((r) => r.addEventListener('change', () => { if (r.checked) { c().layout = r.value; update(); } }));
    on('#c-align', 'change', (e) => { c().align = e.target.value; update(); });
    on('#c-title', 'input', (e) => { c().title = e.target.value; scheduleUpdate(); });
    on('#c-legend', 'change', (e) => { c().legend = e.target.value; update(); });
    on('#c-height', 'change', (e) => { c().height = e.target.value; update(); });
    on('#c-linewidth', 'input', (e) => {
      c().lineWidth = Number(e.target.value);
      $('#c-linewidth-out').textContent = `${c().lineWidth}px`;
      scheduleUpdate();
    });
    on('#c-labels', 'change', (e) => { c().dataLabels = e.target.checked; update(); });
    on('#c-grid', 'change', (e) => { c().gridlines = e.target.checked; update(); });
    on('#c-zero', 'change', (e) => { c().beginAtZero = e.target.checked; update(); });
    on('#c-markers', 'change', (e) => { c().markers = e.target.checked; update(); });
    on('#c-samescale', 'change', (e) => { c().sameScale = e.target.checked; update(); });
    on('#c-preset', 'change', (e) => {
      c().preset = e.target.value;
      const preset = Charts.PRESETS[c().preset];
      if (preset) {
        state.options.forEach((o, i) => { o.color = preset.colors[i % preset.colors.length]; });
        c().componentColors = preset.colors.slice(0, 3);
      }
      update();
    });
    on('#btn-png', 'click', () => {
      const shots = Charts.snapshots();
      if (!shots.length) { toast('There is no chart to download yet.', true); return; }
      shots.forEach((s, i) => {
        setTimeout(() => {
          fetch(s.dataUrl).then((r) => r.blob()).then((b) => download(b, `CBRE-${slug(s.title)}.png`))
            .catch(() => { const a = el('a', { href: s.dataUrl, download: `CBRE-${slug(s.title)}.png` }); a.click(); });
        }, i * 250);
      });
    });
  }

  // ---------------------------------------------------------- schedule

  function renderSchedule() {
    const sel = $('#s-option');
    sel.textContent = '';
    state.options.forEach((o) => sel.appendChild(el('option', { value: o.id, text: displayName(o) })));
    if (!state.options.some((o) => o.id === state.schedule.optionId)) state.schedule.optionId = state.activeId;
    sel.value = state.schedule.optionId;
    $$('input[name="s-view"]').forEach((r) => { r.checked = r.value === state.schedule.view; });

    const table = $('#schedule-table');
    table.textContent = '';
    const opt = state.options.find((o) => o.id === state.schedule.optionId);
    const r = results.get(opt.id);
    const cur = state.settings.currency, unit = state.settings.areaUnit;
    if (!r || !r.ok) {
      table.appendChild(el('tbody', null, el('tr', null, el('td', { class: 'cell-invalid', text: `Complete the inputs for ${displayName(opt)} to see its schedule.` }))));
      return;
    }
    const m = (v) => F.money(v, cur);
    const annual = state.schedule.view === 'annual';
    const heads = annual
      ? ['Lease year', 'Months', 'Period', `Base rent (${cur}/${unit}/yr)`, 'Base rent', 'Free rent', 'OpEx', 'Parking', 'Total annual cost', 'Cumulative cost']
      : ['Month', 'Lease year', 'Period', `Base rent (${cur}/${unit}/yr)`, 'Base rent', 'Free rent', 'OpEx', 'Parking', 'Total monthly cost', 'Cumulative cost'];
    table.appendChild(el('thead', null, el('tr', null, heads.map((h) => el('th', { scope: 'col', text: h })))));
    const body = el('tbody');
    if (annual) {
      let cum = 0;
      r.years.forEach((y) => {
        cum += y.total;
        const months = round2(y.months);
        body.appendChild(el('tr', null,
          el('td', { text: `Year ${y.leaseYear}` }),
          el('td', { text: F.number(months, Number.isInteger(months) ? 0 : 2) }),
          el('td', { text: `${F.date(y.start)} - ${F.date(y.end)}` }),
          el('td', { text: m(y.baseRatePsfYr) }),
          el('td', { text: m(y.baseRent) }),
          el('td', { class: y.freeRent < 0 ? 'cell-free' : 'cell-muted', text: m(y.freeRent) }),
          el('td', { text: m(y.opex) }),
          el('td', { text: m(y.parking) }),
          el('td', { text: m(y.total) }),
          el('td', { text: m(cum) })));
      });
    } else {
      r.months.forEach((row) => {
        body.appendChild(el('tr', null,
          el('td', { text: String(row.month) }),
          el('td', { text: String(row.leaseYear) }),
          el('td', { text: `${F.date(row.start)} - ${F.date(row.end)}${row.proration < 1 ? ` (${F.pct(row.proration * 100)})` : ''}` }),
          el('td', { text: m(row.baseRatePsfYr) }),
          el('td', { text: m(row.baseRent) }),
          el('td', { class: row.freeRent < 0 ? 'cell-free' : 'cell-muted', text: m(row.freeRent) }),
          el('td', { text: m(row.opex) }),
          el('td', { text: m(row.parking) }),
          el('td', { text: m(row.total) }),
          el('td', { text: m(row.cumulative) })));
      });
    }
    table.appendChild(body);
    const s = r.summary;
    const termMonths = round2(s.termMonths);
    table.appendChild(el('tfoot', null, el('tr', null,
      el('td', { text: 'Total' }),
      el('td', { text: annual ? F.number(termMonths, Number.isInteger(termMonths) ? 0 : 2) : '' }),
      el('td', { text: `${F.date(opt.commencement)} - ${F.date(opt.expiration)}` }),
      el('td', { text: '' }),
      el('td', { text: m(s.totalBaseRent) }),
      el('td', { text: m(s.totalFreeRent) }),
      el('td', { text: m(s.totalOpex) }),
      el('td', { text: m(s.totalParking) }),
      el('td', { text: m(s.totalCost) }),
      el('td', { text: '' }))));
  }

  $('#s-option').addEventListener('change', (e) => { state.schedule.optionId = e.target.value; renderSchedule(); persist(); });
  $$('input[name="s-view"]').forEach((r) => r.addEventListener('change', () => {
    if (r.checked) { state.schedule.view = r.value; renderSchedule(); persist(); }
  }));

  // ------------------------------------------------------------ update

  function computeAll() {
    results = new Map(state.options.map((o) => [o.id, Calc.calculate(o, state.settings)]));
  }

  function update() {
    const focusKey = currentFocusKey();
    computeAll();
    const opt = activeOption();
    const r = results.get(opt.id);
    renderTabs();
    renderFieldErrors(r);
    renderOptionSummary(opt, r);
    renderComparison();
    renderChartControls();
    renderCharts();
    renderSchedule();
    persist();
    restoreFocus(focusKey);
  }

  let pending = false;
  function scheduleUpdate() {
    if (pending) return;
    pending = true;
    requestAnimationFrame(() => { pending = false; update(); });
  }

  // Pushes the whole state into the page (first load and after opening a file).
  function applyState() {
    lastSummary = '';
    syncSettingsInputs();
    renderUnitLabels();
    fillForm(activeOption());
    update();
  }

  // --------------------------------------------------- save / open / export

  $('#btn-save').addEventListener('click', () => {
    const data = {
      app: 'cbre-lease-calculator',
      version: 1,
      savedAt: new Date().toISOString(),
      settings: state.settings,
      options: state.options,
      chart: state.chart,
    };
    download(new Blob([JSON.stringify(data, null, 2)], { type: 'application/json' }),
      `CBRE-lease-scenario-${localISO(new Date())}.json`);
    toast('Scenario saved. Open it later with Open scenario.');
  });

  const fileInput = $('#file-open');
  $('#btn-open').addEventListener('click', () => fileInput.click());

  fileInput.addEventListener('change', (e) => {
    const file = e.target.files && e.target.files[0];
    e.target.value = '';
    if (!file) return;
    const reader = new FileReader();
    reader.onload = () => {
      let next = null;
      try { next = normalizeState(JSON.parse(reader.result)); } catch (err) { next = null; }
      if (!next) { toast('That file is not a lease scenario saved from this calculator.', true); return; }
      const prev = state;
      try {
        state = next;
        applyState();
        toast(`Opened ${file.name}: ${state.options.length} ${state.options.length === 1 ? 'option' : 'options'}.`);
      } catch (err) {
        console.error(err);
        state = prev;
        applyState();
        toast('That scenario could not be opened. Your current options are unchanged.', true);
      }
    };
    reader.onerror = () => toast('The file could not be read.', true);
    reader.readAsText(file);
  });

  function loadExcelJS() {
    if (window.ExcelJS) return Promise.resolve(window.ExcelJS);
    return new Promise((resolve, reject) => {
      const s = el('script', { src: 'vendor/exceljs.min.js' });
      s.onload = () => (window.ExcelJS ? resolve(window.ExcelJS) : reject(new Error('ExcelJS did not load')));
      s.onerror = () => reject(new Error('ExcelJS did not load'));
      document.head.appendChild(s);
    });
  }

  $('#btn-export').addEventListener('click', async () => {
    const btn = $('#btn-export');
    const all = state.options.map((o) => ({ input: Object.assign({}, o, { name: displayName(o) }), result: results.get(o.id), color: o.color }));
    const valid = all.filter((e) => e.result && e.result.ok);
    const skipped = all.length - valid.length;
    if (!valid.length) { toast('Complete at least one lease option before exporting.', true); return; }
    btn.disabled = true;
    const label = btn.textContent;
    btn.textContent = 'Preparing...';
    try {
      const ExcelJS = await loadExcelJS();
      const wb = window.LeaseExcel.buildWorkbook(ExcelJS, {
        options: valid,
        settings: state.settings,
        charts: Charts.snapshots(),
        logo: window.CBRE_LOGOS && window.CBRE_LOGOS.green,
        generatedAt: new Date(),
      });
      const buf = await wb.xlsx.writeBuffer();
      download(new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }),
        window.LeaseExcel.fileName(new Date()));
      toast(`Excel file ready with ${valid.length} ${valid.length === 1 ? 'option' : 'options'}${skipped ? `. ${skipped} incomplete ${skipped === 1 ? 'option was' : 'options were'} left out` : ''}.`);
    } catch (err) {
      console.error(err);
      toast(`Export failed: ${err.message}`, true);
    } finally {
      btn.disabled = false;
      btn.textContent = label;
    }
  });

  // -------------------------------------------------------------- init

  bindSettings();
  bindChartControls();
  try {
    applyState();
  } catch (err) {
    console.error(err);
    state = defaultState();
    applyState();
    firstVisit = false;
    toast('Saved data could not be loaded, so the example options are shown.', true);
  }
  if (firstVisit) toast('Example options loaded. Replace them with your proposal terms.');

  // Exposed for automated checks and the browser console.
  window.LeaseApp = { get state() { return state; }, results: () => results, update };
})();
