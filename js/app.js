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

  function localISO(d) {
    const pad = (x) => String(x).padStart(2, '0');
    return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
  }

  function firstOfNextMonth() {
    const now = new Date();
    return localISO(new Date(now.getFullYear(), now.getMonth() + 1, 1));
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
    if (!over || !over.expiration) syncExpiration(opt);
    return opt;
  }

  function defaultState() {
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
      chart: {
        metric: 'monthlyTotal', layout: 'combined', type: 'line', align: 'leaseMonth', legend: 'bottom',
        height: 'standard', dataLabels: false, gridlines: true, beginAtZero: true, markers: false,
        sameScale: true, lineWidth: 2, title: '', preset: 'cbre', componentColors: Charts.COMPONENT_COLORS.slice(),
      },
      schedule: { optionId: a.id, view: 'monthly' },
    };
  }

  // Accepts saved or imported data and fills any gaps from the defaults.
  function normalizeState(raw) {
    const base = defaultState();
    if (!raw || !Array.isArray(raw.options) || !raw.options.length) return null;
    const options = raw.options.slice(0, MAX_OPTIONS).map((o, i) => {
      const opt = Object.assign({}, Calc.DEFAULT_OPTION, o);
      opt.id = typeof o.id === 'string' && o.id ? o.id : uid();
      opt.name = String(opt.name || `Option ${i + 1}`).slice(0, 60);
      if (!/^#[0-9a-fA-F]{6}$/.test(opt.color || '')) opt.color = PALETTE[i % PALETTE.length];
      return opt;
    });
    const ids = new Set();
    options.forEach((o) => { if (ids.has(o.id)) o.id = uid(); ids.add(o.id); });
    const state = {
      settings: Object.assign({}, base.settings, raw.settings),
      options,
      activeId: options.some((o) => o.id === raw.activeId) ? raw.activeId : options[0].id,
      chart: Object.assign({}, base.chart, raw.chart),
      schedule: Object.assign({}, base.schedule, raw.schedule),
    };
    if (!Array.isArray(state.chart.componentColors) || state.chart.componentColors.length !== 3) {
      state.chart.componentColors = Charts.COMPONENT_COLORS.slice();
    }
    if (!options.some((o) => o.id === state.schedule.optionId)) state.schedule.optionId = state.activeId;
    return state;
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

  // ------------------------------------------------------ term syncing

  function syncExpiration(opt) {
    const months = Calc.termMonthsFromInput(opt.termValue, opt.termUnit);
    if (Calc.parseDate(opt.commencement) && months > 0) {
      opt.expiration = Calc.expirationFromTerm(opt.commencement, months);
    }
  }

  function syncTermFromExpiration(opt) {
    const t = Calc.termFromDates(opt.commencement, opt.expiration);
    if (!(t.months > 0)) return;
    if (t.fraction > 1e-9) {
      opt.termUnit = 'months';
      opt.termValue = Math.round(t.months * 100) / 100;
    } else if (opt.termUnit === 'years' && t.whole % 12 === 0) {
      opt.termValue = t.whole / 12;
    } else {
      opt.termUnit = 'months';
      opt.termValue = t.whole;
    }
  }

  function convertTermUnit(opt, from, to) {
    const v = Number(opt.termValue);
    if (!isFinite(v) || v <= 0 || from === to) return;
    opt.termValue = to === 'years' ? Math.round((v / 12) * 100) / 100 : Math.round(v * 12 * 100) / 100;
  }

  // ------------------------------------------------------------- form

  const form = $('#option-form');

  function fillForm(opt) {
    Array.from(form.elements).forEach((input) => {
      const name = input.name;
      if (!name || !(name in opt)) return;
      const v = opt[name];
      if (input.type === 'radio') input.checked = String(v) === input.value;
      else if (input.type === 'checkbox') input.checked = !!v;
      else input.value = displayValue(input, v, opt);
    });
    form.setAttribute('aria-labelledby', `tab-${opt.id}`);
    applyVisibility(opt);
    updateOpexNote(opt);
    $('#btn-remove').disabled = state.options.length <= 1;
    $('#btn-duplicate').disabled = state.options.length >= MAX_OPTIONS;
  }

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
      convertTermUnit(opt, opt.termUnit, value);
      opt.termUnit = value;
      form.elements.termValue.value = opt.termValue;
    } else {
      opt[name] = value;
    }

    if (name === 'commencement' || name === 'termValue') {
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

  function renderFieldErrors(result) {
    const errors = (result && result.errors) || {};
    $$('[data-error-for]', form).forEach((node) => {
      const key = node.dataset.errorFor;
      node.textContent = errors[key] || '';
      const input = form.elements[key];
      if (input && input.setAttribute) {
        if (errors[key]) input.setAttribute('aria-invalid', 'true');
        else input.removeAttribute('aria-invalid');
      }
    });
  }

  function stat(label, value, sub) {
    return el('div', { class: 'stat' },
      el('div', { class: 'stat__label', text: label }),
      el('div', { class: 'stat__value', text: value }),
      sub ? el('div', { class: 'stat__sub', text: sub }) : null);
  }

  function renderOptionSummary(opt, result) {
    const box = $('#option-summary');
    box.textContent = '';
    const cur = state.settings.currency, unit = state.settings.areaUnit;
    if (!result || !result.ok) {
      box.className = 'option-summary option-summary--error';
      box.textContent = 'Complete the highlighted fields to see results for this option.';
      $('#deposit-value').textContent = '-';
      return;
    }
    box.className = 'option-summary';
    const s = result.summary;
    box.append(
      stat('Total lease cost', F.money(s.totalCost, cur, 0), F.termLabel(s.termMonths)),
      stat('Average monthly cost', F.money(s.avgMonthlyCost, cur, 0), `${F.money(s.startingMonthlyBase, cur, 0)} base rent in month 1`),
      stat(`Effective rent per ${unit}/yr`, F.money(s.effectiveRentPsfYr, cur, 2), 'All-in, net of free rent'),
      stat('Free rent value', F.money(-s.totalFreeRent, cur, 0), F.freeRentLabel(opt, s.freeMonths)),
    );
    $('#deposit-value').textContent = opt.depositBasis === 'none' ? 'None' : F.money(s.securityDeposit, cur, 2);
  }

  // ------------------------------------------------------------- tabs

  function renderTabs() {
    const bar = $('#option-tabs');
    bar.textContent = '';
    state.options.forEach((o) => {
      const r = results.get(o.id);
      const selected = o.id === state.activeId;
      const tab = el('button', {
        type: 'button', class: 'tab', role: 'tab', id: `tab-${o.id}`,
        'aria-selected': selected ? 'true' : 'false', 'aria-controls': 'option-form',
        tabindex: selected ? '0' : '-1', 'data-id': o.id,
        onclick: () => setActive(o.id),
      },
      el('span', { class: 'tab__dot', style: { background: o.color }, 'aria-hidden': 'true' }),
      el('span', { class: 'tab__name', text: o.name || 'Untitled' }),
      r && !r.ok ? el('span', { class: 'tab__warn', text: '!', title: 'Incomplete inputs', 'aria-label': 'Incomplete inputs' }) : null);
      bar.appendChild(tab);
    });
    bar.appendChild(el('button', {
      type: 'button', class: 'tab tab--add', disabled: state.options.length >= MAX_OPTIONS,
      title: state.options.length >= MAX_OPTIONS ? `Up to ${MAX_OPTIONS} options` : 'Add a lease option',
      onclick: addOption,
    }, '+ Add option'));
  }

  $('#option-tabs').addEventListener('keydown', (e) => {
    if (e.key !== 'ArrowRight' && e.key !== 'ArrowLeft') return;
    const idx = state.options.findIndex((o) => o.id === state.activeId);
    const next = (idx + (e.key === 'ArrowRight' ? 1 : -1) + state.options.length) % state.options.length;
    setActive(state.options[next].id);
    const tab = $(`#tab-${state.options[next].id}`);
    if (tab) tab.focus();
    e.preventDefault();
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
      baseRate: '',
      opex: from.opex,
    });
    state.options.push(opt);
    setActive(opt.id);
    form.elements.baseRate.focus();
  }

  $('#btn-duplicate').addEventListener('click', () => {
    if (state.options.length >= MAX_OPTIONS) return;
    const from = activeOption();
    const copy = Object.assign({}, from, {
      id: uid(),
      name: `${from.name} (copy)`.slice(0, 60),
      color: nextColor(state.options),
    });
    state.options.splice(state.options.indexOf(from) + 1, 0, copy);
    setActive(copy.id);
    toast(`Duplicated ${from.name}.`);
  });

  $('#btn-remove').addEventListener('click', () => {
    if (state.options.length <= 1) return;
    const opt = activeOption();
    if (!window.confirm(`Remove ${opt.name}? This cannot be undone.`)) return;
    const idx = state.options.indexOf(opt);
    state.options.splice(idx, 1);
    setActive(state.options[Math.max(0, idx - 1)].id);
    toast(`Removed ${opt.name}.`);
  });

  // ---------------------------------------------------------- settings

  function renderUnitLabels() {
    const cur = state.settings.currency, unit = state.settings.areaUnit;
    $$('[data-cur]').forEach((n) => { n.textContent = cur; });
    $$('[data-area]').forEach((n) => { n.textContent = unit; });
    $$('[data-rate-unit]').forEach((n) => { n.textContent = `/${unit}/yr`; });
    const basis = form.elements.rateBasis;
    const current = activeOption().rateBasis;
    basis.textContent = '';
    [['psf_yr', `/${unit}/yr`], ['psf_mo', `/${unit}/mo`], ['monthly', '/month total'], ['annual', '/year total']]
      .forEach(([v, t]) => basis.appendChild(el('option', { value: v, text: t })));
    basis.value = current;
  }

  function bindSettings() {
    const cur = $('#set-currency'), area = $('#set-area'), disc = $('#set-discount');
    cur.value = state.settings.currency;
    area.value = state.settings.areaUnit;
    disc.value = state.settings.discountRate;
    cur.addEventListener('change', () => { state.settings.currency = cur.value; renderUnitLabels(); update(); });
    area.addEventListener('change', () => { state.settings.areaUnit = area.value; renderUnitLabels(); update(); });
    disc.addEventListener('input', () => {
      state.settings.discountRate = disc.value === '' ? 0 : Number(disc.value);
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
        e.opt.name || 'Untitled',
        lowest && lowest === e ? [el('br'), el('span', { class: 'badge' }, '✓ Lowest total cost')] : null,
        e.r && !e.r.ok ? [el('br'), el('span', { class: 'badge', style: { background: '#FBEEEE', color: '#A03530' } }, '! Incomplete inputs')] : null));
    });
    table.appendChild(el('thead', null, head));

    const body = el('tbody');
    const cols = entries.length + 1;
    const section = (title) => body.appendChild(el('tr', { class: 'row-section' }, el('th', { colspan: cols, scope: 'colgroup', text: title })));
    const row = (label, get, opts) => {
      const tr = el('tr', { class: opts && opts.strong ? 'row-strong' : null }, el('th', { scope: 'row', text: label }));
      entries.forEach((e) => {
        if (!e.r || !e.r.ok) { tr.appendChild(el('td', { class: 'cell-muted', text: '-' })); return; }
        const v = get(e.r.summary, e.opt, e);
        tr.appendChild(typeof v === 'object' && v !== null ? el('td', null, v) : el('td', { text: v }));
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
      ? `${F.number(o.parkingSpaces)} spaces at ${F.money(o.parkingRate, cur, 0)}/month` : 'None'));

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
    if (lowest) meta.textContent = `${lowest.opt.name} has the lowest total lease cost.`;
    else if (valid.length === 1 && entries.length === 1) meta.textContent = 'Add an option to compare proposals side by side.';
    else meta.textContent = '';
  }

  // ------------------------------------------------------------ charts

  const chartArea = $('#chart-area');

  function chartSeries() {
    return state.options
      .map((o) => ({ id: o.id, name: o.name || 'Untitled', color: o.color, result: results.get(o.id) }))
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

  function swatchRow(current, onPick) {
    const wrap = el('div', { class: 'swatches', role: 'group', 'aria-label': 'CBRE colours' });
    Charts.SWATCHES.forEach((hex) => {
      wrap.appendChild(el('button', {
        type: 'button', class: 'swatch', style: { background: hex }, title: hex,
        'aria-label': `Use ${hex}`, 'aria-pressed': hex.toUpperCase() === String(current).toUpperCase() ? 'true' : 'false',
        onclick: () => onPick(hex),
      }));
    });
    return wrap;
  }

  function colourRow(name, color, onPick) {
    const input = el('input', { type: 'color', value: color, 'aria-label': `Colour for ${name}` });
    input.addEventListener('input', () => onPick(input.value.toUpperCase(), true));
    input.addEventListener('change', () => onPick(input.value.toUpperCase()));
    return el('div', { class: 'colour-row' },
      input,
      el('span', { class: 'colour-row__name', text: name, title: name }),
      swatchRow(color, (hex) => onPick(hex)));
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
        }));
      });
      return;
    }
    state.options.forEach((o) => {
      list.appendChild(colourRow(o.name || 'Untitled', o.color, (hex, live) => {
        o.color = hex;
        state.chart.preset = 'custom';
        afterColourChange(live);
      }));
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
    if (invalid.length) notes.push(`${invalid.map((o) => o.name).join(', ')} ${invalid.length === 1 ? 'is' : 'are'} left out until the inputs are complete.`);
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

  function bindChartControls() {
    const c = state.chart;
    const on = (sel, evt, fn) => $(sel).addEventListener(evt, fn);
    on('#c-metric', 'change', (e) => {
      c.metric = e.target.value;
      const types = Charts.typesFor(c.metric);
      if (!types.includes(c.type)) c.type = types[0];
      update();
    });
    on('#c-type', 'change', (e) => { c.type = e.target.value; update(); });
    $$('input[name="c-layout"]').forEach((r) => r.addEventListener('change', () => { if (r.checked) { c.layout = r.value; update(); } }));
    on('#c-align', 'change', (e) => { c.align = e.target.value; update(); });
    on('#c-title', 'input', (e) => { c.title = e.target.value; scheduleUpdate(); });
    on('#c-legend', 'change', (e) => { c.legend = e.target.value; update(); });
    on('#c-height', 'change', (e) => { c.height = e.target.value; update(); });
    on('#c-linewidth', 'input', (e) => { c.lineWidth = Number(e.target.value); $('#c-linewidth-out').textContent = `${c.lineWidth}px`; scheduleUpdate(); });
    on('#c-labels', 'change', (e) => { c.dataLabels = e.target.checked; update(); });
    on('#c-grid', 'change', (e) => { c.gridlines = e.target.checked; update(); });
    on('#c-zero', 'change', (e) => { c.beginAtZero = e.target.checked; update(); });
    on('#c-markers', 'change', (e) => { c.markers = e.target.checked; update(); });
    on('#c-samescale', 'change', (e) => { c.sameScale = e.target.checked; update(); });
    on('#c-preset', 'change', (e) => {
      c.preset = e.target.value;
      const preset = Charts.PRESETS[c.preset];
      if (preset) {
        state.options.forEach((o, i) => { o.color = preset.colors[i % preset.colors.length]; });
        c.componentColors = preset.colors.slice(0, 3);
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
    state.options.forEach((o) => sel.appendChild(el('option', { value: o.id, text: o.name || 'Untitled' })));
    if (!state.options.some((o) => o.id === state.schedule.optionId)) state.schedule.optionId = state.activeId;
    sel.value = state.schedule.optionId;
    $$('input[name="s-view"]').forEach((r) => { r.checked = r.value === state.schedule.view; });

    const table = $('#schedule-table');
    table.textContent = '';
    const opt = state.options.find((o) => o.id === state.schedule.optionId);
    const r = results.get(opt.id);
    const cur = state.settings.currency, unit = state.settings.areaUnit;
    if (!r || !r.ok) {
      table.appendChild(el('tbody', null, el('tr', null, el('td', { class: 'cell-invalid', text: `Complete the inputs for ${opt.name} to see its schedule.` }))));
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
        body.appendChild(el('tr', null,
          el('td', { text: `Year ${y.leaseYear}` }),
          el('td', { text: F.number(y.months, Number.isInteger(Math.round(y.months * 100) / 100) ? 0 : 2) }),
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
    table.appendChild(el('tfoot', null, el('tr', null,
      el('td', { text: 'Total' }),
      el('td', { text: annual ? F.number(s.termMonths, Number.isInteger(s.termMonths) ? 0 : 2) : '' }),
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
  }

  let pending = false;
  function scheduleUpdate() {
    if (pending) return;
    pending = true;
    requestAnimationFrame(() => { pending = false; update(); });
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

  $('#file-open').addEventListener('change', (e) => {
    const file = e.target.files && e.target.files[0];
    e.target.value = '';
    if (!file) return;
    const reader = new FileReader();
    reader.onload = () => {
      let next = null;
      try { next = normalizeState(JSON.parse(reader.result)); } catch (err) { next = null; }
      if (!next) { toast('That file is not a lease scenario saved from this calculator.', true); return; }
      state = next;
      $('#set-currency').value = state.settings.currency;
      $('#set-area').value = state.settings.areaUnit;
      $('#set-discount').value = state.settings.discountRate;
      renderUnitLabels();
      fillForm(activeOption());
      update();
      toast(`Opened ${file.name}: ${state.options.length} ${state.options.length === 1 ? 'option' : 'options'}.`);
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
    const all = state.options.map((o) => ({ input: o, result: results.get(o.id), color: o.color }));
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
  renderUnitLabels();
  bindChartControls();
  fillForm(activeOption());
  update();
  if (firstVisit) toast('Example options loaded. Replace them with your proposal terms.');

  // Exposed for automated checks and the browser console.
  window.LeaseApp = { get state() { return state; }, results: () => results, update };
})();
