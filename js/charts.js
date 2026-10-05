/*
 * Comparison charts built on Chart.js 4. Browser global: window.LeaseCharts.
 *
 * Colours are CBRE brand colours. The default order is the CBRE Charts &
 * Graphs palette with accent green and wheat skipped, which keeps every
 * adjacent pair distinguishable under protanopia and deuteranopia
 * (checked with an OKLab colour-difference validator, worst pair dE 16.5).
 */
(function (root) {
  'use strict';

  const F = root.LeaseFormat;

  const PRESETS = {
    cbre: { label: 'CBRE charts', colors: ['#80BBAD', '#435254', '#D2785A', '#885073', '#A388BF', '#1F3765', '#3E7CA6'] },
    contrast: { label: 'CBRE high contrast', colors: ['#003F2D', '#D2785A', '#435254', '#A388BF', '#1F3765', '#3E7CA6'] },
    greens: { label: 'CBRE greens', colors: ['#003F2D', '#17E88F', '#538184', '#C0D4CB', '#012A2D', '#80BBAD'] },
  };

  // Swatches offered in the colour picker: brand primary, secondary and data colours.
  const SWATCHES = [
    '#003F2D', '#17E88F', '#012A2D', '#435254', '#80BBAD', '#538184',
    '#DBD99A', '#D2785A', '#885073', '#A388BF', '#1F3765', '#3E7CA6',
    '#032842', '#778F9C', '#96B3B6', '#C0D4CB', '#7F8480', '#CAD1D3',
  ];

  const COMPONENTS = [
    { key: 'netBaseRent', label: 'Net base rent' },
    { key: 'opex', label: 'OpEx' },
    { key: 'parking', label: 'Parking' },
  ];
  const COMPONENT_COLORS = ['#80BBAD', '#435254', '#D2785A'];

  const METRICS = {
    monthlyTotal: { label: 'Total monthly cost', kind: 'monthly', value: (m) => m.total, money: true },
    monthlyBase: { label: 'Net base rent per month', kind: 'monthly', value: (m) => m.netBaseRent, money: true },
    rate: { label: 'Base rent rate', kind: 'monthly', value: (m) => m.baseRatePsfYr, money: true, rate: true },
    cumulative: { label: 'Cumulative lease cost', kind: 'monthly', value: (m) => m.cumulative, money: true },
    annual: { label: 'Annual cost by lease year', kind: 'annual', value: (y) => y.total, money: true },
    totals: { label: 'Total lease cost by option', kind: 'totals', value: (r) => r.summary.totalCost, money: true },
    effective: { label: 'Effective rent by option', kind: 'totals', value: (r) => r.summary.effectiveRentPsfYr, money: true, rate: true },
    breakdown: { label: 'Lease cost breakdown by option', kind: 'breakdown', money: true },
  };

  const TYPES = {
    line: 'Line',
    stepped: 'Stepped line',
    area: 'Area',
    bar: 'Column',
    hbar: 'Horizontal bar',
  };

  function typesFor(metricKey) {
    const kind = (METRICS[metricKey] || METRICS.monthlyTotal).kind;
    if (kind === 'monthly') return ['line', 'stepped', 'area', 'bar'];
    if (kind === 'annual') return ['bar', 'hbar', 'line', 'stepped', 'area'];
    return ['bar', 'hbar'];
  }

  const INK = '#435254';
  const MUTED = '#767676';
  const GRID = '#E3E8E8';
  const FONT = '"Calibre", Arial, Helvetica, sans-serif';
  const HEIGHTS = { compact: 280, standard: 380, tall: 520 };

  function alpha(hex, a) {
    const h = hex.replace('#', '');
    const n = parseInt(h.length === 3 ? h.split('').map((c) => c + c).join('') : h, 16);
    return `rgba(${(n >> 16) & 255}, ${(n >> 8) & 255}, ${n & 255}, ${a})`;
  }

  // ------------------------------------------------------------ plugins

  // Paints a white canvas so PNG downloads and Excel images are not transparent.
  const background = {
    id: 'cbreBackground',
    beforeDraw(chart) {
      const { ctx } = chart;
      ctx.save();
      ctx.globalCompositeOperation = 'destination-over';
      ctx.fillStyle = '#FFFFFF';
      ctx.fillRect(0, 0, chart.width, chart.height);
      ctx.restore();
    },
  };

  // Vertical hairline that follows the hovered x position on line charts.
  const crosshair = {
    id: 'cbreCrosshair',
    afterDatasetsDraw(chart) {
      const active = chart.tooltip && chart.tooltip.getActiveElements();
      if (!active || !active.length || chart.config.type !== 'line') return;
      const x = active[0].element.x;
      const { top, bottom } = chart.chartArea;
      const { ctx } = chart;
      ctx.save();
      ctx.strokeStyle = 'rgba(67, 82, 84, 0.45)';
      ctx.lineWidth = 1;
      ctx.beginPath();
      ctx.moveTo(x, top);
      ctx.lineTo(x, bottom);
      ctx.stroke();
      ctx.restore();
    },
  };

  // Selective direct labels: the end of each line, every bar when there are
  // few of them, and stack totals on breakdown charts. Text uses ink colours,
  // never the series colour.
  const directLabels = {
    id: 'cbreLabels',
    afterDatasetsDraw(chart, _args, opts) {
      if (!opts || !opts.enabled) return;
      const { ctx } = chart;
      const horizontal = chart.options.indexAxis === 'y';
      const fmt = opts.format;
      ctx.save();
      ctx.font = `600 11px ${FONT}`;
      ctx.fillStyle = INK;
      ctx.lineWidth = 3;
      ctx.strokeStyle = '#FFFFFF';

      // Each label is a candidate box; a label that would overlap one already
      // drawn is skipped (the tooltip and the tables still carry the value).
      const LINE_H = 13;
      const candidate = (text, x, y, align, baseline) => {
        const w = ctx.measureText(text).width;
        const x1 = align === 'center' ? x - w / 2 : align === 'left' ? x : x - w;
        const y1 = baseline === 'bottom' ? y - LINE_H : baseline === 'middle' ? y - LINE_H / 2 : y;
        return { text, x, y, align, baseline, box: [x1 - 2, y1 - 1, x1 + w + 2, y1 + LINE_H + 1] };
      };
      const overlaps = (a, b) => a[0] < b[2] && b[0] < a[2] && a[1] < b[3] && b[1] < a[3];
      const placed = [];
      const draw = (c) => {
        if (placed.some((p) => overlaps(p, c.box))) return;
        placed.push(c.box);
        ctx.textAlign = c.align;
        ctx.textBaseline = c.baseline;
        ctx.strokeText(c.text, c.x, c.y);
        ctx.fillText(c.text, c.x, c.y);
      };

      if (opts.stackTotals) {
        const metas = chart.data.datasets.map((_, i) => chart.getDatasetMeta(i)).filter((m) => !m.hidden);
        const n = chart.data.labels.length;
        for (let i = 0; i < n; i++) {
          const total = chart.data.datasets.reduce((acc, ds, k) => acc + (chart.getDatasetMeta(k).hidden ? 0 : ds.data[i] || 0), 0);
          const tip = metas.reduce((best, m) => {
            const el = m.data[i];
            if (!el) return best;
            if (!best) return el;
            return horizontal ? (el.x > best.x ? el : best) : (el.y < best.y ? el : best);
          }, null);
          if (!tip) continue;
          if (horizontal) draw(candidate(fmt(total), tip.x + 6, tip.y, 'left', 'middle'));
          else draw(candidate(fmt(total), tip.x, tip.y - 6, 'center', 'bottom'));
        }
        ctx.restore();
        return;
      }

      // Bars get a label each when they all fit without touching; otherwise,
      // like lines, only the last value of each series is labelled.
      const build = (endsOnly) => {
        const out = [];
        chart.data.datasets.forEach((ds, di) => {
          const meta = chart.getDatasetMeta(di);
          if (meta.hidden) return;
          const isLine = meta.type === 'line';
          const points = meta.data;
          const valued = points.map((el, i) => (ds.data[i] === null || ds.data[i] === undefined ? -1 : i)).filter((i) => i >= 0);
          if (!valued.length) return;
          const indices = isLine || endsOnly || valued.length > opts.maxBars ? [valued[valued.length - 1]] : valued;
          indices.forEach((i) => {
            const el = points[i];
            const v = ds.data[i];
            if (horizontal) out.push(candidate(fmt(v), el.x + 6, el.y, 'left', 'middle'));
            else if (isLine) out.push(candidate(fmt(v), el.x + 6, el.y - 8, 'left', 'bottom'));
            else out.push(candidate(fmt(v), el.x, el.y - 6, 'center', 'bottom'));
          });
        });
        return out;
      };
      let labels = build(false);
      const crowded = labels.some((a, i) => labels.some((b, j) => j > i && overlaps(a.box, b.box)));
      if (crowded) labels = build(true);
      labels.forEach(draw);
      ctx.restore();
    },
  };

  // --------------------------------------------------------- data shaping

  function calendarKey(iso) {
    return iso.slice(0, 7);
  }

  function monthKeysBetween(startKey, endKey) {
    const keys = [];
    let [y, m] = startKey.split('-').map(Number);
    const [ey, em] = endKey.split('-').map(Number);
    while (y < ey || (y === ey && m <= em)) {
      keys.push(`${y}-${String(m).padStart(2, '0')}`);
      m++;
      if (m > 12) { m = 1; y++; }
    }
    return keys;
  }

  // Builds { labels, series: [{ name, color, data }] } for a monthly metric.
  function monthlyData(series, metric, align) {
    if (align === 'calendar') {
      const starts = series.map((s) => calendarKey(s.result.months[0].start)).sort();
      const ends = series.map((s) => calendarKey(s.result.months[s.result.months.length - 1].start)).sort();
      const keys = monthKeysBetween(starts[0], ends[ends.length - 1]);
      const index = new Map(keys.map((k, i) => [k, i]));
      return {
        labels: keys.map((k) => F.monthYear(k + '-01')),
        keys,
        series: series.map((s) => {
          const data = new Array(keys.length).fill(null);
          s.result.months.forEach((m) => { data[index.get(calendarKey(m.start))] = metric.value(m); });
          return { name: s.name, color: s.color, data };
        }),
      };
    }
    const n = Math.max(...series.map((s) => s.result.months.length));
    return {
      labels: Array.from({ length: n }, (_, i) => `Month ${i + 1}`),
      series: series.map((s) => ({
        name: s.name,
        color: s.color,
        data: Array.from({ length: n }, (_, i) => (s.result.months[i] ? metric.value(s.result.months[i]) : null)),
      })),
    };
  }

  function annualData(series, metric) {
    const n = Math.max(...series.map((s) => s.result.years.length));
    return {
      labels: Array.from({ length: n }, (_, i) => `Year ${i + 1}`),
      series: series.map((s) => ({
        name: s.name,
        color: s.color,
        data: Array.from({ length: n }, (_, i) => (s.result.years[i] ? metric.value(s.result.years[i]) : null)),
        months: s.result.years.map((y) => y.months),
      })),
    };
  }

  // ------------------------------------------------------------ builders

  function seriesColor(ds, index) {
    if (ds.type === 'line') return ds.borderColor;
    return Array.isArray(ds.backgroundColor) ? ds.backgroundColor[index] || ds.backgroundColor[0] : ds.backgroundColor;
  }

  function baseOptions(cfg, ctx) {
    const horizontal = ctx.type === 'hbar';
    const valueAxis = horizontal ? 'x' : 'y';
    const catAxis = horizontal ? 'y' : 'x';
    const fmtTick = ctx.rate ? (v) => F.money(v, ctx.currency, 2) : (v) => F.moneyCompact(v, ctx.currency);
    const scales = {};
    scales[catAxis] = {
      grid: { display: false },
      border: { color: '#CAD1D3' },
      ticks: {
        color: MUTED,
        font: { family: FONT, size: 11 },
        autoSkip: true,
        autoSkipPadding: 14,
        maxRotation: 0,
        callback: function (value, index) {
          const label = this.getLabelForValue(value);
          return ctx.tickFilter ? (ctx.tickFilter(index, label) ? label : null) : label;
        },
      },
    };
    scales[valueAxis] = {
      beginAtZero: !!cfg.beginAtZero,
      grid: { display: cfg.gridlines !== false, color: GRID, drawTicks: false },
      border: { display: false },
      ticks: { color: MUTED, font: { family: FONT, size: 11 }, padding: 6, callback: fmtTick, maxTicksLimit: 7 },
      stacked: !!ctx.stacked,
    };
    if (ctx.stacked) scales[catAxis].stacked = true;
    if (ctx.range) {
      if (Number.isFinite(ctx.range.max)) scales[valueAxis].suggestedMax = ctx.range.max;
      if (Number.isFinite(ctx.range.min) && !cfg.beginAtZero) scales[valueAxis].suggestedMin = ctx.range.min;
    }

    const legendPos = cfg.legend || 'bottom';
    return {
      responsive: true,
      maintainAspectRatio: false,
      animation: false,
      devicePixelRatio: Math.max(2, root.devicePixelRatio || 1),
      indexAxis: horizontal ? 'y' : 'x',
      interaction: { mode: ctx.pointHover ? 'nearest' : 'index', intersect: !!ctx.pointHover, axis: catAxis },
      layout: { padding: { top: cfg.dataLabels ? 22 : 8, right: cfg.dataLabels ? 56 : 12, left: 4, bottom: 4 } },
      scales,
      plugins: {
        legend: {
          display: legendPos !== 'none' && ctx.showLegend,
          position: legendPos === 'none' ? 'bottom' : legendPos,
          align: 'start',
          labels: {
            color: INK,
            font: { family: FONT, size: 12 },
            usePointStyle: true,
            pointStyle: ctx.lineLegend ? 'line' : 'rectRounded',
            boxWidth: 18,
            boxHeight: 10,
            pointStyleWidth: ctx.lineLegend ? 22 : 12,
            padding: 14,
            generateLabels(chart) {
              const items = root.Chart.defaults.plugins.legend.labels.generateLabels(chart);
              items.forEach((it) => {
                const c = seriesColor(chart.data.datasets[it.datasetIndex], 0);
                it.fillStyle = c;
                it.strokeStyle = c;
                it.lineWidth = ctx.lineLegend ? 3 : 0;
              });
              return items;
            },
          },
        },
        title: {
          display: !!ctx.title,
          text: ctx.title,
          align: 'start',
          color: INK,
          font: { family: FONT, size: 15, weight: '600' },
          padding: { top: 2, bottom: ctx.subtitle ? 2 : 12 },
        },
        subtitle: {
          display: !!ctx.subtitle,
          text: ctx.subtitle,
          align: 'start',
          color: MUTED,
          font: { family: FONT, size: 12 },
          padding: { bottom: 12 },
        },
        tooltip: {
          backgroundColor: '#012A2D',
          titleColor: '#FFFFFF',
          bodyColor: '#FFFFFF',
          titleFont: { family: FONT, size: 12, weight: '600' },
          bodyFont: { family: FONT, size: 12 },
          padding: 10,
          cornerRadius: 4,
          boxPadding: 4,
          usePointStyle: true,
          callbacks: {
            label(item) {
              const v = horizontal ? item.parsed.x : item.parsed.y;
              if (v === null || v === undefined) return null;
              const name = item.dataset.label;
              return ` ${F.money(v, ctx.currency, ctx.rate ? 2 : 0)}${ctx.rate ? ctx.rateSuffix : ''}  ${name}`;
            },
            labelPointStyle() {
              return { pointStyle: ctx.lineLegend ? 'line' : 'rectRounded', rotation: 0 };
            },
            labelColor(item) {
              const c = seriesColor(item.dataset, item.dataIndex);
              return { borderColor: c, backgroundColor: c, borderWidth: 3 };
            },
            footer: ctx.tooltipFooter,
          },
        },
        cbreLabels: {
          enabled: !!cfg.dataLabels,
          maxBars: 24,
          stackTotals: !!ctx.stacked,
          format: ctx.rate ? (v) => F.money(v, ctx.currency, 2) : (v) => F.moneyCompact(v, ctx.currency),
        },
      },
    };
  }

  function dataset(s, type, cfg) {
    const width = Number(cfg.lineWidth) || 2;
    if (type === 'bar' || type === 'hbar') {
      return {
        type: 'bar',
        label: s.name,
        data: s.data,
        backgroundColor: s.colors || s.color,
        hoverBackgroundColor: s.colors ? s.colors.map((c) => alpha(c, 0.8)) : alpha(s.color, 0.8),
        borderRadius: 4,
        borderSkipped: 'start',
        maxBarThickness: 28,
        categoryPercentage: 0.8,
        barPercentage: 0.92,
      };
    }
    return {
      type: 'line',
      label: s.name,
      data: s.data,
      borderColor: s.color,
      backgroundColor: type === 'area' ? alpha(s.color, 0.12) : s.color,
      fill: type === 'area' ? 'origin' : false,
      stepped: type === 'stepped' ? true : false,
      tension: 0,
      borderWidth: width,
      borderJoinStyle: 'round',
      borderCapStyle: 'round',
      pointRadius: cfg.markers ? 3.5 : 0,
      pointHoverRadius: 5,
      pointBackgroundColor: s.color,
      pointBorderColor: '#FFFFFF',
      pointBorderWidth: 2,
      pointHitRadius: 12,
      spanGaps: false,
    };
  }

  function lineTicks(labels, align, keys) {
    if (align === 'calendar' && keys) {
      const span = keys.length;
      const step = span > 60 ? 12 : span > 24 ? 6 : span > 9 ? 3 : 1;
      return (index) => {
        const m = Number(keys[index].slice(5, 7)) - 1;
        return m % step === 0;
      };
    }
    const span = labels.length;
    const step = span > 48 ? 12 : span > 18 ? 6 : span > 8 ? 3 : 1;
    return (index) => index % step === 0;
  }

  function seriesRange(seriesList) {
    let min = Infinity, max = -Infinity;
    seriesList.forEach((s) => s.data.forEach((v) => {
      if (v === null || v === undefined) return;
      if (v < min) min = v;
      if (v > max) max = v;
    }));
    return { min, max };
  }

  // Returns [{ title, config }] Chart.js configs for the current settings.
  function buildConfigs(series, cfg, settings) {
    const metricKey = METRICS[cfg.metric] ? cfg.metric : 'monthlyTotal';
    const metric = METRICS[metricKey];
    const types = typesFor(metricKey);
    const type = types.includes(cfg.type) ? cfg.type : types[0];
    const currency = settings.currency || '$';
    const unit = settings.areaUnit || 'SF';
    const rateSuffix = `/${unit}/yr`;
    const separate = cfg.layout === 'separate' && metric.kind !== 'totals' && series.length > 1;
    const titleText = (cfg.title || '').trim() || defaultTitle(metricKey, unit);
    const common = { currency, rate: !!metric.rate, rateSuffix, type };
    const out = [];

    if (metric.kind === 'monthly' || metric.kind === 'annual') {
      const shaped = metric.kind === 'monthly'
        ? monthlyData(series, metric, cfg.align)
        : annualData(series, metric);
      const lineLike = type === 'line' || type === 'stepped' || type === 'area';
      const ticks = metric.kind === 'monthly' ? lineTicks(shaped.labels, cfg.align, shaped.keys) : null;
      const range = cfg.sameScale !== false ? seriesRange(shaped.series) : null;
      // Partial lease years are noted per hovered series, using the months
      // carried on each dataset (correct in combined and separate layouts).
      const footer = metric.kind === 'annual'
        ? (items) => items.map((it) => {
          const months = it.dataset.months ? it.dataset.months[it.dataIndex] : 12;
          return months && months < 12 - 1e-9 ? `${it.dataset.label}: partial year, ${F.number(months, Number.isInteger(months) ? 0 : 1)} months` : null;
        }).filter(Boolean)
        : undefined;
      const groups = separate ? shaped.series.map((s) => [s]) : [shaped.series];
      groups.forEach((group) => {
        const ctx = Object.assign({}, common, {
          title: separate ? group[0].name : titleText,
          subtitle: separate ? titleText : subtitleText(cfg, metric, unit, currency),
          showLegend: !separate && group.length > 1,
          lineLegend: lineLike,
          tickFilter: ticks,
          range: separate ? range : null,
          tooltipFooter: footer,
        });
        const options = baseOptions(cfg, ctx);
        out.push({
          title: separate ? `${titleText}: ${group[0].name}` : titleText,
          config: {
            type: lineLike ? 'line' : 'bar',
            data: { labels: shaped.labels, datasets: group.map((s) => Object.assign(dataset(s, type, cfg), { months: s.months })) },
            options,
            plugins: [background, crosshair, directLabels],
          },
        });
      });
      return { charts: out, type, separate };
    }

    if (metric.kind === 'totals') {
      const ctx = Object.assign({}, common, {
        title: titleText,
        subtitle: subtitleText(cfg, metric, unit, currency),
        showLegend: false,
      });
      const data = series.map((s) => metric.value(s.result));
      out.push({
        title: titleText,
        config: {
          type: 'bar',
          data: {
            labels: series.map((s) => s.name),
            datasets: [dataset({ name: metric.label, data, colors: series.map((s) => s.color) }, type, cfg)],
          },
          options: baseOptions(cfg, ctx),
          plugins: [background, directLabels],
        },
      });
      return { charts: out, type, separate: false };
    }

    // Breakdown: stacked components per option, or one chart per option.
    const compColors = cfg.componentColors || COMPONENT_COLORS;
    if (separate) {
      const range = cfg.sameScale !== false
        ? { min: 0, max: Math.max(...series.map((s) => Math.max(...COMPONENTS.map((c) => s.result.summary[sumKey(c.key)])))) }
        : null;
      series.forEach((s) => {
        const ctx = Object.assign({}, common, { title: s.name, subtitle: titleText, showLegend: false, range });
        out.push({
          title: `${titleText}: ${s.name}`,
          config: {
            type: 'bar',
            data: {
              labels: COMPONENTS.map((c) => c.label),
              datasets: [dataset({ name: s.name, data: COMPONENTS.map((c) => s.result.summary[sumKey(c.key)]), colors: compColors }, type, cfg)],
            },
            options: baseOptions(cfg, ctx),
            plugins: [background, directLabels],
          },
        });
      });
      return { charts: out, type, separate: true };
    }
    const ctx = Object.assign({}, common, {
      title: titleText,
      subtitle: subtitleText(cfg, metric, unit, currency),
      showLegend: true,
      stacked: true,
    });
    const options = baseOptions(cfg, ctx);
    out.push({
      title: titleText,
      config: {
        type: 'bar',
        data: {
          labels: series.map((s) => s.name),
          datasets: COMPONENTS.map((c, i) => Object.assign(
            dataset({ name: c.label, data: series.map((s) => s.result.summary[sumKey(c.key)]), color: compColors[i] }, type, cfg),
            { borderRadius: 0, borderSkipped: false, borderColor: '#FFFFFF', borderWidth: horizontalGap(type) },
          )),
        },
        options,
        plugins: [background, directLabels],
      },
    });
    return { charts: out, type, separate: false };
  }

  // A 2px white gap between stacked segments, drawn as a surface-coloured edge.
  function horizontalGap(type) {
    return type === 'hbar' ? { left: 0, right: 2, top: 0, bottom: 0 } : { top: 2, bottom: 0, left: 0, right: 0 };
  }

  function sumKey(componentKey) {
    return { netBaseRent: 'netBaseRent', opex: 'netOpex', parking: 'netParking' }[componentKey];
  }

  function defaultTitle(metricKey, unit) {
    switch (metricKey) {
      case 'monthlyTotal': return 'Total monthly cost';
      case 'monthlyBase': return 'Net base rent per month';
      case 'rate': return `Base rent rate per ${unit} per year`;
      case 'cumulative': return 'Cumulative lease cost';
      case 'annual': return 'Annual cost by lease year';
      case 'totals': return 'Total lease cost by option';
      case 'effective': return `Effective rent per ${unit} per year`;
      case 'breakdown': return 'Lease cost breakdown by option';
      default: return '';
    }
  }

  function subtitleText(cfg, metric, unit, currency) {
    if (metric.kind === 'monthly') {
      return cfg.align === 'calendar' ? 'By calendar month' : 'By lease month, from each option\'s commencement';
    }
    if (metric.kind === 'annual') return 'Lease years run 12 months from commencement';
    if (metric.kind === 'breakdown') return 'Net of free rent';
    if (metric.rate) return `All-in cost, ${currency}/${unit}/yr, net of free rent`;
    return 'All-in cost over the full term, net of free rent';
  }

  // ------------------------------------------------------------- render

  let instances = [];

  function destroy() {
    instances.forEach((i) => i.chart.destroy());
    instances = [];
  }

  /*
   * Renders into `container`. series = [{ name, color, result }] (valid only).
   * Returns { type, separate, charts: [{ title, chart }] }.
   */
  function render(container, series, cfg, settings) {
    destroy();
    container.textContent = '';
    if (!series.length || !root.Chart) return { charts: [], type: cfg.type, separate: false };
    root.Chart.defaults.font.family = FONT;
    root.Chart.defaults.color = INK;

    const built = buildConfigs(series, cfg, settings);
    const height = HEIGHTS[cfg.height] || HEIGHTS.standard;
    container.classList.toggle('chart-grid', built.separate);
    built.charts.forEach((c) => {
      const box = document.createElement('figure');
      box.className = 'chart-box';
      box.style.height = `${built.separate ? Math.round(height * 0.8) : height}px`;
      const canvas = document.createElement('canvas');
      canvas.setAttribute('role', 'img');
      canvas.setAttribute('aria-label', `${c.title}. Values are listed in the comparison and schedule tables.`);
      box.appendChild(canvas);
      container.appendChild(box);
      const chart = new root.Chart(canvas, c.config);
      instances.push({ title: c.title, chart });
    });
    return { type: built.type, separate: built.separate, charts: instances.slice() };
  }

  function snapshots() {
    return instances.map((i) => ({
      title: i.title,
      dataUrl: i.chart.toBase64Image('image/png', 1),
      width: i.chart.width,
      height: i.chart.height,
    }));
  }

  root.LeaseCharts = {
    PRESETS,
    SWATCHES,
    COMPONENTS,
    COMPONENT_COLORS,
    METRICS,
    TYPES,
    typesFor,
    buildConfigs,
    render,
    destroy,
    snapshots,
    defaultTitle,
  };
})(typeof self !== 'undefined' ? self : this);
