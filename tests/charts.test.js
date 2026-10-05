'use strict';

// charts.js is a browser script; give it the globals it expects.
global.self = global;
global.LeaseFormat = require('../js/format.js');
require('../js/charts.js');

const test = require('node:test');
const assert = require('node:assert/strict');
const C = require('../js/calc.js');
const Charts = global.LeaseCharts;

const opt = (over) => Object.assign({}, C.DEFAULT_OPTION, {
  size: 10000, baseRate: 38, opex: 14.5, commencement: '2027-01-01', expiration: '2031-12-31',
  parkingSpaces: 20, parkingRate: 150, freeMonths: 6, freeAppliesTo: 'all',
}, over);
const series = (opts) => opts.map((o, i) => ({ name: `Option ${i + 1}`, color: '#80BBAD', result: C.calculate(o) }));
const cfg = (over) => Object.assign({ layout: 'combined', type: 'bar', align: 'leaseMonth', sameScale: true }, over);
const settings = { currency: '$', areaUnit: 'SF' };

test('breakdown stacks sum to the total lease cost when free rent covers OpEx and parking', () => {
  const s = series([opt(), opt({ freeAppliesTo: 'gross' })]);
  const { charts } = Charts.buildConfigs(s, cfg({ metric: 'breakdown' }), settings);
  const datasets = charts[0].config.data.datasets;
  s.forEach((x, i) => {
    const stack = datasets.reduce((a, d) => a + d.data[i], 0);
    assert.ok(Math.abs(stack - x.result.summary.totalCost) < 0.01);
  });
});

test('annual tooltip notes the partial year of the hovered option, in both layouts', () => {
  const s = series([opt(), opt({ expiration: '2029-06-30' })]);
  ['combined', 'separate'].forEach((layout) => {
    const { charts } = Charts.buildConfigs(s, cfg({ metric: 'annual', layout }), settings);
    const chart = layout === 'combined' ? charts[0] : charts[1];
    const ds = chart.config.data.datasets[layout === 'combined' ? 1 : 0];
    const footer = chart.config.options.plugins.tooltip.callbacks.footer;
    const lines = footer([{ dataset: ds, dataIndex: 2 }]);
    assert.deepEqual(lines, ['Option 2: partial year, 6 months'], layout);
    assert.deepEqual(footer([{ dataset: ds, dataIndex: 1 }]), [], `${layout}: full year has no note`);
  });
});

test('every metric and allowed chart type builds a config', () => {
  const s = series([opt(), opt({ commencement: '2027-03-15', expiration: '2030-03-14' })]);
  Object.keys(Charts.METRICS).forEach((metric) => {
    Charts.typesFor(metric).forEach((type) => {
      ['combined', 'separate'].forEach((layout) => {
        ['leaseMonth', 'calendar'].forEach((align) => {
          const built = Charts.buildConfigs(s, cfg({ metric, type, layout, align }), settings);
          assert.ok(built.charts.length >= 1, `${metric}/${type}/${layout}/${align}`);
        });
      });
    });
  });
});
