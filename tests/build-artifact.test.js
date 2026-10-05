'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const { build } = require('../scripts/build-artifact.js');
const versions = require('../vendor/VERSIONS.json');

const page = build();

test('artifact page starts with its title and inline styles', () => {
  assert.match(page, /^<title>CBRE Lease Rent Calculator<\/title>\n<style>/);
  assert.ok(!/<!doctype|<html|<head>|<body>/i.test(page), 'the publish skeleton supplies these');
});

test('libraries load from jsDelivr at the vendored versions', () => {
  assert.ok(page.includes(`https://cdn.jsdelivr.net/npm/chart.js@${versions['chart.js']}/dist/chart.umd.min.js`));
  assert.ok(page.includes(`https://cdn.jsdelivr.net/npm/exceljs@${versions.exceljs}/dist/exceljs.min.js`));
  const firstInline = page.indexOf('<script>\n');
  assert.ok(page.lastIndexOf('<script src="https://') < firstInline, 'libraries load before the app scripts');
});

test('everything else is inline: no local file references', () => {
  assert.ok(!/(src|href)="(assets|vendor|js|css)\//.test(page));
  assert.equal((page.match(/src="data:image\/png;base64,/g) || []).length, 2, 'both logos');
  ['window.CBRE_LOGOS', 'root.LeaseCalc', 'root.LeaseFormat', 'root.LeaseCharts', 'root.LeaseExcel', "window.claude.use('downloads')"]
    .forEach((marker) => assert.ok(page.includes(marker), marker));
});

test('artifact page stays well under the 16 MB limit', () => {
  assert.ok(page.length < 1024 * 1024);
});
