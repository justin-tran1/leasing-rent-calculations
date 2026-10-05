#!/usr/bin/env node
// Builds dist/cbre-lease-rent-calculator.html: the calculator as one
// self-contained page for publishing as a claude.ai artifact. CSS, app code
// and logos are inlined; Chart.js and ExcelJS load from jsDelivr, pinned to
// the versions in vendor/VERSIONS.json (the artifact sandbox only runs
// scripts from allow-listed CDNs). The page itself detects the claude.ai
// viewer and saves files through its downloads capability.
// Run: npm run build:artifact
'use strict';

const fs = require('fs');
const path = require('path');

const root = path.resolve(__dirname, '..');
const read = (p) => fs.readFileSync(path.join(root, p), 'utf8');

function build() {
  const versions = JSON.parse(read('vendor/VERSIONS.json'));
  const libs = [
    `https://cdn.jsdelivr.net/npm/chart.js@${versions['chart.js']}/dist/chart.umd.min.js`,
    `https://cdn.jsdelivr.net/npm/exceljs@${versions.exceljs}/dist/exceljs.min.js`,
  ];

  const html = read('index.html');
  const bodyStart = html.indexOf('<body>') + '<body>'.length;
  const scriptsStart = html.indexOf('<script src="vendor/');
  if (bodyStart < 6 || scriptsStart < bodyStart) throw new Error('index.html layout changed: cannot find <body> or the vendor scripts');
  let body = html.slice(bodyStart, scriptsStart).trim();

  // Logos become data URIs (the artifact serves a single file).
  body = body.replace(/src="assets\/([\w-]+\.png)"/g, (m, file) =>
    `src="data:image/png;base64,${fs.readFileSync(path.join(root, 'assets', file)).toString('base64')}"`);
  if (/src="(assets|vendor|js|css)\//.test(body)) throw new Error('index.html references a local file the build does not inline');

  // Same scripts, same order as index.html, minus the vendored libraries.
  const local = [...html.slice(scriptsStart).matchAll(/<script src="(js\/[\w.-]+\.js)"><\/script>/g)].map((m) => m[1]);
  const scripts = local.map((p) => {
    const src = read(p);
    if (/<\/script/i.test(src)) throw new Error(`${p} contains "</script", which would end the inline script early`);
    return `<script>\n${src}\n</script>`;
  });

  // The publish skeleton supplies doctype, head and body; title and style
  // come first so the title is found in the first 8 KB.
  return `<title>CBRE Lease Rent Calculator</title>
<style>
${read('css/styles.css')}
</style>
${body}

${libs.map((u) => `<script src="${u}"></script>`).join('\n')}
${scripts.join('\n')}
`;
}

if (require.main === module) {
  const out = path.join(root, 'dist', 'cbre-lease-rent-calculator.html');
  fs.mkdirSync(path.dirname(out), { recursive: true });
  const page = build();
  fs.writeFileSync(out, page);
  console.log(`${path.relative(root, out)} (${(page.length / 1024).toFixed(1)} KB)`);
}

module.exports = { build };
