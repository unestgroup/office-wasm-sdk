// Lists the files the viewer origin serves while documents are opened (what a packaged site must contain).
//   node tools/trace-requests.cjs [--samples a.docx,b.xlsx] [--modes embedded,edit] [--viewer-url http://host:port/] [--out file.json] [--baseline trace.json]
// Starts demo/server.cjs (which starts the viewer from ../office-wasm-viewer) unless --viewer-url points at an already
// running viewer site (for example a packed one, to check that nothing is missing). Prints the union of URL paths per
// sample/mode and writes them to --out. Exit code 1 when a request failed (4xx/5xx) or went outside the two origins.
// With --baseline (a trace.json of the unpacked source) only failures the baseline did not have count, and files the
// baseline fetched but this run did not are listed (a packed site that serves less than the source is fine, one that
// fails where the source works is not).
// SPDX-License-Identifier: Apache-2.0
'use strict';
const { spawn } = require('node:child_process');
const fs = require('node:fs');
const path = require('node:path');

const arg = (n, d) => { const i = process.argv.indexOf('--' + n); return i < 0 ? d : process.argv[i + 1]; };
const port = Number(arg('port', 8765));
const base = 'http://127.0.0.1:' + port;
const viewerUrl = arg('viewer-url', '');
const viewerBase = (viewerUrl || 'http://127.0.0.1:' + (port + 1) + '/').replace(/\/$/, '');
const modes = arg('modes', 'embedded,edit').split(',');
const samples = arg('samples', 'medium.docx,test1.xlsx,simple.pptx,emf-math.docx,word-fonts.docx,vn-fonts.docx').split(',');
const baselineFile = arg('baseline', '');
const outFile = arg('out', path.join(__dirname, '..', 'out', 'trace.json'));

function chromiumPath() {
  if (process.env.PW_CHROMIUM) return process.env.PW_CHROMIUM;
  const dir = path.join(process.env.LOCALAPPDATA || '', 'ms-playwright');
  if (!fs.existsSync(dir)) return undefined;
  const revs = fs.readdirSync(dir).filter((d) => /^chromium-\d+$/.test(d)).sort((a, b) => Number(b.split('-')[1]) - Number(a.split('-')[1]));
  for (const rev of revs) for (const sub of ['chrome-win64', 'chrome-win']) {
    const exe = path.join(dir, rev, sub, 'chrome.exe');
    if (fs.existsSync(exe)) return exe;
  }
  return undefined;
}

(async () => {
  const { chromium } = require('playwright');
  const args = [path.join(__dirname, '..', 'demo', 'server.cjs'), String(port)];
  if (viewerUrl) args.push('--viewer-url', viewerUrl);
  const server = spawn(process.execPath, args, { stdio: ['ignore', 'ignore', 'inherit'] });
  await new Promise((r) => setTimeout(r, 1500));
  const browser = await chromium.launch({ executablePath: chromiumPath(), headless: true });
  const seen = new Map(); // url path -> {bytes, samples:Set}
  const bad = [];
  const external = [];
  try {
    for (const sample of samples) for (const mode of modes) {
      const ctx = await browser.newContext({ viewport: { width: 1400, height: 900 } });
      await ctx.route('**/*', (route) => {
        const u = route.request().url();
        if (u.startsWith(base + '/') || u.startsWith(viewerBase + '/') || u.startsWith('blob:') || u.startsWith('data:')) return route.continue();
        external.push(u.slice(0, 160));
        return route.abort();
      });
      const page = await ctx.newPage();
      page.on('response', (r) => {
        const u = r.url();
        if (!u.startsWith(viewerBase + '/')) return;
        const p = new URL(u).pathname;
        if (r.status() >= 400) bad.push(r.status() + ' ' + p);
        const e = seen.get(p) || { samples: new Set() };
        e.samples.add(sample + ':' + mode);
        seen.set(p, e);
      });
      await page.goto(`${base}/?sample=${encodeURIComponent(sample)}&mode=${mode}`);
      try {
        await page.waitForFunction(() => window.__state && (window.__state.phase === 'ready' || window.__state.phase === 'error'), null, { timeout: 90000 });
        await page.waitForTimeout(2000);
        if (await page.evaluate(() => window.__state.phase) === 'ready') await page.evaluate(() => window.__poc.exportPdf());
      } catch (e) { bad.push('timeout ' + sample + ':' + mode); }
      await ctx.close();
      console.log(`${sample} [${mode}] done, ${seen.size} files so far`);
    }
  } finally {
    await browser.close();
    server.kill();
  }
  const list = [...seen.keys()].sort();
  fs.mkdirSync(path.dirname(outFile), { recursive: true });
  fs.writeFileSync(outFile, JSON.stringify({ viewer: viewerBase, files: list, bad, external }, null, 1));
  let fresh = bad;
  if (baselineFile) {
    const base = JSON.parse(fs.readFileSync(baselineFile, 'utf8'));
    const known = new Set(base.bad);
    fresh = [...new Set(bad)].filter((b) => !known.has(b));
    const unused = base.files.filter((f) => !seen.has(f));
    console.log(`baseline: ${known.size} known failures, ${fresh.length} new, ${unused.length} baseline files not fetched here`);
    if (unused.length) console.log(unused.slice(0, 20).join('\n'));
  }
  console.log(`${list.length} files, ${bad.length} bad, ${external.length} external -> ${outFile}`);
  if (fresh.length) console.log([...new Set(fresh)].slice(0, 20).join('\n'));
  process.exit(fresh.length || external.length ? 1 : 0);
})().catch((e) => { console.error(e); process.exit(2); });
