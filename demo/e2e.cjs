// End-to-end check of the offline preview PoC in headless Chromium.
//   node demo/e2e.cjs [--modes embedded,desktop,edit] [--samples medium.docx,test1.xlsx] [--port 8765] [--headed]
//   --local-fonts   Windows only: the page gets a queryLocalFonts() that serves some fonts of C:\Windows\Fonts (the real API needs a
//                   permission prompt), the demo opens with "fonts of this computer" and the run checks they reach editor and PDF
//                            [--coi] [--x2t prebuilt|st|mt] [--threads N] [--report report.json]
//   --coi  serve with COOP + COEP (cross-origin isolated page -> the multi-threaded x2t build is used when present)
//   --x2t  force one x2t build instead of the automatic choice
//   --threads N  thread pool size of the multi-threaded build
//   --math-fonts embedded|installed-first|off   substitute fonts for MathType equations (viewer default: embedded)
//   --tag  suffix for output file names
// For every sample x mode: open -> wait for the document -> screenshot -> export PDF -> extract text (JS and x2t);
// in edit mode also: type a marker, save, check that the saved Office file contains it.
// Outputs: out/e2e/report.json, *.png, *.pdf. Exit code 1 when any run fails.
'use strict';
const { spawn } = require('node:child_process');
const fs = require('node:fs');
const path = require('node:path');

const root = path.join(__dirname, '..');
const outDir = path.join(root, 'out', 'e2e');
const arg = (name, dflt) => {
  const i = process.argv.indexOf('--' + name);
  return i < 0 ? dflt : process.argv[i + 1];
};
const port = Number(arg('port', 8765));
const base = 'http://127.0.0.1:' + port;
const viewerBase = 'http://127.0.0.1:' + (port + 1); // the viewer runs on its own origin (server.cjs starts it)
const modes = arg('modes', 'embedded,desktop').split(',');
const samples = arg('samples', 'medium.docx,fonts_and_images.docx,test1.xlsx,simple.pptx').split(',');
const headed = process.argv.includes('--headed');
const coi = process.argv.includes('--coi');
const forceX2t = arg('x2t', '');
const threads = arg('threads', ''); // thread pool size of the multi-threaded build (default: chosen by the client)
const reportName = arg('report', 'report.json');
const mathFonts = arg('math-fonts', '');
const tagSuffix = arg('tag', '');
const localFonts = process.argv.includes('--local-fonts');
// [family, style, file in C:\Windows\Fonts]: the fonts the stub of queryLocalFonts offers
const WINDOWS_FONTS = [
  ['Arial', 'Regular', 'arial.ttf'], ['Arial', 'Bold', 'arialbd.ttf'], ['Times New Roman', 'Regular', 'times.ttf'], ['Calibri', 'Regular', 'calibri.ttf'],
  ['Segoe UI', 'Regular', 'segoeui.ttf'], ['Segoe UI', 'Bold', 'segoeuib.ttf'], ['Georgia', 'Regular', 'georgia.ttf'], ['Century Gothic', 'Regular', 'GOTHIC.TTF'],
  ['Tahoma', 'Regular', 'tahoma.ttf'], ['Verdana', 'Regular', 'verdana.ttf'], ['Consolas', 'Regular', 'consola.ttf'], ['Comic Sans MS', 'Regular', 'comic.ttf'],
  ['Trebuchet MS', 'Regular', 'trebuc.ttf'], ['Garamond', 'Regular', 'GARA.TTF'], ['Palatino Linotype', 'Regular', 'pala.ttf'],
].filter(([, , f]) => fs.existsSync(path.join(process.env.WINDIR || 'C:\Windows', 'Fonts', f)));
const READY_TIMEOUT = Number(arg('timeout', 90000));

function findChromium() {
  if (process.env.PW_CHROMIUM) return process.env.PW_CHROMIUM;
  const dir = path.join(process.env.LOCALAPPDATA || '', 'ms-playwright');
  if (!fs.existsSync(dir)) return null;
  const revs = fs.readdirSync(dir).filter((d) => /^chromium-\d+$/.test(d)).sort((a, b) => Number(b.split('-')[1]) - Number(a.split('-')[1]));
  for (const rev of revs) {
    for (const sub of ['chrome-win64', 'chrome-win']) {
      const exe = path.join(dir, rev, sub, 'chrome.exe');
      if (fs.existsSync(exe)) return exe;
    }
  }
  return null;
}

async function launch(chromium) {
  try {
    return await chromium.launch({ headless: !headed });
  } catch (e) {
    const exe = findChromium();
    if (!exe) throw e;
    console.log('playwright browser not installed for this version; using ' + exe);
    return chromium.launch({ headless: !headed, executablePath: exe });
  }
}

async function waitForServer() {
  for (let i = 0; i < 50; i++) {
    try {
      const [a, b] = await Promise.all([fetch(base + '/server-stats.json'), fetch(viewerBase + '/server-stats.json')]);
      if (a.ok && b.ok) return;
    } catch { /* not up yet */ }
    await new Promise((r) => setTimeout(r, 200));
  }
  throw new Error('preview server did not start on ' + base);
}

// requests and bytes served by both origins (host page + viewer)
const stats = async () => {
  const [a, b] = await Promise.all([base, viewerBase].map(async (u) => (await fetch(u + '/server-stats.json')).json()));
  return { requests: a.requests + b.requests, bytes: a.bytes + b.bytes };
};

// Text of the named parts of a zip (central directory + raw inflate); enough to look for a typed marker in an OOXML file.
function zipText(buf, pick) {
  const zlib = require('node:zlib');
  let eocd = buf.length - 22;
  while (eocd >= 0 && buf.readUInt32LE(eocd) !== 0x06054b50) eocd--;
  if (eocd < 0) throw new Error('not a zip');
  const count = buf.readUInt16LE(eocd + 10);
  let p = buf.readUInt32LE(eocd + 16);
  const out = {};
  for (let i = 0; i < count; i++) {
    const method = buf.readUInt16LE(p + 10);
    const csize = buf.readUInt32LE(p + 20);
    const nlen = buf.readUInt16LE(p + 28);
    const elen = buf.readUInt16LE(p + 30);
    const clen = buf.readUInt16LE(p + 32);
    const local = buf.readUInt32LE(p + 42);
    const name = buf.toString('utf8', p + 46, p + 46 + nlen);
    if (pick.test(name)) {
      const start = local + 30 + buf.readUInt16LE(local + 26) + buf.readUInt16LE(local + 28);
      const raw = buf.subarray(start, start + csize);
      out[name] = (method === 0 ? raw : zlib.inflateRawSync(raw)).toString('utf8');
    }
    p += 46 + nlen + elen + clen;
  }
  return out;
}

// Edit mode: type a marker into the document, save it through the viewer, and look for the marker in the saved file.
async function editCheck(page, sample, tag) {
  const marker = 'EDITOK' + Date.now() % 100000;
  const box = await page.locator('#stage').boundingBox();
  const ext = sample.split('.').pop();
  const cx = box.x + box.width / 2;
  if (ext === 'pptx') {
    await page.mouse.click(cx, box.y + 520);
    await page.mouse.dblclick(cx, box.y + 520);
  } else if (ext === 'xlsx') {
    await page.mouse.click(box.x + 200, box.y + 300);
  } else {
    await page.mouse.click(cx, box.y + 300);
  }
  await page.keyboard.type(marker, { delay: 15 });
  if (ext === 'xlsx') await page.keyboard.press('Enter');
  await page.waitForTimeout(1500);
  const saved = await page.evaluate(async () => {
    try {
      const r = await window.__poc.saveFile();
      let s = '';
      for (let i = 0; i < r.bytes.length; i += 0x8000) s += String.fromCharCode.apply(null, r.bytes.subarray(i, i + 0x8000));
      return { ok: true, name: r.name, b64: btoa(s) };
    } catch (e) {
      return { ok: false, error: String(e && e.message) };
    }
  });
  if (!saved.ok) return saved;
  const buf = Buffer.from(saved.b64, 'base64');
  fs.writeFileSync(path.join(outDir, tag + '.saved.' + ext), buf);
  const parts = zipText(buf, /^(word\/document\.xml|xl\/sharedStrings\.xml|ppt\/slides\/slide\d+\.xml)$/);
  const found = Object.entries(parts).filter(([, t]) => t.includes(marker)).map(([n]) => n);
  return { ok: found.length > 0, name: saved.name, bytes: buf.length, marker, found };
}

function pdfInfo(buf) {
  const s = buf.toString('latin1');
  const pages = (s.match(/\/Type\s*\/Page(?![a-zA-Z])/g) || []).length;
  const names = [...new Set([...s.matchAll(/\/FontName\s*\/(?:[A-Z]{6}\+)?([\w-]+)/g)].map((m) => m[1]))];
  return { fontNames: names, header: s.slice(0, 8).trim(), pages: pages || null, images: (s.match(/\/Subtype\s*\/Image/g) || []).length, fonts: (s.match(/\/FontFile\d?/g) || []).length };
}

async function runOne(browser, sample, mode) {
  const tag = sample + '.' + mode + (coi ? '.coi' : '') + (forceX2t ? '.' + forceX2t : '') + (mathFonts ? '.math-' + mathFonts : '') + (localFonts ? '.localfonts' : '') + tagSuffix;
  fs.mkdirSync(path.dirname(path.join(outDir, tag)), { recursive: true });
  const run = { sample, mode, ok: false, console: [], pageErrors: [], badResponses: [], external: [] };
  const before = await stats();
  const context = await browser.newContext({ viewport: { width: 1400, height: 900 } });
  // Offline guarantee: anything that is not our local server is blocked and recorded.
  await context.route('**/*', (route) => {
    const url = route.request().url();
    if (url.startsWith(base + '/') || url.startsWith(viewerBase + '/') || url.startsWith('blob:') || url.startsWith('data:')) return route.continue();
    run.external.push(url.slice(0, 200));
    return route.abort();
  });
  if (localFonts) {
    await context.route(base + '/__testfonts/*', (route) => {
      const f = decodeURIComponent(new URL(route.request().url()).pathname.split('/').pop());
      if (!WINDOWS_FONTS.some(([, , n]) => n === f)) return route.abort();
      return route.fulfill({ body: fs.readFileSync(path.join(process.env.WINDIR || 'C:\Windows', 'Fonts', f)), contentType: 'font/ttf' });
    });
    await context.addInitScript((list) => {
      window.queryLocalFonts = async () => list.map(([family, style, file]) => ({
        family, style, fullName: family + (style === 'Regular' ? '' : ' ' + style), postscriptName: family.replace(/ /g, '') + '-' + style,
        blob: async () => (await fetch('/__testfonts/' + encodeURIComponent(file))).blob(),
      }));
    }, WINDOWS_FONTS);
  }
  const page = await context.newPage();
  page.on('console', (m) => {
    if ((m.type() === 'error' || m.type() === 'warning') && run.console.length < 40) run.console.push(m.type() + ': ' + m.text().slice(0, 300));
  });
  page.on('pageerror', (e) => run.pageErrors.push(String(e.message).slice(0, 400)));
  page.on('response', (r) => {
    if (r.status() >= 400 && run.badResponses.length < 40) run.badResponses.push(r.status() + ' ' + r.url().replace(base, '').replace(viewerBase, 'viewer:'));
  });
  try {
    const t0 = Date.now();
    await page.goto(`${base}/?sample=${encodeURIComponent(sample)}&mode=${mode}` + (forceX2t ? '&x2t=' + forceX2t : '') + (threads ? '&threads=' + threads : '') + (mathFonts ? '&mathFonts=' + mathFonts : '') + (localFonts ? '&localfonts=1' : ''));
    try {
      await page.waitForFunction(() => window.__state && (window.__state.phase === 'ready' || window.__state.phase === 'error'), null, { timeout: READY_TIMEOUT });
    } catch {
      run.timeout = 'document not ready after ' + READY_TIMEOUT + ' ms';
    }
    run.readyWallMs = Date.now() - t0;
    await page.waitForTimeout(2500); // let fonts / images paint
    await page.screenshot({ path: path.join(outDir, tag + '.png') });
    run.frames = page.frames().map((f) => f.url().replace(base, '').replace(viewerBase, 'viewer:').slice(0, 120));
    run.phase = await page.evaluate(() => window.__state.phase);
    run.load = await stats();
    run.load = { requests: run.load.requests - before.requests, bytes: run.load.bytes - before.bytes };

    if (run.phase === 'ready') {
      // The editor runs in a frame of the cross-origin viewer frame: ask the viewer, and check that the frame exists.
      const vs = await page.evaluate(() => window.__poc.getState());
      if (localFonts) run.localFonts = vs.localFonts;
      run.editorApi = { iframe: page.frames().some((f) => f.name() === 'frameEditor'), ...vs.editor, document: vs.document };
      run.pdf = await page.evaluate(async () => {
        try {
          const p = await window.__poc.exportPdf();
          return p ? { ok: true, bytes: p.bytes.length, pdfBinBytes: p.pdfBinBytes, x2tMs: p.x2tMs, mathPictures: p.mathPictures } : { ok: false, error: 'no result' };
        } catch (e) {
          return { ok: false, error: String(e && e.message) };
        }
      });
      if (run.pdf.ok) {
        const buf = Buffer.from(await page.evaluate(() => window.__poc.pdfBase64()), 'base64');
        fs.writeFileSync(path.join(outDir, tag + '.pdf'), buf);
        Object.assign(run.pdf, pdfInfo(buf));
      }
      run.text = await page.evaluate(async () => {
        const out = {};
        try {
          const r = await window.__poc.extractJs();
          out.js = { chars: r.text.length, type: r.type, stats: r.stats };
        } catch (e) {
          out.js = { error: String(e && e.message) };
        }
        try {
          out.x2t = { chars: (await window.__poc.extractX2t()).length };
        } catch (e) {
          out.x2t = { error: String(e && e.message) };
        }
        return out;
      });
      if (mode === 'edit') run.edit = await editCheck(page, sample, tag).catch((e) => ({ ok: false, error: String(e && e.message) }));
      await page.screenshot({ path: path.join(outDir, tag + '.after.png') });
    }
    const st = await page.evaluate(() => ({ x2t: window.__state.x2t, timings: window.__state.timings, error: window.__state.error, events: window.__state.events }));
    Object.assign(run, st);
    run.ok = run.phase === 'ready' && !!(run.pdf && run.pdf.ok && run.pdf.header.startsWith('%PDF')) && !(run.text.js.error || run.text.x2t.error) && (mode !== 'edit' || !!(run.edit && run.edit.ok)) && (!localFonts || !!(run.localFonts && run.localFonts.installed.length));
  } catch (e) {
    run.exception = String(e && e.stack ? e.stack : e).slice(0, 1500);
  } finally {
    await context.close();
  }
  return run;
}

(async () => {
  fs.mkdirSync(outDir, { recursive: true });
  const { chromium } = require('playwright');
  const server = spawn(process.execPath, [path.join(__dirname, 'server.cjs'), String(port), ...(coi ? ['--coi'] : [])], { stdio: ['ignore', 'inherit', 'inherit'] });
  const report = { startedAt: new Date().toISOString(), node: process.version, coi, forceX2t: forceX2t || null, mathFonts: mathFonts || null, runs: [] };
  let browser;
  try {
    await waitForServer();
    browser = await launch(chromium);
    report.browser = browser.version();
    for (const mode of modes) {
      for (const sample of samples) {
        const run = await runOne(browser, sample, mode);
        report.runs.push(run);
        const t = run.timings || {};
        console.log(
          `${run.ok ? 'PASS' : 'FAIL'} ${sample} [${mode}] x2t=${run.x2t ? run.x2t.build + (run.x2t.threads ? '/' + run.x2t.threads + 't' : '') : '?'} phase=${run.phase} ready=${t.documentReadyMs ?? '-'}ms bin=${t.toEditorBinMs ?? '-'}ms `
          + `pdf=${run.pdf ? (run.pdf.ok ? run.pdf.bytes + 'B/' + run.pdf.pages + 'p/' + t.pdfTotalMs + 'ms' : 'ERR ' + run.pdf.error) : '-'} `
          + `load=${run.load ? run.load.requests + 'req/' + (run.load.bytes / 1048576).toFixed(1) + 'MB' : '-'} `
          + `errors: console=${run.console.length} page=${run.pageErrors.length} http=${run.badResponses.length} external=${run.external.length}`
          + (run.edit ? ` edit=${run.edit.ok ? 'saved ' + run.edit.bytes + 'B in ' + run.edit.found.join(',') : 'FAIL ' + (run.edit.error || 'marker not found')}` : '')
          + (run.localFonts ? ` localFonts=${run.localFonts.installed.join('/')} pdfFonts=${(run.pdf.fontNames || []).join('/')}` : '')
          + (run.timeout ? ' TIMEOUT' : '') + (run.exception ? ' EXCEPTION' : ''),
        );
      }
    }
  } catch (e) {
    report.fatal = String(e && e.stack ? e.stack : e);
    console.error(report.fatal);
  } finally {
    if (browser) await browser.close().catch(() => {});
    server.kill();
    fs.writeFileSync(path.join(outDir, reportName), JSON.stringify(report, null, 1));
  }
  const failed = report.fatal || report.runs.some((r) => !r.ok);
  console.log('report: ' + path.join(outDir, reportName));
  process.exit(failed ? 1 : 0);
})();
