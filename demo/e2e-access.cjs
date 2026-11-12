// Checks who may use a hosted viewer: the allowed-domain list (frame-ancestors, with *. wildcards, changed while the server runs)
// and the signed-URL gate. Starts the viewer server itself with a temporary config file.
//   node demo/e2e-access.cjs [--site <dir>] [--port 9765] [--headed]
//   --site  serve a packed site (office-wasm-viewer: npm run pack, then dist/<package>/site) instead of the source tree
// Browser cases open the viewer from hosts app.localhost / a.b.localhost / localhost / 127.0.0.1 (Chromium resolves *.localhost to the
// loopback address, so wildcard rules can be tried without DNS). Exit code 1 when any check fails.
'use strict';
const { spawn } = require('node:child_process');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const http = require('node:http');
const crypto = require('node:crypto');

const root = path.join(__dirname, '..');
const VIEWER = process.env.OFFICE_WASM_VIEWER || path.resolve(root, '..', 'office-wasm-viewer');
const arg = (name, dflt) => {
  const i = process.argv.indexOf('--' + name);
  return i < 0 ? dflt : process.argv[i + 1];
};
const hostPort = Number(arg('port', 9765));
const viewerPort = hostPort + 1;
const site = arg('site', '');
const headed = process.argv.includes('--headed');
const SECRET = crypto.randomBytes(24).toString('base64url');
const viewerBase = `http://127.0.0.1:${viewerPort}`;

const results = [];
const check = (name, ok, detail = '') => {
  results.push({ name, ok });
  console.log(`${ok ? 'PASS' : 'FAIL'} ${name}${detail ? '  (' + detail + ')' : ''}`);
};
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

function sign({ secret = SECRET, scope = 'viewer', expires = Math.floor(Date.now() / 1000) + 600, badSig = false } = {}) {
  let sig = crypto.createHmac('sha256', secret).update(`${scope}\n${expires}`).digest('base64url');
  if (badSig) sig = sig.slice(0, -1) + (sig.endsWith('A') ? 'B' : 'A');
  return `${viewerBase}/s/${expires}.${sig}/`;
}
async function status(url) {
  const r = await fetch(url, { redirect: 'manual' });
  return { status: r.status, csp: r.headers.get('content-security-policy') };
}

const cfgFile = path.join(fs.mkdtempSync(path.join(os.tmpdir(), 'access-')), 'viewer.config.json');
const writeCfg = (obj) => fs.writeFileSync(cfgFile, typeof obj === 'string' ? obj : JSON.stringify(obj, null, 2));
const BASE_CFG = { frameAncestors: [`http://*.localhost:${hostPort}`], tokenScope: 'viewer' };
async function waitCsp(pred, ms = 10000) {
  const t0 = Date.now();
  let last = null;
  while (Date.now() - t0 < ms) {
    last = (await status(sign())).csp;
    if (pred(last)) return last;
    await sleep(500);
  }
  return last;
}

// Host page: a tiny page that embeds the viewer with the SDK (the same code an application would use).
const HOST_HTML = `<!doctype html><meta charset="utf-8"><title>host</title><div id="box" style="width:900px;height:600px"></div>
<script type="module">
import { OfficeViewer } from '/sdk/office-viewer.mjs';
const q = new URLSearchParams(location.search);
(async () => {
  try {
    const v = await OfficeViewer.create(document.getElementById('box'), { url: q.get('v'), timeoutMs: Number(q.get('timeout') || 6000) });
    let opened = null;
    if (q.get('open')) {
      const data = await (await fetch('/samples/medium.docx')).arrayBuffer();
      await v.open({ name: 'medium.docx', data, mode: q.get('open') });
      opened = (await v.exportPdf()).data.length;
    }
    window.__result = { ok: true, opened };
  } catch (e) { window.__result = { ok: false, error: String(e.message || e) }; }
})();
</script>`;
const hostServer = http.createServer((req, res) => {
  const u = new URL(req.url, 'http://x').pathname;
  const map = { '/sdk/office-viewer.mjs': ['sdk/office-viewer.mjs', 'text/javascript'], '/samples/medium.docx': ['samples/medium.docx', 'application/octet-stream'] };
  if (u === '/') return res.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8' }).end(HOST_HTML);
  if (!map[u]) return res.writeHead(404).end();
  res.writeHead(200, { 'Content-Type': map[u][1] }).end(fs.readFileSync(path.join(root, map[u][0])));
});

async function tryEmbed(browser, host, viewerUrl, { open = '', timeout = 6000 } = {}) {
  const page = await browser.newPage();
  try {
    await page.goto(`http://${host}:${hostPort}/?v=${encodeURIComponent(viewerUrl)}&timeout=${timeout}&open=${open}`);
    await page.waitForFunction(() => window.__result, null, { timeout: open ? 120000 : timeout + 15000 });
    return await page.evaluate(() => window.__result);
  } finally {
    await page.close();
  }
}

async function launch(chromium) {
  try {
    return await chromium.launch({ headless: !headed });
  } catch (e) {
    const dir = path.join(process.env.LOCALAPPDATA || '', 'ms-playwright');
    const rev = fs.existsSync(dir) ? fs.readdirSync(dir).filter((d) => /^chromium-\d+$/.test(d)).sort().pop() : null;
    const exe = process.env.PW_CHROMIUM || (rev && ['chrome-win64', 'chrome-win'].map((s) => path.join(dir, rev, s, 'chrome.exe')).find(fs.existsSync));
    if (!exe) throw e;
    return chromium.launch({ headless: !headed, executablePath: exe });
  }
}

(async () => {
  writeCfg(BASE_CFG);
  const args = [path.join(VIEWER, 'tools', 'serve.cjs'), String(viewerPort), '--config', cfgFile];
  if (site) args.push('--site', site);
  const child = spawn(process.execPath, args, { stdio: ['ignore', 'inherit', 'inherit'], env: { ...process.env, OFFICE_WASM_TOKEN_SECRET: SECRET } });
  await new Promise((r) => hostServer.listen(hostPort, '127.0.0.1', r));
  let browser;
  try {
    for (let i = 0; i < 50; i++) { try { await fetch(viewerBase + '/'); break; } catch { await sleep(200); } }

    // 1. the token gate (plain HTTP)
    const ok = await status(sign());
    check('valid token: 200', ok.status === 200, String(ok.status));
    check('Content-Security-Policy lists the configured domain and self', ok.csp === `frame-ancestors 'self' http://*.localhost:${hostPort}`, ok.csp);
    check('no token: 403', (await status(viewerBase + '/')).status === 403);
    check('wrong signature: 403', (await status(sign({ badSig: true }))).status === 403);
    check('expired token: 403', (await status(sign({ expires: Math.floor(Date.now() / 1000) - 5 }))).status === 403);
    check('token of another secret: 403', (await status(sign({ secret: 'z'.repeat(32) }))).status === 403);
    check('token of another scope: 403', (await status(sign({ scope: 'other' }))).status === 403);
    check('token without trailing slash: 301', (await status(sign().slice(0, -1))).status === 301);

    const { chromium } = require('playwright');
    browser = await launch(chromium);

    // 2. wildcard: subdomains of localhost are allowed, the bare name and the IP are not
    let r = await tryEmbed(browser, 'app.localhost', sign(), { open: 'desktop' });
    check('app.localhost (matches *.localhost): viewer starts, opens a document in the editor, exports PDF', r.ok && r.opened > 1000, JSON.stringify(r));
    r = await tryEmbed(browser, 'a.b.localhost', sign());
    check('a.b.localhost (two levels): viewer starts', r.ok, JSON.stringify(r));
    r = await tryEmbed(browser, 'localhost', sign());
    check('localhost (bare domain, not matched by *.localhost): refused', !r.ok, JSON.stringify(r));
    r = await tryEmbed(browser, '127.0.0.1', sign());
    check('127.0.0.1: refused', !r.ok, JSON.stringify(r));
    r = await tryEmbed(browser, 'app.localhost', sign({ badSig: true }));
    check('allowed host with a bad signature: refused', !r.ok, JSON.stringify(r));

    // 3. change the list while the server runs: add the bare domain, add educora.vn and *.educora.vn
    writeCfg({ ...BASE_CFG, frameAncestors: [`http://*.localhost:${hostPort}`, `http://localhost:${hostPort}`, 'https://educora.vn', 'https://*.educora.vn'] });
    const csp2 = await waitCsp((c) => c && c.includes('https://*.educora.vn'));
    check('config edited: new list in force without restart', !!csp2 && csp2.includes(`http://localhost:${hostPort}`) && csp2.includes('https://educora.vn ') && csp2.includes('https://*.educora.vn'), csp2);
    r = await tryEmbed(browser, 'localhost', sign());
    check('localhost now allowed', r.ok, JSON.stringify(r));
    r = await tryEmbed(browser, 'app.localhost', sign());
    check('app.localhost still allowed', r.ok, JSON.stringify(r));
    r = await tryEmbed(browser, '127.0.0.1', sign());
    check('127.0.0.1 still refused', !r.ok, JSON.stringify(r));

    // 4. a broken edit keeps the previous list
    for (const bad of ['{ not json', JSON.stringify({ frameAncestors: ['https://*.vn'] }), JSON.stringify({ frameAncestors: ['https://a.com; script-src *'] }), JSON.stringify({ frameAncestors: ['http://evil.example'] })]) {
      writeCfg(bad);
      await sleep(3500);
      const c = (await status(sign())).csp;
      check(`invalid config ${bad.slice(0, 40)}: previous list kept`, c === csp2, c);
    }

    // 5. remove a domain: it stops working at once
    writeCfg({ ...BASE_CFG, frameAncestors: ['https://educora.vn', 'https://*.educora.vn'] });
    const csp3 = await waitCsp((c) => c && !c.includes('localhost'));
    check('config edited: localhost domains removed', !!csp3 && !csp3.includes('localhost'), csp3);
    r = await tryEmbed(browser, 'app.localhost', sign());
    check('app.localhost refused after removal', !r.ok, JSON.stringify(r));
  } catch (e) {
    check('run completed without an exception', false, String(e.stack || e));
  } finally {
    if (browser) await browser.close();
    hostServer.close();
    child.kill();
    fs.rmSync(path.dirname(cfgFile), { recursive: true, force: true });
  }
  const failed = results.filter((x) => !x.ok).length;
  console.log(`\n${results.length - failed}/${results.length} passed`);
  process.exit(failed ? 1 : 0);
})();
