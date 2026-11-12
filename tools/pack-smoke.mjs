// Checks a packed office-wasm-sdk (.tgz from `npm run pack`) as a consumer would use it: unpacks it into a temp directory,
// imports both entry points, checks that no AGPL code or notice is inside and extracts the text of a sample document.
//   node tools/pack-smoke.mjs dist/office-wasm-sdk-<version>.tgz
// SPDX-License-Identifier: Apache-2.0
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { spawnSync } from 'node:child_process';
import { pathToFileURL, fileURLToPath } from 'node:url';

const tgz = path.resolve(process.argv[2] || '');
if (!process.argv[2] || !fs.existsSync(tgz)) { console.error('usage: node tools/pack-smoke.mjs <office-wasm-sdk-x.tgz>'); process.exit(1); }
const tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'sdkpack-'));
const fails = [];
try {
  // relative paths only: GNU tar reads "E:" as a host name
  const x = spawnSync('tar', ['-xzf', path.basename(tgz), '-C', path.relative(path.dirname(tgz), tmp).split(path.sep).join('/')], { cwd: path.dirname(tgz), stdio: 'inherit' });
  if (x.status !== 0) throw new Error('tar failed');
  const dir = path.join(tmp, 'package');
  const pkg = JSON.parse(fs.readFileSync(path.join(dir, 'package.json'), 'utf8'));
  const walk = (d) => fs.readdirSync(d, { withFileTypes: true }).flatMap((e) => (e.isDirectory() ? walk(path.join(d, e.name)) : [path.join(d, e.name)]));
  const files = walk(dir).map((f) => path.relative(dir, f).split(path.sep).join('/'));
  for (const need of ['LICENSE', 'NOTICE', 'README.md', 'sdk/office-viewer.mjs', 'src/extract/index.mjs']) if (!files.includes(need)) fails.push('missing ' + need);
  for (const f of files.filter((n) => /\.(mjs|js|cjs|md)$/.test(n) || n === 'NOTICE')) {
    const text = fs.readFileSync(path.join(dir, f), 'utf8');
    if (/SPDX-License-Identifier:\s*AGPL/i.test(text)) fails.push('AGPL header in ' + f);
  }
  if (files.some((f) => /^(demo|samples|tools|out)\//.test(f))) fails.push('demo/samples/tools must not be in the package');
  const viewer = await import(pathToFileURL(path.join(dir, pkg.exports['.'])).href);
  if (typeof viewer.OfficeViewer?.create !== 'function') fails.push('OfficeViewer.create missing');
  const sign = await import(pathToFileURL(path.join(dir, pkg.exports['./server'])).href);
  const signed = sign.signViewerUrl({ base: 'https://v.example/', secret: 'x'.repeat(32), ttlSeconds: 60, now: 1000000000 });
  if (!/^https:\/\/v\.example\/s\/1000060\.[A-Za-z0-9_-]{43}\/$/.test(signed)) fails.push('signViewerUrl format: ' + signed);
  const ex = await import(pathToFileURL(path.join(dir, pkg.exports['./extract'])).href);
  const sample = path.join(path.dirname(fileURLToPath(import.meta.url)), '..', 'samples', 'medium.docx');
  const text = await ex.extractText(new Uint8Array(fs.readFileSync(sample)), { name: 'medium.docx' });
  if (!text || text.length < 100) fails.push('extractText returned too little');
  console.log(`${pkg.name} ${pkg.version}: ${files.length} files, license ${pkg.license}, extracted ${text?.length ?? 0} characters`);
} catch (e) { fails.push(String(e.stack || e)); } finally { fs.rmSync(tmp, { recursive: true, force: true }); }
for (const f of fails) console.log('FAIL ' + f);
process.exit(fails.length ? 1 : 0);
