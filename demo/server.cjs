// Demo host server. No dependencies. Starts the viewer server (tools/serve.cjs of the viewer repo) on port+1 as a child process,
// so the demo page and the viewer run on two origins, as they would in production.
// usage: node demo/server.cjs [port] [--coi] [--viewer-url <url>] [--no-font-aliases]   (default 8765, 127.0.0.1)
//   --coi         host and viewer send COOP/COEP (+ CORP on the viewer): the viewer frame is cross-origin isolated
//                 and uses the multi-threaded x2t build
//   --viewer-url  use a viewer that is already running somewhere else instead of starting one
//   /            -> demo/
//   /sdk/...     -> sdk/
//   /src/...     -> src/            (src/extract: pure-JS text extraction)
//   /samples/... -> samples/
//   /out/...     -> out/
//   /config.json -> {viewerUrl}
'use strict';
const http = require('node:http');
const fs = require('node:fs');
const path = require('node:path');
const { spawn } = require('node:child_process');

const ROOT = path.resolve(__dirname, '..');
// the viewer repo (AGPL) next to this one, or OFFICE_WASM_VIEWER
const VIEWER = process.env.OFFICE_WASM_VIEWER || path.resolve(ROOT, '..', 'office-wasm-viewer');
const MOUNTS = [
  ['/sdk/', path.join(ROOT, 'sdk')],
  ['/src/', path.join(ROOT, 'src')],
  ['/samples/', path.join(ROOT, 'samples')],
  ['/out/', path.join(ROOT, 'out')],
  ['/', __dirname],
];
const TYPES = {
  '.html': 'text/html; charset=utf-8', '.js': 'text/javascript; charset=utf-8', '.mjs': 'text/javascript; charset=utf-8',
  '.css': 'text/css; charset=utf-8', '.json': 'application/json; charset=utf-8', '.png': 'image/png', '.pdf': 'application/pdf',
  '.svg': 'image/svg+xml', '.txt': 'text/plain; charset=utf-8',
  '.docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  '.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  '.pptx': 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
};

const argv = process.argv.slice(2);
const opt = (name) => {
  const i = argv.indexOf('--' + name);
  return i < 0 ? null : argv[i + 1];
};
const COI = argv.includes('--coi');
const port = Number(argv.find((a, i) => /^\d+$/.test(a) && argv[i - 1] !== '--viewer-url') || process.env.PORT || 8765);
let viewerUrl = opt('viewer-url');
let child = null;
if (!viewerUrl) {
  const vport = port + 1;
  viewerUrl = `http://127.0.0.1:${vport}/`;
  const args = [path.join(VIEWER, 'tools', 'serve.cjs'), String(vport), '--frame-ancestors', `'self' http://127.0.0.1:${port}`];
  if (COI) args.push('--coi');
  if (argv.includes('--no-font-aliases')) args.push('--no-font-aliases');
  child = spawn(process.execPath, args, { stdio: ['ignore', 'inherit', 'inherit'] });
  child.on('exit', (code) => {
    if (code) console.error('viewer server exited with ' + code);
  });
  for (const sig of ['SIGINT', 'SIGTERM']) process.on(sig, () => { child.kill(); process.exit(0); });
  process.on('exit', () => child.kill());
}

const HEADERS = { 'Cache-Control': 'no-cache', 'X-Content-Type-Options': 'nosniff', ...(COI ? { 'Cross-Origin-Opener-Policy': 'same-origin', 'Cross-Origin-Embedder-Policy': 'require-corp' } : {}) };
const stats = { requests: 0, bytes: 0 };

function resolve(urlPath) {
  for (const [prefix, dir] of MOUNTS) {
    if (!urlPath.startsWith(prefix)) continue;
    const file = path.resolve(dir, urlPath.slice(prefix.length));
    if (file !== dir && !file.startsWith(dir + path.sep)) return null; // path traversal
    return file;
  }
  return null;
}

const server = http.createServer((req, res) => {
  let urlPath;
  try {
    urlPath = decodeURIComponent(new URL(req.url, 'http://x').pathname);
  } catch {
    res.writeHead(400).end();
    return;
  }
  if (urlPath === '/config.json') {
    res.writeHead(200, { ...HEADERS, 'Content-Type': TYPES['.json'] }).end(JSON.stringify({ viewerUrl }));
    return;
  }
  if (urlPath === '/server-stats.json') {
    res.writeHead(200, { 'Content-Type': TYPES['.json'] }).end(JSON.stringify(stats));
    return;
  }
  if (urlPath.endsWith('/')) urlPath += 'index.html';
  const file = resolve(urlPath);
  if (!file || !fs.existsSync(file) || !fs.statSync(file).isFile()) {
    res.writeHead(404, { 'Content-Type': 'text/plain' }).end('not found: ' + urlPath);
    return;
  }
  const size = fs.statSync(file).size;
  stats.requests++;
  stats.bytes += size;
  res.writeHead(200, { ...HEADERS, 'Content-Type': TYPES[path.extname(file).toLowerCase()] || 'application/octet-stream', 'Content-Length': size });
  if (req.method === 'HEAD') res.end();
  else fs.createReadStream(file).pipe(res);
});

server.listen(port, '127.0.0.1', () => console.log(`demo host: http://127.0.0.1:${port}/  viewer: ${viewerUrl}` + (COI ? ' (cross-origin isolated)' : '')));
