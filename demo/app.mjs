// Demo host page: embeds the office-wasm viewer (separate origin, AGPL-3.0) through sdk/office-viewer.mjs and
// extracts text with src/extract (pure JS). No viewer code is loaded into this page.
import { OfficeViewer } from '/sdk/office-viewer.mjs';

const $ = (id) => document.getElementById(id);
const now = () => performance.now();
const ms = (t0) => Math.round(now() - t0);

// Inspectable from tests: window.__state
const state = (window.__state = { phase: 'starting', events: [], timings: {}, error: null, pdf: null, text: null, x2t: null });
function log(...parts) {
  const line = parts.map((p) => (typeof p === 'string' ? p : JSON.stringify(p))).join(' ');
  state.events.push(line);
  $('log').textContent += line + '\n';
  $('log').scrollTop = $('log').scrollHeight;
}

const q = new URLSearchParams(location.search);
const cfg = await fetch('/config.json').then((r) => r.json());
const t0 = now();
const viewer = await OfficeViewer.create($('stage'), {
  url: cfg.viewerUrl,
  params: { x2t: q.get('x2t'), threads: q.get('threads'), mathFonts: q.get('mathFonts') },
});
state.timings.viewerStartMs = ms(t0);
state.x2t = viewer.info.x2t;
state.phase = 'idle';
log(`viewer ${viewer.info.version} (${cfg.viewerUrl}) sẵn sàng: x2t ${viewer.info.x2t?.build}${viewer.info.x2t?.threads ? '/' + viewer.info.x2t.threads + ' luồng' : ''}, ${state.timings.viewerStartMs} ms`);
viewer.on('log', (line) => log('[viewer] ' + line));
viewer.on('phase', ({ phase, error }) => {
  state.phase = phase;
  if (error) state.error = error;
});
viewer.on('editorError', (e) => log('editor error', e));
viewer.on('pdf', ({ name, data }) => download(new Uint8Array(data), name, 'application/pdf')); // viewer's own Print button

let doc = null; // {name, bytes}

let modified = false;
viewer.on('modified', (e) => {
  modified = !!e.modified;
  $('save').textContent = modified ? 'Lưu file đã sửa *' : 'Lưu file đã sửa';
});

function setButtons(enabled) {
  for (const id of ['pdf', 'text', 'x2ttext', 'showdoc']) $(id).disabled = !enabled;
  $('save').disabled = !(enabled && $('mode').value === 'edit');
}

async function saveFile() {
  const t = now();
  const r = await viewer.save();
  state.timings.saveWallMs = ms(t);
  state.saved = { name: r.name, bytes: r.data };
  log(`đã lưu ${r.name}: ${r.data.length} B, ${state.timings.saveWallMs} ms`);
  download(r.data, r.name, 'application/octet-stream');
  return state.saved;
}

async function openFile(name, bytes) {
  setButtons(false);
  modified = false;
  $('save').textContent = 'Lưu file đã sửa';
  $('drop').style.display = 'none';
  $('textview').style.display = 'none';
  Object.assign(state, { error: null, pdf: null, text: null });
  doc = { name, bytes };
  const t = now();
  try {
    const r = await viewer.open({ name, data: bytes, mode: $('mode').value, localFonts: $('localfonts').checked ? true : undefined });
    state.timings = { ...state.timings, ...r.timings, openWallMs: ms(t) };
    state.opened = r;
    setButtons(true);
  } catch (e) {
    if (e.message === 'superseded') return;
    state.phase = 'error';
    state.error = e.message;
    log('Lỗi mở file: ' + e.message);
  }
}

async function exportPdf() {
  const t = now();
  const r = await viewer.exportPdf();
  state.timings = { ...state.timings, ...r.timings, pdfWallMs: ms(t) };
  state.pdf = { bytes: r.data, pdfBinBytes: r.pdfBinBytes, x2tMs: r.x2tMs, mathPictures: r.mathPictures };
  return state.pdf;
}

function download(bytes, name, type) {
  const a = document.createElement('a');
  a.href = URL.createObjectURL(new Blob([bytes], { type }));
  a.download = name;
  a.click();
  setTimeout(() => URL.revokeObjectURL(a.href), 60000);
}

// ---------- text extraction ----------
function showText(text) {
  $('textview').textContent = text;
  $('textview').style.display = 'block';
}

async function extractJs() {
  const { extract } = await import('/src/extract/index.mjs');
  const t = now();
  const r = await extract(doc.bytes);
  state.timings.extractJsMs = ms(t);
  state.text = { via: 'js', chars: r.text.length, type: r.type };
  log(`trích văn bản (JS thuần): ${r.text.length} ký tự, ${state.timings.extractJsMs} ms`);
  return r;
}

async function extractX2t() {
  const t = now();
  const to = /\.(xlsx|xls|ods|csv|xlsm|xltx|xlsb)$/i.test(doc.name) ? 'out.csv' : 'out.txt';
  const r = await viewer.convert({ to });
  const text = new TextDecoder().decode(r.data);
  state.timings.extractX2tMs = ms(t);
  state.text = { via: 'x2t', chars: text.length };
  log(`trích văn bản (x2t): ${text.length} ký tự, ${state.timings.extractX2tMs} ms (x2t ${r.convertMs} ms)`);
  return text;
}

// ---------- UI wiring ----------
const guard = (fn) => async () => {
  try {
    await fn();
  } catch (e) {
    state.error = e.message;
    log('Lỗi: ' + e.message);
  }
};
$('file').addEventListener('change', async (ev) => {
  const f = ev.target.files[0];
  if (f) await openFile(f.name, new Uint8Array(await f.arrayBuffer()));
});
for (const b of document.querySelectorAll('button[data-sample]')) b.addEventListener('click', () => openSample(b.dataset.sample));
$('pdf').addEventListener('click', guard(async () => {
  const pdf = await exportPdf();
  download(pdf.bytes, doc.name.replace(/\.[^.]+$/, '') + '.pdf', 'application/pdf');
}));
$('text').addEventListener('click', guard(async () => showText((await extractJs()).text)));
$('x2ttext').addEventListener('click', guard(async () => showText(await extractX2t())));
$('showdoc').addEventListener('click', () => { $('textview').style.display = 'none'; });
$('save').addEventListener('click', guard(saveFile));
$('mode').addEventListener('change', () => {
  if (modified && !confirm('Đổi giao diện sẽ mở lại file gốc và mất phần đã sửa chưa lưu. Tiếp tục?')) return;
  if (doc) openFile(doc.name, doc.bytes);
});
// Fonts of this computer (Local Font Access API): the permission prompt needs this click.
async function useLocalFonts(on) {
  if (on) {
    try {
      const n = await viewer.enableLocalFonts();
      log(`font trong máy: ${n} kiểu chữ`);
    } catch (e) {
      $('localfonts').checked = false;
      log('Không dùng được font trong máy: ' + e.message);
      return;
    }
  }
  if (doc) openFile(doc.name, doc.bytes);
}
$('localfonts').addEventListener('change', () => useLocalFonts($('localfonts').checked));
document.addEventListener('dragover', (ev) => ev.preventDefault());
document.addEventListener('drop', async (ev) => {
  ev.preventDefault();
  const f = ev.dataTransfer.files[0];
  if (f) await openFile(f.name, new Uint8Array(await f.arrayBuffer()));
});

async function openSample(name) {
  const r = await fetch('/samples/' + encodeURIComponent(name));
  if (!r.ok) {
    log('không tải được mẫu: ' + name);
    return;
  }
  await openFile(name, new Uint8Array(await r.arrayBuffer()));
}

// Automation hooks (used by demo/e2e.cjs)
window.__poc = {
  openSample,
  exportPdf,
  saveFile,
  extractJs,
  extractX2t,
  getState: () => viewer.getState(),
  pdfBase64() {
    if (!state.pdf) return null;
    let s = '';
    const b = state.pdf.bytes;
    for (let i = 0; i < b.length; i += 0x8000) s += String.fromCharCode.apply(null, b.subarray(i, i + 0x8000));
    return btoa(s);
  },
};

if (q.get('mode')) $('mode').value = q.get('mode');
if (q.get('sample')) {
  (async () => {
    if (q.get('localfonts')) { // automation: no click is possible, the test supplies its own queryLocalFonts
      $('localfonts').checked = true;
      try { await viewer.enableLocalFonts(); } catch (e) { log('font trong máy: ' + e.message); $('localfonts').checked = false; }
    }
    await openSample(q.get('sample'));
  })();
}
