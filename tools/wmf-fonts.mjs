// Lists the fonts and character codes used by the text records of every WMF picture in an OOXML file.
//   node tools/wmf-fonts.mjs samples/local/de-mau.docx
// MathType equation previews are WMF pictures; this shows which fonts (MT Extra, Symbol, Euclid ...) a document needs.
import fs from 'node:fs';
import { openZip } from '../src/extract/zip.mjs';

const META_CREATEFONTINDIRECT = 0x02fb;
const META_SELECTOBJECT = 0x012d;
const META_DELETEOBJECT = 0x01f0;
const META_EXTTEXTOUT = 0x0a32;
const META_TEXTOUT = 0x0521;
// Records that take a slot in the object table without being a font: brush, pen, palette, pattern brush, region, DIB brush.
const OTHER_OBJECTS = [0x02fc, 0x02fa, 0x00f7, 0x01f9, 0x06ff, 0x0142];

/** @returns {{face: string, charset: number, italic: boolean, bold: boolean, codes: number[]}[]} one entry per text record */
export function wmfTextRuns(data) {
  const dv = new DataView(data.buffer, data.byteOffset, data.byteLength);
  let p = dv.getUint32(0, true) === 0x9ac6cdd7 ? 22 : 0; // placeable header
  p += dv.getUint16(p + 2, true) * 2;
  const objs = [];
  const put = (o) => {
    let k = objs.indexOf(null);
    if (k < 0) k = objs.length;
    objs[k] = o;
  };
  const runs = [];
  let cur = null;
  while (p + 6 <= data.length) {
    const size = dv.getUint32(p, true) * 2;
    const fn = dv.getUint16(p + 4, true);
    if (size < 6 || fn === 0) break;
    const a = p + 6;
    if (fn === META_CREATEFONTINDIRECT) {
      let face = '';
      for (let i = a + 18; i < p + size && data[i]; i++) face += String.fromCharCode(data[i]);
      put({ face, charset: data[a + 14], italic: !!data[a + 10], bold: dv.getInt16(a + 8, true) >= 700 });
    } else if (OTHER_OBJECTS.includes(fn)) {
      put({});
    } else if (fn === META_SELECTOBJECT) {
      const o = objs[dv.getUint16(a, true)];
      if (o && o.face !== undefined) cur = o;
    } else if (fn === META_DELETEOBJECT) {
      objs[dv.getUint16(a, true)] = null;
    } else if (fn === META_EXTTEXTOUT || fn === META_TEXTOUT) {
      let len;
      let off;
      if (fn === META_EXTTEXTOUT) {
        len = dv.getUint16(a + 4, true);
        off = a + 8 + (dv.getUint16(a + 6, true) & 6 ? 8 : 0);
      } else {
        len = dv.getUint16(a, true);
        off = a + 2;
      }
      runs.push({ ...(cur || { face: '?', charset: 0, italic: false, bold: false }), codes: [...data.subarray(off, off + len)] });
    }
    p += size;
  }
  return runs;
}

if (process.argv[1] && import.meta.url.endsWith(process.argv[1].replace(/\\/g, '/').split('/').pop())) {
  const zip = openZip(new Uint8Array(fs.readFileSync(process.argv[2])));
  const fonts = new Map();
  let files = 0;
  for (const name of zip.names()) {
    if (!/\.wmf$/i.test(name)) continue;
    files++;
    for (const run of wmfTextRuns(await zip.bytes(name))) {
      const key = `${run.face} (charset ${run.charset}${run.italic ? ', italic' : ''}${run.bold ? ', bold' : ''})`;
      if (!fonts.has(key)) fonts.set(key, new Map());
      const m = fonts.get(key);
      for (const c of run.codes) m.set(c, (m.get(c) || 0) + 1);
    }
  }
  console.log('WMF pictures: ' + files);
  for (const [key, m] of [...fonts].sort()) {
    const codes = [...m].sort((x, y) => x[0] - y[0]);
    console.log(`${key}: ${codes.reduce((s, c) => s + c[1], 0)} characters`);
    console.log('  ' + codes.map(([c, n]) => (c > 0x20 && c < 0x7f ? String.fromCharCode(c) : '') + '<' + c.toString(16).padStart(2, '0') + '>x' + n).join(' '));
  }
}
