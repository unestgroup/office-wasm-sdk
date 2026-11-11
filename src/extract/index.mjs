// Fast text + structure extraction from OOXML files (.docx / .xlsx / .pptx and their macro/template variants).
// Pure JS, no dependencies, no OnlyOffice code: works in Node >= 20, browsers and workers.
// Legacy binary formats (.doc / .xls / .ppt), ODF and RTF are not handled here: convert them with x2t first.
import { openZip } from './zip.mjs';
import { mainPartOf, readMeta } from './opc.mjs';
import { extractDocx } from './docx.mjs';
import { extractXlsx } from './xlsx.mjs';
import { extractPptx } from './pptx.mjs';

export class UnsupportedFormatError extends Error {
  constructor(message) {
    super(message);
    this.name = 'UnsupportedFormatError';
    this.code = 'UNSUPPORTED_FORMAT';
  }
}

/**
 * @param {Uint8Array|ArrayBuffer} input  file bytes
 * @param {{notes?: boolean, meta?: boolean, maxRows?: number, sheets?: (string|number)[], maxEntryBytes?: number}} [opts]
 * @returns {Promise<{type: 'docx'|'xlsx'|'pptx', text: string, meta?: object, stats: object}>} plus
 *   docx: blocks[], headers/footers/footnotes/endnotes/comments, images[]
 *   xlsx: sheets[{name, state, dimension, rows[{r, cells[{c, v, t, f?}]}], merges[]}], date1904
 *   pptx: slides[{index, title, shapes[], text, notes?}], size
 */
export async function extract(input, opts = {}) {
  const u8 = input instanceof Uint8Array ? input : new Uint8Array(input);
  if (u8.length < 4 || u8[0] !== 0x50 || u8[1] !== 0x4b) {
    const ole = u8.length > 8 && u8[0] === 0xd0 && u8[1] === 0xcf && u8[2] === 0x11 && u8[3] === 0xe0;
    throw new UnsupportedFormatError(ole ? 'legacy binary Office file (or encrypted OOXML): convert with x2t first' : 'not an OOXML (zip) file');
  }
  const zip = openZip(u8, { maxEntryBytes: opts.maxEntryBytes });
  const main = await mainPartOf(zip);
  if (!main) throw new UnsupportedFormatError('zip file without an Office main part (ODF? convert with x2t first)');

  let result;
  if (main.startsWith('word/')) result = await extractDocx(zip, main, opts);
  else if (main.startsWith('xl/')) {
    if (main.endsWith('.bin')) throw new UnsupportedFormatError('.xlsb (binary workbook): convert with x2t first');
    result = await extractXlsx(zip, main, opts);
  } else if (main.startsWith('ppt/')) result = await extractPptx(zip, main, opts);
  else throw new UnsupportedFormatError('unknown main part: ' + main);

  if (opts.meta !== false) result.meta = await readMeta(zip);
  return result;
}

/** Plain text only. */
export async function extractText(input, opts = {}) {
  return (await extract(input, { ...opts, meta: false })).text;
}

export { openZip } from './zip.mjs';
