// .xlsx -> sheets with typed cell values (cached results), formulas, merges; text as TSV. No recalculation.
import { attr, decodeXml, dirOf, parseRels, parseXml, relsPathOf } from './xml.mjs';
import { local } from './opc.mjs';

const BUILTIN_DATE_FORMATS = new Set([14, 15, 16, 17, 18, 19, 20, 21, 22, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 45, 46, 47, 50, 51, 52, 53, 54, 55, 56, 57, 58]);

export function isDateFormat(code) {
  const s = code
    .replace(/"[^"]*"/g, '') // literals
    .replace(/\\./g, '') // escaped characters
    .replace(/[_*]./g, '') // padding / fill characters
    .replace(/\[(h+|m+|s+)\]/gi, 'h') // elapsed time
    .replace(/\[[^\]]*\]/g, ''); // colours, conditions, locales
  return /[dmyhs]/i.test(s);
}

/** Excel serial number -> ISO string (date, date-time, or time of day for serials below 1). */
export function serialToIso(serial, date1904 = false) {
  if (!Number.isFinite(serial)) return null;
  const days = Math.floor(serial);
  const secs = Math.round((serial - days) * 86400);
  const base = date1904 ? Date.UTC(1904, 0, 1) : Date.UTC(1899, 11, 30);
  const d = new Date(base + days * 86400000 + secs * 1000);
  if (Number.isNaN(d.getTime())) return null;
  const iso = d.toISOString();
  if (days === 0 && !date1904) return iso.slice(11, 19);
  return secs === 0 ? iso.slice(0, 10) : iso.slice(0, 19);
}

function colIndex(ref) {
  let n = 0;
  for (let i = 0; i < ref.length; i++) {
    const c = ref.charCodeAt(i) & ~32; // upper-case
    if (c < 65 || c > 90) break;
    n = n * 26 + (c - 64);
  }
  return n - 1;
}

function parseSharedStrings(xml) {
  const out = [];
  if (!xml) return out;
  let cur = '';
  let inT = false;
  let rph = 0;
  parseXml(xml, {
    open(n) {
      const name = local(n);
      if (name === 'si') cur = '';
      else if (name === 'rPh') rph++; // phonetic run: not part of the value
      else if (name === 't') inT = rph === 0;
    },
    text(raw, isCdata) {
      if (inT) cur += isCdata ? raw : decodeXml(raw);
    },
    close(n) {
      const name = local(n);
      if (name === 't') inT = false;
      else if (name === 'rPh') rph--;
      else if (name === 'si') out.push(cur);
    },
  });
  return out;
}

/** @returns {boolean[]} per cellXfs index: is the number format a date/time format */
function parseDateStyles(xml) {
  const xfIsDate = [];
  if (!xml) return xfIsDate;
  const custom = new Map();
  let inCellXfs = false;
  parseXml(xml, {
    open(n, a) {
      const name = local(n);
      if (name === 'numFmt') custom.set(Number(attr(a, 'numFmtId')), attr(a, 'formatCode') ?? '');
      else if (name === 'cellXfs') inCellXfs = true;
      else if (name === 'xf' && inCellXfs) {
        const id = Number(attr(a, 'numFmtId')) || 0;
        xfIsDate.push(custom.has(id) ? isDateFormat(custom.get(id)) : BUILTIN_DATE_FORMATS.has(id));
      }
    },
    close(n) {
      if (local(n) === 'cellXfs') inCellXfs = false;
    },
  });
  return xfIsDate;
}

function parseSheet(xml, ctx, maxRows) {
  const rows = [];
  const merges = [];
  let dimension = null;
  let row = null;
  let cell = null;
  let cap = 0; // 1 = <v>, 2 = <f>, 3 = inline string <is><t>
  let buf = '';
  let inIs = false;
  let rph = 0;
  let nextRow = 1;
  let nextCol = 0;
  let cells = 0;

  parseXml(xml, {
    open(n, a) {
      const name = local(n);
      switch (name) {
        case 'c': {
          if (!row) break;
          const ref = attr(a, 'r');
          const c = ref ? colIndex(ref) : nextCol;
          nextCol = c + 1;
          cell = { c, t: attr(a, 't'), s: attr(a, 's'), v: null, f: null, is: null };
          break;
        }
        case 'v':
          if (cell) {
            cap = 1;
            buf = '';
          }
          break;
        case 'f':
          if (cell) {
            cap = 2;
            buf = '';
          }
          break;
        case 'is':
          if (cell) {
            inIs = true;
            cell.is = '';
          }
          break;
        case 'rPh':
          rph++;
          break;
        case 't':
          if (inIs && rph === 0) {
            cap = 3;
            buf = '';
          }
          break;
        case 'row': {
          const r = Number(attr(a, 'r')) || nextRow;
          nextRow = r + 1;
          nextCol = 0;
          row = rows.length < maxRows ? { r, cells: [] } : null;
          break;
        }
        case 'mergeCell': {
          const ref = attr(a, 'ref');
          if (ref) merges.push(ref);
          break;
        }
        case 'dimension':
          dimension = attr(a, 'ref') ?? null;
          break;
        default:
      }
    },
    text(raw, isCdata) {
      if (cap) buf += isCdata ? raw : decodeXml(raw);
    },
    close(n) {
      const name = local(n);
      switch (name) {
        case 'v':
          if (cap === 1) cell.v = buf;
          cap = 0;
          break;
        case 'f':
          if (cap === 2 && buf) cell.f = buf;
          cap = 0;
          break;
        case 't':
          if (cap === 3) cell.is += buf;
          cap = 0;
          break;
        case 'rPh':
          rph--;
          break;
        case 'is':
          inIs = false;
          break;
        case 'c': {
          if (!cell || !row) {
            cell = null;
            break;
          }
          const out = { c: cell.c };
          const raw = cell.v;
          switch (cell.t) {
            case 's':
              out.v = ctx.strings[Number(raw)] ?? '';
              out.t = 's';
              break;
            case 'inlineStr':
              out.v = cell.is ?? '';
              out.t = 's';
              break;
            case 'str':
              out.v = raw ?? '';
              out.t = 's';
              break;
            case 'b':
              out.v = raw === '1' || raw === 'true';
              out.t = 'b';
              break;
            case 'e':
              out.v = raw ?? '';
              out.t = 'e';
              break;
            case 'd':
              out.v = raw ?? '';
              out.t = 'd';
              break;
            default:
              if (raw !== null && raw !== '') {
                const num = Number(raw);
                const iso = ctx.xfIsDate[Number(cell.s)] ? serialToIso(num, ctx.date1904) : null;
                if (iso) {
                  out.v = iso;
                  out.t = 'd';
                  out.n = num; // original serial number
                } else {
                  out.v = num;
                  out.t = 'n';
                }
              }
          }
          if (cell.f) out.f = cell.f;
          if (out.v !== undefined && out.v !== '' || out.f) {
            row.cells.push(out);
            cells++;
          }
          cell = null;
          break;
        }
        case 'row':
          if (row && row.cells.length) rows.push(row);
          row = null;
          break;
        default:
      }
    },
  });
  return { dimension, rows, merges, cells };
}

function cellText(cell) {
  if (cell.v === undefined) return '';
  if (cell.t === 'b') return cell.v ? 'TRUE' : 'FALSE';
  const s = String(cell.v);
  return s.includes('\t') || s.includes('\n') ? s.replace(/[\t\r\n]+/g, ' ') : s;
}

export function sheetToText(sheet) {
  const lines = [];
  for (const row of sheet.rows) {
    const cols = [];
    for (const cell of row.cells) {
      while (cols.length < cell.c) cols.push('');
      cols[cell.c] = cellText(cell);
    }
    lines.push(cols.join('\t'));
  }
  return lines.join('\n');
}

/**
 * @param zip   result of openZip()
 * @param {string} mainPath  e.g. "xl/workbook.xml"
 * @param {{maxRows?: number, sheets?: (string|number)[]}} [opts]  maxRows per sheet; sheets: names or 0-based indexes to read
 */
export async function extractXlsx(zip, mainPath, opts = {}) {
  const rels = parseRels(await zip.text(relsPathOf(mainPath)), dirOf(mainPath));
  const first = (suffix) => [...rels.values()].find((r) => !r.external && r.type.endsWith(suffix))?.target;

  const sheetList = [];
  let date1904 = false;
  parseXml((await zip.text(mainPath)) ?? '', {
    open(n, a) {
      const name = local(n);
      if (name === 'sheet') {
        sheetList.push({ name: attr(a, 'name') ?? '', path: rels.get(attr(a, 'r:id'))?.target ?? null, state: attr(a, 'state') ?? 'visible' });
      } else if (name === 'workbookPr') {
        const v = attr(a, 'date1904');
        date1904 = v === '1' || v === 'true';
      }
    },
  });

  const ctx = {
    strings: parseSharedStrings(await zip.text(first('/sharedStrings') ?? 'xl/sharedStrings.xml')),
    xfIsDate: parseDateStyles(await zip.text(first('/styles') ?? 'xl/styles.xml')),
    date1904,
  };

  const maxRows = opts.maxRows ?? Infinity;
  const wanted = opts.sheets ? new Set(opts.sheets) : null;
  const sheets = [];
  for (let i = 0; i < sheetList.length; i++) {
    const s = sheetList[i];
    if (wanted && !wanted.has(s.name) && !wanted.has(i)) continue;
    const xml = s.path && s.path.endsWith('.xml') ? await zip.text(s.path) : undefined; // chart sheets etc. have no cells
    const parsed = xml ? parseSheet(xml, ctx, maxRows) : { dimension: null, rows: [], merges: [], cells: 0 };
    sheets.push({ name: s.name, index: i, state: s.state, ...parsed });
  }

  const text = sheets
    .filter((s) => s.rows.length)
    .map((s) => (sheets.length > 1 ? '# ' + s.name + '\n' : '') + sheetToText(s))
    .join('\n\n');
  return {
    type: 'xlsx',
    text,
    sheets,
    date1904,
    stats: { sheets: sheets.length, rows: sheets.reduce((n, s) => n + s.rows.length, 0), cells: sheets.reduce((n, s) => n + s.cells, 0) },
  };
}
