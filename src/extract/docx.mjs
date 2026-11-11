// .docx -> text + block structure (headings, paragraphs, list items, tables, links, notes). No layout.
import { attr, decodeXml, dirOf, parseRels, parseXml, relsPathOf } from './xml.mjs';

/** styleId -> heading level (1..9, 0 = Title) resolved through basedOn. */
function loadStyles(xml) {
  const styles = new Map();
  if (xml) {
    let cur = null;
    parseXml(xml, {
      open(name, a) {
        if (name === 'w:style') cur = { name: '', basedOn: null, outline: null };
        else if (!cur) return;
        else if (name === 'w:name') cur.name = attr(a, 'w:val') ?? '';
        else if (name === 'w:basedOn') cur.basedOn = attr(a, 'w:val') ?? null;
        else if (name === 'w:outlineLvl') cur.outline = Number(attr(a, 'w:val'));
        if (name === 'w:style') cur.id = attr(a, 'w:styleId');
      },
      close(name) {
        if (name === 'w:style' && cur) {
          if (cur.id) styles.set(cur.id, cur);
          cur = null;
        }
      },
    });
  }
  const cache = new Map();
  return function headingLevel(styleId) {
    if (!styleId) return null;
    if (cache.has(styleId)) return cache.get(styleId);
    let level = null;
    let s = styles.get(styleId);
    for (let hops = 0; s && hops < 12; hops++, s = styles.get(s.basedOn)) {
      if (s.outline !== null && s.outline >= 0 && s.outline < 9) {
        level = s.outline + 1;
        break;
      }
      const m = /^heading\s*([1-9])$/i.exec(s.name);
      if (m) {
        level = Number(m[1]);
        break;
      }
      if (/^title$/i.test(s.name)) {
        level = 0;
        break;
      }
    }
    cache.set(styleId, level);
    return level;
  };
}

export function blocksToText(blocks) {
  const out = [];
  for (const b of blocks) out.push(b.type === 'table' ? b.rows.map((r) => r.join('\t')).join('\n') : b.text);
  return out.join('\n');
}

/**
 * Walk a WordprocessingML part (document, header, footnotes ...) and return its blocks.
 * Block: {type: 'paragraph'|'heading'|'list-item', text, level?, style?, links?} | {type: 'table', rows: string[][]}
 */
export function parseBlocks(xml, { headingLevel = () => null, rels = new Map() } = {}) {
  const containers = [{ blocks: [] }]; // body, table cells, text boxes
  const paras = []; // a text box nests paragraphs inside a paragraph
  const tables = [];
  let p = null;
  let inT = false;
  let inPPr = 0;
  let skip = 0; // inside mc:Fallback (duplicate VML rendering of the mc:Choice content)
  let link = null;
  const top = () => containers[containers.length - 1];

  parseXml(xml, {
    open(name, a) {
      if (name === 'mc:Fallback') {
        skip++;
        return;
      }
      if (skip) return;
      switch (name) {
        case 'w:p':
          if (p) paras.push(p);
          p = { text: '', style: null, numId: null, ilvl: 0, outline: null, links: null, after: null };
          inPPr = 0;
          break;
        case 'w:pPr':
          if (p) inPPr++; // a nested w:pPr (inside w:pPrChange) holds the old formatting: ignored
          break;
        case 'w:pStyle':
          if (inPPr === 1) p.style = attr(a, 'w:val') ?? null;
          break;
        case 'w:numId':
          if (inPPr === 1) p.numId = attr(a, 'w:val') ?? null;
          break;
        case 'w:ilvl':
          if (inPPr === 1) p.ilvl = Number(attr(a, 'w:val')) || 0;
          break;
        case 'w:outlineLvl':
          if (inPPr === 1) p.outline = Number(attr(a, 'w:val'));
          break;
        case 'w:t':
        case 'm:t':
          inT = true;
          break;
        case 'w:tab':
        case 'w:ptab':
          if (p && !inPPr) p.text += '\t';
          break;
        case 'w:br':
        case 'w:cr':
          if (p && !inPPr) p.text += '\n';
          break;
        case 'w:noBreakHyphen':
          if (p) p.text += '-';
          break;
        case 'w:hyperlink':
          if (p) {
            const id = attr(a, 'r:id');
            const anchor = attr(a, 'w:anchor');
            link = { start: p.text.length, url: (id && rels.get(id)?.target) || (anchor ? '#' + anchor : '') };
          }
          break;
        case 'w:tbl':
          tables.push({ rows: [], row: null });
          break;
        case 'w:tr':
          if (tables.length) tables[tables.length - 1].row = [];
          break;
        case 'w:tc':
        case 'w:txbxContent':
          containers.push({ blocks: [] });
          break;
        default:
      }
    },
    text(raw, isCdata) {
      if (inT && p && !skip) p.text += isCdata ? raw : decodeXml(raw);
    },
    close(name) {
      if (name === 'mc:Fallback') {
        skip--;
        return;
      }
      if (skip) return;
      switch (name) {
        case 'w:t':
        case 'm:t':
          inT = false;
          break;
        case 'w:pPr':
          if (inPPr) inPPr--;
          break;
        case 'w:hyperlink':
          if (link && p) {
            const text = p.text.slice(link.start);
            if (link.url && text) (p.links ??= []).push({ text, url: link.url });
          }
          link = null;
          break;
        case 'w:p': {
          if (!p) break;
          const done = p;
          p = paras.pop() ?? null;
          inPPr = 0;
          if (done.text.trim()) {
            const block = { type: 'paragraph', text: done.text };
            const level = done.outline !== null && done.outline >= 0 && done.outline < 9 ? done.outline + 1 : headingLevel(done.style);
            if (level !== null) {
              block.type = 'heading';
              block.level = level;
            } else if (done.numId && done.numId !== '0') {
              block.type = 'list-item';
              block.level = done.ilvl;
            }
            if (done.style) block.style = done.style;
            if (done.links) block.links = done.links;
            top().blocks.push(block);
          }
          if (done.after) top().blocks.push(...done.after);
          break;
        }
        case 'w:txbxContent': {
          if (containers.length < 2) break;
          const box = containers.pop();
          if (p) (p.after ??= []).push(...box.blocks); // emitted right after the paragraph that anchors the text box
          else top().blocks.push(...box.blocks);
          break;
        }
        case 'w:tc': {
          if (containers.length < 2) break;
          const cell = containers.pop();
          const t = tables[tables.length - 1];
          if (t && t.row) t.row.push(blocksToText(cell.blocks));
          else top().blocks.push(...cell.blocks);
          break;
        }
        case 'w:tr': {
          const t = tables[tables.length - 1];
          if (t && t.row) {
            t.rows.push(t.row);
            t.row = null;
          }
          break;
        }
        case 'w:tbl': {
          const t = tables.pop();
          if (t && t.rows.length) top().blocks.push({ type: 'table', rows: t.rows });
          break;
        }
        default:
      }
    },
  });
  return containers[0].blocks;
}

/**
 * @param zip   result of openZip()
 * @param {string} mainPath  e.g. "word/document.xml"
 * @param {{notes?: boolean}} [opts]  notes: also read headers, footers, footnotes, endnotes, comments (default true)
 */
export async function extractDocx(zip, mainPath, opts = {}) {
  const rels = parseRels(await zip.text(relsPathOf(mainPath)), dirOf(mainPath));
  const byType = (suffix) => [...rels.values()].filter((r) => !r.external && r.type.endsWith(suffix)).map((r) => r.target);
  const headingLevel = loadStyles(await zip.text(byType('/styles')[0] ?? 'word/styles.xml'));

  const blocks = parseBlocks((await zip.text(mainPath)) ?? '', { headingLevel, rels });
  const result = { type: 'docx', text: blocksToText(blocks), blocks };

  if (opts.notes !== false) {
    const part = async (path) => {
      const xml = await zip.text(path);
      if (!xml) return [];
      const partRels = parseRels(await zip.text(relsPathOf(path)), dirOf(path));
      return parseBlocks(xml, { headingLevel, rels: partRels });
    };
    for (const [key, suffix] of [['headers', '/header'], ['footers', '/footer'], ['footnotes', '/footnotes'], ['endnotes', '/endnotes'], ['comments', '/comments']]) {
      const texts = [];
      for (const path of byType(suffix)) {
        const t = blocksToText(await part(path));
        if (t) texts.push(t);
      }
      if (texts.length) result[key] = texts;
    }
  }

  const stats = { paragraphs: 0, headings: 0, listItems: 0, tables: 0 };
  for (const b of blocks) {
    if (b.type === 'table') stats.tables++;
    else if (b.type === 'heading') stats.headings++;
    else if (b.type === 'list-item') stats.listItems++;
    else stats.paragraphs++;
  }
  result.stats = stats;
  result.images = zip.names().filter((n) => n.startsWith('word/media/'));
  return result;
}
