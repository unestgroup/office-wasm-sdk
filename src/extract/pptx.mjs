// .pptx -> slides in presentation order: title, text shapes (with outline levels), tables, pictures, speaker notes.
import { attr, decodeXml, dirOf, parseRels, parseXml, relsPathOf } from './xml.mjs';

const SHAPE_TAGS = new Set(['p:sp', 'p:cxnSp', 'p:pic', 'p:graphicFrame']);

/** Shapes of one slide / notes part, in document (z-) order. */
export function parseShapes(xml, rels = new Map()) {
  const shapes = [];
  let shape = null;
  let para = null;
  let inT = false;
  let tbl = null;
  let row = null;
  let cell = null;
  let skip = 0;

  parseXml(xml, {
    open(name, a) {
      if (name === 'mc:Fallback') {
        skip++;
        return;
      }
      if (skip) return;
      if (SHAPE_TAGS.has(name)) {
        shape = { kind: name === 'p:pic' ? 'picture' : name === 'p:graphicFrame' ? 'graphic' : 'text', name: '', named: false };
        if (shape.kind === 'text') shape.paragraphs = [];
        return;
      }
      if (!shape) return;
      switch (name) {
        case 'p:cNvPr':
          if (!shape.named) {
            shape.named = true;
            shape.name = attr(a, 'name') ?? '';
            const alt = attr(a, 'descr');
            if (alt) shape.alt = alt;
          }
          break;
        case 'p:ph':
          shape.placeholder = attr(a, 'type') ?? 'obj';
          break;
        case 'a:blip':
          if (shape.kind === 'picture') {
            const target = rels.get(attr(a, 'r:embed') ?? attr(a, 'r:link'))?.target;
            if (target) shape.image = target;
          }
          break;
        case 'a:tbl':
          tbl = [];
          break;
        case 'a:tr':
          row = [];
          break;
        case 'a:tc':
          cell = [];
          break;
        case 'a:p':
          para = { text: '', level: 0 };
          break;
        case 'a:pPr':
          if (para) para.level = Number(attr(a, 'lvl')) || 0;
          break;
        case 'a:t':
          inT = true;
          break;
        case 'a:br':
          if (para) para.text += '\n';
          break;
        default:
      }
    },
    text(raw, isCdata) {
      if (inT && para && !skip) para.text += isCdata ? raw : decodeXml(raw);
    },
    close(name) {
      if (name === 'mc:Fallback') {
        skip--;
        return;
      }
      if (skip || !shape) return;
      switch (name) {
        case 'a:t':
          inT = false;
          break;
        case 'a:p':
          if (para) {
            if (cell) cell.push(para.text);
            else if (shape.paragraphs && para.text.trim()) shape.paragraphs.push(para);
          }
          para = null;
          break;
        case 'a:tc':
          if (row && cell) row.push(cell.filter(Boolean).join('\n'));
          cell = null;
          break;
        case 'a:tr':
          if (tbl && row) tbl.push(row);
          row = null;
          break;
        case 'a:tbl':
          if (tbl) {
            shape.kind = 'table';
            shape.rows = tbl;
          }
          tbl = null;
          break;
        default:
          if (SHAPE_TAGS.has(name)) {
            delete shape.named;
            const keep = shape.kind === 'table' || shape.kind === 'picture' || (shape.kind === 'text' && shape.paragraphs.length);
            if (keep) shapes.push(shape);
            shape = null;
          }
      }
    },
  });
  return shapes;
}

export function shapeText(shape) {
  if (shape.kind === 'table') return shape.rows.map((r) => r.join('\t')).join('\n');
  if (shape.kind === 'text') return shape.paragraphs.map((p) => p.text).join('\n');
  return '';
}

const NOT_CONTENT = new Set(['sldNum', 'dt', 'ftr', 'hdr', 'sldImg']);

/**
 * @param zip   result of openZip()
 * @param {string} mainPath  e.g. "ppt/presentation.xml"
 * @param {{notes?: boolean}} [opts]  notes: read speaker notes (default true)
 */
export async function extractPptx(zip, mainPath, opts = {}) {
  const rels = parseRels(await zip.text(relsPathOf(mainPath)), dirOf(mainPath));
  const slidePaths = [];
  let size = null;
  parseXml((await zip.text(mainPath)) ?? '', {
    open(name, a) {
      if (name === 'p:sldId') {
        const target = rels.get(attr(a, 'r:id'))?.target;
        if (target) slidePaths.push(target);
      } else if (name === 'p:sldSz') {
        size = { cx: Number(attr(a, 'cx')), cy: Number(attr(a, 'cy')) }; // EMU (914400 per inch)
      }
    },
  });

  const slides = [];
  for (let i = 0; i < slidePaths.length; i++) {
    const path = slidePaths[i];
    const xml = await zip.text(path);
    if (!xml) continue;
    const slideRels = parseRels(await zip.text(relsPathOf(path)), dirOf(path));
    const shapes = parseShapes(xml, slideRels).filter((s) => !NOT_CONTENT.has(s.placeholder));
    const titleShape = shapes.find((s) => s.placeholder === 'title' || s.placeholder === 'ctrTitle');
    const slide = { index: i + 1, path, title: titleShape ? shapeText(titleShape).replace(/\s*\n\s*/g, ' ') : null, shapes };
    slide.text = shapes.map(shapeText).filter(Boolean).join('\n');

    if (opts.notes !== false) {
      const notesPath = [...slideRels.values()].find((r) => !r.external && r.type.endsWith('/notesSlide'))?.target;
      const notesXml = notesPath && (await zip.text(notesPath));
      if (notesXml) {
        const notes = parseShapes(notesXml)
          .filter((s) => s.kind === 'text' && !NOT_CONTENT.has(s.placeholder))
          .map(shapeText)
          .filter(Boolean)
          .join('\n');
        if (notes) slide.notes = notes;
      }
    }
    slides.push(slide);
  }

  const text = slides
    .map((s) => s.text + (s.notes ? '\n[notes]\n' + s.notes : ''))
    .filter(Boolean)
    .join('\n\n');
  return {
    type: 'pptx',
    text,
    slides,
    size,
    stats: { slides: slides.length, shapes: slides.reduce((n, s) => n + s.shapes.length, 0), notes: slides.filter((s) => s.notes).length },
  };
}
