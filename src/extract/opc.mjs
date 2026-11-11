// Open Packaging Conventions helpers shared by the docx / xlsx / pptx extractors.
import { decodeXml, parseRels, parseXml } from './xml.mjs';

export const local = (name) => {
  const i = name.indexOf(':');
  return i < 0 ? name : name.slice(i + 1);
};

/** Path of the main part (word/document.xml, xl/workbook.xml, ppt/presentation.xml) from the package relationships. */
export async function mainPartOf(zip) {
  const rels = parseRels(await zip.text('_rels/.rels'), '');
  for (const r of rels.values()) {
    if (!r.external && r.type.endsWith('/officeDocument') && zip.has(r.target)) return r.target;
  }
  for (const p of ['word/document.xml', 'xl/workbook.xml', 'ppt/presentation.xml']) if (zip.has(p)) return p;
  return null;
}

/** Text of every leaf element, keyed by local name (enough for docProps/core.xml and docProps/app.xml). */
function leafTexts(xml) {
  const out = {};
  if (!xml) return out;
  let cur = null;
  let buf = '';
  parseXml(xml, {
    open(name) {
      cur = name;
      buf = '';
    },
    text(raw, isCdata) {
      if (cur) buf += isCdata ? raw : decodeXml(raw);
    },
    close(name) {
      if (cur === name && buf.trim()) out[local(name)] = buf.trim();
      cur = null;
    },
  });
  return out;
}

export async function readMeta(zip) {
  const core = leafTexts(await zip.text('docProps/core.xml'));
  const app = leafTexts(await zip.text('docProps/app.xml'));
  const meta = {};
  const pick = (src, from, to = from) => {
    if (src[from] !== undefined) meta[to] = src[from];
  };
  pick(core, 'title');
  pick(core, 'subject');
  pick(core, 'creator');
  pick(core, 'keywords');
  pick(core, 'description');
  pick(core, 'lastModifiedBy');
  pick(core, 'created');
  pick(core, 'modified');
  pick(app, 'Application', 'application');
  for (const k of ['Pages', 'Words', 'Characters', 'Slides', 'Paragraphs']) {
    if (app[k] !== undefined && Number.isFinite(Number(app[k]))) meta[k.toLowerCase()] = Number(app[k]);
  }
  return meta;
}
