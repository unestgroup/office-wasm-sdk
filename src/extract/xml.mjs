// Forward-only XML tokenizer, enough for OOXML parts (no DTD, element names kept with their prefix).

const ENT = { amp: '&', lt: '<', gt: '>', quot: '"', apos: "'" };

export function decodeXml(s) {
  if (s.indexOf('&') < 0) return s;
  return s.replace(/&(#x[0-9a-fA-F]+|#\d+|[a-zA-Z]+);/g, (m, e) => {
    if (e.charCodeAt(0) === 35) {
      const code = e.charCodeAt(1) === 120 ? parseInt(e.slice(2), 16) : parseInt(e.slice(1), 10);
      return Number.isFinite(code) && code > 0 && code <= 0x10ffff ? String.fromCodePoint(code) : '';
    }
    return ENT[e] ?? m;
  });
}

/** Read one attribute from the raw attribute string of a start tag. */
export function attr(attrs, name) {
  if (!attrs) return undefined;
  let p = 0;
  for (;;) {
    p = attrs.indexOf(name, p);
    if (p < 0) return undefined;
    const before = p === 0 ? 32 : attrs.charCodeAt(p - 1);
    let q = p + name.length;
    while (attrs.charCodeAt(q) <= 32) q++;
    if (before <= 32 && attrs.charCodeAt(q) === 61) {
      q++;
      while (attrs.charCodeAt(q) <= 32) q++;
      const quote = attrs[q];
      const end = attrs.indexOf(quote, q + 1);
      return end < 0 ? undefined : decodeXml(attrs.slice(q + 1, end));
    }
    p += name.length;
  }
}

/**
 * @param {string} xml
 * @param {{open?: (name: string, attrs: string) => void, close?: (name: string) => void, text?: (raw: string, isCdata?: boolean) => void}} h
 *   `text` receives the raw (still entity-encoded) text unless isCdata is true.
 */
export function parseXml(xml, h) {
  const n = xml.length;
  let i = 0;
  while (i < n) {
    const lt = xml.indexOf('<', i);
    if (lt < 0) break;
    if (lt > i && h.text) h.text(xml.slice(i, lt), false);
    const c = xml.charCodeAt(lt + 1);
    if (c === 47) {
      // </name>
      const gt = xml.indexOf('>', lt + 2);
      if (gt < 0) break;
      if (h.close) h.close(xml.slice(lt + 2, gt).trimEnd());
      i = gt + 1;
    } else if (c === 63) {
      // <? ... ?>
      const e = xml.indexOf('?>', lt + 2);
      if (e < 0) break;
      i = e + 2;
    } else if (c === 33) {
      if (xml.startsWith('<!--', lt)) {
        const e = xml.indexOf('-->', lt + 4);
        if (e < 0) break;
        i = e + 3;
      } else if (xml.startsWith('<![CDATA[', lt)) {
        const e = xml.indexOf(']]>', lt + 9);
        if (e < 0) break;
        if (h.text) h.text(xml.slice(lt + 9, e), true);
        i = e + 3;
      } else {
        const e = xml.indexOf('>', lt + 2);
        if (e < 0) break;
        i = e + 1;
      }
    } else {
      let j = lt + 1;
      let quote = 0;
      for (; j < n; j++) {
        const ch = xml.charCodeAt(j);
        if (quote) {
          if (ch === quote) quote = 0;
        } else if (ch === 34 || ch === 39) quote = ch;
        else if (ch === 62) break;
      }
      if (j >= n) break;
      const selfClose = xml.charCodeAt(j - 1) === 47;
      const end = selfClose ? j - 1 : j;
      let sp = lt + 1;
      while (sp < end && xml.charCodeAt(sp) > 32) sp++;
      const name = xml.slice(lt + 1, sp);
      if (h.open) h.open(name, sp < end ? xml.slice(sp, end) : '');
      if (selfClose && h.close) h.close(name);
      i = j + 1;
    }
  }
}

/** Parse a .rels part into Map<Id, {type, target, external}>; targets are resolved against baseDir. */
export function parseRels(xml, baseDir) {
  const rels = new Map();
  if (!xml) return rels;
  parseXml(xml, {
    open(name, attrs) {
      if (name !== 'Relationship') return;
      const id = attr(attrs, 'Id');
      const target = attr(attrs, 'Target') ?? '';
      const external = attr(attrs, 'TargetMode') === 'External';
      rels.set(id, { type: attr(attrs, 'Type') ?? '', target: external ? target : resolvePath(baseDir, target), external });
    },
  });
  return rels;
}

/** Resolve a relationship target ("../media/a.png", "/xl/x.xml") against a directory inside the package. */
export function resolvePath(baseDir, target) {
  const parts = target.startsWith('/') ? [] : baseDir.split('/').filter(Boolean);
  for (const seg of target.split('/')) {
    if (!seg || seg === '.') continue;
    if (seg === '..') parts.pop();
    else parts.push(seg);
  }
  return parts.join('/');
}

export const dirOf = (path) => (path.includes('/') ? path.slice(0, path.lastIndexOf('/')) : '');
export const relsPathOf = (path) => {
  const d = dirOf(path);
  return (d ? d + '/' : '') + '_rels/' + path.slice(path.lastIndexOf('/') + 1) + '.rels';
};
