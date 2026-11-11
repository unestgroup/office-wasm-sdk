// Minimal read-only ZIP reader (store + deflate, ZIP64 sizes/offsets). No dependencies.
// Works in Node >= 20 and in browsers / workers.

const SIG_EOCD = 0x06054b50;
const SIG_EOCD64 = 0x06064b50;
const SIG_EOCD64_LOC = 0x07064b50;
const SIG_CEN = 0x02014b50;
const SIG_LOC = 0x04034b50;

let nodeZlib = null;
try {
  nodeZlib = globalThis.process?.getBuiltinModule?.('node:zlib') ?? null;
} catch {
  nodeZlib = null;
}

async function inflateRaw(data, maxBytes) {
  if (nodeZlib) return nodeZlib.inflateRawSync(data, { maxOutputLength: maxBytes });
  const stream = new Blob([data]).stream().pipeThrough(new DecompressionStream('deflate-raw'));
  const reader = stream.getReader();
  const chunks = [];
  let total = 0;
  for (;;) {
    const { done, value } = await reader.read();
    if (done) break;
    total += value.length;
    if (total > maxBytes) {
      await reader.cancel();
      throw new Error('zip entry exceeds maxEntryBytes');
    }
    chunks.push(value);
  }
  const out = new Uint8Array(total);
  let off = 0;
  for (const c of chunks) {
    out.set(c, off);
    off += c.length;
  }
  return out;
}

/**
 * @param {Uint8Array|ArrayBuffer} input
 * @param {{maxEntryBytes?: number}} [opts]
 */
export function openZip(input, opts = {}) {
  const u8 = input instanceof Uint8Array ? input : new Uint8Array(input);
  const maxEntryBytes = opts.maxEntryBytes ?? 512 * 1024 * 1024;
  const dv = new DataView(u8.buffer, u8.byteOffset, u8.byteLength);
  const utf8 = new TextDecoder('utf-8');

  let eocd = -1;
  for (let i = u8.length - 22, min = Math.max(0, u8.length - 22 - 0xffff); i >= min; i--) {
    if (dv.getUint32(i, true) === SIG_EOCD) {
      eocd = i;
      break;
    }
  }
  if (eocd < 0) throw new Error('not a zip file (end of central directory not found)');

  let count = dv.getUint16(eocd + 10, true);
  let cdOff = dv.getUint32(eocd + 16, true);
  if ((count === 0xffff || cdOff === 0xffffffff) && eocd >= 20 && dv.getUint32(eocd - 20, true) === SIG_EOCD64_LOC) {
    const e64 = Number(dv.getBigUint64(eocd - 20 + 8, true));
    if (dv.getUint32(e64, true) !== SIG_EOCD64) throw new Error('bad zip64 end of central directory');
    count = Number(dv.getBigUint64(e64 + 32, true));
    cdOff = Number(dv.getBigUint64(e64 + 48, true));
  }

  /** @type {Map<string, {name: string, method: number, csize: number, usize: number, off: number}>} */
  const entries = new Map();
  let p = cdOff;
  for (let i = 0; i < count; i++) {
    if (p + 46 > u8.length || dv.getUint32(p, true) !== SIG_CEN) throw new Error('bad zip central directory');
    const method = dv.getUint16(p + 10, true);
    let csize = dv.getUint32(p + 20, true);
    let usize = dv.getUint32(p + 24, true);
    const nameLen = dv.getUint16(p + 28, true);
    const extraLen = dv.getUint16(p + 30, true);
    const commentLen = dv.getUint16(p + 32, true);
    let off = dv.getUint32(p + 42, true);
    const name = utf8.decode(u8.subarray(p + 46, p + 46 + nameLen));
    if (usize === 0xffffffff || csize === 0xffffffff || off === 0xffffffff) {
      let x = p + 46 + nameLen;
      const xEnd = x + extraLen;
      while (x + 4 <= xEnd) {
        const id = dv.getUint16(x, true);
        const len = dv.getUint16(x + 2, true);
        if (id === 1) {
          let q = x + 4;
          if (usize === 0xffffffff) { usize = Number(dv.getBigUint64(q, true)); q += 8; }
          if (csize === 0xffffffff) { csize = Number(dv.getBigUint64(q, true)); q += 8; }
          if (off === 0xffffffff) { off = Number(dv.getBigUint64(q, true)); }
          break;
        }
        x += 4 + len;
      }
    }
    entries.set(name, { name, method, csize, usize, off });
    p += 46 + nameLen + extraLen + commentLen;
  }

  async function bytes(name) {
    const e = entries.get(name);
    if (!e) return undefined;
    if (e.usize > maxEntryBytes) throw new Error('zip entry exceeds maxEntryBytes: ' + name);
    if (e.off + 30 > u8.length || dv.getUint32(e.off, true) !== SIG_LOC) throw new Error('bad zip local header: ' + name);
    const start = e.off + 30 + dv.getUint16(e.off + 26, true) + dv.getUint16(e.off + 28, true);
    const raw = u8.subarray(start, start + e.csize);
    if (e.method === 0) return raw;
    if (e.method === 8) return inflateRaw(raw, maxEntryBytes);
    throw new Error('unsupported zip compression method ' + e.method + ': ' + name);
  }

  return {
    names: () => [...entries.keys()],
    has: (name) => entries.has(name),
    size: (name) => entries.get(name)?.usize,
    bytes,
    /** @returns {Promise<string|undefined>} */
    async text(name) {
      const b = await bytes(name);
      return b === undefined ? undefined : utf8.decode(b);
    },
  };
}
