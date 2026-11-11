// node tools/extract-bench.mjs [files...]   -> timings + structure summary; JSON/TXT dumps in out/extract/
import { mkdirSync, readFileSync, readdirSync, writeFileSync } from 'node:fs';
import { basename, dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { extract } from '../src/extract/index.mjs';

const root = join(dirname(fileURLToPath(import.meta.url)), '..');
const outDir = join(root, 'out', 'extract');
mkdirSync(outDir, { recursive: true });
const files = process.argv.length > 2
  ? process.argv.slice(2)
  : readdirSync(join(root, 'samples')).filter((f) => /\.(docx|xlsx|pptx)$/i.test(f)).map((f) => join(root, 'samples', f));

const RUNS = 15;
for (const file of files) {
  const bytes = new Uint8Array(readFileSync(file));
  let result;
  const times = [];
  for (let i = 0; i < RUNS; i++) {
    const t0 = performance.now();
    result = await extract(bytes);
    times.push(performance.now() - t0);
  }
  times.sort((a, b) => a - b);
  const name = basename(file);
  writeFileSync(join(outDir, name + '.json'), JSON.stringify(result, null, 1));
  writeFileSync(join(outDir, name + '.txt'), result.text);
  console.log(`${name}: ${bytes.length} B, type=${result.type}, text=${result.text.length} chars, first=${times.length ? times.at(-1).toFixed(1) : '-'} ms (max), median=${times[RUNS >> 1].toFixed(2)} ms, min=${times[0].toFixed(2)} ms`);
  console.log('  stats=' + JSON.stringify(result.stats) + ' meta=' + JSON.stringify(result.meta));
}
