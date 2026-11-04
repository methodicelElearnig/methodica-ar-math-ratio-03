// Compare two screenshot dirs byte-for-byte (PNG encoder is deterministic for identical pixels).
// usage: node cmp.mjs <baseDir> <newDir>
import { readdir, readFile } from 'node:fs/promises';
import { join } from 'node:path';
const [A, B, ONLY] = process.argv.slice(2);
const PARTS = ONLY ? ONLY.split(',') : null;
const files = (await readdir(A)).filter(f => f.endsWith('.png') && (!PARTS || PARTS.includes(f.slice(0, 2))));
let diff = [], missing = [];
for (const f of files) {
  let b; try { b = await readFile(join(B, f)); } catch { missing.push(f); continue; }
  const a = await readFile(join(A, f));
  if (!a.equals(b)) diff.push(f);
}
const ra = JSON.parse(await readFile(join(A, 'report.json'), 'utf8')), rb = JSON.parse(await readFile(join(B, 'report.json'), 'utf8'));
for (const p of Object.keys(ra).filter(p => !PARTS || PARTS.includes(p))) {
  const x = ra[p], y = rb[p] || {};
  if (JSON.stringify(x.net404) !== JSON.stringify(y.net404)) console.log(p, '404 changed:', y.net404);
  if (JSON.stringify(x.errors) !== JSON.stringify(y.errors)) console.log(p, 'errors changed:', y.errors);
  if (JSON.stringify(x.fonts) !== JSON.stringify(y.fonts)) console.log(p, 'fonts changed:', x.fonts, '->', y.fonts);
  if (y.brokenImgs && y.brokenImgs.length) console.log(p, 'BROKEN IMGS:', y.brokenImgs);
  if (x.imgs !== undefined && y.imgs !== x.imgs) console.log(p, 'img count', x.imgs, '->', y.imgs);
}
console.log(`compared ${files.length}, differ ${diff.length}, missing ${missing.length}`);
if (diff.length) console.log('DIFF:', diff.join(' '));
if (missing.length) console.log('MISSING:', missing.join(' '));
process.exit(diff.length || missing.length ? 1 : 0);
