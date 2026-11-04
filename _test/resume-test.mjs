// Resume oracle: a screen restored after a RELOAD must look exactly like the same screen
// re-entered IN-SESSION, and the restore must send no answered/completed again.
// usage: node resume-test.mjs <repoRoot> <scenarios.json> <outDir>
// scenario: { part:'02', name:'s1-wrong1', screen:1, away:0, setup:"js run on a fresh state" }
import { spawn } from 'node:child_process';
import { createServer } from 'node:http';
import { readFile, stat, mkdir, writeFile } from 'node:fs/promises';
import { readFileSync } from 'node:fs';
import { join, extname, normalize } from 'node:path';
import { tmpdir } from 'node:os';

import zlib from 'node:zlib';
/* Minimal PNG decoder (8-bit RGB/RGBA, non-interlaced — what Chrome's screenshots are). */
function decodePng(buf) {
  let p = 8, w, h, ct, idat = [];
  while (p < buf.length) {
    const len = buf.readUInt32BE(p), type = buf.toString('ascii', p + 4, p + 8), data = buf.subarray(p + 8, p + 8 + len);
    if (type === 'IHDR') { w = data.readUInt32BE(0); h = data.readUInt32BE(4); ct = data[9]; }
    else if (type === 'IDAT') idat.push(data);
    p += 12 + len;
  }
  const bpp = ct === 6 ? 4 : 3, stride = w * bpp, raw = zlib.inflateSync(Buffer.concat(idat)), out = Buffer.alloc(h * stride);
  for (let y = 0; y < h; y++) {
    const f = raw[y * (stride + 1)], src = y * (stride + 1) + 1, dst = y * stride;
    for (let x = 0; x < stride; x++) {
      const a = x >= bpp ? out[dst + x - bpp] : 0, b = y ? out[dst - stride + x] : 0, c = (x >= bpp && y) ? out[dst - stride + x - bpp] : 0;
      let v = raw[src + x];
      if (f === 1) v += a; else if (f === 2) v += b; else if (f === 3) v += (a + b) >> 1;
      else if (f === 4) { const pp = a + b - c, pa = Math.abs(pp - a), pb = Math.abs(pp - b), pc = Math.abs(pp - c); v += (pa <= pb && pa <= pc) ? a : (pb <= pc ? b : c); }
      out[dst + x] = v & 255;
    }
  }
  return { w, h, bpp, px: out };
}
/* Same picture, allowing anti-aliasing noise: no channel off by more than 64, and at most
   40 pixels off by more than 4 (re-rasterised rounded edges after a reload). */
function samePicture(A, B) {
  if (A.equals(B)) return { same: true, off: 0 };
  const a = decodePng(A), b = decodePng(B);
  if (a.w !== b.w || a.h !== b.h) return { same: false, off: -1 };
  let off = 0, max = 0;
  for (let i = 0; i < a.w * a.h; i++) {
    let d = 0; for (let k = 0; k < 3; k++) d = Math.max(d, Math.abs(a.px[i * a.bpp + k] - b.px[i * b.bpp + k]));
    if (d > 4) off++; if (d > max) max = d;
  }
  return { same: max <= 64 && off <= 40, off, max };
}

const [ROOT, SCEN, OUT] = process.argv.slice(2);
await mkdir(OUT, { recursive: true });
const SC = JSON.parse(readFileSync(SCEN, 'utf8').replace(/^\uFEFF/, ''));
const MIME = { '.html':'text/html; charset=utf-8', '.js':'text/javascript', '.css':'text/css', '.png':'image/png', '.jpg':'image/jpeg',
  '.jpeg':'image/jpeg', '.svg':'image/svg+xml', '.ttf':'font/ttf', '.mp4':'video/mp4', '.mp3':'audio/mpeg', '.json':'application/json' };
const srv = createServer(async (req, res) => {
  const p = decodeURIComponent(new URL(req.url, 'http://x').pathname);
  if (!p.startsWith('/sub/')) { res.writeHead(404); res.end(); return; }
  const f = normalize(join(ROOT, p.slice(5)));
  try { const s = await stat(f); if (!s.isFile()) throw 0; res.writeHead(200, { 'content-type': MIME[extname(f).toLowerCase()] || 'application/octet-stream', 'cache-control': 'no-store' }); res.end(await readFile(f)); }
  catch { res.writeHead(404); res.end(); }
});
await new Promise(r => srv.listen(0, '127.0.0.1', r));
const BASE = `http://127.0.0.1:${srv.address().port}/sub/`;
const PORT = 9000 + Math.floor(Math.random() * 900);
const proc = spawn('C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe', ['--headless=new', '--disable-gpu', '--mute-audio', '--hide-scrollbars',
  '--autoplay-policy=no-user-gesture-required', '--remote-debugging-port=' + PORT, '--user-data-dir=' + join(tmpdir(), 'ehp' + PORT),
  '--window-size=1280,720', 'about:blank'], { stdio: 'ignore' });
const sleep = ms => new Promise(r => setTimeout(r, ms));
let page = null;
for (let i = 0; i < 50 && !page; i++) { await sleep(300); try { page = (await (await fetch(`http://127.0.0.1:${PORT}/json/list`)).json()).find(t => t.type === 'page'); } catch {} }
const ws = new WebSocket(page.webSocketDebuggerUrl); await new Promise(r => ws.onopen = r);
let id = 0; const pend = new Map(); let evs = [];
ws.onmessage = m => { const x = JSON.parse(m.data); if (x.id && pend.has(x.id)) { const p = pend.get(x.id); pend.delete(x.id); x.error ? p.rej(new Error(JSON.stringify(x.error))) : p.res(x.result); } else if (x.method) evs.push(x); };
const send = (m, p = {}) => new Promise((res, rej) => { const i = ++id; pend.set(i, { res, rej }); ws.send(JSON.stringify({ id: i, method: m, params: p })); setTimeout(() => { if (pend.has(i)) { pend.delete(i); rej(new Error('timeout ' + m)); } }, 30000); });
const ev = async e => { const r = await send('Runtime.evaluate', { expression: e, returnByValue: true, awaitPromise: true }); if (r.exceptionDetails) throw new Error((r.exceptionDetails.exception?.description || r.exceptionDetails.text).split('\n')[0]); return r.result.value; };
await send('Runtime.enable'); await send('Page.enable'); await send('Log.enable'); await send('Network.setCacheDisabled', { cacheDisabled: true });
await send('Emulation.setDeviceMetricsOverride', { width: 1280, height: 720, deviceScaleFactor: 1, mobile: false });
const FREEZE = `(async()=>{const docs=[document];document.querySelectorAll('iframe').forEach(f=>{try{if(f.contentDocument)docs.push(f.contentDocument);}catch(e){}});
 for(const d of docs){if(!d.head)continue;const s=d.createElement('style');s.textContent='*,*::before,*::after{animation:none!important;transition:none!important;caret-color:transparent!important}';d.head.appendChild(s);
 d.querySelectorAll('video,audio').forEach(v=>{try{v.pause();v.currentTime=0;}catch(e){}});
 try{d.getAnimations().forEach(a=>{try{a.finish();}catch(e){a.cancel();}});}catch(e){}}
 if(document.activeElement&&document.activeElement.blur)document.activeElement.blur();
 await new Promise(r=>requestAnimationFrame(()=>requestAnimationFrame(r)));return true;})()`;
const shot = async f => { await ev(FREEZE); await sleep(200); const r = await send('Page.captureScreenshot', { format: 'png' }); const b = Buffer.from(r.data, 'base64'); await writeFile(join(OUT, f), b); return b; };
const COUNT = `(()=>{const L=__stmts();return {answered:L.filter(s=>s.verb.startsWith('answered')).length, completed:L.filter(s=>s.verb==='completed').length, requested:L.filter(s=>s.verb.startsWith('requested')).length, init:L.filter(s=>s.verb==='initialized').length};})()`;

let fails = 0;
for (const sc of SC) {
  const url = `${BASE}methodica-math-ratio-03-${sc.part}/index.html?xapiLib=../_test/xapi-720-k.js`;
  /* reset from a same-origin page that is NOT the lomda: a lomda page would flush its restored
     document back on pagehide */
  await send('Page.navigate', { url: BASE + 'metadata/methodica-math-ratio-03_unit.json' }); await sleep(600);
  await ev(`sessionStorage.clear(); localStorage.clear(); true`);
  await send('Page.navigate', { url }); await sleep(3000);
  evs = [];
  let A, B, before, after, scr;
  try {
    await ev(`(()=>{ ${sc.setup} ; return true })()`);
    await sleep(400);
    /* revisit: compare a screen OTHER than the landing one — restored stale, rebuilt on first entry */
    const look = sc.revisit ?? sc.screen;
    await ev(`goTo(${sc.away ?? 0}); goTo(${look}); true`);
    await sleep(900);
    A = await shot(`${sc.part}_${sc.name}_A.png`);
    if (sc.revisit != null) { await ev(`goTo(${sc.screen}); true`); await sleep(300); }
    before = await ev(COUNT);
    await sleep(1200);                                   // let the debounced save land
    await send('Page.reload', { ignoreCache: true }); await sleep(3500);
    scr = await ev('currentScreen');
    if (sc.revisit != null) { await ev(`goTo(${sc.revisit}); true`); await sleep(900); }
    B = await shot(`${sc.part}_${sc.name}_B.png`);
    after = await ev(COUNT);
    if (sc.after) { const v = await ev(sc.after); console.log('   after:', JSON.stringify(v)); }
  } catch (e) { fails++; console.log('ERROR', sc.part, sc.name, e.message); continue; }
  const errs = evs.filter(e => e.method === 'Runtime.exceptionThrown' || (e.method === 'Runtime.consoleAPICalled' && e.params.type === 'error'))
    .map(e => e.method === 'Runtime.exceptionThrown' ? (e.params.exceptionDetails.exception?.description || '').split('\n')[0] : e.params.args.map(a => a.value ?? a.description).join(' ').slice(0, 200));
  const cmp = samePicture(A, B), same = cmp.same;
  if (!A.equals(B)) console.log(`   pixels: ${cmp.off} off by >4, max delta ${cmp.max}`);
  const resent = after.answered !== before.answered || after.completed !== before.completed || after.requested !== before.requested;
  const ok = same && scr === sc.screen && !resent && !errs.length;
  if (!ok) fails++;
  console.log(`${ok ? 'PASS' : 'FAIL'} ${sc.part} ${sc.name}: screen ${scr}/${sc.screen} pixels ${same ? 'same' : 'DIFFER'} stmts ${JSON.stringify(before)} -> ${JSON.stringify(after)}${errs.length ? ' ERRORS ' + JSON.stringify(errs) : ''}`);
}
console.log(fails ? `FAILED ${fails}` : 'ALL PASS');
ws.close(); proc.kill(); srv.close(); process.exit(fails ? 1 : 0);
