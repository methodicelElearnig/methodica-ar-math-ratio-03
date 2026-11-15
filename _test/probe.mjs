// Scripted behavioural probe over HTTP (/sub/ subpath) in headless Edge.
// usage: node probe.mjs <repoRoot> <steps.json>
// steps: [{nav:"methodica-ar-math-ratio-03-02/index.html?dev=1"}, {wait:800},
//         {eval:"expr", expect: <json value>, label:"..."}, {eval:"expr", log:true}]
// Prints PASS/FAIL per expectation; exit 1 on any FAIL. Console errors are listed at the end.
import { spawn } from 'node:child_process';
import { createServer } from 'node:http';
import { readFile, stat } from 'node:fs/promises';
import { readFileSync } from 'node:fs';
import { join, extname, normalize } from 'node:path';
import { tmpdir } from 'node:os';

const ROOT = process.argv[2];
const STEPS = JSON.parse(readFileSync(process.argv[3], 'utf8').replace(/^\uFEFF/, ''));
const MIME = { '.html':'text/html; charset=utf-8', '.js':'text/javascript', '.css':'text/css', '.png':'image/png', '.jpg':'image/jpeg',
  '.jpeg':'image/jpeg', '.svg':'image/svg+xml', '.ttf':'font/ttf', '.mp4':'video/mp4', '.mp3':'audio/mpeg', '.json':'application/json' };
const hits = [];
const srv = createServer(async (req, res) => {
  const u = new URL(req.url, 'http://x'); const p = decodeURIComponent(u.pathname);
  if (req.method !== 'GET') { let b = ''; req.on('data', c => b += c); req.on('end', () => { hits.push({ m: req.method, p, b }); res.writeHead(204); res.end(); }); return; }
  hits.push({ m: 'GET', p: p + u.search });
  if (!p.startsWith('/sub/')) { res.writeHead(404); res.end(); return; }
  const f = normalize(join(ROOT, p.slice(5)));
  try { const s = await stat(f); if (!s.isFile()) throw 0; res.writeHead(200, { 'content-type': MIME[extname(f).toLowerCase()] || 'application/octet-stream', 'cache-control': 'no-store' }); res.end(await readFile(f)); }
  catch { res.writeHead(404); res.end(); }
});
await new Promise(r => srv.listen(0, '127.0.0.1', r));
const BASE = `http://127.0.0.1:${srv.address().port}/sub/`;
const PORT = 9000 + Math.floor(Math.random() * 900);
const proc = spawn('C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe', ['--headless=new', '--disable-gpu', '--mute-audio',
  '--autoplay-policy=no-user-gesture-required', '--remote-debugging-port=' + PORT, '--user-data-dir=' + join(tmpdir(), 'ehp' + PORT),
  '--window-size=1280,720', 'about:blank'], { stdio: 'ignore' });
const sleep = ms => new Promise(r => setTimeout(r, ms));
let page = null;
for (let i = 0; i < 50 && !page; i++) { await sleep(300); try { page = (await (await fetch(`http://127.0.0.1:${PORT}/json/list`)).json()).find(t => t.type === 'page'); } catch {} }
const ws = new WebSocket(page.webSocketDebuggerUrl); await new Promise(r => ws.onopen = r);
let id = 0; const pend = new Map(); const evs = [];
ws.onmessage = m => { const x = JSON.parse(m.data); if (x.id && pend.has(x.id)) { const p = pend.get(x.id); pend.delete(x.id); x.error ? p.rej(new Error(JSON.stringify(x.error))) : p.res(x.result); } else if (x.method) evs.push(x); };
const send = (m, p = {}) => new Promise((res, rej) => { const i = ++id; pend.set(i, { res, rej }); ws.send(JSON.stringify({ id: i, method: m, params: p })); setTimeout(() => { if (pend.has(i)) { pend.delete(i); rej(new Error('timeout ' + m)); } }, 30000); });
const ev = e => send('Runtime.evaluate', { expression: e, returnByValue: true, awaitPromise: true }).then(r => r.exceptionDetails ? { __err: r.exceptionDetails.exception?.description || r.exceptionDetails.text } : r.result.value);
await send('Runtime.enable'); await send('Log.enable'); await send('Page.enable'); await send('Network.setCacheDisabled', { cacheDisabled: true });
await send('Emulation.setDeviceMetricsOverride', { width: 1280, height: 720, deviceScaleFactor: 1, mobile: false });
globalThis.__hits = hits;
let fails = 0;
for (const s of STEPS) {
  if (s.nav) { await send('Page.navigate', { url: BASE + s.nav }); await sleep(s.after ?? 2000); }
  else if (s.wait) await sleep(s.wait);
  else if (s.shot) { const r = await send('Page.captureScreenshot', { format: 'png' }); (await import('node:fs')).writeFileSync(s.shot, Buffer.from(r.data, 'base64')); console.log('SHOT', s.shot); }
  else if (s.hits) { const v = hits.filter(h => new RegExp(s.hits).test(h.m + ' ' + h.p)); console.log('HITS', s.hits, v.length, s.show ? JSON.stringify(v.slice(-s.show)) : ''); if ('expect' in s && v.length !== s.expect) { fails++; console.log('  FAIL expected', s.expect); } }
  else if (s.eval) {
    const v = await ev(s.eval);
    if ('expect' in s) { const ok = JSON.stringify(v) === JSON.stringify(s.expect); if (!ok) fails++; console.log(ok ? 'PASS' : 'FAIL', s.label || s.eval.slice(0, 80), ok ? '' : '=> ' + JSON.stringify(v) + ' expected ' + JSON.stringify(s.expect)); }
    else console.log('LOG ', s.label || s.eval.slice(0, 80), JSON.stringify(v));
    if (s.after) await sleep(s.after);
  }
}
const errs = evs.filter(e => (e.method === 'Runtime.consoleAPICalled' && e.params.type === 'error') || e.method === 'Runtime.exceptionThrown' || (e.method === 'Log.entryAdded' && e.params.entry.level === 'error' && !/favicon/.test(e.params.entry.url || '')))
  .map(e => e.method === 'Runtime.exceptionThrown' ? 'EXC ' + (e.params.exceptionDetails.exception?.description || e.params.exceptionDetails.text).split('\n')[0] : e.method === 'Log.entryAdded' ? 'LOG ' + e.params.entry.text + ' ' + e.params.entry.url : 'CON ' + e.params.args.map(a => a.value ?? a.description).join(' ').slice(0, 300));
console.log('console errors:', errs.length); errs.slice(0, 20).forEach(e => console.log('  ', e));
console.log(fails ? `FAILED ${fails}` : 'ALL PASS');
ws.close(); proc.kill(); srv.close(); process.exit(fails ? 1 : 0);
