// Baseline / regression screenshots for every screen of every part.
// usage: node shoot.mjs <repoRoot> <outDir> [parts=01,02,...]
// Serves repoRoot over HTTP under a /sub/ subpath, drives headless Edge via CDP,
// calls goTo(n) for each screen with animations + media frozen, saves PNGs,
// and writes report.json with 404s, console errors and exceptions per part.
import { spawn } from 'node:child_process';
import { createServer } from 'node:http';
import { readFile, mkdir, writeFile, stat } from 'node:fs/promises';
import { join, extname, normalize } from 'node:path';
import { tmpdir } from 'node:os';

const ROOT = process.argv[2], OUT = process.argv[3];
const PARTS = (process.argv[4] || '01,02,03,04,05,06').split(',');
await mkdir(OUT, { recursive: true });
const MIME = { '.html':'text/html; charset=utf-8', '.js':'text/javascript', '.css':'text/css', '.png':'image/png',
  '.jpg':'image/jpeg', '.jpeg':'image/jpeg', '.svg':'image/svg+xml', '.ttf':'font/ttf', '.woff2':'font/woff2',
  '.mp4':'video/mp4', '.mp3':'audio/mpeg', '.json':'application/json' };
const notFound = [];
const srv = createServer(async (req, res) => {
  let p = decodeURIComponent(new URL(req.url, 'http://x').pathname);
  if (!p.startsWith('/sub/')) { res.writeHead(404); res.end(); notFound.push(p); return; }
  const f = normalize(join(ROOT, p.slice(5)));
  try {
    const s = await stat(f); if (!s.isFile()) throw 0;
    res.writeHead(200, { 'content-type': MIME[extname(f).toLowerCase()] || 'application/octet-stream', 'cache-control':'no-store' });
    res.end(await readFile(f));
  } catch { res.writeHead(404); res.end(); notFound.push(p); }
});
await new Promise(r => srv.listen(0, '127.0.0.1', r));
const HP = srv.address().port;

const PORT = 9000 + Math.floor(Math.random() * 900);
const proc = spawn('C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe', ['--headless=new', '--disable-gpu',
  '--hide-scrollbars', '--mute-audio', '--autoplay-policy=no-user-gesture-required', '--remote-debugging-port=' + PORT,
  '--user-data-dir=' + join(tmpdir(), 'ehp' + PORT), '--window-size=1280,720', 'about:blank'], { stdio: 'ignore' });
const sleep = ms => new Promise(r => setTimeout(r, ms));
let page = null;
for (let i = 0; i < 50 && !page; i++) { await sleep(300); try { page = (await (await fetch(`http://127.0.0.1:${PORT}/json/list`)).json()).find(t => t.type === 'page'); } catch {} }
const ws = new WebSocket(page.webSocketDebuggerUrl); await new Promise(r => ws.onopen = r);
let id = 0; const pend = new Map(); let evs = [];
ws.onmessage = m => { const x = JSON.parse(m.data); if (x.id && pend.has(x.id)) { const p = pend.get(x.id); pend.delete(x.id); x.error ? p.rej(new Error(JSON.stringify(x.error))) : p.res(x.result); } else if (x.method) evs.push(x); };
const send = (m, p = {}) => new Promise((res, rej) => { const i = ++id; pend.set(i, { res, rej }); ws.send(JSON.stringify({ id: i, method: m, params: p })); setTimeout(() => { if (pend.has(i)) { pend.delete(i); rej(new Error('timeout ' + m)); } }, 30000); });
const ev = e => send('Runtime.evaluate', { expression: e, returnByValue: true, awaitPromise: true }).then(r => { if (r.exceptionDetails) throw new Error(r.exceptionDetails.exception?.description || 'eval'); return r.result.value; });
await send('Runtime.enable'); await send('Log.enable'); await send('Page.enable'); await send('Network.enable');
await send('Network.setCacheDisabled', { cacheDisabled: true });
await send('Emulation.setDeviceMetricsOverride', { width: 1280, height: 720, deviceScaleFactor: 1, mobile: false });

const FREEZE = `(async()=>{
 const docs=[document];
 for(let t=0;t<40;t++){const fr=[...document.querySelectorAll('.screen.active iframe')];
   if(fr.every(f=>{try{return f.contentDocument&&f.contentDocument.readyState==='complete';}catch(e){return true;}}))break;
   await new Promise(r=>setTimeout(r,150));}
 document.querySelectorAll('iframe').forEach(f=>{try{if(f.contentDocument)docs.push(f.contentDocument);}catch(e){}});
 for(const d of docs){if(!d.head)continue;let s=d.getElementById('__freeze');if(!s){s=d.createElement('style');s.id='__freeze';
  s.textContent='*,*::before,*::after{animation:none!important;transition:none!important;caret-color:transparent!important}';d.head.appendChild(s);}
  d.querySelectorAll('video,audio').forEach(v=>{try{v.pause();v.currentTime=0;}catch(e){}});
  try{d.getAnimations&&d.getAnimations().forEach(a=>{try{a.finish();}catch(e){a.cancel();}});}catch(e){}}
 await new Promise(r=>requestAnimationFrame(()=>requestAnimationFrame(r)));return true;})()`;

const report = {};
for (const part of PARTS) {
  const url = `http://127.0.0.1:${HP}/sub/methodica-math-ratio-03-${part}/index.html`;
  evs = []; const nf0 = notFound.length;
  await send('Page.navigate', { url }); await sleep(2500);
  const total = await ev(`document.querySelectorAll('.screen').length`);
  const shots = [];
  for (let n = 0; n < total; n++) {
    await ev(`(()=>{try{goTo(${n});}catch(e){console.error('goTo threw',e.message)}})()`);
    await sleep(900); await ev(FREEZE); await sleep(250);
    const s = await send('Page.captureScreenshot', { format: 'png' });
    const f = `${part}_s${String(n).padStart(2, '0')}.png`;
    await writeFile(join(OUT, f), Buffer.from(s.data, 'base64')); shots.push(f);
  }
  const host = new RegExp('http://127\\.0\\.0\\.1:' + HP, 'g');
  const fav = s => /favicon\.ico/.test(s);
  const net404 = evs.filter(e => e.method === 'Network.responseReceived' && e.params.response.status >= 400).map(e => e.params.response.url.replace(host, '')).filter(u => !fav(u));
  const errors = evs.filter(e => (e.method === 'Runtime.consoleAPICalled' && e.params.type === 'error') || (e.method === 'Log.entryAdded' && e.params.entry.level === 'error') || e.method === 'Runtime.exceptionThrown')
    .map(e => e.method === 'Runtime.exceptionThrown' ? 'EXC ' + (e.params.exceptionDetails.exception?.description || e.params.exceptionDetails.text) : e.method === 'Log.entryAdded' ? 'LOG ' + e.params.entry.text + ' ' + (e.params.entry.url || '') : 'CON ' + e.params.args.map(a => a.value ?? a.description).join(' '))
    .map(s => s.replace(host, '')).filter(s => !fav(s));
  const fonts = await ev(`document.fonts.ready.then(()=>[...document.fonts].filter(f=>f.status==='loaded').map(f=>f.family+' '+f.weight).sort())`);
  const imgs = await ev(`(async()=>{const docs=[document];document.querySelectorAll('iframe').forEach(f=>{try{if(f.contentDocument)docs.push(f.contentDocument);}catch(e){}});
    const all=[];for(const d of docs)d.querySelectorAll('img[src]').forEach(i=>all.push(i));
    await Promise.all(all.map(i=>i.complete?0:new Promise(r=>{i.onload=i.onerror=r;setTimeout(r,3000);})));
    return {n:all.length,broken:all.filter(i=>!i.naturalWidth).map(i=>i.getAttribute('src').slice(0,120))};})()`);
  report[part] = { total, shots: shots.length, net404: [...new Set(net404)], serverNotFound: notFound.slice(nf0), errors, fonts, imgs: imgs.n, brokenImgs: imgs.broken };
  console.log(part, 'screens', total, '404', report[part].net404.length, 'errors', errors.length, 'imgs', imgs.n, 'broken', imgs.broken.length, 'fonts', fonts.join('|'));
}
await writeFile(join(OUT, 'report.json'), JSON.stringify(report, null, 2));
ws.close(); proc.kill(); srv.close();
