// Statement-flow regression: walks scripted learner paths against the local library stub and
// compares the statement stream with a reviewed golden file.
//   node _test/statement-flow.mjs [repoRoot] [--update] [scenario-name …]
// Scenarios: _test/scenarios/flow-*.txt — first line `# part NN [query=&…] <description>`, then one JS step per line
// (async allowed). Golden: the same name with .expected.txt. A run fails on any difference, on a
// duplicate (__dupes), on a console error / exception, or when the preflight is off (stub not
// loaded, METADATA.id ≠ XAPI_COMP_ID). Same browser setup as shoot.mjs (headless Edge over CDP,
// the repo served under /sub/).
import { spawn } from 'node:child_process';
import { createServer } from 'node:http';
import { readFile, writeFile, readdir, stat } from 'node:fs/promises';
import { join, extname, normalize, dirname } from 'node:path';
import { tmpdir } from 'node:os';
import { fileURLToPath } from 'node:url';

const args = process.argv.slice(2);
const UPDATE = args.includes('--update');
const rest = args.filter(a => a !== '--update');
const ROOT = rest[0] && !rest[0].startsWith('flow-') ? rest.shift() : join(dirname(fileURLToPath(import.meta.url)), '..');
const SCN = join(ROOT, '_test', 'scenarios');
const only = rest;
const names = (await readdir(SCN)).filter(f => /^flow-.*\.txt$/.test(f) && !f.endsWith('.expected.txt'))
  .map(f => f.replace(/\.txt$/, '')).filter(n => !only.length || only.includes(n)).sort();

const MIME = { '.html': 'text/html; charset=utf-8', '.js': 'text/javascript', '.css': 'text/css', '.png': 'image/png',
  '.jpg': 'image/jpeg', '.jpeg': 'image/jpeg', '.svg': 'image/svg+xml', '.ttf': 'font/ttf', '.mp4': 'video/mp4', '.json': 'application/json' };
const srv = createServer(async (req, res) => {
  const p = decodeURIComponent(new URL(req.url, 'http://x').pathname);
  if (!p.startsWith('/sub/')) { res.writeHead(404); res.end(); return; }
  const f = normalize(join(ROOT, p.slice(5)));
  try {
    const s = await stat(f); if (!s.isFile()) throw 0;
    res.writeHead(200, { 'content-type': MIME[extname(f).toLowerCase()] || 'application/octet-stream', 'cache-control': 'no-store' });
    res.end(await readFile(f));
  } catch { res.writeHead(404); res.end(); }
});
await new Promise(r => srv.listen(0, '127.0.0.1', r));
const HP = srv.address().port;
const PORT = 9000 + Math.floor(Math.random() * 900);
const proc = spawn('C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe', ['--headless=new', '--disable-gpu',
  '--mute-audio', '--autoplay-policy=no-user-gesture-required', '--remote-debugging-port=' + PORT,
  '--user-data-dir=' + join(tmpdir(), 'esf' + PORT), '--window-size=1280,720', 'about:blank'], { stdio: 'ignore' });
const sleep = ms => new Promise(r => setTimeout(r, ms));
let page = null;
for (let i = 0; i < 50 && !page; i++) { await sleep(300); try { page = (await (await fetch(`http://127.0.0.1:${PORT}/json/list`)).json()).find(t => t.type === 'page'); } catch {} }
const ws = new WebSocket(page.webSocketDebuggerUrl); await new Promise(r => ws.onopen = r);
let id = 0; const pend = new Map(); let evs = [];
ws.onmessage = m => { const x = JSON.parse(m.data); if (x.id && pend.has(x.id)) { const p = pend.get(x.id); pend.delete(x.id); x.error ? p.rej(new Error(JSON.stringify(x.error))) : p.res(x.result); } else if (x.method) evs.push(x); };
const send = (m, p = {}) => new Promise((res, rej) => { const i = ++id; pend.set(i, { res, rej }); ws.send(JSON.stringify({ id: i, method: m, params: p })); setTimeout(() => { if (pend.has(i)) { pend.delete(i); rej(new Error('timeout ' + m)); } }, 60000); });
const ev = e => send('Runtime.evaluate', { expression: e, returnByValue: true, awaitPromise: true }).then(r => { if (r.exceptionDetails) throw new Error(r.exceptionDetails.exception?.description || 'eval'); return r.result.value; });
await send('Runtime.enable'); await send('Page.enable'); await send('Network.setCacheDisabled', { cacheDisabled: true });

/* One line per statement: verb, object type, the object (COMP, item NNN or NNN/qN), and the result
   (success, scaled score rounded to 3 places, result.response — the string MOE compares; before
   2026-11-11 the line showed student_answer, the only answer this unit sent). */
const short = u => !u ? '-' : String(u).replace(/\/+$/, '').split('/').slice(-2).join('/')
  .replace(/^methodica-math-ratio-03-0\d$/, 'COMP').replace(/^methodica-math-ratio-03-0\d\/methodica-math-ratio-03-0\d-(\d{3})$/, '$1')
  .replace(/^methodica-math-ratio-03-0\d-(\d{3})\/(q\d+)$/, '$1/$2');
const fmt = s => {
  const o = s.opts || {}, r = s.result;
  const res = r ? ` ${r.success}` + (r.score ? ` ${Math.round(r.score.scaled * 1000) / 1000}` : '') +
    (r.response != null ? ` «${r.response}»` : '') : '';
  return `${s.verb} ${s.objectType} ${short(o.objectId || o.questionId)}${res}`;
};

/* MOE runbook §3 (2026-11-11): a right answer's result.response must equal one of the question's
   correctXapiResponsesPattern, character by character. The patterns are Kata's, carried in the
   shipped metadata; a question without one is skipped here (it is listed for Kata instead). */
const PATTERN_CHECK = `(() => {
  const strip = p => String(p).replace(/^(\\{[a-z_]+=[^}]*\\})+/, '');
  const pats = {};
  ((window.METADATA && window.METADATA.subContent) || []).forEach(sc => (sc.questions || []).forEach(q => {
    pats[String(q.questionId).replace(/\\/+$/, '')] = (q.correctXapiResponsesPattern || []).map(strip);
  }));
  return __stmts().filter(s => /^answered/.test(s.verb) && s.result && s.result.success === true).map(s => {
    const qid = String((s.opts || {}).questionId || '').replace(/\\/+$/, ''), want = pats[qid];
    if (!want) return 'no metadata question ' + qid;
    if (!want.length) return null;
    return want.includes(String(s.result.response)) ? null
      : qid.split('/').slice(-2).join('/') + ': response «' + s.result.response + '» is not Kata pattern «' + want.join(' || ') + '»';
  }).filter(Boolean);
})()`;

let failed = 0;
for (const name of names) {
  const lines = (await readFile(join(SCN, name + '.txt'), 'utf8')).split(/\r?\n/);
  const [, part, query = ''] = lines[0].match(/^# part (\d\d)(?:\s+query=(\S+))?/) || [];
  const steps = lines.slice(1).filter(l => l.trim() && !l.startsWith('#'));
  const base = `http://127.0.0.1:${HP}/sub/methodica-math-ratio-03-${part}/index.html?xapiLib=../_test/xapi-720-${process.env.XAPI_STUB || 'l'}.js${query}`;   /* -l = production (2026-11-14); XAPI_STUB=k runs the rollback library */
  /* Reset from a same-origin page that is NOT the lomda: with resume on, a lomda page writes its
     restored document back on pagehide, which undid a reset taken on the lomda itself and leaked
     one scenario's state and ledger into the next (README → Notes). */
  await send('Page.navigate', { url: `http://127.0.0.1:${HP}/sub/metadata/methodica-math-ratio-03_unit.json` }); await sleep(500);
  await ev('sessionStorage.clear(); localStorage.clear(); true');
  evs = [];
  await send('Page.navigate', { url: base }); await sleep(3000);
  const problems = [];
  const pre = await ev(`({using: window.XAPI_USING_G === true, meta: !!window.METADATA && window.METADATA.id === XAPI_COMP_ID})`);
  if (!pre.using || !pre.meta) problems.push('preflight ' + JSON.stringify(pre));
  for (const st of steps) { try { await ev(`(async()=>{${st}})()`); } catch (e) { problems.push('step failed: ' + st + ' — ' + e.message.split('\n')[0]); } await sleep(150); }
  await sleep(500);
  const got = (await ev('typeof __stmts === "function" ? __stmts() : []')).map(fmt);
  (await ev(PATTERN_CHECK)).forEach(p => problems.push(p));
  const dupes = await ev('typeof __dupes === "function" ? __dupes() : []');
  if (dupes.length) problems.push('duplicates ' + JSON.stringify(dupes));
  const errs = evs.filter(e => e.method === 'Runtime.exceptionThrown' || (e.method === 'Runtime.consoleAPICalled' && e.params.type === 'error'));
  if (errs.length) problems.push(errs.length + ' console error(s)');
  const gold = join(SCN, name + '.expected.txt');
  if (UPDATE) await writeFile(gold, got.join('\n') + '\n');
  else {
    let want = []; try { want = (await readFile(gold, 'utf8')).split(/\r?\n/).filter(Boolean); } catch { problems.push('no golden file'); }
    const n = Math.max(want.length, got.length);
    for (let i = 0; i < n; i++) if (want[i] !== got[i]) { problems.push(`line ${i + 1}: want «${want[i] ?? '—'}» got «${got[i] ?? '—'}»`); if (problems.length > 8) break; }
  }
  console.log((problems.length ? 'FAIL ' : 'ok   ') + name + ` (${got.length} statements)` + (UPDATE ? ' [golden written]' : ''));
  problems.forEach(p => console.log('     ' + p));
  if (problems.length) failed++;
}
console.log(failed ? `${failed} scenario(s) FAILED` : `all ${names.length} scenarios pass`);
ws.close(); /* proc.kill() ends only the launcher on Windows; the browser tree would linger (hundreds of
   headless msedge.exe after a day of runs). taskkill /T takes the whole tree. */
try { (await import('node:child_process')).execFileSync('taskkill', ['/T', '/F', '/PID', String(proc.pid)], { stdio: 'ignore' }); } catch { proc.kill(); } srv.close();
process.exit(failed ? 1 : 0);
