// No write before the restore, a failed read closes the writes, and the leave save through xapi-720-l.js
// (Documentation/reporting-and-resume/BOOT-AND-RESUME-HARDENING-2026-11-03.md §2.1–2.3, §2.6, §2.10;
// KNOWN-ISSUE-drain-overwrites-payload.md). Fix R, 2026-11-14. The check list is the one of
// methodica-science-mass-density-01/_test/write-guard.js (jsdom); this unit's harness has no npm
// dependencies, so the same checks run here in headless Edge over the DevTools protocol, like probe.mjs.
//
//   a–j  behaviour (part 01, a document saved on screen 2): a character pick and a navigation before
//        phase B write nothing; the drain writes the document as read; the restore lands on 2; the
//        debounced document (held by reference) is still on 2; markSent before phase B marks in memory
//        only; a failed read (not a 404) closes every write until a read succeeds; a 404 is a fresh
//        document.
//   k–l  source shape of 40-resume.js / 50-loader.js.          s  the leave save's shape, letter l.
//   m–p  every part's markup: one cover net that raises __resumeInFlight itself (n: in the browser, with
//        every external script removed — only with a registration), the preloads equal the loader's
//        URLs, the Kata preconnect, every image off the landing screen lazy, no hidden autoplay clip with
//        an inline src; part 06: the radar clips get their src only when their screen opens.
//   t    every part: the leave flush goes through saveState720OnLeave and carries the new screen; on -k
//        (no such function — the rollback path) through saveState720; no page error.
//
//   node _test/write-guard.mjs [baseDir]
// Acceptance: passes on the repo; FAILS with the previous package as baseDir
// (../../deployments/2026-11-11) — the proof that it sees the bug.
import { spawn } from 'node:child_process';
import { createServer } from 'node:http';
import { readFile, stat } from 'node:fs/promises';
import { readFileSync, existsSync } from 'node:fs';
import { join, extname, normalize, dirname, resolve } from 'node:path';
import { tmpdir } from 'node:os';
import { fileURLToPath } from 'node:url';

const HERE = dirname(fileURLToPath(import.meta.url));
const BASE = resolve(process.argv[2] || join(HERE, '..'));
const UNIT = 'methodica-ar-math-ratio-03';
const PARTS = ['01', '02', '03', '04', '05', '06'];
const CDN = 'https://lomdot.education.gov.il/metodica/720active/common/';
const MIME = { '.html': 'text/html; charset=utf-8', '.js': 'text/javascript', '.css': 'text/css', '.png': 'image/png', '.jpg': 'image/jpeg',
  '.jpeg': 'image/jpeg', '.svg': 'image/svg+xml', '.ttf': 'font/ttf', '.mp4': 'video/mp4', '.mp3': 'audio/mpeg', '.json': 'application/json' };
const read = p => readFileSync(join(BASE, p), 'utf8');

let pass = 0; const failures = [];
const ok = (scope, name, cond, detail) => { if (cond) { pass++; return; } failures.push(`  [${scope}] ${name}` + (detail ? `\n      -> ${String(detail).slice(0, 300)}` : '')); };

/* /wg/  — the page without 50-loader.js / 90-boot.js: the test drives the boot (as the jsdom original).
   /raw/ — the page with EVERY external script removed: only the inline cover net runs (check n).
   Requests to the CDN preloads are left alone; nothing here depends on them. */
const srv = createServer(async (req, res) => {
  const u = new URL(req.url, 'http://x'); const p = decodeURIComponent(u.pathname);
  const m = /^\/(wg|raw)\/(.*)$/.exec(p);
  if (!m) { res.writeHead(404); res.end(); return; }
  const f = normalize(join(BASE, m[2]));
  try {
    const s = await stat(f); if (!s.isFile()) throw 0;
    let body = await readFile(f);
    if (/index\.html$/.test(f)) {
      let h = body.toString('utf8');
      h = m[1] === 'raw' ? h.replace(/<script\b[^>]*\bsrc="[^"]*"[^>]*><\/script>/g, '')
                         : h.replace(/<script\b[^>]*\bsrc="[^"]*(50-loader|90-boot)\.js[^"]*"[^>]*><\/script>/g, '');
      body = Buffer.from(h, 'utf8');
    }
    res.writeHead(200, { 'content-type': MIME[extname(f).toLowerCase()] || 'application/octet-stream', 'cache-control': 'no-store' });
    res.end(body);
  } catch { res.writeHead(404); res.end(); }
});
await new Promise(r => srv.listen(0, '127.0.0.1', r));
const ORIGIN = `http://127.0.0.1:${srv.address().port}/`;

const PORT = 9000 + Math.floor(Math.random() * 900);
const proc = spawn('C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe', ['--headless=new', '--disable-gpu', '--mute-audio',
  '--autoplay-policy=no-user-gesture-required', '--remote-debugging-port=' + PORT, '--user-data-dir=' + join(tmpdir(), 'ewg' + PORT),
  '--window-size=1280,720', 'about:blank'], { stdio: 'ignore' });
const sleep = ms => new Promise(r => setTimeout(r, ms));
let page = null;
for (let i = 0; i < 50 && !page; i++) { await sleep(300); try { page = (await (await fetch(`http://127.0.0.1:${PORT}/json/list`)).json()).find(t => t.type === 'page'); } catch {} }
const ws = new WebSocket(page.webSocketDebuggerUrl); await new Promise(r => ws.onopen = r);
let id = 0; const pend = new Map(); let evs = [];
ws.onmessage = m => { const x = JSON.parse(m.data); if (x.id && pend.has(x.id)) { const p = pend.get(x.id); pend.delete(x.id); x.error ? p.rej(new Error(JSON.stringify(x.error))) : p.res(x.result); } else if (x.method) evs.push(x); };
const send = (m, p = {}) => new Promise((res, rej) => { const i = ++id; pend.set(i, { res, rej }); ws.send(JSON.stringify({ id: i, method: m, params: p })); setTimeout(() => { if (pend.has(i)) { pend.delete(i); rej(new Error('timeout ' + m)); } }, 30000); });
const val = e => send('Runtime.evaluate', { expression: `(function(){ try { return (${e}); } catch (e) { return '__THREW__ ' + e.message; } })()`, returnByValue: true })
  .then(r => r.exceptionDetails ? '__THREW__ ' + (r.exceptionDetails.exception?.description || r.exceptionDetails.text) : r.result.value);
const run = e => send('Runtime.evaluate', { expression: e, returnByValue: true }).then(r => { if (r.exceptionDetails) throw new Error(r.exceptionDetails.exception?.description || r.exceptionDetails.text); });
await send('Runtime.enable'); await send('Page.enable'); await send('Network.enable'); await send('Network.setCacheDisabled', { cacheDisabled: true });
await send('Emulation.setDeviceMetricsOverride', { width: 1280, height: 720, deviceScaleFactor: 1, mobile: false });

async function nav(url) {
  evs = [];
  await send('Page.navigate', { url });
  for (let i = 0; i < 100; i++) { await sleep(100); if (await val('document.readyState') === 'complete') break; }
  await sleep(300);
}
const pageErrors = () => evs.filter(e => e.method === 'Runtime.exceptionThrown').map(e => (e.params.exceptionDetails.exception?.description || e.params.exceptionDetails.text).split('\n')[0]);

/* An in-memory state store behind the library's three state calls; the debounced stub keeps the
   document by reference, as the library does. 'l' adds saveState720OnLeave (counted), 'k' does not. */
const STUBS = (slug, lib) => `
  localStorage.clear(); sessionStorage.clear();
  window.__store = null; window.__debouncedDoc = undefined; window.__leave = [];
  window.__lr = { op: 'load', status: 404, ok: true, reason: 'absent' };
  window.loadState720 = function () {
    var d = window.__store ? JSON.parse(JSON.stringify(window.__store)) : null;
    window.__lr = d ? { op: 'load', status: 200, ok: true, reason: 'ok' } : { op: 'load', status: 404, ok: true, reason: 'absent' };
    return d;
  };
  window.saveState720 = function (id, doc) { window.__store = JSON.parse(JSON.stringify(doc)); return true; };
  window.saveState720Debounced = function (id, doc) { window.__debouncedDoc = doc; };
  window.stateLastResult720 = function () { return window.__lr; };
  ${lib === 'l' ? "window.saveState720OnLeave = function (id, doc) { window.__leave.push(doc && doc.payload ? doc.payload.currentScreen : null); return window.saveState720(id, doc); };" : 'delete window.saveState720OnLeave;'}
  window.sendStatement720 = function () {}; window.XAPI_USING_G = true;
  (function () { var x = new XMLHttpRequest(); x.open('GET', '../metadata/${slug}.json', false); x.send(); window.METADATA = JSON.parse(x.responseText); })();
  true`;
async function boot(c, lib = 'l') {
  const slug = `${UNIT}-${c}`;
  await nav(`${ORIGIN}wg/${slug}/index.html`);
  await run(STUBS(slug, lib));
  return slug;
}
const stored = () => val('window.__store');
const screenOf = d => (d && d.payload ? d.payload.currentScreen : null);

/* the seed: part 01's own payload on screen 2, captured from the page itself (v6) */
let slug = await boot('01');
await run('goTo(2)');
const payload2 = await val('capturePartPayload()');
const SEED = s => ({ v: 6, component: s, payload: payload2, done: {}, doneItems: {}, hints: {}, picks: {}, ui: { character: 'character-1' }, results: {} });
ok('setup', 'part 01 payload captured on screen 2', payload2 && payload2.currentScreen === 2, JSON.stringify(payload2));
const seed = async s => run(`window.__store = ${JSON.stringify(SEED(s))}; true`);

/* ── a–d: the drain ── */
slug = await boot('01'); await seed(slug);
await run("readUnitState(); setUnitCharacter('character-2'); goTo(1);");
let st = await stored();
ok('write-guard', 'a: a card pick + navigation before phase B write nothing',
  screenOf(st) === 2 && await val('_pendingProfile && _pendingProfile.character') === 'character-2' && await val('window.__debouncedDoc === undefined'), JSON.stringify(st && st.payload));
await run('_resumeReady = true; drainPendingUnitState();');
st = await stored();
ok('write-guard', 'b: the drain persists the document AS READ: saved payload intact, character updated',
  screenOf(st) === 2 && st.ui.character === 'character-2' && st.component === slug, JSON.stringify(st && { p: st.payload && st.payload.currentScreen, ui: st.ui }));
await run('applyExecutionState(_unitState.payload);');
ok('write-guard', 'c: the restore lands on the saved screen after the drain', await val('currentScreen') === 2, 'currentScreen=' + await val('currentScreen'));
ok('write-guard', 'd: the debounced document (held by reference) still carries the saved payload',
  await val('window.__debouncedDoc ? window.__debouncedDoc.payload.currentScreen : "none"') === 2, await val('JSON.stringify(window.__debouncedDoc && window.__debouncedDoc.payload && window.__debouncedDoc.payload.currentScreen)'));
ok('write-guard', 'a–d: no page error', pageErrors().length === 0, pageErrors().join(' | '));

/* ── e–f: markSent ── */
slug = await boot('01'); await seed(slug);
await run('readUnitState(); goTo(1); markSent("doneItems", "x");');
st = await stored();
ok('write-guard', 'e: markSent before phase B marks in memory but writes nothing',
  screenOf(st) === 2 && !(st.doneItems || {}).x && await val('_unitState.doneItems.x') === true, JSON.stringify(st && st.doneItems));
await run('_resumeReady = true; flushResumeSave();');
st = await stored();
ok('write-guard', 'f: the key marked in memory rides the next save', st && st.doneItems && st.doneItems.x === true, JSON.stringify(st && st.doneItems));

/* ── g–j: a failed read ── */
slug = await boot('01'); await seed(slug);
await run("window.__realLoad = window.loadState720; window.loadState720 = function () { window.__lr = { op: 'load', status: 500, ok: false, reason: 'rejected' }; return null; }; readUnitState();");
ok('write-guard', 'g: a read that fails (not a 404) is flagged, not taken as a fresh document',
  await val('_resumeReadFailed') === true && await val('_unitState.payload') === null, await val('String(_resumeReadFailed)'));
await run('_resumeReady = true; window.__p = persistUnitState(_unitState); flushResumeSave(); flushResumeSaveOnLeave(); markSent("done", "y"); goTo(1); scheduleResumeSave();');
st = await stored();
ok('write-guard', 'h: after a failed read every write is refused and the stored document is untouched',
  await val('window.__p') === false && screenOf(st) === 2 && st.done.y === undefined && await val('window.__debouncedDoc === undefined') && await val('window.__leave.length') === 0,
  JSON.stringify({ p: await val('window.__p'), cs: screenOf(st), done: st && st.done }));
await run('window.loadState720 = window.__realLoad; readUnitState();');
ok('write-guard', 'i: a read that succeeds afterwards clears the flag and restores the payload',
  await val('_resumeReadFailed') === false && await val('_unitState.payload && _unitState.payload.currentScreen') === 2);
await run('window.__store = null; readUnitState();');
ok('write-guard', 'j: a 404 ("nothing saved yet") is NOT a failed read', await val('_resumeReadFailed') === false && await val('_unitState.payload') === null);

/* ── k–l, s: source shape ── */
const r40 = read('unit-js/40-resume.js');
const r50 = read('unit-js/50-loader.js');
/* comments stripped: the code's own comments name what it must not do ("never captureUnitState()") */
const body = (src, fn) => ((src.split('function ' + fn + '(')[1] || '').split(/\nfunction /)[0]).replace(/\/\*[\s\S]*?\*\//g, '');
ok('shape', 'k: the drain persists _unitState, never a capture',
  /persistUnitState\(_unitState\)/.test(body(r40, 'drainPendingUnitState')) && !/captureUnitState\(/.test(body(r40, 'drainPendingUnitState')));
ok('shape', 'k: markSent persists only once _resumeReady', /if \(!_resumeReady\) return;[^\n]*\n\s*try \{ persistUnitState\(captureUnitState\(\)\)/.test(body(r40, 'markSent')));
ok('shape', 'k: persistUnitState refuses after a failed read', /^[^{]*\{(\s*var ok = false;)?(\s*\/\*[\s\S]*?\*\/)?\s*if \(_resumeReadFailed\) \{[\s\S]{0,160}return false;/.test(body(r40, 'persistUnitState')));
ok('shape', 'k: scheduleResumeSave / flushResumeSave bail after a failed read',
  /_resumeReadFailed\) return;/.test(body(r40, 'scheduleResumeSave')) && /_resumeReadFailed\) return;/.test(body(r40, 'flushResumeSave')));
ok('shape', 'l: phase A keeps the cover on a failed read', /if \(!_payload && !_resumeReadFailed\) dropBootCover\(\);/.test(r50));
ok('shape', 'l: phase B retries the read once', /if \(_resumeReadFailed\) \{\s*_saved = readUnitState\(\);/.test(r50));
ok('shape', 'l: the phase-B catch keeps a failed session closed', /_resumeReady = !_resumeReadFailed;/.test(r50));
ok('shape', 'l: the restore keeps _devHarness', /if \(_payload && !_devHarness\)/.test(r50));
ok('shape', 'l: a screen reached during the window is saved', /if \(!_resumed && currentScreen !== 0\) \{ try \{ scheduleResumeSave\(\); \}/.test(r50));
const letter = (r50.match(/RESUME_ENABLED \? 'xapi-720-([a-z])\.js'/) || [])[1];
ok('shape', 's: the loader loads xapi-720-l.js and the regex lists l', letter === 'l' && /XAPI_USING_G = \/xapi-720-\[ghijkl\]\\\.js\//.test(r50), 'letter ' + letter);
const leaving = body(r40, 'flushResumeSaveLeaving');
ok('shape', 's: the leave flush has the write guards and typeof-guards saveState720OnLeave',
  /_resumeReadFailed\) return;/.test(leaving) && /typeof window\.saveState720OnLeave === 'function'/.test(leaving));
ok('shape', 's: the leave handler goes through flushResumeSaveLeaving', /flushResumeSaveLeaving\(\);/.test(body(r40, 'flushResumeSaveOnLeave')));
ok('shape', 's: the -l test stub exists beside the -k one', existsSync(join(HERE, 'xapi-720-l.js')) && existsSync(join(HERE, 'xapi-720-k.js')));

/* ── m–p (markup) and t (leave, behaviour) for every part ── */
for (const c of PARTS) {
  const s = `${UNIT}-${c}`;
  const html = read(`${s}/index.html`);
  const nets = (html.match(/<script>[\s\S]*?<\/script>/g) || []).filter(x => /boot-cover/.test(x));
  ok(c, 'm: exactly one inline script handles boot-cover', nets.length === 1, String(nets.length));
  const net = nets[0] || '';
  ok(c, 'm: the net raises __resumeInFlight itself on a registration', /if \(\/\[\?&\]registration=\/\.test\(location\.search\)\) window\.__resumeInFlight = true;/.test(net));
  ok(c, 'm: the net keeps 800 ms and the 11000 ceiling', /setTimeout\(tick, 800\)/.test(net) && /< 11000\b/.test(net));
  for (const q of ['?slxapi=1&registration=abc', '']) {
    await nav(`${ORIGIN}raw/${s}/index.html${q}`);
    const on = await val('window.__resumeInFlight === true');
    ok(c, 'n: __resumeInFlight ' + (q ? 'raised with' : 'not raised without') + ' a registration (no external script)', q ? on : !on);
  }
  const head = html.split('</head>')[0];
  const pre = [...head.matchAll(/<link rel="preload" as="script"[^>]*>/g)].map(m => m[0]);
  ok(c, 'o: two script preloads, the exact loader URLs, letter ' + letter + ', no crossorigin / ?',
    pre.length === 2 && pre[0].includes('href="' + CDN + 'xapiwrapper.min.js"') && pre[1].includes('href="' + CDN + 'xapi-720-' + letter + '.js"') &&
    !pre.some(p => /crossorigin|\.js\?/.test(p)), pre.join(' '));
  ok(c, 'o: preconnect to Kata', /<link rel="preconnect" href="https:\/\/kata\.cet\.ac\.il">/.test(head));
  /* p: the landing section is the one carrying `screen active` in the markup (data-screen 0 here) */
  const masked = html.replace(/<!--[\s\S]*?-->/g, m => ' '.repeat(m.length));
  const m0 = masked.search(/<section class="screen active"/);
  const end0 = m0 < 0 ? -1 : m0 + 10 + masked.slice(m0 + 10).search(/<section\b/);
  const bad = [];
  for (const m of masked.matchAll(/<img\b[^>]*>/g)) {
    const inLanding = m.index >= m0 && m.index < end0;
    const src = (/\ssrc="([^"]*)"/.exec(m[0]) || [])[1];
    const lazy = /loading="lazy"/.test(m[0]);
    /* rule (§2.3 C3): under assets/, a .png under ../unit-assets/img/, or no src (script fills it in later);
       the shared SVG icons under ../unit-assets/img/ stay eager */
    const ruled = src === undefined || src.startsWith('assets/') || (src.startsWith('../unit-assets/img/') && src.endsWith('.png'));
    if (inLanding ? lazy : (ruled && !lazy)) bad.push(m[0].slice(0, 90));
  }
  ok(c, 'p: every image off the landing screen is lazy (assets/, unit-assets PNG, script-filled); none on it', bad.length === 0, bad.slice(0, 2).join(' | '));
  const vbad = [];
  for (const m of masked.matchAll(/<video\b[^>]*>/g)) {
    const inLanding = m.index >= m0 && m.index < end0;
    if (!inLanding && /\sautoplay\b/.test(m[0]) && /\ssrc="/.test(m[0])) vbad.push(m[0].slice(0, 80));
  }
  ok(c, 'p: no hidden autoplay <video> off the landing screen carries an inline src', vbad.length === 0, vbad.slice(0, 2).join(' | '));

  /* p (behaviour): after boot, no hidden <video> off the active screen has a src (part 06: the radar clips) */
  for (const lib of ['l', 'k']) {
    await boot(c, lib);
    if (lib === 'l') {
      const early = await val("[...document.querySelectorAll('.screen:not(.active) video')].filter(v => v.getAttribute('src') || v.getAttribute('poster')).map(v => v.id)");
      ok(c, 'p: no hidden <video> off the landing screen has a src or poster after boot', Array.isArray(early) && early.length === 0, JSON.stringify(early));
    }
    /* t: open the writes, move to screen 1, leave */
    await run(`readUnitState(); _resumeReady = true; drainPendingUnitState(); goTo(1); flushResumeSaveOnLeave();`);
    st = await stored();
    ok(c, `t -${lib}: the leave write carries the new screen`, screenOf(st) === 1, JSON.stringify(st && st.payload && st.payload.currentScreen));
    if (lib === 'l') ok(c, 't -l: it goes through saveState720OnLeave', JSON.stringify(await val('window.__leave')) === '[1]', JSON.stringify(await val('window.__leave')));
    ok(c, `t -${lib}: no page error`, pageErrors().length === 0, pageErrors().join(' | '));
  }
}

/* part 06: the radar clips arrive on screen entry */
await boot('06');
const before = await val("[document.getElementById('s6-video').getAttribute('src'), document.getElementById('s6-video').getAttribute('data-src')]");
await run('goTo(6)');
const after = await val("[document.getElementById('s6-video').getAttribute('src'), document.getElementById('s6-video').getAttribute('poster'), document.getElementById('s8-video').getAttribute('src')]");
ok('06', 'p: s6-video has no src at boot and gets src + poster when S6 opens (S8 still waits)',
  before[0] === null && /Blue_airplane/.test(before[1] || '') && /Blue_airplane/.test(after[0] || '') && /blue-video-frame0/.test(after[1] || '') && after[2] === null,
  JSON.stringify({ before, after }));
ok('06', 'p: no page error', pageErrors().length === 0, pageErrors().join(' | '));

ws.close(); proc.kill(); srv.close();
console.log(`\nwrite-guard (${BASE}): ${pass} passed, ${failures.length} failed`);
if (failures.length) { console.log(failures.join('\n')); process.exitCode = 1; }
