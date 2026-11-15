/* Static checks, no browser, no dependencies:  node _test/verify-static.js
   Ported from methodica-math-ratio-05 and retargeted to this repo's layout (js/main.js,
   css/style.css, assets/widgets/*.html).
   1. every file an index.html / widget page / stylesheet / script references exists with its exact
      case (Windows is case-insensitive; the CDN is not) and sits under a shipped folder
   2. every code reference in every index.html carries the SAME ?v= (bump them together)
   3. per component: XAPI_COMP_ID equals the metadata id; SCREEN_TO_SUBCONTENT has exactly
      TOTAL_SCREENS keys (0..N-1), names only catalogued items and maps every catalogued item;
      XAPI_EVAL_ITEMS are exactly the items that declare questions; every reportQ /
      xapiRequestedHint ref exists in the metadata; every declared question has exactly ONE
      reportQ site (D-9); a choice screen offers as many options as its question has answers
      (xapiChoiceAnswer reports answers[] by position)
   4. the library loader line and its XAPI_USING_G regex agree (guide §3.5); resume on loads -l (the
      leave save, BOOT-AND-RESUME-HARDENING-2026-11-03.md §2.10)
   5. guide §3.3: no top-level name is declared both in unit-js/ and in a part's main.js
      (regex scan of top-level declarations — the authoring scripts use the AST equivalent)
   6. per part (fix R, 2026-11-14, hardening doc §2.3 C1/C2): one inline boot-cover net, which raises
      __resumeInFlight itself on ?registration and keeps 800 ms and the 11000 ceiling; two script
      preloads byte-equal to the loader's URLs (letter from 50-loader.js, no ?v=, no crossorigin) and
      the Kata preconnect. Behaviour is in _test/write-guard.mjs. */
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const ROOT = path.join(__dirname, '..');
const PARTS = ['01', '02', '03', '04', '05', '06'].map(u => 'methodica-math-ratio-03-' + u);
let failures = 0, checks = 0;
const ok = (cond, msg) => { checks++; if (!cond) { failures++; console.log('  FAIL ' + msg); } };
const read = p => fs.readFileSync(path.join(ROOT, p), 'utf8');
const trim = s => String(s).replace(/\/+$/, '');

function existsExact(rel) {
  let dir = ROOT;
  for (const seg of rel.split('/')) {
    if (!fs.existsSync(dir) || !fs.readdirSync(dir).includes(seg)) return false;
    dir = path.join(dir, seg);
  }
  return true;
}
const SHIP = /^(unit-js\/[^/]+\.js|unit-css\/[^/]+\.css|unit-assets\/.+|metadata\/[^/]+\.json|methodica-math-ratio-03-0\d\/(index\.html|js\/main\.js|css\/style\.css|assets\/.+))$/;

function checkRef(fromFile, ref) {
  ref = ref.split(/[?#]/)[0];
  if (!ref || /^(https?:|data:|blob:|mailto:|#|javascript:|about:)/.test(ref) || ref.includes('${')) return;
  const rel = path.posix.normalize(path.posix.join(path.posix.dirname(fromFile), decodeURIComponent(ref)));
  ok(existsExact(rel), `${fromFile}: missing (or case-mismatched) ${ref}`);
  ok(SHIP.test(rel), `${fromFile}: ${ref} is not in a shipped folder (${rel})`);
}

/* 1. references */
const ASSET_RE = /['"]((?:\.\.\/)*(?:assets|unit-assets)\/[^'"\\]+\.(?:png|jpe?g|svg|mp4|mp3|ttf|woff2|html))['"]/g;
PARTS.forEach(u => {
  const p = u + '/index.html';
  const html = read(p).replace(/<!--[\s\S]*?-->/g, '');
  for (const m of html.matchAll(/\b(?:src|href|data-zoom-src|poster)="([^"]+)"/g)) checkRef(p, m[1]);
  for (const m of html.matchAll(/url\(['"]?([^'")]+)['"]?\)/g)) checkRef(p, m[1]);
  /* script paths resolve against the page, i.e. the component folder */
  for (const m of read(u + '/js/main.js').matchAll(ASSET_RE)) checkRef(p, m[1]);
  const wdir = path.join(ROOT, u, 'assets', 'widgets');
  if (fs.existsSync(wdir)) fs.readdirSync(wdir).filter(f => f.endsWith('.html')).forEach(w => {
    const wp = `${u}/assets/widgets/${w}`;
    /* bundled pages carry their CSS/markup inside escaped JS strings: unescape the quotes first */
    const t = read(wp).replace(/\\"/g, '"').replace(/\\u002F/g, '/');
    for (const m of t.matchAll(/(?:src|href)="((?:\.\.\/)+[^"]+)"/g)) checkRef(wp, m[1]);
    for (const m of t.matchAll(/url\('((?:\.\.\/)+[^']+)'\)/g)) checkRef(wp, m[1]);
  });
});
const cssFiles = [...fs.readdirSync(path.join(ROOT, 'unit-css')).map(f => 'unit-css/' + f), ...PARTS.map(u => u + '/css/style.css')];
cssFiles.forEach(f => {
  const css = read(f).replace(/\/\*[\s\S]*?\*\//g, '');
  for (const m of css.matchAll(/url\(['"]?([^'")]+)['"]?\)/g)) checkRef(f, m[1]);
});

/* 2. one cache-buster */
const vs = new Set();
PARTS.forEach(u => { for (const m of read(u + '/index.html').matchAll(/(?:src|href)="[^"]+\.(?:js|css)\?v=([^"]+)"/g)) vs.add(m[1]); });
ok(vs.size === 1, 'more than one ?v= across the index.html files: ' + [...vs].join(', '));
PARTS.forEach(u => {
  const html = read(u + '/index.html');
  for (const m of html.matchAll(/(?:src|href)="([^"]+\.(?:js|css))"/g)) if (!/^https?:/.test(m[1])) ok(false, `${u}/index.html: ${m[1]} has no ?v=`);   // the CDN preloads (C2) must NOT carry one
});

/* 3. xAPI seams against metadata */
PARTS.forEach(u => {
  const meta = JSON.parse(read('metadata/' + u + '.json'));
  const sandbox = { window: {}, location: { search: '' }, URLSearchParams, console };
  vm.createContext(sandbox);
  vm.runInContext(read('unit-js/10-identity.js'), sandbox);
  const js = read(u + '/js/main.js');
  const html = read(u + '/index.html');
  const grab = name => { const m = js.match(new RegExp('(?:var|const) ' + name + '\\s*=\\s*([^;]+);')); return m ? vm.runInContext('(' + m[1] + ')', sandbox) : undefined; };
  sandbox.XAPI_COMP_SLUG = grab('XAPI_COMP_SLUG');
  ok(sandbox.XAPI_COMP_SLUG === u, `${u}: XAPI_COMP_SLUG ${sandbox.XAPI_COMP_SLUG} != folder`);
  const compId = grab('XAPI_COMP_ID'), total = grab('TOTAL_SCREENS'), map = grab('SCREEN_TO_SUBCONTENT'), evals = grab('XAPI_EVAL_ITEMS') || {};
  ok(compId === meta.id, `${u}: XAPI_COMP_ID ${compId} != metadata ${meta.id}`);
  ok(grab('XAPI_METADATA_FILE') === '../metadata/' + u + '.json', `${u}: XAPI_METADATA_FILE`);
  ok(total === (html.match(/<section class="screen[ "]/g) || []).length, `${u}: TOTAL_SCREENS ${total} != number of .screen sections`);
  const keys = Object.keys(map || {}).map(Number).sort((a, b) => a - b);
  ok(keys.length === total && keys.every((k, i) => k === i), `${u}: SCREEN_TO_SUBCONTENT keys ${keys} vs TOTAL_SCREENS ${total}`);
  const itemIds = meta.subContent.map(s => trim(s.id).slice(-3));
  meta.subContent.forEach(s => ok(trim(s.id) === trim(meta.id) + '/' + u + '-' + trim(s.id).slice(-3), `${u}: item id shape ${s.id}`));
  const mapped = new Set(Object.values(map || {}).filter(Boolean).map(v => v[0]));
  mapped.forEach(i => ok(itemIds.includes(i), `${u}: SCREEN_TO_SUBCONTENT names unknown item ${i}`));
  itemIds.forEach(i => ok(mapped.has(i), `${u}: catalogued item ${i} has no screen`));
  const withQ = meta.subContent.filter(s => (s.questions || []).length).map(s => trim(s.id).slice(-3)).sort();
  ok(JSON.stringify(Object.keys(evals).sort()) === JSON.stringify(withQ), `${u}: XAPI_EVAL_ITEMS ${Object.keys(evals)} vs items with questions ${withQ}`);
  /* F-6: an all-sections item must be graded and have more than one question (one question needs no rule). */
  const allCorrect = grab('XAPI_ITEM_ALL_CORRECT') || {};
  Object.keys(allCorrect).forEach(i => { const s = meta.subContent.find(x => trim(x.id).slice(-3) === i); ok(evals[i] && s && (s.questions || []).length > 1, `${u}: XAPI_ITEM_ALL_CORRECT lists ${i}, which is not a graded multi-question item`); });
  const qset = new Set(); meta.subContent.forEach(s => (s.questions || []).forEach(q => qset.add(trim(s.id).slice(-3) + '/' + trim(q.questionId).split('/').pop())));
  const reported = {};
  for (const m of js.matchAll(/reportQ\(\['(\d{3})',\s*'(q\d+)'\]/g)) { const k = m[1] + '/' + m[2]; reported[k] = (reported[k] || 0) + 1; ok(qset.has(k), `${u}: reportQ refers to ${k}, not in metadata`); }
  qset.forEach(k => ok(reported[k] === 1, `${u}: question ${k} has ${reported[k] || 0} reportQ sites (want exactly 1, D-9)`));
  for (const m of js.matchAll(/xapiRequestedHint\('(\d{3})',\s*'(q\d+)'\)/g)) ok(qset.has(m[1] + '/' + m[2]), `${u}: hint refers to ${m[1]}/${m[2]}, not in metadata`);
  /* choice screens: option count == answers count */
  for (const m of js.matchAll(/xapiChoiceAnswer\('(\d{3})', '(q\d+)', '#(s\d+) \.scq-opt'/g)) {
    const [, item, q, scr] = m;
    const sec = (html.match(new RegExp(`<section class="screen[^"]*"[^>]*id="${scr}"[\\s\\S]*?</section>`)) || [''])[0].replace(/<!--[\s\S]*?-->/g, '');
    const n = (sec.match(/class="scq-opt[ "]/g) || []).length;
    const it = meta.subContent.find(s => trim(s.id).endsWith('-' + item));
    const qq = it && (it.questions || []).find(x => trim(x.questionId).split('/').pop() === q);
    ok(qq && Array.isArray(qq.answers) && qq.answers.length === n, `${u}: ${scr} has ${n} options but ${item}/${q} declares ${qq && qq.answers ? qq.answers.length : '?'} answers`);
  }
});

/* 4. loader letter vs regex */
const loader = read('unit-js/50-loader.js');
const letters = [...loader.matchAll(/xapi-720-([a-z])\.js'/g)].map(m => m[1]);
const re = (loader.match(/XAPI_USING_G = \/xapi-720-\[([a-z]+)\]/) || [])[1] || '';
letters.forEach(l => ok(re.includes(l), `50-loader.js loads xapi-720-${l}.js but XAPI_USING_G's regex lacks '${l}'`));
if (/RESUME_ENABLED = true;/.test(read('unit-js/10-identity.js'))) ok(letters.includes('l'), 'resume on must load -l (the leave save)');

/* 6. C1 / C2 per part */
const CDN = (loader.match(/var CDN = '([^']+)'/) || [])[1] || '';
const libLetter = (loader.match(/RESUME_ENABLED \? 'xapi-720-([a-z])\.js'/) || [])[1];
PARTS.forEach(u => {
  const html = read(u + '/index.html');
  const nets = (html.match(/<script>[\s\S]*?<\/script>/g) || []).filter(x => /boot-cover/.test(x));
  ok(nets.length === 1, `${u}: ${nets.length} inline scripts handle boot-cover (want 1)`);
  const net = nets[0] || '';
  ok(/if \(\/\[\?&\]registration=\/\.test\(location\.search\)\) window\.__resumeInFlight = true;/.test(net), `${u}: C1 the net does not raise __resumeInFlight on ?registration`);
  ok(/setTimeout\(tick, 800\)/.test(net) && /< 11000\b/.test(net) && /__resumeInFlight &&/.test(net), `${u}: C1 the net lost 800 ms / the 11000 ceiling / its __resumeInFlight wait`);
  const head = html.split('</head>')[0];
  const pre = [...head.matchAll(/<link rel="preload" as="script"[^>]*>/g)].map(m => m[0]);
  ok(pre.length === 2 && pre[0].includes('href="' + CDN + 'xapiwrapper.min.js"') && pre[1].includes('href="' + CDN + 'xapi-720-' + libLetter + '.js"')
     && !pre.some(x => /crossorigin|\.js\?/.test(x)), `${u}: C2 preloads are not the loader's two URLs (letter ${libLetter}): ${pre.join(' ')}`);
  ok(head.indexOf('rel="preload"') >= 0 && head.indexOf('rel="preload"') < head.indexOf('rel="stylesheet"'), `${u}: C2 the preloads must come before the first stylesheet`);
  ok(/<link rel="preconnect" href="https:\/\/kata\.cet\.ac\.il">/.test(head), `${u}: C2 no preconnect to kata.cet.ac.il`);
});

/* 5. no name in both layers */
const topNames = src => {
  const out = new Set();
  for (const m of src.matchAll(/^(?:function\s+([A-Za-z_$][\w$]*)|(?:var|let|const)\s+([A-Za-z_$][\w$]*))/gm)) out.add(m[1] || m[2]);
  return out;
};
const shared = new Map();
fs.readdirSync(path.join(ROOT, 'unit-js')).filter(f => f.endsWith('.js')).forEach(f => topNames(read('unit-js/' + f)).forEach(n => {
  ok(!shared.has(n), `${n} declared in both unit-js/${f} and unit-js/${shared.get(n)}`); shared.set(n, f);
}));
PARTS.forEach(u => topNames(read(u + '/js/main.js')).forEach(n => ok(!shared.has(n), `${u}/js/main.js re-declares ${n} (unit-js/${shared.get(n)})`)));

console.log(`${checks} checks, ${failures ? failures + ' FAILURE(S)' : 'ALL PASS'}`);
process.exit(failures ? 1 : 0);
