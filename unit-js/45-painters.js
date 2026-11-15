'use strict';
/* ═══════════════════════════════════════════════════════════════════
   RESUME — the per-part payload and the screen painters (methodica-ar-math-ratio-03, all parts)
   Definition-only. 40-resume.js owns the document, the ledger and the save triggers; this file
   owns WHAT a part saves and HOW a saved screen is drawn again. Guide:
   Documentation/reporting-and-resume/ADDING-REPORTING-AND-RESUME.md §8.4 / §8.7 / §8.8.

   ── The registry (each part's js/main.js) ──
     PART_STATE_VARS   part-level state: names (any assignable expression, e.g. 'sectionResult' or
                       'window.lomdaState.s7') saved and restored as a whole
     PART_SCREENS[n] = {
       vars:     [names]   this screen's state variables (top-level let/const of main.js)
       inputs:   [ids]     fields whose value is state (it lives only in the DOM)
       pristine: () => bool   nothing attempted, nothing entered — "clean" never means attempts===0
                              alone (rule 4)
       init:     () => …   the screen's own ResetInitial: brings a FRESH DOM to its initial look.
                           It also zeroes the screen's variables; rebuild() restores them around it.
       paint:    () => …   draws the restored state on that DOM with the template's own functions
                           (Mark / Finish / Select / ApplyView / ShowPopup …). DOM only, never
                           resets, never reports, never throws out (rules 3, 5, 7).
     }
   Variables are read and written by NAME through Function() — classic-script top-level let/const
   live in the one global lexical scope, so this file can reach them without a closure per
   variable. const objects / Sets / arrays are refilled in place, never reassigned.

   ── When a painter runs ──
   In-session the DOM of a visited screen is still there, so a re-entry needs nothing beyond the
   part's own resetScreenState() branch (which now resets a screen only while it is pristine).
   After a RESTORE every saved screen's DOM is the fresh markup: applyExecutionState marks those
   screens stale, and repaintScreen(n) — called by every goTo — rebuilds a stale screen once:
     snapshot its live vars + inputs → init() → put them back → paint().
   The snapshot is the LIVE state, so progress made after the restore is never overwritten.
   ═══════════════════════════════════════════════════════════════════ */

/* ── Variable access by name ── */
function _pvGet(name) {
  try { return Function('"use strict"; return (' + name + ');')(); } catch (e) { return undefined; }
}
function _pvSet(name, v) {
  var cur = _pvGet(name);
  if (cur instanceof Set) { cur.clear(); (v || []).forEach(function (x) { cur.add(x); }); return; }
  if (Array.isArray(cur) && Array.isArray(v)) { cur.length = 0; v.forEach(function (x) { cur.push(x); }); return; }
  if (cur && typeof cur === 'object' && v && typeof v === 'object' && !Array.isArray(v)) {
    Object.keys(cur).forEach(function (k) { delete cur[k]; });
    Object.keys(v).forEach(function (k) { cur[k] = v[k]; });
    return;
  }
  try { Function('__v', '"use strict"; ' + name + ' = __v;')(v); }
  catch (e) { console.error('[resume] cannot restore ' + name, e); }
}
/* JSON-safe copy; a Set travels as an array (restored into the live Set by _pvSet). */
function _pvClone(v) {
  if (v instanceof Set) return Array.from(v);
  if (v === undefined) return null;
  try { return JSON.parse(JSON.stringify(v)); } catch (e) { return null; }
}

function _partScreens() { return (typeof PART_SCREENS !== 'undefined' && PART_SCREENS) || {}; }
function _partVars() {
  var names = (typeof PART_STATE_VARS !== 'undefined' && PART_STATE_VARS) ? PART_STATE_VARS.slice() : [];
  var S = _partScreens();
  Object.keys(S).forEach(function (n) { (S[n].vars || []).forEach(function (v) { names.push(v); }); });
  return names;
}
function _partInputs() {
  var ids = [], S = _partScreens();
  Object.keys(S).forEach(function (n) { (S[n].inputs || []).forEach(function (id) { ids.push(id); }); });
  return ids;
}
function _snapVars(names) { var o = {}; names.forEach(function (n) { o[n] = _pvClone(_pvGet(n)); }); return o; }
function _applyVars(o) { Object.keys(o || {}).forEach(function (n) { _pvSet(n, o[n]); }); }
function _snapInputs(ids) {
  var o = {};
  ids.forEach(function (id) { var el = document.getElementById(id); if (el && 'value' in el) o[id] = el.value; });
  return o;
}
function _applyInputs(o) {
  Object.keys(o || {}).forEach(function (id) { var el = document.getElementById(id); if (el && 'value' in el) el.value = o[id]; });
}

/* ── Helpers for the parts' registries ── */

/* Every listed field is empty (whitespace counts as empty). */
function resumeFieldsEmpty(ids) {
  return ids.every(function (id) { var el = document.getElementById(id); return !el || !String(el.value).trim(); });
}

/* Whether screen n is pristine — nothing attempted, nothing entered. A screen with no registry
   entry counts as pristine, so its resetScreenState() branch behaves as before. The parts call
   this where they used to reset an unresolved screen on every entry: re-entry now resets only a
   pristine screen, so a half-answered question keeps its attempts and its pending answer
   (decision 2026-10-01; guide §8.8 rules 4, 5, 9). */
function screenPristine(n) {
  var cfg = _partScreens()[n];
  if (!cfg || typeof cfg.pristine !== 'function') return true;
  try { return !!cfg.pristine(); } catch (e) { return true; }
}

/* ── The four resume hooks 40-resume.js / 50-loader.js call (guide §8.7) ── */

/* The part's payload: the screen, the per-question outcomes, every registered variable and field. */
function capturePartPayload() {
  return {
    currentScreen: currentScreen,
    qr: _pvClone(XAPI_Q_RESULTS),
    v: _snapVars(_partVars()),
    inp: _snapInputs(_partInputs())
  };
}

/* Answer variables. The parameter MUST be named st (guide §8.7). */
function applyResumeVars(st) {
  if (!st) return;
  Object.keys(XAPI_Q_RESULTS).forEach(function (k) { delete XAPI_Q_RESULTS[k]; });
  Object.keys(st.qr || {}).forEach(function (k) { XAPI_Q_RESULTS[k] = st.qr[k]; });
  _applyVars(st.v);
}

/* Values that live only in the DOM. */
function applyResumeDom(st) {
  if (!st) return;
  _applyInputs(st.inp);
}

/* Repaint one screen from the restored state (rebuild if stale). Exception-safe. */
function restoreScreenUI(n) { repaintScreen(n); }

/* ── Stale screens and the rebuild ── */
var _resumeStale = {};

function _rebuildScreen(n) {
  var cfg = _partScreens()[n];
  if (!cfg) return;
  if (typeof cfg.pristine === 'function' && cfg.pristine()) return;   // nothing to draw
  var names = (typeof PART_STATE_VARS !== 'undefined' ? PART_STATE_VARS : []).concat(cfg.vars || []);
  var vars = _snapVars(names);
  var inp  = _snapInputs(cfg.inputs || []);
  beginRepaint();
  try {
    if (typeof cfg.init === 'function') cfg.init();
    _applyVars(vars);
    _applyInputs(inp);
    if (typeof cfg.paint === 'function') cfg.paint();
  } catch (e) {
    console.error('[resume] paint screen ' + n, e);
  } finally {
    endRepaint();
  }
}

/* Called at the end of EVERY goTo (each part's main.js), after resetScreenState(n). */
function repaintScreen(n) {
  if (_resumeStale[n]) {
    delete _resumeStale[n];
    _rebuildScreen(n);
  }
  try { restoreEndedButton(); } catch (e) {}
}

/* The restore itself (guide §8.4). The sender is stubbed for its whole length, so painting a
   resolved screen can never re-report it; _restoring keeps the ledger and every save shut. */
function applyExecutionState(payload, overrideScreen) {
  if (!payload) return;
  var target = (typeof overrideScreen === 'number') ? overrideScreen : payload.currentScreen;
  if (typeof target !== 'number' || target < 0 || target >= TOTAL_SCREENS) target = 0;
  var realSend = window.sendStatement720;
  _restoring = true;
  window.sendStatement720 = function () {};
  try {
    applyResumeVars(payload);
    applyResumeDom(payload);
    _resumeStale = {};
    Object.keys(_partScreens()).forEach(function (n) { _resumeStale[n] = true; });
    goTo(target);                 // resetScreenState(target) → repaintScreen(target) rebuilds it
  } catch (e) {
    console.error('[resume] apply', e);
  } finally {
    window.sendStatement720 = realSend;
    _restoring = false;
  }
  /* The item that is open NOW is the landing screen's: re-open it for real (the stubbed goTo
     only moved the latch). */
  xapiCurrentItem = null;
  try { xapiSeedAnsweredFromResume(); } catch (e) {}
  try { xapiOnScreen(currentScreen); } catch (e) {}
}

/* ── The flush after an answer (D-13) ──
   reportQ (22-scoring.js) calls this right after the answered statement — but the Check that
   called reportQ has not finished writing its own state yet (result, done, marks). Deferring the
   synchronous save to the end of the current task puts it after the handler's LAST write, on every
   branch, including the early returns. One save per task. */
var _flushQueued = false;
function flushAfterCommit() {
  if (_flushQueued) return;
  _flushQueued = true;
  setTimeout(function () { _flushQueued = false; try { flushResumeSave(); } catch (e) {} }, 0);
}
