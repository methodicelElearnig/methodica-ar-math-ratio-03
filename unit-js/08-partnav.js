/* ═══════════════════════════════════════════════════════════
   unit-js/08-partnav.js  —  methodica-ar-math-ratio-03 (all parts)
   Cross-part navigation trail (PartNav). One shared copy: the engine body
   used to be duplicated, byte-identical, in every part's js/partnav.js.
   Mechanism ported from the approved sibling methodica-math-percent-01
   (720-templates → _global-components.md → "Back navigation in a
   branched journey").

   WHY THIS EXISTS
   Each part is its own page, so a part's FIRST screen has no sibling
   screen to go back to. A 720 learning journey is a graph, not a line,
   so "back" has to mean whichever predecessor the learner actually came
   through — which cannot be derived from the part number. The crossing
   is RECORDED as it happens and read back on the next page.

   THE TRAIL
   sessionStorage (per tab, gone on tab close) holds an array of
   { part, screen } entries, one per boundary crossed, in order.
   forward() pushes the part being LEFT plus the screen it was left
   from; back() pops the last entry and returns there. An empty trail
   means this part was opened directly, so there is genuinely nowhere to
   go back to and the destination's Back button stays hidden.

   The three constants are LOMDA-wide (same folder prefix, same two
   storage keys in every part); each page identifies itself through
   PartNav.init({ part: 'NN' }). Each part's OWN routing — which screen
   crosses out, and under which gate — lives in that part's
   advanceScreen() / crossToNextPart() in js/main.js, not here.

   THE PLATFORM OWNS ROUTING (guide §8.12, since 2026-09-16)
   Kata launches every component on its own URL and routes on the
   'completed' statement each part's last click sends (xapiEndPart,
   unit-js/20-xapi.js). So in production this module never navigates:
   forward() and back() are no-ops, forwardOverride() is null and the
   first screen's Back button stays hidden. Everything below is live only
   in a local walkthrough — DEV_NAV (unit-js/10-identity.js: ?dev=1 and no
   ?registration) — and every hop carries the query string along so the
   next part is still in the walkthrough.
   ═══════════════════════════════════════════════════════════ */
var PartNav = (function () {
  'use strict';

  /* Sibling-folder layout — the same relative path every cross-part
     redirect in this project uses literally. */
  const FOLDER_PREFIX = '../methodica-ar-math-ratio-03-';
  /* Namespaced like CHARACTER_STORAGE_KEY, for the same reason: an LMS
     may host several lomdot on one origin. */
  const TRAIL_KEY  = 'methodica_ar_math_ratio_03_navTrail';
  const RETURN_KEY = 'methodica_ar_math_ratio_03_navReturn';
  /* Defensive cap only. The real trail is at most one entry per part. */
  const MAX_TRAIL  = 16;

  let PART = null;                 // this page's part id, e.g. '02'
  /* DEV_NAV is declared by 10-identity.js, which loads after this file —
     so it is read at call time, never at load. */
  function devNav() { return typeof DEV_NAV !== 'undefined' && !!DEV_NAV; }
  let getScreen = function () { return 0; };
  let forwardTo = null;            // set only when entered via back()

  /* Storage access is wrapped everywhere: a sandboxed iframe or a
     privacy mode can make sessionStorage throw on ACCESS, not just on
     write. Every failure degrades to "no trail" = no Back button. */
  function store() {
    try { return window.sessionStorage; } catch (e) { return null; }
  }
  function read(key) {
    const s = store();
    if (!s) return null;
    try {
      const raw = s.getItem(key);
      return raw ? JSON.parse(raw) : null;
    } catch (e) { return null; }
  }
  function write(key, value) {
    const s = store();
    if (!s) return;
    try {
      if (value === null) s.removeItem(key);
      else s.setItem(key, JSON.stringify(value));
    } catch (e) {}
  }

  function trail() {
    const t = read(TRAIL_KEY);
    return Array.isArray(t) ? t : [];
  }
  /* The query string travels with the hop (?dev=1 must survive it; percent-02 0dcf18b). */
  function partUrl(part) { return FOLDER_PREFIX + part + '/index.html' + window.location.search; }

  /* ── Public API ─────────────────────────────────────────── */

  /* Called once per page, AFTER the part's own initial screen setup: if
     this page was entered via Back, it jumps straight to the screen the
     learner left from, through the part's own goTo() so the landing
     screen gets its normal activate + resetScreenState treatment.

     The jump is DEFERRED to a microtask — the landing screen's
     resetScreenState() can read `let` state variables declared FURTHER
     DOWN main.js than this call site; jumping synchronously would hit
     the temporal dead zone and throw, aborting the rest of main.js.
     Deferring runs the jump once main.js has finished executing — still
     inside the same load task, so nothing has painted yet. forwardTo is
     set synchronously, so forwardOverride() is correct immediately. */
  function init(opts) {
    opts = opts || {};
    PART = String(opts.part);
    if (typeof opts.getScreen === 'function') getScreen = opts.getScreen;

    if (!devNav()) return;
    const pending = read(RETURN_KEY);
    if (!pending || pending.part !== PART) return;
    write(RETURN_KEY, null);              // one-shot: consumed on arrival
    forwardTo = pending.forwardTo || null;
    if (typeof opts.goTo !== 'function' || typeof pending.screen !== 'number') return;
    const jump = function () { opts.goTo(pending.screen); };
    if (typeof queueMicrotask === 'function') queueMicrotask(jump);
    else setTimeout(jump, 0);
  }

  /* Cross a part boundary forward. Replaces a bare
     `window.location.href = '../methodica-ar-math-ratio-03-NN/index.html'`
     — same navigation, plus the trail entry that makes the
     destination's Back button work. */
  function forward(part) {
    if (!devNav()) return;
    part = String(part);
    const t = trail();
    /* A cycle would otherwise grow the trail forever. This keeps the
       trail a true path if one is ever added. */
    const seenAt = t.findIndex(function (e) { return e && e.part === part; });
    if (seenAt !== -1) t.length = seenAt;
    t.push({ part: PART, screen: getScreen() });
    while (t.length > MAX_TRAIL) t.shift();
    write(TRAIL_KEY, t);
    forwardTo = null;
    try { if (typeof resumeBeforeLeave === 'function') resumeBeforeLeave(); } catch (e) {}
    window.location.href = partUrl(part);
  }

  function canGoBack() { return devNav() && trail().length > 0; }

  /* Cross a part boundary backwards, to the last screen of the part the
     learner actually arrived from. No-op (not an error) on an empty
     trail, so a stray keyboard ArrowRight on S0 behaves exactly as it
     would if this module did not exist. */
  function back() {
    if (!devNav()) return;
    const t = trail();
    const entry = t.pop();
    if (!entry || !entry.part) return;
    write(TRAIL_KEY, t);
    write(RETURN_KEY, {
      part: entry.part,
      screen: typeof entry.screen === 'number' ? entry.screen : 0,
      forwardTo: PART               // so that part can re-cross forward
    });
    try { if (typeof resumeBeforeLeave === 'function') resumeBeforeLeave(); } catch (e) {}
    window.location.href = partUrl(entry.part);
  }

  /* The part this page was entered FROM via Back, or null. */
  function forwardOverride() { return devNav() ? forwardTo : null; }

  /* Show/hide a first-screen Back button to match trail state. Called
     from the part's resetScreenState(0) so it is correct on initial
     load and on every return to S0 from within the part. */
  function mountBack(id) {
    const btn = document.getElementById(id || 's0-back');
    if (!btn) return;
    /* Hidden by class AND attribute (guide §8.12): .hidden is display:none, and
       [hidden] keeps it out of the accessibility tree even if a rule overrides it. */
    const show = canGoBack();
    btn.classList.toggle('hidden', !show);
    btn.hidden = !show;
  }

  /* ?resetState (40-resume.js) empties the walkthrough trail too. */
  function clearTrail() { write(TRAIL_KEY, null); write(RETURN_KEY, null); }

  return {
    init: init,
    forward: forward,
    back: back,
    canGoBack: canGoBack,
    forwardOverride: forwardOverride,
    mountBack: mountBack,
    clearTrail: clearTrail
  };
})();
