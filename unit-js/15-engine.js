/* ═══════════════════════════════════════════════════════════
   unit-js/15-engine.js  —  methodica-math-ratio-03 (all parts)
   Shared engine helpers: DECLARATIONS ONLY, nothing runs at load.
   Every declaration here used to be pasted, byte-identical, into each
   part's js/main.js (hash-checked at the move). Loaded before main.js,
   so the part's own top-level code (scaleApp() call, resize/load/
   pageshow listeners, PartNav.init, resetScreenState(0)) still runs in
   the same order. A name declared here must NOT be re-declared in any
   part's main.js (guide §3.3) — const/let would throw, function would
   silently win.

   Deliberately NOT here, because the parts' copies have drifted:
   goBack (01 differs), popupPointerDown (+ its popupDragId/OffX/OffY
   state), renderPopupAction, applyCompanionCharacter,
   updateProgressQuestion, openImageZoom, goTo, advanceScreen,
   resetScreenState — and all per-screen template code.
   ═══════════════════════════════════════════════════════════ */
'use strict';

/* ─── Companion character ───────────────────────────────────
   Part 01 WRITES the learner's choice to this namespaced key; every
   later part READS it back — the key string must stay byte-identical
   across all parts (720-templates → _global-components.md → Companion
   character system → "Cross-part persistence"). Confirmed against
   methodica-math-ratio-03-01/js/main.js and -02/js/main.js:
   CHARACTER_STORAGE_KEY = 'methodica_math_ratio_03_selectedCharacter'. */
const CHARACTER_STORAGE_KEY = 'methodica_math_ratio_03_selectedCharacter';

const DEFAULT_CHARACTER = 'character-1';

/* Validates the stored value against this part's own known character
   IDs — the canonical contract's requirement, not a truthiness check: a
   value left over from an older/renamed scheme falls back safely
   instead of reaching an asset lookup. Missing, empty or invalid all
   resolve to the same default. Never reads back into Part 01's own UI —
   this part has no character-selection screen, so that exclusion does
   not apply here. */
function getCharacter() {
  let stored = null;
  try { stored = localStorage.getItem(CHARACTER_STORAGE_KEY); } catch (e) {}
  return CHARACTER_IDS.includes(stored) ? stored : DEFAULT_CHARACTER;
}

/* S0 — TransitionScreen. No state machine: the screen has no
   interaction and nothing to resolve, so this is a pure mount. Runs on
   every entry (not once), so the character re-resolves if the learner
   went back to an earlier part and changed their selection. */
function s0ApplyCharacter() {
  applyCompanionCharacter(S0_CHARACTER, 's0-char-video');
}

/* ─── Canvas scaling ────────────────────────────────────────
   1280 × 710 reference canvas. Identical in every part of this lomda. */
function scaleApp() {
  const app = document.getElementById('app');
  const scale   = Math.min(window.innerWidth / 1280, window.innerHeight / 710);
  const canvasW = window.innerWidth  / scale;
  const canvasH = window.innerHeight / scale;
  app.style.width     = canvasW + 'px';
  app.style.height    = canvasH + 'px';
  app.style.transform = `scale(${scale})`;
  app.style.left      = '0px';
  app.style.top       = '0px';
}

/* Retry-popup title for questions with SEVERAL answer parts (client
   correction, 2026-09-29): when the board already shows some parts
   correct, "התשובה אינה נכונה." reads wrong next to the green marks, so
   those screens switch to the partial wording; a board with nothing
   right keeps the plain one. Single-answer screens are unaffected. */
const RETRY_TITLE         = 'الإجابة غير صحيحة.';

const RETRY_TITLE_PARTIAL = 'الإجابة ليست صحيحة بالكامل.';

/* Arrow keys must not hijack a control that has its own arrow-key
   meaning (text caret movement, radio group traversal, etc.). */
function hasOwnArrowKeyMeaning(el) {
  if (!el) return false;
  const tag = el.tagName;
  if (tag === 'TEXTAREA' || tag === 'SELECT') return true;
  if (tag === 'INPUT') {
    const noArrowMeaning = ['checkbox', 'radio', 'button', 'submit', 'reset', 'file', 'image'];
    return !noArrowMeaning.includes(el.type);
  }
  return false;
}

/* Relabel a button WITHOUT destroying its .nav-label width reservation
   (see the HOVER STABILITY block in css/style.css) — every runtime
   label change must go through here, never through textContent. */
function setNavLabel(el, text) {
  if (!el) return;
  let wrap = el.querySelector('.nav-label');
  if (!wrap) {
    wrap = document.createElement('span');
    wrap.className = 'nav-label';
    wrap.appendChild(document.createElement('span'));
    el.textContent = '';
    el.appendChild(wrap);
  }
  wrap.setAttribute('data-label', text);
  wrap.firstChild.textContent = text;
}

/* ═══════════════════════════════════════════════════════════
   Feedback popup — shared drag/clamp/action-button mechanics.
   Ported verbatim from the approved sibling
   methodica-math-ratio-03-02 (720-templates → _global-components.md →
   Feedback popup system). Reused as-is by every question screen in
   this part; do not reimplement per screen.
   ═══════════════════════════════════════════════════════════ */
const CANVAS_W = 1280, CANVAS_H = 710, BOTTOM_BAR_H = 74;

function resetPopupPosition(popup) {
  popup.style.left   = '2px';
  popup.style.top    = 'auto';
  popup.style.bottom = '76px';
}

/* Read the LIVE canvas size, not the constants above: scaleApp() grows
   #app past the nominal canvas on a surplus viewport, and the clamp has
   to follow the real edges. */
function getCanvasSize() {
  const app = document.getElementById('app');
  return {
    w: parseFloat(app.style.width)  || CANVAS_W,
    h: parseFloat(app.style.height) || CANVAS_H
  };
}

/* Never leave the canvas, never cover the 74px bottom bar. */
function clampPopup(popup) {
  if (!popup || popup.classList.contains('hidden')) return;
  const { w: canvasW, h: canvasH } = getCanvasSize();
  const left = parseFloat(popup.style.left);
  const top  = parseFloat(popup.style.top);
  if (!isNaN(left)) popup.style.left = Math.max(0, Math.min(left, canvasW - popup.offsetWidth)) + 'px';
  if (!isNaN(top))  popup.style.top  = Math.max(0, Math.min(top, canvasH - BOTTOM_BAR_H - popup.offsetHeight)) + 'px';
}

function getAppScale() {
  const app = document.getElementById('app');
  const m = app.style.transform.match(/scale\(([^)]+)\)/);
  return m ? parseFloat(m[1]) : 1;
}

function hideHintButton(id) {
  const hint = document.getElementById(id);
  if (!hint) return;
  hint.classList.add('hidden');
  hint.disabled = true;
}

function showHintButton(id) {
  const hint = document.getElementById(id);
  if (!hint) return;
  hint.classList.remove('hidden');
  hint.disabled = false;
}

/* ─── Scroll-gated popup — shared mechanics ──────────────────
   720-templates → _global-components.md → "Long-feedback variant —
   scroll-gated popup". New in this part: no sibling in this lomda has
   shipped the variant yet, so this is built from that spec rather than
   ported. Kept generic (popup id in, nothing screen-specific) so the
   remaining question screens of this part can reuse it. ── */
const SCROLL_END_TOLERANCE = 2;

/* scrollTop is fractional in most browsers while scrollHeight and
   clientHeight are rounded integers, so an equality test can be false
   by a fraction of a pixel at the very bottom and strand the learner
   with a permanently disabled progression button. The tolerance is
   mandatory, not a nicety. The same call answers "does this content
   overflow at all?" when made right after opening with scrollTop 0. */
function atBottom(el) {
  return el.scrollHeight - el.scrollTop - el.clientHeight <= SCROLL_END_TOLERANCE;
}

/* The bubble lives in the fixed 1280×710 canvas space, so scaleApp()
   scales it with everything else and there is no viewport-pixel maths
   to do — only a canvas-space fit check against the tooltip's own
   overall width (Figma: 210), whose first 4px are the transparent
   margin before the arrow tip. */
function placeScrollHint(popupId) {
  const popup = document.getElementById(popupId);
  if (!popup || popup.classList.contains('hidden')) return;
  const hint = popup.querySelector('.scq-scroll-hint');
  if (!hint) return;
  const { w: canvasW } = getCanvasSize();
  const left  = parseFloat(popup.style.left) || 0;
  const right = left + popup.offsetWidth;
  hint.classList.toggle('is-flipped', (canvasW - right) < 210);
}

function setsEqual(a, b) {
  if (a.size !== b.size) return false;
  for (const v of a) if (!b.has(v)) return false;
  return true;
}

/* ─── Screen-reader announcement ────────────────────────────
   Writes into the #a11y-announcer live region (index.html, next to the report dialogs). Cleared
   first and re-set a beat later so a repeat of the same message is announced again. Used by
   unit-js/25-report.js ("הדיווח נשלח, תודה"). */
function announce(msg) {
  const el = document.getElementById('a11y-announcer');
  if (!el || !msg) return;
  el.textContent = '';
  setTimeout(function () { el.textContent = msg; }, 50);
}
