'use strict';

/* ═══════════════════════════════════════════════════════════
   js/main.js  —  methodica-ar-math-ratio-03-02
   720 reference engine (canvas 1280×710) + Screen 00
   (TransitionScreen — practice-stage intro, "מסך הזנקה").

   Source: the approved production script
   יעד 1.3_יחס וביטויים אלגבריים_להפקה_06092026.pptx, one level above the
   project root. Slide identity re-derived EACH session from
   presentation.xml's sldIdLst (order 29 → slide29.xml, order 30 →
   slide30.xml, both 1:1 with no offset — the pptx's own mtime advanced
   between sessions, so this was re-derived fresh rather than trusted
   from an earlier extraction); presenter notes resolved through each
   slide's own .rels (part-name numbering is NOT slide numbering —
   pptx-script-extraction). Component divider: slide 28 ("רכיב 2",
   "תרגול סטנדרטי", "methodica-ar-math-ratio-03-02").

   Screen roster (built so far):
     S0 — TransitionScreen, practice-stage intro (slide 29).
     S1 — ValueInputQuestion (local ratio variant), practice Q1 of 3
          (slide 30).
     S2 — SingleChoiceQuestion, practice Q2 of 3 (slide 31).
     S3 — ValueInputQuestion (local ratio variant), practice Q3 of 3,
          final question (slide 32).

   S0 ported from the approved Math sibling
   methodica-math-percent-01/methodica-math-percent-01-02's own S0 — the
   identical screen role (practice-stage intro, also slide 29 in that
   deck: "סיימנו את שלב הלימוד, נעבור לתרגול הראשון ונענה על" / "N שאלות"
   / "צריך לענות נכון על N שאלות לפחות."), per 720-templates →
   references/TransitionScreen.md + references/_global-components.md.
   S1's shell (.scq-question / .scq-img-analog / .scq-right-col /
   .progress-question / .scq-popup / .scq-hint-overlay / bottom bar) is
   ported from the same sibling's own S1 (practice Q1 of that part's
   series) — the same "first question after the kickoff transition"
   role. S1's INPUT is a local variant of 720-templates →
   ValueInputQuestion.md's Variant B (fraction): same one-answer-object
   contract (single combined verdict icon, both fields judged together)
   but laid out as a horizontal ratio "3 : 2" (two boxes + a colon) since
   slide 30 authors a ratio, not a stacked fraction — see the S1VIQ block
   below for the full delta note.
   S2 reuses S1's own shell verbatim (canvas, Progress Question nav,
   feedback popup, hint overlay, bottom bar) per 720-templates →
   SingleChoiceQuestion.md — four answer pills, no View My Answer
   Toggle (SCQ shows the chosen and correct options side by side
   instead). See the S2SC block below for the full source-evidence note
   (including how the "3a"/"2a" split-run ambiguity in the popup text
   was resolved by cross-checking this slide's own stem text and its
   own OMML calculation — not by reading DOM/bidi order as intent).
   Engine plumbing (scaleApp / goTo / resetScreenState / keyboard nav /
   dev bridge) ported from the same-lomda sibling
   methodica-ar-math-ratio-03-01, byte-identical, so both parts of this
   lomda behave the same. Cross-part Back uses this part's own
   js/partnav.js (PartNav), same engine as the percent-01 family.
   ═══════════════════════════════════════════════════════════ */

/* ─── Constants ─────────────────────────────────────────── */
const TOTAL_SCREENS = 9;  // S0-S8 — slide 37 is the last screen of part 02


/* ─── S0 companion-character assets ─────────────────────────
   Explicit ID → asset mapping (the supplied filenames,
   "Character_1_Book.mp4" / "Character_2_Book.mp4", don't follow the
   character-N-[state] convention, and resolving by filename/colour
   guessing is forbidden regardless — 720-templates →
   _global-components.md → Companion character system).

   LOGICAL MAPPING confirmed against methodica-ar-math-ratio-03-01's own
   CHARACTER_ASSETS (same lomda, same two companions):
       character-1 = the ORANGE round one-eyed character
       character-2 = the TALL YELLOW character with the purple headband

   Per-instance asset map (TransitionScreen.md → "Companion-character
   asset resolution": each TransitionScreen instance keeps its own map,
   not a lomda-wide one) — this is this part's only TransitionScreen so
   far.

   Verified 2026-09-08 by loading each clip directly in-browser and
   reading videoWidth/videoHeight/duration, then screenshotting a late
   frame: both are 1280×720, 4.01s, and both show the matching
   companion character wearing glasses and reading a book, surrounded by
   piles of open books — matching the presenter note verbatim ("הדמות עם
   משקפיים גדולים ומלא ספרים פתוחים מולה"). No `fit` shim: both
   compositions already fill their own 1280×720 frame edge-to-edge
   (books extend to the left/right/bottom edges in both), so a plain
   contain-fit renders them at the same physical scale in the 640×360
   box — see the note on `.s1-char-media` in css/style.css for the full
   measurement record. */
const S0_CHARACTER = {
  'character-1': {
    video: 'assets/video/Character_1_Book.mp4',
    alt: 'الشخصية البرتقالية المستديرة ذات العين الواحدة، ترتدي نظارات، تقرأ كتابًا ومحاطة بأكوام من الكتب المفتوحة'
  },
  'character-2': {
    video: 'assets/video/Character_2_Book.mp4',
    alt: 'الشخصية الصفراء الطويلة ذات عصبة الرأس البنفسجية، ترتدي نظارات، تقرأ كتابًا وتقف محاطة بأكوام من الكتب المفتوحة'
  }
};
const CHARACTER_IDS = Object.keys(S0_CHARACTER);


/* ─── Global lomda state ─────────────────────────────────── */
window.lomdaState = window.lomdaState || {
  selectedCharacter: getCharacter()
};

/* Decorative character clip: plays once on its own, never loops, and
   stays on its final encoded frame — step 1 of the canonical end-state
   hierarchy (720-templates → _global-components.md → Companion
   character system → "Decorative character video" → End-state
   hierarchy). No `freezeVideoOnEnd()`-style handler is wired: that
   pattern reset clips to their OPENING pose and is explicitly
   prohibited. Both clips were sampled during build verification and
   hold a meaningful final frame, so nothing further is needed here. */

/* Resolve and mount the selected character's clip. Re-entry safe: the
   src is only touched when it actually differs, so returning to the
   screen never reloads/restarts an already-correct clip. */
function applyCompanionCharacter(characterMap, videoId) {
  const char = window.lomdaState.selectedCharacter || DEFAULT_CHARACTER;
  const assets = characterMap[char] || characterMap[DEFAULT_CHARACTER];
  const vid = document.getElementById(videoId);
  if (!vid || !assets) return;
  // Decorative only — never in the accessibility tree.
  vid.setAttribute('aria-hidden', 'true');
  vid.alt = '';
  /* Per-character fit shim. ALWAYS written, including the no-op values,
     so switching characters can never leave a previous one's scale
     behind on the element (none of this part's entries carry `fit` —
     see the note on S0_CHARACTER / css/style.css). */
  const fit = assets.fit || {};
  vid.style.setProperty('--char-scale',   fit.scale  || 1);
  vid.style.setProperty('--char-shift-y', (fit.shiftY || 0) + 'px');
  if (!vid.getAttribute('src') || !vid.getAttribute('src').endsWith(assets.video)) {
    vid.src = assets.video;
    vid.load();
    vid.play().catch(() => {});
  } else if (vid.paused && !vid.ended) {
    /* Re-entry after an interrupted playback (_global-components.md →
       "Re-entry after an interrupted playback"): goTo() paused this clip
       when the learner left mid-play, so continue it from where it stopped.
       An ended clip stays on its final frame; nothing seeks. No freezeAt
       here (see the S0_CHARACTER note). */
    vid.play().catch(() => {});
  }
  /* Old iOS Safari re-runs autoplay when the screen is shown again and would
     replay an ENDED clip from 0: drop the attribute once, at its end. */
  if (!vid.dataset.autoplayGuard) {
    vid.dataset.autoplayGuard = '1';
    vid.addEventListener('ended', () => vid.removeAttribute('autoplay'), { once: true });
  }
}


/* ─── State ─────────────────────────────────────────────── */
let currentScreen = 0;

window.addEventListener('resize', scaleApp);
scaleApp();
/* Re-measure once the real viewport is known: a mobile/tablet browser can report the
   content's own width at the first call and settle without a resize event, leaving the
   canvas oversized with the bottom bar and Continue off-screen. pageshow covers a
   back/forward-cache restore (figma-lomda-builder → Viewport scaling). */
window.addEventListener('load', scaleApp);
window.addEventListener('pageshow', e => { if (e.persisted) scaleApp(); });


/* The first screen is already .active in the markup, so it never goes
   through goTo() — reveal it once here. The engine holds the entrance
   paused until window.load so it never plays against the loading
   white-screen. */
if (window.Motion) Motion.reveal(document.querySelector('.screen.active'));

/* ─── Navigation ────────────────────────────────────────── */
function goTo(n) {
  if (n < 0 || n >= TOTAL_SCREENS) return;
  // Feedback popups / hint overlays live outside .screen divs — close
  // them all on every navigation so none carries over to the next
  // screen (720-templates → _global-components.md → "Advance gates" /
  // Feedback popup system). Every question screen in this part has
  // them; the sweep is the shared engine shape every part carries.
  document.querySelectorAll('[id$="-popup"], [id$="-hint-overlay"]')
    .forEach(el => el.classList.add('hidden'));
  /* Neither zoom surface matches that sweep (they are not feedback
     popups), so they are closed explicitly on EVERY screen change — a
     mouse, button or keyboard transition alike. The applet zoom is closed
     even when the destination is another applet screen: the promoted card
     belongs to the screen being left. */
  closeImageZoomOnNavigate();
  closeAppletZoomOnNavigate();
  const prev = document.querySelector('.screen.active');
  if (prev) {
    prev.querySelectorAll('video').forEach(v => { try { v.pause(); } catch (e) {} });
    /* Clear a reveal that is still mid-flight, so leaving a screen early
       can never strand one of its groups at opacity 0 (animation-system). */
    if (window.Motion) Motion.cancel(prev);
    prev.classList.remove('active');
  }
  currentScreen = n;
  const next = document.getElementById('s' + n);
  if (next) next.classList.add('active');
  /* Straight after .active, per the animation-system wiring contract.
     Only screens carrying [data-reveal] animate; the reveal plays on a
     screen's FIRST entry per session, and a revisit, Back or DEV jump
     shows everything instantly. This is the only call site — scaleApp()
     never calls it, so a resize cannot replay a reveal. */
  if (window.Motion) Motion.reveal(next);
  resetScreenState(n);
  /* Resume: rebuild a screen restored from the state document, once (45-painters.js). */
  repaintScreen(n);
  /* xAPI: item-level initialized/completed on an item boundary (20-xapi.js). */
  xapiOnScreen(n);
  /* Resume: the screen change is the save choke point (debounced, 40-resume.js). */
  scheduleResumeSave();
}

/* Restores a screen's state when the learner navigates away and back
   again WITHIN the same loaded page (no reload). Screen 0 must call
   PartNav.mountBack() so the cross-part Back button stays in sync with
   the navigation trail, and re-resolve the companion character on every
   entry (720-templates → _global-components.md → Companion character
   system → verification checklist). */
function resetScreenState(n) {
  if (n === 0) {
    PartNav.mountBack('s0-back');
    s0ApplyCharacter();
  }
  if (n === 1) {
    /* Resume-state lock: a resolved question must never hand out a
       fresh attempt on return, and a resolved screen never comes back
       without its feedback (_global-components.md → "Reopening
       feedback on return to a resolved screen"). s1viqDone is the only
       flag involved — there is deliberately no second
       "popup was open on exit" flag. */
    if (!s1viqDone) { if (screenPristine(1)) s1viqResetInitial(); }
    else if (s1viqResult === 'correct') s1viqShowPopup('correct', 'solution');
    else { s1viqView = 'solution'; s1viqApplyView(); }
    enterSection(1);
  }
  if (n === 2) {
    /* Same resume/reopen contract as S1 — SingleChoiceQuestion has no
       View My Answer Toggle (it shows the chosen and correct options
       side by side already), so reopening a resolved screen is just
       re-showing the last outcome's popup; the option marks themselves
       are untouched DOM state that already persisted. */
    if (!s2scDone) { if (screenPristine(2)) s2scResetInitial(); }
    else s2scShowPopup(s2scResult);
    enterSection(2);
  }
  if (n === 3) {
    /* Same resume/reopen contract as S1 (View My Answer Toggle applies
       — this is ValueInputQuestion again, not SCQ). */
    if (!s3viqDone) { if (screenPristine(3)) s3viqResetInitial(); }
    else if (s3viqResult === 'correct') s3viqShowPopup('correct', 'solution');
    else { s3viqView = 'solution'; s3viqApplyView(); }
    enterSection(3);
  }
  if (n === 4) {
    /* Same resume/reopen contract as S1/S3 — ValueInputQuestion again,
       so the View My Answer Toggle applies. This screen belongs to the
       section א of Question 2 — the indicator stays on שאלה 2 across
       S4 and S5, and Question 2 resolves only once S5 is answered. */
    if (!s4viqDone) { if (screenPristine(4)) s4viqResetInitial(); }
    else if (s4viqResult === 'correct') s4viqShowPopup('correct', 'solution');
    else { s4viqView = 'solution'; s4viqApplyView(); }
    enterSection(4);
  }
  if (n === 5) {
    /* Same resume/reopen contract as S2 — SingleChoiceQuestion has no
       View My Answer Toggle, so reopening a resolved screen is just
       re-showing the last outcome; the option marks are DOM state that
       already persisted. Section ב of Question 2 — its LAST section,
       so answering it is what resolves Question 2. */
    if (!s5scDone) { if (screenPristine(5)) s5scResetInitial(); }
    else s5scShowPopup(s5scResult);
    enterSection(5);
  }
  if (n === 6) {
    /* Same resume/reopen contract as S1/S3/S4 — ValueInputQuestion, so
       the View My Answer Toggle applies. Section א of Question 3 — the
       indicator stays on שאלה 3 across S6, S7 and S8. The applet is
       deliberately
       untouched here: its iframe is mounted once in the markup and
       never re-created, so re-entering the screen keeps whatever the
       learner marked on the graph. */
    if (!s6viqDone) { if (screenPristine(6)) s6viqResetInitial(); }
    else if (s6viqResult === 'correct') s6viqShowPopup('correct', 'solution');
    else { s6viqView = 'solution'; s6viqApplyView(); }
    enterSection(6);
  }
  if (n === 7) {
    /* Same resume/reopen contract as S6, and the same deliberate
       hands-off treatment of the applet: its iframe is mounted once in
       the markup and never re-created, so re-entering keeps the range
       the learner marked on the graph — which on THIS screen is the
       mechanism of the question, not a convenience. */
    const st = s7dqState();
    if (!st.done) { if (screenPristine(7)) s7dqResetInitial(); }
    else if (st.answer === null) s7dqShowPopup('correct', 'solution');
    else { st.view = 'solution'; s7dqApplyView(); }
    enterSection(7);
  }
  if (n === 8) {
    /* Same resume/reopen contract as S6, and the same hands-off
       treatment of the applet: the iframe is mounted once and never
       re-created, so re-entering keeps the learner's graph state. */
    if (!s8viqDone) { if (screenPristine(8)) s8viqResetInitial(); }
    else if (s8viqResult === 'correct') s8viqShowPopup('correct', 'solution');
    else { s8viqView = 'solution'; s8viqApplyView(); }
    enterSection(8);
  }
}

/* ─── Forward navigation ─────────────────────────────────────
   The gate lives HERE, not only in the Continue button's handler: the
   engine binds ArrowLeft directly to advanceScreen(), so a gate that
   sat only in an onclick would be silently bypassable by keyboard
   (720-templates → _global-components.md → "Advance gates").
   S0 has NO gate: a TransitionScreen carries no interaction to resolve,
   and slide 29 mocks Continue in its enabled state
   (720-templates → TransitionScreen.md → "A. Structural contract").
   S1, S2 and S3 gate on RESOLVED, never on correctness — a final wrong
   attempt still resolves the question and still lets the learner move
   on. */
function advanceScreen() {
  if (currentScreen === 1 && !s1viqDone) return;
  if (currentScreen === 2 && !s2scDone) return;
  if (currentScreen === 3 && !s3viqDone) return;
  if (currentScreen === 4 && !s4viqDone) return;
  if (currentScreen === 5 && !s5scDone) return;
  if (currentScreen === 6 && !s6viqDone) return;
  if (currentScreen === 7 && !s7dqState().done) return;
  if (currentScreen === 8 && !s8viqDone) return;
  /* S8 is the last screen of this part. Leaving it is a Decision Point,
     not an ordinary step, so it goes through crossToNextPart() rather
     than goTo(9) — which TOTAL_SCREENS would silently swallow. */
  if (currentScreen === 8) { crossToNextPart(); return; }
  goTo(currentScreen + 1);
}

/* ═══════════════════════════════════════════════════════════
   END-OF-PART DECISION POINT — S8 → methodica-ar-math-ratio-03-03

   ROUTE (producer decision, 2026-09-13):
     masteryState().passed === true   → open 03-03, land on its S0
     masteryState().passed === false  → stay in 03-02; nothing opens

   WHAT THE CONDITION COUNTS. Only the three AGGREGATED question
   results, never the eight section results:
       שאלה 1 = S1 AND S2 AND S3
       שאלה 2 = S4 AND S5
       שאלה 3 = S6 AND S7 AND S8
   masteryState() derives those through questionResult(); nothing here
   touches sectionResult directly. The 2-of-3 threshold is both the
   authored condition on slide 29 ("צריך לענות נכון על 2 שאלות לפחות")
   and the 720 default ceil(2N/3) for N=3 (720-templates →
   _question-template-defaults.md → "Pass threshold formula"), so there
   is no deviation to surface.

   The route is evaluated only once all three questions have RESOLVED.
   Reaching S8's own resolved state is not sufficient on its own — a
   learner who jumped screens through the progress nav can resolve S8
   with earlier questions still open — so allResolved is checked
   explicitly rather than inferred from linear progression.

   MECHANISM. The crossing itself is PartNav.forward(), the project's
   existing cross-part contract (js/partnav.js): it pushes this part
   and the screen being left onto the sessionStorage trail before
   navigating, which is what gives 03-03's S0 a Back button that
   returns HERE, to S8, instead of to a hardcoded predecessor. No ad
   hoc window.location assignment is used. Entering 03-03 directly
   leaves the trail empty, so no Back destination is fabricated.

   RE-CROSSING. If the learner already crossed this boundary and came
   back, PartNav.forwardOverride() reports the destination recorded for
   their journey, and that destination is REPLAYED rather than the gate
   re-run (720-templates → _global-components.md → "Re-crossing a
   previously-resolved boundary forward must replay the recorded
   destination, not re-evaluate the original gate"). This matters
   precisely because the trail does not restore in-progress screen
   state: re-running the gate against a freshly reset part would block
   a learner who had already legitimately passed.
   ═══════════════════════════════════════════════════════════ */
const NEXT_PART = '03';

/* One navigation per crossing. Click, Enter/Space on the button and
   ArrowLeft all funnel through advanceScreen(), and a browser keeps
   running script while it tears the page down, so rapid activation
   would otherwise assign window.location.href several times. */
let partCrossingInFlight = false;

/* Platform routing (decision 2026-10-01, guide §6.4 / §6.5 / §8.8 rule 12): the last click
   reports the component 'completed' — pass AND fail — disables the button and stops; Kata routes
   on the statement (recommendedAfterFail on a fail). This supersedes the earlier producer rule
   that kept המשך enabled on a fail. The hop to the next part survives only in a ?dev=1
   walkthrough (xapiEndPart, unit-js/20-xapi.js). */
function crossToNextPart() {
  if (partCrossingInFlight) return;
  if (!masteryState().allResolved) return;
  if (xapiEndPart(lastScreenButton(), NEXT_PART)) partCrossingInFlight = true;
}

/* Back returns to the previous valid predecessor in the learner's
   actual route. From this part's FIRST screen, Back crosses the part
   boundary backwards via PartNav instead of no-opping — with no
   recorded crossing (this part opened directly, e.g. in dev) that is
   itself a no-op, which is why the button stays hidden until
   PartNav.mountBack() confirms a real trail entry. */
function goBack() {
  if (currentScreen === 0) { PartNav.back(); return; }
  goTo(currentScreen - 1);
}

document.addEventListener('keydown', e => {
  if (hasOwnArrowKeyMeaning(e.target)) return;
  if (e.key === 'ArrowLeft')  advanceScreen();
  if (e.key === 'ArrowRight') goBack();
});


let popupDragId = null, popupOffX = 0, popupOffY = 0;


function popupPointerDown(e, popupId) {
  /* The View My Answer Toggle's action button is a control, not a drag
     handle — without this guard its own pointerdown would both fire its
     click AND start a popup drag. */
  if (e.target && e.target.closest && e.target.closest('.popup-action')) return;
  const popup = document.getElementById(popupId);
  const app   = document.getElementById('app');
  if (!popup || !app) return;
  popupDragId = popupId;
  const scale = getAppScale();
  const canvasY = (e.clientY - (parseFloat(app.style.top) || 0)) / scale;
  const canvasX = (e.clientX - (parseFloat(app.style.left) || 0)) / scale;
  const topPx = parseFloat(popup.style.top);
  const popupTop = isNaN(topPx) ? popup.offsetTop : topPx;
  popup.style.top    = popupTop + 'px';
  popup.style.bottom = 'auto';
  popupOffX = canvasX - (parseFloat(popup.style.left) || 2);
  popupOffY = canvasY - popupTop;
  e.preventDefault();
  /* S6-S8 put the popup over the applet's <iframe>. Without capture, a fast
     drag that leaves the popup hands the pointer to the iframe's document and
     the parent stops receiving moves and the release. Capturing keeps the whole
     gesture on the popup; the listeners below still see it (events bubble). */
  try { popup.setPointerCapture(e.pointerId); } catch (err) { /* not capturable: fall back to the plain drag */ }
}
window.addEventListener('pointermove', e => {
  if (!popupDragId) return;
  if (e.buttons === 0) { popupDragId = null; return; }   // release was missed: never follow a free cursor
  const popup = document.getElementById(popupDragId);
  const app   = document.getElementById('app');
  if (!popup || !app) return;
  const scale = getAppScale();
  const rawX = (e.clientX - (parseFloat(app.style.left) || 0)) / scale - popupOffX;
  const rawY = (e.clientY - (parseFloat(app.style.top)  || 0)) / scale - popupOffY;
  const { w: canvasW, h: canvasH } = getCanvasSize();
  popup.style.left = Math.min(Math.max(rawX, 0), canvasW - popup.offsetWidth) + 'px';
  popup.style.top  = Math.min(Math.max(rawY, 0), (canvasH - BOTTOM_BAR_H) - popup.offsetHeight) + 'px';
});
window.addEventListener('pointerup',     () => { popupDragId = null; });
window.addEventListener('pointercancel', () => { popupDragId = null; });

/* Feedback popup action button (View My Answer Toggle). `action` falsy
   → the row is removed (retry/correct never carry one). */
function renderPopupAction(popupId, action, onClick) {
  const popup = document.getElementById(popupId);
  if (!popup) return;
  let row = popup.querySelector('.scq-popup-actions');
  if (!action) { if (row) row.remove(); return; }
  if (!row) {
    row = document.createElement('div');
    row.className = 'scq-popup-actions';
    popup.appendChild(row);
  }
  let btn = row.querySelector('.popup-action');
  if (!btn) {
    btn = document.createElement('button');
    btn.type = 'button';
    btn.className = 'popup-action';
    row.appendChild(btn);
  }
  setNavLabel(btn, action.label);
  if (action.view) btn.dataset.view = action.view;
  btn.onclick = onClick || null;
}


/* Escape closes whichever hint overlay is open — never a feedback
   popup, which has no dismissal path at all. */
document.addEventListener('keydown', e => {
  if (e.key !== 'Escape') return;
  document.querySelectorAll('[id$="-hint-overlay"]').forEach(el => {
    if (!el.classList.contains('hidden')) el.classList.add('hidden');
  });
});

/* ═══════════════════════════════════════════════════════════
   GLOBAL COMPONENT — Progress Question nav (generic renderer)
   720-templates → _global-components.md → "Progress Question".

   ⚠ STRUCTURAL CORRECTION, producer 2026-09-13. This part carries THREE
   authored questions, not eight. Each question is split into SECTIONS
   (א / ב / ג) that occupy one screen each:

       Question 1 = S1 (א) · S2 (ב) · S3 (ג)
       Question 2 = S4 (א) · S5 (ב)
       Question 3 = S6 (א) · S7 (ב) · S8 (ג)

   The א/ב/ג labels in the authored content are SECTION labels, not
   question numbers. The earlier implementation read them as eight
   independent questions and carried three separate progress objects
   (practiceProgress / 2 / 3), which advanced the indicator on every
   screen. That is what this model replaces — and it is why screen 0's
   "3 שאלות" was never a contradiction: the learner meets eight screens
   because each question is sectioned.

   Consequences encoded below:
   · the indicator does NOT advance between sections of one question;
   · a question resolves only when its LAST section has resolved;
   · question correctness is the AND of every section
     (חישוב תשובות נכונות במסע למידה.docx — "success in the question
     requires a correct response to every section of that question");
   · mastery counts QUESTIONS, not sections: at least 2 of 3.
   ═══════════════════════════════════════════════════════════ */

/* Which screens make up each authored question, in order. */
const QUESTION_SECTIONS = [[1, 2, 3], [4, 5], [6, 7, 8]];

/* Per-SECTION outcome, 'correct' | 'incorrect' | null (unresolved).
   Sections are the only thing a screen writes; every question-level
   value below is derived, never stored, so the two can never disagree. */
const sectionResult = { 1: null, 2: null, 3: null, 4: null, 5: null, 6: null, 7: null, 8: null };

const practiceProgress = {
  questions: [
    { number: 1, visited: false, state: 'not-answered', screens: [1, 2, 3] },
    { number: 2, visited: false, state: 'not-answered', screens: [4, 5] },
    { number: 3, visited: false, state: 'not-answered', screens: [6, 7, 8] }
  ]
};

function questionIndexOfScreen(screen) {
  return QUESTION_SECTIONS.findIndex(secs => secs.indexOf(screen) !== -1);
}

/* null until EVERY section of the question has an outcome. A partially
   answered question is deliberately not a result — it must not show as
   correct or complete early. */
function questionResult(qi) {
  const secs = QUESTION_SECTIONS[qi];
  if (!secs) return null;
  if (secs.some(sc => sectionResult[sc] === null)) return null;
  return secs.every(sc => sectionResult[sc] === 'correct') ? 'correct' : 'incorrect';
}

/* Recompute every item, then paint the nav on the screen the learner is
   on. A question shows its verdict once resolved; otherwise it is
   'current' if the learner is inside it, and neutral if not. */
function refreshProgress(screen) {
  const cur = questionIndexOfScreen(screen);
  practiceProgress.questions.forEach((q, qi) => {
    const r = questionResult(qi);
    q.state = (r !== null) ? r : (qi === cur ? 'current' : 'not-answered');
  });
  const el = document.getElementById('s' + screen);
  const nav = el && el.querySelector('.progress-question');
  if (nav) updateProgressQuestion(nav, practiceProgress);
}

/* Called on entry to any section screen. */
function enterSection(screen) {
  const qi = questionIndexOfScreen(screen);
  if (qi >= 0) practiceProgress.questions[qi].visited = true;
  refreshProgress(screen);
}

/* Called by a section screen the moment its own SECTION resolves.
   It records the section only — whether the QUESTION has resolved is
   derived by questionResult() above. */
function recordSection(screen, result) {
  sectionResult[screen] = (result === 'correct') ? 'correct' : 'incorrect';
  refreshProgress(screen);
}

/* MASTERY — computed from the three aggregated question results only,
   never from section counts. Exposed for the end-of-part crossing,
   which is NOT wired yet (awaiting the producer's routing decision). */
function masteryState() {
  const results = practiceProgress.questions.map((q, qi) => questionResult(qi));
  const correctQuestionCount = results.filter(r => r === 'correct').length;
  return {
    question1Correct: results[0] === 'correct',
    question2Correct: results[1] === 'correct',
    question3Correct: results[2] === 'correct',
    results,
    correctQuestionCount,
    passed: correctQuestionCount >= 2,
    allResolved: results.every(r => r !== null)
  };
}
function updateProgressQuestion(container, state) {
  state.questions.forEach((q, i) => {
    const n    = i + 1;
    const item = container.querySelector('[data-question="' + n + '"]');
    if (!item) return;
    const icon  = item.querySelector('.progress-question__icon');
    const label = item.querySelector('.progress-question__label');
    icon.classList.remove(
      'progress-question__icon--current',
      'progress-question__icon--correct',
      'progress-question__icon--incorrect'
    );
    if (q.state !== 'not-answered') icon.classList.add('progress-question__icon--' + q.state);
    label.classList.toggle('progress-question__label--visited', q.visited);
    const target = q.screens[0];
    const navigable = q.visited && target != null && target !== currentScreen;
    item.style.cursor = navigable ? 'pointer' : '';
    item.onclick = navigable ? (() => goTo(target)) : null;
  });
  for (let n = 1; n < state.questions.length; n++) {
    const conn = container.querySelector('[data-connector="' + n + '"]');
    if (!conn) continue;
    const qState = state.questions[n - 1].state;
    conn.classList.toggle('progress-question__connector--visited', qState === 'correct' || qState === 'incorrect');
  }
}

/* ═══════════════════════════════════════════════════════════
   S1 — ValueInputQuestion, LOCAL RATIO VARIANT (slide 30)
   Practice Q1 of 3. Spec: 720-templates → ValueInputQuestion.md
   (Variant B contract: one answer object, one combined verdict) +
   _question-template-defaults.md + _global-components.md → "View My
   Answer Toggle". See the css/style.css header block above this
   screen's rules for the full source-evidence note (hint policy,
   attempt model, diagram geometry).

   CORRECT VALUES — script saved 2026-09-29 17:26 (slide at deck position
   30 of "יעד 1.3_יחס וביטויים אלגבריים_להפקה_06092026.pptx"). The question
   asks for LENGTH to WIDTH: 2a : 3a, simplified 2 : 3. The slide draws
   "2" in the LEFT box and "3" in the RIGHT box, read LEFT TO RIGHT like
   S3 and parts 03/04: LEFT field = 2 (length, first term), RIGHT field =
   3 (width, second term). The explanation is verbatim from the slide
   ("2a : 3a"; "נקבל 3 : 2 ." in an RTL paragraph, which renders 2 on the
   left).
   The explanation sentence was aligned in the script saved 18:22
   ("النسبة بين طول المستطيل وعرضه هي"); the presenter note still says
   "3 : 2" and is not rendered.

   History (all 2026-09-29): the first deck asked width-to-length with
   "3" in the right box; a tester reading left to right typed 3 : 2 and
   was marked wrong. The script was then revised twice (width-to-length
   with 3 on the left, then this length-to-width version).
   S3 (deck position 32): "4" left, "3" right, also read left to right.
   ═══════════════════════════════════════════════════════════ */
const S1VIQ = { correctRight: 3, correctLeft: 2, maxAttempts: 2 };

let s1viqAttempts    = 0;
let s1viqDone        = false;
let s1viqResult      = null;   // 'correct' | 'wrong2'
let s1viqAnswer      = null;   // { right, left } raw typed strings, snapshotted before reveal
let s1viqView        = 'solution';
let s1viqLastWrongSig = null;

/* Shared explanation — three authored paragraphs, verbatim from the
   updated script (2026-09-29). Correct and final-wrong popups
   carry byte-identical bodies in the source and differ only in title
   (_question-template-defaults.md → "Feedback: shared explanation"). */
const S1VIQ_EXPLAIN = [
  'النسبة بين طول المستطيل وعرضه هي',
  '<bdi dir="ltr" class="viq-expr">2a : 3a</bdi>',
  'وبعد الاختزال بـ a، الذي لا يساوي 0، نحصل على <span class="viq-expr">3 : 2</span>.'
];
const S1VIQ_POPUP = {
  /* No retry-specific pedagogical explanation is authored for this
     question — only the canonical fallback title is used here, per
     _question-template-defaults.md → "Feedback: shared explanation":
     "A missing pedagogical explanation must not be invented." */
  retry:   { bg: '#ffdbdc', title: 'الإجابة غير صحيحة.', body: ['هل نحاول مرة أخرى؟'] },
  correct: { bg: '#edf8ed', title: 'إجابتكم صحيحة!', body: S1VIQ_EXPLAIN },
  wrong2: {
    bg: '#ffdbdc', title: 'هذا خطأ. لا تقلقوا، فالفهم يتم تدريجيًا‎.', body: S1VIQ_EXPLAIN,
    action: { label: 'الانتقال إلى إجابتي', view: 'answer' },
    answerView: {
      title: 'الإجابة غير صحيحة.',
      body:  ['يمكن مشاهدة الحل الصحيح مرة أخرى‎.'],
      action: { label: 'الإجابة الصحيحة', view: 'solution' }
    }
  }
};

function s1viqInput(k) { return document.getElementById('s1viq-input-' + k); }
function s1viqField(k) { return document.getElementById('s1viq-field-' + k); }
function s1viqNormalize(raw) { return String(raw).trim().replace(/,/g, ''); }
function s1viqSignature() {
  return s1viqNormalize(s1viqInput('right').value) + '|' + s1viqNormalize(s1viqInput('left').value);
}
function s1viqAllFilled() {
  return s1viqNormalize(s1viqInput('right').value) !== '' && s1viqNormalize(s1viqInput('left').value) !== '';
}
function s1viqIsCorrect() {
  return Number(s1viqNormalize(s1viqInput('right').value)) === S1VIQ.correctRight &&
         Number(s1viqNormalize(s1viqInput('left').value))  === S1VIQ.correctLeft;
}

/* type="text" removes the browser's own numeric enforcement, so every
   keystroke AND paste is filtered here (a paste fires `input` too, so
   no separate paste handler is needed) — ValueInputQuestion.md →
   "Numeric input format". */
function s1viqOnInput() {
  if (s1viqDone) return;
  ['right', 'left'].forEach(k => {
    const el = s1viqInput(k);
    const cleaned = el.value.replace(/[^0-9,.-]/g, '');
    if (cleaned !== el.value) el.value = cleaned;
  });
  /* Starting a new attempt dismisses the previous retry popup —
     approved path 2 of "No manual close". Unreachable once answered. */
  if (s1viqAttempts > 0) {
    document.getElementById('s1viq-popup').classList.add('hidden');
    s1viqMark(null);
  }
  s1viqUpdateBar();
}

function s1viqUpdateBar() {
  const chk = document.getElementById('s1viq-check');
  if (!chk || s1viqDone) return;
  /* Retry gate on the FULL joined signature (_question-template-
     defaults.md → "Retry gate"): resubmitting with one field unchanged
     and the other identical is still "the same wrong answer". */
  chk.disabled = !s1viqAllFilled() || s1viqSignature() === s1viqLastWrongSig;
}

/* Case B (one semantic answer object — ValueInputQuestion.md → "Scoring
   vs feedback granularity" → Case B): BOTH fields always share the SAME
   single outcome, and there is exactly one verdict icon for the whole
   group — never a per-field icon. */
function s1viqMark(outcome) {
  const group = document.getElementById('s1viq-ratio-group');
  const ico   = document.getElementById('s1viq-ico');
  ['right', 'left'].forEach(k => {
    const f = s1viqField(k);
    f.classList.remove('viq-correct', 'viq-wrong');
    if (outcome) f.classList.add(outcome === 'correct' ? 'viq-correct' : 'viq-wrong');
  });
  if (group) group.classList.toggle('viq-has-icon', !!outcome);
  if (ico) {
    ico.innerHTML = !outcome ? '' : (outcome === 'correct' ? S1VIQ_ICON_CORRECT : S1VIQ_ICON_WRONG);
  }
}
/* Type 1 icon — 16×16 bare vector, canonical production SVGs
   (720-design-guidelines → components-icons-states.md → "Type 1"). */
const S1VIQ_ICON_CORRECT = '<svg width="16" height="16" viewBox="0 0 14 12" fill="none" aria-hidden="true"><path d="M4.52036 12C4.12296 12.0002 3.74183 11.8126 3.46106 11.4787L0.258452 7.67744C-0.0861507 7.26814 -0.0861507 6.60467 0.258452 6.19536C0.603165 5.78619 1.16194 5.78619 1.50666 6.19536L4.52036 9.77374L12.4933 0.306878C12.8381 -0.102293 13.3968 -0.102293 13.7415 0.306878C14.0862 0.716179 14.0862 1.37966 13.7415 1.78896L5.57966 11.4787C5.29889 11.8126 4.91776 12.0002 4.52036 12Z" fill="#609E12"/></svg>';
const S1VIQ_ICON_WRONG   = '<svg width="16" height="16" viewBox="0 0 14 14" fill="none" aria-hidden="true"><path d="M13.6743 0.325691C13.4657 0.117151 13.1828 0 12.8879 0C12.5929 0 12.31 0.117151 12.1014 0.325691L7 5.42709L1.8986 0.325691C1.69 0.117151 1.40711 0 1.11215 0C0.817183 0 0.534294 0.117151 0.325691 0.325691C0.117151 0.534294 0 0.817183 0 1.11215C0 1.40711 0.117151 1.69 0.325691 1.8986L5.42709 7L0.325691 12.1014C0.117151 12.31 0 12.5929 0 12.8879C0 13.1828 0.117151 13.4657 0.325691 13.6743C0.534294 13.8829 0.817183 14 1.11215 14C1.40711 14 1.69 13.8829 1.8986 13.6743L7 8.57291L12.1014 13.6743C12.31 13.8829 12.5929 14 12.8879 14C13.1828 14 13.4657 13.8829 13.6743 13.6743C13.8829 13.4657 14 13.1828 14 12.8879C14 12.5929 13.8829 12.31 13.6743 12.1014L8.57291 7L13.6743 1.8986C13.8829 1.69 14 1.40711 14 1.11215C14 0.817183 13.8829 0.534294 13.6743 0.325691Z" fill="#B20010"/></svg>';

function s1viqCheck() {
  if (s1viqDone) { advanceScreen(); return; }
  if (!s1viqAllFilled()) return;

  s1viqAttempts++;
  const correct = s1viqIsCorrect();
  reportQ(['001', 'q1'], correct, correct || s1viqAttempts >= S1VIQ.maxAttempts, String((document.getElementById('s1viq-input-left') || {}).value).trim() + ' : ' + String((document.getElementById('s1viq-input-right') || {}).value).trim());

  if (correct) {
    s1viqMark('correct');
    s1viqResult = 'correct';
    s1viqFinish();
    s1viqShowPopup('correct');
    return;
  }

  if (s1viqAttempts < S1VIQ.maxAttempts) {
    s1viqMark('wrong');
    s1viqLastWrongSig = s1viqSignature();
    document.getElementById('s1viq-check').disabled = true;
    s1viqShowPopup('retry');
    return;
  }

  /* Final wrong. The reveal overwrites the very inputs the learner
     typed into, so snapshot their RAW strings first — the only place
     these survive. Preserved exactly as typed, never reformatted. */
  s1viqAnswer = { right: s1viqInput('right').value, left: s1viqInput('left').value };
  s1viqResult = 'wrong2';
  s1viqView   = 'solution';
  s1viqFinish();
  s1viqApplyView();
}

/* Single owner of the resolved-wrong rendering. The toggle is OBJECT-
   level: one button rewrites BOTH fields (and the single combined
   verdict) in the same pass — this is Case B, not Case A. */
function s1viqApplyView() {
  if (s1viqAnswer === null) return;
  if (s1viqView === 'answer') {
    s1viqInput('right').value = s1viqAnswer.right;
    s1viqInput('left').value  = s1viqAnswer.left;
    s1viqMark('wrong');   // the final attempt was wrong in both views by definition
  } else {
    s1viqInput('right').value = String(S1VIQ.correctRight);
    s1viqInput('left').value  = String(S1VIQ.correctLeft);
    s1viqMark('correct');
  }
  s1viqShowPopup('wrong2', s1viqView);
}
/* Pure view switch — NEVER routes through s1viqCheck(), never touches
   attempts/done/scoring. */
function s1viqToggleView() {
  if (s1viqAnswer === null) return;
  s1viqView = (s1viqView === 'solution') ? 'answer' : 'solution';
  s1viqApplyView();
}

function s1viqFinish() {
  s1viqDone = true;
  s1viqInput('right').disabled = true;
  s1viqInput('left').disabled  = true;
  const chk = document.getElementById('s1viq-check');
  /* ValueInputQuestion's own completion label is "متابعة" — NOT
     SingleChoiceQuestion's "هل نتابع؟"; the template doc is explicit the
     two templates do not share one. */
  if (chk) { setNavLabel(chk, 'متابعة'); chk.disabled = false; }
  recordSection(1, s1viqResult);   /* Question 1, section א */
}

function s1viqResetInitial() {
  s1viqAttempts     = 0;
  s1viqDone         = false;
  s1viqResult       = null;
  s1viqAnswer       = null;
  s1viqView         = 'solution';
  s1viqLastWrongSig = null;
  ['right', 'left'].forEach(k => {
    const el = s1viqInput(k);
    el.value = '';
    el.disabled = false;
  });
  s1viqMark(null);
  document.getElementById('s1viq-popup').classList.add('hidden');
  document.getElementById('s1viq-hint-overlay').classList.add('hidden');
  const chk = document.getElementById('s1viq-check');
  if (chk) { setNavLabel(chk, 'هل إجابتي صحيحة؟'); chk.disabled = true; }
  sectionResult[1] = null;   /* this section is unanswered again */
  /* Hint visible from the start — canonical template default, and this
     slide's own hint trigger sits in the base on-canvas layer (see the
     source-evidence note in css/style.css). */
}

function s1viqShowPopup(type, view) {
  const popup = document.getElementById('s1viq-popup');
  if (!popup) return;
  const base = S1VIQ_POPUP[type];
  const cfg  = (view === 'answer' && base.answerView)
             ? Object.assign({}, base, base.answerView)
             : base;
  popup.style.background = cfg.bg;
  /* Re-anchor on a fresh open, AND on every 'wrong2' render (including
     each toggle) even if already open — a height change between the
     long solution body and the short answer body must never leave the
     popup pinned away from the bottom bar. */
  if (popup.classList.contains('hidden') || type === 'wrong2') resetPopupPosition(popup);
  const titleEl = document.getElementById('s1viq-popup-title');
  /* Terminal-incorrect SOLUTION view carries a fixed second line
     (_question-template-defaults.md → "Terminal-incorrect title carries
     a fixed second line") — appended, never replacing the authored
     title. Does not apply to the answerView (shows the learner's own
     answer, not the solution) or to retry/correct. */
  titleEl.innerHTML = (type === 'wrong2' && view !== 'answer')
    ? cfg.title + '<br>الإجابة الصحيحة تَظهَر الآن.'
    : cfg.title;
  document.getElementById('s1viq-popup-body').innerHTML =
    (cfg.body || []).map(p => '<p>' + p + '</p>').join('');
  renderPopupAction('s1viq-popup', cfg.action, s1viqToggleView);
  popup.classList.remove('hidden');
  clampPopup(popup);
}

function s1viqOpenHint() {
  xapiRequestedHint('001', 'q1');
  document.getElementById('s1viq-hint-overlay').classList.remove('hidden');
}
function s1viqCloseHint() {
  document.getElementById('s1viq-hint-overlay').classList.add('hidden');
}
function s1viqCloseHintOnBackdrop(e) {
  if (e.target === e.currentTarget) s1viqCloseHint();
}

/* ═══════════════════════════════════════════════════════════
   S2 — SingleChoiceQuestion (slide 31)
   Practice Q2 of 3. Spec: 720-templates → SingleChoiceQuestion.md +
   _question-template-defaults.md. Shell reused from S1 (progress nav,
   feedback popup, hint overlay, bottom bar); the interaction itself is
   this template's own single-select pill contract — no View My Answer
   Toggle (SCQ shows the chosen and correct options side by side
   already, which the toggle's own applicability rule excludes).

   FOUR options, o1..o4 in the deck's own top→bottom authored order:
   o1=6a² (correct), o2=3a², o3=a², o4=36a — see index.html's own
   source-evidence comment above #s2 for the full option-by-option OMML
   citation, the extraction-miss root cause for o2 (a copy-pasted shape
   group sharing its PowerPoint auto-name with o3's, corrected after
   user QA), and the safeguard this build now applies.

   CORRECT = 'o1' (6a²), confirmed on THREE independent pieces of
   evidence: its container is the only one of the four with a GREEN
   border in the source (00B050 vs the others' blue 00B0F0); a separate
   solid-green-filled ellipse sits at its radio control's exact
   position as an explicit correctness marker; and it is the
   mathematically correct area — width 3a × length 2a = 6a², confirmed
   by this slide's own OMML calculation ("3𝑎⋅2𝑎=3⋅2⋅𝑎⋅𝑎=6𝑎²"), read
   directly from the OMML run tree (not from flat-text bidi order).
   Never resolved from a letter key.

   Shared explanation, verbatim from the slide (both feedback popups
   carry byte-identical bodies, differing only in title — the deck's
   own normal 2-attempt shape). The plain-text "3a"/"2a" values inside
   it were resolved by cross-checking this same slide's own unambiguous
   single-run stem text ("2a"/"3a") and its own OMML calculation
   (opens "3𝑎⋅2𝑎…") — both agree — rather than read off the raw XML
   run order ("a" then "3"/"2"), which pptx-script-extraction flags as
   bidi-sensitive and not reliable reading-order evidence on its own.
   This is a same-slide cross-check, not a DOM-order-based reordering
   of any ratio's own terms. */
const S2SC_EXPLAIN = [
  'عرض المستطيل هو <span dir="ltr">3a</span>.',
  'طول المستطيل هو <span dir="ltr">2a</span>.',
  'لذلك نحسب مساحة المستطيل على النحو التالي:',
  '<span dir="ltr">3a &sdot; 2a = 3 &sdot; 2 &sdot; a &sdot; a = 6a<sup>2</sup></span>',
  'أي إن مساحة المستطيل هي <span dir="ltr">6a<sup>2</sup></span>.'
];
const S2SC = { correctId: 'o1', maxAttempts: 2 };
const S2SC_POPUP = {
  retry:   { bg: '#ffdbdc', title: 'الإجابة غير صحيحة.', body: ['هل نحاول مرة أخرى؟'] },
  correct: { bg: '#edf8ed', title: 'ممتاز‎!', body: S2SC_EXPLAIN },
  wrong2:  { bg: '#ffdbdc', title: 'هذا خطأ، لا بأس!، لأننا نستطيع الآن تعلّم شيء جديد‎.', body: S2SC_EXPLAIN }
};

let s2scSelected  = null;
let s2scAttempts  = 0;
let s2scDone      = false;
let s2scResult    = null;   // 'correct' | 'wrong2'
let s2scLastWrong = null;   // the option id just marked wrong — retry gate compares against THIS

function s2scOpt(id) { return document.querySelector('#s2 .scq-opt[data-id="' + id + '"]'); }
function s2scOpts()  { return document.querySelectorAll('#s2 .scq-opt'); }

/* Single select; clears a prior wrong attempt's marks; re-enables check
   only if the newly selected id differs from the one just marked wrong
   (_question-template-defaults.md → "Retry gate" — comparison shape:
   the selected option id). */
function s2scSelect(id) {
  if (s2scDone) return;
  s2scOpts().forEach(o => {
    o.classList.remove('selected');
    o.setAttribute('aria-checked', 'false');
  });
  const opt = s2scOpt(id);
  opt.classList.add('selected');
  opt.setAttribute('aria-checked', 'true');
  s2scSelected = id;

  if (s2scAttempts > 0) {
    document.getElementById('s2sc-popup').classList.add('hidden');
    s2scOpts().forEach(o => o.classList.remove('correct', 'wrong'));
  }
  const chk = document.getElementById('s2sc-check');
  if (chk) chk.disabled = (id === s2scLastWrong);
}

function s2scCheck() {
  if (s2scDone) { advanceScreen(); return; }
  if (!s2scSelected) return;

  s2scAttempts++;
  const correct = s2scSelected === S2SC.correctId;
  reportQ(['001', 'q2'], correct, correct || s2scAttempts >= S2SC.maxAttempts, xapiChoiceAnswer('001', 'q2', '#s2 .scq-opt', s2scSelected));

  if (correct) {
    s2scOpt(s2scSelected).classList.add('correct');
    s2scResult = 'correct';
    s2scFinish();
    s2scShowPopup('correct');
    return;
  }

  if (s2scAttempts < S2SC.maxAttempts) {
    /* Non-final wrong: mark the chosen option wrong; the correct one is
       NOT revealed yet (SingleChoiceQuestion.md → state 4, "Try Again"). */
    s2scOpt(s2scSelected).classList.add('wrong');
    s2scLastWrong = s2scSelected;
    document.getElementById('s2sc-check').disabled = true;
    s2scShowPopup('retry');
    return;
  }

  /* Final wrong: chosen option red+✗, correct option green+✓ — both
     shown at once (no toggle needed; this is the whole point of the
     View My Answer Toggle's applicability exclusion for SCQ). */
  s2scOpt(s2scSelected).classList.add('wrong');
  s2scOpt(S2SC.correctId).classList.add('correct');
  s2scResult = 'wrong2';
  s2scFinish();
  s2scShowPopup('wrong2');
}

function s2scFinish() {
  s2scDone = true;
  s2scOpts().forEach(o => { o.disabled = true; });
  const chk = document.getElementById('s2sc-check');
  /* SingleChoiceQuestion's own completion label is "هل نتابع؟" — NOT
     ValueInputQuestion's "متابعة"; the two templates do not share one. */
  if (chk) { setNavLabel(chk, 'هل نتابع؟'); chk.disabled = false; }
  hideHintButton('s2sc-hint');
  recordSection(2, s2scResult);    /* Question 1, section ב */
}

function s2scResetInitial() {
  s2scSelected  = null;
  s2scAttempts  = 0;
  s2scDone      = false;
  s2scResult    = null;
  s2scLastWrong = null;
  s2scOpts().forEach(o => {
    o.classList.remove('selected', 'correct', 'wrong');
    o.setAttribute('aria-checked', 'false');
    o.disabled = false;
  });
  document.getElementById('s2sc-popup').classList.add('hidden');
  document.getElementById('s2sc-hint-overlay').classList.add('hidden');
  const chk = document.getElementById('s2sc-check');
  if (chk) { setNavLabel(chk, 'هل إجابتي صحيحة؟'); chk.disabled = true; }
  showHintButton('s2sc-hint');   // visible from the start — see the source-evidence note in index.html
  sectionResult[2] = null;
}

function s2scShowPopup(type) {
  const popup = document.getElementById('s2sc-popup');
  if (!popup) return;
  const cfg = S2SC_POPUP[type];
  popup.style.background = cfg.bg;
  if (popup.classList.contains('hidden')) resetPopupPosition(popup);
  /* Terminal-incorrect title carries the fixed second line
     (_question-template-defaults.md), as on S5. */
  document.getElementById('s2sc-popup-title').innerHTML = (type === 'wrong2')
    ? cfg.title + '<br>الإجابة الصحيحة تَظهَر الآن.'
    : cfg.title;
  document.getElementById('s2sc-popup-body').innerHTML =
    cfg.body.map(p => '<p>' + p + '</p>').join('');
  renderPopupAction('s2sc-popup', null, null);   // no action row on this template
  popup.classList.remove('hidden');
  clampPopup(popup);
}

function s2scOpenHint() {
  xapiRequestedHint('001', 'q2');
  document.getElementById('s2sc-hint-overlay').classList.remove('hidden');
}
function s2scCloseHint() {
  document.getElementById('s2sc-hint-overlay').classList.add('hidden');
}
function s2scCloseHintOnBackdrop(e) {
  if (e.target === e.currentTarget) s2scCloseHint();
}

/* Enter/Space on a focused option selects it — radiogroup keyboard a11y,
   same pattern as Screen 01's (03-01) TwoOptionSelection. Delegated so
   it covers every .scq-opt without per-node inline handlers. */
document.addEventListener('keydown', e => {
  const isEnter = e.key === 'Enter' || e.code === 'Enter' || e.code === 'NumpadEnter';
  const isSpace = e.key === ' ' || e.key === 'Spacebar' || e.code === 'Space';
  if (!isEnter && !isSpace) return;
  const opt = e.target.closest && e.target.closest('#s2 .scq-opt');
  if (!opt || opt.disabled) return;
  e.preventDefault();
  s2scSelect(opt.dataset.id);
});

/* ═══════════════════════════════════════════════════════════
   S3 — ValueInputQuestion, LOCAL RATIO VARIANT (slide 32)
   Practice Q3 of 3 — the final question in this part's series. Same
   local variant as S1 (see the S1VIQ block's own header note for the
   full mechanism/rationale — not restated here); only the config,
   content and diagram differ.

   CORRECT VALUES: length was doubled (2a → 2a·2 = 4a), width stays 3a;
   new ratio length:width = 4a:3a → simplified 4:3 — confirmed on TWO
   independent pieces of evidence: notesSlide29's explicit "התשובה
   הנכונה היא 4 : 3", and this slide's own two feedback popups, whose
   last authored line is "...נקבל את היחס החדש 4 : 3 .".

   UPDATED SCRIPT, 2026-09-29 (deck position 32): the slide now draws
   "4" in the LEFT box and "3" in the right, read LEFT TO RIGHT like S1
   and parts 03/04: LEFT field = 4 (new length, first term), RIGHT field
   = 3 (width, second term). Its explanation now stores "4a : 3a" and a
   reduced "3 : 4" that renders "4 : 3"; the isolates below carry that
   displayed order. (Earlier deck: "4" in the right box, "3a : 4a" /
   "3 : 4" displayed, consistent only when read right to left.)

   ⚠ WITHDRAWN FLAG (2026-09-09): an earlier revision of this block
   reported a "source inconsistency" between the popup's "3a : 4a"
   sentence and its own next line. That was MY error, not the deck's —
   the two lines agree once each is read in the direction PowerPoint
   actually renders it. See the S3VIQ_EXPLAIN comment below for the
   full bidi explanation. Nothing in the source needs producer
   correction on this point. */
const S3VIQ = { correctRight: 3, correctLeft: 4, maxAttempts: 2 };

let s3viqAttempts     = 0;
let s3viqDone         = false;
let s3viqResult       = null;   // 'correct' | 'wrong2'
let s3viqAnswer       = null;   // { right, left } raw typed strings, snapshotted before reveal
let s3viqView         = 'solution';
let s3viqLastWrongSig = null;

/* Shared explanation — one body for correct and wrong2, differing only
   in title (_question-template-defaults.md → "Feedback: shared
   explanation": the deck's own normal 2-attempt shape).

   FIVE authored paragraphs = the five semantic groups (heading, the
   doubling + its complete calculation, the unchanged width, the
   4a-to-3a ratio statement, the reduction). Slide 32 contains ZERO
   <a:br> elements (verified), so NO in-sentence break is authored:
   every line break visible inside a sentence in the PPT is automatic
   wrapping from its own textbox width, and is therefore NOT
   reproduced with <br> here (_question-template-defaults.md →
   "Automatic wrapping is not an authored instruction"). The four
   genuinely EMPTY paragraphs between the content ones ARE authored
   group breaks and are preserved as a real blank-line gap — see the
   #s3viq-popup rule in css/style.css.

   EXPRESSION ISOLATION: each complete expression sits in its own
   `<bdi dir="ltr" class="viq-expr">` — an LTR isolate so the
   surrounding RTL paragraph can never reorder its own tokens, plus
   `white-space: nowrap` so no number, variable, colon or operator can
   ever be split onto a line of its own. Each trailing period is glued
   directly to the closing tag (no whitespace), so it cannot be
   orphaned either.

   ⚠ REDUCED-RATIO ORDER — corrected 2026-09-09, and the earlier
   "source inconsistency" flag on this screen is WITHDRAWN as my own
   error. The deck's last paragraph stores the reduced ratio as the
   logical run "4 : 3" inside a Hebrew (RTL) paragraph. Bare digits are
   bidi-neutral-numeric, so the RTL paragraph reorders that group and
   PowerPoint RENDERS it as "3 : 4" — while "3a : 4a" one line above
   contains Latin letters (strong LTR), so it renders in its stored
   order "3a : 4a". Read right-to-left, as a Hebrew reader reads them,
   those two displayed forms say 4a:3a and 4:3 — length:width, exactly
   consistent with each other and with the presenter note. The source
   was never contradictory; the earlier build forced the reduced ratio
   through an LTR isolate while keeping its stored token order, which
   displayed "4 : 3" and broke that agreement. The isolate content
   below is therefore the DISPLAY order the deck actually renders
   ("3 : 4"), and "3a : 4a" is left exactly as it was.
   SUPERSEDED 2026-09-29: the updated deck displays "4a : 3a" and
   "4 : 3" (read left to right); the isolates carry those, same
   display-order principle — see UPDATED SCRIPT in the S3VIQ header. */
const S3VIQ_EXPLAIN = [
  'تم مضاعفة طول المستطيل بمرتين، ولذلك أصبح طول المستطيل الجديد <bdi dir="ltr" class="viq-expr">2a &sdot; 2 = 4a</bdi>.',
  'عرض المستطيل بقيَ <bdi dir="ltr" class="viq-expr">3a</bdi>.',
  'ولذلك فإن النسبة بين طول المستطيل الجديد وعرضه هي <bdi dir="ltr" class="viq-expr">4a : 3a</bdi>.',
  'بعد الاختزال بـ <bdi dir="ltr" class="viq-expr">a</bdi> الذي لا يساوي 0، نحصل على النسبة الجديدة <bdi dir="ltr" class="viq-expr">4 : 3</bdi>.'
];
const S3VIQ_POPUP = {
  retry:   { bg: '#ffdbdc', title: 'الإجابة غير صحيحة.', body: ['هل نحاول مرة أخرى؟'] },
  correct: { bg: '#edf8ed', title: 'إجابتكم صحيحة!', body: S3VIQ_EXPLAIN },
  wrong2: {
    bg: '#ffdbdc', title: 'هذا خطأ. لا تقلقوا، فالفهم يتم تدريجيًا‎.', body: S3VIQ_EXPLAIN,
    action: { label: 'الانتقال إلى إجابتي', view: 'answer' },
    answerView: {
      title: 'الإجابة غير صحيحة.',
      body:  ['يمكن مشاهدة الحل الصحيح مرة أخرى‎.'],
      action: { label: 'الإجابة الصحيحة', view: 'solution' }
    }
  }
};

function s3viqInput(k) { return document.getElementById('s3viq-input-' + k); }
function s3viqField(k) { return document.getElementById('s3viq-field-' + k); }
function s3viqNormalize(raw) { return String(raw).trim().replace(/,/g, ''); }
function s3viqSignature() {
  return s3viqNormalize(s3viqInput('right').value) + '|' + s3viqNormalize(s3viqInput('left').value);
}
function s3viqAllFilled() {
  return s3viqNormalize(s3viqInput('right').value) !== '' && s3viqNormalize(s3viqInput('left').value) !== '';
}
function s3viqIsCorrect() {
  return Number(s3viqNormalize(s3viqInput('right').value)) === S3VIQ.correctRight &&
         Number(s3viqNormalize(s3viqInput('left').value))  === S3VIQ.correctLeft;
}

function s3viqOnInput() {
  if (s3viqDone) return;
  ['right', 'left'].forEach(k => {
    const el = s3viqInput(k);
    const cleaned = el.value.replace(/[^0-9,.-]/g, '');
    if (cleaned !== el.value) el.value = cleaned;
  });
  if (s3viqAttempts > 0) {
    document.getElementById('s3viq-popup').classList.add('hidden');
    s3viqMark(null);
  }
  s3viqUpdateBar();
}

function s3viqUpdateBar() {
  const chk = document.getElementById('s3viq-check');
  if (!chk || s3viqDone) return;
  chk.disabled = !s3viqAllFilled() || s3viqSignature() === s3viqLastWrongSig;
}

function s3viqMark(outcome) {
  const group = document.getElementById('s3viq-ratio-group');
  const ico   = document.getElementById('s3viq-ico');
  ['right', 'left'].forEach(k => {
    const f = s3viqField(k);
    f.classList.remove('viq-correct', 'viq-wrong');
    if (outcome) f.classList.add(outcome === 'correct' ? 'viq-correct' : 'viq-wrong');
  });
  if (group) group.classList.toggle('viq-has-icon', !!outcome);
  if (ico) {
    ico.innerHTML = !outcome ? '' : (outcome === 'correct' ? S1VIQ_ICON_CORRECT : S1VIQ_ICON_WRONG);
  }
}

function s3viqCheck() {
  if (s3viqDone) { advanceScreen(); return; }
  if (!s3viqAllFilled()) return;

  s3viqAttempts++;
  const correct = s3viqIsCorrect();
  reportQ(['001', 'q3'], correct, correct || s3viqAttempts >= S3VIQ.maxAttempts, String((document.getElementById('s3viq-input-left') || {}).value).trim() + ' : ' + String((document.getElementById('s3viq-input-right') || {}).value).trim());

  if (correct) {
    s3viqMark('correct');
    s3viqResult = 'correct';
    s3viqFinish();
    s3viqShowPopup('correct');
    return;
  }

  if (s3viqAttempts < S3VIQ.maxAttempts) {
    s3viqMark('wrong');
    s3viqLastWrongSig = s3viqSignature();
    document.getElementById('s3viq-check').disabled = true;
    s3viqShowPopup('retry');
    return;
  }

  s3viqAnswer = { right: s3viqInput('right').value, left: s3viqInput('left').value };
  s3viqResult = 'wrong2';
  s3viqView   = 'solution';
  s3viqFinish();
  s3viqApplyView();
}

function s3viqApplyView() {
  if (s3viqAnswer === null) return;
  if (s3viqView === 'answer') {
    s3viqInput('right').value = s3viqAnswer.right;
    s3viqInput('left').value  = s3viqAnswer.left;
    s3viqMark('wrong');
  } else {
    s3viqInput('right').value = String(S3VIQ.correctRight);
    s3viqInput('left').value  = String(S3VIQ.correctLeft);
    s3viqMark('correct');
  }
  s3viqShowPopup('wrong2', s3viqView);
}
function s3viqToggleView() {
  if (s3viqAnswer === null) return;
  s3viqView = (s3viqView === 'solution') ? 'answer' : 'solution';
  s3viqApplyView();
}

function s3viqFinish() {
  s3viqDone = true;
  s3viqInput('right').disabled = true;
  s3viqInput('left').disabled  = true;
  const chk = document.getElementById('s3viq-check');
  if (chk) { setNavLabel(chk, 'متابعة'); chk.disabled = false; }
  recordSection(3, s3viqResult);   /* Question 1, section ג */
}

function s3viqResetInitial() {
  s3viqAttempts     = 0;
  s3viqDone         = false;
  s3viqResult       = null;
  s3viqAnswer       = null;
  s3viqView         = 'solution';
  s3viqLastWrongSig = null;
  ['right', 'left'].forEach(k => {
    const el = s3viqInput(k);
    el.value = '';
    el.disabled = false;
  });
  s3viqMark(null);
  document.getElementById('s3viq-popup').classList.add('hidden');
  document.getElementById('s3viq-hint-overlay').classList.add('hidden');
  const chk = document.getElementById('s3viq-check');
  if (chk) { setNavLabel(chk, 'هل إجابتي صحيحة؟'); chk.disabled = true; }
  sectionResult[3] = null;
  /* Hint visible from the start — same source-evidence basis as S1/S2
     (the hint trigger sits in this slide's base on-canvas layer too). */
}

function s3viqShowPopup(type, view) {
  const popup = document.getElementById('s3viq-popup');
  if (!popup) return;
  const base = S3VIQ_POPUP[type];
  const cfg  = (view === 'answer' && base.answerView)
             ? Object.assign({}, base, base.answerView)
             : base;
  popup.style.background = cfg.bg;
  if (popup.classList.contains('hidden') || type === 'wrong2') resetPopupPosition(popup);
  const titleEl = document.getElementById('s3viq-popup-title');
  titleEl.innerHTML = (type === 'wrong2' && view !== 'answer')
    ? cfg.title + '<br>الإجابة الصحيحة تَظهَر الآن.'
    : cfg.title;
  document.getElementById('s3viq-popup-body').innerHTML =
    (cfg.body || []).map(p => '<p>' + p + '</p>').join('');
  renderPopupAction('s3viq-popup', cfg.action, s3viqToggleView);
  popup.classList.remove('hidden');
  clampPopup(popup);
}

function s3viqOpenHint() {
  xapiRequestedHint('001', 'q3');
  document.getElementById('s3viq-hint-overlay').classList.remove('hidden');
}
function s3viqCloseHint() {
  document.getElementById('s3viq-hint-overlay').classList.add('hidden');
}
function s3viqCloseHintOnBackdrop(e) {
  if (e.target === e.currentTarget) s3viqCloseHint();
}

/* ═══════════════════════════════════════════════════════════
   S4 — ValueInputQuestion, VARIANT C (multi-field, inline-in-prose)
   slide 33 · Q1 of this part's SECOND question series.

   Spec: 720-templates → ValueInputQuestion.md → "Markup contract —
   Variant C: multi-field holistic" + "Scoring vs feedback granularity"
   → CASE A, plus _question-template-defaults.md and
   _global-components.md → "View My Answer Toggle" / "Progress
   Question". Closest approved implementation of this exact variant and
   sub-shape: methodica-math-percent-02-07 S3 (two inline blanks, two
   sentences) — its JS shape is reused here rather than reinvented.

   CASE A, NOT Case B. S1 and S3 in this part are Case B (a ratio is
   ONE answer object → both boxes carry the same colour and there is
   exactly one combined verdict icon). S4 is the opposite: two
   INDEPENDENT answers (the plain distance and the uphill distance), so
   the question scores as one unit — correct only if BOTH are right —
   while each field is MARKED with its own ✓/✕. A half-right attempt
   therefore renders one green field and one red field and still scores
   incorrect. ("Not scored per field" is not "not marked per field.")

   CORRECT VALUES: plain (במישור) = 15, uphill (בעלייה) = 6 — three
   independent pieces of evidence, quoted in full in the source-evidence
   comment above this screen in index.html: notesSlide30's explicit
   "התשובות הנכונות הן 15 במישור, 6 בעלייה"; the two answer boxes' own
   placeholder digits (id=7 upper = "15", id=18 lower = "6", matching
   the upper/lower sentences); and the feedback's own algebra
   (5x + 2x = 21 → x = 3 → 5·3 = 15, 2·3 = 6), which also reduces to
   the stem's displayed ratio 5 : 2. */
const S4VIQ_KEYS = ['plain', 'uphill'];
const S4VIQ = { correct: { plain: 15, uphill: 6 }, maxAttempts: 2 };

let s4viqAttempts     = 0;
let s4viqDone         = false;
let s4viqResult       = null;   // 'correct' | 'wrong2'
let s4viqAnswer       = null;   // { plain, uphill } raw typed strings, snapshotted before the reveal
let s4viqView         = 'solution';
let s4viqLastWrongSig = null;

/* Shared explanation — one body for correct and wrong2, differing only
   in title (_question-template-defaults.md → "Feedback: shared
   explanation"); slide 33's two feedback panels (ids 26 and 28) are
   byte-identical below their titles, which is that default made
   explicit by the producer.

   STRUCTURE. The source panel holds 14 paragraphs, of which p2, p5 and
   p10 are genuinely EMPTY and there are ZERO <a:br> elements on the
   slide — so the authored structure is: title / [gap] / 2 lines /
   [gap] / 4 lines / [gap] / 4 lines. p2's gap is the title-to-body
   separation the popup's own header/body structure already provides;
   p5 and p10 are reproduced as real blank lines by the `null` entries
   below. Every other line break visible inside a sentence in the PPT
   is that textbox's own automatic wrapping and is NOT imitated with
   markup — the panel is widened instead
   (_question-template-defaults.md → "Automatic wrapping is not an
   authored instruction"; the same correction this part's S3 received).

   EXPRESSION ISOLATION + DISPLAY ORDER. Each complete expression sits
   in its own `<bdi dir="ltr" class="viq-expr">` — an LTR isolate so the
   surrounding RTL paragraph cannot reorder its tokens, plus nowrap so
   no digit, variable, operator or colon can be stranded on its own
   line. Because each expression is isolated, the order written inside
   the isolate IS the order displayed: nothing here is inferred from
   XML run order, DOM order or the Hebrew paragraph's direction.

   The required rendered order, confirmed by the producer against the
   PowerPoint (2026-09-09):

     5&#8729;3 = 15      (stored "5∙3 = 15")
     2&#8729;3 = 6       (stored "2∙3 = 6")
     2x + 5x = 21   (stored as displayed)
     7x = 21        (stored as displayed)
     x = 3          (stored "3 = x")

   ⚠ CORRECTED 2026-09-09. An earlier revision of this screen displayed
   the two arithmetic lines reversed ("15 = 3∙5" / "6 = 3∙2"). That came
   from over-generalising slide 32's lesson: I had derived those two
   orders from a BROWSER bidi simulation of the stored runs rather than
   from the PowerPoint's own output, and treated the simulation as
   evidence of what the deck renders. It is not — the producer's
   screenshots of the rendered deck show "5∙3 = 15" and "2∙3 = 6", i.e.
   the stored order. The general rule from slide 32 still stands (never
   infer displayed order from stored order alone), but the only
   admissible evidence for "what the deck displays" is the rendered
   deck, not a re-simulation of its bytes. The two lines below now carry
   the confirmed order.

   The multiplication sign is the deck's own U+2219 BULLET OPERATOR
   (written as `&#8729;`), with the deck's own spacing — no spaces
   around the operator, one on each side of the "=".

   ⚠ ONE MALFORMED SOURCE RUN (genuine, unlike the flag withdrawn on
   S3). The parenthetical in p6 is stored as the run pair
   "…גורם ההרחבה (" + " :(x" — i.e. TWO U+0028 LEFT PARENTHESIS
   characters, no closing parenthesis, and the colon between them. It
   is not a bidi illusion: rendered and measured, the deck itself
   displays a broken "…ההרחבה ) :)x" rather than the intended "(x):".
   This is the familiar PowerPoint RTL bracket/colon typing artifact.
   The unambiguous intent — "the expansion factor (x):", introducing
   the three equation lines that follow — is reproduced here as "(x):";
   written that way it renders correctly (measured), and the isolate
   additionally keeps it unbreakable. Producer note only; no other
   content on the slide is affected. */
const S4VIQ_EXPLAIN = [
  'نرمز إلى المسافة في المنطقة المستوية بـ <bdi dir="ltr" class="viq-expr">5x</bdi>',
  'وإلى المسافة في الصعود بـ <bdi dir="ltr" class="viq-expr">2x</bdi>.',
  null,
  'طول المسار هو 21 كم. نعوّض في المعادلة ونجد عامل التوسيع <bdi dir="ltr" class="viq-expr">(x)</bdi>:',
  '<bdi dir="ltr" class="viq-expr">2x + 5x = 21</bdi>',
  '<bdi dir="ltr" class="viq-expr">7x = 21</bdi>',
  '<bdi dir="ltr" class="viq-expr">x = 3</bdi>',
  null,
  'نحسب المسافة في المنطقة المستوية: <bdi dir="ltr" class="viq-expr">5&#8729;3 = 15</bdi>',
  'والمسافة في الصعود: <bdi dir="ltr" class="viq-expr">2&#8729;3 = 6</bdi>',
  'ونحصل على أن المسافة في المنطقة المستوية هي 15 كم،',
  'والمسافة في الصعود هي 6 كم‎.'
];
/* Titles are the slide's own authored ones (id=26 p1 / id=28 p1). The
   non-final retry state has no authored content on this slide, so it
   uses the canonical fallback title with no invented pedagogical body
   — same treatment as S1/S2/S3 in this part. */
const S4VIQ_POPUP = {
  retry:   { bg: '#ffdbdc', title: 'الإجابة غير صحيحة.', body: ['هل نحاول مرة أخرى؟'] },
  correct: { bg: '#edf8ed', title: 'ممتاز‎!', body: S4VIQ_EXPLAIN },
  wrong2: {
    bg: '#ffdbdc', title: 'هذا خطأ، لكن رائع أنكم حاولتم مواجهة التحدي‎!', body: S4VIQ_EXPLAIN,
    action: { label: 'الانتقال إلى إجابتي', view: 'answer' },
    answerView: {
      title: 'الإجابة غير صحيحة.',
      body:  ['يمكن مشاهدة الحل الصحيح مرة أخرى‎.'],
      action: { label: 'الإجابة الصحيحة', view: 'solution' }
    }
  }
};

function s4viqInput(k) { return document.getElementById('s4viq-input-' + k); }
function s4viqField(k) { return document.getElementById('s4viq-field-' + k); }
function s4viqIco(k)   { return document.getElementById('s4viq-ico-' + k); }
function s4viqNormalize(raw) { return String(raw).trim().replace(/,/g, ''); }
/* Retry gate compares the FULL joined signature across every field
   (ValueInputQuestion.md → "Retry gate comparison shape"): changing
   one field while the other stays identical must still re-enable
   צדקתי?, and a comma-only edit must not count as a change — hence the
   normalized values. */
function s4viqSignature() {
  return S4VIQ_KEYS.map(k => s4viqNormalize(s4viqInput(k).value)).join('|');
}
function s4viqAllFilled() {
  return S4VIQ_KEYS.every(k => s4viqNormalize(s4viqInput(k).value) !== '');
}
/* Per-field truth — the input to Case A marking. */
function s4viqFieldStates(vals) {
  const o = {};
  S4VIQ_KEYS.forEach(k => o[k] = Number(s4viqNormalize(vals[k])) === S4VIQ.correct[k]);
  return o;
}
function s4viqValues() {
  const o = {};
  S4VIQ_KEYS.forEach(k => o[k] = s4viqInput(k).value);
  return o;
}
function s4viqIsCorrect(vals) {
  return S4VIQ_KEYS.every(k => Number(s4viqNormalize(vals[k])) === S4VIQ.correct[k]);
}

function s4viqOnInput() {
  if (s4viqDone) return;
  S4VIQ_KEYS.forEach(k => {
    const el = s4viqInput(k);
    const cleaned = el.value.replace(/[^0-9,.-]/g, '');
    if (cleaned !== el.value) el.value = cleaned;
  });
  if (s4viqAttempts > 0) {
    document.getElementById('s4viq-popup').classList.add('hidden');
    s4viqMark(null);
  }
  s4viqUpdateBar();
}

function s4viqUpdateBar() {
  const chk = document.getElementById('s4viq-check');
  if (!chk || s4viqDone) return;
  chk.disabled = !s4viqAllFilled() || s4viqSignature() === s4viqLastWrongSig;
}

/* Case A marking. `states` is a per-key map of booleans; passing null
   clears every field. Each field gets its OWN border colour and its own
   16px bare verdict icon (720-design-guidelines →
   components-icons-states.md → Type 1), reusing the same two glyphs
   S1/S3 already use. */
function s4viqMark(states) {
  S4VIQ_KEYS.forEach(k => {
    const f = s4viqField(k), ico = s4viqIco(k);
    f.classList.remove('viq-correct', 'viq-wrong');
    if (states) f.classList.add(states[k] ? 'viq-correct' : 'viq-wrong');
    if (ico) ico.innerHTML = !states ? '' : (states[k] ? S1VIQ_ICON_CORRECT : S1VIQ_ICON_WRONG);
  });
}

function s4viqCheck() {
  if (s4viqDone) { advanceScreen(); return; }
  if (!s4viqAllFilled()) return;

  const vals = s4viqValues();
  s4viqAttempts++;
  /* xAPI (2026-11-11): the layer writes Kata's form from the metadata — performance, 'במישור[.]15[,]בעלייה[.]6' (proposed to Kata, which has it as `other` with no pattern) */
  reportQ(['002', 'q1'], s4viqIsCorrect(vals), s4viqIsCorrect(vals) || s4viqAttempts >= S4VIQ.maxAttempts,
    { steps: { 'في المنطقة المستوية': vals.plain, 'في الصعود': vals.uphill } });

  if (s4viqIsCorrect(vals)) {
    s4viqMark(s4viqFieldStates(vals));     // all true — every field green
    s4viqResult = 'correct';
    s4viqFinish();
    s4viqShowPopup('correct');
    return;
  }

  if (s4viqAttempts < S4VIQ.maxAttempts) {
    s4viqMark(s4viqFieldStates(vals));     // a mixed board is the point of Case A
    s4viqLastWrongSig = s4viqSignature();
    document.getElementById('s4viq-check').disabled = true;
    S4VIQ_POPUP.retry.title = Object.values(s4viqFieldStates(vals)).some(Boolean) ? RETRY_TITLE_PARTIAL : RETRY_TITLE;
    s4viqShowPopup('retry');
    return;
  }

  s4viqAnswer = vals;
  s4viqResult = 'wrong2';
  s4viqView   = 'solution';
  s4viqFinish();
  s4viqApplyView();
}

/* View My Answer Toggle — OBJECT level, never per field
   (ValueInputQuestion.md → Variant C → "View My Answer Toggle"): one
   button writes the whole correct map or the whole snapshotted answer
   map across every field in a single pass. The solution view is
   all-green (it is the correct answer being displayed); the answer view
   restores the learner's own values with their own per-field verdicts. */
function s4viqApplyView() {
  if (s4viqAnswer === null) return;
  const showAnswer = s4viqView === 'answer';
  S4VIQ_KEYS.forEach(k => {
    s4viqInput(k).value = showAnswer ? s4viqAnswer[k] : String(S4VIQ.correct[k]);
  });
  s4viqMark(showAnswer ? s4viqFieldStates(s4viqAnswer) : s4viqFieldStates(S4VIQ.correct));
  s4viqShowPopup('wrong2', s4viqView);
}
// Pure view switch — never routes through s4viqCheck(), never touches attempts/scoring.
function s4viqToggleView() {
  if (s4viqAnswer === null) return;
  s4viqView = (s4viqView === 'solution') ? 'answer' : 'solution';
  s4viqApplyView();
}

function s4viqFinish() {
  s4viqDone = true;
  S4VIQ_KEYS.forEach(k => { s4viqInput(k).disabled = true; });
  const chk = document.getElementById('s4viq-check');
  if (chk) { setNavLabel(chk, 'متابعة'); chk.disabled = false; }
  recordSection(4, s4viqResult);   /* Question 2, section א */
}

function s4viqResetInitial() {
  s4viqAttempts     = 0;
  s4viqDone         = false;
  s4viqResult       = null;
  s4viqAnswer       = null;
  s4viqView         = 'solution';
  s4viqLastWrongSig = null;
  S4VIQ_KEYS.forEach(k => {
    const el = s4viqInput(k);
    el.value = '';
    el.disabled = false;
  });
  s4viqMark(null);
  document.getElementById('s4viq-popup').classList.add('hidden');
  document.getElementById('s4viq-hint-overlay').classList.add('hidden');
  const chk = document.getElementById('s4viq-check');
  if (chk) { setNavLabel(chk, 'هل إجابتي صحيحة؟'); chk.disabled = true; }
  sectionResult[4] = null;
  /* Hint visible from the start — slide 33's hint trigger (id=2,
     y=649) sits in its own on-canvas base layer, same as S1/S2/S3. */
}

function s4viqShowPopup(type, view) {
  const popup = document.getElementById('s4viq-popup');
  if (!popup) return;
  const base = S4VIQ_POPUP[type];
  const cfg  = (view === 'answer' && base.answerView)
             ? Object.assign({}, base, base.answerView)
             : base;
  popup.style.background = cfg.bg;
  if (popup.classList.contains('hidden') || type === 'wrong2') resetPopupPosition(popup);
  const titleEl = document.getElementById('s4viq-popup-title');
  titleEl.innerHTML = (type === 'wrong2' && view !== 'answer')
    ? cfg.title + '<br>الإجابة الصحيحة تَظهَر الآن.'
    : cfg.title;
  /* A `null` body entry is one of the source's authored EMPTY
     paragraphs (slide 33 p5 / p10) — rendered as a real blank line
     rather than as an empty <p>, which would collapse to zero height. */
  document.getElementById('s4viq-popup-body').innerHTML =
    (cfg.body || []).map(p => p === null
      ? '<div class="popup-para-gap" aria-hidden="true"></div>'
      : '<p>' + p + '</p>').join('');
  renderPopupAction('s4viq-popup', cfg.action, s4viqToggleView);
  popup.classList.remove('hidden');
  clampPopup(popup);
}

function s4viqOpenHint() {
  xapiRequestedHint('002', 'q1');
  document.getElementById('s4viq-hint-overlay').classList.remove('hidden');
}
function s4viqCloseHint() {
  document.getElementById('s4viq-hint-overlay').classList.add('hidden');
}
function s4viqCloseHintOnBackdrop(e) {
  if (e.target === e.currentTarget) s4viqCloseHint();
}

/* ═══════════════════════════════════════════════════════════
   GLOBAL COMPONENT — Image zoom (clone-based)
   720-templates → _global-components.md → "Image zoom". One shared
   modal per part; first needed by S5, which presents a content image
   (_question-template-defaults.md → "Image zoom" — required, not
   optional, for any question template with a primary learning image).

   Open/close is ONE delegated document handler: any [data-zoom-src]
   opens, any [data-zoom-close="true"] closes (both the backdrop and the
   close button carry it), Escape closes. No inline onclick anywhere.

   SIZING. The stage is the canonical Media viewport. It FILLS the space
   available inside the surface, preserving the media's own aspect ratio
   and capped by the 1200×640 envelope — never a fixed square, and never
   distorted.

   ⚠ SCREEN-SPECIFIC PRODUCER EXCEPTION (screens 4 and 5 of this part
   only, approved 2026-09-09). This is NOT a general reinterpretation of
   the shared zoom sizing guidance and must not be carried elsewhere on
   its strength. The slight pixelation at ~2× is an accepted tradeoff
   here in exchange for a meaningful enlargement; do not restore the
   intrinsic size and do not seek replacement assets.

   An earlier revision clamped the
   scale at 1 so a source smaller than the viewport rendered at its own
   pixel size — which made zoom a no-op for both of this part's images,
   since the deck places each on the slide at roughly its native size.
   That came from reading 720-design-guidelines → "Image Zoom Surface" →
   Sizing ("Do not enlarge a relatively small source merely to fill the
   envelope") as a hard clamp, and from misreading the producer's own
   note — "Figma shows the source image at its available size inside the
   larger modal composition" means the size AVAILABLE TO IT in the modal,
   not its intrinsic size. The producer confirmed the zoom must visibly
   enlarge. The clamp is therefore removed: the media scales up to the
   available viewport, and the guideline sentence is read as "don't hunt
   for a higher-resolution asset just to fill the envelope", not as a
   prohibition on scaling the raster. This also matches the approved
   reference implementation (methodica-math-percent-02-07).

   Resulting geometry: the available viewport is 1152×552 (1200−2×24 by
   640−64−24), so trail-21.png (429×286) renders 828×552 and
   trail-42.png (392×261) renders 829×552 — both height-bound, as the
   visual contract says is normal for common ratios, and both a genuine
   ~2× enlargement over the on-slide size. */
/* The trigger that opened the modal, so focus can go back to it on
   close (keyboard users must not be dropped at the top of the page). */
let zoomReturnFocus = null;

function openImageZoom(btn) {
  /* Both of this part's image wrappers feed the one shared modal —
     screen 4's `.viq-img-frame` and screen 5's `.scq-img-frame`. Keyed on
     the wrapper CLASSES, never on a screen id, because the clone lands
     in the shared stage outside its own screen. */
  const wrapper = btn.closest('.scq-img-frame, .viq-img-frame');
  const stage   = document.getElementById('img-zoom-stage');
  const modal   = document.getElementById('img-zoom-modal');
  const panel   = modal && modal.querySelector('.img-zoom-panel');
  if (!wrapper || !stage || !modal || !panel) return;

  const srcImg = wrapper.querySelector('img');
  /* Clone the WRAPPER, not the bare file, so any overlay living inside
     it comes along by construction; drop the trigger from the clone. */
  const clone = wrapper.cloneNode(true);
  const cloneBtn = clone.querySelector('.img-zoom-btn');
  if (cloneBtn) cloneBtn.remove();
  stage.innerHTML = '';
  stage.appendChild(clone);

  modal.classList.remove('hidden');
  modal.setAttribute('aria-hidden', 'false');

  if (srcImg && srcImg.naturalWidth && srcImg.naturalHeight) {
    const cs   = getComputedStyle(panel);
    const boxW = 1200 - parseFloat(cs.paddingLeft) - parseFloat(cs.paddingRight);
    const boxH = 640  - parseFloat(cs.paddingTop)  - parseFloat(cs.paddingBottom);
    /* Fit the media to the available viewport, preserving its ratio —
       whichever axis binds first. NOT clamped at 1: a source smaller
       than the viewport is scaled UP so the zoom actually enlarges (see
       the SIZING note above). */
    const scale = Math.min(boxW / srcImg.naturalWidth, boxH / srcImg.naturalHeight);
    stage.style.width  = Math.round(srcImg.naturalWidth  * scale) + 'px';
    stage.style.height = Math.round(srcImg.naturalHeight * scale) + 'px';
  }
  zoomReturnFocus = btn;
  const closeBtn = document.getElementById('img-zoom-close');
  if (closeBtn) closeBtn.focus();
}
function closeImageZoom() {
  const modal = document.getElementById('img-zoom-modal');
  const stage = document.getElementById('img-zoom-stage');
  if (modal) { modal.classList.add('hidden'); modal.setAttribute('aria-hidden', 'true'); }
  if (stage) { stage.innerHTML = ''; stage.style.width = ''; stage.style.height = ''; }
  /* Focus restoration — back to the trigger that opened it, whichever
     screen that was. Guarded so closing with nothing open is a no-op. */
  if (zoomReturnFocus && document.contains(zoomReturnFocus)) zoomReturnFocus.focus();
  zoomReturnFocus = null;
}
document.addEventListener('click', e => {
  const zoomBtn = e.target.closest && e.target.closest('[data-zoom-src]');
  if (zoomBtn) { openImageZoom(zoomBtn); return; }
  if (e.target.closest && e.target.closest('[data-zoom-close="true"]')) closeImageZoom();
});
document.addEventListener('keydown', e => {
  if (e.key !== 'Escape') return;
  const modal = document.getElementById('img-zoom-modal');
  if (modal && !modal.classList.contains('hidden')) closeImageZoom();
});
/* Called by goTo(): an enlarged image must never stay over the next screen. */
function closeImageZoomOnNavigate() {
  const modal = document.getElementById('img-zoom-modal');
  if (modal && !modal.classList.contains('hidden')) closeImageZoom();
}

/* ═══════════════════════════════════════════════════════════
   S5 — SingleChoiceQuestionImage (slide 34)
   Q2 of 2 — closes this part's SECOND question series. Same template as
   this part's approved S2 (slide 31), whose JS shape is reused rather
   than reinvented; only the config, content, the image column and the
   image-zoom component differ. See the source-evidence block above this
   screen in index.html for the full ID inventory, the render
   reconciliation and the answer-key evidence.

   CORRECT ANSWER = o4 (the LAST option in authored reading order).
   notesSlide31 states no answer, so three independent signals settle it:
     (1) the producer authored o4's pill AND radio in green 00B050 with
         the radio filled, against neutral cyan 00B0F0 with unfilled
         radios on o1/o2/o3 — slide 31's own convention for marking the
         correct option, and visible in the rendered slide;
     (2) the feedback's closing sentence is o4's text verbatim;
     (3) the feedback's algebra: 2x + 5x = 42 → 7x = 42 → x = 6 against
         the previous route's x = 3, i.e. the factor doubled.
   OPTION ORDER is by geometry (group tops 316.8 / 397.5 / 477.5 /
   557.6), NOT by XML order — the groups appear in the file as
   11, 27, 32, 37, and group 32 is authored last visually. The four gaps
   are uniform (80.7 / 80.0 / 80.1), which is the positive check that no
   fifth capsule is missing.

   No View My Answer Toggle — SingleChoiceQuestion shows the chosen wrong
   option and the correct one side by side, which is the applicability
   rule's explicit exclusion (_global-components.md → "View My Answer
   Toggle" → Applicability). Completion label `هل نتابع؟` — this
   template's own canonical fallback, not ValueInputQuestion's `متابعة`. */
const S5SC_EXPLAIN = [
  /* Slide 34's two panels hold 12 paragraphs each, identical below their
     titles, with p2/p4/p8/p11 genuinely EMPTY and zero <a:br> on the
     slide. p2's gap is the title-to-body separation the popup structure
     already provides; p4, p8 and p11 are reproduced as real blank lines
     by the `null` entries below. Ratio display order verified on the
     PowerPoint render (stored "2 : 5" → displays "5 : 2"); every
     expression is an LTR isolate with nowrap so it can neither reorder
     nor be split mid-expression. */
  'النسبة بين المسافة في المنطقة المستوية والمسافة في الصعود تبقى <bdi dir="ltr" class="viq-expr">5 : 2</bdi>، لكن طول المسار تضاعف (42 كم)، ولذلك‎:',
  null,
  '<bdi dir="ltr" class="viq-expr">2x + 5x = 42</bdi>',
  '<bdi dir="ltr" class="viq-expr">7x = 42</bdi>',
  '<bdi dir="ltr" class="viq-expr">x = 6</bdi>',
  null,
  'في مسار الـ 21 كم وجدنا‎:',
  '<bdi dir="ltr" class="viq-expr">x = 3</bdi>',
  null,
  'أي إنّ <bdi dir="ltr" class="viq-expr">x</bdi> كبرَ بمرتين، لأن طول المسار الإجمالي كبرَ بمرتين‎.'
];
const S5SC = { correctId: 'o4', maxAttempts: 2 };
/* Titles are the slide's own (id=20 p1 / id=22 p1). The non-final retry
   state has no authored content on this slide, so it uses the canonical
   fallback title with no invented pedagogical body — same treatment as
   S1-S4 in this part. */
const S5SC_POPUP = {
  retry:   { bg: '#ffdbdc', title: 'الإجابة غير صحيحة.', body: ['هل نحاول مرة أخرى؟'] },
  correct: { bg: '#edf8ed', title: 'صحيح‎!', body: S5SC_EXPLAIN },
  wrong2:  { bg: '#ffdbdc', title: 'هذا خطأ، لكن مجرد التعامل مع المسألة يطوّر التفكير‎!', body: S5SC_EXPLAIN }
};

let s5scSelected  = null;
let s5scAttempts  = 0;
let s5scDone      = false;
let s5scResult    = null;   // 'correct' | 'wrong2'
let s5scLastWrong = null;   // the option id just marked wrong — the retry gate compares against THIS

function s5scOpt(id) { return document.querySelector('#s5 .scq-opt[data-id="' + id + '"]'); }
function s5scOpts()  { return document.querySelectorAll('#s5 .scq-opt'); }

/* Single select; clears a prior wrong attempt's marks; re-enables check
   only if the newly selected id differs from the one just marked wrong
   (_question-template-defaults.md → "Retry gate" — comparison shape for
   this template: the selected option id). Re-selecting the same wrong
   option keeps צדקתי? disabled. */
function s5scSelect(id) {
  if (s5scDone) return;
  s5scOpts().forEach(o => {
    o.classList.remove('selected');
    o.setAttribute('aria-checked', 'false');
  });
  const opt = s5scOpt(id);
  opt.classList.add('selected');
  opt.setAttribute('aria-checked', 'true');
  s5scSelected = id;

  if (s5scAttempts > 0) {
    document.getElementById('s5sc-popup').classList.add('hidden');
    s5scOpts().forEach(o => o.classList.remove('correct', 'wrong'));
  }
  const chk = document.getElementById('s5sc-check');
  if (chk) chk.disabled = (id === s5scLastWrong);
}

function s5scCheck() {
  if (s5scDone) { advanceScreen(); return; }
  if (!s5scSelected) return;

  s5scAttempts++;
  reportQ(['002', 'q2'], s5scSelected === S5SC.correctId, s5scSelected === S5SC.correctId || s5scAttempts >= S5SC.maxAttempts, xapiChoiceAnswer('002', 'q2', '#s5 .scq-opt', s5scSelected));

  if (s5scSelected === S5SC.correctId) {
    s5scOpt(s5scSelected).classList.add('correct');
    s5scResult = 'correct';
    s5scFinish();
    s5scShowPopup('correct');
    return;
  }

  if (s5scAttempts < S5SC.maxAttempts) {
    /* Non-final wrong: mark the chosen option only — the correct one is
       NOT revealed yet (SingleChoiceQuestion.md → state 4). */
    s5scOpt(s5scSelected).classList.add('wrong');
    s5scLastWrong = s5scSelected;
    document.getElementById('s5sc-check').disabled = true;
    s5scShowPopup('retry');
    return;
  }

  /* Final wrong: chosen option red+✗ AND the correct one green+✓, both
     at once — the reason this template is excluded from the toggle. */
  s5scOpt(s5scSelected).classList.add('wrong');
  s5scOpt(S5SC.correctId).classList.add('correct');
  s5scResult = 'wrong2';
  s5scFinish();
  s5scShowPopup('wrong2');
}

function s5scFinish() {
  s5scDone = true;
  s5scOpts().forEach(o => { o.disabled = true; });
  const chk = document.getElementById('s5sc-check');
  if (chk) { setNavLabel(chk, 'هل نتابع؟'); chk.disabled = false; }
  hideHintButton('s5sc-hint');
  recordSection(5, s5scResult);    /* Question 2, section ב */
}

function s5scResetInitial() {
  s5scSelected  = null;
  s5scAttempts  = 0;
  s5scDone      = false;
  s5scResult    = null;
  s5scLastWrong = null;
  s5scOpts().forEach(o => {
    o.classList.remove('selected', 'correct', 'wrong');
    o.setAttribute('aria-checked', 'false');
    o.disabled = false;
  });
  document.getElementById('s5sc-popup').classList.add('hidden');
  document.getElementById('s5sc-hint-overlay').classList.add('hidden');
  const chk = document.getElementById('s5sc-check');
  if (chk) { setNavLabel(chk, 'هل إجابتي صحيحة؟'); chk.disabled = true; }
  /* Hint visible from the start — slide 34's hint trigger (id=2, y=649)
     sits in its own on-canvas base layer, exactly as on S1-S4, so this
     part's established policy applies. The hint also stays re-openable
     after use, matching the approved S2 (same template) rather than
     SingleChoiceQuestion.md's generic "disables after use" line; that
     deviation is this component's own approved precedent and is flagged
     in the build report. */
  showHintButton('s5sc-hint');
  sectionResult[5] = null;
}

function s5scShowPopup(type) {
  const popup = document.getElementById('s5sc-popup');
  if (!popup) return;
  const cfg = S5SC_POPUP[type];
  popup.style.background = cfg.bg;
  if (popup.classList.contains('hidden')) resetPopupPosition(popup);
  const titleEl = document.getElementById('s5sc-popup-title');
  /* Terminal-incorrect title carries the fixed second line
     (_question-template-defaults.md). */
  titleEl.innerHTML = (type === 'wrong2')
    ? cfg.title + '<br>الإجابة الصحيحة تَظهَر الآن.'
    : cfg.title;
  /* A `null` body entry is one of the source's authored EMPTY paragraphs
     (slide 34 p4 / p8 / p11) — a real blank line, not an empty <p>,
     which would collapse to zero height. */
  document.getElementById('s5sc-popup-body').innerHTML =
    cfg.body.map(p => p === null
      ? '<div class="popup-para-gap" aria-hidden="true"></div>'
      : '<p>' + p + '</p>').join('');
  popup.classList.remove('hidden');
  clampPopup(popup);
}

function s5scOpenHint() {
  xapiRequestedHint('002', 'q2');
  document.getElementById('s5sc-hint-overlay').classList.remove('hidden');
}
function s5scCloseHint() {
  document.getElementById('s5sc-hint-overlay').classList.add('hidden');
}
function s5scCloseHintOnBackdrop(e) {
  if (e.target === e.currentTarget) s5scCloseHint();
}

/* Enter/Space on a focused option selects it — radiogroup keyboard a11y,
   same delegated pattern as S2's. */
document.addEventListener('keydown', e => {
  const isEnter = e.key === 'Enter' || e.code === 'Enter' || e.code === 'NumpadEnter';
  const isSpace = e.key === ' ' || e.key === 'Spacebar' || e.code === 'Space';
  if (!isEnter && !isSpace) return;
  const opt = e.target.closest && e.target.closest('#s5 .scq-opt');
  if (!opt || opt.disabled) return;
  e.preventDefault();
  s5scSelect(opt.dataset.id);
});

/* ═══════════════════════════════════════════════════════════
   APPLET / IFRAME ZOOM — figma-lomda-builder → reference/patterns.md →
   "Zoom modal". A component DISTINCT from the global image zoom above:
   that one clones an image wrapper onto a shared stage; this one
   enlarges a live embedded interactive. They share only the canonical
   backdrop. Required by slide 35's presenter note ("בלחיצה על זכוכית
   המגדלת היישומון גדל").

   NEITHER move NOR clone the iframe. The applet is stateful — the
   learner can click and drag points and mark a range on the graph — and
   an iframe loses its document the moment it is re-parented. Measured,
   not assumed: planting a probe on the iframe's contentWindow and then
   appending the element to another parent returned the probe gone and a
   fresh readyState. cloneNode and a second same-src iframe fail for the
   same or worse reasons.
   So this zoom PROMOTES THE CARD IN PLACE — a class changes the card's
   geometry to the modal surface's media viewport while the card keeps
   its DOM position and its iframe child untouched. The applet is never
   reloaded; it reflows into the larger box, which is the right result
   for a live layout (a raster zoom would only magnify pixels).
   This is the structural difference from the image zoom above: that one
   clones a wrapper onto a shared stage because a static image can be
   copied freely; this one must not copy or move anything. */

/* The applet's fixed logical viewport — the smallest box in which the
   widget composes with zero overflow on either axis (measured across
   eight sizes; 630x374 needs 439px of height, 800x475 still overflows
   by 18px, 900x535 is exact). Its aspect matches the authored card's to
   within 0.08%, so scaling it to the card is distortion-free. See the
   `.viq-applet` rule in css/style.css. */
const APPLET_VIEWPORT = { w: 900, h: 535 };

/* Which screen's card is currently promoted. Slides 35, 36 and 37 all
   embed the same widget, so the zoom is shared and has to know which
   card it is holding — close, Escape, the backdrop and the iframe's own
   Escape bridge all resolve through this. */
let appletZoomPrefix = null;

function openAppletZoom(prefix) {
  prefix = prefix || appletZoomPrefix || 's6';
  const modal = document.getElementById('applet-zoom-modal');
  const card  = document.getElementById(prefix + '-applet-frame');
  const panel = modal && modal.querySelector('.applet-zoom-panel');
  if (!modal || !card || !panel) return;
  appletZoomPrefix = prefix;

  modal.classList.remove('hidden');         // measure it laid out, not hidden
  modal.setAttribute('aria-hidden', 'false');
  card.classList.add('is-zoomed');          // visual overrides only

  /* The card IS the dialog while enlarged — it holds the applet and the
     close control. The role is applied here rather than authored on the
     card, because in the inline state the card is ordinary screen
     content and must not announce itself as a dialog. */
  card.setAttribute('role', 'dialog');
  card.setAttribute('aria-modal', 'true');
  card.setAttribute('aria-label', 'تكبير التطبيق');

  /* Place the card on the panel's own content box, measured rather than
     hardcoded. Both the modal and the card are absolutely positioned
     against boxes that are inset:0 of #app, so their offset coordinates
     are already canvas coordinates and no scale math is needed — and
     because the panel is flex-centred, this follows it when scaleApp()
     grows the canvas past 710 on a surplus viewport (at 1280x900 the
     panel sits at top 130, not 40). */
  const cs = getComputedStyle(panel);
  const padL = parseFloat(cs.paddingLeft),  padR = parseFloat(cs.paddingRight);
  const padT = parseFloat(cs.paddingTop),   padB = parseFloat(cs.paddingBottom);

  /* Fit the applet's fixed logical viewport into the envelope's media
     area, then shrink the surface to what the applet actually occupies —
     the same media-responsive rule the image-zoom surface follows, so no
     white slack is left around the enlarged applet. */
  panel.style.width = ''; panel.style.height = '';       // measure the envelope
  const availW = panel.clientWidth  - padL - padR;
  const availH = panel.clientHeight - padT - padB;
  const fit = Math.min(availW / APPLET_VIEWPORT.w, availH / APPLET_VIEWPORT.h);
  const w = Math.round(APPLET_VIEWPORT.w * fit);
  const h = Math.round(APPLET_VIEWPORT.h * fit);
  panel.style.width  = (w + padL + padR) + 'px';
  panel.style.height = (h + padT + padB) + 'px';

  /* Place the card on the panel's own content box, measured rather than
     hardcoded. Both the modal and the card are absolutely positioned
     against boxes that are inset:0 of #app, so their offset coordinates
     are already canvas coordinates and no scale math is needed — and
     because the panel is flex-centred, this follows it when scaleApp()
     grows the canvas past 710 on a surplus viewport (at 1280x900 the
     panel sits lower than at 1280x720). */
  card.style.left   = (panel.offsetLeft + padL) + 'px';
  card.style.top    = (panel.offsetTop  + padT) + 'px';
  card.style.width  = w + 'px';
  card.style.height = h + 'px';
  card.style.setProperty('--applet-scale', fit);

  /* Each card carries its OWN close control (it lives inside the card,
     at the enlarged frame's physical top-left), so resolve it within the
     promoted card rather than by a single global id. */
  const closeBtn = card.querySelector('.viq-applet-close');
  if (closeBtn) closeBtn.focus();
}
function closeAppletZoom() {
  const modal = document.getElementById('applet-zoom-modal');
  const prefix = appletZoomPrefix;
  const card  = prefix && document.getElementById(prefix + '-applet-frame');
  if (!modal) return;
  if (card) {
    card.classList.remove('is-zoomed');
    card.style.left = card.style.top = card.style.width = card.style.height = '';
    card.style.removeProperty('--applet-scale');   // back to the CSS default
    card.removeAttribute('role');                  // ordinary content again
    card.removeAttribute('aria-modal');
    card.removeAttribute('aria-label');
  }
  const panel = modal.querySelector('.applet-zoom-panel');
  if (panel) { panel.style.width = ''; panel.style.height = ''; }
  modal.classList.add('hidden');
  modal.setAttribute('aria-hidden', 'true');
  /* Focus restoration — back to the trigger, same contract as the image
     zoom, so a keyboard learner is never dropped at the top of the page. */
  const trigger = prefix && document.getElementById(prefix + '-applet-zoom');
  if (trigger && document.contains(trigger)) trigger.focus();
  appletZoomPrefix = null;
}
document.addEventListener('click', e => {
  if (e.target.closest && e.target.closest('[data-applet-zoom-close="true"]')) closeAppletZoom();
});
document.addEventListener('keydown', e => {
  if (e.key !== 'Escape') return;
  const modal = document.getElementById('applet-zoom-modal');
  if (modal && !modal.classList.contains('hidden')) closeAppletZoom();
});
/* Escape pressed INSIDE the applet. A keydown in the iframe's document
   never propagates to this one, so once the learner clicks into the
   enlarged applet the listener above stops seeing Escape. The widget
   forwards it by postMessage instead.
   Authenticated by source, not by origin: a file:// iframe reports the
   opaque origin "null", so an origin comparison is worthless here —
   matching event.source against the iframe's own contentWindow is the
   check that actually means something. */
window.addEventListener('message', e => {
  if (!e.data || e.data.type !== 'DCX_APPLET_ESC') return;
  const frame = appletZoomPrefix && document.getElementById(appletZoomPrefix + '-applet');
  if (!frame || e.source !== frame.contentWindow) return;
  const modal = document.getElementById('applet-zoom-modal');
  if (modal && !modal.classList.contains('hidden')) closeAppletZoom();
});
/* Called by goTo(): an enlarged applet must never stay promoted over the
   next screen — a card left promoted on a hidden screen shows as a blank
   white panel over the destination. */
function closeAppletZoomOnNavigate() {
  const modal = document.getElementById('applet-zoom-modal');
  if (modal && !modal.classList.contains('hidden')) closeAppletZoom();
}

/* ═══════════════════════════════════════════════════════════
   S6 — ValueInputQuestion, VARIANT A (one value, inline in prose)
   slide 35 · Q1 of 3, opening this part's THIRD question series.

   Spec: 720-templates → ValueInputQuestion.md (Variant A) +
   _question-template-defaults.md + _global-components.md → "View My
   Answer Toggle". Mechanism reused from this part's approved S4 (same
   template, same inline-in-prose shape) — only the field count, the
   config and the content differ. The View My Answer Toggle DOES apply
   (ValueInputQuestion with a destructive final-wrong reveal).

   CORRECT VALUE: 6 — on three independent pieces of evidence:
   notesSlide32 states it outright ("الإجابة الصحيحة היא 6"); the source's
   own answer box (id=7) carries the placeholder digit "6"; and the
   feedback derives it (a square's perimeter is 4 equal sides, so
   4 × 1.5 = 6). Consistent with the applet's own graph, where the
   perimeter/side relation is P = 4s.

   The learner reads 1.5 off the graph by interpolating between the
   plotted integer points, which is what the hint points at. Comma and
   decimal input are both normalized, so "6", "6.0" and " 6 " all pass. */
const S6VIQ = { correct: 6, maxAttempts: 2 };

let s6viqAttempts     = 0;
let s6viqDone         = false;
let s6viqResult       = null;   // 'correct' | 'wrong2'
let s6viqAnswer       = null;   // raw typed string, snapshotted before the reveal
let s6viqView         = 'solution';
let s6viqLastWrongSig = null;

/* Shared explanation — one body for correct and wrong2, differing only
   in title; slide 35's two panels (ids 26 and 28) are identical below
   their titles, which is that default made explicit.

   STRUCTURE. Eight paragraphs each, with p2, p4 and p7 genuinely EMPTY
   and zero <a:br> on the slide. p2's gap is the title-to-body separation
   the popup structure already provides; p4 and p7 are reproduced as real
   blank lines by the `null` entries below.

   EXPRESSION ORDER. "1.5 ∙ 4 = 6" is reproduced in the deck's stored
   order, per the order the producer confirmed for this part's arithmetic
   lines on screens 4 and 5. The multiplication sign is the deck's own
   U+2219 BULLET OPERATOR with the deck's own spacing. Each expression
   sits in its own LTR isolate with nowrap, so it can neither reorder
   inside the RTL paragraph nor be split mid-expression. */
const S6VIQ_EXPLAIN = [
  'محيط المربع هو مجموع أطوال أضلاعه الـ 4 المتساوية‎.',
  null,
  'إذا كان طول كل ضلع <bdi dir="ltr" class="viq-expr">1.5</bdi> سم، نحصل على‎:',
  '<bdi dir="ltr" class="viq-expr">1.5 &#8729; 4 = 6</bdi>',
  null,
  'أي إن محيط المربع هو 6 سم‎.'
];
/* Titles are the slide's own (id=26 p1 / id=28 p1). The non-final retry
   state has no authored content on this slide, so it uses the canonical
   fallback title with no invented pedagogical body — same treatment as
   S1-S5 in this part. */
const S6VIQ_POPUP = {
  retry:   { bg: '#ffdbdc', title: 'الإجابة غير صحيحة.', body: ['هل نحاول مرة أخرى؟'] },
  correct: { bg: '#edf8ed', title: 'إجابة صحيحة‎!', body: S6VIQ_EXPLAIN },
  wrong2: {
    bg: '#ffdbdc', title: 'أخطأتم، لا بأس، فكل محاولة تقرّبكم من الفهم‎ أكثر.', body: S6VIQ_EXPLAIN,
    action: { label: 'الانتقال إلى إجابتي', view: 'answer' },
    answerView: {
      title: 'الإجابة غير صحيحة.',
      body:  ['يمكن مشاهدة الحل الصحيح مرة أخرى‎.'],
      action: { label: 'الإجابة الصحيحة', view: 'solution' }
    }
  }
};

function s6viqInput() { return document.getElementById('s6viq-input'); }
function s6viqNormalize(raw) { return String(raw).trim().replace(/,/g, ''); }
function s6viqSignature() { return s6viqNormalize(s6viqInput().value); }
function s6viqFilled() { return s6viqSignature() !== ''; }
function s6viqIsCorrect() { return Number(s6viqSignature()) === S6VIQ.correct; }

function s6viqOnInput() {
  if (s6viqDone) return;
  const el = s6viqInput();
  const cleaned = el.value.replace(/[^0-9,.-]/g, '');
  if (cleaned !== el.value) el.value = cleaned;
  if (s6viqAttempts > 0) {
    document.getElementById('s6viq-popup').classList.add('hidden');
    s6viqMark(null);
  }
  s6viqUpdateBar();
}

function s6viqUpdateBar() {
  const chk = document.getElementById('s6viq-check');
  if (!chk || s6viqDone) return;
  /* Retry gate (_question-template-defaults.md): comparison shape for a
     single-value screen is the normalized value itself, so a comma-only
     edit is not a change and keeps צדקתי? disabled. */
  chk.disabled = !s6viqFilled() || s6viqSignature() === s6viqLastWrongSig;
}

/* One field, one verdict — a 16px bare glyph inside the box, reusing the
   two icons S1/S3/S4 already use (720-design-guidelines →
   components-icons-states.md → Type 1). */
function s6viqMark(outcome) {
  const f = document.getElementById('s6viq-field');
  const ico = document.getElementById('s6viq-ico');
  f.classList.remove('viq-correct', 'viq-wrong');
  if (outcome) f.classList.add(outcome === 'correct' ? 'viq-correct' : 'viq-wrong');
  if (ico) ico.innerHTML = !outcome ? '' : (outcome === 'correct' ? S1VIQ_ICON_CORRECT : S1VIQ_ICON_WRONG);
}

function s6viqCheck() {
  if (s6viqDone) { advanceScreen(); return; }
  if (!s6viqFilled()) return;

  s6viqAttempts++;
  reportQ(['003', 'q1'], s6viqIsCorrect(), s6viqIsCorrect() || s6viqAttempts >= S6VIQ.maxAttempts, String((document.getElementById('s6viq-input') || {}).value).trim());

  if (s6viqIsCorrect()) {
    s6viqMark('correct');
    s6viqResult = 'correct';
    s6viqFinish();
    s6viqShowPopup('correct');
    return;
  }

  if (s6viqAttempts < S6VIQ.maxAttempts) {
    s6viqMark('wrong');
    s6viqLastWrongSig = s6viqSignature();
    document.getElementById('s6viq-check').disabled = true;
    s6viqShowPopup('retry');
    return;
  }

  s6viqAnswer = s6viqInput().value;
  s6viqResult = 'wrong2';
  s6viqView   = 'solution';
  s6viqFinish();
  s6viqApplyView();
}

/* View My Answer Toggle — the final-wrong reveal is destructive (the
   correct value overwrites the learner's own), so it is a toggle, never
   a one-way reveal. */
function s6viqApplyView() {
  if (s6viqAnswer === null) return;
  if (s6viqView === 'answer') {
    s6viqInput().value = s6viqAnswer;
    s6viqMark('wrong');
  } else {
    s6viqInput().value = String(S6VIQ.correct);
    s6viqMark('correct');
  }
  s6viqShowPopup('wrong2', s6viqView);
}
// Pure view switch — never routes through s6viqCheck(), never touches attempts.
function s6viqToggleView() {
  if (s6viqAnswer === null) return;
  s6viqView = (s6viqView === 'solution') ? 'answer' : 'solution';
  s6viqApplyView();
}

function s6viqFinish() {
  s6viqDone = true;
  s6viqInput().disabled = true;
  const chk = document.getElementById('s6viq-check');
  if (chk) { setNavLabel(chk, 'متابعة'); chk.disabled = false; }
  recordSection(6, s6viqResult);   /* Question 3, section א */
}

function s6viqResetInitial() {
  s6viqAttempts     = 0;
  s6viqDone         = false;
  s6viqResult       = null;
  s6viqAnswer       = null;
  s6viqView         = 'solution';
  s6viqLastWrongSig = null;
  const el = s6viqInput();
  el.value = '';
  el.disabled = false;
  s6viqMark(null);
  document.getElementById('s6viq-popup').classList.add('hidden');
  document.getElementById('s6viq-hint-overlay').classList.add('hidden');
  const chk = document.getElementById('s6viq-check');
  if (chk) { setNavLabel(chk, 'هل إجابتي صحيحة؟'); chk.disabled = true; }
  sectionResult[6] = null;
  /* Hint visible from the start — slide 35's hint trigger (id=2, y=649)
     sits in its own on-canvas base layer, as on every screen in this part.
     Deliberately NOT resetting the applet: its src is authored once and
     the iframe is never re-mounted, so the learner's marked range on the
     graph survives leaving and re-entering the screen. */
}

function s6viqShowPopup(type, view) {
  const popup = document.getElementById('s6viq-popup');
  if (!popup) return;
  const base = S6VIQ_POPUP[type];
  const cfg  = (view === 'answer' && base.answerView)
             ? Object.assign({}, base, base.answerView)
             : base;
  popup.style.background = cfg.bg;
  if (popup.classList.contains('hidden') || type === 'wrong2') resetPopupPosition(popup);
  const titleEl = document.getElementById('s6viq-popup-title');
  titleEl.innerHTML = (type === 'wrong2' && view !== 'answer')
    ? cfg.title + '<br>الإجابة الصحيحة تَظهَر الآن.'
    : cfg.title;
  /* A `null` body entry is one of the source's authored EMPTY paragraphs
     (slide 35 p4 / p7) — a real blank line, not an empty <p>. */
  document.getElementById('s6viq-popup-body').innerHTML =
    (cfg.body || []).map(p => p === null
      ? '<div class="popup-para-gap" aria-hidden="true"></div>'
      : '<p>' + p + '</p>').join('');
  renderPopupAction('s6viq-popup', cfg.action, s6viqToggleView);
  popup.classList.remove('hidden');
  clampPopup(popup);
}

function s6viqOpenHint() {
  xapiRequestedHint('003', 'q1');
  document.getElementById('s6viq-hint-overlay').classList.remove('hidden');
}
function s6viqCloseHint() {
  document.getElementById('s6viq-hint-overlay').classList.add('hidden');
}
function s6viqCloseHintOnBackdrop(e) {
  if (e.target === e.currentTarget) s6viqCloseHint();
}

/* Cross-part navigation. init() must run AFTER this part's own initial
   screen setup (resetScreenState(0) has already run via the initial
   DOM state below): if this page was entered via Back, it jumps
   straight to the screen the learner left from. */
PartNav.init({
  part: '02',
  getScreen: function () { return currentScreen; },
  goTo: goTo
});

/* Initial mount for Screen 0, since goTo() is only invoked on
   navigation — the first screen needs the same resetScreenState() pass
   the engine gives every later entry. */
resetScreenState(0);

/* ─── Dev mode: postMessage bridge ─────────────────────── */
window.addEventListener('message', (e) => {
  if (!e.data || e.data.type !== 'DEV_GOTO') return;
  const n = parseInt(e.data.screen, 10);
  if (!isNaN(n)) goTo(n);
});

// Notify parent of actual screen count (DOM-based — auto-updates when screens are added)
if (window.parent !== window) {
  const screenCount = document.querySelectorAll('.screen').length;
  window.parent.postMessage({ type: 'DEV_READY', total: screenCount }, '*');
}


/* ═══════════════════════════════════════════════════════════
   S7 — DropdownQuestion, MULTI-BLANK (two slots in one sentence)
   slide 36 · Q2 of 3, continuing the series slide 35 opened.

   Spec: 720-templates → DropdownQuestion.md ("Multi-blank (sentence
   fill-in)") + _question-template-defaults.md + _global-components.md →
   "View My Answer Toggle". Namespaced state object rather than flat
   variables, per that file's guidance for a multi-slot screen, and
   because per-slot correctness has to survive a mixed final attempt.
   ═══════════════════════════════════════════════════════════ */

/* Slot 1 is the FIRST (right-hand) blank in RTL — "בין ___" — and slot 2
   the second, "לבין ___". Values are stored as the authored fraction
   strings so the comparison is exact and never touches float rounding
   (3/4 and 7/4 are not on the widget's 0.5 grid; see the applet note in
   index.html). Source: the two closed capsules id=7 «3/4» / id=13 «7/4»,
   the producer's green highlight on exactly those two, and both feedback
   bodies. */
const S7DQ = {
  correct: { 1: '3/4', 2: '7/4' },
  maxAttempts: 2,
  slots: [1, 2]
};

window.lomdaState = window.lomdaState || {};
function s7dqState() {
  window.lomdaState.s7 = window.lomdaState.s7 || {
    attempts: 0,
    done: false,
    selections: { 1: null, 2: null },
    locked: { 1: false, 2: false },
    /* answer: per-slot snapshot of the learner's final pick AND whether
       that slot was individually right, captured BEFORE the reveal
       overwrites a wrong slot. null = nothing to toggle. */
    answer: null,
    view: 'solution',
    lastWrongSig: null
  };
  return window.lomdaState.s7;
}

/* Rendered as a stacked fraction, same markup the option list uses. */
function s7dqFracHTML(v) {
  const p = String(v).split('/');
  return '<span class="dq-frac" aria-hidden="true"><span class="dq-frac__n">' + p[0] +
         '</span><span class="dq-frac__d">' + p[1] + '</span></span>';
}
function s7dqSpokenValue(v) {
  const p = String(v).split('/');
  return p[0] + ' على ' + p[1];
}

const S7DQ_CHECK_SVG = '<img class="dq-check-ico" src="assets/img/dq-correct.svg" width="16" height="16" alt="" aria-hidden="true">';
const S7DQ_WRONG_SVG = '<img class="dq-wrong-ico" src="assets/img/dq-incorrect.svg" width="16" height="16" alt="" aria-hidden="true">';

/* Both popups carry the IDENTICAL body (slide 36 id=26 p2.. and id=28
   p2..); only the title differs. `null` entries are the source's own
   authored EMPTY paragraphs — real blank lines, not empty <p>s.
   The two divisions are complete LTR mathematical expressions, so each
   is isolated whole in a <bdi dir="ltr"> and kept unbreakable: a
   bare-digit group in an RTL paragraph would otherwise reorder. */
const S7DQ_EXPLAIN = [
  'في المربع جميع الأضلاع متساوية، ولذلك فإن المحيط يساوي 4 أضعاف طول الضلع‎.',
  null,
  'لإيجاد طول الضلع من المحيط، نقسم على 4‏‎:',
  '<bdi dir="ltr" class="viq-expr">3 : 4 = ' + s7dqFracHTML('3/4') + '</bdi>',
  '<bdi dir="ltr" class="viq-expr">7 : 4 = ' + s7dqFracHTML('7/4') + '</bdi>',
  null,
  'لذلك، إذا كان محيط المربع بين 3 و-7 سم، فإن طول الضلع هو بين ' +
    '<bdi dir="ltr" class="viq-expr">' + s7dqFracHTML('3/4') + '</bdi> و- ' +
    '<bdi dir="ltr" class="viq-expr">' + s7dqFracHTML('7/4') + '</bdi> سم.'
];
/* Titles are the slide's own (id=26 p1 / id=28 p1). The non-final retry
   state has no authored content on this slide, so it takes the canonical
   fallback title with no invented pedagogical body — the same treatment
   S1-S6 use in this part. */
const S7DQ_POPUP = {
  retry:   { bg: '#ffdbdc', title: 'الإجابة غير صحيحة.', body: ['هل نحاول مرة أخرى؟'] },
  correct: { bg: '#edf8ed', title: 'إجابة صحيحة‎!', body: S7DQ_EXPLAIN },
  wrong2: {
    bg: '#ffdbdc', title: 'أخطأتم. لا بأس، سنتعلم ونتحسن ونواصل.', body: S7DQ_EXPLAIN,
    action: { label: 'الانتقال إلى إجابتي', view: 'answer' },
    answerView: {
      title: 'الإجابة غير صحيحة.',
      body:  ['يمكن مشاهدة الحل الصحيح مرة أخرى‎.'],
      action: { label: 'الإجابة الصحيحة', view: 'solution' }
    }
  }
};

function s7dqBtn(n)  { return document.getElementById('s7dq-btn-' + n); }
function s7dqList(n) { return document.getElementById('s7dq-list-' + n); }
function s7dqSel(n)  { return document.getElementById('s7dq-selected-' + n); }

/* Only one list open at a time, on this screen and across the page. */
function s7dqCloseAll(except) {
  S7DQ.slots.forEach(n => {
    if (n === except) return;
    const l = s7dqList(n), b = s7dqBtn(n);
    if (l) l.classList.add('hidden');
    if (b) b.setAttribute('aria-expanded', 'false');
  });
}

function s7dqToggle(n) {
  const st = s7dqState();
  if (st.done || st.locked[n]) return;
  const b = s7dqBtn(n), l = s7dqList(n);
  if (!b || !l) return;
  const willOpen = l.classList.contains('hidden');
  s7dqCloseAll(n);
  l.classList.toggle('hidden', !willOpen);
  b.setAttribute('aria-expanded', willOpen ? 'true' : 'false');
  /* Opening a dropdown while a non-final retry popup is showing is a
     corrective action, so it dismisses that popup
     (_global-components.md → "First-attempt retry feedback dismissal").
     Cannot fire once the question is done — every slot is disabled. */
  if (willOpen && st.attempts > 0 && !st.done) {
    document.getElementById('s7dq-popup').classList.add('hidden');
  }
  if (willOpen) {
    const cur = st.selections[n];
    const opts = [...l.querySelectorAll('.dq-option')];
    const target = opts.find(o => o.dataset.value === cur) || opts[0];
    if (target) target.focus();
    /* Focus may scroll the list to the current selection, so measure after it. */
    l.onscroll = () => s7dqUpdateMore(l);
    s7dqUpdateMore(l);
  }
}

/* Shows the "more below" fade while options remain under the fold
   (css/style.css → .dq-list.dq-has-more). */
function s7dqUpdateMore(l) {
  l.classList.toggle('dq-has-more', l.scrollTop + l.clientHeight < l.scrollHeight - 2);
}

function s7dqSelect(n, value) {
  const st = s7dqState();
  if (st.done || st.locked[n]) return;
  st.selections[n] = value;
  const b = s7dqBtn(n), l = s7dqList(n), sel = s7dqSel(n);
  sel.innerHTML = s7dqFracHTML(value);
  b.removeAttribute('data-empty');
  b.setAttribute('aria-label', (n === 1 ? 'القيمة الأصغر' : 'القيمة الأكبر') + ': ' + s7dqSpokenValue(value));
  l.querySelectorAll('.dq-option').forEach(o =>
    o.setAttribute('aria-selected', o.dataset.value === value ? 'true' : 'false'));
  l.classList.add('hidden');
  b.setAttribute('aria-expanded', 'false');
  if (st.attempts > 0 && !st.done) {
    document.getElementById('s7dq-popup').classList.add('hidden');
    S7DQ.slots.forEach(k => s7dqMarkSlot(k, null));
  }
  b.focus();
  s7dqUpdateBar();
}

/* Keyboard operation of the list: Enter/Space select, arrows move,
   Escape closes back to the button, Home/End jump. */
function s7dqOptionKey(e, n, value) {
  const l = s7dqList(n);
  if (!l) return;
  const opts = [...l.querySelectorAll('.dq-option')];
  const i = opts.indexOf(e.currentTarget);
  let t = null;
  if (e.key === 'Enter' || e.key === ' ' || e.key === 'Spacebar') { e.preventDefault(); s7dqSelect(n, value); return; }
  else if (e.key === 'ArrowDown') t = opts[Math.min(opts.length - 1, i + 1)];
  else if (e.key === 'ArrowUp')   t = opts[Math.max(0, i - 1)];
  else if (e.key === 'Home')      t = opts[0];
  else if (e.key === 'End')       t = opts[opts.length - 1];
  else if (e.key === 'Escape') { e.preventDefault(); s7dqCloseAll(); s7dqBtn(n).focus(); return; }
  else return;
  e.preventDefault();
  if (t) t.focus();
}

/* One shared document-level listener, registered once: clicking outside
   every wrapper closes any open list. */
document.addEventListener('click', e => {
  if (e.target.closest && e.target.closest('.dq-dropdown-wrap')) return;
  s7dqCloseAll();
});

function s7dqFilled() {
  const st = s7dqState();
  return S7DQ.slots.every(n => st.locked[n] || st.selections[n] !== null);
}
/* Retry-gate comparison shape: a normalized snapshot of EVERY slot, not
   a per-slot flag — resubmitting the identical pair must keep צדקתי?
   disabled (_question-template-defaults.md → "Retry gate"). */
function s7dqSignature() { return JSON.stringify(s7dqState().selections); }

function s7dqUpdateBar() {
  const st = s7dqState();
  const chk = document.getElementById('s7dq-check');
  if (!chk || st.done) return;
  chk.disabled = !s7dqFilled() || s7dqSignature() === st.lastWrongSig;
}

/* Per-slot render. Three states plus null (clear).
   'correct'  — green; leaves an already-correct slot's own selection in
                place, never replaced by the reveal.
   'solution' — green, with the correct value overwriting the display.
   'answer'   — red, showing the learner's own final pick. */
function s7dqMarkSlot(n, state) {
  const b = s7dqBtn(n), sel = s7dqSel(n);
  if (!b) return;
  b.classList.remove('dq-correct', 'dq-wrong');
  const old = b.querySelector('.dq-check-ico, .dq-wrong-ico');
  if (old) old.remove();
  if (state === null) return;
  const st = s7dqState();
  if (state === 'answer') {
    const v = st.answer[n].value;
    sel.innerHTML = v === null ? '' : s7dqFracHTML(v);
    b.classList.add('dq-wrong');
    b.insertAdjacentHTML('beforeend', S7DQ_WRONG_SVG);
  } else {
    if (state === 'solution') sel.innerHTML = s7dqFracHTML(S7DQ.correct[n]);
    b.classList.add('dq-correct');
    b.insertAdjacentHTML('beforeend', S7DQ_CHECK_SVG);
  }
}

function s7dqLock() {
  const st = s7dqState();
  S7DQ.slots.forEach(n => {
    st.locked[n] = true;
    const b = s7dqBtn(n);
    if (b) { b.disabled = true; b.setAttribute('aria-expanded', 'false'); }
    const l = s7dqList(n);
    if (l) l.classList.add('hidden');
  });
}

function s7dqCheck() {
  const st = s7dqState();
  if (st.done) { advanceScreen(); return; }
  if (!s7dqFilled()) return;
  st.attempts++;
  s7dqCloseAll();

  const allRight = S7DQ.slots.every(n => st.selections[n] === S7DQ.correct[n]);
  /* xAPI (2026-11-11): the layer writes Kata's form from the metadata — matching, 'בין[.]3/4[,]לבין[.]7/4' */
  reportQ(['003', 'q2'], allRight, allRight || st.attempts >= S7DQ.maxAttempts,
    { pairs: [{ source: 'بين', target: st.selections[1] }, { source: 'وبين', target: st.selections[2] }] });
  if (allRight) {
    s7dqLock();
    S7DQ.slots.forEach(n => s7dqMarkSlot(n, 'correct'));
    st.answer = null;              // nothing to toggle
    s7dqFinish('correct');
    s7dqShowPopup('correct');
    return;
  }

  if (st.attempts < S7DQ.maxAttempts) {
    /* No per-slot marking on a non-final wrong attempt, and every slot
       stays selectable (DropdownQuestion.md → Reuse QA row 5). */
    st.lastWrongSig = s7dqSignature();
    document.getElementById('s7dq-check').disabled = true;
    s7dqShowPopup('retry');
    return;
  }

  /* Final wrong. Snapshot every slot's own pick AND its own correctness
     BEFORE the reveal overwrites a wrong slot — the only moment the
     learner's actual selection still exists. This is what makes MIXED
     correctness render right in both views. */
  st.answer = {};
  S7DQ.slots.forEach(n => {
    st.answer[n] = { value: st.selections[n], correct: st.selections[n] === S7DQ.correct[n] };
  });
  st.view = 'solution';
  s7dqLock();
  s7dqFinish('wrong2');
  s7dqApplyView();
}

/* Single owner of the resolved-wrong state for every slot and the popup. */
function s7dqApplyView() {
  const st = s7dqState();
  if (st.answer === null) return;
  S7DQ.slots.forEach(n => {
    const a = st.answer[n];
    s7dqMarkSlot(n, a.correct ? 'correct' : (st.view === 'answer' ? 'answer' : 'solution'));
  });
  s7dqShowPopup('wrong2', st.view);
}
/* Pure view switch — never routes through s7dqCheck(), never touches
   attempts or done. */
function s7dqToggleView() {
  const st = s7dqState();
  if (st.answer === null) return;
  st.view = (st.view === 'solution') ? 'answer' : 'solution';
  s7dqApplyView();
}

function s7dqFinish(result) {
  const st = s7dqState();
  st.done = true;
  s7dqLock();
  const chk = document.getElementById('s7dq-check');
  /* DropdownQuestion.md → "Completion label": this template's canonical
     fallback, and the label this part's two SingleChoice screens already
     ship. The slide authors no progression copy. */
  if (chk) { setNavLabel(chk, 'هل نتابع؟'); chk.disabled = false; }
  recordSection(7, result);        /* Question 3, section ב */
}

function s7dqResetInitial() {
  const st = s7dqState();
  st.attempts = 0;
  st.done = false;
  st.answer = null;
  st.view = 'solution';
  st.lastWrongSig = null;
  S7DQ.slots.forEach(n => {
    st.selections[n] = null;
    st.locked[n] = false;
    const b = s7dqBtn(n), sel = s7dqSel(n), l = s7dqList(n);
    if (b) {
      b.disabled = false;
      b.setAttribute('data-empty', 'true');
      b.setAttribute('aria-expanded', 'false');
      b.setAttribute('aria-label', (n === 1 ? 'القيمة الأصغر' : 'القيمة الأكبر') + ': ‎اختاروا كسرًا');
    }
    if (sel) sel.textContent = '';
    if (l) {
      l.classList.add('hidden');
      l.querySelectorAll('.dq-option').forEach(o => o.setAttribute('aria-selected', 'false'));
    }
    s7dqMarkSlot(n, null);
  });
  document.getElementById('s7dq-popup').classList.add('hidden');
  document.getElementById('s7dq-hint-overlay').classList.add('hidden');
  const chk = document.getElementById('s7dq-check');
  if (chk) { setNavLabel(chk, 'هل إجابتي صحيحة؟'); chk.disabled = true; }
  sectionResult[7] = null;
  /* Deliberately NOT resetting the applet: on this screen the marked
     range IS the learner's working, and the iframe is never re-mounted. */
}

function s7dqShowPopup(type, view) {
  const popup = document.getElementById('s7dq-popup');
  if (!popup) return;
  const base = S7DQ_POPUP[type];
  const cfg  = (view === 'answer' && base.answerView)
             ? Object.assign({}, base, base.answerView)
             : base;
  popup.style.background = cfg.bg;
  if (popup.classList.contains('hidden') || type === 'wrong2') resetPopupPosition(popup);
  const titleEl = document.getElementById('s7dq-popup-title');
  titleEl.innerHTML = (type === 'wrong2' && view !== 'answer')
    ? cfg.title + '<br>الإجابة الصحيحة تَظهَر الآن.'
    : cfg.title;
  document.getElementById('s7dq-popup-body').innerHTML =
    (cfg.body || []).map(p => p === null
      ? '<div class="popup-para-gap" aria-hidden="true"></div>'
      : '<p>' + p + '</p>').join('');
  renderPopupAction('s7dq-popup', cfg.action, s7dqToggleView);
  popup.classList.remove('hidden');
  clampPopup(popup);
}

function s7dqOpenHint() {
  xapiRequestedHint('003', 'q2');
  document.getElementById('s7dq-hint-overlay').classList.remove('hidden');
}
function s7dqCloseHint() {
  document.getElementById('s7dq-hint-overlay').classList.add('hidden');
}
function s7dqCloseHintOnBackdrop(e) {
  if (e.target === e.currentTarget) s7dqCloseHint();
}


/* ═══════════════════════════════════════════════════════════
   S8 — ValueInputQuestion, VARIANT A (one value, inline in prose)
   slide 37 · Q3 of 3, and the LAST screen of part 02.

   Spec: 720-templates → ValueInputQuestion.md (Variant A) +
   _question-template-defaults.md + _global-components.md → "View My
   Answer Toggle". Mechanism reused verbatim from this part's approved
   S6 — same template, same inline-in-prose shape, same widget at the
   same configuration; only the copy, the answer and the ids differ.
   ═══════════════════════════════════════════════════════════ */

const S8VIQ = { correct: 6, maxAttempts: 2 };

let s8viqAttempts     = 0;
let s8viqDone         = false;
let s8viqResult       = null;
let s8viqAnswer       = null;
let s8viqView         = 'solution';
let s8viqLastWrongSig = null;

/* Both popups carry the IDENTICAL body (slide 37 id=26 p3.. and id=28
   p3..); only the title differs. `null` entries are the source's own
   authored EMPTY paragraphs (p4, p7, p12) — real blank lines, not empty
   <p>s. The source's p2 is NOT reproduced: it sits between the title and
   the body, and the title lives in the popup header here.
   Every algebraic step is a complete LTR mathematical expression, so each
   is isolated whole in a <bdi dir="ltr"> and kept unbreakable — a
   bare-symbol run in an RTL paragraph would otherwise reorder. That
   includes "5x = 30 /:5", where the trailing /:5 is the Israeli
   divide-both-sides notation and belongs to the same expression. */
const S8VIQ_EXPLAIN = [
  'معلوم أن مجموع طول الضلع ومحيط المربع هو <bdi dir="ltr" class="viq-expr">30</bdi> سم.',
  null,
  'نرمز إلى طول ضلع المربع بـ <bdi dir="ltr" class="viq-expr">x</bdi>.',
  'لذلك، محيط المربع هو <bdi dir="ltr" class="viq-expr">4x</bdi>.',
  null,
  'نكتب المعادلة ونحل‎:',
  '<bdi dir="ltr" class="viq-expr">x + 4x = 30</bdi>',
  '<bdi dir="ltr" class="viq-expr">5x = 30 /:5</bdi>',
  '<bdi dir="ltr" class="viq-expr">x = 6</bdi>',
  null,
  'أي إن طول ضلع المربع هو <bdi dir="ltr" class="viq-expr">6</bdi> سم.'
];
/* Titles are the slide's own (id=26 p1 / id=28 p1). The non-final retry
   state has no authored content on this slide, so it takes the canonical
   fallback with no invented pedagogical body — as S1-S7 do in this part. */
const S8VIQ_POPUP = {
  retry:   { bg: '#ffdbdc', title: 'الإجابة غير صحيحة.', body: ['هل نحاول مرة أخرى؟'] },
  correct: { bg: '#edf8ed', title: 'إجابة صحيحة‎!', body: S8VIQ_EXPLAIN },
  wrong2: {
    bg: '#ffdbdc', title: 'أخطأتم، لا بأس، فكل محاولة تقرّبكم من الفهم‎ أكثر.', body: S8VIQ_EXPLAIN,
    action: { label: 'الانتقال إلى إجابتي', view: 'answer' },
    answerView: {
      title: 'الإجابة غير صحيحة.',
      body:  ['يمكن مشاهدة الحل الصحيح مرة أخرى‎.'],
      action: { label: 'الإجابة الصحيحة', view: 'solution' }
    }
  }
};

function s8viqInput() { return document.getElementById('s8viq-input'); }
function s8viqNormalize(raw) { return String(raw).trim().replace(/,/g, ''); }
function s8viqSignature() { return s8viqNormalize(s8viqInput().value); }
function s8viqFilled() { return s8viqSignature() !== ''; }
function s8viqIsCorrect() { return Number(s8viqSignature()) === S8VIQ.correct; }

function s8viqOnInput() {
  if (s8viqDone) return;
  const el = s8viqInput();
  const cleaned = el.value.replace(/[^0-9,.-]/g, '');
  if (cleaned !== el.value) el.value = cleaned;
  if (s8viqAttempts > 0) {
    document.getElementById('s8viq-popup').classList.add('hidden');
    s8viqMark(null);
  }
  s8viqUpdateBar();
}

function s8viqUpdateBar() {
  const chk = document.getElementById('s8viq-check');
  if (!chk || s8viqDone) return;
  /* Retry gate (_question-template-defaults.md): the comparison shape for
     a single-value screen is the normalized value itself, so a comma-only
     edit is not a change and keeps צדקתי? disabled. */
  chk.disabled = !s8viqFilled() || s8viqSignature() === s8viqLastWrongSig;
}

/* One field, one verdict — the same 16px bare glyph and the same two
   classes S1/S3/S4/S6 use (720-design-guidelines -> components-icons-
   states.md -> Type 1). */
function s8viqMark(outcome) {
  const f = document.getElementById('s8viq-field');
  const ico = document.getElementById('s8viq-ico');
  if (!f) return;
  f.classList.remove('viq-correct', 'viq-wrong');
  if (outcome) f.classList.add(outcome === 'correct' ? 'viq-correct' : 'viq-wrong');
  if (ico) ico.innerHTML = !outcome ? '' : (outcome === 'correct' ? S1VIQ_ICON_CORRECT : S1VIQ_ICON_WRONG);
}

function s8viqCheck() {
  if (s8viqDone) { advanceScreen(); return; }
  if (!s8viqFilled()) return;
  s8viqAttempts++;
  reportQ(['003', 'q3'], s8viqIsCorrect(), s8viqIsCorrect() || s8viqAttempts >= S8VIQ.maxAttempts, String((document.getElementById('s8viq-input') || {}).value).trim());

  if (s8viqIsCorrect()) {
    s8viqMark('correct');
    s8viqResult = 'correct';
    s8viqFinish();
    s8viqShowPopup('correct');
    return;
  }

  if (s8viqAttempts < S8VIQ.maxAttempts) {
    s8viqMark('wrong');
    s8viqLastWrongSig = s8viqSignature();
    document.getElementById('s8viq-check').disabled = true;
    s8viqShowPopup('retry');
    return;
  }

  s8viqAnswer = s8viqInput().value;
  s8viqResult = 'wrong2';
  s8viqView   = 'solution';
  s8viqFinish();
  s8viqApplyView();
}

/* View My Answer Toggle — the final-wrong reveal is destructive (the
   correct value overwrites the learner's own), so it is a toggle, never
   a one-way reveal. */
function s8viqApplyView() {
  if (s8viqAnswer === null) return;
  if (s8viqView === 'answer') {
    s8viqInput().value = s8viqAnswer;
    s8viqMark('wrong');
  } else {
    s8viqInput().value = String(S8VIQ.correct);
    s8viqMark('correct');
  }
  s8viqShowPopup('wrong2', s8viqView);
}
// Pure view switch — never routes through s8viqCheck(), never touches attempts.
function s8viqToggleView() {
  if (s8viqAnswer === null) return;
  s8viqView = (s8viqView === 'solution') ? 'answer' : 'solution';
  s8viqApplyView();
}

function s8viqFinish() {
  s8viqDone = true;
  s8viqInput().disabled = true;
  const chk = document.getElementById('s8viq-check');
  if (chk) { setNavLabel(chk, 'متابعة'); chk.disabled = false; }
  recordSection(8, s8viqResult);   /* Question 3, section ג */
}

function s8viqResetInitial() {
  s8viqAttempts     = 0;
  s8viqDone         = false;
  s8viqResult       = null;
  s8viqAnswer       = null;
  s8viqView         = 'solution';
  s8viqLastWrongSig = null;
  const el = s8viqInput();
  el.value = '';
  el.disabled = false;
  s8viqMark(null);
  document.getElementById('s8viq-popup').classList.add('hidden');
  document.getElementById('s8viq-hint-overlay').classList.add('hidden');
  const chk = document.getElementById('s8viq-check');
  if (chk) { setNavLabel(chk, 'هل إجابتي صحيحة؟'); chk.disabled = true; }
  sectionResult[8] = null;
  /* Deliberately NOT resetting the applet: its src is authored once and
     the iframe is never re-mounted, so the learner's graph state survives
     leaving and re-entering. */
}

function s8viqShowPopup(type, view) {
  const popup = document.getElementById('s8viq-popup');
  if (!popup) return;
  const base = S8VIQ_POPUP[type];
  const cfg  = (view === 'answer' && base.answerView)
             ? Object.assign({}, base, base.answerView)
             : base;
  popup.style.background = cfg.bg;
  if (popup.classList.contains('hidden') || type === 'wrong2') resetPopupPosition(popup);
  const titleEl = document.getElementById('s8viq-popup-title');
  titleEl.innerHTML = (type === 'wrong2' && view !== 'answer')
    ? cfg.title + '<br>الإجابة الصحيحة تَظهَر الآن.'
    : cfg.title;
  document.getElementById('s8viq-popup-body').innerHTML =
    (cfg.body || []).map(p => p === null
      ? '<div class="popup-para-gap" aria-hidden="true"></div>'
      : '<p>' + p + '</p>').join('');
  renderPopupAction('s8viq-popup', cfg.action, s8viqToggleView);
  popup.classList.remove('hidden');
  clampPopup(popup);
}

function s8viqOpenHint() {
  xapiRequestedHint('003', 'q3');
  document.getElementById('s8viq-hint-overlay').classList.remove('hidden');
}
function s8viqCloseHint() {
  document.getElementById('s8viq-hint-overlay').classList.add('hidden');
}
function s8viqCloseHintOnBackdrop(e) {
  if (e.target === e.currentTarget) s8viqCloseHint();
}

/* ═══════════════════════════════════════════════════════════
   xAPI seams — read at call time by unit-js/20-xapi.js, 22-scoring.js and 50-loader.js.
   Ids must match metadata/methodica-ar-math-ratio-03-02.json byte-for-byte (the loader
   checks XAPI_COMP_ID against METADATA.id on every load; _test/verify-static.js checks the map).
   ═══════════════════════════════════════════════════════════ */
const XAPI_COMP_SLUG     = 'methodica-ar-math-ratio-03-02';
const XAPI_COMP_ID       = XAPI_ID_PREFIX + XAPI_COMP_SLUG + '/';
const XAPI_METADATA_FILE = '../metadata/' + XAPI_COMP_SLUG + '.json';
/* Items carrying a code-graded question (each declared question is reported by its screen). */
const XAPI_EVAL_ITEMS    = { '001': 1, '002': 1, '003': 1 };
/* Items scored as ONE question: success only when every section is right (QA F-6, 22-scoring.js).
   QUESTION_SECTIONS maps 1:1 onto the three items (questionResult = AND). */
const XAPI_ITEM_ALL_CORRECT = { '001': 1, '002': 1, '003': 1 };
/* screen → [item suffix, page within the item]; null = no catalogue item. Exactly TOTAL_SCREENS
   keys, 0-based (keep it a literal object — the static check reads it). */
const SCREEN_TO_SUBCONTENT = {
  0: null,
  1: ['001', 1],
  2: ['001', 2],
  3: ['001', 3],
  4: ['002', 1],
  5: ['002', 2],
  6: ['003', 1],
  7: ['003', 2],
  8: ['003', 3]
};

/* The gate S0 promises: "2 of 3 questions" (masteryState). Score = questions correct / 3. */
function partGateResult() {
  const m = masteryState();
  return { success: !!m.passed, score: { scaled: m.correctQuestionCount / m.results.length } };
}
function lastScreenButton() { return currentScreen === PART_LAST ? document.getElementById('s8viq-check') : null; }

/* ═══════════════════════════════════════════════════════════
   Resume painters — unit-js/45-painters.js reads this registry.
   What each screen saves (vars + the fields that hold state) and how a screen restored from the
   state document is drawn again on the fresh markup: init() is the screen's own ResetInitial,
   paint() replays the template's own terminal look (Mark / Finish / ShowPopup / ApplyView) or the
   retry look, without counting an attempt and without reporting. In-session re-entry still goes
   through resetScreenState(), which now resets only a pristine screen.
   ═══════════════════════════════════════════════════════════ */
const PART_STATE_VARS = ['sectionResult', 'practiceProgress'];
const PART_SCREENS = {
  1: {
    vars: ['s1viqAttempts', 's1viqDone', 's1viqResult', 's1viqAnswer', 's1viqView', 's1viqLastWrongSig'],
    inputs: ['s1viq-input-right', 's1viq-input-left'],
    pristine: () => !s1viqAttempts && !s1viqDone && resumeFieldsEmpty(PART_SCREENS[1].inputs),
    init: s1viqResetInitial,
    paint: () => {
      if (s1viqDone) {
        s1viqDone = false;
        if (s1viqResult === 'correct') { s1viqMark('correct'); s1viqFinish(); s1viqShowPopup('correct', 'solution'); }
        else { s1viqFinish(); s1viqView = 'solution'; s1viqApplyView(); }
        return;
      }
      s1viqUpdateBar();
      if (s1viqAttempts && s1viqSignature() === s1viqLastWrongSig) s1viqMark('wrong');
    }
  },
  2: {
    vars: ['s2scSelected', 's2scAttempts', 's2scDone', 's2scResult', 's2scLastWrong'],
    pristine: () => !s2scSelected && !s2scAttempts && !s2scDone,
    init: s2scResetInitial,
    paint: () => {
      const done = s2scDone, sel = s2scSelected;
      s2scDone = false;
      if (sel) s2scSelect(sel);                    // selection look + the retry gate
      if (done) {
        s2scOpt(sel).classList.add(s2scResult === 'correct' ? 'correct' : 'wrong');
        if (s2scResult !== 'correct') s2scOpt(S2SC.correctId).classList.add('correct');
        s2scFinish();
        s2scShowPopup(s2scResult);
      } else if (s2scAttempts && sel === s2scLastWrong) {
        s2scOpt(sel).classList.add('wrong');
      }
    }
  },
  3: {
    vars: ['s3viqAttempts', 's3viqDone', 's3viqResult', 's3viqAnswer', 's3viqView', 's3viqLastWrongSig'],
    inputs: ['s3viq-input-right', 's3viq-input-left'],
    pristine: () => !s3viqAttempts && !s3viqDone && resumeFieldsEmpty(PART_SCREENS[3].inputs),
    init: s3viqResetInitial,
    paint: () => {
      if (s3viqDone) {
        s3viqDone = false;
        if (s3viqResult === 'correct') { s3viqMark('correct'); s3viqFinish(); s3viqShowPopup('correct', 'solution'); }
        else { s3viqFinish(); s3viqView = 'solution'; s3viqApplyView(); }
        return;
      }
      s3viqUpdateBar();
      if (s3viqAttempts && s3viqSignature() === s3viqLastWrongSig) s3viqMark('wrong');
    }
  },
  4: {
    vars: ['s4viqAttempts', 's4viqDone', 's4viqResult', 's4viqAnswer', 's4viqView', 's4viqLastWrongSig'],
    inputs: ['s4viq-input-plain', 's4viq-input-uphill'],
    pristine: () => !s4viqAttempts && !s4viqDone && resumeFieldsEmpty(PART_SCREENS[4].inputs),
    init: s4viqResetInitial,
    paint: () => {
      if (s4viqDone) {
        s4viqDone = false;
        if (s4viqResult === 'correct') { s4viqMark(s4viqFieldStates(s4viqValues())); s4viqFinish(); s4viqShowPopup('correct', 'solution'); }
        else { s4viqFinish(); s4viqView = 'solution'; s4viqApplyView(); }
        return;
      }
      s4viqUpdateBar();
      if (s4viqAttempts && s4viqSignature() === s4viqLastWrongSig) s4viqMark(s4viqFieldStates(s4viqValues()));
    }
  },
  5: {
    vars: ['s5scSelected', 's5scAttempts', 's5scDone', 's5scResult', 's5scLastWrong'],
    pristine: () => !s5scSelected && !s5scAttempts && !s5scDone,
    init: s5scResetInitial,
    paint: () => {
      const done = s5scDone, sel = s5scSelected;
      s5scDone = false;
      if (sel) s5scSelect(sel);                    // selection look + the retry gate
      if (done) {
        s5scOpt(sel).classList.add(s5scResult === 'correct' ? 'correct' : 'wrong');
        if (s5scResult !== 'correct') s5scOpt(S5SC.correctId).classList.add('correct');
        s5scFinish();
        s5scShowPopup(s5scResult);
      } else if (s5scAttempts && sel === s5scLastWrong) {
        s5scOpt(sel).classList.add('wrong');
      }
    }
  },
  6: {
    vars: ['s6viqAttempts', 's6viqDone', 's6viqResult', 's6viqAnswer', 's6viqView', 's6viqLastWrongSig'],
    inputs: ['s6viq-input'],
    pristine: () => !s6viqAttempts && !s6viqDone && resumeFieldsEmpty(PART_SCREENS[6].inputs),
    init: s6viqResetInitial,
    paint: () => {
      if (s6viqDone) {
        s6viqDone = false;
        if (s6viqResult === 'correct') { s6viqMark('correct'); s6viqFinish(); s6viqShowPopup('correct', 'solution'); }
        else { s6viqFinish(); s6viqView = 'solution'; s6viqApplyView(); }
        return;
      }
      s6viqUpdateBar();
      if (s6viqAttempts && s6viqSignature() === s6viqLastWrongSig) s6viqMark('wrong');
    }
  },
  7: {
    vars: ['window.lomdaState.s7'],
    pristine: () => { const st = s7dqState(); return !st.attempts && !st.done && S7DQ.slots.every(k => st.selections[k] === null); },
    init: s7dqResetInitial,
    paint: () => {
      const st = s7dqState();
      const done = st.done, sels = Object.assign({}, st.selections);
      st.done = false;
      S7DQ.slots.forEach(k => { st.locked[k] = false; });
      S7DQ.slots.forEach(k => { if (sels[k] !== null) s7dqSelect(k, sels[k]); });
      if (document.activeElement && document.activeElement.blur) document.activeElement.blur();
      if (done) {
        s7dqLock();
        if (st.answer === null) { S7DQ.slots.forEach(k => s7dqMarkSlot(k, 'correct')); s7dqFinish('correct'); s7dqShowPopup('correct', 'solution'); }
        else { s7dqFinish('wrong2'); st.view = 'solution'; s7dqApplyView(); }
      }
    }
  },
  8: {
    vars: ['s8viqAttempts', 's8viqDone', 's8viqResult', 's8viqAnswer', 's8viqView', 's8viqLastWrongSig'],
    inputs: ['s8viq-input'],
    pristine: () => !s8viqAttempts && !s8viqDone && resumeFieldsEmpty(PART_SCREENS[8].inputs),
    init: s8viqResetInitial,
    paint: () => {
      if (s8viqDone) {
        s8viqDone = false;
        if (s8viqResult === 'correct') { s8viqMark('correct'); s8viqFinish(); s8viqShowPopup('correct', 'solution'); }
        else { s8viqFinish(); s8viqView = 'solution'; s8viqApplyView(); }
        return;
      }
      s8viqUpdateBar();
      if (s8viqAttempts && s8viqSignature() === s8viqLastWrongSig) s8viqMark('wrong');
    }
  }
};
