'use strict';

/* ═══════════════════════════════════════════════════════════
   js/main.js  —  methodica-ar-math-ratio-03-03
   720 reference engine (canvas 1280×710) + this component's screens.

   Source: the approved production script
   יעד 1.3_יחס וביטויים אלגבריים_להפקה_06092026.pptx, two levels above
   this part's folder. Slide identity re-derived from a FRESH extraction
   each session, from presentation.xml's <p:sldIdLst> r:id values
   resolved through ppt/_rels/presentation.xml.rels (presentation order
   38 → slide38.xml, 39 → slide39.xml, 40 → slide40.xml, 41 →
   slide41.xml, 42 → slide42.xml, 43 → slide43.xml, 44 → slide44.xml,
   45 → slide45.xml, 46 → slide46.xml — identity
   mapping in this deck, but derived rather than assumed). Presenter
   notes resolved through each slide's OWN .rels: slide39 →
   notesSlides/notesSlide35.xml, slide40 → notesSlide36.xml, slide41 →
   notesSlide37.xml, slide42 → notesSlide38.xml, slide43 →
   notesSlide39.xml, slide44 → notesSlide40.xml, slide45 →
   notesSlide41.xml, slide46 → notesSlide42.xml — part-name
   numbering is NOT slide numbering (pptx-script-extraction).
   Component dividers bound this part at BOTH ends — production
   dividers, not learner-facing screens, and not built:
     slide 38 — "רכיב 3" / "תרגול בסיסי וסטנדרטי ב" /
                "methodica-ar-math-ratio-03-03"   opens this part
     slide 50 — "רכיב 4" / "methodica-ar-math-ratio-03-04"   closes it
   So this part owns slides 39-49, and S5 (slide 44) is NOT its last
   screen. Slides 45-49 are built as S6-S10:
     45 — TransitionScreen. notesSlide41: "שם המסך בפיגמה:
          TransitionScreen", "הדמות עם מדלייה". Copy: "יופי של עבודה!" +
          "הנה עוד 2 תרגילים ברמת קושי גבוהה יותר" + a character speech
          bubble "ברור לי שתצליחו!".
          (S6_CHARACTER below carries its clips.)
     46-49 — the two harder questions slide 45 announces.
   Resolved this pass from presentation.xml's sldIdLst through
   presentation.xml.rels, and every slide's own cached number
   placeholder agrees with its resolved position.

   SOURCE-CURRENTNESS: the deck's own bytes were re-hashed at the start
   of the S1 pass and are unchanged from the S0 pass (md5
   1dc38e9f7061157af79fbf1e9c9c926b). Note that opening the file in
   PowerPoint REWRITES it, so every render is taken from a throwaway
   copy and every extraction from a copy PowerPoint has never touched —
   otherwise the deck appears to change under you.

   PLANNED SERIES (from the deck, for context — S0-S2 are built):
     slide 44 → question 3 sub-part ב (DropdownQuestion)

   Screen roster (built so far):
     S0 — TransitionScreen, practice-stage intro (slide 39).
     S1 — ValueInputQuestion (local ratio variant), practice question 1
          sub-part א (slide 40).
     S2 — ValueInputQuestion (Variant A, single value), practice
          question 1 sub-part ב (slide 41) — resolving it resolves
          progress item 1.
     S3 — ValueInputQuestion (local ratio variant) + a DECORATIVE
          content image (no zoom — producer decision, see index.html),
          practice QUESTION 2 (slide 42) — a single-screen question, so
          it resolves progress item 2 on its own.
     S4 — MultipleChoiceQuestion (text pills, 2 of 3 correct) + a
          decorative content image, practice question 3 sub-part א
          (slide 43). Its feedback is the ORDINARY popup: the
          scroll-gated variant was removed for this screen by producer
          decision (see S4MC_POPUP). Does not resolve progress item 3 on
          its own — slide 44 is sub-part ב.
     S5 — DropdownQuestion (single slot, 3 options) + the same
          decorative content image and a character speech bubble,
          practice question 3 sub-part ב (slide 44). Resolving THIS
          screen is what resolves progress item 3.
     S6 — TransitionScreen, the part's SECOND one (slide 45): closes
          the basic practice block and opens the harder one that
          slides 46-49 carry. Character speech bubble, medal pose.
          Ungated, like every TransitionScreen.
     S7 — ValueInputQuestion (Variant A, TWO independent values, Case A
          scoring) + a decorative content image, slide 46 — the FIRST
          question of the SECOND practice block, which has its own
          2-item progress series (practiceProgress2).

   S0 ported from the same-lomda sibling methodica-ar-math-ratio-03-02's
   own S0 — the identical screen role (practice-stage kickoff
   transition, slide 29 in the same deck), which in turn ported it from
   the approved Math sibling methodica-math-percent-01-02. Contract:
   720-templates → references/TransitionScreen.md +
   references/_global-components.md.
   Engine plumbing (scaleApp / goTo / resetScreenState / advanceScreen /
   goBack / keyboard nav / setNavLabel / dev bridge) ported from
   methodica-ar-math-ratio-03-02, so all three parts of this lomda behave
   the same. Cross-part Back uses this part's own js/partnav.js
   (PartNav) — the same engine body, byte-identical to 03-02's copy.

   Parts 01 and 02 are read-only references here (they are being built
   in parallel sessions); nothing in this file touches or assumes an
   unbuilt change in either.
   ═══════════════════════════════════════════════════════════ */

/* ─── Constants ─────────────────────────────────────────── */
/* S0-S5 — the screens built so far. This part is NOT complete at S5:
   its slide range is 39-49, established by the component dividers at
   both ends (see the divider note below). S6 (slide 45) and S7
   (slide 46) and S8-S10 (slides 47-49) are built, so every slide of
   this part is now produced. MUST be
   bumped as each is added.

   END-OF-PART BOUNDARY (resolved 2026-09-14 — see crossFromCheckpoint()):
   goTo() bounds-checks n against TOTAL_SCREENS, so advancing past the
   last screen through goTo() alone would be a silent no-op —
   720-templates → _global-components.md → "Advance gates" is explicit
   that "no error" is never evidence that an end-of-part boundary is
   implemented. The forward crossing is implemented explicitly: S10
   goes through crossFromCheckpoint() and PartNav (NEXT_PART, below),
   never through goTo().
   Do not paper over it with an invented PartNav.forward() target. */
const TOTAL_SCREENS = 11;


/* ─── S0 companion-character assets ─────────────────────────
   Explicit ID → asset mapping. The supplied filenames
   ("Character_1_Proud.mp4" / "Character_2_Proud.mp4") don't follow the
   character-N-[state] convention, and resolving by filename or colour
   guessing is forbidden regardless — 720-templates →
   _global-components.md → Companion character system.

   LOGICAL MAPPING confirmed against methodica-ar-math-ratio-03-01's own
   CHARACTER_ASSETS (same lomda, same two companions, read this
   session):
       character-1 = the ORANGE/CORAL round one-eyed character
       character-2 = the TALL YELLOW character with the purple headband

   Per-instance asset map (TransitionScreen.md → "Companion-character
   asset resolution": each TransitionScreen instance keeps its own map,
   not a lomda-wide one) — this is this part's only TransitionScreen.

   ASSET PROVENANCE: the clips implement slide 39's presenter note
   ("הדמות עם מלא דפים מלאים בתרגילים, נראית מותשת אבל מרוצה מעצמה")
   and were supplied by the producer directly into this part's own
   assets/video/ on 2026-09-09. Character_2's file arrived with a
   doubled extension ("...mp4.mp4") and was renamed to the name the
   producer stated. They are this part's own local copies — nothing here
   links at runtime to Parts 01/02 or to the source-material folders.

   No `fit` shim — a DELIBERATE, measured decision, not an omission.
   The full measurement record lives on `.s1-char-media` in
   css/style.css; the short version: both clips are 1280×720 / 4.01s
   full SCENES (companion + a spread of exercise pages filling the
   frame), not a character floating on a blank backdrop, so there is no
   transparent-margin mismatch of the kind TransitionScreen.md → "C"
   exists to correct, and a plain contain-fit renders both at the same
   physical scale. The ~14% difference that remains between the two
   clips' character ink heights is the GENUINE height difference
   between a short round companion and a tall thin one — the same
   relationship methodica-ar-math-ratio-03-01 measured on the standing
   PNGs (73.37% vs 94.74% of canvas height) and also declined to
   equalize. Scaling one up to match the other would misrepresent the
   characters. Re-measure if either asset is ever replaced.

   No `freezeAt`: both clips were played to completion during build
   verification and hold a meaningful final frame, so step 1 of the
   end-state hierarchy applies and no end-handler is wired at all.
   The prohibited `freezeVideoOnEnd()` / `currentTime = 0` pattern is
   deliberately absent. */
/* ⚠ SLIDE 45 IS BLOCKED ON A MISSING COMPANION CLIP — recorded here
   2026-09-10 because this is the map the next pass will extend.

   Slide 45 is this part's second TransitionScreen and its presenter
   note (notesSlide41, resolved through slide45.xml.rels → rId3) asks
   for "הדמות עם מדלייה" — the character WITH A MEDAL. The delivered
   assets hold only ONE of the two clips that pose needs:

     assets/video/Character_2_Medal.mp4   present
     Character_1_Medal.mp4                ABSENT

   Searched for exhaustively before reporting: the whole project tree
   (the only "*Medal*" hit is Character_2_Medal.mp4), every sibling
   part's assets/video, and ppt/media inside the deck itself, which
   embeds no video at all. This part's own clips are
   Character_1_Proud / Character_2_Proud (slide 39) plus
   Character_2_Medal, so the missing file is character-1's counterpart
   for slide 45 and nothing else.

   Why this blocks rather than degrades: DEFAULT_CHARACTER is
   'character-1', so the DEFAULT learner is exactly the one with no
   clip. It also blocks the screen's speech bubble, which this template
   anchors to the character's own measured INK bounds (Part 01's
   --ink-top/--ink-bottom mechanism) — bounds that cannot be measured
   from a clip that does not exist.

   Do NOT silently substitute Character_1_Proud.mp4: the script names
   the medal pose, and the two poses are different content. Awaiting
   the producer's decision.
   Note the sibling parts show that the pose name may legitimately
   DIFFER per character for one screen (03-04 ships
   Character_1_Telescope + Character_2_Binoculars), so the missing file
   need not be named "Medal" — any character-1 clip authored for slide
   45 satisfies this. */
/* Slide 45's own pose. Same two logical ids and the same resolution
   path as S0_CHARACTER — only the clips differ. Both were measured
   from a decoded frame: Character_1_Medal 720x720 / 4.00s (the orange
   round one-eyed character, medal round its neck) and
   Character_2_Medal 960x960 / 4.04s (the tall yellow character with
   the purple headband). */
const S6_CHARACTER = {
  'character-1': { video: 'assets/video/Character_1_Medal.mp4' },
  'character-2': { video: 'assets/video/Character_2_Medal.mp4' }
};

const S0_CHARACTER = {
  'character-1': {
    video: 'assets/video/Character_1_Proud.mp4',
    alt: 'الشخصية البرتقالية المستديرة ذات العين الواحدة، محاطة بأوراق مليئة بالتمارين، تبدو منهكة لكنها راضية عن نفسها'
  },
  'character-2': {
    video: 'assets/video/Character_2_Proud.mp4',
    alt: 'الشخصية الصفراء الطويلة ذات عصابة الرأس البنفسجية، محاطة بأوراق مليئة بالتمارين، تبدو منهكة لكنها راضية عن نفسها'
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
  // Decorative only — never in the accessibility tree. The map's `alt`
  // is production/QA documentation of what each clip shows, not an
  // exposed accessible name: this media conveys no instructional
  // content, and the screen's meaning lives entirely in .s1-text.
  vid.setAttribute('aria-hidden', 'true');
  /* Per-character fit shim. ALWAYS written, including the no-op values,
     so switching characters can never leave a previous one's scale
     behind on the element (neither of this part's entries carries
     `fit` — see the note on S0_CHARACTER above). */
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
       without its feedback (_global-components.md → "Reopening feedback
       on return to a resolved screen"). s1viqDone is the only flag
       involved — there is deliberately no second "popup was open on
       exit" flag. */
    if (!s1viqDone) { if (screenPristine(1)) s1viqResetInitial(); }
    else if (s1viqResult === 'correct') s1viqShowPopup('correct', 'solution');
    else { s1viqView = 'solution'; s1viqApplyView(); }
    syncPracticeProgressNav(document.getElementById('s1'));
  }
  if (n === 2) {
    /* Same resume/reopen contract as S1: a resolved question never hands
       out a fresh attempt, and never comes back without its feedback. */
    if (!s2viqDone) { if (screenPristine(2)) s2viqResetInitial(); }
    else if (s2viqResult === 'correct') s2viqShowPopup('correct', 'solution');
    else { s2viqView = 'solution'; s2viqApplyView(); }
    syncPracticeProgressNav(document.getElementById('s2'));
  }
  if (n === 3) {
    /* Same resume/reopen contract as S1/S2. */
    if (!s3viqDone) { if (screenPristine(3)) s3viqResetInitial(); }
    else if (s3viqResult === 'correct') s3viqShowPopup('correct', 'solution');
    else { s3viqView = 'solution'; s3viqApplyView(); }
    syncPracticeProgressNav(document.getElementById('s3'));
  }
  if (n === 4) {
    /* Same resume/reopen contract as the value-input screens: a resolved
       question never hands out a fresh attempt and never comes back
       without its feedback. This screen's feedback is ungated (see
       S4MC_POPUP), so a reopen restores it with Continue enabled. */
    if (!s4mcDone) { if (screenPristine(4)) s4mcResetInitial(); }
    else s4mcShowPopup(s4mcResult);
    syncPracticeProgressNav(document.getElementById('s4'));
  }
  if (n === 5) {
    /* Same resume/reopen contract as every other question screen here: a
       resolved question never hands out a fresh attempt and never comes
       back without its feedback. DropdownQuestion.md → "On re-entry":
       a resolved final-wrong screen always returns on the SOLUTION view,
       and a first-try-correct screen has nothing to toggle. */
    const st = s5dqState();
    if (!st.done) { if (screenPristine(5)) s5dqResetInitial(); }
    else if (st.answer === null) s5dqShowPopup('correct');
    else { st.view = 'solution'; s5dqApplyView(); }
    syncPracticeProgressNav(document.getElementById('s5'));
  }
  if (n === 7) {
    /* Same resume/reopen contract as every other question screen here. */
    if (!s7viqDone) { if (screenPristine(7)) s7viqResetInitial(); }
    else if (s7viqAnswer === null) s7viqShowPopup('correct');
    else { s7viqView = 'solution'; s7viqApplyView(); }
    syncPracticeProgress2Nav(document.getElementById('s7'));
  }
  if (n === 8) {
    /* Same resume/reopen contract as every other question screen here.
       s8ddqEnter() owns the whole branch because a drag screen has to
       re-render its board before anything else can be decided. */
    s8ddqEnter();
  }
  if (n === 9) {
    s9ddqEnter();
  }
  if (n === 10) {
    /* Same resume/reopen contract as every other question screen here. */
    if (!s10viqDone) { if (screenPristine(10)) s10viqResetInitial(); }
    else if (s10viqAnswer === null) s10viqShowPopup('correct');
    else { s10viqView = 'solution'; s10viqApplyView(); }
    syncPracticeProgress2Nav(document.getElementById('s10'));
  }
  if (n === 6) {
    /* A TransitionScreen holds no answer state — the only per-entry work
       is re-resolving the companion character, exactly as S0 does
       (_global-components.md → Companion character system). */
    applyCompanionCharacter(S6_CHARACTER, 's6-char-video');
  }
}

/* ─── Forward navigation ─────────────────────────────────────
   Gates live HERE, not only in a Continue button's handler: the engine
   binds ArrowLeft directly to advanceScreen(), so a gate that sat only
   in an onclick would be silently bypassable by keyboard
   (720-templates → _global-components.md → "Advance gates").
   S0 has NO gate: a TransitionScreen carries no interaction to resolve,
   and slide 39 mocks Continue in its enabled state (720-templates →
   TransitionScreen.md → "A. Structural contract"). There is therefore
   deliberately no `if (currentScreen === 0 && …) return;` line — the
   absence is the contract, not a missing gate.
   S1 gates on RESOLVED, never on correctness — a final wrong attempt
   still resolves the question and still lets the learner move on
   (_global-components.md → "Gate on the resolved state, not on
   correctness"). Keyboard ArrowLeft hits this same function, so the
   gate is not bypassable by keyboard.
   See the TOTAL_SCREENS note above for the pending end-of-part
   boundary. */
function advanceScreen() {
  if (currentScreen === 1 && !s1viqDone) return;
  if (currentScreen === 2 && !s2viqDone) return;
  if (currentScreen === 3 && !s3viqDone) return;
  if (currentScreen === 4 && !s4mcDone) return;
  if (currentScreen === 5 && !s5dqState().done) return;
  /* S7 and S10 both carry the Long-feedback scroll gate on their terminal
     panel (_global-components.md → "Long-feedback variant"). Done alone
     is not enough: a learner can be Done and still have an unread,
     overflowing panel locking Continue. GateOk is the same flag
     s7viqCheck()/s10viqCheck() already gate the Continue button on
     (2026-09-24 fix — 03-04's advanceScreen() already did this; S7/S10
     did not, so ArrowLeft could bypass the button's own gate). */
  if (currentScreen === 7 && (!s7viqDone || !s7viqGateOk)) return;
  if (currentScreen === 8 && !s8ddqDone) return;
  if (currentScreen === 9 && !s9ddqDone) return;
  if (currentScreen === 10 && (!s10viqDone || !s10viqGateOk)) return;
  /* Decision points. S10 closes the part and goes through its mastery
     checkpoint rather than goTo(), which would be swallowed by
     TOTAL_SCREENS. */
  if (currentScreen === 10) { crossFromCheckpoint(2); return; }
  /* Leaving the last question of the first bundle IS the entry gate the
     opening screen announces ("ענו נכון על 2 שאלות ומעלה כדי להתקדם"):
     2 or 3 correct → S6 and on into the harder questions; 0 or 1 correct →
     the learner stays on S5 with המשך enabled and nothing advances
     (producer decision 2026-09-30; what a failed learner is offered
     instead belongs to the external layer, see MASTERY CHECKPOINTS). */
  if (currentScreen === 5)  { if (mastery1State().passed) goTo(6); else endAtCheckpoint1(); return; }
  goTo(currentScreen + 1);
}

/* Back returns to the previous valid predecessor in the learner's
   actual route. From this part's FIRST screen, Back crosses the part
   boundary backwards via PartNav instead of no-opping — with no
   recorded crossing (this part opened directly, e.g. in dev) that is
   itself a no-op, which is why the button stays hidden until
   PartNav.mountBack() confirms a real trail entry. The destination is
   never hardcoded: it is whichever part the learner actually arrived
   from (720-templates → _global-components.md → "Back navigation in a
   branched journey"). */
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
  /* A pointer-down that starts inside a scrollable popup body must NOT
     begin a drag, or the learner cannot scroll a scroll-gated popup by
     dragging its scrollbar or swiping its text (_global-components.md →
     "Long-feedback variant" rule 11). The header and the rest of the
     popup chrome stay draggable. */
  if (e.target && e.target.closest && e.target.closest('.scq-popup-scroll')) return;
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
   popup, which has no dismissal path at all
   (_global-components.md → "No manual close"). */
document.addEventListener('keydown', e => {
  if (e.key !== 'Escape') return;
  document.querySelectorAll('[id$="-hint-overlay"]').forEach(el => {
    if (!el.classList.contains('hidden')) el.classList.add('hidden');
  });
});

/* ═══════════════════════════════════════════════════════════
   GLOBAL COMPONENT — Pointer-based drag and drop.
   Ported verbatim (only comments reflowed) from the approved
   methodica-science-mass-measure-01-01, where it drives S3, S11 and
   S16 — S16 being DragAndDropQuestion.md's own Image Top reference
   implementation, which is the layout this part's slide 47 uses.

   Native HTML5 drag is deliberately NOT used: it does not fire on
   touch, and these lomdot ship to tablets in both orientations.
   Pointer events give mouse, pen and touch from one code path.

   The floating ghost is a CLONE of the source positioned in VIEWPORT
   coordinates, so it must be scaled by #app's own transform to line up
   with the canvas it was lifted from — that is what getAppScale() is
   for. Drop hit-testing goes through elementFromPoint(), i.e. the
   browser's real hit-testing, not a rectangle comparison.

     const dnd = createPointerDnd({
       canDrag:  (dragId, elem) => boolean,
       onPick:   (dragId, elem) => void,
       onDrop:   (dragId, targetId) => void,
       onCancel: (dragId) => void,
     });
     dnd.attachSource(elem, dragId);   // idempotent
     dnd.attachTarget(elem, targetId);
   ═══════════════════════════════════════════════════════════ */
/* Properties the preview must inherit from the source card itself.
   Deliberately NOT a blanket copy of the root's computed style: border,
   background and opacity stay under the .pointer-drag-ghost rule, which
   is what paints the template's "Selected" state (a Subject-coloured
   stroke) while the card is in flight. */
const PDRAG_ROOT_PROPS = [
  'display', 'flex-direction', 'align-items', 'justify-content', 'gap',
  'padding-top', 'padding-right', 'padding-bottom', 'padding-left',
  'box-sizing', 'overflow', 'border-radius',
  'font-family', 'font-size', 'font-weight', 'font-style', 'font-variant',
  'line-height', 'letter-spacing', 'word-spacing',
  'color', 'text-align', 'text-transform', 'white-space',
  'direction', 'unicode-bidi'
];

/* Make `clone` render identically to `src` even though it lives outside
   `src`'s styled ancestors. The root gets the allowlist above; every
   DESCENDANT gets its full resolved style, since nothing in the ghost
   rule styles descendants and the label is exactly what was breaking.
   Both trees are walked in the same order, which cloneNode(true)
   guarantees. */
function pdragCloneStyles(src, clone) {
  const rootCs = getComputedStyle(src);
  PDRAG_ROOT_PROPS.forEach(p => clone.style.setProperty(p, rootCs.getPropertyValue(p)));

  const srcKids   = src.querySelectorAll('*');
  const cloneKids = clone.querySelectorAll('*');
  for (let i = 0; i < srcKids.length && i < cloneKids.length; i++) {
    const cs = getComputedStyle(srcKids[i]);
    let css = '';
    for (let j = 0; j < cs.length; j++) {
      const prop = cs[j];
      css += prop + ':' + cs.getPropertyValue(prop) + ';';
    }
    cloneKids[i].style.cssText = css;
  }
}

function createPointerDnd(opts) {
  const targets = new Map();   // elem → targetId
  let active = null;

  function getAppScale() {
    const app = document.getElementById('app');
    const m = app && app.style.transform.match(/scale\(([^)]+)\)/);
    return m ? parseFloat(m[1]) : 1;
  }

  function attachSource(elem, dragId) {
    if (!elem || elem.dataset.pdragAttached === '1') {
      if (elem) elem.dataset.pdragId = dragId;
      return;
    }
    elem.dataset.pdragId = dragId;
    elem.dataset.pdragAttached = '1';
    elem.style.touchAction = 'none';
    elem.addEventListener('pointerdown', onSourceDown);
  }

  function attachTarget(elem, targetId) {
    if (!elem) return;
    targets.set(elem, targetId);
  }

  function onSourceDown(e) {
    if (e.button !== undefined && e.button !== 0 && e.pointerType === 'mouse') return;
    const src    = e.currentTarget;
    const dragId = src.dataset.pdragId;
    if (!dragId) return;
    if (opts.canDrag && !opts.canDrag(dragId, src)) return;
    e.preventDefault();

    const rect  = src.getBoundingClientRect();
    const scale = getAppScale();

    const ghost = src.cloneNode(true);
    /* Strip ids from the clone so the document never holds duplicates. */
    ghost.removeAttribute('id');
    ghost.querySelectorAll('[id]').forEach(n => n.removeAttribute('id'));
    /* The clone is about to be parented to <body>, where none of the
       screen-scoped rules reach it. Carry the source's resolved styles
       across FIRST, so what follows only has to position it. */
    pdragCloneStyles(src, ghost);
    ghost.style.position        = 'fixed';
    ghost.style.left            = rect.left + 'px';
    ghost.style.top             = rect.top  + 'px';
    /* rect is POST-scale; divide the scale back out so the ghost keeps
       the source's own authored size once its transform is applied. */
    ghost.style.width           = (rect.width  / scale) + 'px';
    ghost.style.height          = (rect.height / scale) + 'px';
    ghost.style.boxSizing       = 'border-box';
    ghost.style.margin          = '0';
    ghost.style.pointerEvents   = 'none';
    ghost.style.zIndex          = '9999';
    ghost.style.opacity         = '0.92';
    ghost.style.transform       = 'scale(' + scale + ')';
    ghost.style.transformOrigin = 'top left';
    ghost.classList.remove('dragging');
    ghost.classList.add('pointer-drag-ghost');
    document.body.appendChild(ghost);

    active = {
      dragId, srcElem: src, ghost,
      pointerId: e.pointerId,
      offX: e.clientX - rect.left,
      offY: e.clientY - rect.top,
      currentTarget: null
    };
    src.classList.add('dragging');
    if (opts.onPick) opts.onPick(dragId, src);

    document.addEventListener('pointermove',   onMove, { passive: false });
    document.addEventListener('pointerup',     onUp);
    document.addEventListener('pointercancel', onUp);
  }

  function findTargetUnder(clientX, clientY) {
    const el = document.elementFromPoint(clientX, clientY);
    if (!el) return null;
    for (const [tElem, tId] of targets) {
      if (tElem === el || tElem.contains(el)) return { elem: tElem, id: tId };
    }
    return null;
  }

  function onMove(e) {
    if (!active) return;
    if (active.pointerId !== undefined && e.pointerId !== active.pointerId) return;
    e.preventDefault();
    active.ghost.style.left = (e.clientX - active.offX) + 'px';
    active.ghost.style.top  = (e.clientY - active.offY) + 'px';

    const hit   = findTargetUnder(e.clientX, e.clientY);
    const hitId = hit ? hit.id : null;
    if (active.currentTarget !== hitId) {
      if (active.currentTarget) {
        for (const [elem, tid] of targets) {
          if (tid === active.currentTarget) elem.classList.remove('drag-over');
        }
      }
      if (hit) hit.elem.classList.add('drag-over');
      active.currentTarget = hitId;
    }
  }

  function onUp(e) {
    if (!active) return;
    if (active.pointerId !== undefined && e && e.pointerId !== undefined &&
        e.pointerId !== active.pointerId) return;

    document.removeEventListener('pointermove',   onMove);
    document.removeEventListener('pointerup',     onUp);
    document.removeEventListener('pointercancel', onUp);

    const dragId   = active.dragId;
    const targetId = active.currentTarget;
    const srcElem  = active.srcElem;
    const ghost    = active.ghost;

    for (const [elem] of targets) elem.classList.remove('drag-over');
    srcElem.classList.remove('dragging');
    ghost.remove();
    active = null;

    if (targetId && opts.onDrop) opts.onDrop(dragId, targetId);
    else if (opts.onCancel)      opts.onCancel(dragId);
  }

  return { attachSource, attachTarget };
}

/* ═══════════════════════════════════════════════════════════
   GLOBAL COMPONENT — Progress Question nav (generic renderer)
   720-templates → _global-components.md → "Progress Question".
   One shared state object + one renderer for the whole practice series,
   generic over item count — written once, reused by every question
   screen as the later ones are built.

   SERIES SHAPE for this component (רכיב 3, basic/standard practice):
   THREE progress items, from slide 39's own kickoff copy ("3 שאלות").
   A progress item is a SEMANTIC QUESTION, not a screen, and in this
   stage an item resolves only once ALL of its sub-parts are answered:
     item 1 → slides 40 + 41  (stems lettered "א." / "ב.")
     item 2 → slide 42        (single screen, no sub-parts)
     item 3 → slides 43 + 44  (stems lettered "א." / "ב.")
   Slide 40's presenter note "שאלה 1 מתוך 2" is therefore the SUB-PART
   index inside item 1, not the series length — slide 43's note says
   "שאלה 1 מתוך 2" as well, which only makes sense on that reading.

   `screens: []` = not-yet-built → never clickable, per the component's
   own click-to-navigate rule. Only slide 40 is built so far, so item 1
   currently lists just screen 1; append screen 2 when slide 41 lands.

   NEVER RESOLVE EARLY: item 1 stays 'current' after S1 is answered,
   because its sub-part ב (slide 41) is still unanswered. That is the
   stage's stated rule, not a bug — a wrong first sub-part must not show
   its ✕ until the rest of the question has been answered. `state` is
   therefore DERIVED from the sub-screens' own resolutions rather than
   stored, so it cannot drift.
   ═══════════════════════════════════════════════════════════ */
const practiceProgress = {
  questions: [
    { number: 1, visited: true,  state: 'current',      screens: [1, 2] },
    { number: 2, visited: false, state: 'not-answered', screens: [3] },
    { number: 3, visited: false, state: 'not-answered', screens: [4, 5] }
  ]
};

/* Sub-part resolution table for each progress item: screen index →
   a getter returning null (unanswered) | 'correct' | 'incorrect'.
   Every authored sub-part in this part is now built, so no getter
   returns `undefined` any more; the guard in derivePracticeState()
   stays, because it is what holds an item unresolved if one ever does. */
const PRACTICE_SUBPARTS = {
  1: [
    { screen: 1, resolved: () => (s1viqDone ? (s1viqResult === 'correct' ? 'correct' : 'incorrect') : null) },
    { screen: 2, resolved: () => (s2viqDone ? (s2viqResult === 'correct' ? 'correct' : 'incorrect') : null) }
  ],
  2: [
    { screen: 3, resolved: () => (s3viqDone ? (s3viqResult === 'correct' ? 'correct' : 'incorrect') : null) }
  ],
  3: [
    { screen: 4, resolved: () => (s4mcDone ? (s4mcResult === 'correct' ? 'correct' : 'incorrect') : null) },
    { screen: 5, resolved: () => {
        const st = window.lomdaState && window.lomdaState.s5;
        if (!st || !st.done) return null;
        return st.selection === S5DQ.correct ? 'correct' : 'incorrect';
      } }
  ]
};
/* Number of sub-parts each item is AUTHORED to have (from the deck),
   independent of how many are built yet — an item cannot resolve until
   every authored sub-part has actually been answered. */
const PRACTICE_SUBPART_COUNT = { 1: 2, 2: 1, 3: 2 };

function derivePracticeState(itemNumber, isCurrent) {
  const parts    = PRACTICE_SUBPARTS[itemNumber] || [];
  const authored = PRACTICE_SUBPART_COUNT[itemNumber] || 0;
  const results  = parts.map(p => p.resolved());
  const answered = results.filter(r => r !== null && r !== undefined);
  if (answered.length === 0) return isCurrent ? 'current' : 'not-answered';
  /* Not all authored sub-parts answered yet → stay unresolved. */
  if (answered.length < authored) return 'current';
  return answered.every(r => r === 'correct') ? 'correct' : 'incorrect';
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

/* Re-derives every item's state from the screens' own results, then
   renders. Called on entry into any question screen (not just once), so
   returning to an earlier question after resolving a later one shows
   the correct state on both. */
/* ── SECOND practice block (slides 46-49) ──
   Its own series, exactly as the approved sibling 03-02 carries two.
   Slide 45 announces "עוד 2 תרגילים", so TWO items: slide 46 on its
   own, and slides 47-49 as one three-sub-part question. Only item 1 is
   built yet; item 2's getters return undefined and correctly hold it
   unresolved. The derive rule is the first block's, unchanged. */
const practiceProgress2 = {
  questions: [
    { number: 1, visited: true,  state: 'current',      screens: [7] },
    { number: 2, visited: false, state: 'not-answered', screens: [8, 9, 10] }
  ]
};
const PRACTICE2_SUBPARTS = {
  1: [ { screen: 7, resolved: () => (s7viqDone ? (s7viqResult === 'correct' ? 'correct' : 'incorrect') : null) } ],
  2: [
    { screen: 8, resolved: () => (s8ddqDone ? (s8ddqResult === 'correct' ? 'correct' : 'incorrect') : null) },
    { screen: 9, resolved: () => (s9ddqDone ? (s9ddqResult === 'correct' ? 'correct' : 'incorrect') : null) },
    { screen: 10, resolved: () => (s10viqDone ? (s10viqResult === 'correct' ? 'correct' : 'incorrect') : null) }
  ]
};
const PRACTICE2_SUBPART_COUNT = { 1: 1, 2: 3 };

function derivePractice2State(itemNumber, isCurrent) {
  const parts    = PRACTICE2_SUBPARTS[itemNumber] || [];
  const authored = PRACTICE2_SUBPART_COUNT[itemNumber] || 0;
  const answered = parts.map(p => p.resolved()).filter(r => r !== null && r !== undefined);
  if (answered.length === 0) return isCurrent ? 'current' : 'not-answered';
  if (answered.length < authored) return 'current';
  return answered.every(r => r === 'correct') ? 'correct' : 'incorrect';
}

function syncPracticeProgress2Nav(sectionEl) {
  const currentItem = practiceProgress2.questions.find(q => q.screens.indexOf(currentScreen) !== -1);
  practiceProgress2.questions.forEach(q => {
    q.state   = derivePractice2State(q.number, q === currentItem);
    q.visited = q.state !== 'not-answered';
  });
  const nav = sectionEl && sectionEl.querySelector('.progress-question');
  if (nav) updateProgressQuestion(nav, practiceProgress2);
}

function syncPracticeProgressNav(sectionEl) {
  const currentItem = practiceProgress.questions.find(
    q => q.screens.indexOf(currentScreen) !== -1
  );
  practiceProgress.questions.forEach(q => {
    q.state   = derivePracticeState(q.number, q === currentItem);
    q.visited = q.state !== 'not-answered';
  });
  const nav = sectionEl && sectionEl.querySelector('.progress-question');
  if (nav) updateProgressQuestion(nav, practiceProgress);
}

/* ═══════════════════════════════════════════════════════════
   S1 — ValueInputQuestion, LOCAL RATIO VARIANT (slide 40)
   Practice question 1, sub-part א. Spec: 720-templates →
   ValueInputQuestion.md (Variant B's one-answer-object contract, laid
   out as a horizontal ratio rather than a stacked fraction) +
   _question-template-defaults.md + _global-components.md → "View My
   Answer Toggle". The whole state machine is ported from the same-lomda
   sibling methodica-ar-math-ratio-03-02's own S1, which is the same
   template and the same screen role; only geometry, copy and answer
   differ.

   CORRECT VALUES: boys:girls simplified = 1 : 2 — on three independent
   pieces of evidence: notesSlide36's explicit "الإجابة الصحيحة היא 1 : 2";
   the deck's own two placeholder-digit boxes ("1" and "2"); and the
   mathematics (boys = x, girls = 2x → x : 2x → 1 : 2), which both
   feedback panels spell out themselves.

   ⚠ RATIO ORDER — a genuine SOURCE CONFLICT, reported not silently
   normalized (pptx-script-extraction: contradictory evidence is
   reported, not resolved by preference).
   Slide 40 places the answer boxes at these resolved on-slide offsets
   (both at nesting depth 0, so the raw offsets ARE the real positions —
   verified, because a grouped shape's offset would have been in child
   space):
       box "1"  x =  9179312     box "2"  x = 10461460
   so "1" is the LEFT slot and "2" the RIGHT one. Confirmed against a
   fresh PowerPoint render of the slide, which shows literally
   "[1] : [2]". Read left-to-right that is "1 : 2" — exactly the note's
   answer string. Slide 42 uses the SAME two slots with the same hand
   ("5" left, "7" right; note "5 : 7").
   BUT the approved sibling 03-02's S1 does the OPPOSITE: its own
   slide 30 puts the first term in the RIGHT slot ("3" right, "2" left;
   note "3 : 2"). Those boxes sit inside a GROUP, which is why that
   part's note cites resolved coordinates (9274221 / 10534521) that
   differ from the raw child-space ones — the ORDER conclusion there is
   unaffected and correct.
   And slide 42 contradicts ITSELF: its own inline prose ratio
   "(x + 5) : (2x + 1)" renders with the FIRST term physically
   RIGHTMOST (ordinary bidi for an LTR group inside RTL text — verified
   in the render), the opposite hand from its own answer boxes.
   RESOLUTION SHIPPED HERE: follow slide 40's own rendered box order
   (first term LEFT), the same method 03-02 applied to its own slide, so
   each screen matches the slide it was built from and this screen reads
   identically to its note. The cross-part inconsistency is a DECK
   issue, not an implementation choice, and needs producer confirmation:
   if the producer wants one hand across the lomda, this is a two-line
   change here (swap the two fields' ids/labels) and the mirror-image
   change in 03-02.

   ATTEMPT MODEL: the 2-attempt default (maxAttempts: 2). The notes
   state no deviation, and the screen carries a hint plus both a correct
   and a single final-wrong feedback panel — the normal shape of a
   2-attempt question. The absence of a retry-popup or toggle-label
   mockup in the deck is expected and proves nothing
   (_question-template-defaults.md → "Two traps when reading the
   source").
   ═══════════════════════════════════════════════════════════ */
const S1VIQ = { correctFirst: 1, correctSecond: 2, maxAttempts: 2 };

let s1viqAttempts     = 0;
let s1viqDone         = false;
let s1viqResult       = null;   // 'correct' | 'wrong2'
let s1viqAnswer       = null;   // { first, second } raw typed strings, snapshotted before reveal
let s1viqView         = 'solution';
let s1viqLastWrongSig = null;

/* Shared explanation — the authored body of BOTH feedback panels
   (id=26 and id=28), which are byte-identical in the source and differ
   only in title (_question-template-defaults.md → "Feedback: shared
   explanation"). Preserved verbatim, including the source's own empty
   paragraph as an explicit group break.
   `expr: true` marks a line that is a bare algebraic/ratio expression:
   it is rendered as an isolated LTR island so bidi cannot reorder its
   terms, keeping "x : 2x" and "1 : 2" in the same hand as the answer
   boxes above. */
const S1VIQ_EXPLAIN = [
  { text: 'النسبة بين عدد الأولاد وعدد البنات هي' },
  { text: 'x : 2x', expr: true },
  { gap: true },
  { text: 'نختزل بـ x الذي لا يساوي 0 ونحصل على‎:' },   // client correction, 2026-09-29
  { text: '1 : 2', expr: true }
];
const S1VIQ_POPUP = {
  /* No retry-specific pedagogical explanation is authored for this
     question, so only the canonical fallback title is used —
     _question-template-defaults.md: "A missing pedagogical explanation
     must not be invented." */
  retry:   { bg: '#ffdbdc', title: 'الإجابة غير صحيحة.', body: [{ text: 'هل نحاول مرة أخرى؟' }] },
  /* Authored title, verbatim (id=26 P1). */
  correct: { bg: '#edf8ed', title: 'هذا صحيح‎!', body: S1VIQ_EXPLAIN },
  wrong2: {
    /* Authored title, verbatim — TWO bold lines in the source
       (id=28 P1 + P2), so a real line break between them. */
    bg: '#ffdbdc', title: 'هذا خطأ.<br>هيا نفهم لماذا:', body: S1VIQ_EXPLAIN,
    action: { label: 'الانتقال إلى إجابتي', view: 'answer' },
    answerView: {
      title: 'الإجابة غير صحيحة.',
      body:  [{ text: 'يمكن مشاهدة الحل الصحيح مرة أخرى‎.' }],
      action: { label: 'الإجابة الصحيحة', view: 'solution' }
    }
  }
};

function s1viqInput(k) { return document.getElementById('s1viq-input-' + k); }
function s1viqField(k) { return document.getElementById('s1viq-field-' + k); }
function s1viqNormalize(raw) { return String(raw).trim().replace(/,/g, ''); }
function s1viqSignature() {
  return s1viqNormalize(s1viqInput('first').value) + '|' + s1viqNormalize(s1viqInput('second').value);
}
function s1viqAllFilled() {
  return s1viqNormalize(s1viqInput('first').value) !== '' && s1viqNormalize(s1viqInput('second').value) !== '';
}
function s1viqIsCorrect() {
  return Number(s1viqNormalize(s1viqInput('first').value))  === S1VIQ.correctFirst &&
         Number(s1viqNormalize(s1viqInput('second').value)) === S1VIQ.correctSecond;
}

/* type="text" (required so a thousands comma can be typed at all)
   removes the browser's own numeric enforcement, so every keystroke AND
   paste is filtered here — a paste fires `input` too, so no separate
   paste handler is needed (ValueInputQuestion.md → "Numeric input
   format"). */
function s1viqOnInput() {
  if (s1viqDone) return;
  ['first', 'second'].forEach(k => {
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
     defaults.md → "Retry gate"): resubmitting with one field changed
     and the other identical is still "the same wrong answer" only if
     the JOINED value matches, and a comma-only edit normalizes away so
     it does not count as a change. */
  chk.disabled = !s1viqAllFilled() || s1viqSignature() === s1viqLastWrongSig;
}

/* Case B (one semantic answer object): BOTH fields always share the
   SAME single outcome, and there is exactly one verdict icon for the
   whole group — never a per-field icon. */
function s1viqMark(outcome) {
  const ico = document.getElementById('s1viq-ico');
  ['first', 'second'].forEach(k => {
    const f = s1viqField(k);
    f.classList.remove('viq-correct', 'viq-wrong');
    if (outcome) f.classList.add(outcome === 'correct' ? 'viq-correct' : 'viq-wrong');
  });
  if (ico) ico.innerHTML = !outcome ? '' : (outcome === 'correct' ? S1VIQ_ICON_CORRECT : S1VIQ_ICON_WRONG);
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
  reportQ(['001', 'q1'], correct, correct || s1viqAttempts >= S1VIQ.maxAttempts, String((document.getElementById('s1viq-input-first') || {}).value).trim() + ' : ' + String((document.getElementById('s1viq-input-second') || {}).value).trim());

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
  s1viqAnswer = { first: s1viqInput('first').value, second: s1viqInput('second').value };
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
    s1viqInput('first').value  = s1viqAnswer.first;
    s1viqInput('second').value = s1viqAnswer.second;
    s1viqMark('wrong');   // the final attempt was wrong in both views by definition
  } else {
    s1viqInput('first').value  = String(S1VIQ.correctFirst);
    s1viqInput('second').value = String(S1VIQ.correctSecond);
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
  s1viqInput('first').disabled  = true;
  s1viqInput('second').disabled = true;
  const chk = document.getElementById('s1viq-check');
  /* ValueInputQuestion's own completion label is "متابعة" — NOT
     SingleChoiceQuestion's "هل نتابع؟"; the template doc is explicit that
     the two templates do not share one. Slide 40 authors no progression
     copy of its own, so this canonical fallback applies. */
  if (chk) { setNavLabel(chk, 'متابعة'); chk.disabled = false; }
  /* Hint stays reachable after resolution? No — the template hides it
     on finish (ValueInputQuestion.md → "{prefix}Finish(result) … hides
     hint"). */
  const hint = document.getElementById('s1viq-hint');
  if (hint) hint.classList.add('hidden');
  syncPracticeProgressNav(document.getElementById('s1'));
}

function s1viqResetInitial() {
  s1viqAttempts     = 0;
  s1viqDone         = false;
  s1viqResult       = null;
  s1viqAnswer       = null;
  s1viqView         = 'solution';
  s1viqLastWrongSig = null;
  ['first', 'second'].forEach(k => {
    const el = s1viqInput(k);
    el.value = '';
    el.disabled = false;
  });
  s1viqMark(null);
  document.getElementById('s1viq-popup').classList.add('hidden');
  document.getElementById('s1viq-hint-overlay').classList.add('hidden');
  const chk = document.getElementById('s1viq-check');
  if (chk) { setNavLabel(chk, 'هل إجابتي صحيحة؟'); chk.disabled = true; }
  /* Hint visible and enabled from the start — the canonical template
     default; slide 40 states no availability policy and mocks the
     trigger in its enabled state in the base on-canvas layer. */
  const hint = document.getElementById('s1viq-hint');
  if (hint) { hint.classList.remove('hidden'); hint.disabled = false; }
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
     a fixed second line") — APPENDED to, never replacing, the authored
     title. Does not apply to the answerView (which shows the learner's
     own answer, not the solution) or to retry/correct. */
  titleEl.innerHTML = (type === 'wrong2' && view !== 'answer')
    ? cfg.title + '<br>الإجابة الصحيحة تَظهَر الآن.'
    : cfg.title;
  document.getElementById('s1viq-popup-body').innerHTML = (cfg.body || []).map(p => {
    if (p.gap)  return '<p class="scq-popup-gap" aria-hidden="true"></p>';
    if (p.expr) return '<p class="scq-popup-expr" dir="ltr">' + p.text + '</p>';
    return '<p>' + p.text + '</p>';
  }).join('');
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
   S2 — ValueInputQuestion, VARIANT A (single value) — slide 41
   Practice question 1, sub-part ב — the second and LAST sub-part of the
   question S1 opened, so answering this screen is what finally resolves
   progress item 1.

   Spec: 720-templates → ValueInputQuestion.md (Variant A: one inline
   input box, ONE verdict icon INSIDE the field's own border) +
   _question-template-defaults.md + _global-components.md → "View My
   Answer Toggle". State-machine shape ported from this part's own S1 and
   from the approved sibling methodica-ar-math-ratio-03-02's S4 (the same
   inline single-field shape); only the answer, copy and geometry differ.

   CORRECT VALUE: 36 — on three independent pieces of evidence:
   notesSlide37's explicit "الإجابة الصحيحة היא 36"; the deck's own
   placeholder digit in the single value box (id=7, "36"); and the
   mathematics, which this slide's own feedback panels spell out in full
   (girls = 2x = 24 → x = 12 boys → 12 + 24 = 36 pupils).

   No ratio order question arises on this screen: there is ONE box, and
   the authored unit word "תלמידים." sits to its left because the source
   paragraph puts its spacer FIRST (confirmed against a fresh render).

   ATTEMPT MODEL: the 2-attempt default. The notes state no deviation,
   and the screen carries a hint plus one correct and one final-wrong
   panel — the normal shape of a 2-attempt question.

   ACCEPTED FORMAT: `inputmode="decimal"` per Variant A, and the shared
   character filter allows digits, comma, dot and minus; the comparison
   normalizes thousands commas away, so "36" and "36.0" both pass while
   the retry gate does not treat a comma-only edit as a new answer.
   ═══════════════════════════════════════════════════════════ */
const S2VIQ = { correct: 36, maxAttempts: 2 };

let s2viqAttempts     = 0;
let s2viqDone         = false;
let s2viqResult       = null;   // 'correct' | 'wrong2'
let s2viqAnswer       = null;   // the RAW typed string, snapshotted before the reveal
let s2viqView         = 'solution';
let s2viqLastWrongSig = null;

/* Shared explanation — the authored body of BOTH feedback panels
   (id=26 and id=28), byte-identical in the source and differing only in
   title (_question-template-defaults.md → "Feedback: shared
   explanation"). Preserved verbatim, including all four of the source's
   own EMPTY paragraphs as real group breaks.
     `expr: true` — the line is a bare LTR calculation: rendered as an
     isolated LTR island so bidi cannot reorder its terms, and nowrap so
     automatic wrapping cannot split it.
     Inline `<span dir="ltr">` — a single Latin/digit token embedded in
     Hebrew prose needs one isolated element PER TOKEN. A bare number
     like "24" or "12" needs none; only "2x" does. */
const S2VIQ_EXPLAIN = [
  { text: 'النسبة بين عدد الأولاد وعدد البنات هي' },
  { text: 'x : 2x', expr: true },
  { gap: true },
  { text: 'معلوم أن عدد البنات هو <span dir="ltr">2x</span>، وهو يساوي 24. لذلك‎:' },
  /* ⚠ PRODUCER OVERRIDE — DO NOT "correct" THIS BACK TO THE PPT FORM.
     The PPTX stores and renders this run as "2x = 24 / :2" — verified
     twice: the source run hex-dumps to
     2,x,SP,=,SP,2,4,SP,/,SP,:,2 with NO U+0020 between ":" and "2",
     and an ink-column analysis of the slide rendered from PowerPoint
     measures only 2.67px there (pure glyph side bearing; a real space
     on that line measures 6.9-8.0px).
     The producer has nevertheless approved adding ONE ordinary space
     between the colon and the final 2, so the operator reads
     "/ : 2". This is an intentional, documented deviation from source
     fidelity, instructed 2026-09-09 — literal PPT fidelity does not
     override it. The space below is a real U+0020 text character
     inside the isolated LTR run, NOT a margin, pseudo-element or
     bearing trick, so it is a genuine space in the text itself.
     One authored entry serves BOTH feedback panels (correct and
     wrong2 share S2VIQ_EXPLAIN), so this is the only place it lives. */
  { text: '2x = 24 / : 2', expr: true },
  { text: 'x = 12', expr: true },
  { gap: true },
  { text: 'أي إن عدد الأولاد هو 12‏‎.' },
  { text: 'نجمع عدد الأولاد وعدد البنات‎:' },
  { text: '12 + 24 = 36', expr: true },
  { gap: true },
  { text: 'الاستنتاج: عدد الطلاب في الصف هو 36‏‎.' }
];
const S2VIQ_POPUP = {
  /* No retry-specific pedagogical explanation is authored, so only the
     canonical fallback title is used — a missing pedagogical explanation
     must never be invented. */
  retry:   { bg: '#ffdbdc', title: 'الإجابة غير صحيحة.', body: [{ text: 'هل نحاول مرة أخرى؟' }] },
  /* Authored titles, verbatim (id=26 P1 / id=28 P1). */
  correct: { bg: '#edf8ed', title: 'رائع جدًا!', body: S2VIQ_EXPLAIN },
  wrong2: {
    bg: '#ffdbdc', title: 'هذا خطأ. لا بأس، دعونا نشرح‎:', body: S2VIQ_EXPLAIN,
    action: { label: 'الانتقال إلى إجابتي', view: 'answer' },
    answerView: {
      title: 'الإجابة غير صحيحة.',
      body:  [{ text: 'يمكن مشاهدة الحل الصحيح مرة أخرى‎.' }],
      action: { label: 'الإجابة الصحيحة', view: 'solution' }
    }
  }
};

function s2viqInput() { return document.getElementById('s2viq-input'); }
function s2viqNormalize(raw) { return String(raw).trim().replace(/,/g, ''); }
function s2viqSignature() { return s2viqNormalize(s2viqInput().value); }
function s2viqFilled() { return s2viqSignature() !== ''; }
function s2viqIsCorrect() { return Number(s2viqSignature()) === S2VIQ.correct; }

/* type="text" (required so a thousands comma can be typed at all)
   removes the browser's own numeric enforcement, so every keystroke AND
   paste is filtered here — a paste fires `input` too, so no separate
   paste handler is needed (ValueInputQuestion.md → "Numeric input
   format"). */
function s2viqOnInput() {
  if (s2viqDone) return;
  const el = s2viqInput();
  const cleaned = el.value.replace(/[^0-9,.-]/g, '');
  if (cleaned !== el.value) el.value = cleaned;
  /* Starting a new attempt dismisses the previous retry popup —
     approved path 2 of "No manual close". Unreachable once answered. */
  if (s2viqAttempts > 0) {
    document.getElementById('s2viq-popup').classList.add('hidden');
    s2viqMark(null);
  }
  s2viqUpdateBar();
}

function s2viqUpdateBar() {
  const chk = document.getElementById('s2viq-check');
  if (!chk || s2viqDone) return;
  /* Retry gate (_question-template-defaults.md → "Retry gate"): a live
     comparison against the last-wrong value, not a one-shot flag — so
     reverting to the same wrong answer disables the button again, and a
     comma-only edit normalizes away and does not count as a change. */
  chk.disabled = !s2viqFilled() || s2viqSignature() === s2viqLastWrongSig;
}

/* ONE field = ONE answer, so exactly one verdict icon, and it lives
   INSIDE the field's bordered box (ValueInputQuestion.md → Variant A).
   The field surface itself stays Natural/100 in every state — only the
   border and the icon carry outcome colour ("Field surface invariant"). */
function s2viqMark(outcome) {
  const f   = document.getElementById('s2viq-field');
  const ico = document.getElementById('s2viq-ico');
  f.classList.remove('viq-correct', 'viq-wrong');
  if (outcome) f.classList.add(outcome === 'correct' ? 'viq-correct' : 'viq-wrong');
  if (ico) ico.innerHTML = !outcome ? '' : (outcome === 'correct' ? S1VIQ_ICON_CORRECT : S1VIQ_ICON_WRONG);
}

function s2viqCheck() {
  if (s2viqDone) { advanceScreen(); return; }
  if (!s2viqFilled()) return;

  s2viqAttempts++;
  reportQ(['001', 'q2'], s2viqIsCorrect(), s2viqIsCorrect() || s2viqAttempts >= S2VIQ.maxAttempts, String((document.getElementById('s2viq-input') || {}).value).trim());
  if (s2viqIsCorrect()) {
    s2viqMark('correct');
    s2viqResult = 'correct';
    s2viqFinish();
    s2viqShowPopup('correct');
    return;
  }

  if (s2viqAttempts < S2VIQ.maxAttempts) {
    s2viqMark('wrong');
    s2viqLastWrongSig = s2viqSignature();
    document.getElementById('s2viq-check').disabled = true;
    s2viqShowPopup('retry');
    return;
  }

  /* Final wrong. The reveal overwrites the very input the learner typed
     into, so snapshot the RAW string first — this is the only place it
     survives. Preserved exactly as typed (a leading zero included),
     never reformatted or re-derived from the parsed number. */
  s2viqAnswer = s2viqInput().value;
  s2viqResult = 'wrong2';
  s2viqView   = 'solution';
  s2viqFinish();
  s2viqApplyView();
}

/* Single owner of the resolved-wrong rendering. */
function s2viqApplyView() {
  if (s2viqAnswer === null) return;
  if (s2viqView === 'answer') {
    s2viqInput().value = s2viqAnswer;
    s2viqMark('wrong');   // the final attempt was wrong in both views by definition
  } else {
    s2viqInput().value = String(S2VIQ.correct);
    s2viqMark('correct');
  }
  s2viqShowPopup('wrong2', s2viqView);
}
/* Pure view switch — NEVER routes through s2viqCheck(), never touches
   attempts/done/scoring. */
function s2viqToggleView() {
  if (s2viqAnswer === null) return;
  s2viqView = (s2viqView === 'solution') ? 'answer' : 'solution';
  s2viqApplyView();
}

function s2viqFinish() {
  s2viqDone = true;
  s2viqInput().disabled = true;
  const chk = document.getElementById('s2viq-check');
  /* ValueInputQuestion's own completion label is "متابعة"; slide 41
     authors no progression copy of its own, so the canonical fallback
     applies. */
  if (chk) { setNavLabel(chk, 'متابعة'); chk.disabled = false; }
  const hint = document.getElementById('s2viq-hint');
  if (hint) hint.classList.add('hidden');
  /* This is question 1's LAST sub-part, so resolving it is what finally
     resolves progress item 1 — and the item goes ✓ only if BOTH
     sub-parts were correct. syncPracticeProgressNav() re-derives that
     from the two screens' own results; nothing is stored here. */
  syncPracticeProgressNav(document.getElementById('s2'));
}

function s2viqResetInitial() {
  s2viqAttempts     = 0;
  s2viqDone         = false;
  s2viqResult       = null;
  s2viqAnswer       = null;
  s2viqView         = 'solution';
  s2viqLastWrongSig = null;
  const el = s2viqInput();
  el.value = '';
  el.disabled = false;
  s2viqMark(null);
  document.getElementById('s2viq-popup').classList.add('hidden');
  document.getElementById('s2viq-hint-overlay').classList.add('hidden');
  const chk = document.getElementById('s2viq-check');
  if (chk) { setNavLabel(chk, 'هل إجابتي صحيحة؟'); chk.disabled = true; }
  /* Hint visible and enabled from the start — canonical template
     default; slide 41 states no availability policy and mocks the
     trigger enabled in its base on-canvas layer. */
  const hint = document.getElementById('s2viq-hint');
  if (hint) { hint.classList.remove('hidden'); hint.disabled = false; }
}

function s2viqShowPopup(type, view) {
  const popup = document.getElementById('s2viq-popup');
  if (!popup) return;
  const base = S2VIQ_POPUP[type];
  const cfg  = (view === 'answer' && base.answerView)
             ? Object.assign({}, base, base.answerView)
             : base;
  popup.style.background = cfg.bg;
  if (popup.classList.contains('hidden') || type === 'wrong2') resetPopupPosition(popup);
  const titleEl = document.getElementById('s2viq-popup-title');
  /* Terminal-incorrect SOLUTION view carries a fixed second line,
     APPENDED to the authored title, never replacing it. Not applied to
     the answerView (which shows the learner's own answer) or to
     retry/correct. */
  titleEl.innerHTML = (type === 'wrong2' && view !== 'answer')
    ? cfg.title + '<br>الإجابة الصحيحة تَظهَر الآن.'
    : cfg.title;
  document.getElementById('s2viq-popup-body').innerHTML = (cfg.body || []).map(p => {
    if (p.gap)  return '<p class="scq-popup-gap" aria-hidden="true"></p>';
    /* An expression line is emitted as ONE isolated LTR run: the <p>
       carries dir="ltr" + unicode-bidi: isolate from CSS, and the inner
       <bdi dir="ltr"> isolates it again at the ELEMENT level so the
       complete operation survives even if the CSS property were ever
       dropped or overridden. `2x = 24 / : 2` is a single mathematical
       statement — equation plus the "divide both sides by 2" operator —
       and must never be split across runs, since bidi reorders a
       trailing Common Separator (":") against an adjacent European
       Number ("2") and would render the operator as "2 :". <bdi> is the
       same idiom the approved sibling 03-02 uses for its inline ratios.
       The isolation is also what keeps the producer-approved space
       between ":" and "2" (see the PRODUCER OVERRIDE note on
       S2VIQ_EXPLAIN) in place: unisolated, that space sits on a
       direction boundary and bidi is free to move it. */
    if (p.expr) return '<p class="scq-popup-expr viq-expr" dir="ltr"><bdi dir="ltr">' + p.text + '</bdi></p>';
    return '<p>' + p.text + '</p>';
  }).join('');
  renderPopupAction('s2viq-popup', cfg.action, s2viqToggleView);
  popup.classList.remove('hidden');
  clampPopup(popup);
}

function s2viqOpenHint() {
  xapiRequestedHint('001', 'q2');
  document.getElementById('s2viq-hint-overlay').classList.remove('hidden');
}
function s2viqCloseHint() {
  document.getElementById('s2viq-hint-overlay').classList.add('hidden');
}
function s2viqCloseHintOnBackdrop(e) {
  if (e.target === e.currentTarget) s2viqCloseHint();
}

/* ═══════════════════════════════════════════════════════════
   S3 — ValueInputQuestion, LOCAL RATIO VARIANT (slide 42)
   Practice QUESTION 2 of the series — a single-screen question with no
   lettered sub-parts, so it is progress item 2 and resolves as soon as
   this screen resolves.

   Spec: ValueInputQuestion.md (Variant B's one-answer-object contract,
   laid out as a horizontal ratio) + _question-template-defaults.md +
   _global-components.md → "View My Answer Toggle".
   The state machine is this part's own S1 shape, reused rather than
   reinvented; only the answer, copy, geometry and the image differ.

   CORRECT VALUES: 5 : 7 — on three independent pieces of evidence:
   notesSlide38's explicit "الإجابة الصحيحة היא 5 : 7"; the deck's own two
   placeholder digits ("5" in the LEFT slot at x=963.7, "7" in the RIGHT
   at x=1098.3 — the SAME slots and the same hand slide 40 uses, so no
   new ratio-order question arises); and the mathematics, which both
   feedback panels spell out in full (x = 10 → Dana (10+5) = 15, Yael
   (2·10+1) = 21 → 15 : 21, ÷3 → 5 : 7).

   ATTEMPT MODEL: the 2-attempt default — the notes state no deviation,
   and the screen carries a hint plus one correct and one final-wrong
   panel, the normal shape of a 2-attempt question.
   ═══════════════════════════════════════════════════════════ */
const S3VIQ = { correctFirst: 5, correctSecond: 7, maxAttempts: 2 };

let s3viqAttempts     = 0;
let s3viqDone         = false;
let s3viqResult       = null;   // 'correct' | 'wrong2'
let s3viqAnswer       = null;   // { first, second } raw typed strings, snapshotted before the reveal
let s3viqView         = 'solution';
let s3viqLastWrongSig = null;

/* Shared explanation — the authored body of BOTH panels (id=26 and
   id=28), preserved verbatim including all three of the source's own
   EMPTY paragraphs as real group breaks. The two panels are
   structurally identical (13 paragraphs each) and differ only in title,
   so this is the shared-explanation default, not an assumption.
     One line — "(10 + 5) : (2 · 10 + 1)" — is authored as an ORDINARY
   RUN in the correct panel but as an <m:oMath> EQUATION in the
   final-wrong panel. Both carry the same characters; the text below is
   the OMML's own m:t content, read from the equation tree rather than
   from flat text (a flat-text read of id=28 drops the line entirely).
     `expr: true` marks a bare LTR expression line: rendered as an
   isolated LTR island so bidi cannot reorder its terms, and nowrap so
   automatic wrapping cannot split it. This matters more here than
   anywhere else in the part — measured in real Edge, the unisolated
   string "(x + 5) : (2x + 1)" inside RTL prose is shredded by bidi into
   "1) (2x + : 5) + (x", whereas PowerPoint renders it intact. */
const S3VIQ_EXPLAIN = [
  { text: 'النسبة بين عدد كتب دنيا وعدد كتب أمل هي‎:' },
  { text: '(x + 5) : (2x + 1)', expr: true },
  { gap: true },
  { text: 'نعوّض <span dir="ltr">x = 10</span> في النسبة المعطاة، ونحصل على‎:' },
  { text: '(10 + 5) : (2 · 10 + 1)', expr: true },
  { gap: true },
  { text: 'إذًا، النسبة هي‎:' },
  { text: '15 : 21', expr: true },
  { gap: true },
  { text: 'نختزل بـ 3 ونحصل على النسبة المختزلة‎:' },
  { text: '5 : 7', expr: true }
];
const S3VIQ_POPUP = {
  /* No retry-specific pedagogical explanation is authored — only the
     canonical fallback title; a missing explanation is never invented. */
  retry:   { bg: '#ffdbdc', title: 'الإجابة غير صحيحة.', body: [{ text: 'هل نحاول مرة أخرى؟' }] },
  /* Authored titles, verbatim (id=26 P1 / id=28 P1). */
  correct: { bg: '#edf8ed', title: 'هذا صحيح‎!', body: S3VIQ_EXPLAIN },
  wrong2: {
    bg: '#ffdbdc', title: 'هذا خطأ. لا بأس، فهكذا نتعلم‎.', body: S3VIQ_EXPLAIN,
    action: { label: 'الانتقال إلى إجابتي', view: 'answer' },
    answerView: {
      title: 'الإجابة غير صحيحة.',
      body:  [{ text: 'يمكن مشاهدة الحل الصحيح مرة أخرى‎.' }],
      action: { label: 'الإجابة الصحيحة', view: 'solution' }
    }
  }
};

function s3viqInput(k) { return document.getElementById('s3viq-input-' + k); }
function s3viqField(k) { return document.getElementById('s3viq-field-' + k); }
function s3viqNormalize(raw) { return String(raw).trim().replace(/,/g, ''); }
function s3viqSignature() {
  return s3viqNormalize(s3viqInput('first').value) + '|' + s3viqNormalize(s3viqInput('second').value);
}
function s3viqAllFilled() {
  return s3viqNormalize(s3viqInput('first').value) !== '' && s3viqNormalize(s3viqInput('second').value) !== '';
}
function s3viqIsCorrect() {
  return Number(s3viqNormalize(s3viqInput('first').value))  === S3VIQ.correctFirst &&
         Number(s3viqNormalize(s3viqInput('second').value)) === S3VIQ.correctSecond;
}

function s3viqOnInput() {
  if (s3viqDone) return;
  ['first', 'second'].forEach(k => {
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
  /* Retry gate on the FULL joined signature, compared live against the
     last wrong answer, with thousands commas normalized away. */
  chk.disabled = !s3viqAllFilled() || s3viqSignature() === s3viqLastWrongSig;
}

/* Case B: both fields always share the SAME single outcome, and there is
   exactly one verdict icon for the whole ratio. */
function s3viqMark(outcome) {
  const ico = document.getElementById('s3viq-ico');
  ['first', 'second'].forEach(k => {
    const f = s3viqField(k);
    f.classList.remove('viq-correct', 'viq-wrong');
    if (outcome) f.classList.add(outcome === 'correct' ? 'viq-correct' : 'viq-wrong');
  });
  if (ico) ico.innerHTML = !outcome ? '' : (outcome === 'correct' ? S1VIQ_ICON_CORRECT : S1VIQ_ICON_WRONG);
}

function s3viqCheck() {
  if (s3viqDone) { advanceScreen(); return; }
  if (!s3viqAllFilled()) return;

  s3viqAttempts++;
  reportQ(['002', 'q1'], s3viqIsCorrect(), s3viqIsCorrect() || s3viqAttempts >= S3VIQ.maxAttempts, String((document.getElementById('s3viq-input-first') || {}).value).trim() + ' : ' + String((document.getElementById('s3viq-input-second') || {}).value).trim());
  if (s3viqIsCorrect()) {
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

  /* Final wrong — snapshot the RAW typed strings before the reveal
     overwrites the very inputs they came from. */
  s3viqAnswer = { first: s3viqInput('first').value, second: s3viqInput('second').value };
  s3viqResult = 'wrong2';
  s3viqView   = 'solution';
  s3viqFinish();
  s3viqApplyView();
}

/* Object-level toggle: one button rewrites BOTH fields and the single
   combined verdict in the same pass. */
function s3viqApplyView() {
  if (s3viqAnswer === null) return;
  if (s3viqView === 'answer') {
    s3viqInput('first').value  = s3viqAnswer.first;
    s3viqInput('second').value = s3viqAnswer.second;
    s3viqMark('wrong');
  } else {
    s3viqInput('first').value  = String(S3VIQ.correctFirst);
    s3viqInput('second').value = String(S3VIQ.correctSecond);
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
  s3viqInput('first').disabled  = true;
  s3viqInput('second').disabled = true;
  const chk = document.getElementById('s3viq-check');
  if (chk) { setNavLabel(chk, 'متابعة'); chk.disabled = false; }
  const hint = document.getElementById('s3viq-hint');
  if (hint) hint.classList.add('hidden');
  /* This question has ONE sub-part, so resolving this screen resolves
     progress item 2 outright — syncPracticeProgressNav() re-derives it
     from this screen's own result; nothing is stored here. */
  syncPracticeProgressNav(document.getElementById('s3'));
}

function s3viqResetInitial() {
  s3viqAttempts     = 0;
  s3viqDone         = false;
  s3viqResult       = null;
  s3viqAnswer       = null;
  s3viqView         = 'solution';
  s3viqLastWrongSig = null;
  ['first', 'second'].forEach(k => {
    const el = s3viqInput(k);
    el.value = '';
    el.disabled = false;
  });
  s3viqMark(null);
  document.getElementById('s3viq-popup').classList.add('hidden');
  document.getElementById('s3viq-hint-overlay').classList.add('hidden');
  const chk = document.getElementById('s3viq-check');
  if (chk) { setNavLabel(chk, 'هل إجابتي صحيحة؟'); chk.disabled = true; }
  const hint = document.getElementById('s3viq-hint');
  if (hint) { hint.classList.remove('hidden'); hint.disabled = false; }
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
  document.getElementById('s3viq-popup-body').innerHTML = (cfg.body || []).map(p => {
    if (p.gap)  return '<p class="scq-popup-gap" aria-hidden="true"></p>';
    if (p.expr) return '<p class="scq-popup-expr viq-expr" dir="ltr"><bdi dir="ltr">' + p.text + '</bdi></p>';
    return '<p>' + p.text + '</p>';
  }).join('');
  renderPopupAction('s3viq-popup', cfg.action, s3viqToggleView);
  popup.classList.remove('hidden');
  clampPopup(popup);
}

function s3viqOpenHint() {
  xapiRequestedHint('002', 'q1');
  document.getElementById('s3viq-hint-overlay').classList.remove('hidden');
}
function s3viqCloseHint() {
  document.getElementById('s3viq-hint-overlay').classList.add('hidden');
}
function s3viqCloseHintOnBackdrop(e) {
  if (e.target === e.currentTarget) s3viqCloseHint();
}

   // px


/* ═══════════════════════════════════════════════════════════
   S4 — MultipleChoiceQuestion (slide 43)
   Practice question 3, sub-part א — the FIRST of two sub-parts, so
   resolving this screen does NOT resolve progress item 3.

   Spec: 720-templates → MultipleChoiceQuestion.md +
   _question-template-defaults.md + _global-components.md → Feedback
   popup system → "Long-feedback variant". The question workflow,
   attempts, hint policy, popup and bottom bar are SingleChoiceQuestion's
   — reused from the approved sibling 03-02's own S2 — with this
   template's single behavioural delta: selection is a Set.

   CORRECT SET = {o1, o3} — "a : 3a" and the fraction a/3a. Four
   independent pieces of evidence, recorded in full on this screen's
   markup in index.html: the option rects' own green/red fills, the
   three marker squares, the RESOLVED producer comment naming "a : 3a"
   as the correct answer, and the mathematics the feedback panels spell
   out. o2 ("a : 3") is the distractor — the plain count 3 instead of
   the 3a the rule yields.

   DOES NOT use the View My Answer Toggle: this template shows the full
   correct set alongside the learner's own selections simultaneously, so
   the "one slot, two possible values" applicability rule excludes it.

   ATTEMPT MODEL: the 2-attempt default; the notes state no deviation.
   ═══════════════════════════════════════════════════════════ */
const S4MC = { correct: new Set(['o1', 'o3']), maxAttempts: 2 };

let s4mcSelected  = new Set();
let s4mcAttempts  = 0;
let s4mcDone      = false;
let s4mcResult    = null;   // 'correct' | 'wrong2'
let s4mcLastWrong = null;   // serialized Set of the immediately-preceding wrong attempt

/* Stable serialization for the retry gate — sorted so membership, not
   click order, is what is compared. */
function s4mcSig(set) { return [...set].sort().join('|'); }

/* Shared explanation — the authored body of BOTH panels (id=26 correct,
   id=16 final-wrong), which are structurally identical (14 paragraphs,
   4 empties) and differ only in title, so this is the shared-explanation
   default rather than an assumption. Preserved verbatim, including the
   source's own empty paragraphs as real breaks and the ONE bold run
   inside the body (P9, the ratio's labels).
     The source's leading empty paragraph (P2, between title and body)
   is not repeated here: the popup's own canonical 4px title→body gap is
   what stands in for it. (It used to be realized by the scroll column's
   24px gap; that column is gone — see S4MC_POPUP.)
     `math: true` marks the final line, which the source authors as an
   <m:oMath> stacked fraction — rendered with the .frac component, not
   as the flat text "a3a" a naive read produces.
     `expr: true` marks the bare ratio line so bidi cannot reorder its
   terms. Inline single tokens ("a", "3a") are isolated per token. */
const S4MC_EXPLAIN = [
  { text: 'معطى أنه مقابل كل <bdi dir="ltr">a</bdi> طوابق تحت الأرض توجد 3 طوابق فوق الأرض‎.' },
  { gap: true },
  { text: '‎عدد الطوابق تحت الأرض‎ ممثَّل بـ <bdi dir="ltr">a</bdi>.' },
  { text: 'ولذلك فإن عدد الطوابق فوق الأرض هو <bdi dir="ltr">3a</bdi>.' },
  { gap: true },
  { text: 'ومن هنا فإن النسبة' },
  { text: '<strong>تحت الأرض : فوق الأرض</strong>' },
  { text: 'هي‎:' },
  { text: 'a : 3a', expr: true },
  { gap: true },
  { text: 'ويمكن تمثيل النسبة نفسها أيضًا بكسر‎:' },
  { html: '<span class="frac"><span class="frac-num">a</span><span class="frac-den">3a</span></span>', math: true }
];
/* NO `scrollGate` on any outcome — PRODUCER DECISION 2026-09-10, a LOCAL
   EXCEPTION FOR THIS SCREEN ONLY. Slide 43's note does ask for the
   long-feedback variant in as many words ("שימו לב שבמשוב ארוך, לאפשר
   גלילה"), and this screen shipped with it; measured afterwards, the
   gated 474px box overflowed by just 4px, and that 4px came from the
   gated variant's own 24px title→body gap rather than from the content.
   At the canonical .scq-popup gap of 4px the explanation fits entirely,
   so the gate asked the learner to scroll a panel with nothing below
   the fold. The mechanism is intact and still available (the generic
   .is-scroll CSS, atBottom(), placeScrollHint(), the drag guard): a
   later screen whose feedback genuinely overflows sets `scrollGate` on
   its own config. Do NOT re-add it here. */
const S4MC_POPUP = {
  retry:   { bg: '#ffdbdc', title: 'الإجابة غير صحيحة.', body: [{ text: 'هل نحاول مرة أخرى؟' }] },
  correct: { bg: '#edf8ed', title: 'إجابتكم صحيحة، كل الاحترام!', body: S4MC_EXPLAIN },
  wrong2:  { bg: '#ffdbdc', title: 'هذا خطأ. وهذه فرصة لتوضيح الفكرة‎.', body: S4MC_EXPLAIN }
};

function s4mcOpts()  { return [...document.querySelectorAll('#s4 .scq-opt')]; }

/* Toggle — adds/removes ONLY this id. Never clears the other
   selections; this is the one behaviour that must not regress into
   SingleChoiceQuestion's exclusive select. */
function s4mcToggle(id) {
  if (s4mcDone) return;
  if (s4mcSelected.has(id)) s4mcSelected.delete(id);
  else                      s4mcSelected.add(id);
  /* A genuine change clears the previous attempt's per-option marks and
     dismisses its retry popup — approved path 2 of "No manual close". */
  if (s4mcAttempts > 0) {
    document.getElementById('s4mc-popup').classList.add('hidden');
    s4mcOpts().forEach(o => o.classList.remove('correct', 'wrong'));
  }
  s4mcRender();
}

function s4mcRender() {
  s4mcOpts().forEach(o => {
    const on = s4mcSelected.has(o.dataset.id);
    o.classList.toggle('selected', on);
    o.setAttribute('aria-checked', on ? 'true' : 'false');
  });
  const chk = document.getElementById('s4mc-check');
  if (!chk || s4mcDone) return;
  /* Gate on "at least one selected" — never on matching the number of
     correct answers; the learner may legitimately try a wrong count.
     Retry gate on top: the full sorted Set signature, so an off/on
     round-trip back to the same Set does NOT re-enable the button. */
  chk.disabled = s4mcSelected.size < 1 || s4mcSig(s4mcSelected) === s4mcLastWrong;
}

function s4mcCheck() {
  if (s4mcDone) { advanceScreen(); return; }
  if (s4mcSelected.size < 1) return;

  s4mcAttempts++;
  const correct = setsEqual(s4mcSelected, S4MC.correct);
  reportQ(['003', 'q1'], correct, correct || s4mcAttempts >= S4MC.maxAttempts, xapiChoiceAnswer('003', 'q1', '#s4 .scq-opt', [...s4mcSelected]));

  if (correct) {
    /* Fully correct: mark every option in the correct set green. */
    s4mcOpts().forEach(o => {
      o.classList.remove('selected', 'wrong');
      if (S4MC.correct.has(o.dataset.id)) o.classList.add('correct');
    });
    s4mcResult = 'correct';
    s4mcFinish();
    s4mcShowPopup('correct');
    return;
  }

  if (s4mcAttempts < S4MC.maxAttempts) {
    /* Non-final wrong — per-SELECTED-option feedback only. Every option
       the learner actually submitted gets a verdict; a correct option
       they did NOT select stays neutral and unrevealed. */
    s4mcOpts().forEach(o => {
      o.classList.remove('correct', 'wrong');
      if (!s4mcSelected.has(o.dataset.id)) return;
      o.classList.remove('selected');
      o.classList.add(S4MC.correct.has(o.dataset.id) ? 'correct' : 'wrong');
    });
    s4mcLastWrong = s4mcSig(s4mcSelected);
    document.getElementById('s4mc-check').disabled = true;
    S4MC_POPUP.retry.title = [...s4mcSelected].some(id => S4MC.correct.has(id)) ? RETRY_TITLE_PARTIAL : RETRY_TITLE;
    s4mcShowPopup('retry');
    return;
  }

  /* Final wrong — reveal the FULL correct set (including options never
     selected) and mark the learner's incorrect selections red. */
  s4mcOpts().forEach(o => {
    o.classList.remove('selected', 'correct', 'wrong');
    if (S4MC.correct.has(o.dataset.id)) o.classList.add('correct');
    else if (s4mcSelected.has(o.dataset.id)) o.classList.add('wrong');
  });
  s4mcResult = 'wrong2';
  s4mcFinish();
  s4mcShowPopup('wrong2');
}

function s4mcFinish() {
  s4mcDone = true;
  s4mcOpts().forEach(o => { o.disabled = true; });
  const chk = document.getElementById('s4mc-check');
  /* MultipleChoiceQuestion inherits SingleChoiceQuestion's completion
     fallback "هل نتابع؟" — slide 43 authors no progression copy of its
     own. This is NOT ValueInputQuestion's "متابعة"; the two templates
     deliberately do not share a label. */
  if (chk) setNavLabel(chk, 'هل نتابع؟');
  const hint = document.getElementById('s4mc-hint');
  if (hint) hint.classList.add('hidden');
  /* The Continue enable is left to s4mcShowPopup, which owns every
     button state this outcome implies. */
  syncPracticeProgressNav(document.getElementById('s4'));
}

function s4mcResetInitial() {
  s4mcSelected  = new Set();
  s4mcAttempts  = 0;
  s4mcDone      = false;
  s4mcResult    = null;
  s4mcLastWrong = null;
  s4mcOpts().forEach(o => {
    o.disabled = false;
    o.classList.remove('selected', 'correct', 'wrong');
    o.setAttribute('aria-checked', 'false');
  });
  const popup = document.getElementById('s4mc-popup');
  popup.classList.add('hidden');
  popup.classList.remove('is-scroll', 'needs-scroll');   /* defensive: this screen never sets them */
  document.getElementById('s4mc-hint-overlay').classList.add('hidden');
  const chk = document.getElementById('s4mc-check');
  if (chk) { setNavLabel(chk, 'هل إجابتي صحيحة؟'); chk.disabled = true; }
  const hint = document.getElementById('s4mc-hint');
  if (hint) { hint.classList.remove('hidden'); hint.disabled = false; }
}

function s4mcShowPopup(type) {
  const popup = document.getElementById('s4mc-popup');
  if (!popup) return;
  const cfg = S4MC_POPUP[type];
  const chk = document.getElementById('s4mc-check');
  popup.style.background = cfg.bg;
  /* A popup that is already open keeps the position the learner dragged
     it to; a fresh open starts at the shared default. */
  if (popup.classList.contains('hidden')) resetPopupPosition(popup);
  /* Terminal-incorrect feedback carries a fixed second line
     (_question-template-defaults.md → "Terminal-incorrect title carries
     a fixed second line") — APPENDED to, never replacing, the authored
     title, and only on the outcome that REVEALS the correct answer.
     This screen qualifies: wrong2 marks the two correct options and
     spells the reasoning out. There is no View My Answer Toggle here,
     so there is no answerView to exclude, and retry/correct are
     excluded by the rule itself. Same implementation as S1-S3. */
  document.getElementById('s4mc-popup-title').innerHTML =
    (type === 'wrong2') ? cfg.title + '<br>الإجابة الصحيحة تَظهَر الآن.' : cfg.title;
  document.getElementById('s4mc-popup-body').innerHTML = (cfg.body || []).map(p => {
    if (p.gap)  return '<p class="scq-popup-gap" aria-hidden="true"></p>';
    if (p.math) return '<p class="scq-math">' + p.html + '</p>';
    if (p.expr) return '<p class="scq-popup-expr viq-expr" dir="ltr"><bdi dir="ltr">' + p.text + '</bdi></p>';
    return '<p>' + p.text + '</p>';
  }).join('');
  renderPopupAction('s4mc-popup', null, null);   // this template carries no action row

  popup.classList.remove('hidden');
  clampPopup(popup);
  /* Nothing gates the progression button on this screen, so a resolved
     question hands it straight back. */
  if (chk && s4mcDone) chk.disabled = false;
}

function s4mcOpenHint() {
  xapiRequestedHint('003', 'q1');
  document.getElementById('s4mc-hint-overlay').classList.remove('hidden');
}
function s4mcCloseHint() {
  document.getElementById('s4mc-hint-overlay').classList.add('hidden');
}
function s4mcCloseHintOnBackdrop(e) {
  if (e.target === e.currentTarget) s4mcCloseHint();
}

/* Enter/Space on a focused option toggles it, mirroring native checkbox
   semantics. Delegated so it covers every option without per-node
   handlers. A <button> already fires click on both keys, so this only
   normalizes Space (which scrolls by default) and prevents the double
   toggle that would otherwise follow. */
document.addEventListener('keydown', e => {
  const isSpace = e.key === ' ' || e.key === 'Spacebar' || e.code === 'Space';
  if (!isSpace) return;
  const opt = e.target.closest && e.target.closest('#s4 .scq-opt');
  if (!opt || opt.disabled) return;
  e.preventDefault();
  s4mcToggle(opt.dataset.id);
});

/* ═══════════════════════════════════════════════════════════
   S5 — DropdownQuestion, SINGLE slot (slide 44)
   Practice question 3, sub-part ב — the SECOND of two sub-parts, so
   this screen is the one that RESOLVES progress item 3.

   Spec: 720-templates → DropdownQuestion.md + _question-template-
   defaults.md + _global-components.md → Feedback popup system. Ported
   from the approved sibling methodica-ar-math-ratio-03-02's own S7
   (two-slot DropdownQuestion) and collapsed to one slot, which
   DropdownQuestion.md names as the supported reduction ("For the
   single-dropdown variant this collapses to one slot with no per-slot
   loop"). Everything the two share — the retry gate's snapshot shape,
   the outside-click listener, keyboard operation of the list, the View
   My Answer Toggle, the lock rules — is the sibling's code, not a
   re-derivation.

   CORRECT VALUE = 'أكبر من'. Evidence, four independent strands, all
   recorded on this screen's markup in index.html: the producer's green
   a:highlight on exactly that option's run; both feedback bodies'
   closing sentence; the arithmetic the feedback spells out; and the
   slide's resolved producer comments.

   ATTEMPT MODEL: the 2-attempt default. The notes state no deviation,
   and the slide carries a hint plus one correct and one final-wrong
   panel — the normal shape of a 2-attempt question.

   USES the View My Answer Toggle: one slot whose display can hold
   either the learner's own pick or the correct value, which is exactly
   the "one slot, two possible values" applicability rule (contrast S4,
   whose multi-select shows the whole correct set alongside the
   learner's own and is therefore excluded).
   ═══════════════════════════════════════════════════════════ */
const S5DQ = { correct: 'أكبر من', maxAttempts: 2 };

window.lomdaState = window.lomdaState || {};
function s5dqState() {
  window.lomdaState.s5 = window.lomdaState.s5 || {
    attempts: 0,
    done: false,
    selection: null,
    locked: false,
    /* answer: the learner's final pick, captured BEFORE the reveal
       overwrites the display. null = nothing to toggle. */
    answer: null,
    view: 'solution',
    lastWrongSig: null
  };
  return window.lomdaState.s5;
}

/* Verdict icons — the CANONICAL component's own exported assets
   (node 110:1703, state=correct 120:2136 icon 120:3332 and
   state=Incorrect 120:2125 icon 120:3329), downloaded to
   assets/icons/ rather than hand-drawn, and rendered at the
   component's 16x16. Their fills are the component's own
   Feedback/Correct/300 #609E12 and Feedback/Incorrect/300 #B20010. */
const S5DQ_CHECK_SVG = '<img class="dq-check-ico" src="assets/icons/dropdown-correct.svg" alt="" aria-hidden="true">';
const S5DQ_WRONG_SVG = '<img class="dq-wrong-ico" src="assets/icons/dropdown-incorrect.svg" alt="" aria-hidden="true">';

/* Shared explanation — the authored body of BOTH panels (id=26 correct,
   id=16 final-wrong), which are structurally identical (P3..P10, same
   runs) and differ only in title, so this is the shared-explanation
   default rather than an assumption. `null` entries are the source's
   own authored EMPTY paragraphs (P4) and its space-only paragraph (P9),
   reproduced as real breaks.
     Every mathematical island is written in its MEASURED physical order
   inside a <bdi dir="ltr"> — see the measurement table in index.html.
   P6/P7 are standalone display lines with no Hebrew in them at all. */
const S5DQ_EXPLAIN = [
  'النسبة بين عدد الطوابق تحت الأرض وعدد الطوابق فوق الأرض هي <bdi class="viq-expr" dir="ltr">a : 3a</bdi> .',
  null,
  'إذا عوّضنا <bdi class="viq-expr" dir="ltr">a = 3</bdi> سنحصل على أن النسبة هي‎:',
  { math: '3 : (3 ⋅ 3)' },   // client (Nitza): parentheses around the product
  { math: '3 : 9' },
  'بعد الاختزال ستكون النسبة <bdi class="viq-expr" dir="ltr">1 : 3</bdi> .',
  null,
  'ولذلك يكون عدد الطوابق فوق الأرض أكبر من عدد الطوابق تحت الأرض‎.'
];

/* Titles are the slide's own (id=26 P1 / id=16 P1). The non-final retry
   state has no authored content on this slide, so it takes this part's
   canonical fallback — the same treatment S1-S4 use.
   NO `scrollGate` on any outcome: see the popup's comment in
   index.html for the measurement and the producer decision. */
const S5DQ_POPUP = {
  retry:   { bg: '#ffdbdc', title: 'الإجابة غير صحيحة.', body: ['هل نحاول مرة أخرى؟'] },
  correct: { bg: '#edf8ed', title: 'صحيح جدًا‎!', body: S5DQ_EXPLAIN },
  wrong2: {
    bg: '#ffdbdc', title: 'هذا خطأ. لا بأس، فكل خطأ يقوّي قدراتكم‎.', body: S5DQ_EXPLAIN,
    action: { label: 'الانتقال إلى إجابتي', view: 'answer' },
    answerView: {
      title: 'الإجابة غير صحيحة.',
      body:  ['يمكن مشاهدة الحل الصحيح مرة أخرى‎.'],
      action: { label: 'الإجابة الصحيحة', view: 'solution' }
    }
  }
};

function s5dqBtn()  { return document.getElementById('s5dq-btn'); }
function s5dqList() { return document.getElementById('s5dq-list'); }
function s5dqSel()  { return document.getElementById('s5dq-selected'); }

function s5dqCloseList() {
  const l = s5dqList(), b = s5dqBtn();
  if (l) l.classList.add('hidden');
  if (b) b.setAttribute('aria-expanded', 'false');
}

function s5dqToggle() {
  const st = s5dqState();
  if (st.done || st.locked) return;
  const b = s5dqBtn(), l = s5dqList();
  if (!b || !l) return;
  const willOpen = l.classList.contains('hidden');
  l.classList.toggle('hidden', !willOpen);
  b.setAttribute('aria-expanded', willOpen ? 'true' : 'false');
  /* Opening the dropdown while a non-final retry popup is showing is a
     corrective action, so it dismisses that popup
     (_global-components.md → "First-attempt retry feedback dismissal").
     Cannot fire once the question is done — the slot is disabled. */
  if (willOpen && st.attempts > 0 && !st.done) {
    document.getElementById('s5dq-popup').classList.add('hidden');
  }
  if (willOpen) {
    const opts = [...l.querySelectorAll('.dq-option')];
    const target = opts.find(o => o.dataset.value === st.selection) || opts[0];
    if (target) target.focus();
  }
}

function s5dqSelect(value) {
  const st = s5dqState();
  if (st.done || st.locked) return;
  st.selection = value;
  const b = s5dqBtn(), l = s5dqList(), sel = s5dqSel();
  sel.textContent = value;
  b.removeAttribute('data-empty');
  b.setAttribute('aria-label', 'المقارنة: ' + value);
  l.querySelectorAll('.dq-option').forEach(o =>
    o.setAttribute('aria-selected', o.dataset.value === value ? 'true' : 'false'));
  s5dqCloseList();
  if (st.attempts > 0 && !st.done) {
    document.getElementById('s5dq-popup').classList.add('hidden');
    s5dqMark(null);
  }
  /* Focus returns to the trigger, but on the NEXT task. Selecting with
     the keyboard dispatches keydown → keypress → keyup for one physical
     press, and only the keydown is ours: if the trigger is focused
     synchronously, the keypress that follows lands on a <button>, whose
     default Enter/Space action is a click — which would immediately
     re-open the list the selection just closed. Deferring the focus by
     one task lets the rest of the key sequence land on the (inert)
     option and keeps the closed list closed. Verified with real
     dispatched keys, not synthetic handler calls. */
  setTimeout(() => { if (!b.disabled) b.focus(); }, 0);
  s5dqUpdateBar();
}

/* Keyboard operation of the list: Enter/Space select, arrows move,
   Escape closes back to the button, Home/End jump. */
function s5dqOptionKey(e, value) {
  const l = s5dqList();
  if (!l) return;
  const opts = [...l.querySelectorAll('.dq-option')];
  const i = opts.indexOf(e.currentTarget);
  let t = null;
  if (e.key === 'Enter' || e.key === ' ' || e.key === 'Spacebar') { e.preventDefault(); s5dqSelect(value); return; }
  else if (e.key === 'ArrowDown') t = opts[Math.min(opts.length - 1, i + 1)];
  else if (e.key === 'ArrowUp')   t = opts[Math.max(0, i - 1)];
  else if (e.key === 'Home')      t = opts[0];
  else if (e.key === 'End')       t = opts[opts.length - 1];
  else if (e.key === 'Escape') { e.preventDefault(); s5dqCloseList(); s5dqBtn().focus(); return; }
  else return;
  e.preventDefault();
  if (t) t.focus();
}

/* One shared document-level listener, registered once: clicking outside
   the wrapper closes the open list. */
document.addEventListener('click', e => {
  if (e.target.closest && e.target.closest('.dq-dropdown-wrap')) return;
  s5dqCloseList();
});

/* Retry-gate comparison: the selection itself, normalized — reselecting
   the identical value must keep צדקתי? disabled
   (_question-template-defaults.md → "Retry gate"). */
function s5dqSignature() { return JSON.stringify(s5dqState().selection); }

function s5dqUpdateBar() {
  const st = s5dqState();
  const chk = document.getElementById('s5dq-check');
  if (!chk || st.done) return;
  chk.disabled = st.selection === null || s5dqSignature() === st.lastWrongSig;
}

/* Slot render. Three states plus null (clear).
   'correct'  — green, leaving the learner's own (correct) pick shown.
   'solution' — green, with the correct value overwriting the display.
   'answer'   — red, showing the learner's own final pick. */
function s5dqMark(state) {
  const b = s5dqBtn(), sel = s5dqSel();
  if (!b) return;
  b.classList.remove('dq-correct', 'dq-wrong');
  const old = b.querySelector('.dq-check-ico, .dq-wrong-ico');
  if (old) old.remove();
  if (state === null) return;
  const st = s5dqState();
  if (state === 'answer') {
    sel.textContent = st.answer === null ? 'اختاروا إجابة' : st.answer;
    b.classList.add('dq-wrong');
    b.insertAdjacentHTML('beforeend', S5DQ_WRONG_SVG);
  } else {
    if (state === 'solution') sel.textContent = S5DQ.correct;
    b.classList.add('dq-correct');
    b.insertAdjacentHTML('beforeend', S5DQ_CHECK_SVG);
  }
}

function s5dqLock() {
  const st = s5dqState();
  st.locked = true;
  const b = s5dqBtn();
  if (b) { b.disabled = true; b.setAttribute('aria-expanded', 'false'); }
  s5dqCloseList();
}

function s5dqCheck() {
  const st = s5dqState();
  if (st.done) { advanceScreen(); return; }
  if (st.selection === null) return;
  st.attempts++;
  s5dqCloseList();
  /* xAPI (2026-11-11): the layer writes Kata's form from the metadata — matching, the one sentence [.] the pick */
  reportQ(['003', 'q2'], st.selection === S5DQ.correct, st.selection === S5DQ.correct || st.attempts >= S5DQ.maxAttempts,
    { pairs: [{ source: 0, target: st.selection }] });

  if (st.selection === S5DQ.correct) {
    s5dqLock();
    s5dqMark('correct');
    st.answer = null;              // nothing to toggle
    s5dqFinish('correct');
    s5dqShowPopup('correct');
    return;
  }

  if (st.attempts < S5DQ.maxAttempts) {
    /* No marking on a non-final wrong attempt, and the slot stays
       selectable (DropdownQuestion.md → Reuse QA row 5). */
    st.lastWrongSig = s5dqSignature();
    document.getElementById('s5dq-check').disabled = true;
    s5dqShowPopup('retry');
    return;
  }

  /* Final wrong. Snapshot the learner's own pick BEFORE the reveal
     overwrites the display — the only moment it still exists. */
  st.answer = st.selection;
  st.view = 'solution';
  s5dqLock();
  s5dqFinish('wrong2');
  s5dqApplyView();
}

/* Single owner of the resolved-wrong state for the slot and the popup. */
function s5dqApplyView() {
  const st = s5dqState();
  if (st.answer === null) return;
  s5dqMark(st.view === 'answer' ? 'answer' : 'solution');
  s5dqShowPopup('wrong2', st.view);
}
/* Pure view switch — never routes through s5dqCheck(), never touches
   attempts or done. */
function s5dqToggleView() {
  const st = s5dqState();
  if (st.answer === null) return;
  st.view = (st.view === 'solution') ? 'answer' : 'solution';
  s5dqApplyView();
}

function s5dqFinish(result) {
  const st = s5dqState();
  st.done = true;
  s5dqLock();
  const chk = document.getElementById('s5dq-check');
  /* DropdownQuestion.md → "Completion label": this template's canonical
     fallback, and the label every other screen in this part ships. The
     slide authors no progression copy. */
  if (chk) { setNavLabel(chk, 'هل نتابع؟'); chk.disabled = false; }
  syncPracticeProgressNav(document.getElementById('s5'));
}

function s5dqResetInitial() {
  const st = s5dqState();
  st.attempts = 0;
  st.done = false;
  st.selection = null;
  st.locked = false;
  st.answer = null;
  st.view = 'solution';
  st.lastWrongSig = null;
  const b = s5dqBtn(), sel = s5dqSel(), l = s5dqList();
  if (b) {
    b.disabled = false;
    b.setAttribute('data-empty', 'true');
    b.setAttribute('aria-expanded', 'false');
    b.setAttribute('aria-label', 'المقارنة: اختاروا إجابة');
  }
  if (sel) sel.textContent = 'اختاروا إجابة';
  if (l) {
    l.classList.add('hidden');
    l.querySelectorAll('.dq-option').forEach(o => o.setAttribute('aria-selected', 'false'));
  }
  s5dqMark(null);
  document.getElementById('s5dq-popup').classList.add('hidden');
  document.getElementById('s5dq-hint-overlay').classList.add('hidden');
  const chk = document.getElementById('s5dq-check');
  if (chk) { setNavLabel(chk, 'هل إجابتي صحيحة؟'); chk.disabled = true; }
  const hint = document.getElementById('s5dq-hint');
  if (hint) { hint.classList.remove('hidden'); hint.disabled = false; }
}

function s5dqShowPopup(type, view) {
  const popup = document.getElementById('s5dq-popup');
  if (!popup) return;
  const base = S5DQ_POPUP[type];
  const cfg  = (view === 'answer' && base.answerView)
             ? Object.assign({}, base, base.answerView)
             : base;
  popup.style.background = cfg.bg;
  if (popup.classList.contains('hidden') || type === 'wrong2') resetPopupPosition(popup);
  /* Terminal-incorrect feedback carries a fixed second line
     (_question-template-defaults.md → "Terminal-incorrect title carries
     a fixed second line") — appended to, never replacing, the authored
     title, and NOT on the answerView, which shows the learner's own
     answer rather than the solution. */
  document.getElementById('s5dq-popup-title').innerHTML =
    (type === 'wrong2' && view !== 'answer')
      ? cfg.title + '<br>الإجابة الصحيحة تَظهَر الآن.'
      : cfg.title;
  document.getElementById('s5dq-popup-body').innerHTML =
    (cfg.body || []).map(p => {
      if (p === null)  return '<p class="scq-popup-gap" aria-hidden="true"></p>';
      if (p && p.math) return '<p class="scq-math"><bdi dir="ltr">' + p.math + '</bdi></p>';
      return '<p>' + p + '</p>';
    }).join('');
  renderPopupAction('s5dq-popup', cfg.action, s5dqToggleView);
  popup.classList.remove('hidden');
  clampPopup(popup);
  if (document.getElementById('s5dq-check') && s5dqState().done) {
    document.getElementById('s5dq-check').disabled = false;
  }
}

function s5dqOpenHint() {
  xapiRequestedHint('003', 'q2');
  document.getElementById('s5dq-hint-overlay').classList.remove('hidden');
}
function s5dqCloseHint() {
  document.getElementById('s5dq-hint-overlay').classList.add('hidden');
}
function s5dqCloseHintOnBackdrop(e) {
  if (e.target === e.currentTarget) s5dqCloseHint();
}

/* ═══════════════════════════════════════════════════════════
   S7 — ValueInputQuestion, Variant A with TWO independent values
   (slide 46). The FIRST question of this part's second practice block,
   the harder one slide 45 announces ("הנה עוד 2 תרגילים ברמת קושי
   גבוהה יותר").

   Spec: 720-templates → ValueInputQuestion.md + _question-template-
   defaults.md + _global-components.md. The state machine is this part's
   own S2 (the single-value Variant A screen) extended to two slots;
   nothing is re-derived.

   SCORING — CASE A, explicitly (ValueInputQuestion.md → "Scoring vs
   feedback granularity"): the two fields are two INDEPENDENT answers
   ("how many did עמית have" / "how many did עומר have"), so the
   question is ONE scored unit that needs both right, while each field
   is MARKED with its own actual result. A mixed board is correct
   behaviour: one green ✓ and one red ✕ while the question scores
   incorrect. This is the opposite of S1/S3, which are Case B — one
   ratio object, one combined icon.

   CORRECT VALUES: עמית 70, עומר 40. Three independent strands:
   notesSlide42 states them outright ("התשובות הנכונות הן: לעמית 70,
   לעומר 40"); the deck's own two answer boxes are pre-filled with
   exactly 70 (id=7) and 40 (id=31); and the mathematics the feedback
   spells out (7x − 15 = 4x + 15 → x = 10 → 70 and 40).

   ATTEMPT MODEL: the 2-attempt default; the notes state no deviation,
   and the slide carries a hint plus one correct and one final-wrong
   panel. Uses the View My Answer Toggle — the final-wrong reveal
   overwrites the very inputs the learner typed into, which is the
   destructive-reveal case the contract names.
   ═══════════════════════════════════════════════════════════ */
const S7VIQ = { correct: { amit: '70', omer: '40' }, maxAttempts: 2, slots: ['amit', 'omer'] };

let s7viqAttempts = 0;
let s7viqDone     = false;
let s7viqResult   = null;   // 'correct' | 'wrong2'
let s7viqLastWrong = null;  // signature of the immediately-preceding wrong attempt
let s7viqAnswer   = null;   // per-slot snapshot for the toggle; null = nothing to toggle
let s7viqView     = 'solution';

function s7viqInput(k) { return document.getElementById('s7viq-input-' + k); }
function s7viqField(k) { return document.getElementById('s7viq-field-' + k); }
function s7viqIco(k)   { return document.getElementById('s7viq-ico-' + k); }

/* Normalized snapshot of BOTH slots — the retry gate compares the whole
   pair, so retyping the identical pair keeps צדקתי? disabled
   (_question-template-defaults.md → "Retry gate"). */
function s7viqSig() {
  return JSON.stringify(S7VIQ.slots.map(k => (s7viqInput(k) ? s7viqInput(k).value.trim() : '')));
}
function s7viqFilled() {
  return S7VIQ.slots.every(k => s7viqInput(k) && s7viqInput(k).value.trim() !== '');
}

function s7viqOnInput() {
  if (s7viqDone) return;
  /* Numbers only (client, 2026-09-29): the same keystroke-and-paste
     filter as every other numeric ValueInput screen in the lomda. */
  S7VIQ.slots.forEach(k => {
    const el = s7viqInput(k);
    const cleaned = el.value.replace(/[^0-9,.-]/g, '');
    if (cleaned !== el.value) el.value = cleaned;
  });
  /* A genuine edit clears the previous attempt's per-field marks and
     dismisses its retry popup — approved path 2 of "No manual close". */
  if (s7viqAttempts > 0) {
    document.getElementById('s7viq-popup').classList.add('hidden');
    S7VIQ.slots.forEach(k => s7viqMark(k, null));
  }
  s7viqUpdateBar();
}

function s7viqUpdateBar() {
  const chk = document.getElementById('s7viq-check');
  if (!chk || s7viqDone) return;
  chk.disabled = !s7viqFilled() || s7viqSig() === s7viqLastWrong;
}

/* Per-field mark. Case A: every field shows its OWN result.
   null = clear | 'correct' | 'wrong' | 'solution' (the revealed value). */
function s7viqMark(k, state) {
  const f = s7viqField(k), ico = s7viqIco(k), inp = s7viqInput(k);
  if (!f) return;
  /* Same classes and the same two icon constants this part's S1/S2/S3
     already use — the verdict glyph lives INSIDE the bordered field. */
  f.classList.remove('viq-correct', 'viq-wrong');
  if (ico) ico.innerHTML = '';
  if (state === null) return;
  if (state === 'wrong') { f.classList.add('viq-wrong'); if (ico) ico.innerHTML = S1VIQ_ICON_WRONG; return; }
  if (state === 'solution' && inp) inp.value = S7VIQ.correct[k];
  f.classList.add('viq-correct');
  if (ico) ico.innerHTML = S1VIQ_ICON_CORRECT;
}

function s7viqLock() {
  S7VIQ.slots.forEach(k => { const i = s7viqInput(k); if (i) { i.readOnly = true; i.setAttribute('aria-readonly', 'true'); } });
}

function s7viqCheck() {
  if (s7viqDone) { advanceScreen(); return; }
  if (!s7viqFilled()) return;
  s7viqAttempts++;
  const got = {};
  S7VIQ.slots.forEach(k => { got[k] = s7viqInput(k).value.trim(); });
  const perSlot = S7VIQ.slots.map(k => ({ k, ok: got[k] === S7VIQ.correct[k] }));
  const allRight = perSlot.every(s => s.ok);
  /* xAPI (2026-11-11): the layer writes Kata's form from the metadata — performance, 'עמית[.]70[,]עומר[.]40' (proposed to Kata, `other` there) */
  reportQ(['004', 'q1'], allRight, allRight || s7viqAttempts >= S7VIQ.maxAttempts,
    { steps: { 'حنان': got.amit, 'سامر': got.omer } });

  if (allRight) {
    S7VIQ.slots.forEach(k => s7viqMark(k, 'correct'));
    s7viqLock();
    s7viqAnswer = null;                 // nothing to toggle
    s7viqFinish('correct');
    s7viqShowPopup('correct');
    return;
  }

  if (s7viqAttempts < S7VIQ.maxAttempts) {
    /* Retry: mark ONLY the fields that are actually wrong — scoring is
       combined, marking is per field. */
    perSlot.forEach(s => s7viqMark(s.k, s.ok ? 'correct' : 'wrong'));
    s7viqLastWrong = s7viqSig();
    document.getElementById('s7viq-check').disabled = true;
    S7VIQ_POPUP.retry.title = perSlot.some(s => s.ok) ? RETRY_TITLE_PARTIAL : RETRY_TITLE;
    s7viqShowPopup('retry');
    return;
  }

  /* Final wrong. Snapshot each slot's own value AND its own correctness
     BEFORE the reveal overwrites a wrong field — the only moment the
     learner's actual answer still exists. */
  s7viqAnswer = {};
  perSlot.forEach(s => { s7viqAnswer[s.k] = { value: got[s.k], correct: s.ok }; });
  s7viqView = 'solution';
  s7viqLock();
  s7viqFinish('wrong2');
  s7viqApplyView();
}

/* Single owner of the resolved-wrong state for both fields and the popup. */
function s7viqApplyView() {
  if (s7viqAnswer === null) return;
  S7VIQ.slots.forEach(k => {
    const a = s7viqAnswer[k];
    if (a.correct) { s7viqMark(k, 'correct'); s7viqInput(k).value = a.value; return; }
    if (s7viqView === 'answer') { s7viqInput(k).value = a.value; s7viqMark(k, 'wrong'); }
    else s7viqMark(k, 'solution');
  });
  s7viqShowPopup('wrong2', s7viqView);
}
function s7viqToggleView() {
  if (s7viqAnswer === null) return;
  s7viqView = (s7viqView === 'solution') ? 'answer' : 'solution';
  s7viqApplyView();
}

function s7viqFinish(result) {
  s7viqDone = true;
  s7viqResult = result;
  const chk = document.getElementById('s7viq-check');
  /* ValueInputQuestion.md → "Completion label": this template's canonical
     fallback is "متابعة", NOT SingleChoiceQuestion's "هل نتابع؟" — the same
     label screen 1 already ships in this part. Slide 46 authors no
     progression copy of its own. The enable here is provisional: when the
     outcome is long, s7viqShowPopup → s7viqArmScrollGate re-disables it and
     s7viqCheckScrollGate is what finally releases it. */
  if (chk) { setNavLabel(chk, 'متابعة'); chk.disabled = false; }
  /* The template hides the hint on finish (ValueInputQuestion.md →
     "{prefix}Finish(result) … hides hint"), as screen 1 does. */
  const hint = document.getElementById('s7viq-hint');
  if (hint) hint.classList.add('hidden');
  syncPracticeProgress2Nav(document.getElementById('s7'));
}

function s7viqResetInitial() {
  s7viqGateOk = true;
  s7viqGateKey = null;
  s7viqGateRead = null;
  s7viqAttempts = 0;
  s7viqDone = false;
  s7viqResult = null;
  s7viqLastWrong = null;
  s7viqAnswer = null;
  s7viqView = 'solution';
  S7VIQ.slots.forEach(k => {
    const i = s7viqInput(k);
    if (i) { i.value = ''; i.readOnly = false; i.removeAttribute('aria-readonly'); }
    s7viqMark(k, null);
  });
  document.getElementById('s7viq-popup').classList.add('hidden');
  document.getElementById('s7viq-hint-overlay').classList.add('hidden');
  const chk = document.getElementById('s7viq-check');
  if (chk) { setNavLabel(chk, 'هل إجابتي صحيحة؟'); chk.disabled = true; }
  const hint = document.getElementById('s7viq-hint');
  if (hint) { hint.classList.remove('hidden'); hint.disabled = false; }
}

/* Shared explanation — the authored body of BOTH panels (id=26 P3..P19
   and id=28 P4..P20), which are identical run for run and differ only
   in the title, so this is the shared-explanation default rather than an
   assumption. `null` entries are the source's own EMPTY paragraphs.
     Every expression is written in its MEASURED physical order. The
   equation is the one that matters: its source runs are "4x + 15" /
   " " / "7x – 15 =", which PowerPoint orders right-to-left AS RUNS and
   renders as 7x – 15 = 4x + 15 — read off a COM render, not off the run
   concatenation. Single-run expressions render as authored (verified on
   the same render). The en dash in the equation and the hyphen in
   "7x - 15" are each the source's own character. */
const S7VIQ_EXPLAIN = [
  'نرمز: عدد بطاقات حنان هو <bdi class="viq-expr" dir="ltr">7x</bdi> وعدد بطاقات سامر هو <bdi class="viq-expr" dir="ltr">4x</bdi>.',
  null,
  'أعطت حنان لسامر 15 بطاقة، لذلك‎:',
  'بقيَت لدى حنان <bdi class="viq-expr" dir="ltr">7x - 15</bdi> بطاقة.',
  'أصبح لدى سامر الآن <bdi class="viq-expr" dir="ltr">4x + 15</bdi> بطاقة.',
  null,
  'وأصبح لدى الاثنتين الآن نفس عدد البطاقات، ولذلك‎:',
  { math: '7x – 15 = 4x + 15' },
  { math: '3x = 30' },
  { math: 'x = 10' },
  null,
  'إذًا، قبل نقل البطاقات‎:',
  'جمعت حنان <bdi class="viq-expr" dir="ltr">10 ⋅ 7 = 70</bdi>',
  'أي 70 بطاقة‎.',
  null,
  'وجمع سامر <bdi class="viq-expr" dir="ltr">10 ⋅ 4 = 40</bdi>',
  'أي 40 بطاقة‎.'
];

/* Titles are the slide's own. The final-wrong panel authors TWO bold
   title paragraphs (id=28 P1 and P2), so both are title lines. The
   non-final retry state has no authored content, so it takes this
   part's canonical fallback. */
const S7VIQ_POPUP = {
  retry:   { bg: '#ffdbdc', title: 'الإجابة غير صحيحة.', body: ['هل نحاول مرة أخرى؟'] },
  correct: { bg: '#edf8ed', title: 'هذا صحيح‎!', body: S7VIQ_EXPLAIN },
  wrong2: {
    bg: '#ffdbdc', title: 'غير دقيق.<br>لا بأس، دعونا نشرح‎:', body: S7VIQ_EXPLAIN,
    action: { label: 'الانتقال إلى إجابتي', view: 'answer' },
    answerView: {
      title: 'الإجابة غير صحيحة.',
      body:  ['يمكن مشاهدة الحل الصحيح مرة أخرى‎.'],
      action: { label: 'الإجابة الصحيحة', view: 'solution' }
    }
  }
};

function s7viqShowPopup(type, view) {
  const popup = document.getElementById('s7viq-popup');
  if (!popup) return;
  const base = S7VIQ_POPUP[type];
  const cfg  = (view === 'answer' && base.answerView) ? Object.assign({}, base, base.answerView) : base;
  popup.style.background = cfg.bg;
  if (popup.classList.contains('hidden') || type === 'wrong2') resetPopupPosition(popup);
  /* Terminal-incorrect title carries the fixed second line — appended to,
     never replacing, the authored title, and never on the answerView. */
  document.getElementById('s7viq-popup-title').innerHTML =
    (type === 'wrong2' && view !== 'answer') ? cfg.title + '<br>الإجابة الصحيحة تَظهَر الآن.' : cfg.title;
  document.getElementById('s7viq-popup-body').innerHTML =
    (cfg.body || []).map(p => {
      if (p === null)  return '<p class="scq-popup-gap" aria-hidden="true"></p>';
      if (p && p.math) return '<p class="scq-math"><bdi dir="ltr">' + p.math + '</bdi></p>';
      return '<p>' + p + '</p>';
    }).join('');
  renderPopupAction('s7viq-popup', cfg.action, s7viqToggleView);
  popup.classList.remove('hidden');
  clampPopup(popup);
  s7viqArmScrollGate(popup, type, view);
  const chk = document.getElementById('s7viq-check');
  if (chk && s7viqDone && s7viqGateOk) chk.disabled = false;
}

/* ── Scroll gate ──
   Slide 46's note asks for it ("שימו לב שבמשוב ארוך, לאפשר גלילה") and
   here — unlike screens 4 and 5 — the body genuinely overflows: the
   authored panels are 437.6x645.1, half as tall again as slides 43/44's
   495.1, and the rendered body does not fit the canvas. The gate is
   therefore ARMED on the two long outcomes and never on the short retry.
   Uses this part's existing generic machinery (atBottom / placeScrollHint
   / the .is-scroll CSS / popupPointerDown's scroll guard) — the same
   mechanism screens 4 and 5 keep available but do not need. */
/* SCROLL-GATE MEMORY (retrofit, 2026-09-15). The gate belongs to a
   FEEDBACK STATE, not to an opening. `GateKey` is the state on screen
   and `GateRead` the state already read to the end, so re-showing a
   state that was read leaves the gate satisfied — closing and reopening
   the same feedback never asks for the same read twice. A different
   state re-arms it by inequality, and s7viqResetInitial() clears the
   memory because a new question must be read again. */
let s7viqGateKey = null;
let s7viqGateRead = null;
let s7viqGateOk = true;
function s7viqArmScrollGate(popup, type, view) {
  const scroller = document.getElementById('s7viq-popup-scroll');
  const long = popup.querySelector('#s7viq-popup-body').children.length > 3;
  popup.classList.toggle('is-scroll', long);
  s7viqGateKey = long ? (type + '|' + (view || 'solution')) : null;
  /* Already read in THIS state → stay satisfied across a close/reopen. */
  s7viqGateOk = !long || s7viqGateRead === s7viqGateKey;
  if (!long) { popup.classList.remove('needs-scroll'); scroller.removeAttribute('tabindex'); return; }
  const chk = document.getElementById('s7viq-check');
  if (chk && !s7viqGateOk) chk.disabled = true;
  scroller.setAttribute('tabindex', '0');
  scroller.scrollTop = 0;
  placeScrollHint('s7viq-popup');
  s7viqCheckScrollGate();     // unlocks at once if it does not actually overflow
}
function s7viqCheckScrollGate() {
  if (s7viqGateOk) {
    /* Already satisfied — clear a hint left over from a previous,
       still-locked presentation of this popup. */
    const p0 = document.getElementById('s7viq-popup');
    if (p0) p0.classList.remove('needs-scroll');
    return;
  }
  const scroller = document.getElementById('s7viq-popup-scroll');
  const popup = document.getElementById('s7viq-popup');
  if (!scroller || !popup) return;
  if (!atBottom(scroller)) { popup.classList.add('needs-scroll'); return; }
  s7viqGateOk = true;
  s7viqGateRead = s7viqGateKey;      // this state has now been read to the end
  const chk = document.getElementById('s7viq-check');
  if (chk && s7viqDone) chk.disabled = false;
  popup.classList.remove('needs-scroll');
}
let s7viqResizeRaf = null;
window.addEventListener('resize', () => {
  const popup = document.getElementById('s7viq-popup');
  if (!popup || popup.classList.contains('hidden')) return;
  placeScrollHint('s7viq-popup');
  if (s7viqGateOk) return;
  if (s7viqResizeRaf) cancelAnimationFrame(s7viqResizeRaf);
  s7viqResizeRaf = requestAnimationFrame(s7viqCheckScrollGate);
});

function s7viqOpenHint() {
  xapiRequestedHint('004', 'q1');
  document.getElementById('s7viq-hint-overlay').classList.remove('hidden');
}
function s7viqCloseHint() {
  document.getElementById('s7viq-hint-overlay').classList.add('hidden');
}
function s7viqCloseHintOnBackdrop(e) {
  if (e.target === e.currentTarget) s7viqCloseHint();
}


/* ═══════════════════════════════════════════════════════════
   S8 — DragAndDropQuestion, IMAGE TOP layout (slide 47)
   Second practice block, question 2, section א. Spec:
   720-templates → DragAndDropQuestion.md (Image Top variant) +
   _question-template-defaults.md + _global-components.md → "View My
   Answer Toggle". The state machine is ported from that template's own
   Image Top reference implementation (methodica-science-mass-measure-01-01
   S16) — same template, same variant, same 2-attempt + toggle contract;
   only geometry, copy and the answer key differ. The popup, hint overlay
   and nav-label apparatus are this part's own, as on screens 1-7.

   ANSWER KEY — machine 1 = 7x, machine 2 = 3x, on two independent
   strands: the source marks exactly "7x" and "3x" with a green
   <a:highlight val="00FF00"/> (the producer's answer key) and leaves
   "10x" / "3+x" unmarked; and both feedback panels say it in words
   ("במכונה 1 יש 7x ממתקים ובמכונה 2 יש 3x ממתקים"). The ratio itself
   agrees: machine 1 holds MORE, and the ratio is 7 : 3.

   RATIO ORDER — physical left-to-right 7 : 3, the 3 adjacent to the
   preceding Hebrew word. PRODUCER RULING 2026-09-14 ("3 מימין ו7
   משמאל"), which also governs screen 9 / slide 48. Slides 47 and 48
   render this same sentence MIRRORED from each other in the deck —
   47 splits the ratio across three runs, which PowerPoint reorders as
   runs, while 48 uses one run, which plain bidi mirrors — so one of
   the two had to be normalised; the ruling picked slide 47's hand,
   which is also slide 46's and screen 4's.
   ═══════════════════════════════════════════════════════════ */

const S8DDQ = {
  /* targetId → draggableId */
  correctMap: {
    's8target-m1': 's8drag-7x',
    's8target-m2': 's8drag-3x'
  },
  revealMap: {
    's8target-m1': 's8drag-7x',
    's8target-m2': 's8drag-3x'
  },
  maxAttempts: 2
};

/* Draggable labels, for the card the render injects into a target.
   Every one is a pure LTR mathematical token, so each is written in its
   own displayed order and isolated — never left to automatic bidi. */
const S8DDQ_TEXTS = {
  's8drag-3x':    '3x',
  's8drag-10x':   '10x',
  's8drag-3plusx': '3+x',
  's8drag-7x':    '7x'
};

/* Shared explanation — the authored body of BOTH panels (id=15 P2..P4
   and id=7 P3..P5), identical run for run, differing only in the title.
   `null` is the source's own empty paragraph. */
const S8DDQ_EXPLAIN = [
  'النسبة بين عدد قطع الحلوى في الآلة 1 وعدد قطع الحلوى في الآلة 2 هي <bdi class="viq-expr" dir="ltr" role="img" aria-label="نسبة سبعة إلى ثلاثة">7 : 3</bdi> .',
  null,
  'ولذلك، بعد التوسيع بـ <bdi class="viq-expr" dir="ltr">x</bdi> مرات، يمكن القول إنه في الآلة 1 يوجد <bdi class="viq-expr" dir="ltr">7x</bdi> قطع حلوى، وفي الآلة 2 يوجد <bdi class="viq-expr" dir="ltr">3x</bdi> قطع حلوى.'
];

/* Titles are the slide's own. The final-wrong panel authors TWO bold
   title paragraphs (id=7 P0 and P1), so both are title lines. The
   non-final retry state has no authored content, so it takes this
   part's canonical fallback — the same one screens 4, 5 and 7 ship. */
const S8DDQ_POPUP = {
  retry:   { bg: '#ffdbdc', title: 'الإجابة غير صحيحة.', body: ['هل نحاول مرة أخرى؟'] },
  correct: { bg: '#edf8ed', title: 'رائع!', body: S8DDQ_EXPLAIN },
  wrong2: {
    bg: '#ffdbdc', title: 'الإجابة غير صحيحة.<br>لا بأس، تعالوا نفهم ذلك معًا.', body: S8DDQ_EXPLAIN,
    action: { label: 'الانتقال إلى إجابتي', view: 'answer' },
    answerView: {
      title: 'الإجابة غير صحيحة.',
      body:  ['يمكن مشاهدة الحل الصحيح مرة أخرى‎.'],
      action: { label: 'الإجابة الصحيحة', view: 'solution' }
    }
  }
};

/* draggableId → 'source' | targetId */
const s8ddqPlacement = {
  's8drag-3x':     'source',
  's8drag-10x':    'source',
  's8drag-3plusx': 'source',
  's8drag-7x':     'source'
};
let s8ddqChecked      = false;
let s8ddqDone         = false;
let s8ddqAttempts     = 0;
let s8ddqResult       = null;   // 'correct' | 'wrong2'
let s8ddqShowFeedback = false;  // per-zone ✓/✕ currently displayed
let s8ddqDnd          = null;
let s8ddqLastWrong    = null;   // placement signature of the last non-final wrong attempt
let s8ddqAnswer       = null;   // placement snapshot for the toggle; null = nothing to toggle
let s8ddqView         = 'solution';

/* Verdict icons — the same Figma vectors screens 1-7 already use for a
   value field's verdict, at the drag zone's larger badge size. */
const S8DDQ_ICON_CORRECT =
  '<svg viewBox="0 0 12.7273 9.18763" fill="none" aria-hidden="true"><path d="M4.10942 9.18763C3.74814 9.18778 3.40167 9.04418 3.14642 8.78851L0.234956 5.87812C-0.0783188 5.56475 -0.0783188 5.05677 0.234956 4.74339C0.548332 4.43012 1.05631 4.43012 1.36969 4.74339L4.10942 7.48312L11.3576 0.234956C11.671 -0.0783188 12.1789 -0.0783188 12.4923 0.234956C12.8056 0.548332 12.8056 1.05631 12.4923 1.36969L5.07242 8.78851C4.81717 9.04418 4.47069 9.18778 4.10942 9.18763Z" fill="currentColor"/></svg>';
const S8DDQ_ICON_WRONG =
  '<svg viewBox="0 0 10 10" fill="none" aria-hidden="true"><path d="M9.76736 0.232637C9.61836 0.0836796 9.4163 0 9.20561 0C8.99492 0 8.79286 0.0836796 8.64386 0.232637L5 3.87649L1.35615 0.232637C1.20714 0.0836796 1.00508 0 0.794391 0C0.583702 0 0.381639 0.0836796 0.232637 0.232637C0.0836796 0.381639 0 0.583702 0 0.794391C0 1.00508 0.0836796 1.20714 0.232637 1.35615L3.87649 5L0.232637 8.64386C0.0836796 8.79286 0 8.99492 0 9.20561C0 9.4163 0.0836796 9.61836 0.232637 9.76736C0.381639 9.91632 0.583702 10 0.794391 10C1.00508 10 1.20714 9.91632 1.35615 9.76736L5 6.12351L8.64386 9.76736C8.79286 9.91632 8.99492 10 9.20561 10C9.4163 10 9.61836 9.91632 9.76736 9.76736C9.91632 9.61836 10 9.4163 10 9.20561C10 8.99492 9.91632 8.79286 9.76736 8.64386L6.12351 5L9.76736 1.35615C9.91632 1.20714 10 1.00508 10 0.794391C10 0.583702 9.91632 0.381639 9.76736 0.232637Z" fill="currentColor"/></svg>';

function s8ddqSig() { return JSON.stringify(s8ddqPlacement); }

/* ── Render — one state object → DOM ── */
function s8ddqRender() {
  Object.keys(s8ddqPlacement).forEach(id => {
    const card = document.getElementById(id);
    if (card) card.classList.toggle('ghost', s8ddqPlacement[id] !== 'source');
  });

  const targetIds = [...new Set([
    ...Object.keys(S8DDQ.correctMap),
    ...Object.values(s8ddqPlacement).filter(v => v !== 'source')
  ])];

  targetIds.forEach(targetId => {
    const zone = document.getElementById(targetId);
    if (!zone) return;
    const existing = zone.querySelector('.s8ddq-placed-card');
    if (existing) existing.remove();

    const dragId = Object.keys(s8ddqPlacement).find(k => s8ddqPlacement[k] === targetId) || null;
    if (dragId) {
      const card = document.createElement('div');
      card.className = 's8ddq-placed-card';
      const txt = document.createElement('span');
      txt.className = 's8ddq-card-text';
      const bdi = document.createElement('bdi');
      bdi.setAttribute('dir', 'ltr');
      bdi.textContent = S8DDQ_TEXTS[dragId];
      txt.appendChild(bdi);
      card.appendChild(txt);
      if (!s8ddqChecked && s8ddqDnd) s8ddqDnd.attachSource(card, dragId);
      zone.appendChild(card);
      zone.classList.add('occupied');
    } else {
      zone.classList.remove('occupied', 'drag-over');
    }
  });

  s8ddqUpdateBar();
}

function s8ddqAllFilled() {
  return Object.keys(S8DDQ.correctMap).every(tId =>
    Object.keys(s8ddqPlacement).some(dId => s8ddqPlacement[dId] === tId));
}

function s8ddqUpdateBar() {
  const chk = document.getElementById('s8ddq-check');
  if (!chk || s8ddqDone) return;
  /* Retry gate (_question-template-defaults.md → "Retry gate"): the FULL
     placement signature, so re-submitting the same board — even after
     moving a distractor around — stays disabled. */
  chk.disabled = !s8ddqAllFilled() || s8ddqSig() === s8ddqLastWrong;
}

/* ── Pointer DnD wiring — once, idempotent ── */
function s8ddqInitDnd() {
  if (s8ddqDnd) return;
  s8ddqDnd = createPointerDnd({
    canDrag: (dragId, elem) => {
      if (s8ddqChecked) return false;
      if (elem.id === dragId) return s8ddqPlacement[dragId] === 'source';
      return true;   // a placed card carries no id
    },
    onPick: (dragId) => {
      s8ddqClearAttemptFeedback();
      if (s8ddqPlacement[dragId] !== 'source') {
        s8ddqPlacement[dragId] = 'source';
        s8ddqRender();
      }
      const el = document.getElementById(dragId);
      if (el) el.classList.add('dragging');
    },
    onDrop: (dragId, targetId) => {
      /* One item per zone: an occupant is evicted back to the tray. */
      const evicted = Object.keys(s8ddqPlacement).find(k => s8ddqPlacement[k] === targetId);
      if (evicted && evicted !== dragId) s8ddqPlacement[evicted] = 'source';
      s8ddqPlacement[dragId] = targetId;
      s8ddqRender();
    },
    onCancel: () => { s8ddqRender(); }
  });
  Object.keys(s8ddqPlacement).forEach(id => {
    const el = document.getElementById(id);
    if (el) s8ddqDnd.attachSource(el, id);
  });
  document.querySelectorAll('#s8 .s8ddq-target').forEach(el => s8ddqDnd.attachTarget(el, el.id));
}

/* ── Check ── */
function s8ddqCheck() {
  if (s8ddqDone) { advanceScreen(); return; }
  if (!s8ddqAllFilled()) return;

  s8ddqClearAttemptFeedback();
  s8ddqAttempts++;

  const allCorrect = Object.keys(S8DDQ.correctMap)
    .every(tId => s8ddqPlacement[S8DDQ.correctMap[tId]] === tId);
  reportQ(['005', 'q1'], allCorrect, allCorrect || s8ddqAttempts >= S8DDQ.maxAttempts, s8ddqAnswerSpec());

  if (allCorrect) {
    s8ddqChecked = true;
    s8ddqRender();
    s8ddqLock();
    s8ddqMarkZones(true);
    s8ddqAnswer = null;               // nothing to toggle
    s8ddqFinish('correct');
    s8ddqShowPopup('correct');
    return;
  }

  if (s8ddqAttempts < S8DDQ.maxAttempts) {
    /* Non-final wrong: per-zone ✓/✕ against the learner's own board,
       board stays live, check disabled until the board actually changes. */
    s8ddqShowFeedback = true;
    s8ddqLastWrong = s8ddqSig();
    s8ddqMarkZones(false);
    S8DDQ_POPUP.retry.title = Object.keys(S8DDQ.correctMap).some(tId => s8ddqPlacement[S8DDQ.correctMap[tId]] === tId) ? RETRY_TITLE_PARTIAL : RETRY_TITLE;
    s8ddqShowPopup('retry');
    const chk = document.getElementById('s8ddq-check');
    if (chk) chk.disabled = true;
    return;
  }

  /* Terminal wrong. Snapshot the learner's board BEFORE the reveal
     overwrites the very slots it lives in, then resolve into the
     solution view. */
  s8ddqChecked = true;
  s8ddqAnswer = Object.assign({}, s8ddqPlacement);
  s8ddqView = 'solution';
  s8ddqFinish('wrong2');
  s8ddqApplyView();
}

function s8ddqLock() {
  Object.keys(s8ddqPlacement).forEach(id => {
    const card = document.getElementById(id);
    if (card) card.classList.add('locked');
  });
  document.querySelectorAll('#s8 .s8ddq-placed-card').forEach(c => c.classList.add('locked'));
}

/* Per-zone verdict icon + border. `allGreen` forces the correct state,
   for the solution view where every zone shows the right answer. */
function s8ddqMarkZones(allGreen) {
  Object.keys(S8DDQ.correctMap).forEach(tId => {
    const zone = document.getElementById(tId);
    if (!zone) return;
    const ok = allGreen || (s8ddqPlacement[S8DDQ.correctMap[tId]] === tId);
    const card = zone.querySelector('.s8ddq-placed-card');
    if (card) {
      const old = card.querySelector('.s8ddq-ficon');
      if (old) old.remove();
      const ico = document.createElement('span');
      ico.className = 's8ddq-ficon ' + (ok ? 's8ddq-ficon--correct' : 's8ddq-ficon--wrong');
      ico.innerHTML = ok ? S8DDQ_ICON_CORRECT : S8DDQ_ICON_WRONG;
      card.appendChild(ico);
    }
    zone.classList.remove('s8ddq-correct', 's8ddq-wrong');
    zone.classList.add(ok ? 's8ddq-correct' : 's8ddq-wrong');
  });
}

/* The retry feedback must vanish the moment the learner starts
   correcting the board (_global-components.md → Feedback popup system). */
function s8ddqClearAttemptFeedback() {
  if (!s8ddqShowFeedback) return;
  s8ddqShowFeedback = false;
  const popup = document.getElementById('s8ddq-popup');
  if (popup) popup.classList.add('hidden');
  Object.keys(S8DDQ.correctMap).forEach(tId => {
    const zone = document.getElementById(tId);
    if (!zone) return;
    const ico = zone.querySelector('.s8ddq-ficon');
    if (ico) ico.remove();
    zone.classList.remove('s8ddq-correct', 's8ddq-wrong');
  });
}

function s8ddqRevealCorrect() {
  Object.keys(S8DDQ.revealMap).forEach(tId => { s8ddqPlacement[S8DDQ.revealMap[tId]] = tId; });
  const assigned = new Set(Object.values(S8DDQ.revealMap));
  Object.keys(s8ddqPlacement).forEach(id => { if (!assigned.has(id)) s8ddqPlacement[id] = 'source'; });
  s8ddqRender();
}

/* Single owner of the resolved-wrong state. */
function s8ddqApplyView() {
  if (s8ddqAnswer === null) return;
  if (s8ddqView === 'answer') {
    Object.assign(s8ddqPlacement, s8ddqAnswer);
    s8ddqRender();
    s8ddqMarkZones(false);     // per-zone verdict on the learner's own board
  } else {
    s8ddqRevealCorrect();
    s8ddqMarkZones(true);
  }
  s8ddqLock();                 // re-created placed cards need re-locking
  s8ddqShowPopup('wrong2', s8ddqView);
}

function s8ddqToggleView() {
  if (s8ddqAnswer === null) return;
  s8ddqView = (s8ddqView === 'solution') ? 'answer' : 'solution';
  s8ddqApplyView();
}

function s8ddqFinish(result) {
  s8ddqDone = true;
  s8ddqResult = result;
  const chk = document.getElementById('s8ddq-check');
  /* DragAndDropQuestion.md states no fallback label of its own; its own
     Image Top reference implementation ships "هل نتابع؟", which is also
     what every non-ValueInput screen in this part uses. Slide 47
     authors no progression copy. */
  if (chk) { setNavLabel(chk, 'هل نتابع؟'); chk.disabled = false; }
  const hint = document.getElementById('s8ddq-hint');
  if (hint) hint.classList.add('hidden');
  syncPracticeProgress2Nav(document.getElementById('s8'));
}

function s8ddqResetInitial() {
  s8ddqChecked = false;
  s8ddqDone = false;
  s8ddqAttempts = 0;
  s8ddqResult = null;
  s8ddqShowFeedback = false;
  s8ddqLastWrong = null;
  s8ddqAnswer = null;
  s8ddqView = 'solution';
  Object.keys(s8ddqPlacement).forEach(id => { s8ddqPlacement[id] = 'source'; });
  document.querySelectorAll('#s8 .s8ddq-drag-card').forEach(c => c.classList.remove('locked', 'dragging'));
  Object.keys(S8DDQ.correctMap).forEach(tId => {
    const zone = document.getElementById(tId);
    if (zone) zone.classList.remove('s8ddq-correct', 's8ddq-wrong', 'occupied', 'drag-over');
  });
  const popup = document.getElementById('s8ddq-popup');
  if (popup) popup.classList.add('hidden');
  const overlay = document.getElementById('s8ddq-hint-overlay');
  if (overlay) overlay.classList.add('hidden');
  const chk = document.getElementById('s8ddq-check');
  if (chk) { setNavLabel(chk, 'هل إجابتي صحيحة؟'); chk.disabled = true; }
  const hint = document.getElementById('s8ddq-hint');
  if (hint) { hint.classList.remove('hidden'); hint.disabled = false; }
  s8ddqRender();
}

/* Entry point for resetScreenState(8). Same resume/reopen contract as
   every other question screen here: a resolved question never hands out
   a fresh attempt and never comes back without its feedback, always on
   the SOLUTION view. */
function s8ddqEnter() {
  s8ddqInitDnd();
  if (!s8ddqDone) { if (screenPristine(8)) s8ddqResetInitial(); }
  else if (s8ddqAnswer === null) {
    s8ddqRender(); s8ddqLock(); s8ddqMarkZones(true); s8ddqShowPopup('correct');
  } else {
    s8ddqView = 'solution';
    s8ddqApplyView();
  }
  syncPracticeProgress2Nav(document.getElementById('s8'));
}

function s8ddqShowPopup(type, view) {
  const popup = document.getElementById('s8ddq-popup');
  if (!popup) return;
  const base = S8DDQ_POPUP[type];
  const cfg  = (view === 'answer' && base.answerView) ? Object.assign({}, base, base.answerView) : base;
  popup.style.background = cfg.bg;
  if (popup.classList.contains('hidden') || type === 'wrong2') resetPopupPosition(popup);
  /* Terminal-incorrect title carries the fixed second line — appended to,
     never replacing, the authored title, and never on the answerView. */
  document.getElementById('s8ddq-popup-title').innerHTML =
    (type === 'wrong2' && view !== 'answer') ? cfg.title + '<br>الإجابة الصحيحة تَظهَر الآن.' : cfg.title;
  document.getElementById('s8ddq-popup-body').innerHTML =
    (cfg.body || []).map(p => (p === null
      ? '<p class="scq-popup-gap" aria-hidden="true"></p>'
      : '<p>' + p + '</p>')).join('');
  renderPopupAction('s8ddq-popup', cfg.action, s8ddqToggleView);
  popup.classList.remove('hidden');
  clampPopup(popup);
}

function s8ddqOpenHint()  { xapiRequestedHint('005', 'q1'); document.getElementById('s8ddq-hint-overlay').classList.remove('hidden'); }
function s8ddqCloseHint() { document.getElementById('s8ddq-hint-overlay').classList.add('hidden'); }
function s8ddqCloseHintOnBackdrop(e) { if (e.target === e.currentTarget) s8ddqCloseHint(); }


/* ═══════════════════════════════════════════════════════════
   S9 — DragAndDropQuestion, IMAGE TOP layout (slide 48)
   Second practice block, question 2, section ב. Same template, same
   variant and the same state machine as screen 8 — only the board shape
   (three equation cards into ONE zone), the copy and the answer key
   differ.

   ANSWER KEY — "3x + 120 = 7x", on two independent strands: the source
   green-highlights exactly that card and leaves "3x + 7x = 120" and
   "7x + 120 = 3x" unmarked, and both feedback panels end on the same
   equation. Every expression here is written in its MEASURED physical
   order, read off a PowerPoint render of the slide with the three cards
   moved aside so nothing overlapped the panel — not off run
   concatenation.

   RATIO ORDER — physical left-to-right 7 : 3, the 3 adjacent to
   "להיות", per the producer's 2026-09-14 ruling. Slide 48 alone renders
   it the other way round because its sentence is a SINGLE run (plain
   bidi mirrors a spaced `A : B`) where slide 47 splits it across three
   runs (PowerPoint reorders those as runs); the ruling normalised both
   to slide 47's hand. NOTE the emphasis differs from slide 47 on
   purpose: slide 48 authors this whole sentence as one un-bolded run,
   so nothing in it is bold here.
   ═══════════════════════════════════════════════════════════ */

const S9DDQ = {
  correctMap: { 's9target': 's9drag-c' },
  revealMap:  { 's9target': 's9drag-c' },
  maxAttempts: 2
};

const S9DDQ_TEXTS = {
  's9drag-a': '3x + 7x = 120',
  's9drag-b': '7x + 120 = 3x',
  's9drag-c': '3x + 120 = 7x'
};

/* Shared explanation — the authored body of BOTH panels (id=15 P2..P6
   and id=7 P3..P7), identical run for run. `null` is the source's own
   empty paragraph; the closing equation is its own display line. */
const S9DDQ_EXPLAIN = [
  'في الآلة 1 يوجد 120 قطعة حلوى أكثر‎.',
  null,
  'ولذلك نضيف 120 إلى الكمية الموجودة في الآلة الصغيرة ونساويها مع الكمية الموجودة في الآلة الكبيرة‎:',
  null,
  { math: '3x + 120 = 7x' }
];

/* Titles are the slide's own; the final-wrong panel authors two bold
   title paragraphs (id=7 P0 and P1). The non-final retry state has no
   authored content and takes this part's canonical fallback. */
const S9DDQ_POPUP = {
  retry:   { bg: '#ffdbdc', title: 'الإجابة غير صحيحة.', body: ['هل نحاول مرة أخرى؟'] },
  correct: { bg: '#edf8ed', title: 'صحيح‎!', body: S9DDQ_EXPLAIN },
  wrong2: {
    bg: '#ffdbdc', title: 'هذا خطأ.<br>لا بأس، دعونا نشرح‎.', body: S9DDQ_EXPLAIN,
    action: { label: 'الانتقال إلى إجابتي', view: 'answer' },
    answerView: {
      title: 'الإجابة غير صحيحة.',
      body:  ['يمكن مشاهدة الحل الصحيح مرة أخرى‎.'],
      action: { label: 'الإجابة الصحيحة', view: 'solution' }
    }
  }
};

const s9ddqPlacement = { 's9drag-a': 'source', 's9drag-b': 'source', 's9drag-c': 'source' };
let s9ddqChecked      = false;
let s9ddqDone         = false;
let s9ddqAttempts     = 0;
let s9ddqResult       = null;   // 'correct' | 'wrong2'
let s9ddqShowFeedback = false;
let s9ddqDnd          = null;
let s9ddqLastWrong    = null;
let s9ddqAnswer       = null;
let s9ddqView         = 'solution';

function s9ddqSig() { return JSON.stringify(s9ddqPlacement); }

function s9ddqRender() {
  Object.keys(s9ddqPlacement).forEach(id => {
    const card = document.getElementById(id);
    if (card) card.classList.toggle('ghost', s9ddqPlacement[id] !== 'source');
  });

  const targetIds = [...new Set([
    ...Object.keys(S9DDQ.correctMap),
    ...Object.values(s9ddqPlacement).filter(v => v !== 'source')
  ])];

  targetIds.forEach(targetId => {
    const zone = document.getElementById(targetId);
    if (!zone) return;
    const existing = zone.querySelector('.s9ddq-placed-card');
    if (existing) existing.remove();

    const dragId = Object.keys(s9ddqPlacement).find(k => s9ddqPlacement[k] === targetId) || null;
    if (dragId) {
      const card = document.createElement('div');
      card.className = 's9ddq-placed-card';
      const txt = document.createElement('span');
      txt.className = 's9ddq-card-text';
      const bdi = document.createElement('bdi');
      bdi.setAttribute('dir', 'ltr');
      bdi.textContent = S9DDQ_TEXTS[dragId];
      txt.appendChild(bdi);
      card.appendChild(txt);
      if (!s9ddqChecked && s9ddqDnd) s9ddqDnd.attachSource(card, dragId);
      zone.appendChild(card);
      zone.classList.add('occupied');
    } else {
      zone.classList.remove('occupied', 'drag-over');
    }
  });

  s9ddqUpdateBar();
}

function s9ddqAllFilled() {
  return Object.keys(S9DDQ.correctMap).every(tId =>
    Object.keys(s9ddqPlacement).some(dId => s9ddqPlacement[dId] === tId));
}

function s9ddqUpdateBar() {
  const chk = document.getElementById('s9ddq-check');
  if (!chk || s9ddqDone) return;
  chk.disabled = !s9ddqAllFilled() || s9ddqSig() === s9ddqLastWrong;
}

function s9ddqInitDnd() {
  if (s9ddqDnd) return;
  s9ddqDnd = createPointerDnd({
    canDrag: (dragId, elem) => {
      if (s9ddqChecked) return false;
      if (elem.id === dragId) return s9ddqPlacement[dragId] === 'source';
      return true;
    },
    onPick: (dragId) => {
      s9ddqClearAttemptFeedback();
      if (s9ddqPlacement[dragId] !== 'source') {
        s9ddqPlacement[dragId] = 'source';
        s9ddqRender();
      }
      const el = document.getElementById(dragId);
      if (el) el.classList.add('dragging');
    },
    onDrop: (dragId, targetId) => {
      const evicted = Object.keys(s9ddqPlacement).find(k => s9ddqPlacement[k] === targetId);
      if (evicted && evicted !== dragId) s9ddqPlacement[evicted] = 'source';
      s9ddqPlacement[dragId] = targetId;
      s9ddqRender();
    },
    onCancel: () => { s9ddqRender(); }
  });
  Object.keys(s9ddqPlacement).forEach(id => {
    const el = document.getElementById(id);
    if (el) s9ddqDnd.attachSource(el, id);
  });
  document.querySelectorAll('#s9 .s9ddq-target').forEach(el => s9ddqDnd.attachTarget(el, el.id));
}

function s9ddqCheck() {
  if (s9ddqDone) { advanceScreen(); return; }
  if (!s9ddqAllFilled()) return;

  s9ddqClearAttemptFeedback();
  s9ddqAttempts++;

  const allCorrect = Object.keys(S9DDQ.correctMap)
    .every(tId => s9ddqPlacement[S9DDQ.correctMap[tId]] === tId);
  reportQ(['005', 'q2'], allCorrect, allCorrect || s9ddqAttempts >= S9DDQ.maxAttempts, s9ddqAnswerSpec());

  if (allCorrect) {
    s9ddqChecked = true;
    s9ddqRender();
    s9ddqLock();
    s9ddqMarkZones(true);
    s9ddqAnswer = null;
    s9ddqFinish('correct');
    s9ddqShowPopup('correct');
    return;
  }

  if (s9ddqAttempts < S9DDQ.maxAttempts) {
    s9ddqShowFeedback = true;
    s9ddqLastWrong = s9ddqSig();
    s9ddqMarkZones(false);
    S9DDQ_POPUP.retry.title = Object.keys(S9DDQ.correctMap).some(tId => s9ddqPlacement[S9DDQ.correctMap[tId]] === tId) ? RETRY_TITLE_PARTIAL : RETRY_TITLE;
    s9ddqShowPopup('retry');
    const chk = document.getElementById('s9ddq-check');
    if (chk) chk.disabled = true;
    return;
  }

  s9ddqChecked = true;
  s9ddqAnswer = Object.assign({}, s9ddqPlacement);
  s9ddqView = 'solution';
  s9ddqFinish('wrong2');
  s9ddqApplyView();
}

function s9ddqLock() {
  Object.keys(s9ddqPlacement).forEach(id => {
    const card = document.getElementById(id);
    if (card) card.classList.add('locked');
  });
  document.querySelectorAll('#s9 .s9ddq-placed-card').forEach(c => c.classList.add('locked'));
}

function s9ddqMarkZones(allGreen) {
  Object.keys(S9DDQ.correctMap).forEach(tId => {
    const zone = document.getElementById(tId);
    if (!zone) return;
    const ok = allGreen || (s9ddqPlacement[S9DDQ.correctMap[tId]] === tId);
    const card = zone.querySelector('.s9ddq-placed-card');
    if (card) {
      const old = card.querySelector('.s9ddq-ficon');
      if (old) old.remove();
      const ico = document.createElement('span');
      ico.className = 's9ddq-ficon ' + (ok ? 's9ddq-ficon--correct' : 's9ddq-ficon--wrong');
      ico.innerHTML = ok ? S8DDQ_ICON_CORRECT : S8DDQ_ICON_WRONG;
      card.appendChild(ico);
    }
    zone.classList.remove('s9ddq-correct', 's9ddq-wrong');
    zone.classList.add(ok ? 's9ddq-correct' : 's9ddq-wrong');
  });
}

function s9ddqClearAttemptFeedback() {
  if (!s9ddqShowFeedback) return;
  s9ddqShowFeedback = false;
  const popup = document.getElementById('s9ddq-popup');
  if (popup) popup.classList.add('hidden');
  Object.keys(S9DDQ.correctMap).forEach(tId => {
    const zone = document.getElementById(tId);
    if (!zone) return;
    const ico = zone.querySelector('.s9ddq-ficon');
    if (ico) ico.remove();
    zone.classList.remove('s9ddq-correct', 's9ddq-wrong');
  });
}

function s9ddqRevealCorrect() {
  Object.keys(S9DDQ.revealMap).forEach(tId => { s9ddqPlacement[S9DDQ.revealMap[tId]] = tId; });
  const assigned = new Set(Object.values(S9DDQ.revealMap));
  Object.keys(s9ddqPlacement).forEach(id => { if (!assigned.has(id)) s9ddqPlacement[id] = 'source'; });
  s9ddqRender();
}

function s9ddqApplyView() {
  if (s9ddqAnswer === null) return;
  if (s9ddqView === 'answer') {
    Object.assign(s9ddqPlacement, s9ddqAnswer);
    s9ddqRender();
    s9ddqMarkZones(false);
  } else {
    s9ddqRevealCorrect();
    s9ddqMarkZones(true);
  }
  s9ddqLock();
  s9ddqShowPopup('wrong2', s9ddqView);
}

function s9ddqToggleView() {
  if (s9ddqAnswer === null) return;
  s9ddqView = (s9ddqView === 'solution') ? 'answer' : 'solution';
  s9ddqApplyView();
}

function s9ddqFinish(result) {
  s9ddqDone = true;
  s9ddqResult = result;
  const chk = document.getElementById('s9ddq-check');
  /* Same completion label as screen 8 — approved 2026-09-14. */
  if (chk) { setNavLabel(chk, 'هل نتابع؟'); chk.disabled = false; }
  const hint = document.getElementById('s9ddq-hint');
  if (hint) hint.classList.add('hidden');
  syncPracticeProgress2Nav(document.getElementById('s9'));
}

function s9ddqResetInitial() {
  s9ddqChecked = false;
  s9ddqDone = false;
  s9ddqAttempts = 0;
  s9ddqResult = null;
  s9ddqShowFeedback = false;
  s9ddqLastWrong = null;
  s9ddqAnswer = null;
  s9ddqView = 'solution';
  Object.keys(s9ddqPlacement).forEach(id => { s9ddqPlacement[id] = 'source'; });
  document.querySelectorAll('#s9 .s9ddq-drag-card').forEach(c => c.classList.remove('locked', 'dragging'));
  Object.keys(S9DDQ.correctMap).forEach(tId => {
    const zone = document.getElementById(tId);
    if (zone) zone.classList.remove('s9ddq-correct', 's9ddq-wrong', 'occupied', 'drag-over');
  });
  const popup = document.getElementById('s9ddq-popup');
  if (popup) popup.classList.add('hidden');
  const overlay = document.getElementById('s9ddq-hint-overlay');
  if (overlay) overlay.classList.add('hidden');
  const chk = document.getElementById('s9ddq-check');
  if (chk) { setNavLabel(chk, 'هل إجابتي صحيحة؟'); chk.disabled = true; }
  const hint = document.getElementById('s9ddq-hint');
  if (hint) { hint.classList.remove('hidden'); hint.disabled = false; }
  s9ddqRender();
}

function s9ddqEnter() {
  s9ddqInitDnd();
  if (!s9ddqDone) { if (screenPristine(9)) s9ddqResetInitial(); }
  else if (s9ddqAnswer === null) {
    s9ddqRender(); s9ddqLock(); s9ddqMarkZones(true); s9ddqShowPopup('correct');
  } else {
    s9ddqView = 'solution';
    s9ddqApplyView();
  }
  syncPracticeProgress2Nav(document.getElementById('s9'));
}

function s9ddqShowPopup(type, view) {
  const popup = document.getElementById('s9ddq-popup');
  if (!popup) return;
  const base = S9DDQ_POPUP[type];
  const cfg  = (view === 'answer' && base.answerView) ? Object.assign({}, base, base.answerView) : base;
  popup.style.background = cfg.bg;
  if (popup.classList.contains('hidden') || type === 'wrong2') resetPopupPosition(popup);
  document.getElementById('s9ddq-popup-title').innerHTML =
    (type === 'wrong2' && view !== 'answer') ? cfg.title + '<br>الإجابة الصحيحة تَظهَر الآن.' : cfg.title;
  document.getElementById('s9ddq-popup-body').innerHTML =
    (cfg.body || []).map(p => {
      if (p === null)  return '<p class="scq-popup-gap" aria-hidden="true"></p>';
      if (p && p.math) return '<p class="scq-math"><bdi dir="ltr">' + p.math + '</bdi></p>';
      return '<p>' + p + '</p>';
    }).join('');
  renderPopupAction('s9ddq-popup', cfg.action, s9ddqToggleView);
  popup.classList.remove('hidden');
  clampPopup(popup);
}

function s9ddqOpenHint()  { xapiRequestedHint('005', 'q2'); document.getElementById('s9ddq-hint-overlay').classList.remove('hidden'); }
function s9ddqCloseHint() { document.getElementById('s9ddq-hint-overlay').classList.add('hidden'); }
function s9ddqCloseHintOnBackdrop(e) { if (e.target === e.currentTarget) s9ddqCloseHint(); }


/* ═══════════════════════════════════════════════════════════
   S10 — ValueInputQuestion (slide 49)
   Second practice block, question 2, section ג — and the LAST screen of
   this part. Variant A with TWO independent values, scored as ONE
   outcome but marked field by field (ValueInputQuestion.md → "Scoring
   vs feedback granularity", Case A) — the same shape screen 7 ships,
   and the state machine is that screen's, with slide 49's own copy,
   answers and geometry.

   CORRECT VALUES — machine 1 = 210, machine 2 = 90, on two independent
   strands: the presenter note states them outright ("למכונה 1: 210",
   "למכונה 2: 90"), and both feedback panels derive them in full
   (x = 30, then 7·30 = 210 and 3·30 = 90). The deck's own placeholder
   boxes hold the same two numbers.

   Every expression below is written in its MEASURED physical order,
   read off a PowerPoint render of the slide with the panels moved
   on-canvas and everything that overlapped them moved aside — not off
   run concatenation. All four display lines render exactly as authored.

   RESOLVES QUESTION 2. This is the third and last sub-part of the
   second progress item, so answering it is what finally gives that item
   a ✓ or ✕: correct only when sections א, ב AND ג are all correct,
   incorrect for every other completed combination. That rule lives in
   derivePractice2State() and needs no special case here — registering
   screen 10 as the third sub-part is the whole change.

   END OF PART. advanceScreen() past this screen is a deliberate no-op:
   this part has no confirmed forward cross-part destination yet (see
   the TOTAL_SCREENS note at the top of this file). The Continue control
   still ENABLES on completion, which is the established behaviour.
   ═══════════════════════════════════════════════════════════ */

const S10VIQ = { correct: { m1: '210', m2: '90' }, maxAttempts: 2, slots: ['m1', 'm2'] };

let s10viqAttempts  = 0;
let s10viqDone      = false;
let s10viqResult    = null;   // 'correct' | 'wrong2'
let s10viqLastWrong = null;   // signature of the immediately-preceding wrong attempt
let s10viqAnswer    = null;   // per-slot snapshot for the toggle; null = nothing to toggle
let s10viqView      = 'solution';

function s10viqInput(k) { return document.getElementById('s10viq-input-' + k); }
function s10viqField(k) { return document.getElementById('s10viq-field-' + k); }
function s10viqIco(k)   { return document.getElementById('s10viq-ico-' + k); }

function s10viqSig() {
  return JSON.stringify(S10VIQ.slots.map(k => (s10viqInput(k) ? s10viqInput(k).value.trim() : '')));
}
function s10viqFilled() {
  return S10VIQ.slots.every(k => s10viqInput(k) && s10viqInput(k).value.trim() !== '');
}

function s10viqOnInput() {
  if (s10viqDone) return;
  /* Numbers only (client, 2026-09-29) — same filter as S7. */
  S10VIQ.slots.forEach(k => {
    const el = s10viqInput(k);
    const cleaned = el.value.replace(/[^0-9,.-]/g, '');
    if (cleaned !== el.value) el.value = cleaned;
  });
  /* Retry feedback clears the moment the learner starts correcting. */
  if (s10viqAttempts > 0) {
    document.getElementById('s10viq-popup').classList.add('hidden');
    S10VIQ.slots.forEach(k => s10viqMark(k, null));
  }
  s10viqUpdateBar();
}

function s10viqUpdateBar() {
  const chk = document.getElementById('s10viq-check');
  if (!chk || s10viqDone) return;
  chk.disabled = !s10viqFilled() || s10viqSig() === s10viqLastWrong;
}

function s10viqMark(k, state) {
  const f = s10viqField(k), ico = s10viqIco(k), inp = s10viqInput(k);
  if (!f) return;
  f.classList.remove('viq-correct', 'viq-wrong');
  if (state === 'correct' || state === 'solution') f.classList.add('viq-correct');
  else if (state === 'wrong') f.classList.add('viq-wrong');
  if (ico) {
    ico.innerHTML = !state ? ''
      : ((state === 'wrong') ? S1VIQ_ICON_WRONG : S1VIQ_ICON_CORRECT);
  }
  if (state === 'solution' && inp) inp.value = S10VIQ.correct[k];
}

function s10viqLock() {
  S10VIQ.slots.forEach(k => {
    const i = s10viqInput(k);
    if (i) { i.readOnly = true; i.setAttribute('aria-readonly', 'true'); }
  });
}

function s10viqCheck() {
  if (s10viqDone) { advanceScreen(); return; }
  if (!s10viqFilled()) return;
  s10viqAttempts++;

  const got = {};
  S10VIQ.slots.forEach(k => { got[k] = s10viqInput(k).value.trim(); });
  const perSlot = S10VIQ.slots.map(k => ({ k, ok: got[k] === S10VIQ.correct[k] }));
  const allOk = perSlot.every(s => s.ok);
  /* xAPI (2026-11-11): the layer writes Kata's form from the metadata — performance, 'מכונה 1[.]210[,]מכונה 2[.]90' (proposed to Kata, `other` there) */
  reportQ(['005', 'q3'], allOk, allOk || s10viqAttempts >= S10VIQ.maxAttempts,
    { steps: { 'الآلة 1': got.m1, 'الآلة 2': got.m2 } });

  if (allOk) {
    S10VIQ.slots.forEach(k => s10viqMark(k, 'correct'));
    s10viqLock();
    s10viqAnswer = null;
    s10viqFinish('correct');
    s10viqShowPopup('correct');
    return;
  }

  if (s10viqAttempts < S10VIQ.maxAttempts) {
    /* Case A: one combined score, but each box still shows its own
       verdict so the learner knows which half to fix. */
    perSlot.forEach(s => s10viqMark(s.k, s.ok ? 'correct' : 'wrong'));
    s10viqLastWrong = s10viqSig();
    document.getElementById('s10viq-check').disabled = true;
    S10VIQ_POPUP.retry.title = perSlot.some(s => s.ok) ? RETRY_TITLE_PARTIAL : RETRY_TITLE;
    s10viqShowPopup('retry');
    return;
  }

  /* Terminal wrong — snapshot the learner's values BEFORE the solution
     overwrites the same inputs they live in. */
  s10viqAnswer = {};
  perSlot.forEach(s => { s10viqAnswer[s.k] = { value: got[s.k], correct: s.ok }; });
  s10viqView = 'solution';
  s10viqLock();
  s10viqFinish('wrong2');
  s10viqApplyView();
}

function s10viqApplyView() {
  if (s10viqAnswer === null) return;
  S10VIQ.slots.forEach(k => {
    const a = s10viqAnswer[k];
    if (a.correct) { s10viqMark(k, 'correct'); s10viqInput(k).value = a.value; return; }
    if (s10viqView === 'answer') { s10viqInput(k).value = a.value; s10viqMark(k, 'wrong'); }
    else s10viqMark(k, 'solution');
  });
  s10viqShowPopup('wrong2', s10viqView);
}

function s10viqToggleView() {
  if (s10viqAnswer === null) return;
  s10viqView = (s10viqView === 'solution') ? 'answer' : 'solution';
  s10viqApplyView();
}

function s10viqFinish(result) {
  s10viqDone = true;
  s10viqResult = result;
  const chk = document.getElementById('s10viq-check');
  /* ValueInputQuestion.md → "Completion label": this template's canonical
     fallback "متابعة", as screens 1 and 7 ship. Slide 49 authors no
     progression copy. The enable is provisional — s10viqArmScrollGate()
     re-disables it while the long feedback is unread. */
  if (chk) { setNavLabel(chk, 'متابعة'); chk.disabled = false; }
  const hint = document.getElementById('s10viq-hint');
  if (hint) hint.classList.add('hidden');
  syncPracticeProgress2Nav(document.getElementById('s10'));
}

function s10viqResetInitial() {
  s10viqGateOk = true;
  s10viqGateKey = null;
  s10viqGateRead = null;
  s10viqAttempts = 0;
  s10viqDone = false;
  s10viqResult = null;
  s10viqLastWrong = null;
  s10viqAnswer = null;
  s10viqView = 'solution';
  S10VIQ.slots.forEach(k => {
    const i = s10viqInput(k);
    if (i) { i.value = ''; i.readOnly = false; i.removeAttribute('aria-readonly'); }
    s10viqMark(k, null);
  });
  document.getElementById('s10viq-popup').classList.add('hidden');
  document.getElementById('s10viq-hint-overlay').classList.add('hidden');
  const chk = document.getElementById('s10viq-check');
  if (chk) { setNavLabel(chk, 'هل إجابتي صحيحة؟'); chk.disabled = true; }
  const hint = document.getElementById('s10viq-hint');
  if (hint) { hint.classList.remove('hidden'); hint.disabled = false; }
}

/* Shared explanation — the authored body of BOTH panels (id=26 P2..P13
   and id=28 P3..P14), identical run for run and differing only in the
   title. `null` entries are the source's own EMPTY paragraphs. The
   multiplication sign is the source's own U+2219. */
const S10VIQ_EXPLAIN = [
  'نحل المعادلة‎:',
  { math: '3x + 120 = 7x' },
  { math: '4x = 120' },
  { math: 'x = 30' },
  null,
  'في الآلة 1 يوجد‎:',
  { math: '7∙30 = 210' },
  'أي 210 قطع حلوى‎.',
  null,
  'في الآلة 2 يوجد‎:',
  { math: '3∙30 = 90' },
  'أي 90 قطعة حلوى‎.'
];

const S10VIQ_POPUP = {
  retry:   { bg: '#ffdbdc', title: 'الإجابة غير صحيحة.', body: ['هل نحاول مرة أخرى؟'] },
  correct: { bg: '#edf8ed', title: 'هذا صحيح‎!', body: S10VIQ_EXPLAIN },
  wrong2: {
    /* The final-wrong panel authors TWO bold title paragraphs (id=28 P0
       and P1), so both are title lines. */
    bg: '#ffdbdc', title: 'غير دقيق.<br>التمرين الإضافي سيساعد على الفهم‎.', body: S10VIQ_EXPLAIN,
    action: { label: 'الانتقال إلى إجابتي', view: 'answer' },
    answerView: {
      title: 'الإجابة غير صحيحة.',
      body:  ['يمكن مشاهدة الحل الصحيح مرة أخرى‎.'],
      action: { label: 'الإجابة الصحيحة', view: 'solution' }
    }
  }
};

function s10viqShowPopup(type, view) {
  const popup = document.getElementById('s10viq-popup');
  if (!popup) return;
  const base = S10VIQ_POPUP[type];
  const cfg  = (view === 'answer' && base.answerView) ? Object.assign({}, base, base.answerView) : base;
  popup.style.background = cfg.bg;
  if (popup.classList.contains('hidden') || type === 'wrong2') resetPopupPosition(popup);
  document.getElementById('s10viq-popup-title').innerHTML =
    (type === 'wrong2' && view !== 'answer') ? cfg.title + '<br>الإجابة الصحيحة تَظهَر الآن.' : cfg.title;
  document.getElementById('s10viq-popup-body').innerHTML =
    (cfg.body || []).map(p => {
      if (p === null)  return '<p class="scq-popup-gap" aria-hidden="true"></p>';
      if (p && p.math) return '<p class="scq-math"><bdi dir="ltr">' + p.math + '</bdi></p>';
      return '<p>' + p + '</p>';
    }).join('');
  renderPopupAction('s10viq-popup', cfg.action, s10viqToggleView);
  popup.classList.remove('hidden');
  clampPopup(popup);
  s10viqArmScrollGate(popup, type, view);
  const chk = document.getElementById('s10viq-check');
  if (chk && s10viqDone && s10viqGateOk) chk.disabled = false;
}

/* ── Scroll gate ──
   Slide 49's note asks for it by name ("שימו לב שבמשוב ארוך, לאפשר
   גלילה או הוא נראה מגוכך מעל התמונה") and the body genuinely
   overflows: the authored panels are 437.6x560.1 with twelve body
   paragraphs. Armed on the two long outcomes, never on the short
   retry — the same machinery screen 7 uses. */
/* SCROLL-GATE MEMORY (retrofit, 2026-09-15). The gate belongs to a
   FEEDBACK STATE, not to an opening. `GateKey` is the state on screen
   and `GateRead` the state already read to the end, so re-showing a
   state that was read leaves the gate satisfied — closing and reopening
   the same feedback never asks for the same read twice. A different
   state re-arms it by inequality, and s10viqResetInitial() clears the
   memory because a new question must be read again. */
let s10viqGateKey = null;
let s10viqGateRead = null;
let s10viqGateOk = true;
function s10viqArmScrollGate(popup, type, view) {
  const scroller = document.getElementById('s10viq-popup-scroll');
  const long = popup.querySelector('#s10viq-popup-body').children.length > 3;
  popup.classList.toggle('is-scroll', long);
  s10viqGateKey = long ? (type + '|' + (view || 'solution')) : null;
  /* Already read in THIS state → stay satisfied across a close/reopen. */
  s10viqGateOk = !long || s10viqGateRead === s10viqGateKey;
  if (!long) { popup.classList.remove('needs-scroll'); scroller.removeAttribute('tabindex'); return; }
  const chk = document.getElementById('s10viq-check');
  if (chk && !s10viqGateOk) chk.disabled = true;
  scroller.setAttribute('tabindex', '0');
  scroller.scrollTop = 0;
  placeScrollHint('s10viq-popup');
  s10viqCheckScrollGate();     // unlocks at once if it does not actually overflow
}
function s10viqCheckScrollGate() {
  if (s10viqGateOk) {
    /* Already satisfied — clear a hint left over from a previous,
       still-locked presentation of this popup. */
    const p0 = document.getElementById('s10viq-popup');
    if (p0) p0.classList.remove('needs-scroll');
    return;
  }
  const scroller = document.getElementById('s10viq-popup-scroll');
  const popup = document.getElementById('s10viq-popup');
  if (!scroller || !popup) return;
  if (!atBottom(scroller)) { popup.classList.add('needs-scroll'); return; }
  s10viqGateOk = true;
  s10viqGateRead = s10viqGateKey;      // this state has now been read to the end
  const chk = document.getElementById('s10viq-check');
  if (chk && s10viqDone) chk.disabled = false;
  popup.classList.remove('needs-scroll');
}
let s10viqResizeRaf = null;
window.addEventListener('resize', () => {
  const popup = document.getElementById('s10viq-popup');
  if (!popup || popup.classList.contains('hidden')) return;
  placeScrollHint('s10viq-popup');
  if (s10viqGateOk) return;
  if (s10viqResizeRaf) cancelAnimationFrame(s10viqResizeRaf);
  s10viqResizeRaf = requestAnimationFrame(s10viqCheckScrollGate);
});

function s10viqOpenHint()  { xapiRequestedHint('005', 'q3'); document.getElementById('s10viq-hint-overlay').classList.remove('hidden'); }
function s10viqCloseHint() { document.getElementById('s10viq-hint-overlay').classList.add('hidden'); }
function s10viqCloseHintOnBackdrop(e) { if (e.target === e.currentTarget) s10viqCloseHint(); }


/* ═══════════════════════════════════════════════════════════
   MASTERY CHECKPOINTS AND THE CROSSING INTO PART 04
   Producer decision 2026-09-14. This part carries TWO consecutive,
   SEPARATE assessment stages; their scores are never combined.

   CHECKPOINT 1 — first bundle, three scored questions, decided when
   leaving S5 (advanceScreen()); producer decision 2026-09-30, which
   replaces the 2026-09-14 route that sent a failing learner to S7:
     2 or 3 correct → Pass → S6, then the two harder questions S7-S10.
     0 or 1 correct → Fail → remain on S5; המשך stays ENABLED and
                             nothing advances internally.
   S6-S10 are therefore reachable only by a learner who passed.

   CHECKPOINT 2 — the harder questions, decided at S10 (the last screen):
     Q1 (S7) correct AND Q2 (S8-S10, sections א/ב/ג) correct → Pass →
                             cross to 03-04, landing on its S0.
     anything else          → Fail → remain on S10; המשך stays ENABLED and
                             nothing advances internally.
   ⚠️ Since 2026-11-09 (08.11 sheet fix 4) checkpoint 2 no longer decides the REPORTED result: the
   component 'completed' carries checkpoint 1 alone (partGateResult), the only threshold S0 promises.
   The S10 end itself is the platform routing below (decision 2026-10-01), pass or fail.

   MECHANISM. Both Pass routes go through PartNav.forward() — the
   project's one cross-part contract (js/partnav.js, the same engine
   03-02 crosses with) — which records the trail entry that gives 03-04's
   S0 its Back button. No second navigation implementation exists here.
   A Fail never calls it, so a Fail leaves no trail entry and no route.

   REPORTING THE OUTCOME IS NOT THIS LOMDA'S JOB. The Ministry
   reporting/integration layer is owned by Vadim and is added outside
   lomda production: no SCORM, no xAPI, no cmi.*, no postMessage. The
   Fail branches therefore do nothing but not-advance — deliberately.

   RE-CROSSING. The part's answers come back from the resume state
   document (unit-js/45-painters.js + the PART_SCREENS registry at the
   end of this file), which replaced the earlier sessionStorage snapshot
   taken at the crossing. Only S10 crosses.
   ═══════════════════════════════════════════════════════════ */
const NEXT_PART = '04';

function resolvedState(s) { return s === 'correct' || s === 'incorrect'; }

/* Checkpoint 1 — the first series' three items, derived exactly as the
   progress nav derives them (an item resolves only once all its authored
   sub-parts are answered), so the gate and the dots can never disagree. */
function mastery1State() {
  const results = practiceProgress.questions.map(q => derivePracticeState(q.number, false));
  const correctCount = results.filter(r => r === 'correct').length;
  return {
    checkpoint: 1, results, correctCount,
    allResolved: results.every(resolvedState),
    passed: correctCount >= 2
  };
}

/* Checkpoint 2 — the second series' two items; Q2 is ONE item that
   resolves only after sections א, ב and ג have all been answered. */
function mastery2State() {
  const results = practiceProgress2.questions.map(q => derivePractice2State(q.number, false));
  return {
    checkpoint: 2, results,
    correctCount: results.filter(r => r === 'correct').length,
    allResolved: results.every(resolvedState),
    passed: results.every(r => r === 'correct')
  };
}

/* One navigation per crossing. Click, Enter/Space on the button and
   ArrowLeft all funnel through advanceScreen(), and a browser keeps
   running script while it tears the page down, so rapid activation
   would otherwise assign window.location.href several times. */
let partCrossingInFlight = false;

function crossToNextPart() {
  if (partCrossingInFlight) return;
  partCrossingInFlight = true;
  PartNav.forward(NEXT_PART);
}

/* Decision point for checkpoint N. Returns true when the learner is
   crossing out of the part, false when they stay (or continue inside). */
/* Platform routing (decision 2026-10-01, guide §6.4 / §6.5 / §8.8 rule 12): the last click
   reports the component 'completed' — pass AND fail — disables the button and stops; Kata routes
   on the statement (recommendedAfterFail on a fail). This supersedes the earlier producer rule
   that kept המשך enabled on a fail. The hop to the next part survives only in a ?dev=1
   walkthrough (xapiEndPart, unit-js/20-xapi.js). */
function crossFromCheckpoint(n) {
  /* Only checkpoint 2 (S10) reaches here; checkpoint 1 is decided in advanceScreen (S5). */
  if (partCrossingInFlight) return true;
  if (!mastery2State().allResolved) return false;
  if (xapiEndPart(lastScreenButton(), NEXT_PART, crossToNextPart)) return true;
  return false;
}

/* Cross-part navigation. init() must run AFTER this part's own initial
   screen setup: if this page was entered via Back, it jumps straight to
   the screen the learner left from. The jump is deferred to a
   microtask inside PartNav for the temporal-dead-zone reason
   documented there. */
PartNav.init({
  part: '03',
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
   xAPI seams — read at call time by unit-js/20-xapi.js, 22-scoring.js and 50-loader.js.
   Ids must match metadata/methodica-ar-math-ratio-03-03.json byte-for-byte (the loader
   checks XAPI_COMP_ID against METADATA.id on every load; _test/verify-static.js checks the map).
   ═══════════════════════════════════════════════════════════ */
const XAPI_COMP_SLUG     = 'methodica-ar-math-ratio-03-03';
const XAPI_COMP_ID       = XAPI_ID_PREFIX + XAPI_COMP_SLUG + '/';
const XAPI_METADATA_FILE = '../metadata/' + XAPI_COMP_SLUG + '.json';
/* Items carrying a code-graded question (each declared question is reported by its screen). */
const XAPI_EVAL_ITEMS    = { '001': 1, '002': 1, '003': 1, '004': 1, '005': 1 };
/* Items scored as ONE question: success only when every section is right (QA F-6, 22-scoring.js).
   derivePracticeState / derivePractice2State: AND over each question's sub-parts. */
const XAPI_ITEM_ALL_CORRECT = { '001': 1, '003': 1, '005': 1 };
/* screen → [item suffix, page within the item]; null = no catalogue item. Exactly TOTAL_SCREENS
   keys, 0-based (keep it a literal object — the static check reads it). */
const SCREEN_TO_SUBCONTENT = {
  0: null,
  1: ['001', 1],
  2: ['001', 2],
  3: ['002', 1],
  4: ['003', 1],
  5: ['003', 2],
  6: ['004', 1],
  7: ['004', 2],
  8: ['005', 1],
  9: ['005', 2],
  10: ['005', 3]
};

/* Drag-board answer text for result.response: each target's card(s), named as the metadata
   names the targets. Read at check time, before any reveal rewrites the board. */
/* The two drag questions as matching pairs (2026-11-11): every card placed on a target is one
   pair, target name [.] card text; the xAPI layer orders them as Kata's pattern
   ('מכונה 1[.]7x[,]מכונה 2[.]3x', 'המשוואה[.]3x + 120 = 7x'). They used to be one string,
   'מכונה 1: 7x, מכונה 2: 3x', which never matched. */
function s8ddqAnswerSpec() {
  const names = { 's8target-m1': 'الآلة 1', 's8target-m2': 'الآلة 2' };
  return { pairs: Object.keys(s8ddqPlacement).filter(d => names[s8ddqPlacement[d]])
    .map(d => ({ source: names[s8ddqPlacement[d]], target: S8DDQ_TEXTS[d] })) };
}
function s9ddqAnswerSpec() {
  return { pairs: Object.keys(s9ddqPlacement).filter(d => s9ddqPlacement[d] === 's9target')
    .map(d => ({ source: 'المعادلة', target: S9DDQ_TEXTS[d] })) };
}

/* Checkpoint 1 failed (fewer than 2 of 3 at S5): the component ENDS here, reported as a fail —
   S6-S10 are never reached, so the platform routes the learner (recommendedAfterFail). */
function endAtCheckpoint1() {
  if (!mastery1State().allResolved) return;
  xapiEndPart(document.getElementById('s5dq-check'), null);
}
/* The component result (2026-11-09, 08.11 sheet fixes 1, 2, 4, 6): exactly S0's promise, "3 שאלות
   … ענו נכון על 2 שאלות ומעלה כדי להתקדם". success and score.scaled come from ONE count — checkpoint 1's
   three questions (mastery1State, each correct only when all its sub-parts are), the count the S5 gate
   reads. The two harder questions (S7–S10, checkpoint 2) report their own answered and never enter.
   Until now: success also needed both harder questions right (S6 promises no such thing), and the
   score was the share of all 9 catalogued sub-questions (partScaled). */
const XAPI_PASS_K = 2, XAPI_PASS_N = 3;
function partGateResult() {
  const ok = mastery1State().correctCount;
  return { success: ok >= XAPI_PASS_K, score: { scaled: ok / XAPI_PASS_N } };
}
/* Two end screens: S10 always, S5 only when checkpoint 1 failed there. */
const PART_END_SCREENS = [5, 10];
function lastScreenButton() {
  if (currentScreen === 10) return document.getElementById('s10viq-check');
  if (currentScreen === 5) { const m1 = mastery1State(); if (m1.allResolved && !m1.passed) return document.getElementById('s5dq-check'); }
  return null;
}

/* ═══════════════════════════════════════════════════════════
   Resume painters — unit-js/45-painters.js reads this registry.
   What each screen saves (vars + the fields that hold state) and how a screen restored from the
   state document is drawn again on the fresh markup: init() is the screen's own ResetInitial,
   paint() replays the template's own terminal look (Mark / Finish / ShowPopup / ApplyView) or the
   retry look, without counting an attempt and without reporting. In-session re-entry still goes
   through resetScreenState(), which now resets only a pristine screen.
   This replaces the sessionStorage snapshot the part used to take when crossing into 03-04.
   practiceProgress / practiceProgress2 are derived (every sync re-derives state + visited from the
   screens' own results); they travel only so no nav is ever drawn from a stale object. The S7 /
   S10 scroll gates are state (guide §8.8 rule 11): GateRead keeps a feedback that was read to the
   end satisfied after a reload. s8ddqDnd / s9ddqDnd are DnD instances, not state: never saved.
   ═══════════════════════════════════════════════════════════ */
const PART_STATE_VARS = ['practiceProgress', 'practiceProgress2'];
const PART_SCREENS = {
  1: {
    vars: ['s1viqAttempts', 's1viqDone', 's1viqResult', 's1viqAnswer', 's1viqView', 's1viqLastWrongSig'],
    inputs: ['s1viq-input-first', 's1viq-input-second'],
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
    vars: ['s2viqAttempts', 's2viqDone', 's2viqResult', 's2viqAnswer', 's2viqView', 's2viqLastWrongSig'],
    inputs: ['s2viq-input'],
    pristine: () => !s2viqAttempts && !s2viqDone && resumeFieldsEmpty(PART_SCREENS[2].inputs),
    init: s2viqResetInitial,
    paint: () => {
      if (s2viqDone) {
        s2viqDone = false;
        if (s2viqResult === 'correct') { s2viqMark('correct'); s2viqFinish(); s2viqShowPopup('correct', 'solution'); }
        else { s2viqFinish(); s2viqView = 'solution'; s2viqApplyView(); }
        return;
      }
      s2viqUpdateBar();
      if (s2viqAttempts && s2viqSignature() === s2viqLastWrongSig) s2viqMark('wrong');
    }
  },
  3: {
    vars: ['s3viqAttempts', 's3viqDone', 's3viqResult', 's3viqAnswer', 's3viqView', 's3viqLastWrongSig'],
    inputs: ['s3viq-input-first', 's3viq-input-second'],
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
    vars: ['s4mcSelected', 's4mcAttempts', 's4mcDone', 's4mcResult', 's4mcLastWrong'],
    pristine: () => !s4mcSelected.size && !s4mcAttempts && !s4mcDone,
    init: s4mcResetInitial,
    paint: () => {
      const done = s4mcDone;
      s4mcDone = false;
      s4mcRender();                                // selection look + aria-checked + the retry gate
      if (done) {
        /* s4mcCheck's terminal marks: the correct set green; on a final wrong the learner's
           wrong picks red as well. */
        s4mcOpts().forEach(o => {
          const id = o.dataset.id;
          o.classList.remove('selected', 'correct', 'wrong');
          if (S4MC.correct.has(id)) o.classList.add('correct');
          else if (s4mcResult !== 'correct' && s4mcSelected.has(id)) o.classList.add('wrong');
        });
        s4mcFinish();
        s4mcShowPopup(s4mcResult);
      } else if (s4mcAttempts && s4mcSig(s4mcSelected) === s4mcLastWrong) {
        /* Non-final wrong, selection unchanged: each submitted option carries its verdict
           instead of 'selected' (s4mcCheck's retry branch). */
        s4mcOpts().forEach(o => {
          o.classList.remove('correct', 'wrong');
          if (!s4mcSelected.has(o.dataset.id)) return;
          o.classList.remove('selected');
          o.classList.add(S4MC.correct.has(o.dataset.id) ? 'correct' : 'wrong');
        });
      }
    }
  },
  5: {
    vars: ['window.lomdaState.s5'],
    pristine: () => { const st = s5dqState(); return !st.attempts && !st.done && st.selection === null; },
    init: s5dqResetInitial,
    paint: () => {
      const st = s5dqState();
      const done = st.done, sel = st.selection;
      st.done = false;
      st.locked = false;
      if (sel !== null) {
        /* s5dqSelect's DOM writes (label, data-empty, aria), without its deferred focus. */
        const b = s5dqBtn(), l = s5dqList(), lab = s5dqSel();
        if (lab) lab.textContent = sel;
        if (b) { b.removeAttribute('data-empty'); b.setAttribute('aria-label', 'المقارنة: ' + sel); }
        if (l) l.querySelectorAll('.dq-option').forEach(o =>
          o.setAttribute('aria-selected', o.dataset.value === sel ? 'true' : 'false'));
      }
      if (done) {
        s5dqLock();
        if (st.answer === null) { s5dqMark('correct'); s5dqFinish('correct'); s5dqShowPopup('correct'); }
        else { s5dqFinish('wrong2'); st.view = 'solution'; s5dqApplyView(); }
        return;                                    // restoreEndedButton() runs after this paint
      }
      s5dqUpdateBar();                             // a non-final wrong marks nothing; the gate holds
    }
  },
  7: {
    vars: ['s7viqAttempts', 's7viqDone', 's7viqResult', 's7viqLastWrong', 's7viqAnswer', 's7viqView',
           's7viqGateOk', 's7viqGateKey', 's7viqGateRead'],
    inputs: ['s7viq-input-amit', 's7viq-input-omer'],
    pristine: () => !s7viqAttempts && !s7viqDone && resumeFieldsEmpty(PART_SCREENS[7].inputs),
    init: s7viqResetInitial,
    paint: () => {
      if (s7viqDone) {
        s7viqDone = false;
        s7viqLock();
        if (s7viqResult === 'correct') {
          S7VIQ.slots.forEach(k => s7viqMark(k, 'correct'));
          s7viqFinish('correct');
          s7viqShowPopup('correct');               // re-arms the scroll gate against GateRead
        } else { s7viqFinish('wrong2'); s7viqView = 'solution'; s7viqApplyView(); }
        return;
      }
      s7viqUpdateBar();
      if (s7viqAttempts && s7viqSig() === s7viqLastWrong) {
        S7VIQ.slots.forEach(k => s7viqMark(k, s7viqInput(k).value.trim() === S7VIQ.correct[k] ? 'correct' : 'wrong'));
      }
    }
  },
  8: {
    vars: ['s8ddqPlacement', 's8ddqChecked', 's8ddqDone', 's8ddqAttempts', 's8ddqResult',
           's8ddqShowFeedback', 's8ddqLastWrong', 's8ddqAnswer', 's8ddqView'],
    pristine: () => !s8ddqAttempts && !s8ddqDone && Object.keys(s8ddqPlacement).every(k => s8ddqPlacement[k] === 'source'),
    init: s8ddqResetInitial,
    paint: () => {
      if (s8ddqDone) {
        s8ddqDone = false;
        if (s8ddqResult === 'correct') {
          s8ddqRender(); s8ddqLock(); s8ddqMarkZones(true);
          s8ddqFinish('correct');
          s8ddqShowPopup('correct');
        } else { s8ddqFinish('wrong2'); s8ddqView = 'solution'; s8ddqApplyView(); }
        return;
      }
      s8ddqRender();                               // board + the retry gate
      /* Per-zone verdicts stay only while the submitted board is untouched (a pick clears them). */
      if (s8ddqShowFeedback && s8ddqSig() === s8ddqLastWrong) s8ddqMarkZones(false);
    }
  },
  9: {
    vars: ['s9ddqPlacement', 's9ddqChecked', 's9ddqDone', 's9ddqAttempts', 's9ddqResult',
           's9ddqShowFeedback', 's9ddqLastWrong', 's9ddqAnswer', 's9ddqView'],
    pristine: () => !s9ddqAttempts && !s9ddqDone && Object.keys(s9ddqPlacement).every(k => s9ddqPlacement[k] === 'source'),
    init: s9ddqResetInitial,
    paint: () => {
      if (s9ddqDone) {
        s9ddqDone = false;
        if (s9ddqResult === 'correct') {
          s9ddqRender(); s9ddqLock(); s9ddqMarkZones(true);
          s9ddqFinish('correct');
          s9ddqShowPopup('correct');
        } else { s9ddqFinish('wrong2'); s9ddqView = 'solution'; s9ddqApplyView(); }
        return;
      }
      s9ddqRender();
      if (s9ddqShowFeedback && s9ddqSig() === s9ddqLastWrong) s9ddqMarkZones(false);
    }
  },
  10: {
    vars: ['s10viqAttempts', 's10viqDone', 's10viqResult', 's10viqLastWrong', 's10viqAnswer', 's10viqView',
           's10viqGateOk', 's10viqGateKey', 's10viqGateRead'],
    inputs: ['s10viq-input-m1', 's10viq-input-m2'],
    pristine: () => !s10viqAttempts && !s10viqDone && resumeFieldsEmpty(PART_SCREENS[10].inputs),
    init: s10viqResetInitial,
    paint: () => {
      if (s10viqDone) {
        s10viqDone = false;
        s10viqLock();
        if (s10viqResult === 'correct') {
          S10VIQ.slots.forEach(k => s10viqMark(k, 'correct'));
          s10viqFinish('correct');
          s10viqShowPopup('correct');
        } else { s10viqFinish('wrong2'); s10viqView = 'solution'; s10viqApplyView(); }
        return;
      }
      s10viqUpdateBar();
      if (s10viqAttempts && s10viqSig() === s10viqLastWrong) {
        S10VIQ.slots.forEach(k => s10viqMark(k, s10viqInput(k).value.trim() === S10VIQ.correct[k] ? 'correct' : 'wrong'));
      }
    }
  }
};
