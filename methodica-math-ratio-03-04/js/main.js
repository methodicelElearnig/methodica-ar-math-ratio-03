'use strict';

/* ═══════════════════════════════════════════════════════════
   js/main.js  —  methodica-math-ratio-03-04
   720 reference engine (canvas 1280×710) + Screen 00
   (TransitionScreen — advanced-practice intro).

   Source: the approved production script
   יעד 1.3_יחס וביטויים אלגבריים_להפקה_06092026.pptx, two levels above
   this part's folder. Slide identity re-derived THIS session from a
   fresh extraction of the .pptx (the file's mtime advances between
   sessions, so nothing was trusted from an earlier extraction):
   presentation.xml's <p:sldIdLst> r:id → ppt/_rels/presentation.xml.rels
   gives order 50 → slide50.xml and order 51 → slide51.xml (1:1 with no
   offset across all 83 slides — derived, not assumed). Presenter notes
   resolved through each slide's OWN .rels: slide51 → rId3 →
   notesSlides/notesSlide46.xml (part-name numbering is NOT slide
   numbering — pptx-script-extraction), whose cached slide-number field
   reads "51".
   Component divider: slide 50 ("רכיב 4", "תרגול מתקדם",
   "methodica-math-ratio-03-04") — production-only, not a screen.

   Screen roster (built so far):
     S0 — TransitionScreen, advanced-practice intro (slide 51).
     S1 — SingleChoiceQuestionImage, advanced-practice question 1,
          with the scroll-gated long-feedback popup (slide 52).
     S2 — DropdownQuestion (single-dropdown variant), advanced-practice
          question 2, with the View My Answer Toggle and the same
          scroll gate (slide 53).

   S0 ported from the same-lomda sibling methodica-math-ratio-03-03's
   own S0 — the identical screen role (practice-stage kickoff
   transition, slide 39 in the same deck) — which took it from -02's own
   S0 (slide 29), per 720-templates → references/TransitionScreen.md +
   references/_global-components.md.
   Engine plumbing (scaleApp / goTo / resetScreenState / keyboard nav /
   setNavLabel / dev bridge) ported from methodica-math-ratio-03-02,
   byte-identical apart from this part's own screen roster, so every
   part of this lomda behaves the same. Cross-part Back uses this part's
   own js/partnav.js (PartNav).
   ═══════════════════════════════════════════════════════════ */

/* ─── Constants ─────────────────────────────────────────────
   THE PART IS NOW COMPLETE: S0-S11 cover slides 51-62, and slide 63 is
   the deck's own component divider for part 05. goTo() bounds-checks
   against this, so S11 is the last index it will activate — and that is
   now correct rather than provisional, because S11's Continue does NOT
   go through goTo() at all. It crosses the part boundary through
   crossToNextPart() / PartNav.forward(NEXT_PART), which is what
   720-templates → _global-components.md → "Advance gates" requires: a
   missing cross-part redirect looks exactly like working code, so the
   real crossing is wired rather than left to a silent no-op. */
const TOTAL_SCREENS = 12;  // S0-S11 — the whole part (slides 51-62)


/* ─── S0 companion-character assets ─────────────────────────
   Explicit ID → asset mapping. The supplied filenames
   ("Character_1_Telescope.mp4" / "Character_2_Binoculars.mp4") do not
   follow one convention — they do not even name the same prop — and
   resolving a character by filename or colour guessing is forbidden
   regardless (720-templates → _global-components.md → Companion
   character system). The mapping below was verified by playing each
   clip and identifying the character in it, not read off the names.

   LOGICAL MAPPING confirmed against methodica-math-ratio-03-01's own
   CHARACTER_ASSETS (same lomda, same two companions, and the same
   pair pictured on slide 51's own placeholder still, image4.png):
       character-1 = the ORANGE/CORAL round one-eyed character
       character-2 = the TALL YELLOW character with the purple headband

   Per-instance asset map (TransitionScreen.md → "Companion-character
   asset resolution": each TransitionScreen instance keeps its own map,
   not a lomda-wide one) — this is this part's only TransitionScreen so
   far.

   PRESENTER NOTE MATCH — "הדמות עם משקפת, בפנים מופתעות". Both clips
   were supplied by the producer for this screen during this session
   and were verified in real Edge under file:// — metadata read off the
   elements (BOTH 1280×720, 4.01s, matching this lomda's other clip
   pairs) and frames sampled at 1280×720 and inspected: clip 1 shows
   the orange character holding a spyglass with a surprised round
   mouth, clip 2 the yellow character holding binoculars with the same
   surprised mouth. "משקפת" covers both props; each matches its own
   character's build.

   NO `fit` ENTRY — measured, not assumed. TransitionScreen.md → "C"
   requires a shim to be derived from measured ink bounds of THIS clip
   pair, never by analogy. Measured against each clip's own #FEFEFE
   backdrop (threshold 40/765 to reject codec noise), as a fraction of
   the 720px frame height:

       clip                    frame 0            final frame (3.9s)
       Character_1_Telescope   top .054 bot .938  top .047 bot .971
                               ink height .885    ink height .924
       Character_2_Binoculars  top .008 bot .978  top .008 bot .978
                               ink height .969    ink height .969

   Two reasons this ships with no shim:
     1. The clip holds its FINAL frame (nothing loops), and there the
        ground lines agree to within 0.7% of frame height (.971 vs
        .978) and the ink-height gap is 4.7% — near the ~3% variance
        band both of the template's references confirm, and far from
        the 8%+ gaps they cite as visibly noticeable.
     2. The residual gap runs in the same direction as the two
        companions' GENUINE build difference (character-1 is the short
        round one, character-2 the tall one). Scaling character-1 up
        to match would equalise a real height difference, which this
        project does not do — the approved sibling
        methodica-math-ratio-03-01 anchors the same two companions by
        measured ink and deliberately preserves a much larger gap
        (standing-PNG ink heights .734 vs .947).
   The learner only ever sees one of the two, so there is no
   side-by-side comparison for a mismatch to show up in. */
const S0_CHARACTER = {
  'character-1': {
    video: 'assets/video/Character_1_Telescope.mp4',
    alt: 'الشخصية البرتقالية المستديرة ذات العين الواحدة، تنظر بمنظار بتعبير مندهش'
  },
  'character-2': {
    video: 'assets/video/Character_2_Binoculars.mp4',
    alt: 'الشخصية الصفراء الطويلة ذات عصبة الرأس البنفسجية، تنظر بمنظار بتعبير مندهش'
  }
};
/* S7 — slide 58's TransitionScreen. Its own clip pair, measured on its
   own evidence; TransitionScreen.md → "C" forbids deriving a fit shim
   by analogy to another pair, so S0's decision is NOT carried over and
   this pair was measured from scratch.

   MEASURED INK, union over all 96 frames of each corrected clip, in the
   clip's own 1280×720 canvas (never the canvas box itself):
       character-1  solid ink 825 × 633  — 87.9% of canvas height,
                    baseline (feet) y=666, centre x=651
       character-2  solid ink 588 × 667  — 92.6% of canvas height,
                    baseline (feet) y=686, centre x=614
   The two differ by 4.7 points of ink height. That is the companions'
   GENUINE build difference — character-1 is the short round one,
   character-2 the tall one — and it is deliberately NOT equalised
   (720-design-guidelines: anchor by measured ink, never scale one
   companion to match the other's media canvas).

   NO FIT SHIM is applied. The rule that would call for one — anchor a
   character-adjacent element to the character's own ink — has no
   subject here: slide 58 authors no speech bubble and no other
   character-adjacent element (its only shapes are the headline text
   box, the character picture, the report icon and the two nav
   buttons). The residual 10px on-screen baseline difference between
   the two clips is never visible, because the learner only ever sees
   the one companion they selected. */
const S7_CHARACTER = {
  'character-1': {
    video: 'assets/video/Character_1_paint_palette.mp4',
    alt: 'الشخصية البرتقالية المستديرة ذات العين الواحدة، بشارب فنان وقبعة فنان ولوحة ألوان في يدها'
  },
  'character-2': {
    video: 'assets/video/Character_2_paint_palette.mp4',
    alt: 'الشخصية الصفراء الطويلة، بشارب فنان وقبعة فنان ولوحة ألوان في يدها'
  }
};

/* S11 — slide 62's TransitionScreen, the part's closing beat. Its own
   clip pair, measured on its own evidence; TransitionScreen.md → "C"
   forbids deriving a fit shim by analogy to another pair, so neither
   S0's nor S7's decision is carried over.

   THE SLIDE'S STILL IS NOT USED. Slide 62 draws image69.png — the
   deck's embed of the supplied still קשקוש.png (removed from the tree
   2026-09-23, in git) — and that art
   pictures CHARACTER-2 specifically. Shipping it would show every
   learner the yellow companion whatever they chose in part 01, which
   is the exact failure the companion-character system prevents. The
   producer ruled this same case on 2026-09-11 for
   methodica-math-ratio-03-01's #s19 (an instructional-designer still
   marking where a clip was needed), and the matching pair was supplied
   into this part's own assets/video/. Both clips show the same scene
   as the still: the companion in a beret, palette and brush in hand,
   beside an easel bearing the coloured scribble the note names,
   "קשקוש".

   BOTH CLIPS SHIP BACKGROUND-CORRECTED, like this part's other pair.
   As supplied they carried a NEAR-WHITE plate — limited-range Y 233/234
   against the 235 white point, decoding to RGB #FDFDFD/#FEFEFE — which
   renders as a visible grey rectangle on this screen's #fff ground. The
   producer has already ruled that plate not approved, and
   Character_1_paint_palette.mp4 / Character_2_paint_palette.mp4 were
   corrected on the same grounds (their supplied files are archived in
   assets/video/_pre-correction/, which is non-shipping material kept
   until the part is finally approved). The supplied painting clips are
   archived there too, as Character_N_painting.ORIGINAL.mp4.
   The correction is a Y-only highlight ramp — never CSS, which cannot
   reach inside the frame — recovered by measuring the approved pair
   against its own archived original: identity below Y 217, then a
   ×1.5 ramp clamped at the limited-range white point 235. Re-encoded
   h264_nvenc, yuv420p, audio copied, matching the approved pair's own
   encode. Validated three ways:
     · plate:    Y 233/234 → Y 235 on every sampled frame, and the four
                 corners of the decoded frame now read 255,255,255 in
                 the live page (they read 253/254 before).
     · control:  the SAME encode with NO filter leaves the plate at
                 233/234 — so the lift is the ramp's, not the encoder's.
     · midtones: below the knee the corrected clip differs from the
                 source by a mean of 2.264, against the encoder-only
                 control's 2.269 — i.e. the ramp changes nothing there
                 and all of that delta is ordinary re-encode noise.
   The ink boxes below were re-measured AFTER the correction and are
   unchanged by it (the companion's solid ink lies far below the knee).

   MEASURED INK of the HELD FINAL FRAME, solid ink only (threshold
   90/765, which rejects the soft ground shadow), in each clip's own
   1280×720 canvas. Each frame holds a companion AND an easel separated
   by a clear column band, so the COMPANION's own box is measured
   apart from the scene — it is what the bubble points at:
       character-1  companion x 128..666  y  52..656   (easel 748..1092)
       character-2  companion x 309..565  y  27..674   (easel 652..972)
   Companion ink heights 605 and 648 of 720 — 84.0% and 90.0% of frame
   height. NO FIT SHIM: that 6-point gap runs in the direction of the
   two companions' GENUINE build difference (character-1 short and
   round, character-2 tall), the same finding S0 and S7 record for
   their own pairs, and this project does not equalise a real height
   difference. NO freezeAt either — both clips were sampled at 3.98s
   and hold a meaningful final frame, so the natural end of playback is
   the correct end state.

   'ink' is this screen's own addition to the map shape, published to
   CSS on every entry (see s11ApplyCharacter). The bubble is computed
   from it rather than hand-placed per character, because the two clips
   frame the scene very differently and one fixed bubble position
   cannot serve both — the mechanism 03-01's #s19 established. Values
   are fractions of the clip's own frame: l/t/b of the companion. */
const S11_CHARACTER = {
  'character-1': {
    video: 'assets/video/Character_1_painting.mp4',
    alt: 'الشخصية البرتقالية المستديرة ذات العين الواحدة، بشارب وقبعة فنان، تحمل لوحة ألوان وفرشاة بجانب حامل رسم وعليه خربشة ملوّنة',
    ink: { l: 0.100000, t: 0.072222, b: 0.912500 }
  },
  'character-2': {
    video: 'assets/video/Character_2_painting.mp4',
    alt: 'الشخصية الصفراء الطويلة، بشارب وقبعة فنان، تحمل لوحة ألوان وفرشاة بجانب حامل رسم وعليه خربشة ملوّنة',
    ink: { l: 0.241406, t: 0.037500, b: 0.937500 }
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
   hold a meaningful final frame, so no `freezeAt` is needed either. */

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
  /* Per-character fit shim. ALWAYS written, including the no-op values,
     so switching characters can never leave a previous one's scale
     behind on the element (neither entry in this part carries `fit` —
     see the note on S0_CHARACTER above). */
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


/* S7 — slide 58's TransitionScreen. Same pure mount as S0: no state
   machine, nothing to resolve, and the same re-entry-safe applier, so
   returning to the screen never restarts an already-correct clip. The
   clip plays once and holds its final frame; it is never looped. */
function s7ApplyCharacter() {
  applyCompanionCharacter(S7_CHARACTER, 's7-char-video');
}

/* S11 — slide 62's TransitionScreen, the part's last screen. The same
   pure mount as S0 and S7, plus ONE extra step this instance needs: the
   resolved clip's measured companion ink is published to CSS so the
   speech bubble's tail lands on the companion actually on screen. The
   two clips frame the scene differently, so a bubble placed in fixed
   pixels would point at the right companion in one clip and at empty
   backdrop in the other.
   Written unconditionally on every entry, never only once, so a
   selection changed in an earlier part is picked up — and so the
   previous character's ink can never persist on the element. */
function s11ApplyCharacter() {
  applyCompanionCharacter(S11_CHARACTER, 's11-char-video');
  const wrap = document.querySelector('#s11 .s1-video-wrap');
  if (!wrap) return;
  const char   = window.lomdaState.selectedCharacter || DEFAULT_CHARACTER;
  const assets = S11_CHARACTER[char] || S11_CHARACTER[DEFAULT_CHARACTER];
  const ink    = (assets && assets.ink) || {};
  /* The CSS defaults are character-2's, so a missing value falls back
     to a correct composition rather than to zero. */
  if (ink.l != null) wrap.style.setProperty('--ink-l', ink.l);
  if (ink.t != null) wrap.style.setProperty('--ink-t', ink.t);
  if (ink.b != null) wrap.style.setProperty('--ink-b', ink.b);
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
  // The applet-zoom modal is NOT a feedback popup and deliberately does
  // not match the id sweep above, so it is closed explicitly — an
  // enlarged applet must never stay promoted over the next screen.
  closeAppletZoomOnNavigate();
  closeImageZoomOnNavigate();
  const prev = document.querySelector('.screen.active');
  if (prev) {
    prev.querySelectorAll('video').forEach(v => { try { v.pause(); } catch (e) {} });
    /* Clear a reveal that is still mid-flight, so leaving a screen
       early can never strand one of its groups at opacity 0. */
    if (window.Motion) Motion.cancel(prev);
    /* S10's own sequence is cancelled the same way, so leaving it
       mid-flight strands nothing. */
    if (prev.id === 's10') s10Cancel();
    prev.classList.remove('active');
  }
  currentScreen = n;
  const next = document.getElementById('s' + n);
  if (next) next.classList.add('active');
  /* Straight after '.active', per the engine's wiring contract. Only
     screens carrying [data-reveal] animate; every other screen is
     untouched. The engine's canonical policy applies unchanged: the
     staged reveal plays on a screen's FIRST entry in a session, and a
     revisit, Back, or DEV jump shows everything instantly in its
     settled state. Note this call sits in goTo() and NOWHERE else —
     scaleApp() never calls it, so a resize or a responsive reflow
     cannot replay a reveal either. */
  if (window.Motion) Motion.reveal(next);
  /* S10 carries no [data-reveal] groups — its authored sequence is its
     own, and runs here on the same first-visit-per-session rule. */
  if (n === 10) s10Reveal();
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
       feedback on return to a resolved screen"). s1scDone is the only
       flag involved — there is deliberately no second "popup was open
       on exit" flag. Reopening REDISPLAYS the feedback without asking
       for its reading requirement a second time: the scroll gate
       remembers the state it has already been read in (producer
       contract, 2026-09-15), and only a new question state re-arms
       it. */
    if (!s1scDone) { if (screenPristine(1)) s1scResetInitial(); }
    else s1scShowPopup(s1scResult);
    syncPracticeProgressNav(document.getElementById('s1'));
  }
  if (n === 2) {
    /* Same resume/reopen contract as S1. DropdownQuestion adds the
       View My Answer Toggle, so a resolved final-wrong screen always
       comes back on the SOLUTION view (DropdownQuestion.md → "On
       re-entry"), never on whichever view the learner left on. */
    if (!s2dqDone) { if (screenPristine(2)) s2dqResetInitial(); }
    else if (s2dqAnswer !== null) { s2dqView = 'solution'; s2dqApplyView(); }
    else s2dqShowPopup('correct');
    syncPracticeProgressNav(document.getElementById('s2'));
  }
  if (n === 3) {
    /* Same resume/reopen contract as S1 and S2. ValueInputQuestion's
       final-wrong reveal is destructive (it overwrites the very inputs
       the learner typed into), so a resolved final-wrong screen always
       comes back on the SOLUTION view, never on whichever view the
       learner left on (ValueInputQuestion.md → "{prefix}ResetInitial()").
       The nav is re-synced on EVERY entry, not just once, so returning
       to this question after resolving a later one still shows the
       right state (_global-components.md → Progress Question →
       Behavioral rules). */
    if (!s3viqDone) { if (screenPristine(3)) s3viqResetInitial(); }
    else if (s3viqAnswer !== null) { s3viqView = 'solution'; s3viqApplyView(); }
    else s3viqShowPopup('correct');
    syncPracticeProgressNav(document.getElementById('s3'));
  }
  if (n === 4) {
    /* Same resume/reopen contract as S3 — this is the same template
       and the same state machine, re-prefixed. */
    if (!s4viqDone) { if (screenPristine(4)) s4viqResetInitial(); }
    else if (s4viqAnswer !== null) { s4viqView = 'solution'; s4viqApplyView(); }
    else s4viqShowPopup('correct');
    syncPracticeProgressNav(document.getElementById("s4"));
  }
  if (n === 5) {
    /* Same resume/reopen contract as S3 and S4 — the same template,
       reduced to a single value field. */
    if (!s5viqDone) { if (screenPristine(5)) s5viqResetInitial(); }
    else if (s5viqAnswer !== null) { s5viqView = "solution"; s5viqApplyView(); }
    else s5viqShowPopup("correct");
    syncPracticeProgressNav(document.getElementById("s5"));
  }
  if (n === 8) {
    /* Resume-state lock, the same contract every question screen in this
       part carries: a resolved question never hands out a fresh attempt
       on return. This variant has no popup to reopen, so re-entry simply
       re-paints the resolved marking over the locked options. */
    if (!s8scDone) { if (screenPristine(8)) s8scResetInitial(); }
    else s8scApply();
  }
  if (n === 7) {
    /* TransitionScreen: re-resolve the companion on EVERY entry, so a
       selection changed in an earlier part is picked up. Back is a
       normal in-part button here (S6 is its predecessor), so there is
       no PartNav.mountBack() call — that belongs to screen 0 alone. */
    s7ApplyCharacter();
  }
  if (n === 11) {
    /* Same pure mount as S0 and S7, and for the same reason: the
       companion must re-resolve on every entry, not once. This one also
       republishes the clip's measured ink, which is what keeps the
       speech bubble's tail on the companion after a character switch.
       Back is a normal in-part button here (S10 is its predecessor). */
    s11ApplyCharacter();
  }
  if (n === 6) {
    /* Same resume/reopen contract as the other question screens.
       MultipleChoiceQuestion has no View My Answer Toggle, so a
       resolved screen simply reopens the feedback of its outcome over
       the locked pills. */
    if (!s6mcDone) { if (screenPristine(6)) s6mcResetInitial(); }
    else s6mcShowPopup(s6mcResult);
    syncPracticeProgressNav(document.getElementById('s6'));
  }
}

/* ─── Forward navigation ─────────────────────────────────────
   The gate lives HERE, not only in the Continue button's handler: the
   engine binds ArrowLeft directly to advanceScreen(), so a gate that
   sat only in an onclick would be silently bypassable by keyboard
   (720-templates → _global-components.md → "Advance gates").
   S0 has NO gate: a TransitionScreen carries no interaction to resolve
   (720-templates → TransitionScreen.md → "A. Structural contract").
   S1 gates on RESOLVED, never on correctness — a final wrong attempt
   still resolves the question and still lets the learner move on.
   S1 ALSO honours the scroll gate here, not only through the button's
   disabled state. The scroll-gate spec says the progression button is
   the only CONTROL the gate touches, which is about not disabling a
   close button — it is not licence to leave the keyboard path open:
   the engine binds ArrowLeft straight to advanceScreen(), so a gate
   that lived only in `button.disabled` would let a learner skip the
   feedback they have not read with one key press, which is precisely
   the silent bypass _global-components.md → "Advance gates" exists to
   prevent. Both input methods must resolve the same way. */
/* ─── End of part: the crossing into part 05 ──────────────────
   DESTINATION. Part 04 is slides 51-62 and part 05 is slides 64-72;
   the deck's own component-divider slides are what say so (slide 50
   "רכיב 4 / תרגול מתקדם / methodica-math-ratio-03-04", slide 63
   "רכיב 5 / שאלת שיא א / methodica-math-ratio-03-05"), not the part
   numbers. Slide 62 is therefore this part's last screen, and its
   "متابعة" is a part boundary rather than an in-part step.

   UNCONDITIONAL. Slide 62 is a TransitionScreen: it authors no
   checkpoint, no branch and no pass condition, and nothing between
   slides 51 and 62 authors one either — this part has no mastery gate
   at all (unlike methodica-math-ratio-03-03, whose own deck authors two
   checkpoints). So every learner who reaches S11 crosses, exactly as
   the approved 03-01 → 03-02 handoff does.

   MECHANISM. PartNav.forward() (js/partnav.js), the project's existing
   cross-part contract, byte-identical to the copies in 03-01/02/03. It
   pushes this part and the screen being left onto the sessionStorage
   trail before navigating, which is what gives 03-05's S0 a Back button
   returning HERE, to S11, rather than to a hardcoded predecessor. No ad
   hoc window.location assignment is used. The selected character needs
   no special handling: it lives in localStorage under the lomda-wide
   CHARACTER_STORAGE_KEY that 03-05 reads.

   RE-CROSSING. If the learner already crossed and came back,
   PartNav.forwardOverride() reports the destination recorded for their
   journey and that destination is REPLAYED rather than re-derived
   (_global-components.md → "Re-crossing a previously-resolved boundary
   forward must replay the recorded destination"). The trail lives in
   sessionStorage, which a learner can edit, so a replayed id that is
   not this part's own destination is refused rather than followed.

   NO OUTCOME IS REPORTED HERE. The Ministry reporting/integration layer
   is owned outside lomda production: no SCORM, no xAPI, no cmi.*, no
   completion postMessage. This is internal routing only.
   ─────────────────────────────────────────────────────────── */
const NEXT_PART = '05';

/* One navigation per crossing. Click, Enter/Space on the button and
   ArrowLeft all funnel through advanceScreen(), and a browser keeps
   running script while it tears the page down, so rapid activation
   would otherwise assign window.location.href several times and push
   several trail entries. */
let partCrossingInFlight = false;

/* Platform routing (decision 2026-10-01, guide §6.4 / §6.5 / §8.8 rule 12): the last click
   reports the component 'completed' — pass AND fail — disables the button and stops; Kata routes
   on the statement (recommendedAfterFail on a fail). This supersedes the earlier producer rule
   that kept המשך enabled on a fail. The hop to the next part survives only in a ?dev=1
   walkthrough (xapiEndPart, unit-js/20-xapi.js). */
function crossToNextPart() {
  if (partCrossingInFlight) return;
  if (xapiEndPart(lastScreenButton(), NEXT_PART)) partCrossingInFlight = true;
}

function advanceScreen() {
  if (currentScreen === 1 && (!s1scDone || !s1scGateOk)) return;
  if (currentScreen === 2 && (!s2dqDone || !s2dqGateOk)) return;
  if (currentScreen === 3 && (!s3viqDone || !s3viqGateOk)) return;
  if (currentScreen === 4 && (!s4viqDone || !s4viqGateOk)) return;
  if (currentScreen === 5 && (!s5viqDone || !s5viqGateOk)) return;
  if (currentScreen === 6 && (!s6mcDone || !s6mcGateOk)) return;
  /* S7 is a TransitionScreen and carries no gate.
     S8 gates on ANSWERED, never on correct — slide 59's own note:
     "להתנות את ההמשך רק אחרי לחיצה את אחת התשובות". There is no scroll
     gate to honour here because this variant has no feedback popup. */
  if (currentScreen === 8 && !s8scDone) return;
  /* S9, S10 and S11 are ungated too. S11 is the END OF THE PART: its
     Continue crosses the boundary instead of calling goTo(), which
     TOTAL_SCREENS would otherwise swallow into a silent no-op. */
  if (currentScreen === 11) { crossToNextPart(); return; }
  goTo(currentScreen + 1);
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
  /* A pointer-down that starts inside the scrollable body must NOT
     begin a popup drag — otherwise the learner cannot scroll a
     scroll-gated popup by dragging its scrollbar or swiping its text
     (720-templates → _global-components.md → "Long-feedback variant"
     rule 11). The header and the rest of the popup chrome stay draggable. */
  if (e.target && e.target.closest && e.target.closest('.scq-popup-scroll')) return;
  /* An action button is a control, not a drag handle. */
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
  /* The scrolling hint rides along with the popup and needs its own
     canvas-fit check on every move — the drag clamp above keeps the
     POPUP inside the canvas, not the bubble hanging off its side. */
  placeScrollHint(popupDragId);
});
window.addEventListener('pointerup',     () => { popupDragId = null; });
window.addEventListener('pointercancel', () => { popupDragId = null; });

/* Feedback popup action button (View My Answer Toggle). `action` falsy
   → the row is removed. SingleChoiceQuestion never carries one (it
   shows the chosen and the correct option side by side instead), but
   the helper stays part of the shared popup shape every part carries. */
function renderPopupAction(popupId, action, onClick) {
  const popup = document.getElementById(popupId);
  if (!popup) return;
  let row = popup.querySelector('.scq-popup-actions');
  if (!action) { if (row) row.remove(); return; }
  /* The row lives INSIDE the scroll column, as the last thing after the
     feedback body — not as a sibling pinned to the popup's bottom edge.
     Producer decision 2026-09-10: on a scroll-gated popup a pinned row
     reads as a floating control hovering over the text, and it offers
     the toggle before the learner has read the explanation. Placed in
     the flow it simply arrives at the end of the scroll, which is also
     where the scroll gate is satisfied. On an ungated popup the scroll
     column is a plain flex column, so the row still renders directly
     under the body exactly as before. */
  const host = popup.querySelector('.scq-popup-scroll') || popup;
  if (!row) {
    row = document.createElement('div');
    row.className = 'scq-popup-actions';
    host.appendChild(row);
  } else if (row.parentElement !== host) {
    host.appendChild(row);
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

   // px — see below


/* ═══════════════════════════════════════════════════════════
   S1 — SingleChoiceQuestionImage (slide 52)
   Advanced-practice question 1. Spec: 720-templates →
   SingleChoiceQuestion.md (+ its SingleChoiceQuestionImage section) +
   _question-template-defaults.md. The interaction is this template's
   own single-select pill contract — no View My Answer Toggle (SCQ
   shows the chosen and the correct option side by side already, which
   the toggle's own applicability rule excludes).

   THREE options, o1..o3 in the deck's own top→bottom authored order
   (group y 297 / 372 / 448 — an even ~75px rhythm with no gap large
   enough to hide a fourth):
       o1 = כן             (correct)
       o2 = לא
       o3 = אי אפשר לדעת
   CORRECT = 'o1', on the deck's own two independent markers: o1's
   capsule is the ONLY one with a GREEN border (#00B050 vs the others'
   #00B0F0), and its radio ellipse is the only SOLID-FILLED green one.
   Never resolved from a letter key. See index.html's own
   source-evidence block above #s1 for the full inventory, the ratio
   bidi finding, and the two flagged source content issues.

   SHARED EXPLANATION, from the slide. Both feedback shapes (id=20
   correct, id=22 incorrect) carry the SAME explanation — the same
   sentences in the same order, differing only in authoring markup
   (id=20 uses plain runs and spaced ratios, id=22 uses OMML for
   "BD = x"/"CD = 3x" and unspaced ratios) and in their title. Rendered
   they are identical, which is the deck's own normal 2-attempt shape
   (_question-template-defaults.md → "Feedback: shared explanation"), so
   one body is used for both outcomes.

   Every ratio below is written in the physical order the PowerPoint
   render produces, as an isolated LTR unit — never as a raw "3 : 1"
   source string left to the browser's bidi to flip. The empty source
   paragraphs are preserved as real one-line gaps exactly where they
   sit in the source, per "Paragraph structure maps to line-height".
   ═══════════════════════════════════════════════════════════ */
const S1SC_GAP = '';   // rendered as <p class="scq-gap"> — an authored blank line

const S1SC_EXPLAIN = [
  'النسبة بين مساحة المثلث‎ ADB ‎ومساحة المثلث‎ ACD ‎هي أيضًا <bdi class="scq-expr" dir="ltr">1 : 3‏‎</bdi>.',
  S1SC_GAP,
  '<bdi class="scq-expr" dir="ltr">BD = x</bdi> والنسبة بين‎ BD ‎و-CD هي <bdi class="scq-expr" dir="ltr">1 : 3</bdi> .',
  'لذلك يمكننا أن نرمز <bdi class="scq-expr" dir="ltr">CD = 3x</bdi> .',
  S1SC_GAP,
  'نلاحظ أن المثلثين‎ ACD ‎و-ADB ‎ لهما نفس الارتفاع (AE).',
  S1SC_GAP,
  'المثلث‎ ABD ‎مثلث منفرج الزاوية، و-AE هو الارتفاع على الضلع DB.',
  S1SC_GAP,
  'طول ارتفاع المثلث‎ ACD ‎هو 5 سم. لذلك مساحة المثلث‎ ACD ‎هي‎:',
  '@math:<span class="math-expr"><span class="frac"><span class="frac-num">3x &middot; 5</span><span class="frac-den">2</span></span> = <span class="frac"><span class="frac-num">15x</span><span class="frac-den">2</span></span> = 7.5x</span>',
  S1SC_GAP,
  'طول ارتفاع المثلث‎ ABD ‎هو 5 سم‎.',
  'ولذلك مساحة المثلث‎ ABD ‎هي‎:',
  /* PRODUCER-APPROVED CORRECTION (2026-09-09), NOT a fidelity slip —
     do not "restore" the source string in the name of PPT fidelity.
     Both of slide 52's feedback shapes author this numerator as 5h.
     Triangle ADB has base BD = x and height AE = 5, so its area is
     x · 5 / 2 = 2.5x: the numerator is 5x, and `h` is this diagram's
     own label for that same height (AE), which makes "5h" read as
     5 × 5. The producer reviewed the flag raised on the previous
     build and approved the one-character fix, so it ships as 5x here
     and identically in the terminal-incorrect state (both outcomes
     render this one shared array). */
  '@math:<span class="math-expr"><span class="frac"><span class="frac-num">5x</span><span class="frac-den">2</span></span> = 2.5x</span>',
  S1SC_GAP,
  'إذًا النسبة بين مساحة المثلث‎ ABD ‎ومساحة المثلث‎ ACD ‎هي <bdi class="scq-expr" dir="ltr">2.5x : 7.5x</bdi>',
  'نختزل بـ <bdi class="scq-expr" dir="ltr">2.5x</bdi> الذي لا يساوي 0، ونحصل على <bdi class="scq-expr" dir="ltr">1:3‏‎</bdi>.'
];

const S1SC = { correctId: 'o1', maxAttempts: 2 };

/* `scrollGate` is set ONLY on the two long outcomes. The retry popup is
   two short lines and keeps the standard auto-height, ungated popup —
   the variant is a per-feedback-state decision, never a per-template
   default (_global-components.md → "Long-feedback variant").
   `titleHtml` carries the canonical second line the template contract
   appends to a terminal-incorrect title in the solution view
   (_question-template-defaults.md → "Terminal-incorrect title carries
   a fixed second line"): a real line break, both lines bold. */
const S1SC_POPUP = {
  retry:   { bg: '#ffdbdc', title: 'الإجابة غير صحيحة.', body: ['هل نحاول مرة أخرى؟'] },
  correct: { bg: '#edf8ed', title: 'صحيح‎!', body: S1SC_EXPLAIN, scrollGate: true },
  wrong2:  {
    bg: '#ffdbdc',
    title: 'هذا خطأ. أخطأتم، لا تقلقوا، فالفهم يتمّ تدريجيًا‎.',
    titleSecondLine: 'الإجابة الصحيحة تَظهَر الآن.',
    body: S1SC_EXPLAIN,
    scrollGate: true
  }
};

let s1scSelected  = null;
let s1scAttempts  = 0;
let s1scDone      = false;
let s1scResult    = null;   // 'correct' | 'wrong2'
let s1scLastWrong = null;   // the option id just marked wrong — the retry gate compares against THIS
/* SCROLL-GATE MEMORY (retrofit, 2026-09-15). The gate is per FEEDBACK
   STATE, not per opening. `GateKey` names the state currently on
   screen; `GateRead` names the state whose content has already been
   read to the end. Re-showing a state that is already in `GateRead`
   leaves the gate satisfied, so closing and reopening the same feedback
   never asks the learner to read it twice. Showing a DIFFERENT state
   re-arms the gate by simple inequality, and s1scResetInitial() clears
   the memory because a fresh question must be read again. */
let s1scGateKey = null;
let s1scGateRead = null;
let s1scGateOk    = true;   // false only while a scroll-gated popup has not been read to the end

function s1scOpt(id) { return document.querySelector('#s1 .scq-opt[data-id="' + id + '"]'); }
function s1scOpts()  { return document.querySelectorAll('#s1 .scq-opt'); }

/* Single select; clears a prior wrong attempt's marks; re-enables the
   check button only if the newly selected id differs from the one just
   marked wrong (_question-template-defaults.md → "Retry gate" —
   comparison shape for this template: the selected option id). */
function s1scSelect(id) {
  if (s1scDone) return;
  s1scOpts().forEach(o => {
    o.classList.remove('selected');
    o.setAttribute('aria-checked', 'false');
  });
  const opt = s1scOpt(id);
  opt.classList.add('selected');
  opt.setAttribute('aria-checked', 'true');
  s1scSelected = id;

  if (s1scAttempts > 0) {
    document.getElementById('s1sc-popup').classList.add('hidden');
    s1scOpts().forEach(o => o.classList.remove('correct', 'wrong'));
  }
  const chk = document.getElementById('s1sc-check');
  if (chk) chk.disabled = (id === s1scLastWrong);
}

function s1scCheck() {
  if (s1scDone) { advanceScreen(); return; }
  if (!s1scSelected) return;

  s1scAttempts++;
  const correct = s1scSelected === S1SC.correctId;
  reportQ(['001', 'q1'], correct, correct || s1scAttempts >= S1SC.maxAttempts, xapiChoiceAnswer('001', 'q1', '#s1 .scq-opt', s1scSelected));

  if (correct) {
    s1scOpt(s1scSelected).classList.add('correct');
    s1scResult = 'correct';
    s1scFinish();
    s1scShowPopup('correct');
    return;
  }

  if (s1scAttempts < S1SC.maxAttempts) {
    /* Non-final wrong: mark the chosen option wrong; the correct one is
       NOT revealed yet (SingleChoiceQuestion.md → state 4, "Try Again"). */
    s1scOpt(s1scSelected).classList.add('wrong');
    s1scLastWrong = s1scSelected;
    document.getElementById('s1sc-check').disabled = true;
    s1scShowPopup('retry');
    return;
  }

  /* Final wrong: chosen option red + ✗, correct option green + ✓, both
     shown at once — which is why this template has no View My Answer
     Toggle. */
  s1scOpt(s1scSelected).classList.add('wrong');
  s1scOpt(S1SC.correctId).classList.add('correct');
  s1scResult = 'wrong2';
  s1scFinish();
  s1scShowPopup('wrong2');
}

function s1scFinish() {
  s1scDone = true;
  s1scOpts().forEach(o => { o.disabled = true; });
  const chk = document.getElementById('s1sc-check');
  /* SingleChoiceQuestion's own completion label is "هل نتابع؟" — NOT
     ValueInputQuestion's "متابعة"; the two templates do not share one,
     and slide 52 authors no override of it. s1scShowPopup() may
     immediately re-disable this button while the scroll gate is
     unsatisfied; that is the gate's one and only target. */
  if (chk) { setNavLabel(chk, 'هل نتابع؟'); chk.disabled = false; }
  hideHintButton('s1sc-hint');
  /* Question 1 is resolved by this screen alone, so the
     progress model is re-derived here too. This screen carries no
     progress bar of its own (its slide authors none), so nothing is
     rendered — only the shared state is brought up to date. */
  syncPracticeProgressNav(document.getElementById('s1'));

}

function s1scResetInitial() {
  s1scSelected  = null;
  s1scAttempts  = 0;
  s1scDone      = false;
  s1scResult    = null;
  s1scLastWrong = null;
  s1scGateOk    = true;
  s1scGateKey   = null;
  s1scGateRead  = null;
  s1scOpts().forEach(o => {
    o.classList.remove('selected', 'correct', 'wrong');
    o.setAttribute('aria-checked', 'false');
    o.disabled = false;
  });
  document.getElementById('s1sc-popup').classList.add('hidden');
  document.getElementById('s1sc-hint-overlay').classList.add('hidden');
  const chk = document.getElementById('s1sc-check');
  if (chk) { setNavLabel(chk, 'هل إجابتي صحيحة؟'); chk.disabled = true; }
  showHintButton('s1sc-hint');   // visible from the start — see the source-evidence note in index.html
}

function s1scShowPopup(type) {
  const popup = document.getElementById('s1sc-popup');
  if (!popup) return;
  const cfg  = S1SC_POPUP[type];
  const body = document.getElementById('s1sc-popup-body');
  const scroller = document.getElementById('s1sc-popup-scroll');
  const chk  = document.getElementById('s1sc-check');

  popup.style.background = cfg.bg;
  if (popup.classList.contains('hidden')) resetPopupPosition(popup);

  /* innerHTML, not textContent: a terminal-incorrect title is two bold
     lines separated by a REAL break, which plain text cannot carry. */
  const titleEl = document.getElementById('s1sc-popup-title');
  titleEl.innerHTML = cfg.titleSecondLine
    ? (cfg.title + '<br>' + cfg.titleSecondLine)
    : cfg.title;

  body.innerHTML = cfg.body.map(p => {
    if (p === '')                return '<p class="scq-gap" aria-hidden="true"></p>';
    if (p.indexOf('@math:') === 0) return '<p class="scq-math">' + p.slice(6) + '</p>';
    return '<p>' + p + '</p>';
  }).join('');
  renderPopupAction('s1sc-popup', null, null);   // no action row on this template

  /* Scroll gate — arm before showing, then reset and measure AFTER the
     popup is visible. Both of those have to happen after the unhide:
     a display:none element ignores a scrollTop write and reports
     scrollHeight/clientHeight as 0. Resetting before the unhide is a
     silent no-op that leaves the previous scroll position in place, so
     re-entering a resolved screen reopened the feedback already
     scrolled to the end and the gate immediately re-satisfied itself. */
  popup.classList.toggle('is-scroll', !!cfg.scrollGate);
  s1scGateKey = cfg.scrollGate ? (type) : null;
  /* Already read in THIS state → stay satisfied across a close/reopen. */
  s1scGateOk = !cfg.scrollGate || s1scGateRead === s1scGateKey;
  if (chk && !s1scGateOk) chk.disabled = true;
  /* The SCROLL CONTAINER is the title+body column, not the body
     alone — see the .scq-popup.is-scroll block in css/style.css. */
  if (cfg.scrollGate) scroller.setAttribute('tabindex', '0');   // the only way a keyboard user can satisfy the gate
  else                scroller.removeAttribute('tabindex');

  popup.classList.remove('hidden');
  clampPopup(popup);
  scroller.scrollTop = 0;
  if (cfg.scrollGate) {
    placeScrollHint('s1sc-popup');
    s1scCheckScrollGate();      // unlocks immediately when the content does not actually overflow
  } else {
    popup.classList.remove('needs-scroll');
  }
}

/* Bound from the popup body's own onscroll, and called once right after
   the popup is shown. Once satisfied the gate stays satisfied for that
   feedback STATE — not merely for that opening. Scrolling back up never
   re-locks it, and neither does closing the feedback and reopening the
   same state; `GateRead` is what carries that across (producer
   contract, 2026-09-15). Only a state that has not been read to the end
   arms it. */
function s1scCheckScrollGate() {
  if (s1scGateOk) {
    /* Already satisfied — make sure the hint bubble is not left over
       from a previous, still-locked presentation of this popup. */
    const p0 = document.getElementById('s1sc-popup');
    if (p0) p0.classList.remove('needs-scroll');
    return;
  }
  const scroller = document.getElementById('s1sc-popup-scroll');
  const popup = document.getElementById('s1sc-popup');
  if (!scroller || !popup) return;
  if (!atBottom(scroller)) {
    popup.classList.add('needs-scroll');   // shows the hint bubble
    return;
  }
  s1scGateOk = true;
  s1scGateRead = s1scGateKey;      // this state has now been read to the end
  const chk = document.getElementById('s1sc-check');
  if (chk && s1scDone) chk.disabled = false;
  popup.classList.remove('needs-scroll');  // hides the hint bubble
}

/* A resize can turn overflowing content into fitting content (or the
   reverse). A satisfied gate is never re-armed by a resize. */
let s1scResizeRaf = null;
window.addEventListener('resize', () => {
  const popup = document.getElementById('s1sc-popup');
  if (!popup || popup.classList.contains('hidden')) return;
  placeScrollHint('s1sc-popup');
  if (s1scGateOk) return;
  if (s1scResizeRaf) cancelAnimationFrame(s1scResizeRaf);
  s1scResizeRaf = requestAnimationFrame(s1scCheckScrollGate);
});

function s1scOpenHint() {
  xapiRequestedHint('001', 'q1');
  document.getElementById('s1sc-hint-overlay').classList.remove('hidden');
}
function s1scCloseHint() {
  document.getElementById('s1sc-hint-overlay').classList.add('hidden');
}
function s1scCloseHintOnBackdrop(e) {
  if (e.target === e.currentTarget) s1scCloseHint();
}

/* Enter/Space on a focused option selects it — radiogroup keyboard a11y.
   Delegated so it covers every .scq-opt without per-node handlers. */
document.addEventListener('keydown', e => {
  const isEnter = e.key === 'Enter' || e.code === 'Enter' || e.code === 'NumpadEnter';
  const isSpace = e.key === ' ' || e.key === 'Spacebar' || e.code === 'Space';
  if (!isEnter && !isSpace) return;
  const opt = e.target.closest && e.target.closest('#s1 .scq-opt');
  if (!opt || opt.disabled) return;
  e.preventDefault();
  s1scSelect(opt.dataset.id);
});

/* ═══════════════════════════════════════════════════════════
   S2 — DropdownQuestion, single-dropdown variant (slide 53)
   Advanced-practice question 2. Spec: 720-templates →
   DropdownQuestion.md + _question-template-defaults.md +
   _global-components.md (Feedback popup system, View My Answer Toggle,
   Long-feedback scroll gate).

   CORRECT VALUE = 'هي 1:1', from the deck's own marker: that option's
   run carries a green text highlight (srgbClr 00FF00) and the other
   two do not. Never resolved from position or a letter key. It is also
   the mathematically correct answer, which this screen's own feedback
   body derives: BD = DC (given 1 : 1) and both triangles share the
   height AE, so their areas are equal.

   SHARED EXPLANATION — slide 53's two feedback shapes (id=26 correct,
   id=16 incorrect) carry byte-identical bodies: the same twelve
   paragraphs, the same four empty group breaks, the same two OMML
   fractions. Only the title differs. That is the documented
   shared-explanation shape (_question-template-defaults.md), so one
   body renders both outcomes.

   Unlike slide 52, the two equations are authored INLINE at the end of
   their own sentence rather than as standalone centred lines, and are
   reproduced that way. Every ratio here is "1 : 1" — symmetric, so it
   is immune to the bidi term-swap slide 52 exhibited — but each is
   still an isolated no-wrap LTR unit so no line break can land inside
   one.
   ═══════════════════════════════════════════════════════════ */
const S2DQ_FRAC_XH =
  '<span class="math-expr"><span class="frac"><span class="frac-num">x &middot; h</span>' +
  '<span class="frac-den">2</span></span></span>';

const S2DQ_EXPLAIN = [
  'النسبة بين القطعة‎ BD ‎والقطعة‎ DC ‎هي <bdi class="scq-expr" dir="ltr">1 : 1</bdi> ولذلك <bdi class="scq-expr" dir="ltr">BD = DC</bdi>.',
  'نرمز إلى كل واحد منهما بـ <bdi class="scq-expr" dir="ltr">x</bdi>.',
  '',
  'AE هو ارتفاع للمثلثات ADB، ABC و-ACD.',
  '',
  'في المثلث منفرج الزاوية، يقع ارتفاعان من الارتفاعات الثلاثة خارج المثلث‎.',
  '',
  'نعبّر عن مساحة المثلث‎ ADB ‎على النحو التالي: ' + S2DQ_FRAC_XH,
  '',
  'نعبّر عن مساحة المثلث ACD ‎على النحو التالي: ' + S2DQ_FRAC_XH,
  '',
  'المساحتان متساويتان، ولذلك فإن النسبة بينهما هي <bdi class="scq-expr" dir="ltr">1 : 1‏‎</bdi>.'
];

const S2DQ = { correctValue: 'هي 1:1', maxAttempts: 2 };

/* `scrollGate` on the two long outcomes only — the note asks for it and
   both source panels are 438×690, taller than the canvas. `action` /
   `answerView` implement the View My Answer Toggle, which THIS template
   uses (SingleChoiceQuestion's exclusion does not apply: a dropdown has
   one slot whose single display value is either the learner's pick or
   the solution, never both at once). The toggle's own labels and copy
   are canonical product UI text, not producer content. */
const S2DQ_POPUP = {
  retry:   { bg: '#ffdbdc', title: 'الإجابة غير صحيحة.', body: ['هل نحاول مرة أخرى؟'] },
  correct: { bg: '#edf8ed', title: 'صحيح جدًا‎!', body: S2DQ_EXPLAIN, scrollGate: true },
  wrong2:  {
    bg: '#ffdbdc',
    title: 'أخطأتم. لا بأس، فكل خطأ يقوّي قدراتكم‎.',
    titleSecondLine: 'الإجابة الصحيحة تَظهَر الآن.',
    body: S2DQ_EXPLAIN,
    scrollGate: true,
    action: { label: 'الانتقال إلى إجابتي', view: 'answer' },
    answerView: {
      title: 'الإجابة غير صحيحة.',
      /* NULL, not omitted. The canonical "الإجابة الصحيحة تَظهَر الآن." second
         line belongs to the SOLUTION view only — the answer view shows
         the learner's own answer, not the solution, and keeps its own
         fixed copy (_question-template-defaults.md → "Terminal-incorrect
         title carries a fixed second line" → "Does NOT apply to ... the
         answerView"). Because the view is built by merging this object
         onto the base outcome, the key has to be explicitly cleared or
         the base's second line leaks into it. */
      titleSecondLine: null,
      body:  ['يمكن مشاهدة الحل الصحيح مرة أخرى‎.'],
      action: { label: 'الإجابة الصحيحة', view: 'solution' },
      scrollGate: false          /* two short lines — never gated */
    }
  }
};

let s2dqSelected  = null;
let s2dqAttempts  = 0;
let s2dqDone      = false;
let s2dqResult    = null;   // 'correct' | 'wrong2'
let s2dqLastWrong = null;   // the value just marked wrong — the retry gate compares against THIS
let s2dqAnswer    = null;   // the learner's own final pick, snapshotted BEFORE the reveal overwrites it
let s2dqView      = 'solution';
/* SCROLL-GATE MEMORY (retrofit, 2026-09-15). The gate is per FEEDBACK
   STATE, not per opening. `GateKey` names the state currently on
   screen; `GateRead` names the state whose content has already been
   read to the end. Re-showing a state that is already in `GateRead`
   leaves the gate satisfied, so closing and reopening the same feedback
   never asks the learner to read it twice. Showing a DIFFERENT state
   re-arms the gate by simple inequality, and s2dqResetInitial() clears
   the memory because a fresh question must be read again. */
let s2dqGateKey = null;
let s2dqGateRead = null;
let s2dqGateOk    = true;

/* Verdict icons — the EXPORTED Figma vectors from the canonical
   Dropdown Menu states (120:2136 correct, 120:2125 Incorrect), each a
   16x16 viewBox with the component-s own Feedback/Correct/300 and
   Feedback/Incorrect/300 fills. Not redrawn, and not the 14x12 / 14x14
   radio glyphs this project uses elsewhere. */
const S2DQ_ICON_CORRECT = '<svg width="16" height="16" viewBox="0 0 16 16" fill="none" aria-hidden="true" xmlns="http://www.w3.org/2000/svg"><path d="M6.52036 15C6.12296 15.0002 5.74183 14.8126 5.46106 14.4787L2.25845 10.6774C1.91385 10.2681 1.91385 9.60467 2.25845 9.19536C2.60317 8.78619 3.16194 8.78619 3.50666 9.19536L6.52036 12.7737L14.4933 3.30688C14.8381 2.89771 15.3968 2.89771 15.7415 3.30688C16.0862 3.71618 16.0862 4.37966 15.7415 4.78896L7.57966 14.4787C7.29889 14.8126 6.91776 15.0002 6.52036 15Z" fill="#609E12"/></svg>';
const S2DQ_ICON_WRONG   = '<svg width="16" height="16" viewBox="0 0 16 16" fill="none" aria-hidden="true" xmlns="http://www.w3.org/2000/svg"><path d="M15.6743 2.32569C15.4657 2.11715 15.1828 2 14.8879 2C14.5929 2 14.31 2.11715 14.1014 2.32569L9 7.42709L3.8986 2.32569C3.69 2.11715 3.40711 2 3.11215 2C2.81718 2 2.53429 2.11715 2.32569 2.32569C2.11715 2.53429 2 2.81718 2 3.11215C2 3.40711 2.11715 3.69 2.32569 3.8986L7.42709 9L2.32569 14.1014C2.11715 14.31 2 14.5929 2 14.8879C2 15.1828 2.11715 15.4657 2.32569 15.6743C2.53429 15.8829 2.81718 16 3.11215 16C3.40711 16 3.69 15.8829 3.8986 15.6743L9 10.5729L14.1014 15.6743C14.31 15.8829 14.5929 16 14.8879 16C15.1828 16 15.4657 15.8829 15.6743 15.6743C15.8829 15.4657 16 15.1828 16 14.8879C16 14.5929 15.8829 14.31 15.6743 14.1014L10.5729 9L15.6743 3.8986C15.8829 3.69 16 3.40711 16 3.11215C16 2.81718 15.8829 2.53429 15.6743 2.32569Z" fill="#B20010"/></svg>';

function s2dqBtn()  { return document.getElementById('s2dq-btn'); }
function s2dqList() { return document.getElementById('s2dq-list'); }
function s2dqOptions() { return Array.from(s2dqList().querySelectorAll('.dq-option')); }

/* Keyboard model: focus never leaves the field. The list is a listbox
   whose current option is published through aria-activedescendant, and
   the options are not tab stops (tabindex="-1") — the standard
   collapsible-listbox pattern, and the only one that keeps a single
   predictable tab stop inside an RTL sentence. */
let s2dqActive = -1;

function s2dqSetActive(i) {
  const opts = s2dqOptions();
  s2dqActive = i;
  opts.forEach((o, n) => o.classList.toggle('is-active', n === i));
  const btn = s2dqBtn();
  if (i >= 0) {
    btn.setAttribute('aria-activedescendant', opts[i].id);
    opts[i].scrollIntoView({ block: 'nearest' });
  } else {
    btn.removeAttribute('aria-activedescendant');
  }
}

function s2dqCloseList(refocus) {
  const l = s2dqList(); if (!l) return;
  l.classList.add('hidden');
  const btn = s2dqBtn();
  btn.setAttribute('aria-expanded', 'false');
  s2dqSetActive(-1);
  if (refocus) btn.focus();
}

function s2dqOpenList() {
  const btn = s2dqBtn(), list = s2dqList();
  if (!btn || btn.disabled) return;
  /* Only one list open at a time on a screen (DropdownQuestion.md →
     Interaction rules). This part has one slot; the sweep keeps the
     contract true if a second is ever added. */
  document.querySelectorAll('.dq-list').forEach(l => l.classList.add('hidden'));
  document.querySelectorAll('.dq-btn').forEach(b => b.setAttribute('aria-expanded', 'false'));
  list.classList.remove('hidden');
  btn.setAttribute('aria-expanded', 'true');
  /* Start on the current selection if there is one, else the first. */
  const opts = s2dqOptions();
  const cur = opts.findIndex(o => o.dataset.value === s2dqSelected);
  s2dqSetActive(cur >= 0 ? cur : 0);
  /* Opening the dropdown while a non-final retry popup is showing is a
     corrective action and dismisses it (_global-components.md →
     "First-attempt retry feedback dismissal"). Never fires once done:
     the field is disabled by then. */
  if (s2dqAttempts > 0 && !s2dqDone) {
    document.getElementById('s2dq-popup').classList.add('hidden');
  }
}

function s2dqToggle() {
  const btn = s2dqBtn();
  if (!btn || btn.disabled) return;
  if (btn.getAttribute('aria-expanded') === 'true') s2dqCloseList(false);
  else s2dqOpenList();
}

/* All keyboard interaction for the field. Enter/Space are handled here
   and default-prevented so the button's own implicit click cannot also
   fire and immediately re-toggle the list. */
function s2dqKeydown(e) {
  const btn = s2dqBtn();
  if (!btn || btn.disabled) return;
  const open = btn.getAttribute('aria-expanded') === 'true';
  const opts = s2dqOptions();
  const isEnter = e.key === 'Enter';
  const isSpace = e.key === ' ' || e.key === 'Spacebar' || e.code === 'Space';

  if (e.key === 'Escape') {
    if (open) { e.preventDefault(); s2dqCloseList(true); }
    return;
  }
  if (e.key === 'ArrowDown' || e.key === 'ArrowUp') {
    e.preventDefault();
    e.stopPropagation();               // never let it reach the engine's screen nav
    if (!open) { s2dqOpenList(); return; }
    const step = (e.key === 'ArrowDown') ? 1 : -1;
    const next = (s2dqActive + step + opts.length) % opts.length;
    s2dqSetActive(next);
    return;
  }
  if (e.key === 'Home' || e.key === 'End') {
    if (!open) return;
    e.preventDefault();
    s2dqSetActive(e.key === 'Home' ? 0 : opts.length - 1);
    return;
  }
  if (isEnter || isSpace) {
    e.preventDefault();
    if (!open) { s2dqOpenList(); return; }
    if (s2dqActive >= 0) s2dqSelect(opts[s2dqActive].dataset.value, true);
    return;
  }
}

function s2dqSelect(value, refocus) {
  const btn = s2dqBtn();
  if (!btn || btn.disabled) return;
  s2dqSelected = value;
  const sel = document.getElementById('s2dq-selected');
  sel.textContent = value;
  btn.dataset.empty = 'false';
  s2dqOptions().forEach(o =>
    o.setAttribute('aria-selected', String(o.dataset.value === value)));
  s2dqCloseList(!!refocus);
  /* Retry gate — comparison shape for a single slot is the selected
     value itself (_question-template-defaults.md → "Retry gate"). */
  const chk = document.getElementById('s2dq-check');
  if (chk) chk.disabled = (value === s2dqLastWrong);
}

/* One shared document-level listener: a click outside every dropdown
   wrapper closes any open list. */
document.addEventListener('click', e => {
  if (e.target.closest && e.target.closest('.dq-dropdown-wrap')) return;
  document.querySelectorAll('.dq-list').forEach(l => l.classList.add('hidden'));
  document.querySelectorAll('.dq-btn').forEach(b => {
    b.setAttribute('aria-expanded', 'false');
    b.removeAttribute('aria-activedescendant');
  });
  document.querySelectorAll('.dq-option.is-active').forEach(o => o.classList.remove('is-active'));
});

/* Render the single slot in one of three states. The verdict icon's box
   is reserved in every state by CSS (visibility, not display), so this
   never changes the control's geometry. */
function s2dqMark(state) {
  const btn = s2dqBtn();
  const sel = document.getElementById('s2dq-selected');
  const ico = document.getElementById('s2dq-verdict');
  btn.classList.remove('dq-correct', 'dq-wrong');
  ico.innerHTML = '';
  if (state === 'answer') {
    sel.textContent = s2dqAnswer;            // the learner's own (wrong) pick
    btn.classList.add('dq-wrong');
    ico.innerHTML = S2DQ_ICON_WRONG;
  } else {
    if (state === 'solution') sel.textContent = S2DQ.correctValue;
    btn.classList.add('dq-correct');
    ico.innerHTML = S2DQ_ICON_CORRECT;
  }
}

/* Single owner of the resolved-wrong state for the slot and the popup. */
function s2dqApplyView() {
  if (s2dqAnswer === null) return;
  s2dqMark(s2dqView === 'answer' ? 'answer' : 'solution');
  s2dqShowPopup('wrong2', s2dqView);
}
/* A pure view switch: never routes through s2dqCheck(), never touches
   attempts/done/scoring. */
function s2dqToggleView() {
  if (s2dqAnswer === null) return;
  s2dqView = (s2dqView === 'solution') ? 'answer' : 'solution';
  s2dqApplyView();
}

function s2dqCheck() {
  if (s2dqDone) { advanceScreen(); return; }
  if (!s2dqSelected) return;

  s2dqAttempts++;
  const correct = s2dqSelected === S2DQ.correctValue;
  /* xAPI (2026-11-11): the layer writes Kata's form from the metadata — matching, the one sentence [.] the pick */
  reportQ(['002', 'q1'], correct, correct || s2dqAttempts >= S2DQ.maxAttempts,
    { pairs: [{ source: 0, target: s2dqSelected }] });

  if (correct) {
    s2dqResult = 'correct';
    s2dqAnswer = null;                       // nothing to toggle to
    s2dqFinish();
    s2dqMark('correct');
    s2dqShowPopup('correct');
    return;
  }

  if (s2dqAttempts < S2DQ.maxAttempts) {
    /* Non-final wrong: NO per-slot marking (DropdownQuestion.md → reuse
       QA row 5), the slot stays selectable, and the check button
       disables until a genuinely different value is chosen. */
    s2dqLastWrong = s2dqSelected;
    document.getElementById('s2dq-check').disabled = true;
    s2dqShowPopup('retry');
    return;
  }

  /* Final wrong: snapshot the learner's own pick BEFORE the reveal
     overwrites the slot's displayed value — this is the only moment it
     still exists. */
  s2dqAnswer = s2dqSelected;
  s2dqResult = 'wrong2';
  s2dqView   = 'solution';
  s2dqFinish();
  s2dqApplyView();
}

function s2dqFinish() {
  s2dqDone = true;
  s2dqCloseList(false);
  const btn = s2dqBtn();
  if (btn) btn.disabled = true;             // a locked slot never reopens
  const chk = document.getElementById('s2dq-check');
  /* DropdownQuestion's own canonical completion label. */
  if (chk) { setNavLabel(chk, 'هل نتابع؟'); chk.disabled = false; }
  hideHintButton('s2dq-hint');
  /* Question 2 is resolved by this screen alone, so the
     progress model is re-derived here too. This screen carries no
     progress bar of its own (its slide authors none), so nothing is
     rendered — only the shared state is brought up to date. */
  syncPracticeProgressNav(document.getElementById('s2'));

}

function s2dqResetInitial() {
  s2dqSelected  = null;
  s2dqAttempts  = 0;
  s2dqDone      = false;
  s2dqResult    = null;
  s2dqLastWrong = null;
  s2dqAnswer    = null;
  s2dqView      = 'solution';
  s2dqGateOk    = true;
  s2dqGateKey   = null;
  s2dqGateRead  = null;
  const btn = s2dqBtn();
  if (btn) {
    btn.disabled = false;
    btn.dataset.empty = 'true';
    btn.classList.remove('dq-correct', 'dq-wrong');
    btn.setAttribute('aria-expanded', 'false');
  }
  document.getElementById('s2dq-verdict').innerHTML = '';
  document.getElementById('s2dq-selected').textContent = 'اختاروا إجابة';
  s2dqList().classList.add('hidden');
  s2dqOptions().forEach(o => { o.setAttribute('aria-selected', 'false'); o.classList.remove('is-active'); });
  s2dqSetActive(-1);
  document.getElementById('s2dq-popup').classList.add('hidden');
  document.getElementById('s2dq-hint-overlay').classList.add('hidden');
  const chk = document.getElementById('s2dq-check');
  if (chk) { setNavLabel(chk, 'هل إجابتي صحيحة؟'); chk.disabled = true; }
  showHintButton('s2dq-hint');
}

/* `view` merges the config's answerView override onto the base outcome,
   exactly as _global-components.md → "View My Answer Toggle" specifies. */
function s2dqShowPopup(type, view) {
  const popup = document.getElementById('s2dq-popup');
  if (!popup) return;
  const base = S2DQ_POPUP[type];
  const cfg  = (view === 'answer' && base.answerView)
    ? Object.assign({}, base, base.answerView)
    : base;
  const body     = document.getElementById('s2dq-popup-body');
  const scroller = document.getElementById('s2dq-popup-scroll');
  const chk      = document.getElementById('s2dq-check');

  popup.style.background = cfg.bg;
  if (popup.classList.contains('hidden')) resetPopupPosition(popup);

  const titleEl = document.getElementById('s2dq-popup-title');
  titleEl.innerHTML = cfg.titleSecondLine
    ? (cfg.title + '<br>' + cfg.titleSecondLine)
    : cfg.title;

  body.innerHTML = cfg.body.map(p => {
    if (p === '') return '<p class="scq-gap" aria-hidden="true"></p>';
    return '<p>' + p + '</p>';
  }).join('');
  renderPopupAction('s2dq-popup', cfg.action, s2dqToggleView);

  /* Scroll gate — same contract and the same ordering fix as S1: the
     reset and the measurement must both run AFTER the popup is visible,
     because a display:none element ignores a scrollTop write and
     reports zero for scrollHeight/clientHeight. */
  popup.classList.toggle('is-scroll', !!cfg.scrollGate);
  s2dqGateKey = cfg.scrollGate ? (type + '|' + (view || 'solution')) : null;
  /* Already read in THIS state → stay satisfied across a close/reopen. */
  s2dqGateOk = !cfg.scrollGate || s2dqGateRead === s2dqGateKey;
  if (chk && !s2dqGateOk) chk.disabled = true;
  if (cfg.scrollGate) scroller.setAttribute('tabindex', '0');
  else                scroller.removeAttribute('tabindex');

  popup.classList.remove('hidden');
  clampPopup(popup);
  scroller.scrollTop = 0;
  if (cfg.scrollGate) {
    placeScrollHint('s2dq-popup');
    s2dqCheckScrollGate();
  } else {
    popup.classList.remove('needs-scroll');
    if (chk && s2dqDone) chk.disabled = false;
  }
}

function s2dqCheckScrollGate() {
  if (s2dqGateOk) {
    /* Already satisfied — make sure the hint bubble is not left over
       from a previous, still-locked presentation of this popup. */
    const p0 = document.getElementById('s2dq-popup');
    if (p0) p0.classList.remove('needs-scroll');
    return;
  }
  const scroller = document.getElementById('s2dq-popup-scroll');
  const popup    = document.getElementById('s2dq-popup');
  if (!scroller || !popup) return;
  if (!atBottom(scroller)) { popup.classList.add('needs-scroll'); return; }
  s2dqGateOk = true;
  s2dqGateRead = s2dqGateKey;      // this state has now been read to the end
  const chk = document.getElementById('s2dq-check');
  if (chk && s2dqDone) chk.disabled = false;
  popup.classList.remove('needs-scroll');
}

let s2dqResizeRaf = null;
window.addEventListener('resize', () => {
  const popup = document.getElementById('s2dq-popup');
  if (!popup || popup.classList.contains('hidden')) return;
  placeScrollHint('s2dq-popup');
  if (s2dqGateOk) return;
  if (s2dqResizeRaf) cancelAnimationFrame(s2dqResizeRaf);
  s2dqResizeRaf = requestAnimationFrame(s2dqCheckScrollGate);
});

function s2dqOpenHint()  { xapiRequestedHint('002', 'q1'); document.getElementById('s2dq-hint-overlay').classList.remove('hidden'); }
function s2dqCloseHint() { document.getElementById('s2dq-hint-overlay').classList.add('hidden'); }
function s2dqCloseHintOnBackdrop(e) { if (e.target === e.currentTarget) s2dqCloseHint(); }

/* ═══════════════════════════════════════════════════════════
   PROGRESS + SCORING MODEL — FOUR questions; Question 3 owns S3, S4
   and S5 as its sections א / ב / ג
   720-templates → _global-components.md → "Progress Question"
   (display) + _question-template-defaults.md → "Pass threshold formula
   (bundle scoring)" (scoring).

   ── AUTHORITY ──
   The producer corrected the presenter notes of slides 52-57 in the
   source deck on 2026-09-14 and supplied the mapping below. Everything
   here was re-derived from a FRESH extraction of that corrected deck
   (presentation order resolved through presentation.xml's <p:sldIdLst>
   → ppt/_rels/presentation.xml.rels; each slide's notes resolved
   through its OWN ppt/slides/_rels/slideNN.xml.rels; each note's
   cached slide-number field checked against its position).

       slide 52  →  screen 1 (s1sc,  SingleChoiceQuestion)  = QUESTION 1
       slide 53  →  screen 2 (s2dq,  DropdownQuestion)      = QUESTION 2
       slide 54  →  screen 3 (s3viq, ValueInputQuestion)  ┐
       slide 55  →  screen 4 (s4viq, ValueInputQuestion)  ├ QUESTION 3
       slide 56  →  screen 5 (s5viq, ValueInputQuestion)  ┘ sections א/ב/ג
       slide 57  →  screen 6 (s6mc,  MultipleChoiceQuestion) = QUESTION 4

   Slide numbers are NOT screen numbers: slide 51 is the transition
   screen s0, so every question screen sits one index below its slide.
   The mapping above was verified against the build itself, not against
   these comments — each screen's template matches its slide's
   "שם המסך בפיגמה" note, and each ValueInput screen's answer key
   matches its slide's "الإجابة الصحيحة" note (54 → 1 : 4, 55 → 1 : 2,
   56 → 16).

   ── EVIDENCE, FROM THE CORRECTED NOTES ──
     · 52 "שאלה 1 מתוך 4"  · 53 "שאלה 2 מתוך 4"
     · 54, 55 and 56 ALL carry the same "שאלה 3 מתוך 4" — one question,
       three screens
     · 57 "שאלה 4 מתוך 4"
     · 51 (transition) and 58+ (PopupPage) carry no marker at all.
   Supporting, from the slides themselves: 54/55/56 share one stem
   lettered א. → ב. → ג. ("עזרו לנוי לעצב את ההזמנה."), repeat the same
   surprise-party scenario and the same 2 : 1 ratio, and section ג opens
   by recalling what section ב established. No other slide in 51-62
   carries a lettered stem.

   ── WHAT A DOT IS ──
   Every dot is ONE COMPLETE AUTHORED QUESTION, never a section:
     · Questions 1, 2 and 4 each occupy one screen and resolve as soon
       as that screen is completed;
     · Question 3 occupies three screens. Its dot stays `current` after
       section א and after section ב — whether those sections were
       answered correctly or not — and resolves ONLY after section ג.
   This is the shared component's Basic/Standard row
   (_global-components.md → "What one progress item represents"): "one
   progress item = one question in the series", "never resolve early —
   while the learner is still working through a multi-sub-part
   question, its item stays visually unresolved (current), even when
   the outcome is already determined mathematically". THIS IS NOT A
   PEAK QUESTION; the Peak row (item = one lettered sub-part, verdict
   shown at once) does not apply to this sequence.

   ── SCORING ──
   Question 3 is correct only if a && b && c. Any incorrect section
   makes it incorrect, and that verdict appears only after section ג.
   No partial credit, and no 2-of-3 rule inside the question — the
   sections' results are stored separately for diagnostics and for each
   screen's own restoration, never as scored units of their own
   (_question-template-defaults.md → "What counts as a scored unit":
   a question with lettered sub-parts "is still ONE scored unit, not
   one unit per sub-part").

   Nothing is stored: every question's state is derived on each sync
   from the screens' own `sNDone` / `sNResult` flags, so a second copy
   cannot drift. `resolved` and `scoredCorrect` stay separate fields
   (_question-template-defaults.md → "Outcome semantics") — reaching a
   terminal state is not being scored correct, and the final-wrong
   reveal does not convert failure into success.

   NO PART-LEVEL VERDICT IS COMPUTED HERE. `ceil(2N/3)` scores a bundle
   and would need a producer decision about the bundle's membership and
   about what a failed threshold does; no transitionScreen in this part
   authors a "succeed in at least K of N" condition, so none is
   invented. The four question results are published for whoever
   assembles that later.

   ── WHERE THE BAR RENDERS ──
   On EVERY question screen of the series: s1, s2, s3, s4, s5 and s6.
   The same shared component, byte-identical markup in all six, synced
   by the same syncPracticeProgressNav() on entry to each.

   Only slides 54-56 carry the bar's raster mock (image22.png, id=14);
   slides 52, 53 and 57 carry no progress-bar shape or image at all.
   That absence is a gap in the deck's mock-ups, not authored intent —
   the producer confirmed on 2026-09-14 that those slides are Questions
   1, 2 and 4 of the same four-question series, and ruled that the live
   component must appear consistently across it. The PowerPoint raster
   is never used; it is only evidence about geometry
   (_global-components.md → Progress Question → "the nav lives inside
   each question screen's own markup … every screen in the series
   repeats the same markup", which this now satisfies).
   ═══════════════════════════════════════════════════════════ */

/* Question 3's three sections: display order, authored letter, the
   screen each occupies, and the LOCAL result of that screen. Each
   resolver returns null (unanswered) | 'correct' | 'incorrect'. These
   are stored separately and are NOT progress items. */
const Q3_SECTIONS = [
  { letter: 'א', screen: 3, resolved: () => (s3viqDone ? (s3viqResult === 'correct' ? 'correct' : 'incorrect') : null) },
  { letter: 'ב', screen: 4, resolved: () => (s4viqDone ? (s4viqResult === 'correct' ? 'correct' : 'incorrect') : null) },
  { letter: 'ג', screen: 5, resolved: () => (s5viqDone ? (s5viqResult === 'correct' ? 'correct' : 'incorrect') : null) }
];

/* Question 3's own verdict: null until section ג has been completed,
   then correct only if a && b && c. */
function q3Resolved() {
  const r = Q3_SECTIONS.map(s => s.resolved());
  if (r.some(x => x === null)) return null;          // ג not completed yet
  return r.every(x => x === 'correct') ? 'correct' : 'incorrect';
}

/* The four authored questions, in authored order. `screens` lists every
   data-screen index the question spans; `resolved()` returns null while
   the question is still open. A single-screen question resolves the
   moment its screen is completed. */
const PRACTICE_QUESTIONS = [
  { number: 1, slides: '52', screens: [1],
    resolved: () => (s1scDone ? (s1scResult === 'correct' ? 'correct' : 'incorrect') : null) },
  { number: 2, slides: '53', screens: [2],
    resolved: () => (s2dqDone ? (s2dqResult === 'correct' ? 'correct' : 'incorrect') : null) },
  { number: 3, slides: '54-56', screens: [3, 4, 5], sections: Q3_SECTIONS,
    resolved: q3Resolved },
  { number: 4, slides: '57', screens: [6],
    resolved: () => (s6mcDone ? (s6mcResult === 'correct' ? 'correct' : 'incorrect') : null) }
];

/* One entry per authored QUESTION, in the shared component's own field
   shape. Four entries, four dots, four authored labels. */
const practiceProgress = {
  questions: PRACTICE_QUESTIONS.map(q => ({
    number: q.number, visited: false, state: 'not-answered', screens: q.screens.slice()
  })),
  results: []   /* filled by every sync — see deriveQuestionResult() */
};

/* A question's DISPLAY state.
   Resolved → its verdict. Otherwise `current` while the learner is on
   one of its screens, or while it is part-done (Question 3 after
   section א or ב — the case this whole model exists for). Anything
   else is a question not yet begun. */
function deriveQuestionState(q) {
  const verdict = q.resolved();
  if (verdict) return verdict;
  const onIt    = q.screens.indexOf(currentScreen) !== -1;
  const started = !!q.sections && q.sections.some(s => s.resolved() !== null);
  return (onIt || started) ? 'current' : 'not-answered';
}

/* One question as a scored unit, derived — never stored. Section
   fields are present only on the question that has sections, and they
   are diagnostics: `sectionsCorrect` is never a score, and two correct
   sections with one incorrect section is ONE INCORRECT QUESTION. */
function deriveQuestionResult(q) {
  const verdict = q.resolved();
  const out = {
    id: 'ratio-03-04-q' + q.number,
    number: q.number,
    slides: q.slides,
    screens: q.screens.slice(),
    state: deriveQuestionState(q),
    resolved: verdict !== null,
    scoredCorrect: verdict === 'correct'
  };
  if (q.sections) {
    const r = q.sections.map(s => s.resolved());
    out.sectionCount    = q.sections.length;
    out.sectionLetters  = q.sections.map(s => s.letter);
    out.sectionScreens  = q.sections.map(s => s.screen);
    out.sectionResults  = r;                                   /* stored separately */
    out.sectionsResolved = r.filter(x => x !== null).length;
    out.sectionsCorrect  = r.filter(x => x === 'correct').length;  /* diagnostic, NOT a score */
  }
  return out;
}

/* The documented generic renderer, generic over question count.
   Written once per part, never once per screen. */
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
    const st = state.questions[n - 1].state;
    /* Connector "visited" is NOT item "visited": it lights up only once
       the whole QUESTION on its right-hand side is resolved, not merely
       seen — so it stays unlit while Question 3 runs through its
       sections. */
    conn.classList.toggle('progress-question__connector--visited', st === 'correct' || st === 'incorrect');
  }
}

/* Re-derives all four questions from the screens' own results and
   renders the bar if the current screen carries one. Called on entry
   into every question screen and from every question's finish handler,
   so returning to an earlier question after resolving a later one
   shows the correct state on both. */
function syncPracticeProgressNav(sectionEl) {
  PRACTICE_QUESTIONS.forEach((q, i) => {
    const item = practiceProgress.questions[i];
    item.state   = deriveQuestionState(q);
    item.visited = item.state !== 'not-answered';
  });
  practiceProgress.results = PRACTICE_QUESTIONS.map(deriveQuestionResult);
  /* The four question results, for diagnostics and for whoever later
     assembles a part result. Deliberately NOT a part-level score —
     see the block comment. */
  window.lomdaState.practiceQuestions = practiceProgress.results;
  const nav = sectionEl && sectionEl.querySelector('.progress-question');
  if (nav) updateProgressQuestion(nav, practiceProgress);
}

/* ═══════════════════════════════════════════════════════════
   S3 — ValueInputQuestion, RATIO-PAIR VARIANT (slide 54)
   SECTION א of QUESTION 3 (its other two sections are S4 and S5) —
   stem lettered "א.", with the
   "מסיבת הפתעה" widget embedded.
   Spec: 720-templates → ValueInputQuestion.md (Variant B's
   one-answer-object contract, laid out as a horizontal ratio rather
   than a stacked fraction) + _question-template-defaults.md +
   _global-components.md → "View My Answer Toggle" / "Long-feedback
   variant — scroll-gated popup" / "Progress Question".
   The state machine is ported from the approved sibling
   methodica-math-ratio-03-03's own S1 — the same template, the same
   ratio-pair shape and the same screen role; only geometry, copy, the
   answer and the scroll gate differ. The scroll-gate mechanics are
   this part's own shared helpers, already shipped on S1 and S2.

   CORRECT VALUES: 4 (left slot) : 1 (right slot), on three independent
   pieces of evidence:
     1. notesSlide49's explicit "الإجابة الصحيحة היא 1 : 4" — whose
        rendered form is "4 : 1" (see the ORDER note below);
     2. the deck's own two placeholder digits — box id=31 carries "4"
        at x 963.7 and box id=7 carries "1" at x 1098.3, i.e. 4 in the
        LEFT slot and 1 in the RIGHT one;
     3. the mathematics, which both feedback panels spell out: the
        small square's side is x and the large one's is 2x, so the
        areas are x² and 4x², and 4x² : x² reduces to 4 : 1.

   RATIO ORDER — no conflict on this screen, unlike 03-03's S1.
   All three sources agree on the same hand. The one thing that does
   need care is that the NOTE and the FEEDBACK PROSE store the string
   in logical order ("1 : 4") while PowerPoint DISPLAYS it swapped
   ("4 : 1"), which a browser's plain UBA would not reproduce. Verified
   glyph by glyph on a fresh 1280×720 export of both the slide and its
   off-canvas incorrect panel: in "לאחר צמצום ב-x² נקבל 1 : 4 ." the
   leftmost glyph is the period, then 4, then the colon, then 1. The
   displayed order is the authored intent and is also the correct
   mathematics, so every ratio below is written out in DISPLAY order
   inside an explicit LTR isolate rather than left to the browser's
   bidi algorithm.

   ATTEMPT MODEL: the 2-attempt default (maxAttempts: 2). The notes
   state no deviation, and the screen carries a hint plus exactly one
   correct and one final-wrong panel — the normal shape of a 2-attempt
   question. The absence of a retry-popup or toggle-label mockup in the
   deck is expected and proves nothing (_question-template-defaults.md
   → "Two traps when reading the source").
   ═══════════════════════════════════════════════════════════ */
const S3VIQ = { correctFirst: 4, correctSecond: 1, maxAttempts: 2 };

let s3viqAttempts     = 0;
let s3viqDone         = false;
let s3viqResult       = null;   // 'correct' | 'wrong2'
let s3viqAnswer       = null;   // { first, second } raw typed strings, snapshotted before the reveal
let s3viqView         = 'solution';
let s3viqLastWrongSig = null;
/* SCROLL-GATE MEMORY (retrofit, 2026-09-15). The gate is per FEEDBACK
   STATE, not per opening. `GateKey` names the state currently on
   screen; `GateRead` names the state whose content has already been
   read to the end. Re-showing a state that is already in `GateRead`
   leaves the gate satisfied, so closing and reopening the same feedback
   never asks the learner to read it twice. Showing a DIFFERENT state
   re-arms the gate by simple inequality, and s3viqResetInitial() clears
   the memory because a fresh question must be read again. */
let s3viqGateKey = null;
let s3viqGateRead = null;
let s3viqGateOk       = true;   // false only while a scroll-gated popup has not been read to the end

/* Shared explanation — the authored body of BOTH feedback panels
   (id=26 "هذا صحيح‎!" and id=28 "זו טעות, / בואו נבין למה:"), which are
   identical in the source and differ only in title
   (_question-template-defaults.md → "Feedback: shared explanation").
   Preserved verbatim, including all five of the source's own empty
   paragraphs as explicit group breaks.
   '' renders as a blank line; a '@math:' prefix renders the paragraph
   as a centred display equation, which is how the source authors its
   four OMML zones (each on its own algn="ctr" paragraph). Every
   expression — display or inline — is an isolated, non-wrapping LTR
   unit so no operator, operand or ratio term can be reordered or split
   by the surrounding RTL paragraph. */
const S3VIQ_GAP = '';

const S3VIQ_EXPLAIN = [
  'نرمز بـ <bdi class="scq-expr" dir="ltr">x</bdi> إلى طول ضلع المربع الصغير.',
  'بحسب النسبة، طول ضلع المربع الكبير هو <bdi class="scq-expr" dir="ltr">2x</bdi>.',
  S3VIQ_GAP,
  'مساحة المربع الصغير هي‎:',
  '@math:<span class="scq-expr">x &middot; x = x<sup>2</sup></span>',
  S3VIQ_GAP,
  'مساحة المربع الكبير هي‎:',
  '@math:<span class="scq-expr">2x &middot; 2x = 4x<sup>2</sup></span>',
  S3VIQ_GAP,
  'لذلك فإن النسبة بين مساحة المربع الكبير ومساحة المربع الصغير هي‎:',
  '@math:<span class="scq-expr">4x<sup>2</sup> : x<sup>2</sup></span>',
  S3VIQ_GAP,
  'بعد الاختزال بـ <bdi class="scq-expr" dir="ltr">x<sup>2</sup></bdi> الذي لا يساوي 0، نحصل على <bdi class="scq-expr" dir="ltr">4 : 1‏‎</bdi>.'
];

/* `scrollGate` on the two long outcomes only — the note asks for it
   ("שימו לב שבמשוב ארוך, לאפשר גלילה") and both source panels are
   438×635, taller than the canvas. The retry popup is two short lines
   and keeps the standard auto-height, ungated popup: the variant is a
   per-feedback-state decision, never a per-template default.
   `action` / `answerView` implement the View My Answer Toggle, which
   this template REQUIRES on the final-wrong path because the reveal is
   destructive — it overwrites the same two inputs the learner typed
   into. Its labels are canonical product UI text, not producer
   content. */
const S3VIQ_POPUP = {
  /* No retry-specific pedagogical explanation is authored for this
     question, so only the canonical fallback title is used —
     _question-template-defaults.md: "A missing pedagogical explanation
     must not be invented." */
  retry:   { bg: '#ffdbdc', title: 'الإجابة غير صحيحة.', body: ['هل نحاول مرة أخرى؟'] },
  /* Authored title, verbatim (id=26 ¶1). */
  correct: { bg: '#edf8ed', title: 'هذا صحيح‎!', body: S3VIQ_EXPLAIN, scrollGate: true },
  wrong2: {
    /* PRODUCER-APPROVED WORDING (2026-09-10) — do not "restore" the
       source string in the name of PPT fidelity. Slide 54 authors this
       title as TWO bold paragraphs, "זו טעות," and "בואו נבין למה:".
       The producer replaced the second with "בואו נלמד ממנה." and
       directed that it sit on the SAME line, so the authored line break
       between them is deliberately dropped and the two runs render as
       one sentence. (The deck itself uses this phrasing elsewhere —
       slide 56's own incorrect panel opens "זו טעות. לא נורא, בואו נלמד
       ממנה.") The canonical terminal-incorrect second line still goes
       on its own line below, via titleSecondLine. */
    bg: '#ffdbdc',
    title: 'هذا خطأ، تعالوا نتعلم منه.',
    titleSecondLine: 'الإجابة الصحيحة تَظهَر الآن.',
    body: S3VIQ_EXPLAIN,
    scrollGate: true,
    action: { label: 'الانتقال إلى إجابتي', view: 'answer' },
    answerView: {
      title: 'الإجابة غير صحيحة.',
      /* NULL, not omitted. The canonical second line belongs to the
         SOLUTION view only — the answer view shows the learner's own
         answer, not the solution. Because the view is built by merging
         this object onto the base outcome, the key has to be explicitly
         cleared or the base's second line leaks into it. */
      titleSecondLine: null,
      body:  ['يمكن مشاهدة الحل الصحيح مرة أخرى‎.'],
      action: { label: 'الإجابة الصحيحة', view: 'solution' },
      scrollGate: false          /* two short lines — never gated */
    }
  }
};

/* Type 1 icons — 16×16 bare vectors, canonical production SVGs
   (720-design-guidelines → components-icons-states.md → "Type 1"). */
const S3VIQ_ICON_CORRECT = '<svg width="16" height="16" viewBox="0 0 14 12" fill="none" aria-hidden="true"><path d="M4.52036 12C4.12296 12.0002 3.74183 11.8126 3.46106 11.4787L0.258452 7.67744C-0.0861507 7.26814 -0.0861507 6.60467 0.258452 6.19536C0.603165 5.78619 1.16194 5.78619 1.50666 6.19536L4.52036 9.77374L12.4933 0.306878C12.8381 -0.102293 13.3968 -0.102293 13.7415 0.306878C14.0862 0.716179 14.0862 1.37966 13.7415 1.78896L5.57966 11.4787C5.29889 11.8126 4.91776 12.0002 4.52036 12Z" fill="#609E12"/></svg>';
const S3VIQ_ICON_WRONG   = '<svg width="16" height="16" viewBox="0 0 14 14" fill="none" aria-hidden="true"><path d="M13.6743 0.325691C13.4657 0.117151 13.1828 0 12.8879 0C12.5929 0 12.31 0.117151 12.1014 0.325691L7 5.42709L1.8986 0.325691C1.69 0.117151 1.40711 0 1.11215 0C0.817183 0 0.534294 0.117151 0.325691 0.325691C0.117151 0.534294 0 0.817183 0 1.11215C0 1.40711 0.117151 1.69 0.325691 1.8986L5.42709 7L0.325691 12.1014C0.117151 12.31 0 12.5929 0 12.8879C0 13.1828 0.117151 13.4657 0.325691 13.6743C0.534294 13.8829 0.817183 14 1.11215 14C1.40711 14 1.69 13.8829 1.8986 13.6743L7 8.57291L12.1014 13.6743C12.31 13.8829 12.5929 14 12.8879 14C13.1828 14 13.4657 13.8829 13.6743 13.6743C13.8829 13.4657 14 13.1828 14 12.8879C14 12.5929 13.8829 12.31 13.6743 12.1014L8.57291 7L13.6743 1.8986C13.8829 1.69 14 1.40711 14 1.11215C14 0.817183 13.8829 0.534294 13.6743 0.325691Z" fill="#B20010"/></svg>';

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

/* type="text" (required so a thousands comma can be typed at all)
   removes the browser's own numeric enforcement, so every keystroke AND
   paste is filtered here — a paste fires `input` too, so no separate
   paste handler is needed (ValueInputQuestion.md → "Numeric input
   format"). */
function s3viqOnInput() {
  if (s3viqDone) return;
  ['first', 'second'].forEach(k => {
    const el = s3viqInput(k);
    const cleaned = el.value.replace(/[^0-9,.-]/g, '');
    if (cleaned !== el.value) el.value = cleaned;
  });
  /* Starting a new attempt dismisses the previous retry popup —
     approved path 2 of "No manual close". Unreachable once answered. */
  if (s3viqAttempts > 0) {
    document.getElementById('s3viq-popup').classList.add('hidden');
    s3viqMark(null);
  }
  s3viqUpdateBar();
}

function s3viqUpdateBar() {
  const chk = document.getElementById('s3viq-check');
  if (!chk || s3viqDone) return;
  /* Retry gate on the FULL joined signature (_question-template-
     defaults.md → "Retry gate"): resubmitting the same pair keeps the
     button disabled, and a comma-only edit normalizes away so it does
     not count as a change. */
  chk.disabled = !s3viqAllFilled() || s3viqSignature() === s3viqLastWrongSig;
}

/* Case B (one semantic answer object): BOTH fields always share the
   SAME single outcome, and there is exactly one verdict icon for the
   whole group — never a per-field icon. */
function s3viqMark(outcome) {
  const ico = document.getElementById('s3viq-ico');
  ['first', 'second'].forEach(k => {
    const f = s3viqField(k);
    f.classList.remove('viq-correct', 'viq-wrong');
    if (outcome) f.classList.add(outcome === 'correct' ? 'viq-correct' : 'viq-wrong');
  });
  if (ico) ico.innerHTML = !outcome ? '' : (outcome === 'correct' ? S3VIQ_ICON_CORRECT : S3VIQ_ICON_WRONG);
}

function s3viqCheck() {
  if (s3viqDone) { advanceScreen(); return; }
  if (!s3viqAllFilled()) return;

  s3viqAttempts++;
  reportQ(['003', 'q1'], s3viqIsCorrect(), s3viqIsCorrect() || s3viqAttempts >= S3VIQ.maxAttempts, String((document.getElementById('s3viq-input-first') || {}).value).trim() + ' : ' + String((document.getElementById('s3viq-input-second') || {}).value).trim());

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

  /* Final wrong. The reveal overwrites the very inputs the learner
     typed into, so snapshot their RAW strings first — the only place
     these survive. Preserved exactly as typed, never reformatted. */
  s3viqAnswer = { first: s3viqInput('first').value, second: s3viqInput('second').value };
  s3viqResult = 'wrong2';
  s3viqView   = 'solution';
  s3viqFinish();
  s3viqApplyView();
}

/* Single owner of the resolved-wrong rendering. The toggle is
   OBJECT-level: one button rewrites BOTH fields (and the single
   combined verdict) in the same pass — this is Case B, not Case A. */
function s3viqApplyView() {
  if (s3viqAnswer === null) return;
  if (s3viqView === 'answer') {
    s3viqInput('first').value  = s3viqAnswer.first;
    s3viqInput('second').value = s3viqAnswer.second;
    s3viqMark('wrong');   // the final attempt was wrong in both views by definition
  } else {
    s3viqInput('first').value  = String(S3VIQ.correctFirst);
    s3viqInput('second').value = String(S3VIQ.correctSecond);
    s3viqMark('correct');
  }
  s3viqShowPopup('wrong2', s3viqView);
}
/* Pure view switch — NEVER routes through s3viqCheck(), never touches
   attempts/done/scoring. */
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
  /* ValueInputQuestion's own completion label is "متابعة" — NOT
     SingleChoiceQuestion's "هل نتابع؟"; the template doc is explicit that
     the two do not share one. Slide 54 authors no progression copy of
     its own, so this canonical fallback applies. The button stays
     DISABLED here when the outcome is scroll-gated; s3viqCheckScrollGate()
     is what releases it. */
  if (chk) setNavLabel(chk, 'متابعة');
  hideHintButton('s3viq-hint');
  syncPracticeProgressNav(document.getElementById('s3'));
}

function s3viqResetInitial() {
  s3viqAttempts     = 0;
  s3viqDone         = false;
  s3viqResult       = null;
  s3viqAnswer       = null;
  s3viqView         = 'solution';
  s3viqLastWrongSig = null;
  s3viqGateOk       = true;
  s3viqGateKey      = null;
  s3viqGateRead     = null;
  ['first', 'second'].forEach(k => {
    const el = s3viqInput(k);
    el.value = '';
    el.disabled = false;
  });
  s3viqMark(null);
  const popup = document.getElementById('s3viq-popup');
  popup.classList.add('hidden');
  popup.classList.remove('is-scroll', 'needs-scroll');
  document.getElementById('s3viq-hint-overlay').classList.add('hidden');
  const chk = document.getElementById('s3viq-check');
  if (chk) { setNavLabel(chk, 'هل إجابتي صحيحة؟'); chk.disabled = true; }
  /* Hint visible and enabled from the start — the canonical template
     default; slide 54 states no availability policy and mocks the
     trigger in its enabled state in the base on-canvas layer. */
  showHintButton('s3viq-hint');
}

function s3viqShowPopup(type, view) {
  const popup = document.getElementById('s3viq-popup');
  if (!popup) return;
  const base = S3VIQ_POPUP[type];
  const cfg  = (view === 'answer' && base.answerView)
    ? Object.assign({}, base, base.answerView)
    : base;
  const body     = document.getElementById('s3viq-popup-body');
  const scroller = document.getElementById('s3viq-popup-scroll');
  const chk      = document.getElementById('s3viq-check');

  popup.style.background = cfg.bg;
  /* Re-anchor on a fresh open, AND on every 'wrong2' render (including
     each toggle) even if already open — the height difference between
     the long solution body and the short answer body must never leave
     the popup pinned away from the bottom bar. */
  if (popup.classList.contains('hidden') || type === 'wrong2') resetPopupPosition(popup);

  /* innerHTML, not textContent: the authored incorrect title is two
     bold lines separated by a REAL break, and the terminal-incorrect
     solution view appends the canonical fixed second line on top of
     that (_question-template-defaults.md → "Terminal-incorrect title
     carries a fixed second line"). */
  const titleEl = document.getElementById('s3viq-popup-title');
  titleEl.innerHTML = cfg.titleSecondLine
    ? (cfg.title + '<br>' + cfg.titleSecondLine)
    : cfg.title;

  body.innerHTML = cfg.body.map(p => {
    if (p === '')                 return '<p class="scq-gap" aria-hidden="true"></p>';
    if (p.indexOf('@math:') === 0) return '<p class="scq-math">' + p.slice(6) + '</p>';
    return '<p>' + p + '</p>';
  }).join('');
  renderPopupAction('s3viq-popup', cfg.action, s3viqToggleView);

  /* Scroll gate — arm before showing, then reset and measure AFTER the
     popup is visible: a display:none element ignores a scrollTop write
     and reports scrollHeight/clientHeight as 0. */
  popup.classList.toggle('is-scroll', !!cfg.scrollGate);
  s3viqGateKey = cfg.scrollGate ? (type + '|' + (view || 'solution')) : null;
  /* Already read in THIS state → stay satisfied across a close/reopen. */
  s3viqGateOk = !cfg.scrollGate || s3viqGateRead === s3viqGateKey;
  if (chk && !s3viqGateOk) chk.disabled = true;
  if (cfg.scrollGate) scroller.setAttribute('tabindex', '0');   // the only way a keyboard user can satisfy the gate
  else                scroller.removeAttribute('tabindex');

  popup.classList.remove('hidden');
  clampPopup(popup);
  scroller.scrollTop = 0;
  if (cfg.scrollGate) {
    placeScrollHint('s3viq-popup');
    s3viqCheckScrollGate();     // unlocks immediately when the content does not actually overflow
  } else {
    popup.classList.remove('needs-scroll');
    if (chk && s3viqDone) chk.disabled = false;
  }
}

/* Bound from the scroll column's own onscroll, and called once right
   after the popup is shown. Once satisfied the gate stays satisfied for
   that feedback STATE — not merely for that opening. Scrolling back up
   never re-locks it, and neither does closing the feedback and
   reopening the same state; `GateRead` is what carries that across
   (producer contract, 2026-09-15). Only a state that has not been read
   to the end arms it. */
function s3viqCheckScrollGate() {
  if (s3viqGateOk) {
    /* Already satisfied — make sure the hint bubble is not left over
       from a previous, still-locked presentation of this popup. */
    const p0 = document.getElementById('s3viq-popup');
    if (p0) p0.classList.remove('needs-scroll');
    return;
  }
  const scroller = document.getElementById('s3viq-popup-scroll');
  const popup    = document.getElementById('s3viq-popup');
  if (!scroller || !popup) return;
  if (!atBottom(scroller)) { popup.classList.add('needs-scroll'); return; }
  s3viqGateOk = true;
  s3viqGateRead = s3viqGateKey;      // this state has now been read to the end
  const chk = document.getElementById('s3viq-check');
  if (chk && s3viqDone) chk.disabled = false;
  popup.classList.remove('needs-scroll');
}

let s3viqResizeRaf = null;
window.addEventListener('resize', () => {
  const popup = document.getElementById('s3viq-popup');
  if (!popup || popup.classList.contains('hidden')) return;
  placeScrollHint('s3viq-popup');
  if (s3viqGateOk) return;
  if (s3viqResizeRaf) cancelAnimationFrame(s3viqResizeRaf);
  s3viqResizeRaf = requestAnimationFrame(s3viqCheckScrollGate);
});

function s3viqOpenHint()  { xapiRequestedHint('003', 'q1'); document.getElementById('s3viq-hint-overlay').classList.remove('hidden'); }
function s3viqCloseHint() { document.getElementById('s3viq-hint-overlay').classList.add('hidden'); }
function s3viqCloseHintOnBackdrop(e) { if (e.target === e.currentTarget) s3viqCloseHint(); }


/* ═══════════════════════════════════════════════════════════
   S4 — ValueInputQuestion, RATIO-PAIR VARIANT (slide 55)
   SECTION ב of QUESTION 3 — stem lettered "ב.", with the same
   "מסיבת הפתעה" widget embedded.
   Spec and contract: identical to S3 (see the S3 block above) —
   720-templates → ValueInputQuestion.md (Variant B's one-answer-object
   contract laid out as a horizontal ratio) + _question-template-
   defaults.md + _global-components.md → "View My Answer Toggle" /
   "Long-feedback variant — scroll-gated popup" / "Progress Question".
   The whole state machine is this part's own S3, re-prefixed; only the
   answer, the explanation and the two titles below differ, because
   slide 55 is slide 54's question asked about PERIMETER instead of
   AREA.

   CORRECT VALUES: 2 (left slot) : 1 (right slot), on three independent
   pieces of evidence:
     1. notesSlide50's explicit "الإجابة الصحيحة היא 1 : 2" — whose
        rendered form is "2 : 1" (see the ORDER note below);
     2. the deck's own two placeholder digits — box id=31 carries "2"
        at x 963.7 and box id=7 carries "1" at x 1098.3, i.e. 2 in the
        LEFT slot and 1 in the RIGHT one;
     3. the mathematics, which both feedback panels spell out: the
        small square's side is x and the large one's is 2x, so the
        perimeters are 4x and 8x, and 8x : 4x reduces to 2 : 1.

   RATIO ORDER — no conflict; all three sources agree on the same hand.
   As on slide 54, the note and the feedback prose store the string in
   logical order ("1 : 2") while PowerPoint DISPLAYS it swapped
   ("2 : 1"), which a browser's plain UBA would not reproduce. Verified
   glyph by glyph on fresh 1280×720 exports of both the slide and its
   off-canvas incorrect panel: in "לאחר צמצום ב-x נקבל 1 : 2 ." the
   leftmost glyph is the period, then 2, then the colon, then 1. Every
   ratio below is therefore written out in DISPLAY order inside an
   explicit LTR isolate rather than left to the browser's bidi.

   TITLES. The correct panel's "هذا صحيح‎!" is authored verbatim (id=26
   ¶1) and matches slide 54's. The incorrect panel authors TWO bold
   paragraphs of its own — "לא בדיוק." and "בואו נסביר:" — whose
   WORDING is preserved verbatim; they are rendered on ONE line rather
   than two, applying the shape the producer approved for S3 on
   2026-09-10 ("המשפט המתוקן צריך להופיע באותו שורה"). That direction
   named S3's own sentence, so only its SHAPE is carried over here and
   nothing is reworded; if the producer wants slide 55's two sentences
   back on separate lines it is a one-character change (a <br> between
   them).

   ATTEMPT MODEL: the 2-attempt default, same evidence as S3 — the
   notes state no deviation and the screen carries a hint plus exactly
   one correct and one final-wrong panel.
   ═══════════════════════════════════════════════════════════ */
const S4VIQ = { correctFirst: 2, correctSecond: 1, maxAttempts: 2 };

let s4viqAttempts     = 0;
let s4viqDone         = false;
let s4viqResult       = null;   // 'correct' | 'wrong2'
let s4viqAnswer       = null;   // { first, second } raw typed strings, snapshotted before the reveal
let s4viqView         = 'solution';
let s4viqLastWrongSig = null;
/* SCROLL-GATE MEMORY (retrofit, 2026-09-15). The gate is per FEEDBACK
   STATE, not per opening. `GateKey` names the state currently on
   screen; `GateRead` names the state whose content has already been
   read to the end. Re-showing a state that is already in `GateRead`
   leaves the gate satisfied, so closing and reopening the same feedback
   never asks the learner to read it twice. Showing a DIFFERENT state
   re-arms the gate by simple inequality, and s4viqResetInitial() clears
   the memory because a fresh question must be read again. */
let s4viqGateKey = null;
let s4viqGateRead = null;
let s4viqGateOk       = true;   // false only while a scroll-gated popup has not been read to the end

/* Shared explanation — the authored body of BOTH feedback panels
   (id=26 "هذا صحيح‎!" and id=28 "לא בדיוק. / בואו נסביר:"), which are
   identical in the source and differ only in title
   (_question-template-defaults.md → "Feedback: shared explanation").
   Preserved verbatim, including all five of the source's own empty
   paragraphs as explicit group breaks.
   '' renders as a blank line; a '@math:' prefix renders the paragraph
   as a centred display equation, which is how the source authors its
   four OMML zones (each on its own algn="ctr" paragraph). Every
   expression — display or inline — is an isolated, non-wrapping LTR
   unit so no operator, operand or ratio term can be reordered or split
   by the surrounding RTL paragraph. */
const S4VIQ_GAP = '';

const S4VIQ_EXPLAIN = [
  'نرمز بـ <bdi class="scq-expr" dir="ltr">x</bdi> إلى طول ضلع المربع الصغير.',
  'بحسب النسبة، طول ضلع المربع الكبير هو <bdi class="scq-expr" dir="ltr">2x</bdi>.',
  S4VIQ_GAP,
  'محيط المربع الصغير هو‎:',
  '@math:<span class="scq-expr">4 &middot; x = 4x</span>',
  S4VIQ_GAP,
  'محيط المربع الكبير هو‎:',
  '@math:<span class="scq-expr">4 &middot; 2x = 8x</span>',
  S4VIQ_GAP,
  'لذلك فإن النسبة بين محيط المربع الكبير ومحيط المربع الصغير هي‎:',
  '@math:<span class="scq-expr">8x : 4x</span>',
  S4VIQ_GAP,
  'بعد الاختزال بـ <bdi class="scq-expr" dir="ltr">x</bdi> الذي لا يساوي 0، نحصل على <bdi class="scq-expr" dir="ltr">2 : 1</bdi>.'
];

/* `scrollGate` on the two long outcomes only — the note asks for it
   ("שימו לב שבמשוב ארוך, לאפשר גלילה") and both source panels are
   438×635, taller than the canvas. The retry popup is two short lines
   and keeps the standard auto-height, ungated popup: the variant is a
   per-feedback-state decision, never a per-template default.
   `action` / `answerView` implement the View My Answer Toggle, which
   this template REQUIRES on the final-wrong path because the reveal is
   destructive — it overwrites the same two inputs the learner typed
   into. Its labels are canonical product UI text, not producer
   content. */
const S4VIQ_POPUP = {
  /* No retry-specific pedagogical explanation is authored for this
     question, so only the canonical fallback title is used —
     _question-template-defaults.md: "A missing pedagogical explanation
     must not be invented." */
  retry:   { bg: '#ffdbdc', title: 'الإجابة غير صحيحة.', body: ['هل نحاول مرة أخرى؟'] },
  /* Authored title, verbatim (id=26 ¶1) — identical to slide 54's. */
  correct: { bg: '#edf8ed', title: 'هذا صحيح‎!', body: S4VIQ_EXPLAIN, scrollGate: true },
  wrong2: {
    /* Authored wording verbatim (id=28 ¶1 + ¶2), rendered on one
       line — see the TITLES note in this block's header. */
    bg: '#ffdbdc',
    title: 'غير دقيق. دعونا نشرح‎:',
    titleSecondLine: 'الإجابة الصحيحة تَظهَر الآن.',
    body: S4VIQ_EXPLAIN,
    scrollGate: true,
    action: { label: 'الانتقال إلى إجابتي', view: 'answer' },
    answerView: {
      title: 'الإجابة غير صحيحة.',
      /* NULL, not omitted. The canonical second line belongs to the
         SOLUTION view only — the answer view shows the learner's own
         answer, not the solution. Because the view is built by merging
         this object onto the base outcome, the key has to be explicitly
         cleared or the base's second line leaks into it. */
      titleSecondLine: null,
      body:  ['يمكن مشاهدة الحل الصحيح مرة أخرى‎.'],
      action: { label: 'الإجابة الصحيحة', view: 'solution' },
      scrollGate: false          /* two short lines — never gated */
    }
  }
};

/* Type 1 icons — 16×16 bare vectors, canonical production SVGs
   (720-design-guidelines → components-icons-states.md → "Type 1"). */
const S4VIQ_ICON_CORRECT = '<svg width="16" height="16" viewBox="0 0 14 12" fill="none" aria-hidden="true"><path d="M4.52036 12C4.12296 12.0002 3.74183 11.8126 3.46106 11.4787L0.258452 7.67744C-0.0861507 7.26814 -0.0861507 6.60467 0.258452 6.19536C0.603165 5.78619 1.16194 5.78619 1.50666 6.19536L4.52036 9.77374L12.4933 0.306878C12.8381 -0.102293 13.3968 -0.102293 13.7415 0.306878C14.0862 0.716179 14.0862 1.37966 13.7415 1.78896L5.57966 11.4787C5.29889 11.8126 4.91776 12.0002 4.52036 12Z" fill="#609E12"/></svg>';
const S4VIQ_ICON_WRONG   = '<svg width="16" height="16" viewBox="0 0 14 14" fill="none" aria-hidden="true"><path d="M13.6743 0.325691C13.4657 0.117151 13.1828 0 12.8879 0C12.5929 0 12.31 0.117151 12.1014 0.325691L7 5.42709L1.8986 0.325691C1.69 0.117151 1.40711 0 1.11215 0C0.817183 0 0.534294 0.117151 0.325691 0.325691C0.117151 0.534294 0 0.817183 0 1.11215C0 1.40711 0.117151 1.69 0.325691 1.8986L5.42709 7L0.325691 12.1014C0.117151 12.31 0 12.5929 0 12.8879C0 13.1828 0.117151 13.4657 0.325691 13.6743C0.534294 13.8829 0.817183 14 1.11215 14C1.40711 14 1.69 13.8829 1.8986 13.6743L7 8.57291L12.1014 13.6743C12.31 13.8829 12.5929 14 12.8879 14C13.1828 14 13.4657 13.8829 13.6743 13.6743C13.8829 13.4657 14 13.1828 14 12.8879C14 12.5929 13.8829 12.31 13.6743 12.1014L8.57291 7L13.6743 1.8986C13.8829 1.69 14 1.40711 14 1.11215C14 0.817183 13.8829 0.534294 13.6743 0.325691Z" fill="#B20010"/></svg>';

function s4viqInput(k) { return document.getElementById('s4viq-input-' + k); }
function s4viqField(k) { return document.getElementById('s4viq-field-' + k); }
function s4viqNormalize(raw) { return String(raw).trim().replace(/,/g, ''); }
function s4viqSignature() {
  return s4viqNormalize(s4viqInput('first').value) + '|' + s4viqNormalize(s4viqInput('second').value);
}
function s4viqAllFilled() {
  return s4viqNormalize(s4viqInput('first').value) !== '' && s4viqNormalize(s4viqInput('second').value) !== '';
}
function s4viqIsCorrect() {
  return Number(s4viqNormalize(s4viqInput('first').value))  === S4VIQ.correctFirst &&
         Number(s4viqNormalize(s4viqInput('second').value)) === S4VIQ.correctSecond;
}

/* type="text" (required so a thousands comma can be typed at all)
   removes the browser's own numeric enforcement, so every keystroke AND
   paste is filtered here — a paste fires `input` too, so no separate
   paste handler is needed (ValueInputQuestion.md → "Numeric input
   format"). */
function s4viqOnInput() {
  if (s4viqDone) return;
  ['first', 'second'].forEach(k => {
    const el = s4viqInput(k);
    const cleaned = el.value.replace(/[^0-9,.-]/g, '');
    if (cleaned !== el.value) el.value = cleaned;
  });
  /* Starting a new attempt dismisses the previous retry popup —
     approved path 2 of "No manual close". Unreachable once answered. */
  if (s4viqAttempts > 0) {
    document.getElementById('s4viq-popup').classList.add('hidden');
    s4viqMark(null);
  }
  s4viqUpdateBar();
}

function s4viqUpdateBar() {
  const chk = document.getElementById('s4viq-check');
  if (!chk || s4viqDone) return;
  /* Retry gate on the FULL joined signature (_question-template-
     defaults.md → "Retry gate"): resubmitting the same pair keeps the
     button disabled, and a comma-only edit normalizes away so it does
     not count as a change. */
  chk.disabled = !s4viqAllFilled() || s4viqSignature() === s4viqLastWrongSig;
}

/* Case B (one semantic answer object): BOTH fields always share the
   SAME single outcome, and there is exactly one verdict icon for the
   whole group — never a per-field icon. */
function s4viqMark(outcome) {
  const ico = document.getElementById('s4viq-ico');
  ['first', 'second'].forEach(k => {
    const f = s4viqField(k);
    f.classList.remove('viq-correct', 'viq-wrong');
    if (outcome) f.classList.add(outcome === 'correct' ? 'viq-correct' : 'viq-wrong');
  });
  if (ico) ico.innerHTML = !outcome ? '' : (outcome === 'correct' ? S4VIQ_ICON_CORRECT : S4VIQ_ICON_WRONG);
}

function s4viqCheck() {
  if (s4viqDone) { advanceScreen(); return; }
  if (!s4viqAllFilled()) return;

  s4viqAttempts++;
  reportQ(['003', 'q2'], s4viqIsCorrect(), s4viqIsCorrect() || s4viqAttempts >= S4VIQ.maxAttempts, String((document.getElementById('s4viq-input-first') || {}).value).trim() + ' : ' + String((document.getElementById('s4viq-input-second') || {}).value).trim());

  if (s4viqIsCorrect()) {
    s4viqMark('correct');
    s4viqResult = 'correct';
    s4viqFinish();
    s4viqShowPopup('correct');
    return;
  }

  if (s4viqAttempts < S4VIQ.maxAttempts) {
    s4viqMark('wrong');
    s4viqLastWrongSig = s4viqSignature();
    document.getElementById('s4viq-check').disabled = true;
    s4viqShowPopup('retry');
    return;
  }

  /* Final wrong. The reveal overwrites the very inputs the learner
     typed into, so snapshot their RAW strings first — the only place
     these survive. Preserved exactly as typed, never reformatted. */
  s4viqAnswer = { first: s4viqInput('first').value, second: s4viqInput('second').value };
  s4viqResult = 'wrong2';
  s4viqView   = 'solution';
  s4viqFinish();
  s4viqApplyView();
}

/* Single owner of the resolved-wrong rendering. The toggle is
   OBJECT-level: one button rewrites BOTH fields (and the single
   combined verdict) in the same pass — this is Case B, not Case A. */
function s4viqApplyView() {
  if (s4viqAnswer === null) return;
  if (s4viqView === 'answer') {
    s4viqInput('first').value  = s4viqAnswer.first;
    s4viqInput('second').value = s4viqAnswer.second;
    s4viqMark('wrong');   // the final attempt was wrong in both views by definition
  } else {
    s4viqInput('first').value  = String(S4VIQ.correctFirst);
    s4viqInput('second').value = String(S4VIQ.correctSecond);
    s4viqMark('correct');
  }
  s4viqShowPopup('wrong2', s4viqView);
}
/* Pure view switch — NEVER routes through s4viqCheck(), never touches
   attempts/done/scoring. */
function s4viqToggleView() {
  if (s4viqAnswer === null) return;
  s4viqView = (s4viqView === 'solution') ? 'answer' : 'solution';
  s4viqApplyView();
}

function s4viqFinish() {
  s4viqDone = true;
  s4viqInput('first').disabled  = true;
  s4viqInput('second').disabled = true;
  const chk = document.getElementById('s4viq-check');
  /* ValueInputQuestion's own completion label is "متابعة" — NOT
     SingleChoiceQuestion's "هل نتابع؟"; the template doc is explicit that
     the two do not share one. Slide 54 authors no progression copy of
     its own, so this canonical fallback applies. The button stays
     DISABLED here when the outcome is scroll-gated; s4viqCheckScrollGate()
     is what releases it. */
  if (chk) setNavLabel(chk, 'متابعة');
  hideHintButton('s4viq-hint');
  syncPracticeProgressNav(document.getElementById('s4'));
}

function s4viqResetInitial() {
  s4viqAttempts     = 0;
  s4viqDone         = false;
  s4viqResult       = null;
  s4viqAnswer       = null;
  s4viqView         = 'solution';
  s4viqLastWrongSig = null;
  s4viqGateOk       = true;
  s4viqGateKey      = null;
  s4viqGateRead     = null;
  ['first', 'second'].forEach(k => {
    const el = s4viqInput(k);
    el.value = '';
    el.disabled = false;
  });
  s4viqMark(null);
  const popup = document.getElementById('s4viq-popup');
  popup.classList.add('hidden');
  popup.classList.remove('is-scroll', 'needs-scroll');
  document.getElementById('s4viq-hint-overlay').classList.add('hidden');
  const chk = document.getElementById('s4viq-check');
  if (chk) { setNavLabel(chk, 'هل إجابتي صحيحة؟'); chk.disabled = true; }
  /* Hint visible and enabled from the start — the canonical template
     default; slide 55 states no availability policy and mocks the
     trigger in its enabled state in the base on-canvas layer. */
  showHintButton('s4viq-hint');
}

function s4viqShowPopup(type, view) {
  const popup = document.getElementById('s4viq-popup');
  if (!popup) return;
  const base = S4VIQ_POPUP[type];
  const cfg  = (view === 'answer' && base.answerView)
    ? Object.assign({}, base, base.answerView)
    : base;
  const body     = document.getElementById('s4viq-popup-body');
  const scroller = document.getElementById('s4viq-popup-scroll');
  const chk      = document.getElementById('s4viq-check');

  popup.style.background = cfg.bg;
  /* Re-anchor on a fresh open, AND on every 'wrong2' render (including
     each toggle) even if already open — the height difference between
     the long solution body and the short answer body must never leave
     the popup pinned away from the bottom bar. */
  if (popup.classList.contains('hidden') || type === 'wrong2') resetPopupPosition(popup);

  /* innerHTML, not textContent: the authored incorrect title is two
     bold lines separated by a REAL break, and the terminal-incorrect
     solution view appends the canonical fixed second line on top of
     that (_question-template-defaults.md → "Terminal-incorrect title
     carries a fixed second line"). */
  const titleEl = document.getElementById('s4viq-popup-title');
  titleEl.innerHTML = cfg.titleSecondLine
    ? (cfg.title + '<br>' + cfg.titleSecondLine)
    : cfg.title;

  body.innerHTML = cfg.body.map(p => {
    if (p === '')                 return '<p class="scq-gap" aria-hidden="true"></p>';
    if (p.indexOf('@math:') === 0) return '<p class="scq-math">' + p.slice(6) + '</p>';
    return '<p>' + p + '</p>';
  }).join('');
  renderPopupAction('s4viq-popup', cfg.action, s4viqToggleView);

  /* Scroll gate — arm before showing, then reset and measure AFTER the
     popup is visible: a display:none element ignores a scrollTop write
     and reports scrollHeight/clientHeight as 0. */
  popup.classList.toggle('is-scroll', !!cfg.scrollGate);
  s4viqGateKey = cfg.scrollGate ? (type + '|' + (view || 'solution')) : null;
  /* Already read in THIS state → stay satisfied across a close/reopen. */
  s4viqGateOk = !cfg.scrollGate || s4viqGateRead === s4viqGateKey;
  if (chk && !s4viqGateOk) chk.disabled = true;
  if (cfg.scrollGate) scroller.setAttribute('tabindex', '0');   // the only way a keyboard user can satisfy the gate
  else                scroller.removeAttribute('tabindex');

  popup.classList.remove('hidden');
  clampPopup(popup);
  scroller.scrollTop = 0;
  if (cfg.scrollGate) {
    placeScrollHint('s4viq-popup');
    s4viqCheckScrollGate();     // unlocks immediately when the content does not actually overflow
  } else {
    popup.classList.remove('needs-scroll');
    if (chk && s4viqDone) chk.disabled = false;
  }
}

/* Bound from the scroll column's own onscroll, and called once right
   after the popup is shown. Once satisfied the gate stays satisfied for
   that feedback STATE — not merely for that opening. Scrolling back up
   never re-locks it, and neither does closing the feedback and
   reopening the same state; `GateRead` is what carries that across
   (producer contract, 2026-09-15). Only a state that has not been read
   to the end arms it. */
function s4viqCheckScrollGate() {
  if (s4viqGateOk) {
    /* Already satisfied — make sure the hint bubble is not left over
       from a previous, still-locked presentation of this popup. */
    const p0 = document.getElementById('s4viq-popup');
    if (p0) p0.classList.remove('needs-scroll');
    return;
  }
  const scroller = document.getElementById('s4viq-popup-scroll');
  const popup    = document.getElementById('s4viq-popup');
  if (!scroller || !popup) return;
  if (!atBottom(scroller)) { popup.classList.add('needs-scroll'); return; }
  s4viqGateOk = true;
  s4viqGateRead = s4viqGateKey;      // this state has now been read to the end
  const chk = document.getElementById('s4viq-check');
  if (chk && s4viqDone) chk.disabled = false;
  popup.classList.remove('needs-scroll');
}

let s4viqResizeRaf = null;
window.addEventListener('resize', () => {
  const popup = document.getElementById('s4viq-popup');
  if (!popup || popup.classList.contains('hidden')) return;
  placeScrollHint('s4viq-popup');
  if (s4viqGateOk) return;
  if (s4viqResizeRaf) cancelAnimationFrame(s4viqResizeRaf);
  s4viqResizeRaf = requestAnimationFrame(s4viqCheckScrollGate);
});

function s4viqOpenHint()  { xapiRequestedHint('003', 'q2'); document.getElementById('s4viq-hint-overlay').classList.remove('hidden'); }
function s4viqCloseHint() { document.getElementById('s4viq-hint-overlay').classList.add('hidden'); }
function s4viqCloseHintOnBackdrop(e) { if (e.target === e.currentTarget) s4viqCloseHint(); }


/* ═══════════════════════════════════════════════════════════
   S5 — ValueInputQuestion, VARIANT A (single value) — slide 56
   SECTION ג of QUESTION 3 — stem lettered "ג.", the LAST section,
   so answering it is what finally RESOLVES Question 3 and lets its
   progress dot leave `current` for ✓ (only if א, ב and ג are all
   correct) or ✕. Until this screen resolves, the dot stays `current`
   however א and ב went. Local feedback on THIS screen is immediate as
   always — it is the QUESTION's dot that must not resolve early.

   Spec: 720-templates → ValueInputQuestion.md ("Markup contract —
   Variant A": one inline input box, ONE verdict icon INSIDE the
   field's own border) + _question-template-defaults.md +
   _global-components.md → "View My Answer Toggle" / "Long-feedback
   variant — scroll-gated popup" / "Progress Question". The state
   machine is this part's own S3/S4 reduced to a single field, and the
   inline single-field shape is the approved sibling
   methodica-math-ratio-03-03's S2 (slide 41).

   CORRECT VALUE: 16 — on three independent pieces of evidence:
     1. notesSlide51's explicit "الإجابة الصحيحة היא 16";
     2. the deck's own placeholder digit in the single value box
        (id=7, "16");
     3. the mathematics, which both feedback panels spell out in full:
        the perimeters are x and 2x with x + 2x = 48, so x = 16 is the
        SMALL square's PERIMETER, its side is 16 ÷ 4 = 4, and its area
        is 4 · 4 = 16 cm². The answer and that intermediate share the
        value 16 by coincidence — the question asks for the AREA.

   NO RATIO ORDER QUESTION on this screen: there is ONE box, and the
   authored unit word "סמ\"ר." sits to its physical LEFT because the
   source paragraph puts its spacer before it (confirmed against a
   fresh render). The two ratios in the STEM are a different matter and
   are handled in index.html — slide 56 authors each as its own en-US
   run reading "2 : 1", which PowerPoint renders unswapped, so unlike
   slides 54/55 the stored order already is the displayed order.

   ATTEMPT MODEL: the 2-attempt default. The notes state no deviation,
   and the screen carries a hint plus exactly one correct and one
   final-wrong panel — the normal shape of a 2-attempt question.

   ACCEPTED FORMAT: `inputmode="decimal"` per Variant A, and the shared
   character filter allows digits, comma, dot and minus; the comparison
   normalizes thousands commas away, so "16" and "16.0" both pass while
   the retry gate does not treat a comma-only edit as a new answer.
   ═══════════════════════════════════════════════════════════ */
const S5VIQ = { correct: 16, maxAttempts: 2 };

let s5viqAttempts     = 0;
let s5viqDone         = false;
let s5viqResult       = null;   // 'correct' | 'wrong2'
let s5viqAnswer       = null;   // the RAW typed string, snapshotted before the reveal
let s5viqView         = 'solution';
let s5viqLastWrongSig = null;
/* SCROLL-GATE MEMORY (retrofit, 2026-09-15). The gate is per FEEDBACK
   STATE, not per opening. `GateKey` names the state currently on
   screen; `GateRead` names the state whose content has already been
   read to the end. Re-showing a state that is already in `GateRead`
   leaves the gate satisfied, so closing and reopening the same feedback
   never asks the learner to read it twice. Showing a DIFFERENT state
   re-arms the gate by simple inequality, and s5viqResetInitial() clears
   the memory because a fresh question must be read again. */
let s5viqGateKey = null;
let s5viqGateRead = null;
let s5viqGateOk       = true;

/* Shared explanation — the authored body of BOTH feedback panels
   (id=26 "صحيح جدًا‎!" and id=28 "هذا خطأ، لا بأس، تعالوا نتعلّم منه."),
   which are identical in the source and differ only in title
   (_question-template-defaults.md → "Feedback: shared explanation").
   Preserved verbatim, including all four of the source's own empty
   paragraphs as explicit group breaks.
   The ratio "1 : 2" in the source's ¶5 is written here in the order
   the slide RENDERS it ("2 : 1") — verified glyph by glyph on a fresh
   export of the off-canvas panel — inside an explicit LTR isolate. */
const S5VIQ_GAP = '';

const S5VIQ_EXPLAIN = [
  'مجموع محيطَيّ المربع الكبير والمربع الصغير هو 48 سم‎',
  S5VIQ_GAP,
  'النسبة بين محيطَيّ المربعين هي <bdi class="scq-expr" dir="ltr">2 : 1‏‎</bdi>.',
  S5VIQ_GAP,
  'نرمز بـ <bdi class="scq-expr" dir="ltr">x</bdi> إلى محيط المربع الصغير، فيكون محيط المربع الكبير <bdi class="scq-expr" dir="ltr">2x</bdi>.',
  'يمكننا الآن بناء معادلة‎:',
  '@math:<span class="scq-expr">2x + x = 48</span>',
  '@math:<span class="scq-expr">3x = 48</span>',
  '@math:<span class="scq-expr">x = 16</span>',
  S5VIQ_GAP,
  'أي إن محيط المربع الصغير هو 16 سم‎.',
  'نقسم محيط المربع الصغير على 4 فنحصل على طول ضلع المربع الصغير: 4 سم‎.',
  '@math:<span class="scq-expr">4 &middot; 4 = 16</span>',
  'الاستنتاج: مساحة المربع هي 16 سم مربع.'
];

/* `scrollGate` on the two long outcomes only — the note asks for it and
   both source panels are 438×585, taller than the canvas. `action` /
   `answerView` implement the View My Answer Toggle, which this template
   REQUIRES on the final-wrong path because the reveal is destructive:
   it overwrites the same input the learner typed into. */
const S5VIQ_POPUP = {
  /* No retry-specific pedagogical explanation is authored, so only the
     canonical fallback title is used — a missing pedagogical
     explanation must not be invented. */
  retry:   { bg: '#ffdbdc', title: 'الإجابة غير صحيحة.', body: ['هل نحاول مرة أخرى؟'] },
  /* Authored title, verbatim (id=26 ¶1). */
  correct: { bg: '#edf8ed', title: 'صحيح جدًا‎!', body: S5VIQ_EXPLAIN, scrollGate: true },
  wrong2: {
    /* Authored title, verbatim (id=28 ¶1) — a SINGLE source paragraph
       here, already one sentence on one line, which is also the shape
       the producer approved for S3 on 2026-09-10. */
    bg: '#ffdbdc',
    title: 'هذا خطأ، لا بأس، تعالوا نتعلّم منه.',
    titleSecondLine: 'الإجابة الصحيحة تَظهَر الآن.',
    body: S5VIQ_EXPLAIN,
    scrollGate: true,
    action: { label: 'الانتقال إلى إجابتي', view: 'answer' },
    answerView: {
      title: 'الإجابة غير صحيحة.',
      titleSecondLine: null,   // NULL, not omitted — see the S3 note
      body:  ['يمكن مشاهدة الحل الصحيح مرة أخرى‎.'],
      action: { label: 'الإجابة الصحيحة', view: 'solution' },
      scrollGate: false
    }
  }
};

function s5viqInput() { return document.getElementById('s5viq-input'); }
function s5viqNormalize(raw) { return String(raw).trim().replace(/,/g, ''); }
function s5viqFilled() { return s5viqNormalize(s5viqInput().value) !== ''; }
function s5viqIsCorrect() { return Number(s5viqNormalize(s5viqInput().value)) === S5VIQ.correct; }

/* type="text" (required so a thousands comma can be typed at all)
   removes the browser's own numeric enforcement, so every keystroke AND
   paste is filtered here — a paste fires `input` too, so no separate
   paste handler is needed. */
function s5viqOnInput() {
  if (s5viqDone) return;
  const el = s5viqInput();
  const cleaned = el.value.replace(/[^0-9,.-]/g, '');
  if (cleaned !== el.value) el.value = cleaned;
  /* Starting a new attempt dismisses the previous retry popup —
     approved path 2 of "No manual close". Unreachable once answered. */
  if (s5viqAttempts > 0) {
    document.getElementById('s5viq-popup').classList.add('hidden');
    s5viqMark(null);
  }
  s5viqUpdateBar();
}

function s5viqUpdateBar() {
  const chk = document.getElementById('s5viq-check');
  if (!chk || s5viqDone) return;
  /* Retry gate on the normalized value: resubmitting the same answer
     keeps the button disabled, and a comma-only edit normalizes away so
     it does not count as a change. */
  chk.disabled = !s5viqFilled() || s5viqNormalize(s5viqInput().value) === s5viqLastWrongSig;
}

/* ONE field = one answer = exactly ONE verdict icon, and it lives
   INSIDE the field's own bordered box. */
function s5viqMark(outcome) {
  const f   = document.getElementById('s5viq-field');
  const ico = document.getElementById('s5viq-ico');
  f.classList.remove('viq-correct', 'viq-wrong');
  if (outcome) f.classList.add(outcome === 'correct' ? 'viq-correct' : 'viq-wrong');
  if (ico) ico.innerHTML = !outcome ? '' : (outcome === 'correct' ? S3VIQ_ICON_CORRECT : S3VIQ_ICON_WRONG);
}

function s5viqCheck() {
  if (s5viqDone) { advanceScreen(); return; }
  if (!s5viqFilled()) return;

  s5viqAttempts++;
  reportQ(['003', 'q3'], s5viqIsCorrect(), s5viqIsCorrect() || s5viqAttempts >= S5VIQ.maxAttempts, String((document.getElementById('s5viq-input') || {}).value).trim());

  if (s5viqIsCorrect()) {
    s5viqMark('correct');
    s5viqResult = 'correct';
    s5viqFinish();
    s5viqShowPopup('correct');
    return;
  }

  if (s5viqAttempts < S5VIQ.maxAttempts) {
    s5viqMark('wrong');
    s5viqLastWrongSig = s5viqNormalize(s5viqInput().value);
    document.getElementById('s5viq-check').disabled = true;
    s5viqShowPopup('retry');
    return;
  }

  /* Final wrong. The reveal overwrites the very input the learner typed
     into, so snapshot the RAW string first — the only place it
     survives. Preserved exactly as typed, never reformatted. */
  s5viqAnswer = s5viqInput().value;
  s5viqResult = 'wrong2';
  s5viqView   = 'solution';
  s5viqFinish();
  s5viqApplyView();
}

/* Single owner of the resolved-wrong rendering. */
function s5viqApplyView() {
  if (s5viqAnswer === null) return;
  if (s5viqView === 'answer') {
    s5viqInput().value = s5viqAnswer;
    s5viqMark('wrong');   // the final attempt was wrong in both views by definition
  } else {
    s5viqInput().value = String(S5VIQ.correct);
    s5viqMark('correct');
  }
  s5viqShowPopup('wrong2', s5viqView);
}
/* Pure view switch — NEVER routes through s5viqCheck(), never touches
   attempts/done/scoring. */
function s5viqToggleView() {
  if (s5viqAnswer === null) return;
  s5viqView = (s5viqView === 'solution') ? 'answer' : 'solution';
  s5viqApplyView();
}

function s5viqFinish() {
  s5viqDone = true;
  s5viqInput().disabled = true;
  const chk = document.getElementById('s5viq-check');
  /* ValueInputQuestion's own completion label. Slide 56 authors no
     progression copy of its own, so this canonical fallback applies.
     The button stays DISABLED here when the outcome is scroll-gated;
     s5viqCheckScrollGate() is what releases it. */
  if (chk) setNavLabel(chk, 'متابعة');
  hideHintButton('s5viq-hint');
  syncPracticeProgressNav(document.getElementById('s5'));
}

function s5viqResetInitial() {
  s5viqAttempts     = 0;
  s5viqDone         = false;
  s5viqResult       = null;
  s5viqAnswer       = null;
  s5viqView         = 'solution';
  s5viqLastWrongSig = null;
  s5viqGateOk       = true;
  s5viqGateKey      = null;
  s5viqGateRead     = null;
  const el = s5viqInput();
  el.value = '';
  el.disabled = false;
  s5viqMark(null);
  const popup = document.getElementById('s5viq-popup');
  popup.classList.add('hidden');
  popup.classList.remove('is-scroll', 'needs-scroll');
  document.getElementById('s5viq-hint-overlay').classList.add('hidden');
  const chk = document.getElementById('s5viq-check');
  if (chk) { setNavLabel(chk, 'هل إجابتي صحيحة؟'); chk.disabled = true; }
  showHintButton('s5viq-hint');
}

function s5viqShowPopup(type, view) {
  const popup = document.getElementById('s5viq-popup');
  if (!popup) return;
  const base = S5VIQ_POPUP[type];
  const cfg  = (view === 'answer' && base.answerView)
    ? Object.assign({}, base, base.answerView)
    : base;
  const body     = document.getElementById('s5viq-popup-body');
  const scroller = document.getElementById('s5viq-popup-scroll');
  const chk      = document.getElementById('s5viq-check');

  popup.style.background = cfg.bg;
  if (popup.classList.contains('hidden') || type === 'wrong2') resetPopupPosition(popup);

  const titleEl = document.getElementById('s5viq-popup-title');
  titleEl.innerHTML = cfg.titleSecondLine
    ? (cfg.title + '<br>' + cfg.titleSecondLine)
    : cfg.title;

  body.innerHTML = cfg.body.map(p => {
    if (p === '')                 return '<p class="scq-gap" aria-hidden="true"></p>';
    if (p.indexOf('@math:') === 0) return '<p class="scq-math">' + p.slice(6) + '</p>';
    return '<p>' + p + '</p>';
  }).join('');
  renderPopupAction('s5viq-popup', cfg.action, s5viqToggleView);

  /* Scroll gate — arm before showing, then reset and measure AFTER the
     popup is visible: a display:none element ignores a scrollTop write
     and reports scrollHeight/clientHeight as 0. */
  popup.classList.toggle('is-scroll', !!cfg.scrollGate);
  s5viqGateKey = cfg.scrollGate ? (type + '|' + (view || 'solution')) : null;
  /* Already read in THIS state → stay satisfied across a close/reopen. */
  s5viqGateOk = !cfg.scrollGate || s5viqGateRead === s5viqGateKey;
  if (chk && !s5viqGateOk) chk.disabled = true;
  if (cfg.scrollGate) scroller.setAttribute('tabindex', '0');
  else                scroller.removeAttribute('tabindex');

  popup.classList.remove('hidden');
  clampPopup(popup);
  scroller.scrollTop = 0;
  if (cfg.scrollGate) {
    placeScrollHint('s5viq-popup');
    s5viqCheckScrollGate();
  } else {
    popup.classList.remove('needs-scroll');
    if (chk && s5viqDone) chk.disabled = false;
  }
}

function s5viqCheckScrollGate() {
  if (s5viqGateOk) {
    /* Already satisfied — make sure the hint bubble is not left over
       from a previous, still-locked presentation of this popup. */
    const p0 = document.getElementById('s5viq-popup');
    if (p0) p0.classList.remove('needs-scroll');
    return;
  }
  const scroller = document.getElementById('s5viq-popup-scroll');
  const popup    = document.getElementById('s5viq-popup');
  if (!scroller || !popup) return;
  if (!atBottom(scroller)) { popup.classList.add('needs-scroll'); return; }
  s5viqGateOk = true;
  s5viqGateRead = s5viqGateKey;      // this state has now been read to the end
  const chk = document.getElementById('s5viq-check');
  if (chk && s5viqDone) chk.disabled = false;
  popup.classList.remove('needs-scroll');
}

let s5viqResizeRaf = null;
window.addEventListener('resize', () => {
  const popup = document.getElementById('s5viq-popup');
  if (!popup || popup.classList.contains('hidden')) return;
  placeScrollHint('s5viq-popup');
  if (s5viqGateOk) return;
  if (s5viqResizeRaf) cancelAnimationFrame(s5viqResizeRaf);
  s5viqResizeRaf = requestAnimationFrame(s5viqCheckScrollGate);
});

function s5viqOpenHint()  { xapiRequestedHint('003', 'q3'); document.getElementById('s5viq-hint-overlay').classList.remove('hidden'); }
function s5viqCloseHint() { document.getElementById('s5viq-hint-overlay').classList.add('hidden'); }
function s5viqCloseHintOnBackdrop(e) { if (e.target === e.currentTarget) s5viqCloseHint(); }

/* ═══════════════════════════════════════════════════════════
   S6 — MultipleChoiceQuestion, text-pill variant (slide 57)
   QUESTION 4 of four — the last question in the series, resolved by
   this screen alone ("שאלה 4 מתוך 4" in its corrected note). It
   carries the same shared progress bar as every other question screen
   in the series. The
   note names the template outright ("מסך לא בתבנית, צריך להיות
   multiplechoice").
   Spec: 720-templates → MultipleChoiceQuestion.md — SingleChoice-
   Question's workflow with exactly one delta: selection is a Set,
   with per-selected-option feedback on a non-final wrong attempt and a
   full-set reveal on the final one. + _question-template-defaults.md
   + _global-components.md → Feedback popup / Long-feedback scroll gate.
   The state machine is the approved sibling methodica-math-ratio-03-03's
   S4 (the same lomda's only shipped MultipleChoiceQuestion), with this
   part's own scroll-gate mechanics (already shipped on S1-S5) layered
   on its popup.

   CORRECT SET = {o1, o2, o4} — three independent, agreeing signals,
   recorded in full on this screen's markup in index.html: the capsule
   strokes (#00B050 ×3, #C00000 ×1), the marker squares (filled green
   ×3, empty red outline ×1), and the mathematics the feedback spells
   out. o3 ("adding 10 minutes to each part keeps the ratio") is the
   distractor — additive change does not preserve a ratio; the feedback's
   own last paragraph says exactly that.

   DOES NOT use the View My Answer Toggle: this template shows the full
   correct set alongside the learner's own selections simultaneously, so
   the "one slot, two possible values" applicability rule excludes it.

   ATTEMPT MODEL: the 2-attempt default. The notes state no deviation,
   and the screen carries a hint plus exactly one correct and one
   final-wrong panel.
   ═══════════════════════════════════════════════════════════ */
const S6MC = { correct: new Set(['o1', 'o2', 'o4']), maxAttempts: 2 };

let s6mcSelected  = new Set();
let s6mcAttempts  = 0;
let s6mcDone      = false;
let s6mcResult    = null;   // 'correct' | 'wrong2'
let s6mcLastWrong = null;   // serialized Set of the immediately-preceding wrong attempt
/* SCROLL-GATE MEMORY (retrofit, 2026-09-15). The gate is per FEEDBACK
   STATE, not per opening. `GateKey` names the state currently on
   screen; `GateRead` names the state whose content has already been
   read to the end. Re-showing a state that is already in `GateRead`
   leaves the gate satisfied, so closing and reopening the same feedback
   never asks the learner to read it twice. Showing a DIFFERENT state
   re-arms the gate by simple inequality, and s6mcResetInitial() clears
   the memory because a fresh question must be read again. */
let s6mcGateKey = null;
let s6mcGateRead = null;
let s6mcGateOk    = true;   // false only while a scroll-gated popup has not been read to the end

/* Stable serialization for the retry gate — sorted so membership, not
   click order, is what is compared. */
function s6mcSig(set) { return [...set].sort().join('|'); }

/* Shared explanation — the authored body of BOTH panels (id=26 correct,
   id=16 final-wrong), structurally identical (11 paragraphs, 4 empties,
   the same two OMML fractions) and differing only in title, so this is
   the shared-explanation default rather than an assumption. Preserved
   verbatim, including the source's own empty paragraphs as real breaks.
   Every Latin expression is an isolated, non-wrapping LTR unit; the
   fractions are the .frac component inside such a unit, inline where
   the source sets them inline (both end their own sentence).

   ONE DELIBERATE, SCOPED OMISSION: the source's ¶7 ends its OMML
   fraction with a lone "." run — "…לכן [a/(a+b)] ." — while ¶8
   continues the very same sentence ("הוא חלקו של…") and carries its
   real full stop. On the rendered slide that stray mark is a barely
   visible dot at the fraction's baseline; rendered at full contrast in
   HTML it reads as a typo mid-sentence. It is dropped here, documented
   as an authoring slip rather than wording; every other character is
   as authored. Restoring it is a one-character change if the producer
   disagrees. */
const S6MC_GAP = '';
const S6MC_FRAC_AB   = '<bdi class="scq-expr" dir="ltr"><span class="frac"><span class="frac-num">a</span><span class="frac-den">b</span></span></bdi>';
const S6MC_FRAC_AAB  = '<bdi class="scq-expr" dir="ltr"><span class="frac"><span class="frac-num">a</span><span class="frac-den">a+b</span></span></bdi>';

const S6MC_EXPLAIN = [
  'النسبة بين زمن المشي في المنطقة المستوية وزمن المشي في الوادي هي <bdi class="scq-expr" dir="ltr">a : b</bdi>، ويمكننا أيضًا كتابتها بصيغة كسر: ' + S6MC_FRAC_AB,
  'نوسّع بـ <bdi class="scq-expr" dir="ltr">x</bdi>، ونحصل على النسبة <bdi class="scq-expr" dir="ltr">ax : bx</bdi> .',
  'ولذلك زمن المسار الكلّي سيكون <bdi class="scq-expr" dir="ltr">ax + bx</bdi> .',
  S6MC_GAP,
  'زمن المشي الكلي هو <bdi class="scq-expr" dir="ltr">a+b</bdi> لذلك ' + S6MC_FRAC_AAB,
  'هو جزء زمن المشي في المنطقة المستوية من زمن المشي الكلّي.',
  S6MC_GAP,
  'إذا أُضيفت 10 دقائق إلى كل جزء من أجزاء المسار، لن تُحفظ النسبة، لأن إضافة عدد إلى كل حدّ من حدود النسبة لا تحافظ على النسبة.'
];

/* `scrollGate` on the two long outcomes — the note asks for it in as
   many words and both source panels are 438×653, taller than the
   canvas. The mechanism unlocks immediately when the body does not
   actually overflow the gated box (rule 6), so it can never strand the
   learner; whether it is genuinely needed here is MEASURED after the
   build and reported (the approved sibling 03-03 removed a vacuous gate
   on its own MCQ screen at the producer's request). No `action` /
   `answerView`: this template has no View My Answer Toggle. */
const S6MC_POPUP = {
  retry:   { bg: '#ffdbdc', title: 'الإجابة غير صحيحة.', body: ['هل نحاول مرة أخرى؟'] },
  /* Authored title, verbatim (id=26 ¶1). */
  correct: { bg: '#edf8ed', title: 'إجابتكم صحيحة، كل الاحترام!', body: S6MC_EXPLAIN, scrollGate: true },
  /* Authored title, verbatim (id=16 ¶1) — a single source paragraph. */
  wrong2:  { bg: '#ffdbdc', title: 'هذا خطأ. وهذه فرصة للتوقف وتوضيح الفكرة‎.', titleSecondLine: 'الإجابة الصحيحة تَظهَر الآن.', body: S6MC_EXPLAIN, scrollGate: true }
};

function s6mcOpts()  { return [...document.querySelectorAll('#s6 .scq-opt')]; }

/* Toggle — adds/removes ONLY this id. Never clears the other
   selections; this is the one behaviour that must not regress into
   SingleChoiceQuestion's exclusive select. */
function s6mcToggle(id) {
  if (s6mcDone) return;
  if (s6mcSelected.has(id)) s6mcSelected.delete(id);
  else                      s6mcSelected.add(id);
  /* A genuine change clears the previous attempt's per-option marks and
     dismisses its retry popup — approved path 2 of "No manual close". */
  if (s6mcAttempts > 0) {
    document.getElementById('s6mc-popup').classList.add('hidden');
    s6mcOpts().forEach(o => o.classList.remove('correct', 'wrong'));
  }
  s6mcRender();
}

function s6mcRender() {
  s6mcOpts().forEach(o => {
    const on = s6mcSelected.has(o.dataset.id);
    o.classList.toggle('selected', on);
    o.setAttribute('aria-checked', on ? 'true' : 'false');
  });
  const chk = document.getElementById('s6mc-check');
  if (!chk || s6mcDone) return;
  /* Gate on "at least one selected" — never on matching the number of
     correct answers; the learner may legitimately try a wrong count.
     Retry gate on top: the full sorted Set signature, so an off/on
     round-trip back to the same Set does NOT re-enable the button. */
  chk.disabled = s6mcSelected.size < 1 || s6mcSig(s6mcSelected) === s6mcLastWrong;
}

function s6mcCheck() {
  if (s6mcDone) { advanceScreen(); return; }
  if (s6mcSelected.size < 1) return;

  s6mcAttempts++;
  const correct = setsEqual(s6mcSelected, S6MC.correct);
  reportQ(['004', 'q1'], correct, correct || s6mcAttempts >= S6MC.maxAttempts, xapiChoiceAnswer('004', 'q1', '#s6 .scq-opt', [...s6mcSelected]));

  if (correct) {
    /* Fully correct: mark every option in the correct set green. */
    s6mcOpts().forEach(o => {
      o.classList.remove('selected', 'wrong');
      if (S6MC.correct.has(o.dataset.id)) o.classList.add('correct');
    });
    s6mcResult = 'correct';
    s6mcFinish();
    s6mcShowPopup('correct');
    return;
  }

  if (s6mcAttempts < S6MC.maxAttempts) {
    /* Non-final wrong — per-SELECTED-option feedback only. Every option
       the learner actually submitted gets a verdict; a correct option
       they did NOT select stays neutral and unrevealed. */
    s6mcOpts().forEach(o => {
      o.classList.remove('correct', 'wrong');
      if (!s6mcSelected.has(o.dataset.id)) return;
      o.classList.remove('selected');
      o.classList.add(S6MC.correct.has(o.dataset.id) ? 'correct' : 'wrong');
    });
    s6mcLastWrong = s6mcSig(s6mcSelected);
    document.getElementById('s6mc-check').disabled = true;
    S6MC_POPUP.retry.title = [...s6mcSelected].some(id => S6MC.correct.has(id)) ? RETRY_TITLE_PARTIAL : RETRY_TITLE;
    s6mcShowPopup('retry');
    return;
  }

  /* Final wrong — reveal the FULL correct set (including options never
     selected) and mark the learner's incorrect selections red. */
  s6mcOpts().forEach(o => {
    o.classList.remove('selected', 'correct', 'wrong');
    if (S6MC.correct.has(o.dataset.id)) o.classList.add('correct');
    else if (s6mcSelected.has(o.dataset.id)) o.classList.add('wrong');
  });
  s6mcResult = 'wrong2';
  s6mcFinish();
  s6mcShowPopup('wrong2');
}

function s6mcFinish() {
  s6mcDone = true;
  s6mcOpts().forEach(o => { o.disabled = true; });
  const chk = document.getElementById('s6mc-check');
  /* MultipleChoiceQuestion inherits SingleChoiceQuestion's completion
     fallback "هل نتابع؟" — slide 57 authors no progression copy of its
     own. NOT ValueInputQuestion's "متابعة". The button stays DISABLED
     here when the outcome is scroll-gated; s6mcCheckScrollGate() is
     what releases it. */
  if (chk) setNavLabel(chk, 'هل نتابع؟');
  hideHintButton('s6mc-hint');
  /* Question 4 is resolved by this screen alone, so the
     progress model is re-derived here too. This screen carries no
     progress bar of its own (its slide authors none), so nothing is
     rendered — only the shared state is brought up to date. */
  syncPracticeProgressNav(document.getElementById('s6'));

}

function s6mcResetInitial() {
  s6mcSelected  = new Set();
  s6mcAttempts  = 0;
  s6mcDone      = false;
  s6mcResult    = null;
  s6mcLastWrong = null;
  s6mcGateOk    = true;
  s6mcGateKey   = null;
  s6mcGateRead  = null;
  s6mcOpts().forEach(o => {
    o.disabled = false;
    o.classList.remove('selected', 'correct', 'wrong');
    o.setAttribute('aria-checked', 'false');
  });
  const popup = document.getElementById('s6mc-popup');
  popup.classList.add('hidden');
  popup.classList.remove('is-scroll', 'needs-scroll');
  document.getElementById('s6mc-hint-overlay').classList.add('hidden');
  const chk = document.getElementById('s6mc-check');
  if (chk) { setNavLabel(chk, 'هل إجابتي صحيحة؟'); chk.disabled = true; }
  showHintButton('s6mc-hint');
}

function s6mcShowPopup(type) {
  const popup = document.getElementById('s6mc-popup');
  if (!popup) return;
  const cfg      = S6MC_POPUP[type];
  const body     = document.getElementById('s6mc-popup-body');
  const scroller = document.getElementById('s6mc-popup-scroll');
  const chk      = document.getElementById('s6mc-check');

  popup.style.background = cfg.bg;
  if (popup.classList.contains('hidden')) resetPopupPosition(popup);

  /* Terminal-incorrect title carries the canonical fixed second line
     (_question-template-defaults.md) — appended to, never replacing,
     the authored title; only the outcome that reveals the answer. */
  const titleEl = document.getElementById('s6mc-popup-title');
  titleEl.innerHTML = cfg.titleSecondLine
    ? (cfg.title + '<br>' + cfg.titleSecondLine)
    : cfg.title;

  body.innerHTML = cfg.body.map(p => {
    if (p === '') return '<p class="scq-gap" aria-hidden="true"></p>';
    return '<p>' + p + '</p>';
  }).join('');
  renderPopupAction('s6mc-popup', null, null);   // this template carries no action row

  /* Scroll gate — arm before showing, then reset and measure AFTER the
     popup is visible (a display:none element ignores a scrollTop write
     and reports 0 for scrollHeight/clientHeight). */
  popup.classList.toggle('is-scroll', !!cfg.scrollGate);
  s6mcGateKey = cfg.scrollGate ? (type) : null;
  /* Already read in THIS state → stay satisfied across a close/reopen. */
  s6mcGateOk = !cfg.scrollGate || s6mcGateRead === s6mcGateKey;
  if (chk && !s6mcGateOk) chk.disabled = true;
  if (cfg.scrollGate) scroller.setAttribute('tabindex', '0');
  else                scroller.removeAttribute('tabindex');

  popup.classList.remove('hidden');
  clampPopup(popup);
  scroller.scrollTop = 0;
  if (cfg.scrollGate) {
    placeScrollHint('s6mc-popup');
    s6mcCheckScrollGate();     // unlocks immediately when the content does not actually overflow
  } else {
    popup.classList.remove('needs-scroll');
    if (chk && s6mcDone) chk.disabled = false;
  }
}

function s6mcCheckScrollGate() {
  if (s6mcGateOk) {
    /* Already satisfied — make sure the hint bubble is not left over
       from a previous, still-locked presentation of this popup. */
    const p0 = document.getElementById('s6mc-popup');
    if (p0) p0.classList.remove('needs-scroll');
    return;
  }
  const scroller = document.getElementById('s6mc-popup-scroll');
  const popup    = document.getElementById('s6mc-popup');
  if (!scroller || !popup) return;
  if (!atBottom(scroller)) { popup.classList.add('needs-scroll'); return; }
  s6mcGateOk = true;
  s6mcGateRead = s6mcGateKey;      // this state has now been read to the end
  const chk = document.getElementById('s6mc-check');
  if (chk && s6mcDone) chk.disabled = false;
  popup.classList.remove('needs-scroll');
}

let s6mcResizeRaf = null;
window.addEventListener('resize', () => {
  const popup = document.getElementById('s6mc-popup');
  if (!popup || popup.classList.contains('hidden')) return;
  placeScrollHint('s6mc-popup');
  if (s6mcGateOk) return;
  if (s6mcResizeRaf) cancelAnimationFrame(s6mcResizeRaf);
  s6mcResizeRaf = requestAnimationFrame(s6mcCheckScrollGate);
});

function s6mcOpenHint()  { xapiRequestedHint('004', 'q1'); document.getElementById('s6mc-hint-overlay').classList.remove('hidden'); }
function s6mcCloseHint() { document.getElementById('s6mc-hint-overlay').classList.add('hidden'); }
function s6mcCloseHintOnBackdrop(e) { if (e.target === e.currentTarget) s6mcCloseHint(); }

/* ═══════════════════════════════════════════════════════════
   S8 — SingleChoiceQuestion, NO-POPUP VARIANT (slide 59)
   The riddle that opens the golden-ratio sequence: three pictures, one
   question, three options, and a verdict shown on the click itself.

   Spec: 720-templates → SingleChoiceQuestion.md, with the four
   departures slide 59's own note authors explicitly (no feedback popup,
   one attempt, no "هل إجابتي صحيحة؟" button, Continue gated on having answered).
   Those are producer instructions — the top of the source hierarchy —
   not inferences from neighbouring slides, and each is quoted in the
   #s8 markup comment in index.html.

   CORRECT = 'o2', taken from the slide's own drawing: id=7 is the only
   capsule filled with the marked tint (accent3 lumMod 20%/lumOff 80%)
   and the only one carrying the check-mark graphic id=19. It is also
   the only option the unit's content supports — slides 60-61 go on to
   name the golden ratio. It is the MIDDLE option in DOM order, which
   RTL renders as the middle capsule, and it stays the correct answer.

   ITS WORDING CHANGED (content correction, 2026-09-14). The slide's
   original — "מסתתר בהן יחס מתמטי מיוחד", the ratio is HIDDEN in them —
   asserted that the proportion is really present in all three pictures.
   For the galaxy and the Mona Lisa there is no evidence of that: shapes
   resembling the golden ratio can be seen in them, but neither was made
   according to it. The approved replacement, "אפשר להמחיש בעזרתן יחס
   מתמטי מיוחד" — they can be used to ILLUSTRATE it — is true of a
   visual comparison and does not claim more than one. The same
   correction runs through S9's closing line and S10's conclusion, and
   is why the third picture is now Le Corbusier's National Museum of
   Western Art in Tokyo: the one example where the ratio was used
   deliberately (through the Modulor).

   ONE ATTEMPT, so there is no attempts counter, no retry gate and no
   last-wrong signature to compare against: the first click resolves the
   screen for good. s8scResult records 'correct' | 'wrong' for the
   same reason every other screen here records one — so a later results
   surface reads the outcome from the screen's own state. This screen is
   NOT part of the four-question progress series (slide 59 carries no
   "שאלה N מתוך" marker), so nothing here touches practiceProgress.

   IMAGE CREDITS — CLOSED (re-audited 2026-09-23 against the shipped
   files). The note's three source URLs name pictures the producer has
   since replaced, so they match nothing shipped; docs/asset-rights.md
   keeps them for traceability only. Of the three pictures that DO ship,
   only the Tokyo museum photograph (CC BY-SA 4.0) requires attribution,
   and its credit control is rendered beside it here and on S10; the
   Mona Lisa (Public Domain) and M51 (Unsplash License) require none.

   MOTION — "תוכן עולה בהדרגה" (content appears gradually) IS WIRED
   through the animation-system engine, in index.html: the three-picture
   row (.s8-images) is reveal group 1 and the question column
   (.scq-question — riddle + question + options as ONE unit) is group 2.
   Motion.reveal() plays them from goTo() on the screen's first entry per
   session; nothing in this file drives the timing, and no ad-hoc CSS is
   involved.
   ═══════════════════════════════════════════════════════════ */
const S8SC = { correctId: 'o2' };

let s8scSelected = null;   // the option the learner actually clicked
let s8scDone     = false;
let s8scResult   = null;   // 'correct' | 'wrong'

function s8scOpts()  { return document.querySelectorAll('#s8 .scq-opt'); }

/* Single owner of the resolved rendering, so entry and first answer
   paint identically. Per slide 59's note: the correct option is marked,
   and the chosen one is marked wrong when it was wrong. Every option is
   locked — one attempt. */
function s8scApply() {
  s8scOpts().forEach(o => {
    const id = o.dataset.id;
    o.classList.remove('selected', 'correct', 'wrong');
    if (id === S8SC.correctId) o.classList.add('correct');
    else if (id === s8scSelected) o.classList.add('wrong');
    o.setAttribute('aria-checked', id === s8scSelected ? 'true' : 'false');
    o.disabled = true;
  });
  const chk = document.getElementById('s8sc-check');
  if (chk) chk.disabled = false;
}

/* The click IS the submission: no select-then-check step exists on this
   screen, because slide 59 authors no "هل إجابتي صحيحة؟" button. */
function s8scAnswer(id) {
  if (s8scDone) return;
  s8scSelected = id;
  s8scResult   = (id === S8SC.correctId) ? 'correct' : 'wrong';
  s8scDone     = true;
  reportQ(['005', 'q1'], id === S8SC.correctId, true, xapiChoiceAnswer('005', 'q1', '#s8 .scq-opt', id));
  s8scApply();
}

function s8scResetInitial() {
  s8scOpts().forEach(o => {
    o.classList.remove('selected', 'correct', 'wrong');
    o.setAttribute('aria-checked', 'false');
    o.disabled = false;
  });
  const chk = document.getElementById('s8sc-check');
  if (chk) chk.disabled = true;
}

/* ═══════════════════════════════════════════════════════════
   S10 — the authored three-image sequence (slide 61)
   "התמונות מופיעות בהדרגה – תמונה על כל המסך ואז קטנה ומשתבצת למקום
   שלה בדף, ככה ל-3 התמונות. אחרי התמונות יופיע המשפט."

   A producer-approved, SCREEN-SPECIFIC exception to the shared motion
   vocabulary (which allows fade and float only). It lives here, scoped
   to #s10; js/motion.js and css/motion.css are the canonical engine,
   byte-for-byte, and nothing below touches them or adds scale to the
   shared language. The engine's own tokens and its m-float-in keyframe
   are reused for the closing sentence so the timing still belongs to
   one system.

   FLIP, not a second layout. Each picture already sits in its final
   authored frame; the "full-stage" state is a transform derived from
   that frame, so the aspect ratio and the slide's srcRect crop are
   carried along untouched and nothing is measured twice.

   THE STAGE is the usable content area only — below the report flag
   (which ends at y 81.8) and above the 74px bottom bar (top edge at
   y 636). A small inset keeps a settling picture off both edges. The
   flag, Continue, Back and the DEV harness are outside #s10's content
   and are never moved or covered.
   ═══════════════════════════════════════════════════════════ */
const S10_STAGE = { left: 40, top: 96, right: 1240, bottom: 624 };
const S10_STEP  = 620;   /* one picture's settle, ms */
const S10_GAP   = 120;   /* pause between pictures, ms */

let s10Seen = false;     /* first visit per session — same rule the engine uses */
let s10Timers = [];

function s10ReducedMotion() {
  return window.matchMedia &&
         window.matchMedia('(prefers-reduced-motion: reduce)').matches;
}

/* Full reset: clears pending timers and every inline style the sequence
   writes, so leaving mid-sequence can never strand a picture scaled up,
   shifted or invisible. Safe to call at any time, including before the
   first run. */
function s10Cancel() {
  s10Timers.forEach(t => clearTimeout(t));
  s10Timers = [];
  document.querySelectorAll('#s10 .s10-img').forEach(el => {
    el.classList.remove('s10-anim', 's10-pending');
    el.style.transform = '';
    el.style.opacity = '';
  });
  /* The credit control rides with the conclusion: it is chrome for the
     picture, not part of it, so it must not be left hanging over a
     full-stage photograph while that photograph is still shrinking. */
  document.querySelectorAll('#s10 .s10-line, #s10 .s10-credit').forEach(el => {
    el.classList.remove('s10-pending', 's10-enter');
  });
}

/* The transform that maps a picture's final frame onto the stage,
   in DESIGN pixels: offsetLeft/Width are pre-transform layout values, so
   they are unaffected by scaleApp()'s scaling of the whole app. */
function s10StageTransform(el) {
  const fw = el.offsetWidth, fh = el.offsetHeight;
  /* Sum the offsets up to #s10 rather than reading offsetLeft alone:
     the centring parent .s10-stage sits between the picture and the
     screen, so the picture's own offsetLeft is relative to the CLUSTER,
     while S10_STAGE is expressed in screen coordinates. Walking the
     offsetParent chain keeps the two in the same frame of reference
     without hardcoding the centring offset in two places. */
  const sec = document.getElementById('s10');
  let fl = 0, ft = 0;
  for (let n = el; n && n !== sec; n = n.offsetParent) { fl += n.offsetLeft; ft += n.offsetTop; }
  if (!fw || !fh) return null;
  const sw = S10_STAGE.right - S10_STAGE.left;
  const sh = S10_STAGE.bottom - S10_STAGE.top;
  /* contain: the picture fills the stage on its tighter axis, so its
     aspect ratio is preserved exactly. */
  const k  = Math.min(sw / fw, sh / fh);
  /* S10_STAGE is written for the 1280-wide design canvas; the settled
     composition centres on the LIVE canvas (left: 50%), so the stage moves
     with it — otherwise on a wider canvas each picture would start from a
     stage 640px off the middle of where it lands. */
  const stageLeft = S10_STAGE.left + (getCanvasSize().w - 1280) / 2;
  const dx = (stageLeft + sw / 2) - (fl + fw / 2);
  const dy = (S10_STAGE.top  + sh / 2) - (ft + fh / 2);
  return 'translate(' + dx.toFixed(2) + 'px, ' + dy.toFixed(2) + 'px) scale(' + k.toFixed(4) + ')';
}

/* Runs on entry into S10. First visit only; a revisit — and reduced
   motion — leave the settled composition exactly as the stylesheet
   renders it, with no inline styles at all. */
function s10Reveal() {
  const sec = document.getElementById('s10');
  if (!sec) return;
  s10Cancel();
  if (s10Seen || s10ReducedMotion()) return;
  s10Seen = true;

  const imgs = [...sec.querySelectorAll('.s10-img')];
  const line = sec.querySelector('.s10-line');
  const credit = sec.querySelector('.s10-credit');
  if (!imgs.length) return;

  /* Hide what has not had its turn yet. Continue stays enabled
     throughout — the sequence gates nothing. */
  imgs.forEach(el => el.classList.add('s10-pending'));
  if (line) line.classList.add('s10-pending');
  if (credit) credit.classList.add('s10-pending');

  imgs.forEach((el, i) => {
    const at = i * (S10_STEP + S10_GAP);
    s10Timers.push(setTimeout(() => {
      const from = s10StageTransform(el);
      if (!from) { el.classList.remove('s10-pending'); return; }
      /* Appear at full-stage size... */
      el.style.transform = from;
      el.classList.remove('s10-pending');
      el.style.opacity = '1';
      /* ...then, on the next frame, let the transition carry it home.
         Two rAFs so the from-state is guaranteed to have been painted
         before the property changes. */
      requestAnimationFrame(() => requestAnimationFrame(() => {
        el.classList.add('s10-anim');
        el.style.transform = 'none';
      }));
      /* Drop the inline styles once it has arrived, so the settled
         picture is styled purely by the stylesheet again. */
      s10Timers.push(setTimeout(() => {
        el.classList.remove('s10-anim');
        el.style.transform = '';
        el.style.opacity = '';
      }, S10_STEP + 60));
    }, at));
  });

  /* The sentence only after all three have settled. */
  if (line) {
    const after = imgs.length * (S10_STEP + S10_GAP) + 40;
    s10Timers.push(setTimeout(() => {
      line.classList.remove('s10-pending');
      line.classList.add('s10-enter');
      if (credit) { credit.classList.remove('s10-pending'); credit.classList.add('s10-enter'); }
      s10Timers.push(setTimeout(() => {
        line.classList.remove('s10-enter');
        if (credit) credit.classList.remove('s10-enter');
      }, 900));
    }, after));
  }
}

/* ═══════════════════════════════════════════════════════════
   IMAGE ZOOM — canonical clone-based modal, one per part.
   720-templates → _global-components.md → "Image zoom" (markup, DOM,
   clone behaviour) + _question-template-defaults.md → "Image zoom"
   (REQUIRED for any question that shows a content image). Ported
   verbatim from the approved sibling methodica-math-ratio-03-02, whose
   producer confirmed the zoom must visibly ENLARGE: the media is scaled
   to the surface's available viewport, not clamped at its intrinsic
   size. Fully delegated — the trigger carries data-zoom-src, the
   backdrop and the X carry data-zoom-close — so no screen needs inline
   handlers for it. Distinct from the applet zoom above, which promotes
   a live iframe card in place and must never clone.
   ═══════════════════════════════════════════════════════════ */
let zoomReturnFocus = null;
/* Monotonic, so two opens in one session can never mint the same
   cloned id. */
let zoomCloneSeq = 0;

function openImageZoom(btn) {
  const wrapper = btn.closest('.viq-img-frame');
  const stage   = document.getElementById('img-zoom-stage');
  const modal   = document.getElementById('img-zoom-modal');
  const panel   = modal && modal.querySelector('.img-zoom-panel');
  if (!wrapper || !stage || !modal || !panel) return;

  const srcImg = wrapper.querySelector('img');
  /* Clone the WRAPPER, not the bare file, so the picture keeps its own
     structure; drop the trigger from the clone. */
  const clone = wrapper.cloneNode(true);
  clone.removeAttribute('id');
  /* The modal must never carry a second trigger — the clone's own zoom
     button is removed before it is mounted. */
  const cloneBtn = clone.querySelector('.img-zoom-btn');
  if (cloneBtn) cloneBtn.remove();
  /* IDS. Two ids can never coexist in one document, so the clone's are
     re-issued rather than simply dropped: a live SVG diagram refers to
     its own <filter> by url(#id), and stripping the id would leave that
     reference dangling and silently lose the drop shadow in the modal.
     Each id gets a unique suffix and every url(#…) / href="#…" inside
     THIS clone is repointed to it. A frame with no internal references
     (a plain raster, like S6's) is unaffected: its ids are renamed,
     nothing points at them, and the document still holds no duplicate. */
  const idStamp = '-zoom' + (++zoomCloneSeq);
  clone.querySelectorAll('[id]').forEach(el => {
    const was = el.getAttribute('id');
    const now = was + idStamp;
    el.setAttribute('id', now);
    clone.querySelectorAll('*').forEach(ref => {
      for (const at of Array.from(ref.attributes)) {
        if (at.value === 'url(#' + was + ')') ref.setAttribute(at.name, 'url(#' + now + ')');
        else if (at.value === '#' + was) ref.setAttribute(at.name, '#' + now);
      }
    });
  });
  stage.innerHTML = '';
  stage.appendChild(clone);

  modal.classList.remove('hidden');
  modal.setAttribute('aria-hidden', 'false');

  const cs   = getComputedStyle(panel);
  const boxW = 1200 - parseFloat(cs.paddingLeft) - parseFloat(cs.paddingRight);
  const boxH = 640  - parseFloat(cs.paddingTop)  - parseFloat(cs.paddingBottom);
  /* TWO KINDS OF FRAME, one fit rule: scale to the surface's available
     viewport preserving the source's own ratio, whichever axis binds
     first, and NOT clamped at 1 — a source smaller than the viewport is
     scaled UP, which is the producer's own requirement that the zoom
     must visibly enlarge. */
  if (srcImg && srcImg.naturalWidth && srcImg.naturalHeight) {
    /* Raster: the <img> reports its intrinsic size and object-fit does
       the fitting inside the sized stage. */
    const scale = Math.min(boxW / srcImg.naturalWidth, boxH / srcImg.naturalHeight);
    stage.style.width  = Math.round(srcImg.naturalWidth  * scale) + 'px';
    stage.style.height = Math.round(srcImg.naturalHeight * scale) + 'px';
  } else if (clone.querySelector('.scq-diagram-art')) {
    /* LIVE SVG DIAGRAM (S1, S2). There is no <img> to measure, and
       object-fit cannot be used: the vertex labels are real text
       positioned in AUTHORED PIXELS over the drawing, so stretching the
       box would move every letter off its vertex. The whole frame is
       therefore scaled as ONE unit by a uniform transform — art and
       labels together, ratio exact by construction, no distortion
       possible. The frame's own authored box is the source size
       (offsetWidth/Height are layout px, so #app's canvas transform
       does not leak into them). */
    const w = wrapper.offsetWidth, h = wrapper.offsetHeight;
    if (w && h) {
      const scale = Math.min(boxW / w, boxH / h);
      clone.style.width  = w + 'px';
      clone.style.height = h + 'px';
      clone.style.transform = 'scale(' + scale + ')';
      stage.style.width  = Math.round(w * scale) + 'px';
      stage.style.height = Math.round(h * scale) + 'px';
    }
  }
  zoomReturnFocus = btn;
  /* FOCUS MOVE IS DEFERRED — measured, not stylistic. Activating the
     trigger with ENTER used to open the modal and shut it again in the
     same keystroke: Chromium's Enter-on-button default action was still
     in flight when this handler moved focus to the Close control, so a
     SECOND click landed on Close — which carries data-zoom-close — and
     closed what the first click had just opened. The instrumented trace
     read:
         keydown → keypress → click(trigger) → modal OPEN
                            → click(Close)   → modal HIDDEN → keyup
     Space never showed it, because a button activates on Space's KEYUP,
     so no default action remains to be delivered afterwards.
     Deferring to a task lets the key event finish before focus moves.
     The learner cannot act inside that gap, so "focus moves to the
     modal's close control" still holds exactly.
     This is the SHARED component, so the same fix covers S6's raster
     trigger, which had the identical defect. */
  const closeBtn = document.getElementById('img-zoom-close');
  if (closeBtn) closeBtn.focus();
}
function closeImageZoom() {
  const modal = document.getElementById('img-zoom-modal');
  const stage = document.getElementById('img-zoom-stage');
  if (modal) { modal.classList.add('hidden'); modal.setAttribute('aria-hidden', 'true'); }
  if (stage) { stage.innerHTML = ''; stage.style.width = ''; stage.style.height = ''; }
  /* Focus restoration — back to the trigger, so a keyboard learner is
     never dropped at the top of the page. */
  if (zoomReturnFocus && document.contains(zoomReturnFocus)) zoomReturnFocus.focus();
  zoomReturnFocus = null;
}
document.addEventListener('click', e => {
  /* Two trigger markers, one component: data-zoom-src names a raster
     frame, data-zoom-diagram a live-SVG one. The marker is what makes a
     button a trigger; openImageZoom() reads neither value, because it
     clones the FRAME rather than re-fetching a file. */
  const zoomBtn = e.target.closest && e.target.closest('[data-zoom-src], [data-zoom-diagram]');
  if (zoomBtn) { openImageZoom(zoomBtn); return; }
  if (e.target.closest && e.target.closest('[data-zoom-close="true"]')) closeImageZoom();
});
document.addEventListener('keydown', e => {
  if (e.key !== 'Escape') return;
  const modal = document.getElementById('img-zoom-modal');
  if (modal && !modal.classList.contains('hidden')) closeImageZoom();
});
/* A screen change must never leave an enlarged image over the next
   screen — goTo() sweeps popups and hint overlays by id suffix, which
   this modal deliberately does not match, so it is closed explicitly. */
function closeImageZoomOnNavigate() {
  const modal = document.getElementById('img-zoom-modal');
  if (modal && !modal.classList.contains('hidden')) closeImageZoom();
}

/* ═══════════════════════════════════════════════════════════
   APPLET / IFRAME ZOOM — figma-lomda-builder → reference/patterns.md →
   "Zoom modal". Ported from the approved sibling
   methodica-math-ratio-03-02 (its S6/S7 applet screens); only the
   enlarged target size below is this part's own, and this part's
   applet carries no scale transform.

   Added on producer request (2026-09-10). Slide 54 draws no zoom
   affordance of its own, so this is a documented addition, not a
   fidelity claim.

   NEITHER move NOR clone the iframe. The applet is stateful — the
   learner sets a slider and toggles area/perimeter — and an iframe
   loses its document the moment it is re-parented. So this zoom
   PROMOTES THE CARD IN PLACE: a class plus an inline rect move and
   resize the card onto the modal surface while the card keeps its DOM
   position and its iframe child untouched. The applet is never
   reloaded.

   THE ENLARGED SIZE IS MEASURED, NOT GUESSED. This widget lays itself
   out from its own viewport (a flex column with vh-based clamps), and
   its iframe is 100%/100% of the card with no transform, so enlarging
   the card makes it genuinely RE-LAY-OUT rather than magnify. Sweeping
   boxes in the live page gave:
       334×476 (inline)  invitation 121px · legend 2 rows · buttons 2 lines
       380×552           invitation 177px · legend 2 rows · buttons 1 line
       440×552           invitation 197px · legend 1 row  · buttons 1 line
       500..720×552      identical to 440 — the stage is height-bound,
                         so extra WIDTH never grows the invitation
       580..620×592      invitation 213px at the default slider
   440×552 was the first shipped size: the smallest box in which the
   widget composes fully. It was ENLARGED to 620×592 on 2026-09-10 at
   the producer's request (marked-up screenshot), and the extra width
   is load-bearing rather than cosmetic — see below.

   WHY THE CARD IS NOW WIDER THAN THE INVITATION NEEDS. The close
   control sits inside the card at the library's 16/16 top-left inset,
   i.e. x 17..61. The invitation is square and bounded by the stage's
   HEIGHT, and it is centred in the stage's WIDTH — so widening the
   card pushes the artwork away from that corner without changing its
   size. Measured at the top of the widget's own slider (picSide 7,
   where the invitation is at its largest and fills the stage):
       440×552   invitation 341px, its left edge 18px inside the card
                 → directly under the close button (the reported bug)
       580×592   invitation 378px, left edge 69px inside the card
       620×592   invitation 378px, left edge ~89px inside the card
   620 is chosen for that clearance margin, not for the applet. The
   resulting surface is 668×680 (card + the panel's 24/44 padding),
   which is also roughly the enlarged frame the producer marked.
   ═══════════════════════════════════════════════════════════ */
const APPLET_ZOOM_TARGET = { w: 620, h: 592 };

/* Which screen's card is currently promoted — close, Escape, the
   backdrop and the iframe's own Escape bridge all resolve through it. */
let appletZoomPrefix = null;

/* Container-driven resize must not animate (producer, 2026-09-22).
   The applet lays itself out from its own viewport and animates its
   invitation with a width/height transition, which is right when the
   LEARNER moves the slider and wrong when the CARD changes size: opening
   or closing the zoom made the picture visibly grow or shrink by itself.
   The widget cannot tell the two apart from the inside — its own
   ResizeObserver sees only "the box changed" — so the host, which knows
   exactly when a container change happens, announces it and the widget
   freezes its own transitions until the new size has landed (the receive
   side is the DCX_APPLET_CONTAINER_RESIZE script in
   assets/widgets/מסיבת הפתעה.html).
   Why a message and not a stylesheet injected from here: the lomda is
   opened from file://, where the iframe has the opaque origin "null" and
   frame.contentDocument is null — an earlier version of this helper did
   exactly that and silently did nothing in the real browser (the same
   reason the Escape bridge below is a postMessage). Target origin "*"
   because a file:// child cannot be named; the child authenticates by
   event.source === window.parent. Called BEFORE the card's geometry is
   written so the freeze is in place when the iframe's resize lands; even
   if the message arrived a frame late, transition:none on a running
   transition snaps it to its end value, so the worst case is one frame. */
function freezeAppletTransitions(prefix) {
  const frame = prefix && document.getElementById(prefix + '-applet');
  const win = frame && frame.contentWindow;
  if (!win) return;
  try { win.postMessage({ type: 'DCX_APPLET_CONTAINER_RESIZE' }, '*'); } catch (e) { /* not loaded yet */ }
}

function openAppletZoom(prefix) {
  prefix = prefix || appletZoomPrefix || 's3';
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
     card, because inline the card is ordinary screen content and must
     not announce itself as a dialog. */
  card.setAttribute('role', 'dialog');
  card.setAttribute('aria-modal', 'true');
  card.setAttribute('aria-label', 'تكبير التطبيق');

  const cs = getComputedStyle(panel);
  const padL = parseFloat(cs.paddingLeft), padR = parseFloat(cs.paddingRight);
  const padT = parseFloat(cs.paddingTop),  padB = parseFloat(cs.paddingBottom);

  /* Fit the measured target into the envelope's media area, then shrink
     the surface to what the applet actually occupies — the same
     media-responsive rule the image-zoom surface follows, so no white
     slack is left around the enlarged applet. */
  panel.style.width = ''; panel.style.height = '';       // measure the envelope
  const availW = panel.clientWidth  - padL - padR;
  const availH = panel.clientHeight - padT - padB;
  const fit = Math.min(availW / APPLET_ZOOM_TARGET.w, availH / APPLET_ZOOM_TARGET.h);
  const w = Math.round(APPLET_ZOOM_TARGET.w * fit);
  const h = Math.round(APPLET_ZOOM_TARGET.h * fit);
  panel.style.width  = (w + padL + padR) + 'px';
  panel.style.height = (h + padT + padB) + 'px';

  /* Place the card on the panel's own content box, MEASURED rather than
     hardcoded. Both the modal and the card are absolutely positioned
     against boxes that are inset:0 of #app, so their offset coordinates
     are already canvas coordinates and no scale maths is needed — and
     because the panel is flex-centred, this follows it when scaleApp()
     grows the canvas past 710 on a surplus viewport. */
  freezeAppletTransitions(prefix);
  card.style.left   = (panel.offsetLeft + padL) + 'px';
  card.style.top    = (panel.offsetTop  + padT) + 'px';
  card.style.width  = w + 'px';
  card.style.height = h + 'px';

  /* Each card carries its OWN close control (inside the card, at the
     enlarged frame's physical top-left), so resolve it within the
     promoted card rather than by a single global id. */
  const closeBtn = card.querySelector('.viq-applet-close');
  if (closeBtn) closeBtn.focus();
}

function closeAppletZoom() {
  const modal  = document.getElementById('applet-zoom-modal');
  const prefix = appletZoomPrefix;
  const card   = prefix && document.getElementById(prefix + '-applet-frame');
  if (!modal) return;
  freezeAppletTransitions(prefix);
  if (card) {
    card.classList.remove('is-zoomed');
    card.style.left = card.style.top = card.style.width = card.style.height = '';
    card.removeAttribute('role');                  // ordinary content again
    card.removeAttribute('aria-modal');
    card.removeAttribute('aria-label');
  }
  const panel = modal.querySelector('.applet-zoom-panel');
  if (panel) { panel.style.width = ''; panel.style.height = ''; }
  modal.classList.add('hidden');
  modal.setAttribute('aria-hidden', 'true');
  /* Focus restoration — back to the trigger, so a keyboard learner is
     never dropped at the top of the page. */
  const trigger = prefix && document.getElementById(prefix + '-applet-zoom');
  if (trigger && document.contains(trigger)) trigger.focus();
  appletZoomPrefix = null;
}

/* Close is fully delegated: the backdrop and the in-card X both carry
   data-applet-zoom-close="true", so neither needs an inline handler. */
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
   forwards it by postMessage instead (the same three-line bridge the
   sibling's widget carries, added to this part's own copy of
   assets/widgets/מסיבת הפתעה.html).
   Authenticated by SOURCE, not by origin: a file:// iframe reports the
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
/* A screen change must never leave an enlarged applet promoted over the
   next screen — goTo() sweeps popups and hint overlays by id suffix,
   which this modal deliberately does not match (it is not a feedback
   popup), so it is closed here instead. */
function closeAppletZoomOnNavigate() {
  const modal = document.getElementById('applet-zoom-modal');
  if (modal && !modal.classList.contains('hidden')) closeAppletZoom();
}
/* Cross-part navigation. init() must run AFTER this part's own initial
   screen setup: if this page was entered via Back, it jumps straight to
   the screen the learner left from. */
PartNav.init({
  part: '04',
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
   Ids must match metadata/methodica-math-ratio-03-04.json byte-for-byte (the loader
   checks XAPI_COMP_ID against METADATA.id on every load; _test/verify-static.js checks the map).
   ═══════════════════════════════════════════════════════════ */
const XAPI_COMP_SLUG     = 'methodica-math-ratio-03-04';
const XAPI_COMP_ID       = XAPI_ID_PREFIX + XAPI_COMP_SLUG + '/';
const XAPI_METADATA_FILE = '../metadata/' + XAPI_COMP_SLUG + '.json';
/* Items carrying a code-graded question (each declared question is reported by its screen). */
const XAPI_EVAL_ITEMS    = { '001': 1, '002': 1, '003': 1, '004': 1, '005': 1 };
/* Items scored as ONE question: success only when every section is right (QA F-6, 22-scoring.js).
   q3Resolved: two correct sections and one incorrect = ONE INCORRECT QUESTION. */
const XAPI_ITEM_ALL_CORRECT = { '003': 1 };
/* screen → [item suffix, page within the item]; null = no catalogue item. Exactly TOTAL_SCREENS
   keys, 0-based (keep it a literal object — the static check reads it). */
const SCREEN_TO_SUBCONTENT = {
  0: null,
  1: ['001', 1],
  2: ['002', 1],
  3: ['003', 1],
  4: ['003', 2],
  5: ['003', 3],
  6: ['004', 1],
  7: ['005', 1],
  8: ['005', 2],
  9: ['005', 3],
  10: ['005', 4],
  11: ['005', 5]
};

/* No gate on screen here ("4 advanced questions"). The score is the share of the 7 questions the
   metadata declares (partScaled: unanswered counts as wrong).
   MOE 2026-10-08: real score, success only ≥60%, no gate, nothing shown. */
function partGateResult() { const sc = partScaled(); return sc == null ? { success: true } : { success: sc >= 0.6, score: { scaled: sc } }; }
function lastScreenButton() { return currentScreen === PART_LAST ? document.querySelector('#s11 .btn-continue') : null; }

/* ═══════════════════════════════════════════════════════════
   Resume painters — unit-js/45-painters.js reads this registry.
   What each screen saves (vars + the fields that hold state) and how a screen restored from the
   state document is drawn again on the fresh markup: init() is the screen's own ResetInitial,
   paint() replays the template's own terminal look (Mark / Finish / ShowPopup / ApplyView) or the
   retry look, without counting an attempt and without reporting. In-session re-entry still goes
   through resetScreenState(), which now resets only a pristine screen.
   Scroll gates are state (guide §8.8 rule 11): every screen saves its GateKey / GateRead / GateOk,
   and ShowPopup re-derives the gate from them, so feedback already read stays read and feedback
   not yet read keeps the button locked. The progress model (practiceProgress /
   lomdaState.practiceQuestions) is derived from the screens' own results by
   syncPracticeProgressNav on every entry, so it is not saved. S10's s10Seen is animation only.
   ═══════════════════════════════════════════════════════════ */
const PART_STATE_VARS = [];
const PART_SCREENS = {
  1: {
    vars: ['s1scSelected', 's1scAttempts', 's1scDone', 's1scResult', 's1scLastWrong',
           's1scGateKey', 's1scGateRead', 's1scGateOk'],
    pristine: () => !s1scSelected && !s1scAttempts && !s1scDone,
    init: s1scResetInitial,
    paint: () => {
      const done = s1scDone, sel = s1scSelected;
      s1scDone = false;
      if (sel) s1scSelect(sel);                    // selection look + the retry gate
      if (done) {
        s1scOpt(sel).classList.add(s1scResult === 'correct' ? 'correct' : 'wrong');
        if (s1scResult !== 'correct') s1scOpt(S1SC.correctId).classList.add('correct');
        s1scFinish();
        s1scShowPopup(s1scResult);                 // re-derives the scroll gate from GateRead
      } else if (s1scAttempts && sel === s1scLastWrong) {
        s1scOpt(sel).classList.add('wrong');
      }
    }
  },
  2: {
    vars: ['s2dqSelected', 's2dqAttempts', 's2dqDone', 's2dqResult', 's2dqLastWrong', 's2dqAnswer', 's2dqView',
           's2dqGateKey', 's2dqGateRead', 's2dqGateOk'],
    pristine: () => !s2dqSelected && !s2dqAttempts && !s2dqDone,
    init: s2dqResetInitial,
    paint: () => {
      const done = s2dqDone, sel = s2dqSelected;
      s2dqDone = false;
      if (sel) s2dqSelect(sel);                    // slot label + the retry gate (no retry marks on this template)
      if (done) {
        s2dqFinish();
        if (s2dqAnswer === null) { s2dqMark('correct'); s2dqShowPopup('correct'); }
        else s2dqApplyView();
      }
    }
  },
  3: {
    vars: ['s3viqAttempts', 's3viqDone', 's3viqResult', 's3viqAnswer', 's3viqView', 's3viqLastWrongSig',
           's3viqGateKey', 's3viqGateRead', 's3viqGateOk'],
    inputs: ['s3viq-input-first', 's3viq-input-second'],
    pristine: () => !s3viqAttempts && !s3viqDone && resumeFieldsEmpty(PART_SCREENS[3].inputs),
    init: s3viqResetInitial,
    paint: () => {
      const done = s3viqDone;
      s3viqDone = false;
      s3viqUpdateBar();                            // the button as it stood when Check was pressed
      if (done) {
        if (s3viqAnswer === null) { s3viqMark('correct'); s3viqFinish(); s3viqShowPopup('correct'); }
        else { s3viqFinish(); s3viqApplyView(); }
        return;
      }
      if (s3viqAttempts && s3viqSignature() === s3viqLastWrongSig) s3viqMark('wrong');
    }
  },
  4: {
    vars: ['s4viqAttempts', 's4viqDone', 's4viqResult', 's4viqAnswer', 's4viqView', 's4viqLastWrongSig',
           's4viqGateKey', 's4viqGateRead', 's4viqGateOk'],
    inputs: ['s4viq-input-first', 's4viq-input-second'],
    pristine: () => !s4viqAttempts && !s4viqDone && resumeFieldsEmpty(PART_SCREENS[4].inputs),
    init: s4viqResetInitial,
    paint: () => {
      const done = s4viqDone;
      s4viqDone = false;
      s4viqUpdateBar();
      if (done) {
        if (s4viqAnswer === null) { s4viqMark('correct'); s4viqFinish(); s4viqShowPopup('correct'); }
        else { s4viqFinish(); s4viqApplyView(); }
        return;
      }
      if (s4viqAttempts && s4viqSignature() === s4viqLastWrongSig) s4viqMark('wrong');
    }
  },
  5: {
    vars: ['s5viqAttempts', 's5viqDone', 's5viqResult', 's5viqAnswer', 's5viqView', 's5viqLastWrongSig',
           's5viqGateKey', 's5viqGateRead', 's5viqGateOk'],
    inputs: ['s5viq-input'],
    pristine: () => !s5viqAttempts && !s5viqDone && resumeFieldsEmpty(PART_SCREENS[5].inputs),
    init: s5viqResetInitial,
    paint: () => {
      const done = s5viqDone;
      s5viqDone = false;
      s5viqUpdateBar();
      if (done) {
        if (s5viqAnswer === null) { s5viqMark('correct'); s5viqFinish(); s5viqShowPopup('correct'); }
        else { s5viqFinish(); s5viqApplyView(); }
        return;
      }
      if (s5viqAttempts && s5viqNormalize(s5viqInput().value) === s5viqLastWrongSig) s5viqMark('wrong');
    }
  },
  6: {
    vars: ['s6mcSelected', 's6mcAttempts', 's6mcDone', 's6mcResult', 's6mcLastWrong',
           's6mcGateKey', 's6mcGateRead', 's6mcGateOk'],
    pristine: () => !s6mcSelected.size && !s6mcAttempts && !s6mcDone,
    init: s6mcResetInitial,
    paint: () => {
      const done = s6mcDone;
      s6mcDone = false;
      s6mcRender();                                // selection look + the retry gate
      if (done) {
        s6mcOpts().forEach(o => {
          const id = o.dataset.id;
          o.classList.remove('selected', 'correct', 'wrong');
          if (S6MC.correct.has(id)) o.classList.add('correct');
          else if (s6mcSelected.has(id)) o.classList.add('wrong');
        });
        s6mcFinish();
        s6mcShowPopup(s6mcResult);
      } else if (s6mcAttempts && s6mcSig(s6mcSelected) === s6mcLastWrong) {
        s6mcOpts().forEach(o => {
          if (!s6mcSelected.has(o.dataset.id)) return;
          o.classList.remove('selected');
          o.classList.add(S6MC.correct.has(o.dataset.id) ? 'correct' : 'wrong');
        });
      }
    }
  },
  8: {
    vars: ['s8scSelected', 's8scDone', 's8scResult'],
    pristine: () => !s8scDone,                     // the click is the submission
    init: s8scResetInitial,
    paint: () => { if (s8scDone) s8scApply(); }
  }
};
