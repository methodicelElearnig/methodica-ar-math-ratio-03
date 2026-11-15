'use strict';

/* ═══════════════════════════════════════════════════════════
   js/main.js  —  methodica-math-ratio-03-01
   720 reference engine (canvas 1280×710) + Screen 01
   (TwoOptionSelection — companion-character selection).

   Ported from the approved Math sibling
   methodica-math-percent-02/methodica-math-percent-02-01, per
   720-templates → references/TwoOptionSelection.md +
   references/_global-components.md.

   Screens implemented so far: 01 (TwoOptionSelection), 02
   (TransitionScreen local variant -- hummus-stand briefing), 03
   (FullImageWithTextOverlay C -- embedded hummus-stand widget), 04
   (TransitionScreen local variant -- hook closing text), 05
   (TransitionScreen -- "מה קרה בחומוסייה?", centred text + illustration),
   06 (comic sequence -- six staged units, Continue gated on the reveal),
   07 (TransitionScreen -- "מה אפשר ללמוד מזה?", centred text only),
   08 (FlipCardsReveal -- 3 cards, unlock-on-all-flipped Continue gate).
   ═══════════════════════════════════════════════════════════ */

/* ─── Constants ─────────────────────────────────────────── */
const TOTAL_SCREENS = 24;  // S0-S23 — bump as screens are added

/* ═══════════════════════════════════════════════════════════
   Screen 03 widget completion gate
   ───────────────────────────────────────────────────────────
   Slide 5's presenter note: "המשך מתאפשר רק לאחר סיום המשחק" —
   Continue is enabled only after the game is finished.

   The widget as originally delivered published NO completion event. An
   earlier attempt therefore intercepted the component's own
   `StreamableLogic.prototype.setState` from the parent to watch for its
   terminal phase. That was removed: it is fragile (couples the screen to
   the DC runtime's internals and to same-origin access) and it did not
   survive the real learner path.

   The widget now emits an EXPLICIT contract instead, from the component's
   own `onCallNext` — the single transition that sets phase 'owner' and so
   renders the completion modal (the "כל הכבוד" card carrying "שחקו שוב").
   It is the widget's "המשיכו" button that invokes it, so the message
   fires when the modal APPEARS, not when "שחקו שוב" is clicked:

     onCallNext = () => {
       this.setState({ phase: 'owner' });
       if (!this.__lomdaCompleteSent) {
         window.parent.postMessage({ type: 'lomda:widget-complete', widget: 'hummus-stand' }, '*');
         this.__lomdaCompleteSent = true;
       }
     };

   `componentDidUpdate` keeps its original single responsibility
   (`this.paint()`), and `onRestart` deliberately does NOT clear
   __lomdaCompleteSent — replaying the game must never re-lock Continue.

   targetOrigin is '*', NOT window.location.origin: a document opened from
   the filesystem reports its origin as the STRING "null", which is not a
   valid targetOrigin and makes postMessage throw. '*' weakens nothing
   here because the parent below validates source, type, widget id and
   origin — and the e.source check is what actually establishes trust.

   Only the component-local copy is patched; the producer's original in
   the root יישומונים folder is untouched. If the widget is regenerated,
   this one line has to be re-applied — that is the known maintenance cost
   of the explicit contract, and it is the version that also keeps working
   cross-origin (e.g. once packaged for SCORM), which no parent-side hook
   would.

   Nothing here unlocks on a timer, on iframe load, on an interaction
   count, on scraped text, or by calling the game's internal handlers.
   ═══════════════════════════════════════════════════════════ */
const WIDGET_COMPLETE_CONTRACT = {
  screen: 2,
  iframeId: 's2-widget',
  type: 'lomda:widget-complete',
  widget: 'hummus-stand'
};


/* Resolved flag. Within the same loaded page it survives navigating away
   and back, which is the canonical resolved-state policy for a completed
   screen. With resume on it is also saved in the state document (a gate
   is state — PART_SCREENS[2] at the end of this file), so a learner who
   finished the game and reloads keeps Continue open; the widget itself
   restarts. Only the real completion message ever sets it live. */
let s2WidgetCompleted = false;

function unlockS2Continue() {
  // Idempotent: genuine completion unlocks exactly once, and a repeat
  // signal cannot re-fire side effects.
  if (s2WidgetCompleted) return;
  s2WidgetCompleted = true;
  const btn = document.getElementById('s2-continue');
  if (!btn) { console.error('[s2-gate] #s2-continue not found — cannot unlock'); return; }
  btn.disabled = false;
  btn.removeAttribute('aria-disabled');
}

/* Is `origin` acceptable for a message from our own widget iframe?
   - http(s): it must equal this document's own origin exactly.
   - file:   Chromium is inconsistent about how it serializes file origins.
             A file document's own window.location.origin / window.origin
             can read 'file://', while a file CHILD's MessageEvent.origin
             arrives as the opaque 'null' — and some builds do the reverse.
             So under file: we accept either spelling, and ONLY when this
             document is itself a file: document. The origin string carries
             no trust information under file: anyway; the e.source identity
             check in the listener is what establishes trust. */
function isAcceptableWidgetOrigin(origin) {
  if (window.location.protocol !== 'file:') return origin === window.location.origin;
  return origin === 'null' || /^file:/.test(origin);
}

window.addEventListener('message', function (e) {
  const frame = document.getElementById(WIDGET_COMPLETE_CONTRACT.iframeId);
  const d = e.data;
  if (!frame) return;

  // 1. The message must come from THIS iframe's window — not the dev
  //    harness, not a sibling frame, not the page itself. This is the
  //    load-bearing check and is never relaxed.
  if (e.source !== frame.contentWindow) return;

  const isOurs = d && typeof d === 'object' &&
                 d.type === WIDGET_COMPLETE_CONTRACT.type &&
                 d.widget === WIDGET_COMPLETE_CONTRACT.widget;
  if (!isOurs) return;                    // bundler relay traffic etc.

  // 2. Origin check, evaluated only for our own recognised payload. A
  //    refused completion is always reported: a silently-refused one is
  //    exactly how this gate failed under file:// before.
  if (!isAcceptableWidgetOrigin(e.origin)) {
    console.error('[s2-gate] REJECTED: origin', e.origin,
      'is not acceptable for page origin', window.location.origin);
    return;
  }

  unlockS2Continue();
});

/* ─── Companion character system ─────────────────────────────
   Logical character IDs (never a filename, colour name, asset path or
   DOM id) — _global-components.md → "Companion character system" →
   State contract. The producer's presenter-note aliases
   (<ירוק בהיר עומד> / <גבוהה>) are human-facing descriptions only and
   are deliberately NOT the stored values.

   Explicit ID → asset mapping, required because the supplied filenames
   ("Character_1_waving.mp4") don't follow the character-N-[state]
   convention. The side each character occupies is authored content from
   slide 3's presenter note, not something derived from filenames:
     physical RIGHT = <ירוק בהיר עומד> = character-1
     physical LEFT  = <גבוהה>          = character-2
   ───────────────────────────────────────────────────────────── */
const CHARACTER_ASSETS = {
  'character-1': {
    selection: 'assets/video/Character_1_waving.mp4',
    standing:  'assets/img/Character-1-standing.png',
    standingAlt: 'الشخصية البرتقالية المستديرة ذات العين الواحدة، واقفة',
    /* Measured ink bounds of the standing PNG, as fractions of its canvas
       height (alpha > 16, canvas 1254x1254). Consumed by the CSS to anchor
       the character to the ground line and the bubble to its head — see
       the anchoring note in css/style.css. Re-measure if the asset is
       ever replaced; a stale value here silently mis-places both. */
    standingInk: { top: 0.1477, bottom: 0.8813 },
    /* Horizontal ink bounds, measured the same way (alpha > 16 on the
       1254x1254 canvas). Needed by any screen that has to clear the
       character horizontally rather than just stand it on a ground
       line — the two characters' ink widths differ by 64%, so centring
       the PNG canvas puts them in very different places. */
    standingInkX: { left: 0.1834, right: 0.8182 },
    /* Speech-bubble anchor, as fractions of this character's own PNG
       canvas: x = the head's horizontal centre, top = the head's topmost
       ink. Measured, never hand-fitted. */
    standingHead: { x: 0.5008, top: 0.1475 },

    /* HEIGHT RELATIONSHIP — how tall this character is relative to the
       TALLER of the pair, independent of any single asset's canvas.
       0.7744 is the deck's own authored ratio, read off its
       two-character reference art (image4.png) and confirmed against the
       standing pair's measured ink (0.7337 / 0.9474 = 0.7744). It is
       what stops a pose whose canvas happens to be cropped tighter from
       silently stretching the round character to the lanky one's height.
       Never equalize the two. */
    heightRel: 0.7744,

    /* POSES. Every entry carries its own measured geometry because the
       assets do not share a canvas: standing is 1254x1254, ticket is
       911x911. Expressing everything as FRACTIONS of each asset's own
       canvas is what makes that harmless. Re-measure with
       scratchpad/png-ink.cjs if an asset is replaced. */
    poses: {
      standing: {
        src: 'assets/img/Character-1-standing.png',
        alt: 'الشخصية البرتقالية المستديرة ذات العين الواحدة، واقفة',
        ink:  { top: 0.1475, bottom: 0.8812 },
        inkX: { left: 0.1834, right: 0.8182 },
        head: { x: 0.5008, top: 0.1475 }
      },
      ticket: {
        src: 'assets/img/Character-1-ticket.png',
        alt: 'الشخصية البرتقالية المستديرة تحمل بطاقة قُرعة بكلتا يديها وتنظر إليها، وتبدو محبطة',
        ink:  { top: 0.0450, bottom: 0.9528 },
        inkX: { left: 0.1262, right: 0.8738 },
        head: { x: 0.4989, top: 0.0450 }
      }
    }
    /* No `fit` entry — see the measurement note below. */
  },
  'character-2': {
    selection: 'assets/video/Character_2_waving.mp4',
    standing:  'assets/img/Character-2-standing.png',
    standingAlt: 'الشخصية الصفراء الطويلة ذات عصبة الرأس البنفسجية، واقفة',
    standingInk: { top: 0.0143, bottom: 0.9617 },
    standingInkX: { left: 0.3070, right: 0.6938 },
    standingHead: { x: 0.5004, top: 0.0144 },
    /* The taller of the pair — the reference the other is measured against. */
    heightRel: 1,
    poses: {
      standing: {
        src: 'assets/img/Character-2-standing.png',
        alt: 'الشخصية الصفراء الطويلة ذات عصبة الرأس البنفسجية، واقفة',
        ink:  { top: 0.0143, bottom: 0.9617 },
        inkX: { left: 0.3070, right: 0.6938 },
        head: { x: 0.5004, top: 0.0143 }
      },
      ticket: {
        src: 'assets/img/Character-2-ticket.png',
        alt: 'الشخصية الصفراء الطويلة تحمل بطاقة قُرعة بكلتا يديها وتنظر إليها، وتبدو محبطة',
        ink:  { top: 0.0378, bottom: 0.9612 },
        inkX: { left: 0.3361, right: 0.6629 },
        head: { x: 0.5260, top: 0.0378 }
      }
    }
  }
};

/* PER-CHARACTER FIT SHIM (TransitionScreen.md → C): deliberately NOT
   applied to this asset pair — measured, not assumed.
   Both standing PNGs share a 1254x1254 canvas. Measured ink bounds:
     character-1  ink 796x920   → 73.37% of canvas height
     character-2  ink 485x1188  → 94.74% of canvas height
   That 21-point spread is NOT a framing error — it is the two
   characters' genuine height difference. Cross-checked against the
   deck's own authored artwork (slide 4 "Picture 5", both characters
   drawn together in one scene on a shared ground line): there the round
   character measures 0.7782 of the tall one's height, and in the
   supplied PNGs the same ratio is 73.37/94.74 = 0.7744 — a 0.5%
   difference, well inside the template's ~3% "normal variance, ship
   with no shim" threshold.
   Scaling character-1 up to match character-2's ink height would
   therefore be wrong: it would stretch the short round character to the
   lanky one's height and destroy the authored size relationship.
   The shim MECHANISM stays wired (applyStandingCharacter always writes
   both custom properties, which default to a no-op) so a future asset
   pair that genuinely needs one can use it — but do not add values here
   without re-measuring this pair. */
const CHARACTER_IDS = Object.keys(CHARACTER_ASSETS);


/* Resolve the companion character for a screen that displays one.
   Reads the CANONICAL in-memory state for Part 01 —
   window.lomdaState.selectedCharacter, set by Screen 01's selection —
   and never re-reads localStorage here: within Part 01 the in-memory
   value is authoritative. A RESUMED attempt gets its choice back on
   purpose — 50-loader.js adopts it from this component's state document
   (adoptUnitCharacter) before the cover drops. The selection is written
   through setUnitCharacter() (40-resume.js), which also keeps the
   localStorage mirror parts 02+ read.
   Validates against the IDs this part actually has assets for, so a
   value from an older/renamed scheme falls back instead of breaking. */
function resolveCharacter() {
  const sel = window.lomdaState.selectedCharacter;
  return (sel && Object.prototype.hasOwnProperty.call(CHARACTER_ASSETS, sel))
    ? sel
    : DEFAULT_CHARACTER;
}

/* Mount the selected character's standing image + alt + fit shim onto a
   screen's character <img>. Called from resetScreenState() on EVERY
   entry, so returning to the screen after changing the selection
   re-resolves it (companion character system verification checklist). */
/* Mount a character in a named POSE. Screens that show the companion in a
   specific authored pose (slide 14's "holding a raffle ticket") pass their
   own pose key; everything else uses 'standing'. All geometry comes from
   the pose's own measured fractions, so two poses on different canvases
   position identically without any per-screen nudging. */
function applyCharacterPose(imgId, poseKey) {
  const assets = CHARACTER_ASSETS[resolveCharacter()];
  const img = document.getElementById(imgId);
  if (!img || !assets) return;
  const pose = (assets.poses && assets.poses[poseKey]) || null;
  /* Custom properties go on the STAGE, not the <img>: the bubble is a
     SIBLING of the image, so it can only inherit them from their common
     ancestor. Both children then position themselves off the same values. */
  const stage = img.closest('.s1-char-stage') || img;
  // Ink bounds → ground line + bubble anchor. Written unconditionally,
  // including the neutral fallback, so a previous character's geometry
  // can never linger on the stage after a switch.
  const ink = (pose && pose.ink) || assets.standingInk || { top: 0, bottom: 1 };
  stage.style.setProperty('--ink-top', ink.top);
  stage.style.setProperty('--ink-bottom', ink.bottom);
  /* Ink HEIGHT and the character's height relationship let a screen ask
     for a target ink height and get it whatever the asset's canvas is —
     see the --char-box rule in css/style.css. */
  stage.style.setProperty('--ink-h', (ink.bottom - ink.top).toFixed(4));
  stage.style.setProperty('--height-rel', assets.heightRel != null ? assets.heightRel : 1);
  /* Horizontal bounds too — written unconditionally, and defaulting to
     the whole canvas so a screen that does not use them is unaffected. */
  const inkX = (pose && pose.inkX) || assets.standingInkX || { left: 0, right: 1 };
  stage.style.setProperty('--ink-left', inkX.left);
  stage.style.setProperty('--ink-w', (inkX.right - inkX.left).toFixed(4));
  /* Head anchor for the speech bubble — written unconditionally, with a
     neutral fallback, so no previous character's anchor can linger. */
  const head = (pose && pose.head) || assets.standingHead || { x: (inkX.left + inkX.right) / 2, top: ink.top };
  stage.style.setProperty('--head-x', head.x);
  stage.style.setProperty('--head-top', head.top);
  // Fit shim written unconditionally — including no-op values — so
  // switching characters never leaves the previous one's scale behind.
  const fit = assets.fit || {};
  stage.style.setProperty('--char-scale', fit.scale || 1);
  stage.style.setProperty('--char-shift-y', (fit.shiftY || 0) + 'px');
  const src = (pose && pose.src) || assets.standing;
  img.alt = (pose && pose.alt) || assets.standingAlt;
  // Only touch src when it actually differs, so re-entry doesn't cause a
  // needless reload/flicker of an already-correct image.
  if (img.getAttribute('src') !== src) img.src = src;
}

/* Existing entry point, unchanged for every screen that already calls it. */
function applyStandingCharacter(imgId) { applyCharacterPose(imgId, 'standing'); }


/* ─── Global lomda state ─────────────────────────────────────
   Single source of truth that persists across every screen for the
   whole session. Always starts unselected (see note above). */
window.lomdaState = window.lomdaState || {
  selectedCharacter: null
};

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
  // Feedback popups and hint overlays live OUTSIDE the .screen divs, so a
  // screen swap alone never hides them — navigation has to close them
  // itself, before the swap (_global-components.md → Feedback popup
  // system → "Popups are contextual to the current screen only"). Relies
  // on the id naming contract: popup ids end in -popup, hint overlays in
  // -hint-overlay. Closing never touches answer state.
  document.querySelectorAll('[id$="-popup"], [id$="-hint-overlay"]')
    .forEach(el => el.classList.add('hidden'));
  closeZoom();
  const prev = document.querySelector('.screen.active');
  if (prev) {
    /* Clear a reveal that is still mid-flight, so leaving a screen early
       can never strand one of its groups at opacity 0 (animation-system). */
    if (window.Motion) Motion.cancel(prev);
    prev.classList.remove('active');
  }
  // Leaving screen 08 stops its gesture hint (_gesture-hints.md lifecycle
  // rule: nothing keeps animating on a hidden screen).
  if (prev && prev.id === 's7') frcGhStop();
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
   again WITHIN the same loaded page (no reload). For Screen 01 this is
   ordinary same-session restoration of the in-memory selection — it
   never reads localStorage, per the cross-part contract. */
function resetScreenState(n) {
  // Leaving (or re-entering) any screen kills the comic sequence's pending
  // steps, so navigating away mid-reveal can never leave a stray timer
  // firing into a screen the learner is no longer on.
  cmcClearTimers();
  // Same rule for screen 11's exercise reveal: a pending unlock must never
  // fire into a screen the learner has already left.
  if (typeof s10Timer !== 'undefined' && s10Timer) { clearTimeout(s10Timer); s10Timer = null; }

  if (n === 5) {
    /* Screen 06: six authored units enter in the producer's 1..6 order.
       Re-entry within the same loaded page shows the finished state
       immediately (and keeps Continue open) rather than replaying — the
       standard 720 seen-once convention. With resume on the flag is saved
       in the state document (PART_SCREENS[5]), so a restore keeps it too. */
    if (cmcSequenceDone || cmcPrefersReducedMotion()) cmcShowAll();
    else cmcStartSequence();
  }

  if (n === 7) frcEnter();

  if (n === 8) {
    /* Resume-state lock: a resolved question never hands out a fresh
       attempt on return, and a resolved screen never comes back without
       its feedback (_global-components.md → "Reopening feedback on
       return to a resolved screen"). s8scDone is the only flag involved
       — there is deliberately no second "popup was open on exit" flag.
       SingleChoiceQuestion has no View My Answer Toggle (it already
       shows the chosen and the correct option at once), so reopening is
       just re-showing the last outcome's popup; the option marks
       themselves are DOM state that already persisted. */
    if (!s8scDone) { if (screenPristine(8)) s8scResetInitial(); }
    else s8scShowPopup(s8scResult);
  }

  if (n === 9) {
    /* Stateless TransitionScreen — nothing to restore, no gate. The two
       characters are re-mounted on every entry, matching the general
       re-render-on-re-entry contract for any character-displaying
       screen. */
    s9MountClip();
  }

  if (n === 10) {
    /* Slide 14's own note: "הדמות מחזיקה כרטיס, מסתכלת עליו ונראית
       מתאפקת שלא לבכות" — the authored pose, now that its artwork
       exists. */
    applyCharacterPose('s10-char-img', 'ticket');
    s10Enter();
  }

  if (n === 11) s11Enter();

  if (n === 12) s12Enter();

  if (n === 13) s13Enter();

  if (n === 14) s14Enter();

  if (n === 15) s15Enter();

  if (n === 16) s16Enter();

  if (n === 17) s17Enter();

  if (n === 18) s18Enter();

  if (n === 19) s19MountClip();

  if (n === 20) s20MountClip();

  if (n === 21) {
    /* Same resume-state lock as screen 09: a resolved question never hands
       out a fresh attempt on return, and never comes back without its
       feedback (_global-components.md → "Reopening feedback on return to a
       resolved screen"). Within one page load the hotspot marks are DOM
       state that persisted; s21hsRestoreMarks() is idempotent, so running
       it on every entry costs nothing. After a resume restore the screen
       is rebuilt by its painter (PART_SCREENS[21], end of file). */
    if (!s21hsDone) { if (screenPristine(21)) s21hsResetInitial(); }
    else { s21hsRestoreMarks(); s21hsShowPopup(s21hsResult); }
  }

  if (n === 22) {
    s22MountClip();
    /* Same resume-state lock as every other question screen: a resolved
       section never hands out a fresh attempt on return, and never comes
       back without its feedback. s22viqApplyView() restores whichever
       view the learner left the toggle on. */
    if (!s22viqDone) { if (screenPristine(22)) s22viqResetInitial(); }
    else if (s22viqResult === 'wrong2') s22viqApplyView();
    else s22viqShowPopup(s22viqResult);
  }

  if (n === 23) {
    if (!s23viqDone) { if (screenPristine(23)) s23viqResetInitial(); }
    else if (s23viqResult === 'wrong2') s23viqApplyView();
    else s23viqShowPopup(s23viqResult);
  }

  if (n === 0) {
    const saved = window.lomdaState.selectedCharacter;
    document.querySelectorAll('#s0 .option-card').forEach(card => {
      const isSelected = !!saved && card.dataset.value === saved;
      card.classList.toggle('selected', isSelected);
      card.setAttribute('aria-checked', isSelected ? 'true' : 'false');
    });
    const continueBtn = document.getElementById('s0-continue');
    if (continueBtn) continueBtn.disabled = !saved;
    /* Decorative option-card clips: play once, hold the final frame.
       Per the end-state hierarchy's step 1 (_global-components.md →
       "Decorative character video" → End-state hierarchy) NOTHING is
       wired here — a <video> with no `loop` already stops on its last
       encoded frame, which is exactly the required behaviour. Both
       clips were sampled during build verification and hold a
       meaningful final frame, so the blank-tail exception (step 2)
       does not apply. Never add a `currentTime = 0` end handler: that
       rewinds to the OPENING pose and is explicitly prohibited. */
  }
  if (n === 1) {
    /* TransitionScreen (local variant): re-resolve the companion
       character on every entry, so returning here after changing the
       Screen 01 selection shows the newly chosen character. Continue is
       ungated on this screen, so there is nothing to enable/disable. */
    applyStandingCharacter('s1-char-img');
  }
  if (n === 2) {
    /* FullImageWithTextOverlay Variant C: stateless re-entry apart from
       the completion gate. The iframe is NEVER re-sourced here — doing so
       would reload the widget and wipe the learner's game progress on
       every return visit. Restore only the resolved Continue state. */
    const btn = document.getElementById('s2-continue');
    if (btn) btn.disabled = !s2WidgetCompleted;
  }
}

/* ═══════════════════════════════════════════════════════════
   Screen 06 (#s5) — comic sequence controller
   ───────────────────────────────────────────────────────────
   Six authored units reveal in the producer's 1..6 order ("להעלות את
   התמונות בהדרגה לפי המספור"). Each unit is one image plus its own
   speech bubble — they enter together, never separately.

   DWELL between units is derived from the unit's OWN text at runtime, so
   it can never drift from a copy edit:
       reading time = words / 3 words-per-second
       + 1s visual orientation
       clamped to [2.5s, 8s]
   A "word" is a whitespace-separated token containing at least one letter
   or digit, which drops bare punctuation and operators (":", "!", "=",
   "·") from the count, as specified.

   THE GATE: Continue ships disabled and opens by itself the moment unit 6
   finishes entering — wired to that unit's own `animationend`, not to a
   parallel timer that could drift out of step with the visuals. The
   learner never has to click anything to unlock it.

   REDUCED MOTION: with `prefers-reduced-motion: reduce`, nothing is
   scheduled at all — every unit is shown at once and Continue opens
   immediately.
   ═══════════════════════════════════════════════════════════ */
const CMC = { screen: 5, wordsPerSecond: 3, orientationSeconds: 1, minSeconds: 2.5, maxSeconds: 8 };
let cmcSequenceDone = false;
let cmcTimers = [];

function cmcPrefersReducedMotion() {
  return window.matchMedia && window.matchMedia('(prefers-reduced-motion: reduce)').matches;
}
function cmcUnits() {
  return Array.prototype.slice.call(document.querySelectorAll('#s5 .cmc-unit'))
    .sort((a, b) => (+a.dataset.cmcStage) - (+b.dataset.cmcStage));
}
/* Words = tokens carrying a letter or digit, counted on the rendered text
   (textContent, so markup never inflates the count). */
function cmcWordCount(el) {
  return el.textContent.trim().split(/\s+/)
    .filter(w => /[\p{L}\p{N}]/u.test(w)).length;
}
function cmcDwellMs(el) {
  const raw = cmcWordCount(el) / CMC.wordsPerSecond + CMC.orientationSeconds;
  return Math.min(CMC.maxSeconds, Math.max(CMC.minSeconds, raw)) * 1000;
}
function cmcClearTimers() {
  cmcTimers.forEach(clearTimeout);
  cmcTimers = [];
}
function cmcUnlockContinue() {
  cmcSequenceDone = true;
  const btn = document.getElementById('s5-continue');
  if (!btn) return;
  btn.disabled = false;
  btn.removeAttribute('aria-disabled');
}
/* Finished state: every unit visible, no animation, Continue open. Used on
   re-entry after the sequence has run and under reduced motion. */
function cmcShowAll() {
  cmcClearTimers();
  const screen = document.getElementById('s5');
  if (!screen) return;
  cmcUnits().forEach(u => u.classList.remove('is-in'));
  screen.classList.add('cmc-all-shown');
  cmcUnlockContinue();
}
function cmcStartSequence() {
  const screen = document.getElementById('s5');
  const units = cmcUnits();
  if (!screen || !units.length) return;

  cmcClearTimers();
  screen.classList.remove('cmc-all-shown');
  units.forEach(u => u.classList.remove('is-in'));
  void screen.offsetWidth;                 // reflow, so a restart re-runs the animation

  const last = units[units.length - 1];
  if (!last._cmcWired) {
    last._cmcWired = true;
    // Semantic end of the sequence: the final unit has finished entering.
    last.addEventListener('animationend', function () {
      if (screen.classList.contains('active')) cmcUnlockContinue();
    });
  }

  let at = 0;
  units.forEach((unit, i) => {
    if (i === 0) unit.classList.add('is-in');
    else cmcTimers.push(setTimeout(() => unit.classList.add('is-in'), at));
    at += cmcDwellMs(unit);                // this unit's own read time
  });
}

/* ═══════════════════════════════════════════════════════════
   Screen 08 (#s7) — FlipCardsReveal controller
   ───────────────────────────────────────────────────────────
   720-templates → FlipCardsReveal.md, "Interaction Rules":

     "Canonical unlock rule: The Continue button remains disabled until
      every card has been flipped at least once. ... The unlock tracks
      'seen at least once' (frcRevealed — entries only ever turn true),
      independent of each card's current visual flip state."

     "Flip-back behavior is a per-lomda producer decision: the default is
      one-way (cards stay flipped, as described below), but producer notes
      may request two-way cards ('ניתנים להיפוך חזרה') — in that case cards
      toggle freely and a separate flip-state array tracks the current
      visual state. The canonical unlock rule holds unchanged in both
      variants, because it tracks seen-at-least-once, not the current
      state."

   Slide 10's presenter notes (notesSlide8.xml, resolved through the
   slide's own _rels) name the screen — "שם המסך בפגימה: FlipCardsReveal" —
   and the three images with their ALT texts. They contain NO reversibility
   request, so this lomda takes the MD's documented default: ONE-WAY.

   State is nonetheless tracked in two independent arrays, exactly as the
   MD's two variants require:
     frcSeen[] — ever revealed. Only ever turns true. Drives the gate.
     frcFace[] — the card's CURRENT visual face (false = image, true = text).
   Flipping FRC_FLIP_BACK to true is therefore the whole two-way variant:
   the gate keeps reading frcSeen, so Continue stays enabled even when a
   learner flips cards back, and completion survives navigation either way.
   ═══════════════════════════════════════════════════════════ */
const FRC_FLIP_BACK = true;                // two-way on producer request, 2026-09-22
let frcSeen = [false, false, false];       // ever revealed — one per .frc-card, DOM order
let frcFace = [false, false, false];       // current face: false = image, true = text
let frcDone = false;

/* ── Gesture hint (rightmost card) — 720-templates → _gesture-hints.md.
   Reuses the validated mass-measure-01-01 S6 flip-card pattern: one
   class-gated looping CSS animation, JS owns start/stop. Click/tap family
   → persistent: loops until the learner flips ANY card (any equivalent
   element counts as "gesture found"), then stays off for the session.
   frcGhSeen is DELIBERATELY separate from frcSeen/frcFace/frcDone: it
   records only that the learner has performed the real gesture, is set
   only by the real flip handler, and is never cleared — so no reveal-
   state change can revive the hint, and the hint can never touch the
   reveal state or the Continue gate. */
let frcGhSeen = false;
function frcGhStart() {
  if (frcGhSeen) return;                    // learner already found the gesture
  if (window.matchMedia && window.matchMedia('(prefers-reduced-motion: reduce)').matches) return;
  const gh = document.getElementById('frc-gh');
  if (gh) gh.classList.remove('frc-gh-stopped');
}
function frcGhStop() {
  const gh = document.getElementById('frc-gh');
  if (gh) gh.classList.add('frc-gh-stopped');
}

/* Apply a card's current face to the DOM: visual state, aria-expanded on
   the button and aria-hidden on the text face stay synchronized, so the
   accessible name/expanded state always describes what is actually shown. */
function frcApplyFace(cardEl, showText) {
  cardEl.classList.toggle('is-flipped', showText);
  cardEl.setAttribute('aria-expanded', showText ? 'true' : 'false');
  const back = cardEl.querySelector('.frc-card-back');
  if (back) back.setAttribute('aria-hidden', showText ? 'false' : 'true');
}

/* Click / Enter / Space on the card. The card is a real <button>, so the
   keyboard path is the browser's own activation — no key handler of our
   own, and focus stays on the card across the flip. */
function frcFlip(cardEl) {
  frcGhSeen = true;                         // the real gesture, on any card, ends the demo
  frcGhStop();
  const i = +cardEl.dataset.index;
  if (frcFace[i] && !FRC_FLIP_BACK) return;   // one-way: a revealed card ignores further clicks
  frcFace[i] = !frcFace[i];
  if (frcFace[i]) frcSeen[i] = true;          // seen-at-least-once, never cleared
  frcApplyFace(cardEl, frcFace[i]);
  frcCheckUnlock();
}

function frcCheckUnlock() {
  if (!frcSeen.every(Boolean)) return;        // gate reads frcSeen, never frcFace
  frcDone = true;
  const btn = document.getElementById('s7-continue');
  if (!btn) return;
  btn.disabled = false;
  btn.removeAttribute('aria-disabled');
}

/* Called from resetScreenState(7) on every entry — restores each card's
   visual state from frcFace (idempotent) and the Continue button's
   resolved state from frcDone, so returning to this screen never re-locks
   a finished one nor re-flips a card the learner left face-down. */
function frcEnter() {
  document.querySelectorAll('#s7 .frc-card').forEach(card => {
    frcApplyFace(card, !!frcFace[+card.dataset.index]);
  });
  const btn = document.getElementById('s7-continue');
  if (btn) btn.disabled = !frcDone;
  frcGhStart();                             // no-op once the gesture was found, or under reduced motion
}

/* ─── Forward navigation ─────────────────────────────────────
   The gate lives HERE, not only in the Continue button's handler: the
   engine binds ArrowLeft directly to advanceScreen(), so a gate that
   sat only in an onclick would be silently bypassable by keyboard
   (_global-components.md → "Advance gates"). */
function advanceScreen() {
  // Screen 01 gate: a character must be selected before advancing
  if (currentScreen === 0 && !window.lomdaState.selectedCharacter) return;
  // Screen 03 gate: the widget game must be completed before advancing.
  // Lives here, not only in the button handler, so the ArrowLeft key
  // cannot bypass it (_global-components.md -> "Advance gates").
  if (currentScreen === 2 && !s2WidgetCompleted) return;
  // Screen 06 gate: the comic sequence must have finished revealing. Lives
  // here too, so ArrowLeft cannot bypass what the button enforces.
  if (currentScreen === 5 && !cmcSequenceDone) return;
  // Screen 08 gate: every card must have been flipped at least once.
  if (currentScreen === 7 && !frcDone) return;
  /* Gate on RESOLVED, not on correctness: a final wrong attempt still
     resolves the question and still lets the learner move on
     (_global-components.md → "Gate on the resolved state"). */
  if (currentScreen === 8 && !s8scDone) return;
  /* Screen 11 is an EXPLANATORY guided step, so this is not an answer
     gate — it is the producer's own authored condition, "המשך יהיה זמין
     ללחיצה אחרי הופעת התרגיל". */
  if (currentScreen === 10 && !s10Revealed) return;
  /* Screen 12 is an OPTION step: the gate is the step being resolved by
     a selection (GuidedMathPracticeScreen.md → D). */
  if (currentScreen === 11 && !s11Done) return;
  /* Screens 13–15 are EXPLANATORY guided steps (B1): nothing to
     complete, so nothing to gate on. */
  /* Screen 16 is an OPTION step again: the gate is the step being
     resolved by a selection (GuidedMathPracticeScreen.md → D). */
  if (currentScreen === 15 && !s15Done) return;
  if (currentScreen === 17 && !s17Done) return;
  /* Screen 22 — gate on RESOLVED, not on correctness: a final wrong
     attempt still resolves the question and still lets the learner move
     on (_global-components.md → "Gate on the resolved state"). */
  if (currentScreen === 21 && !s21hsDone) return;
  /* Screen 23 — same rule: gate on the SECTION being resolved, not on
     it being correct. */
  if (currentScreen === 22 && !s22viqDone) return;
  /* Screen 24 — the last section. Same rule: gate on RESOLVED, not on
     correct. Note this gate governs leaving the SECTION, not the
     question: the question's own outcome is an aggregate and never
     gates navigation. */
  if (currentScreen === 23 && !s23viqDone) return;
  /* The part boundary. S23 is this part's last screen, so there is no
     goTo() target: crossing is the advance. */
  if (currentScreen === 23) { crossToNextPart(); return; }
  goTo(currentScreen + 1);
}

/* ═══════════════════════════════════════════════════════════
   PART BOUNDARY — 03-01 → 03-02, landing on its S0

   UNCONDITIONAL. This part carries one warm-up question (sections
   א/ב/ג, slides 25-27) and its outcome must NOT change the route —
   producer decision, 2026-09-14. The only precondition is that section
   ג has RESOLVED, which advanceScreen()'s own gate above already
   enforces; correct and incorrect resolutions take the identical path.
   That is the one way this part differs from 03-02, whose own boundary
   is gated: 2 of its 3 aggregated questions correct routes on to 03-03's
   S0, and a fail keeps the learner there with המשך still enabled
   (approved, 2026-09-14). Neither part emits an outcome.

   MECHANISM. PartNav.forward() (js/partnav.js), the project's existing
   cross-part contract, byte-identical to the copies in 03-02/03/04. It
   pushes this part and the screen being left onto the sessionStorage
   trail before navigating, which is what gives 03-02's S0 a Back button
   returning HERE, to S23, rather than to a hardcoded predecessor. No ad
   hoc window.location assignment is used anywhere. The selected
   character needs no special handling: it lives in localStorage under
   the lomda-wide CHARACTER_STORAGE_KEY that 03-02 reads.

   RE-CROSSING. If the learner already crossed and came back,
   PartNav.forwardOverride() reports the destination recorded for their
   journey and that destination is REPLAYED rather than re-derived
   (_global-components.md → "Re-crossing a previously-resolved boundary
   forward must replay the recorded destination"). The trail lives in
   sessionStorage, which a learner can edit, so a replayed id that is
   not this part's own destination is refused rather than followed.

   NO OUTCOME IS REPORTED HERE. The Ministry reporting/integration layer
   is owned by Vadim, outside lomda production: no SCORM, no xAPI, no
   cmi.*, no completion postMessage. This is internal routing only.
   ═══════════════════════════════════════════════════════════ */
const NEXT_PART = '02';

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

/* Back returns to the previous valid predecessor in the learner's actual
   route. Screen 01 is the lomda's first screen and has no predecessor,
   so no Back control is rendered and this is a no-op for now. */
function goBack() {
  /* From this part's FIRST screen, Back crosses the boundary backwards
     via PartNav rather than no-opping. This part is the journey's entry
     point, so in practice the trail is empty here and PartNav.back() is
     itself a no-op — the branch exists so the part obeys the same shared
     contract as 03-02/03/04 if the graph ever gains a predecessor. */
  if (currentScreen === 0) { PartNav.back(); return; }
  if (currentScreen < 0) return;
  goTo(currentScreen - 1);
}

document.addEventListener('keydown', e => {
  if (hasOwnArrowKeyMeaning(e.target)) return;
  if (e.key === 'ArrowLeft')  advanceScreen();
  if (e.key === 'ArrowRight') goBack();
});

/* ═══════════════════════════════════════════════════════════
   TEMPLATE — TwoOptionSelection (Screen 01, character selection)
   Selecting a card stores the choice and enables Continue. No modal,
   no auto-advance, single selection only, no deselection-to-empty
   (the canonical contract has no "unselect" path — clicking the
   already-selected card re-affirms it and Continue stays enabled;
   clicking the other card moves the selection).
   See 720-templates → references/TwoOptionSelection.md
   ═══════════════════════════════════════════════════════════ */
function selectOption(cardEl) {
  document.querySelectorAll('#s0 .option-card').forEach(c => {
    c.classList.remove('selected');
    c.setAttribute('aria-checked', 'false');
  });
  cardEl.classList.add('selected');
  cardEl.setAttribute('aria-checked', 'true');

  // Persist for the rest of the session and across parts through the
  // shared setter (unit-js/40-resume.js): it sets
  // window.lomdaState.selectedCharacter, writes the localStorage mirror
  // under CHARACTER_STORAGE_KEY and the unit state document. Only the
  // logical ID is stored.
  const value = cardEl.dataset.value;
  setUnitCharacter(value);

  const continueBtn = document.getElementById('s0-continue');
  if (continueBtn) continueBtn.disabled = false;
}

/* Enter/Space on a focused card selects it — radiogroup keyboard a11y
   required by TwoOptionSelection.md → "JS contract". Delegated so it
   covers every .option-card without per-node inline handlers. */
document.addEventListener('keydown', e => {
  // Match on BOTH key and code: `key` is the spec-correct check, while
  // `code === 'Space'` also catches the legacy 'Spacebar' key name and
  // keeps the activation working on layouts/hosts that report the space
  // bar inconsistently.
  const isEnter = e.key === 'Enter' || e.code === 'Enter' || e.code === 'NumpadEnter';
  const isSpace = e.key === ' ' || e.key === 'Spacebar' || e.code === 'Space';
  if (!isEnter && !isSpace) return;
  const card = e.target.closest && e.target.closest('.option-card');
  if (!card) return;
  e.preventDefault();   // stop Space from scrolling the page
  selectOption(card);
});

/* Continue ("בחרתי") → advance. */
function advanceFromS0() {
  advanceScreen();
}

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
   SHARED QUESTION CHROME — feedback popup, hint button, image zoom
   720-templates → _global-components.md (Feedback popup system, Image
   zoom) + _question-template-defaults.md. Ported from the same-lomda
   sibling methodica-math-ratio-03-02 (popup/hint) and the approved
   methodica-math-percent-01-01 (image zoom), so both parts of this
   lomda behave identically. Written once here — every later question
   screen in this part reuses it rather than reimplementing it.
   ═══════════════════════════════════════════════════════════ */


let popupDragId = null, popupOffX = 0, popupOffY = 0;


function popupPointerDown(e, popupId) {
  if (e.target && e.target.closest && e.target.closest('.popup-action')) return;
  const popup = document.getElementById(popupId);
  const app   = document.getElementById('app');
  if (!popup || !app) return;
  popupDragId = popupId;
  const scale = getAppScale();
  const canvasY = (e.clientY - (parseFloat(app.style.top)  || 0)) / scale;
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

/* Feedback-popup action row. `action` falsy → the row is removed;
   SingleChoiceQuestion never carries one (no View My Answer Toggle). */
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
  btn.onclick = onClick || null;
}


/* Escape closes whichever hint overlay is open — never a feedback
   popup, which has no dismissal path at all. */
document.addEventListener('keydown', e => {
  if (e.key !== 'Escape') return;
  document.querySelectorAll('[id$="-hint-overlay"]').forEach(el => {
    if (!el.classList.contains('hidden')) el.classList.add('hidden');
  });
  closeZoom();
});

/* Image zoom — canonical clone-based shared modal. One delegated click
   handler: [data-zoom-src] opens it by cloning the opener's own image
   frame into the shared stage, [data-zoom-close] closes it. */
function closeZoom() {
  const m = document.getElementById('img-zoom-modal');
  if (!m) return;
  m.classList.add('hidden');
  m.classList.remove('img-zoom-modal--wide');   // never carries to the next opener
}
document.addEventListener('click', e => {
  const opener = e.target.closest && e.target.closest('[data-zoom-src]');
  if (opener) {
    const modal = document.getElementById('img-zoom-modal');
    const stage = document.getElementById('img-zoom-stage');
    /* .scq-img-inner is the original frame class; data-zoom-frame lets a
       screen whose zoomable frame is not an .scq-img-inner (S21's hotspot
       surface) name its own without renaming anything that already works. */
    const wrapper = opener.parentElement && opener.parentElement.closest
      ? (opener.closest('[data-zoom-frame]') || opener.closest('.scq-img-inner'))
      : null;
    if (modal && stage && wrapper) {
      stage.innerHTML = '';
      const clone = wrapper.cloneNode(true);
      clone.querySelectorAll('.img-zoom-btn').forEach(b => b.remove());
      /* cloneNode(true) duplicates ids too — strip them so the document
         never carries two elements with the same id at once. */
      if (clone.id) clone.removeAttribute('id');
      clone.querySelectorAll('[id]').forEach(el => el.removeAttribute('id'));
      stage.appendChild(clone);
      /* A landscape picture in the square panel would render SMALLER than
         it does in the page. A trigger opts into the wide panel; every
         other screen keeps the shared square. */
      modal.classList.toggle('img-zoom-modal--wide', opener.hasAttribute('data-zoom-wide'));
      modal.classList.remove('hidden');
    }
    return;
  }
  if (e.target.closest && e.target.closest('[data-zoom-close]')) closeZoom();
});


/* ═══════════════════════════════════════════════════════════
   Screen 09 (#s8) — SingleChoiceQuestionImage controller
   Source: slide 12 (presentation position). See index.html's own
   source-evidence comment above #s8 for the shape inventory, the
   off-canvas popup classification and the correct-answer evidence.

   CORRECT = 'o2' ("2x : 4x"), never resolved from a letter key: the
   slide's own correct-feedback popup states the answer outright, and
   the rendered slide highlights that option's text in green as the
   author's correctness marker.

   ATTEMPTS = 2, the 720 default. Slide 12 authors exactly the deck's
   normal 2-attempt shape — one correct popup and one final-wrong popup
   sharing a single explanation body — and nothing in the slide or its
   notes asks for a different attempt model
   (_question-template-defaults.md → "Attempt model").

   The shared explanation is verbatim from the slide's own two popups
   (#42 / #44), which carry byte-identical bodies and differ only in
   title. The ratio inside it is authored "4 : 2" in the XML and is
   RENDERED by PowerPoint as 2 · : · 4 — a pure-digit run reordered by
   the bidi algorithm — so it is reproduced here as the isolated LTR
   unit "2 : 4", which is the same physical arrangement the slide shows
   and the same one the answer "2x : 4x" uses.
   ═══════════════════════════════════════════════════════════ */
const S8SC_EXPLAIN = [
  'للحفاظ على النسبة <bdi dir="ltr" class="scq-expr">2 : 4</bdi>، نضرب المكوّنَين بنفس العامل، ولذلك التعبير الجبري هو:',
  '<bdi dir="ltr" class="scq-expr">2x : 4x</bdi>'
];
const S8SC = { correctId: 'o2', maxAttempts: 2 };
/* Titles: the correct and final-wrong titles are the slide's own
   authored copy. The retry title is the canonical fallback — the slide
   mocks no retry popup, which is the normal shape for a 2-attempt
   question and not evidence that the state does not exist. The
   final-wrong title carries the canonical second line required by
   _question-template-defaults.md → "Terminal-incorrect title carries a
   fixed second line"; it is APPENDED to the authored line, never a
   replacement, and both lines render bold. */
const S8SC_POPUP = {
  retry:   { bg: '#ffdbdc', title: 'الإجابة غير صحيحة.', body: ['هل نحاول مرة أخرى؟'] },
  correct: { bg: '#edf8ed', title: 'صحيح‎!', body: S8SC_EXPLAIN },
  wrong2:  { bg: '#ffdbdc',
             title: 'أخطأتم، لكن لا بأس، هذه أفضل طريقة للتعلّم.<br>الإجابة الصحيحة تَظهَر الآن.',
             body: S8SC_EXPLAIN }
};

let s8scSelected  = null;
let s8scAttempts  = 0;
let s8scDone      = false;
let s8scResult    = null;   // 'correct' | 'wrong2'
let s8scLastWrong = null;   // the option id just marked wrong — the retry gate compares against THIS

function s8scOpt(id) { return document.querySelector('#s8 .scq-opt[data-id="' + id + '"]'); }
function s8scOpts()  { return document.querySelectorAll('#s8 .scq-opt'); }

/* Single select; clears a prior wrong attempt's marks; re-enables the
   check button ONLY if the newly selected id differs from the one just
   marked wrong (_question-template-defaults.md → "Retry gate" —
   comparison shape for this template: the selected option id). */
function s8scSelect(id) {
  if (s8scDone) return;
  s8scOpts().forEach(o => {
    o.classList.remove('selected');
    o.setAttribute('aria-checked', 'false');
  });
  const opt = s8scOpt(id);
  if (!opt) return;
  opt.classList.add('selected');
  opt.setAttribute('aria-checked', 'true');
  s8scSelected = id;

  if (s8scAttempts > 0) {
    document.getElementById('s8sc-popup').classList.add('hidden');
    s8scOpts().forEach(o => o.classList.remove('correct', 'wrong'));
  }
  const chk = document.getElementById('s8sc-check');
  if (chk) chk.disabled = (id === s8scLastWrong);
}

function s8scCheck() {
  if (s8scDone) { advanceScreen(); return; }
  if (!s8scSelected) return;

  s8scAttempts++;
  const correct = s8scSelected === S8SC.correctId;
  reportQ(['002', 'q1'], correct, correct || s8scAttempts >= S8SC.maxAttempts, xapiChoiceAnswer('002', 'q1', '#s8 .scq-opt', s8scSelected));

  if (correct) {
    s8scOpt(s8scSelected).classList.add('correct');
    s8scResult = 'correct';
    s8scFinish();
    s8scShowPopup('correct');
    return;
  }

  if (s8scAttempts < S8SC.maxAttempts) {
    /* Non-final wrong: mark the chosen option wrong; the correct one is
       NOT revealed yet (SingleChoiceQuestion.md → state 4, Try Again). */
    s8scOpt(s8scSelected).classList.add('wrong');
    s8scLastWrong = s8scSelected;
    document.getElementById('s8sc-check').disabled = true;
    s8scShowPopup('retry');
    return;
  }

  /* Final wrong: the chosen option red + ✗ and the correct one green +
     ✓, both shown at once — which is exactly why this template needs no
     View My Answer Toggle. */
  s8scOpt(s8scSelected).classList.add('wrong');
  s8scOpt(S8SC.correctId).classList.add('correct');
  s8scResult = 'wrong2';
  s8scFinish();
  s8scShowPopup('wrong2');
}

function s8scFinish() {
  s8scDone = true;
  s8scOpts().forEach(o => { o.disabled = true; });
  const chk = document.getElementById('s8sc-check');
  /* SingleChoiceQuestion's own completion label. Slide 12 mocks only
     the unresolved "هل إجابتي صحيحة؟" state, so the template's canonical fallback
     applies (_global-components.md → "Completion state and primary
     button"); the same lomda's part 02 resolves to the same label. */
  if (chk) { setNavLabel(chk, 'هل نتابع؟'); chk.disabled = false; }
  hideHintButton('s8sc-hint');
}

function s8scResetInitial() {
  s8scSelected  = null;
  s8scAttempts  = 0;
  s8scDone      = false;
  s8scResult    = null;
  s8scLastWrong = null;
  s8scOpts().forEach(o => {
    o.classList.remove('selected', 'correct', 'wrong');
    o.setAttribute('aria-checked', 'false');
    o.disabled = false;
  });
  document.getElementById('s8sc-popup').classList.add('hidden');
  document.getElementById('s8sc-hint-overlay').classList.add('hidden');
  const chk = document.getElementById('s8sc-check');
  if (chk) { setNavLabel(chk, 'هل إجابتي صحيحة؟'); chk.disabled = true; }
  /* Hint availability: the template default, visible from the start —
     slide 12 mocks "אפשר רמז?" in its enabled state and neither the
     slide nor its notes ask for the delayed policy. Once available, the
     hint stays reusable indefinitely (open/close/reopen at will) — the
     same shape every other hinted screen in this lomda already uses
     (e.g. 03-02 S2/S5). Corrected 2026-09-24: this screen used to
     disable the button after its first use, per a misreading of
     SingleChoiceQuestion.md's generic "disables after use" line as a
     universal rule; the producer has since clarified that unlock TIMING
     may vary by screen, but once unlocked a hint is always reusable —
     disabling it after one open is a defect, not a variant. */
  showHintButton('s8sc-hint');
}

function s8scShowPopup(type) {
  const popup = document.getElementById('s8sc-popup');
  if (!popup) return;
  const cfg = S8SC_POPUP[type];
  if (!cfg) return;
  popup.style.background = cfg.bg;
  if (popup.classList.contains('hidden')) resetPopupPosition(popup);
  /* innerHTML, not textContent: the final-wrong title is two bold lines
     separated by a real break, and the bodies carry isolated LTR
     expressions. */
  document.getElementById('s8sc-popup-title').innerHTML = cfg.title;
  document.getElementById('s8sc-popup-body').innerHTML =
    cfg.body.map(p => '<p>' + p + '</p>').join('');
  renderPopupAction('s8sc-popup', null, null);   // no action row on this template
  popup.classList.remove('hidden');
  clampPopup(popup);
}

function s8scOpenHint() {
  xapiRequestedHint('002', 'q1');
  document.getElementById('s8sc-hint-overlay').classList.remove('hidden');
}
function s8scCloseHint() {
  document.getElementById('s8sc-hint-overlay').classList.add('hidden');
}
function s8scCloseHintOnBackdrop(e) {
  if (e.target === e.currentTarget) s8scCloseHint();
}

/* Enter/Space on a focused option selects it — radiogroup keyboard
   a11y. Delegated, so it covers every .scq-opt without per-node inline
   handlers. */
document.addEventListener('keydown', e => {
  const isEnter = e.key === 'Enter' || e.code === 'Enter' || e.code === 'NumpadEnter';
  const isSpace = e.key === ' ' || e.key === 'Spacebar' || e.code === 'Space';
  if (!isEnter && !isSpace) return;
  const opt = e.target.closest && e.target.closest('#s8 .scq-opt');
  if (!opt || opt.disabled) return;
  e.preventDefault();
  s8scSelect(opt.dataset.id);
});


/* ═══════════════════════════════════════════════════════════
   Screen 10 (#s9) — TransitionScreen, two-character variant
   Source: slide 13. See index.html's own source-evidence comment above
   #s9 for the shape inventory and for the producer decision behind this
   clip.

   ONE CLIP FOR EVERY LEARNER. TransitionScreen.md's companion-character
   resolution reads window.lomdaState.selectedCharacter because a normal
   instance shows THE learner's own character. This instance shows BOTH
   characters in one authored clip, and the producer instructed that it
   always plays regardless of the Screen 01 choice — so there is
   deliberately no map keyed by character and no resolveCharacter() call
   here. Adding one would be a bug, not a fix.

   No freezeAt: the clip's final frame is real content (measured — luma
   range 2..255, mean 232), so it is left to end naturally and hold that
   frame, per TransitionScreen.md → "Decorative video end-state". Never
   reset currentTime. */
const S9_CLIP = 'assets/video/Characters_High5.mp4';

function s9MountClip() {
  const vid = document.getElementById('s9-char-video');
  if (!vid) return;
  vid.setAttribute('aria-hidden', 'true');
  /* Fit-shim properties written unconditionally, including the no-op
     values, so nothing can linger on the element. */
  vid.style.setProperty('--char-scale', 1);
  vid.style.setProperty('--char-shift-y', '0px');
  /* Re-entry safe: only touch src when it actually differs, so returning
     to this screen never restarts a clip that already played and is
     correctly holding its final frame. */
  const cur = vid.getAttribute('src');
  if (!cur || !cur.endsWith(S9_CLIP)) {
    vid.src = S9_CLIP;
    vid.load();
    vid.play().catch(() => {});
  }
}


/* ═══════════════════════════════════════════════════════════
   Screen 11 (#s10) — GuidedMathPracticeScreen, step 1 (variant B1)
   Source: slide 14. See index.html's own comment above #s10 for the
   shape inventory, the run's extent and the bidi evidence.

   NO ANSWER STATE EXISTS on this step: it asks the learner to look, not
   to solve, and its four lettered options are authored content rather
   than capsules. The only thing tracked is whether the exercise has
   appeared, because that is the condition the producer gated Continue
   on ("כפתור המשך יהיה זמין ללחיצה אחרי הופעת התרגיל").

   s10Revealed is seen-once for the page session: once the exercise has
   appeared, coming back shows it already there and Continue already
   open, rather than replaying the entrance and re-locking a screen the
   learner has finished with. */
let s10Revealed = false;
let s10Timer = null;

function s10PrefersReducedMotion() {
  return window.matchMedia && window.matchMedia('(prefers-reduced-motion: reduce)').matches;
}

function s10Unlock() {
  s10Revealed = true;
  const btn = document.getElementById('s10-continue');
  if (btn) { btn.disabled = false; btn.removeAttribute('aria-disabled'); }
}

function s10Enter() {
  const card = document.getElementById('s10-exercise');
  const btn = document.getElementById('s10-continue');
  if (!card) return;
  if (s10Timer) { clearTimeout(s10Timer); s10Timer = null; }

  /* Already seen, or reduced motion: the exercise is simply there and
     Continue is already available — no entrance, no re-lock. */
  if (s10Revealed || s10PrefersReducedMotion()) {
    card.classList.add('is-in');
    s10Unlock();
    return;
  }

  if (btn) btn.disabled = true;
  card.classList.remove('is-in');
  void card.offsetWidth;              // restart the transition cleanly
  card.classList.add('is-in');
  /* Unlock when the entrance has finished — transitionend would also
     work, but a timer cannot be lost if the transition is interrupted by
     a navigation mid-reveal. */
  s10Timer = setTimeout(() => { s10Timer = null; s10Unlock(); }, 500);
}


/* ═══════════════════════════════════════════════════════════
   Screen 12 (#s11) — GuidedMathPracticeScreen, step 2 (variant B2)
   Source: slide 15. See index.html's own comment above #s11 for the
   shape inventory, the production notes and the correct-answer evidence.

   B2 SEMANTICS, per GuidedMathPracticeScreen.md:
   - Selecting IS submitting. One click resolves the step; there is no
     check button, no second attempt and no attempt counter.
   - Resolution is inline: reveal the correct capsule, mark the
     learner's own pick, disable every capsule, unlock Continue.
   - NO feedback popup. Slide 15's note is explicit — "חיווי של התשובה
     הנכונה (מסומנת), ללא משוב מילולי מעבר לזה" — and the template only
     allows a popup where verbal feedback is actually authored.
   - SELECTION and CORRECTNESS are separate states. s11Picked drives the
     Purple/300 ground; the badge is driven by the correct id alone. A
     revealed-correct capsule the learner did not choose therefore never
     wears the selection treatment.
   - A wrong pick is not branded — slide 15 marks only the correct
     answer.

   Re-entry rebuilds from STORED state (which option was picked), not
   from the DOM, so returning to a resolved step restores the learner's
   own mark and not merely "answered". */
const S11 = { correctId: 'o6' };
let s11Picked = null;   // the learner's own choice — drives the selection ground
let s11Done   = false;  // resolved — drives the gate

function s11Opts() { return document.querySelectorAll('#s11 .gmp-opt'); }

function s11Apply() {
  s11Opts().forEach(o => {
    const id = o.dataset.id;
    const picked = s11Done && id === s11Picked;
    const correct = s11Done && id === S11.correctId;
    o.classList.toggle('picked', picked);
    o.classList.toggle('correct', correct);
    o.setAttribute('aria-checked', id === s11Picked ? 'true' : 'false');
    o.disabled = s11Done;
  });
  const btn = document.getElementById('s11-continue');
  if (btn) btn.disabled = !s11Done;
}

function s11Pick(id) {
  if (s11Done) return;          // one action resolves the step; no retry
  s11Picked = id;
  s11Done = true;
  /* The guided item 003 ("נפתור יחד"), not a test (unit owner, 10.11.26): its steps report NO
     answer — Kata declares no question on 003 — and 003 is not in XAPI_EVAL_ITEMS, so it never
     enters a score. The pick is still committed — save it as reportQ used to. */
  try { flushAfterCommit(); } catch (e) {}
  s11Apply();
}

function s11Enter() {
  /* Resolved steps come back resolved: the gate never re-locks and the
     learner never answers twice. */
  s11Apply();
}

/* Enter/Space on a focused capsule selects it — radiogroup keyboard
   a11y, delegated so it covers every capsule without inline handlers. */
document.addEventListener('keydown', e => {
  const isEnter = e.key === 'Enter' || e.code === 'Enter' || e.code === 'NumpadEnter';
  const isSpace = e.key === ' ' || e.key === 'Spacebar' || e.code === 'Space';
  if (!isEnter && !isSpace) return;
  const opt = e.target.closest && e.target.closest('#s11 .gmp-opt');
  if (!opt || opt.disabled) return;
  e.preventDefault();
  s11Pick(opt.dataset.id);
});


/* ═══════════════════════════════════════════════════════════
   Screen 13 (#s12) — GuidedMathPracticeScreen, step 3 (variant B1)
   Source: slide 16. See index.html's own comment above #s12.

   B1 SEMANTICS, per GuidedMathPracticeScreen.md → B1 / D / E:
   - No learner interaction, so no answer state, no resolved flag, no
     gate and nothing to restore on re-entry.
   - The correctness mark on option א is AUTHORED content (the slide's
     highlight + the note), not a reveal — it is present from the moment
     the screen shows and is not driven by any learner action.
   The enter hook only keeps Continue enabled, so a future markup change
   cannot accidentally gate an explanatory step. */
function s12Enter() {
  const btn = document.getElementById('s12-continue');
  if (btn) btn.disabled = false;
}


/* ═══════════════════════════════════════════════════════════
   Screen 14 (#s13) — GuidedMathPracticeScreen, step 4 (variant B1)
   Source: slide 17. See index.html's own comment above #s13. Same
   B1 semantics as screen 13: no state, no gate, nothing to restore. */
function s13Enter() {
  const btn = document.getElementById('s13-continue');
  if (btn) btn.disabled = false;
}


/* ═══════════════════════════════════════════════════════════
   Screen 15 (#s14) — GuidedMathPracticeScreen, step 5 (variant B1)
   Source: slide 18. See index.html's own comment above #s14. Same
   B1 semantics as screens 13–14: no state, no gate, nothing to restore. */
function s14Enter() {
  const btn = document.getElementById('s14-continue');
  if (btn) btn.disabled = false;
}


/* ═══════════════════════════════════════════════════════════
   Screen 16 (#s15) — GuidedMathPracticeScreen, step 6 (variant B2)
   Source: slide 19. See index.html's own comment above #s15.

   Identical B2 semantics to screen 12, and deliberately the same code
   shape: selecting IS submitting, resolution is inline, SELECTION and
   CORRECTNESS are separate states, a wrong pick is not branded, there
   is no feedback popup, and re-entry rebuilds from STORED state. The
   only differences are the option ids and that this step is binary. */
const S15 = { correctId: 'yes' };
let s15Picked = null;   // the learner's own choice — drives the selection ground
let s15Done   = false;  // resolved — drives the gate

function s15Opts() { return document.querySelectorAll('#s15 .gmp-opt'); }

function s15Apply() {
  s15Opts().forEach(o => {
    const id = o.dataset.id;
    const picked = s15Done && id === s15Picked;
    const correct = s15Done && id === S15.correctId;
    o.classList.toggle('picked', picked);
    o.classList.toggle('correct', correct);
    o.setAttribute('aria-checked', id === s15Picked ? 'true' : 'false');
    o.disabled = s15Done;
  });
  const btn = document.getElementById('s15-continue');
  if (btn) btn.disabled = !s15Done;
}

function s15Pick(id) {
  if (s15Done) return;          // one action resolves the step; no retry
  s15Picked = id;
  s15Done = true;
  try { flushAfterCommit(); } catch (e) {}   /* guided item 003: no answer (see s11Pick) */
  s15Apply();
}

function s15Enter() {
  /* Resolved steps come back resolved: the gate never re-locks and the
     learner never answers twice. */
  s15Apply();
}

/* Enter/Space on a focused capsule selects it — same delegated handler
   shape as screen 12's, scoped to this screen's own radiogroup. */
document.addEventListener('keydown', e => {
  const isEnter = e.key === 'Enter' || e.code === 'Enter' || e.code === 'NumpadEnter';
  const isSpace = e.key === ' ' || e.key === 'Spacebar' || e.code === 'Space';
  if (!isEnter && !isSpace) return;
  const opt = e.target.closest && e.target.closest('#s15 .gmp-opt');
  if (!opt || opt.disabled) return;
  e.preventDefault();
  s15Pick(opt.dataset.id);
});


/* ═══════════════════════════════════════════════════════════
   Screen 17 (#s16) — GuidedMathPracticeScreen, step 7 (variant B1)
   Source: slide 20. See index.html's own comment above #s16. Same
   B1 semantics as screens 13–15: no state, no gate, nothing to restore. */
function s16Enter() {
  const btn = document.getElementById('s16-continue');
  if (btn) btn.disabled = false;
}


/* ═══════════════════════════════════════════════════════════
   Screen 18 (#s17) — GuidedMathPracticeScreen, step 8 (variant B2)
   Source: slide 21. See index.html's own comment above #s17.

   Same B2 semantics and the same code shape as screens 12 and 16:
   selecting IS submitting, resolution is inline, SELECTION and
   CORRECTNESS are separate states, a wrong pick is not branded, there
   is no feedback popup, and re-entry rebuilds from STORED state. */
const S17 = { correctId: 'r25' };
let s17Picked = null;   // the learner's own choice — drives the selection ground
let s17Done   = false;  // resolved — drives the gate

function s17Opts() { return document.querySelectorAll('#s17 .gmp-opt'); }

function s17Apply() {
  s17Opts().forEach(o => {
    const id = o.dataset.id;
    const picked = s17Done && id === s17Picked;
    const correct = s17Done && id === S17.correctId;
    o.classList.toggle('picked', picked);
    o.classList.toggle('correct', correct);
    o.setAttribute('aria-checked', id === s17Picked ? 'true' : 'false');
    o.disabled = s17Done;
  });
  const btn = document.getElementById('s17-continue');
  if (btn) btn.disabled = !s17Done;
}

function s17Pick(id) {
  if (s17Done) return;          // one action resolves the step; no retry
  s17Picked = id;
  s17Done = true;
  try { flushAfterCommit(); } catch (e) {}   /* guided item 003: no answer (see s11Pick) */
  s17Apply();
}

function s17Enter() {
  /* Resolved steps come back resolved: the gate never re-locks and the
     learner never answers twice. */
  s17Apply();
}

/* Enter/Space on a focused capsule selects it — same delegated handler
   shape as screens 12 and 16, scoped to this screen's own radiogroup. */
document.addEventListener('keydown', e => {
  const isEnter = e.key === 'Enter' || e.code === 'Enter' || e.code === 'NumpadEnter';
  const isSpace = e.key === ' ' || e.key === 'Spacebar' || e.code === 'Space';
  if (!isEnter && !isSpace) return;
  const opt = e.target.closest && e.target.closest('#s17 .gmp-opt');
  if (!opt || opt.disabled) return;
  e.preventDefault();
  s17Pick(opt.dataset.id);
});


/* ═══════════════════════════════════════════════════════════
   Screen 19 (#s18) — GuidedMathPracticeScreen, step 9 (variant B1)
   Source: slide 22. See index.html's own comment above #s18. Same
   B1 semantics as screens 13–15 and 17: no state, no gate, nothing to
   restore. Both verdicts on this screen are AUTHORED content, present
   from the moment it shows — neither is driven by a learner action. */
function s18Enter() {
  const btn = document.getElementById('s18-continue');
  if (btn) btn.disabled = false;
}


/* ═══════════════════════════════════════════════════════════
   Screen 20 (#s19) — TransitionScreen, the winning beat (slide 23).

   Slide 23's still was an instructional-designer placeholder; the real
   assets are the producer's two "sign" clips, and this screen shows the
   LEARNER'S OWN companion — the bubble says "אני!", so it has to.

   Each clip's ink fractions are MEASURED from its held final frame (the
   raised sign), by column/row histogram on a canvas so edge compression
   noise cannot widen the box. They are published as CSS variables and the
   bubble is computed from them — the same principle as the still poses,
   which anchor their bubbles by measured ink and never by the canvas box.
   Re-measure these if a clip is ever replaced; a stale value here
   silently mis-places the bubble. */
const S19_CLIPS = {
  'character-1': { src: 'assets/video/Character_1_sign.mp4', ink: { l: 0.1781, t: 0.1000, b: 1.0000 } },
  'character-2': { src: 'assets/video/Character_2_sign.mp4', ink: { l: 0.3937, t: 0.0000, b: 0.9778 } }
};

function s19MountClip() {
  const vid = document.getElementById('s19-char-video');
  const wrap = vid && vid.closest('.s1-video-wrap');
  if (!vid || !wrap) return;
  const id = (window.lomdaState && window.lomdaState.selectedCharacter) || DEFAULT_CHARACTER;
  const clip = S19_CLIPS[id] || S19_CLIPS[DEFAULT_CHARACTER];

  /* Publish this clip's own ink so the bubble's tail lands on THIS
     character, at head height, with the deck's own 36px gap. */
  wrap.style.setProperty('--ink-l', clip.ink.l);
  wrap.style.setProperty('--ink-t', clip.ink.t);
  wrap.style.setProperty('--ink-b', clip.ink.b);

  /* Fit-shim properties written unconditionally, including the no-op
     values, so nothing can linger on the element. */
  vid.style.setProperty('--char-scale', 1);
  vid.style.setProperty('--char-shift-y', '0px');

  /* Re-entry safe: only touch src when it actually differs, so returning
     to this screen never restarts a clip that already played and is
     correctly holding its final frame. */
  const cur = vid.getAttribute('src');
  if (!cur || !cur.endsWith(clip.src)) {
    vid.src = clip.src;
    vid.load();
    vid.play().catch(() => {});
  }
}


/* ═══════════════════════════════════════════════════════════
   Screen 21 (#s20) — the warm-up clips (slide 24).

   Same shape as S19_CLIPS: a per-character map of a source and that
   source's MEASURED visible bounds. The screen's CSS does all the
   arithmetic from those bounds, so replacing an asset — a re-export, a
   rename, a different framing — is a data change here and nothing else.

   Bounds are the UNION across five samples of each clip, because the
   pose moves through the 4.01s. Re-measure with
   scratchpad/edge/measure-warmup.cjs if a clip is ever replaced; a stale
   value here silently mis-aligns the character.

   Character-1's file was supplied as "Character_1_warm-up.mp4.mp4" and
   renamed to the single-extension name on 2026-09-23 (final cleanup). */
const S20_CLIPS = {
  'character-1': { src: 'assets/video/Character_1_warm-up.mp4', ink: { l: 0.2930, r: 0.6836, b: 0.8986 } },
  'character-2': { src: 'assets/video/Character_2_warm-up.mp4',     ink: { l: 0.3328, r: 0.6578, b: 0.9417 } }
};

function s20MountClip() {
  const vid = document.getElementById('s20-char-video');
  const wrap = vid && vid.closest('.s1-video-wrap');
  if (!vid || !wrap) return;
  const id = (window.lomdaState && window.lomdaState.selectedCharacter) || DEFAULT_CHARACTER;
  const clip = S20_CLIPS[id] || S20_CLIPS[DEFAULT_CHARACTER];

  /* Publish this clip's own bounds; the CSS turns them into the ground
     line and centre corrections. */
  wrap.style.setProperty('--ink-l', clip.ink.l);
  wrap.style.setProperty('--ink-r', clip.ink.r);
  wrap.style.setProperty('--ink-b', clip.ink.b);

  /* Scale stays a no-op — the two characters are NOT normalised to the
     same size; each keeps the framing it was shot at. Written anyway so
     nothing can linger on the element. */
  vid.style.setProperty('--char-scale', 1);

  /* Re-entry safe: only touch src when it actually differs, so returning
     to this screen never restarts a clip that already played and is
     correctly holding its final pose. */
  const cur = vid.getAttribute('src');
  if (!cur || !cur.endsWith(clip.src)) {
    vid.src = clip.src;
    vid.load();
    vid.play().catch(() => {});
  }
}


/* ═══════════════════════════════════════════════════════════
   Practice-question section results (screens 21-23).
   No progress bar: removed by producer ruling, 2026-09-29.

   ⚠ STRUCTURAL CORRECTION, producer 2026-09-14. This part's practice
   block carries ONE authored question, split into SECTIONS (א / ב / ג)
   that occupy one screen each:

       Question 1 = S21 (א) · S22 (ב) · S23 (ג)

   The א/ב/ג labels in the authored content are SECTION labels, not
   question numbers, and the presenter notes' "שאלה 1/2/3 מתוך 3" counts
   those sections — it is not a scoring boundary. An earlier build read
   them as three independent questions and scored each on its own; this
   model replaces that.

   Consequences encoded below, identical to the sibling part 03-02:
   · the question resolves only when its LAST section has resolved;
   · question correctness is the AND of every section
     (חישוב תשובות נכונות במסע למידה — "success in the question requires
     a correct response to every section of that question");
   · no partial credit, and no 2-of-3 rule.
   ═══════════════════════════════════════════════════════════ */

/* Which screens make up each authored question, in order. */
const QUESTION_SECTIONS = [[21, 22, 23]];

/* Per-SECTION outcome, 'correct' | 'incorrect' | null (unresolved).
   Sections are the only thing a screen writes; every question-level
   value below is derived, never stored, so the two can never disagree. */
const sectionResult = { 21: null, 22: null, 23: null };

/* The question-level result — the scoring rule; nothing displays it.
   null until EVERY section of the question has an outcome. A partially
   answered question is deliberately not a result — it must not show as
   correct, incorrect or complete early. */
function questionResult(qi) {
  const secs = QUESTION_SECTIONS[qi];
  if (!secs) return null;
  if (secs.some(sc => sectionResult[sc] === null)) return null;
  return secs.every(sc => sectionResult[sc] === 'correct') ? 'correct' : 'incorrect';
}

/* Called by a section screen the moment its own SECTION resolves. It
   records the section only — whether the QUESTION has resolved is
   derived by questionResult() above. */
function recordSection(screen, result) {
  sectionResult[screen] = (result === 'correct') ? 'correct' : 'incorrect';
}


/* ═══════════════════════════════════════════════════════════
   Screen 22 (#s21) — ImageHotspotQuestion controller, region variant
   SECTION א of the one popcorn question (slides 25-27); see the section-
   results block above for what this screen's result does and does not resolve.
   Source: slide 25. See index.html's own source-evidence comment above
   #s21 for the shape inventory, the off-canvas popup classification and
   the correct-answer evidence.

   CORRECT = 'c3' ("x : 6x"). Two independent signals agree: the slide's
   own green marker (#45, a no-fill roundRect on the container grid)
   covers the third container, and the correct-feedback popup states the
   expression outright. Never resolved from position.

   ATTEMPTS = 2, the 720 default — slide 25 authors exactly the deck's
   normal 2-attempt shape, one correct popup and one final-wrong popup
   sharing a single explanation body.

   The shared explanation is verbatim from the slide's own two popups
   (#27 / #29), which carry byte-identical bodies and differ only in
   title. Both ratios inside it were MEASURED in the engine rather than
   derived: the authored "6 : 1" is reordered by the bidi algorithm and
   renders physically as 1 : 6, which is the form reproduced here and the
   same one the question text shows.
   ═══════════════════════════════════════════════════════════ */
const S21HS_EXPLAIN = [
  'النسبة بين عدد أكواب الزيت وعدد أكواب الذرة هي <bdi dir="ltr" class="scq-expr">1 : 6‏‎</bdi>.',
  'عامل التوسيع هو <bdi dir="ltr" class="scq-expr">x</bdi>، ولذلك التعبير الجبري سيكون <bdi dir="ltr" class="scq-expr">x : 6x</bdi>.'
];
const S21HS = { correctId: 'c3', maxAttempts: 2 };
/* Titles: correct and final-wrong are the slide's own authored copy. The
   retry title is the canonical fallback — the slide mocks no retry popup,
   which is normal for a 2-attempt question. The final-wrong title carries
   the canonical second line required by _question-template-defaults.md →
   "Terminal-incorrect title carries a fixed second line"; it is APPENDED
   to the authored line, never a replacement. */
const S21HS_POPUP = {
  retry:   { bg: '#ffdbdc', title: 'الإجابة غير صحيحة.', body: ['هل نحاول مرة أخرى؟'] },
  correct: { bg: '#edf8ed', title: 'كل الاحترام، أجبتم بشكل صحيح‎!', body: S21HS_EXPLAIN },
  wrong2:  { bg: '#ffdbdc',
             title: 'هذا خطأ. ممتاز أنكم حاولتم، والآن يمكننا التدقيق.<br>الإجابة الصحيحة تَظهَر الآن.',
             body: S21HS_EXPLAIN }
};

let s21hsSelected  = null;
let s21hsAttempts  = 0;
let s21hsDone      = false;
let s21hsResult    = null;   // 'correct' | 'wrong2'
let s21hsLastWrong = null;   // the hotspot just marked wrong — the retry gate compares against THIS

function s21hsSpot(id) { return document.querySelector('#s21 .hsq-hotspot[data-id="' + id + '"]'); }
function s21hsSpots()  { return document.querySelectorAll('#s21 .hsq-hotspot'); }

/* Single select; clears a prior wrong attempt's marks; re-enables the
   check button ONLY if the newly selected region differs from the one
   just marked wrong (_question-template-defaults.md → "Retry gate" —
   comparison shape for this template: the selected hotspot id). */
function s21hsSelect(id) {
  if (s21hsDone) return;
  s21hsSpots().forEach(o => o.classList.remove('selected'));
  const spot = s21hsSpot(id);
  if (!spot) return;
  spot.classList.add('selected');
  s21hsSelected = id;

  if (s21hsAttempts > 0) {
    document.getElementById('s21hs-popup').classList.add('hidden');
    s21hsSpots().forEach(o => o.classList.remove('correct', 'wrong'));
  }
  const chk = document.getElementById('s21hs-check');
  if (chk) chk.disabled = (id === s21hsLastWrong);
}

function s21hsCheck() {
  if (s21hsDone) { advanceScreen(); return; }
  if (!s21hsSelected) return;

  s21hsAttempts++;
  const correct = s21hsSelected === S21HS.correctId;
  reportQ(['004', 'q1'], correct, correct || s21hsAttempts >= S21HS.maxAttempts, xapiOptLabel(s21hsSpot(s21hsSelected)));

  if (correct) {
    s21hsMark(s21hsSelected, 'correct');
    s21hsResult = 'correct';
    s21hsFinish();
    s21hsShowPopup('correct');
    return;
  }

  if (s21hsAttempts < S21HS.maxAttempts) {
    /* Non-final wrong: mark the chosen region wrong; the correct one is
       NOT revealed yet. */
    s21hsMark(s21hsSelected, 'wrong');
    s21hsLastWrong = s21hsSelected;
    document.getElementById('s21hs-check').disabled = true;
    s21hsShowPopup('retry');
    return;
  }

  /* Final wrong: the chosen region red + ✗ and the correct one green +
     ✓, both shown at once — which is why this template needs no View My
     Answer Toggle. */
  s21hsMark(s21hsSelected, 'wrong');
  s21hsMark(S21HS.correctId, 'correct');
  s21hsResult = 'wrong2';
  s21hsFinish();
  s21hsShowPopup('wrong2');
}

function s21hsMark(id, cls) {
  const spot = s21hsSpot(id);
  if (spot) spot.classList.add(cls);
}

function s21hsFinish() {
  s21hsDone = true;
  s21hsSpots().forEach(o => { o.disabled = true; });
  const chk = document.getElementById('s21hs-check');
  /* The template's own completion label — slide 25 mocks only the
     unresolved "هل إجابتي صحيحة؟" state, so the canonical fallback applies
     (_global-components.md → "Completion state and primary button"). */
  if (chk) { setNavLabel(chk, 'هل نتابع؟'); chk.disabled = false; }
  hideHintButton('s21hs-hint');
  /* Section א has resolved. This records the SECTION's result only — the
     popcorn question stays unresolved until ב and ג resolve too, so a
     correct answer here must never mark the whole question correct, a
     wrong one must never count as an incorrect question on its own, and
     the question must not take a verdict yet. */
  recordSection(21, s21hsResult === 'correct' ? 'correct' : 'incorrect');
}

function s21hsResetInitial() {
  s21hsSelected  = null;
  s21hsAttempts  = 0;
  s21hsDone      = false;
  s21hsResult    = null;
  s21hsLastWrong = null;
  /* A fresh, unresolved section carries no result. */
  sectionResult[21] = null;
  s21hsSpots().forEach(o => {
    o.classList.remove('selected', 'correct', 'wrong');
    o.disabled = false;
  });
  document.getElementById('s21hs-popup').classList.add('hidden');
  document.getElementById('s21hs-hint-overlay').classList.add('hidden');
  const chk = document.getElementById('s21hs-check');
  if (chk) { setNavLabel(chk, 'هل إجابتي صحيحة؟'); chk.disabled = true; }
  /* Hint availability: the template default, visible from the start —
     slide 25 mocks "אפשר רמז?" enabled and neither slide nor note asks
     for the delayed policy. Once available, the hint stays reusable
     indefinitely (2026-09-24 fix — see the identical correction on S8). */
  showHintButton('s21hs-hint');
}

function s21hsShowPopup(type) {
  const popup = document.getElementById('s21hs-popup');
  if (!popup) return;
  const cfg = S21HS_POPUP[type];
  if (!cfg) return;
  popup.style.background = cfg.bg;
  if (popup.classList.contains('hidden')) resetPopupPosition(popup);
  document.getElementById('s21hs-popup-title').innerHTML = cfg.title;
  document.getElementById('s21hs-popup-body').innerHTML =
    cfg.body.map(p => '<p>' + p + '</p>').join('');
  renderPopupAction('s21hs-popup', null, null);   // no action row on this template
  popup.classList.remove('hidden');
  clampPopup(popup);
}

function s21hsOpenHint() {
  xapiRequestedHint('004', 'q1');
  document.getElementById('s21hs-hint-overlay').classList.remove('hidden');
}
function s21hsCloseHint() {
  document.getElementById('s21hs-hint-overlay').classList.add('hidden');
}
function s21hsCloseHintOnBackdrop(e) {
  if (e.target === e.currentTarget) s21hsCloseHint();
}


/* ═══════════════════════════════════════════════════════════
   Screen 23 (#s22) — ValueInputQuestion controller, Variant A
   SECTION ב of the one popcorn question (slides 25-27). This screen
   records its own section result and nothing more; see the section-
   results block above.
   Source: slide 26. See index.html's own source-evidence comment above
   #s22 for the shape inventory and the character-asset evidence.

   CORRECT = 7, stated outright in the presenter note ("الإجابة الصحيحة
   היא 7") and corroborated by the slide's own feedback body, which
   spells out x + 6x = 7x. The field carries the coefficient only; the
   x is a fixed symbol beside it, which is what the note's "נדרש סימון
   מתמטי x קבוע בשדה ההשלמה" asks for.

   ATTEMPTS = 2, the 720 default — slide 26 authors one correct popup
   and one final-wrong popup sharing a single explanation body.

   The final-wrong reveal overwrites the very field the learner typed
   into, so it is implemented as the View My Answer Toggle rather than a
   one-way reveal (_global-components.md → Feedback popup system → "View
   My Answer Toggle"); the learner's raw string is snapshotted before
   the overwrite and shown back exactly as typed.
   ═══════════════════════════════════════════════════════════ */
const S22VIQ = { correct: 7, maxAttempts: 2 };

/* Verbatim from the slide's own two popups (#26 / #28), which carry
   byte-identical bodies and differ only in title. The authored
   paragraph breaks are preserved — ¶3 ends on "היחס הוא" and ¶4 opens
   with the expression — rather than being reflowed into one line. */
const S22VIQ_EXPLAIN = [
  'عندما نوسّع النسبة بـ <bdi dir="ltr" class="scq-expr">x</bdi> مرات، تصبح النسبة',
  '<bdi dir="ltr" class="scq-expr">x : 6x</bdi> ولذلك، يكون مجموع أكواب الزيت والحبوب <bdi dir="ltr" class="scq-expr">7x</bdi> لأن:',
  '<bdi dir="ltr" class="scq-expr">x + 6x = 7x</bdi>'
];
const S22VIQ_POPUP = {
  /* No retry-specific explanation is authored — only the canonical
     fallback title, per _question-template-defaults.md: a missing
     pedagogical explanation is never invented. */
  retry:   { bg: '#ffdbdc', title: 'الإجابة غير صحيحة.', body: ['هل نحاول مرة أخرى؟'] },
  correct: { bg: '#edf8ed', title: 'رائع جدًا!', body: S22VIQ_EXPLAIN },
  wrong2: {
    bg: '#ffdbdc', title: 'أخطأتم، لا بأس، فكل محاولة تقرّبكم من الفهم‎ أكثر.', body: S22VIQ_EXPLAIN,
    action: { label: 'الانتقال إلى إجابتي', view: 'answer' },
    answerView: {
      title: 'الإجابة غير صحيحة.',
      body:  ['يمكن مشاهدة الحل الصحيح مرة أخرى‎.'],
      action: { label: 'الإجابة الصحيحة', view: 'solution' }
    }
  }
};

/* Type 1 icons — 16x16 bare vectors (720-design-guidelines →
   components-icons-states.md → "Type 1"). */
const S22VIQ_ICON_CORRECT = '<svg width="16" height="16" viewBox="0 0 14 12" fill="none" aria-hidden="true"><path d="M4.52036 12C4.12296 12.0002 3.74183 11.8126 3.46106 11.4787L0.258452 7.67744C-0.0861507 7.26814 -0.0861507 6.60467 0.258452 6.19536C0.603165 5.78619 1.16194 5.78619 1.50666 6.19536L4.52036 9.77374L12.4933 0.306878C12.8381 -0.102293 13.3968 -0.102293 13.7415 0.306878C14.0862 0.716179 14.0862 1.37966 13.7415 1.78896L5.57966 11.4787C5.29889 11.8126 4.91776 12.0002 4.52036 12Z" fill="#609E12"/></svg>';
const S22VIQ_ICON_WRONG   = '<svg width="16" height="16" viewBox="0 0 14 14" fill="none" aria-hidden="true"><path d="M13.6743 0.325691C13.4657 0.117151 13.1828 0 12.8879 0C12.5929 0 12.31 0.117151 12.1014 0.325691L7 5.42709L1.8986 0.325691C1.69 0.117151 1.40711 0 1.11215 0C0.817183 0 0.534294 0.117151 0.325691 0.325691C0.117151 0.534294 0 0.817183 0 1.11215C0 1.40711 0.117151 1.69 0.325691 1.8986L5.42709 7L0.325691 12.1014C0.117151 12.31 0 12.5929 0 12.8879C0 13.1828 0.117151 13.4657 0.325691 13.6743C0.534294 13.8829 0.817183 14 1.11215 14C1.40711 14 1.69 13.8829 1.8986 13.6743L7 8.57291L12.1014 13.6743C12.31 13.8829 12.5929 14 12.8879 14C13.1828 14 13.4657 13.8829 13.6743 13.6743C13.8829 13.4657 14 13.1828 14 12.8879C14 12.5929 13.8829 12.31 13.6743 12.1014L8.57291 7L13.6743 1.8986C13.8829 1.69 14 1.40711 14 1.11215C14 0.817183 13.8829 0.534294 13.6743 0.325691Z" fill="#B20010"/></svg>';

let s22viqAttempts     = 0;
let s22viqDone         = false;
let s22viqResult       = null;   // 'correct' | 'wrong2'
let s22viqAnswer       = null;   // the learner's RAW final-attempt string
let s22viqView         = 'solution';
let s22viqLastWrongSig = null;

function s22viqInputEl() { return document.getElementById('s22viq-input'); }
function s22viqFieldEl() { return document.getElementById('s22viq-field'); }

/* type="text" carries no numeric enforcement of its own, so the allowed
   set is filtered on every keystroke and paste — both fire 'input'
   (ValueInputQuestion.md → "Numeric input format"). */
function s22viqOnInput() {
  const el = s22viqInputEl();
  if (!el) return;
  const cleaned = el.value.replace(/[^0-9,.-]/g, '');
  if (cleaned !== el.value) el.value = cleaned;
  if (s22viqDone) return;
  /* A previous attempt's marks clear the moment the learner edits. */
  if (s22viqAttempts > 0) {
    document.getElementById('s22viq-popup').classList.add('hidden');
    s22viqMark(null);
  }
  const chk = document.getElementById('s22viq-check');
  if (chk) chk.disabled = !s22viqFilled() || (s22viqSignature() === s22viqLastWrongSig);
}

function s22viqFilled() {
  const el = s22viqInputEl();
  return !!(el && el.value.trim() !== '');
}
/* Comparison normalises away thousands commas, so adding or removing
   one is not a "genuinely different" answer for the retry gate and not
   a wrong answer for the check. */
function s22viqNumeric() {
  const el = s22viqInputEl();
  if (!el) return NaN;
  return Number(el.value.trim().replace(/,/g, ''));
}
function s22viqSignature() { const n = s22viqNumeric(); return isNaN(n) ? '∅' : String(n); }
function s22viqIsCorrect() { return s22viqNumeric() === S22VIQ.correct; }

function s22viqMark(kind) {
  const field = s22viqFieldEl(), ico = document.getElementById('s22viq-ico');
  if (!field || !ico) return;
  field.classList.remove('viq-correct', 'viq-wrong', 'viq-has-icon');
  ico.innerHTML = '';
  if (kind === 'correct') { field.classList.add('viq-correct', 'viq-has-icon'); ico.innerHTML = S22VIQ_ICON_CORRECT; }
  else if (kind === 'wrong') { field.classList.add('viq-wrong', 'viq-has-icon'); ico.innerHTML = S22VIQ_ICON_WRONG; }
}

function s22viqCheck() {
  if (s22viqDone) { advanceScreen(); return; }
  if (!s22viqFilled()) return;

  s22viqAttempts++;
  reportQ(['004', 'q2'], s22viqIsCorrect(), s22viqIsCorrect() || s22viqAttempts >= S22VIQ.maxAttempts, String((document.getElementById('s22viq-input') || {}).value).trim());
  if (s22viqIsCorrect()) {
    s22viqMark('correct');
    s22viqResult = 'correct';
    s22viqFinish();
    s22viqShowPopup('correct');
    return;
  }

  if (s22viqAttempts < S22VIQ.maxAttempts) {
    s22viqMark('wrong');
    s22viqLastWrongSig = s22viqSignature();
    document.getElementById('s22viq-check').disabled = true;
    s22viqShowPopup('retry');
    return;
  }

  /* Final wrong. The reveal overwrites the field the learner typed into,
     so snapshot the RAW string first — the only place it survives.
     Preserved exactly as typed, never reformatted. */
  s22viqAnswer = s22viqInputEl().value;
  s22viqResult = 'wrong2';
  s22viqView   = 'solution';
  s22viqFinish();
  s22viqApplyView();
}

/* Single owner of the resolved-wrong rendering. */
function s22viqApplyView() {
  if (s22viqAnswer === null) return;
  const el = s22viqInputEl();
  if (s22viqView === 'answer') {
    el.value = s22viqAnswer;
    s22viqMark('wrong');   // the final attempt was wrong, in either view
  } else {
    el.value = String(S22VIQ.correct);
    s22viqMark('correct');
  }
  s22viqShowPopup('wrong2', s22viqView);
}
/* Pure view switch — NEVER routes through s22viqCheck(), never touches
   attempts, resolution or the section result. */
function s22viqToggleView() {
  if (s22viqAnswer === null) return;
  s22viqView = (s22viqView === 'solution') ? 'answer' : 'solution';
  s22viqApplyView();
  /* The view is canonical state the learner chose, and it changes AFTER
     the section resolved — so the resume document, written at
     resolution time, has to be refreshed here or a reload would always
     come back in the solution view. */
  scheduleResumeSave();
}

function s22viqFinish() {
  s22viqDone = true;
  s22viqInputEl().disabled = true;
  const chk = document.getElementById('s22viq-check');
  /* ValueInputQuestion's own completion label is "متابعة" — NOT the
     "هل نتابع؟" the choice templates use; the sibling part 03-02 resolves
     its ValueInput screens the same way. */
  if (chk) { setNavLabel(chk, 'متابعة'); chk.disabled = false; }
  hideHintButton('s22viq-hint');
  /* Section ב only. The question stays unresolved until ג resolves, so
     this must never give the question a verdict. */
  recordSection(22, s22viqResult === 'correct' ? 'correct' : 'incorrect');
}

function s22viqResetInitial() {
  s22viqAttempts     = 0;
  s22viqDone         = false;
  s22viqResult       = null;
  s22viqAnswer       = null;
  s22viqView         = 'solution';
  s22viqLastWrongSig = null;
  const el = s22viqInputEl();
  if (el) { el.value = ''; el.disabled = false; }
  s22viqMark(null);
  document.getElementById('s22viq-popup').classList.add('hidden');
  document.getElementById('s22viq-hint-overlay').classList.add('hidden');
  const chk = document.getElementById('s22viq-check');
  if (chk) { setNavLabel(chk, 'هل إجابتي صحيحة؟'); chk.disabled = true; }
  sectionResult[22] = null;   /* this section is unanswered again */
  /* Hint visible from the start — the template default; slide 26 mocks
     "אפשר רמז?" enabled and neither slide nor note asks for the delayed
     policy. Once available, the hint stays reusable indefinitely
     (2026-09-24 fix — see the identical correction on S8). */
  showHintButton('s22viq-hint');
}

function s22viqShowPopup(type, view) {
  const popup = document.getElementById('s22viq-popup');
  if (!popup) return;
  const base = S22VIQ_POPUP[type];
  if (!base) return;
  const cfg = (view === 'answer' && base.answerView)
            ? Object.assign({}, base, base.answerView)
            : base;
  popup.style.background = cfg.bg;
  /* Re-anchor on a fresh open, AND on every 'wrong2' render (including
     each toggle): the solution and answer bodies differ in height, and a
     height change must never leave the popup pinned away from the bar. */
  if (popup.classList.contains('hidden') || type === 'wrong2') resetPopupPosition(popup);
  /* The terminal-incorrect SOLUTION view carries the canonical second
     line (_question-template-defaults.md); the answer view shows the
     learner's own answer, so it does not. */
  document.getElementById('s22viq-popup-title').innerHTML =
    (type === 'wrong2' && view !== 'answer')
      ? cfg.title + '<br>الإجابة الصحيحة تَظهَر الآن.'
      : cfg.title;
  document.getElementById('s22viq-popup-body').innerHTML =
    (cfg.body || []).map(p => '<p>' + p + '</p>').join('');
  renderPopupAction('s22viq-popup', cfg.action, s22viqToggleView);
  popup.classList.remove('hidden');
  clampPopup(popup);
}

function s22viqOpenHint() {
  xapiRequestedHint('004', 'q2');
  document.getElementById('s22viq-hint-overlay').classList.remove('hidden');
}
function s22viqCloseHint() {
  document.getElementById('s22viq-hint-overlay').classList.add('hidden');
}
function s22viqCloseHintOnBackdrop(e) {
  if (e.target === e.currentTarget) s22viqCloseHint();
}

/* ── The burnt-popcorn clips (slide 26) ──
   Same shape as S19_CLIPS / S20_CLIPS: a per-character map of a source
   and that source's MEASURED visible bounds. The screen's CSS does the
   arithmetic, so replacing an asset is a data change here and nothing
   else.    Bounds are the union across ALL 96 frames of each clip, measured on
   the corrected production files themselves; the pose moves, so a
   five-sample union was not enough. `t` matters as much as `b` here:
   the bubble's placement is derived from the visible HEIGHT, so a stale
   top silently mis-places the bubble as well as the character.
   Re-measure with scratchpad/vid/ink-final.js if a clip is ever
   replaced. Mapping was confirmed by inspecting the clips themselves,
   not inferred from the filenames. */
const S22_CLIPS = {
  'character-1': { src: 'assets/video/Character_1_burnt_popcorn.mp4', ink: { l: 0.3297, r: 0.6695, t: 0.1528, b: 0.9264 } },
  'character-2': { src: 'assets/video/Character_2_burnt_popcorn.mp4', ink: { l: 0.4094, r: 0.5938, t: 0.0306, b: 0.9569 } }
};

function s22MountClip() {
  const vid = document.getElementById('s22-char-video');
  if (!vid) return;
  const id = (window.lomdaState && window.lomdaState.selectedCharacter) || DEFAULT_CHARACTER;
  const clip = S22_CLIPS[id] || S22_CLIPS[DEFAULT_CHARACTER];

  /* Published on the SCREEN, not on the media wrap: two siblings need
     these numbers — the clip's own fit shim and the speech bubble, which
     places itself against THIS character's head. */
  const screen = document.getElementById('s22');
  if (screen) {
    screen.style.setProperty('--ink-l', clip.ink.l);
    screen.style.setProperty('--ink-r', clip.ink.r);
    screen.style.setProperty('--ink-t', clip.ink.t);
    screen.style.setProperty('--ink-b', clip.ink.b);
  }
  /* Scale stays a no-op — the two companions are NOT normalised to one
     size; each keeps the framing it was shot at. */
  vid.style.setProperty('--char-scale', 1);

  /* Re-entry safe: only touch src when it actually differs, so coming
     back never restarts a clip that already played and is correctly
     holding its final pose. */
  const cur = vid.getAttribute('src');
  if (!cur || !cur.endsWith(clip.src)) {
    vid.src = clip.src;
    vid.load();
    vid.play().catch(() => {});
  }
}


/* ═══════════════════════════════════════════════════════════
   Screen 24 (#s23) — ValueInputQuestion controller, Variant A
   SECTION ג of the one popcorn question (slides 25-27), and its LAST
   section. Resolving this screen is what finally resolves the question:
   recordSection(23, …) completes the set, questionResult(0) stops
   returning null, and the question counts as correct only if א, ב and ג
   are all correct. This controller does not compute that — it records
   its own section and nothing else, exactly like #s21 and #s22.
   Source: slide 27. See index.html's own source-evidence comment above
   #s23 for the shape inventory.

   CORRECT = 8: the slide's own feedback body derives it in full
   (6x = 48, x = 8, "יש להוסיף 8 כוסות שמן") and the mock field is
   pre-filled with 8.

   ATTEMPTS = 2, the 720 default. The final-wrong reveal overwrites the
   learner's own value, so it is the View My Answer Toggle, as on #s22.
   ═══════════════════════════════════════════════════════════ */
const S23VIQ = { correct: 8, maxAttempts: 2 };

/* Verbatim from the slide's own two popups (#26 / #28), which carry
   byte-identical bodies and differ only in title. The authored paragraph
   structure is preserved, INCLUDING its two empty paragraphs — rendered
   as spacers ('---'), not as an invented margin. Both ratios were
   measured in the engine, not derived: the authored "6 : 1" renders
   physically as 1 : 6, and "6x = 48" and "x=8" render as authored. */
const S23VIQ_EXPLAIN = [
  'النسبة بين أكواب الزيت وأكواب الذرة هي <bdi dir="ltr" class="scq-expr">1 : 6</bdi>، ولذلك نمثّل الكميات على النحو التالي:',
  '<bdi dir="ltr" class="scq-expr">x : 6x</bdi>',
  '---',
  'معلوم أن هناك 48 كوبًا من الذرة، لذلك نعوّض‎:',
  '<bdi dir="ltr" class="scq-expr">6x = 48</bdi>',
  '<bdi dir="ltr" class="scq-expr">x = 8</bdi>',
  '---',
  'أي، يجب إضافة 8 أكواب من الزيت‎.'
];
const S23VIQ_POPUP = {
  retry:   { bg: '#ffdbdc', title: 'الإجابة غير صحيحة.', body: ['هل نحاول مرة أخرى؟'] },
  correct: { bg: '#edf8ed', title: 'صحيح‎!', body: S23VIQ_EXPLAIN },
  wrong2: {
    bg: '#ffdbdc', title: 'هذا خطأ، وهذا مقبول جدًا، لأننا نستطيع الآن تعلّم شيء جديد‎.', body: S23VIQ_EXPLAIN,
    action: { label: 'الانتقال إلى إجابتي', view: 'answer' },
    answerView: {
      title: 'الإجابة غير صحيحة.',
      body:  ['يمكن مشاهدة الحل الصحيح مرة أخرى‎.'],
      action: { label: 'الإجابة الصحيحة', view: 'solution' }
    }
  }
};

let s23viqAttempts     = 0;
let s23viqDone         = false;
let s23viqResult       = null;
let s23viqAnswer       = null;
let s23viqView         = 'solution';
let s23viqLastWrongSig = null;

function s23viqInputEl() { return document.getElementById('s23viq-input'); }
function s23viqFieldEl() { return document.getElementById('s23viq-field'); }

function s23viqOnInput() {
  const el = s23viqInputEl();
  if (!el) return;
  const cleaned = el.value.replace(/[^0-9,.-]/g, '');
  if (cleaned !== el.value) el.value = cleaned;
  if (s23viqDone) return;
  if (s23viqAttempts > 0) {
    document.getElementById('s23viq-popup').classList.add('hidden');
    s23viqMark(null);
  }
  const chk = document.getElementById('s23viq-check');
  if (chk) chk.disabled = !s23viqFilled() || (s23viqSignature() === s23viqLastWrongSig);
}
function s23viqFilled() { const el = s23viqInputEl(); return !!(el && el.value.trim() !== ''); }
function s23viqNumeric() { const el = s23viqInputEl(); return el ? Number(el.value.trim().replace(/,/g, '')) : NaN; }
function s23viqSignature() { const n = s23viqNumeric(); return isNaN(n) ? '∅' : String(n); }
function s23viqIsCorrect() { return s23viqNumeric() === S23VIQ.correct; }

function s23viqMark(kind) {
  const field = s23viqFieldEl(), ico = document.getElementById('s23viq-ico');
  if (!field || !ico) return;
  field.classList.remove('viq-correct', 'viq-wrong', 'viq-has-icon');
  ico.innerHTML = '';
  if (kind === 'correct') { field.classList.add('viq-correct', 'viq-has-icon'); ico.innerHTML = S22VIQ_ICON_CORRECT; }
  else if (kind === 'wrong') { field.classList.add('viq-wrong', 'viq-has-icon'); ico.innerHTML = S22VIQ_ICON_WRONG; }
}

function s23viqCheck() {
  if (s23viqDone) { advanceScreen(); return; }
  if (!s23viqFilled()) return;

  s23viqAttempts++;
  reportQ(['004', 'q3'], s23viqIsCorrect(), s23viqIsCorrect() || s23viqAttempts >= S23VIQ.maxAttempts, String((document.getElementById('s23viq-input') || {}).value).trim());
  if (s23viqIsCorrect()) {
    s23viqMark('correct');
    s23viqResult = 'correct';
    s23viqFinish();
    s23viqShowPopup('correct');
    return;
  }
  if (s23viqAttempts < S23VIQ.maxAttempts) {
    s23viqMark('wrong');
    s23viqLastWrongSig = s23viqSignature();
    document.getElementById('s23viq-check').disabled = true;
    s23viqShowPopup('retry');
    return;
  }
  s23viqAnswer = s23viqInputEl().value;
  s23viqResult = 'wrong2';
  s23viqView   = 'solution';
  s23viqFinish();
  s23viqApplyView();
}

function s23viqApplyView() {
  if (s23viqAnswer === null) return;
  const el = s23viqInputEl();
  if (s23viqView === 'answer') { el.value = s23viqAnswer; s23viqMark('wrong'); }
  else { el.value = String(S23VIQ.correct); s23viqMark('correct'); }
  s23viqShowPopup('wrong2', s23viqView);
}
function s23viqToggleView() {
  if (s23viqAnswer === null) return;
  s23viqView = (s23viqView === 'solution') ? 'answer' : 'solution';
  s23viqApplyView();
  /* The view is canonical state the learner chose, and it changes AFTER
     the section resolved — so the resume document, written at
     resolution time, has to be refreshed here or a reload would always
     come back in the solution view. */
  scheduleResumeSave();
}

function s23viqFinish() {
  s23viqDone = true;
  s23viqInputEl().disabled = true;
  const chk = document.getElementById('s23viq-check');
  if (chk) { setNavLabel(chk, 'متابعة'); chk.disabled = false; }
  hideHintButton('s23viq-hint');
  /* Section ג. This is the call that completes the set — after it,
     questionResult(0) returns 'correct' only when א, ב and ג are all
     correct, and 'incorrect' for every other completed combination. */
  recordSection(23, s23viqResult === 'correct' ? 'correct' : 'incorrect');
}

function s23viqResetInitial() {
  s23viqAttempts     = 0;
  s23viqDone         = false;
  s23viqResult       = null;
  s23viqAnswer       = null;
  s23viqView         = 'solution';
  s23viqLastWrongSig = null;
  const el = s23viqInputEl();
  if (el) { el.value = ''; el.disabled = false; }
  s23viqMark(null);
  document.getElementById('s23viq-popup').classList.add('hidden');
  document.getElementById('s23viq-hint-overlay').classList.add('hidden');
  const chk = document.getElementById('s23viq-check');
  if (chk) { setNavLabel(chk, 'هل إجابتي صحيحة؟'); chk.disabled = true; }
  sectionResult[23] = null;
  /* Hint visible from the start, reusable indefinitely once available
     (2026-09-24 fix — see the identical correction on S8). */
  showHintButton('s23viq-hint');
}

function s23viqShowPopup(type, view) {
  const popup = document.getElementById('s23viq-popup');
  if (!popup) return;
  const base = S23VIQ_POPUP[type];
  if (!base) return;
  const cfg = (view === 'answer' && base.answerView)
            ? Object.assign({}, base, base.answerView)
            : base;
  popup.style.background = cfg.bg;
  if (popup.classList.contains('hidden') || type === 'wrong2') resetPopupPosition(popup);
  document.getElementById('s23viq-popup-title').innerHTML =
    (type === 'wrong2' && view !== 'answer')
      ? cfg.title + '<br>الإجابة الصحيحة تَظهَر الآن.'
      : cfg.title;
  /* '---' is an authored EMPTY paragraph, kept as real spacing. */
  document.getElementById('s23viq-popup-body').innerHTML =
    (cfg.body || []).map(p => p === '---' ? '<div class="popup-para-gap"></div>' : '<p>' + p + '</p>').join('');
  renderPopupAction('s23viq-popup', cfg.action, s23viqToggleView);
  popup.classList.remove('hidden');
  clampPopup(popup);
}

function s23viqOpenHint() {
  xapiRequestedHint('004', 'q3');
  document.getElementById('s23viq-hint-overlay').classList.remove('hidden');
}
function s23viqCloseHint() { document.getElementById('s23viq-hint-overlay').classList.add('hidden'); }
function s23viqCloseHintOnBackdrop(e) { if (e.target === e.currentTarget) s23viqCloseHint(); }


/* Rebuild #s21's resolved marks from its flags. Idempotent and
   mirrors exactly what its own finish path produced; resetScreenState(21)
   runs it on every entry to a resolved screen. */
function s21hsRestoreMarks() {
  if (!s21hsDone) return;
  s21hsSpots().forEach(function (o) {
    o.classList.remove('selected', 'correct', 'wrong');
    o.disabled = true;
  });
  if (s21hsSelected) {
    const sel = s21hsSpot(s21hsSelected);
    if (sel) sel.classList.add('selected');
  }
  if (s21hsResult === 'correct') {
    s21hsMark(s21hsSelected, 'correct');
  } else {
    s21hsMark(s21hsSelected, 'wrong');
    s21hsMark(S21HS.correctId, 'correct');
  }
  const chk = document.getElementById('s21hs-check');
  if (chk) { setNavLabel(chk, 'هل نتابع؟'); chk.disabled = false; }
  hideHintButton('s21hs-hint');
}

/* Cross-part navigation. init() must run AFTER this part's own initial
   screen setup: if this page was entered via Back it jumps straight to
   the screen the learner left from, deferred to a microtask so the
   landing screen's resetScreenState() cannot hit a temporal dead zone.
   Answer state is no longer snapshotted here: it travels in the resume
   state document (PART_SCREENS registry below, unit-js/45-painters.js). */
PartNav.init({
  part: '01',
  getScreen: function () { return currentScreen; },
  goTo: goTo
});

/* ═══════════════════════════════════════════════════════════
   xAPI seams — read at call time by unit-js/20-xapi.js, 22-scoring.js and 50-loader.js.
   Ids must match metadata/methodica-math-ratio-03-01.json byte-for-byte (the loader
   checks XAPI_COMP_ID against METADATA.id on every load; _test/verify-static.js checks the map).
   ═══════════════════════════════════════════════════════════ */
const XAPI_COMP_SLUG     = 'methodica-math-ratio-03-01';
const XAPI_COMP_ID       = XAPI_ID_PREFIX + XAPI_COMP_SLUG + '/';
const XAPI_METADATA_FILE = '../metadata/' + XAPI_COMP_SLUG + '.json';
/* Items carrying a code-graded question (each declared question is reported by its screen). */
const XAPI_EVAL_ITEMS    = { '002': 1, '004': 1 };   /* 003 is the guided item: no question, no score (10.11.26) */
/* Items scored as ONE question: success only when every section is right (QA F-6, 22-scoring.js).
   004 = QUESTION_SECTIONS [21, 22, 23]. */
const XAPI_ITEM_ALL_CORRECT = { '004': 1 };
/* screen → [item suffix, page within the item]; null = no catalogue item. Exactly TOTAL_SCREENS
   keys, 0-based (keep it a literal object — the static check reads it). */
const SCREEN_TO_SUBCONTENT = {
  0: ['001', 1],
  1: ['001', 2],
  2: ['001', 3],
  3: ['001', 4],
  4: ['002', 1],
  5: ['002', 2],
  6: ['002', 3],
  7: ['002', 4],
  8: ['002', 5],
  9: ['003', 1],
  10: ['003', 2],
  11: ['003', 3],
  12: ['003', 4],
  13: ['003', 5],
  14: ['003', 6],
  15: ['003', 7],
  16: ['003', 8],
  17: ['003', 9],
  18: ['003', 10],
  19: ['003', 11],
  20: ['004', 1],
  21: ['004', 2],
  22: ['004', 3],
  23: ['004', 4]
};

/* The warm-up outcome does not gate this part (approved 2026-09-14), so the component passes;
   the score is the share of its catalogued questions answered correctly. */
function partGateResult() { const sc = partScaled(); return sc == null ? { success: true } : { success: true, score: { scaled: sc } }; }
function lastScreenButton() { return currentScreen === PART_LAST ? document.getElementById('s23viq-check') : null; }

/* ═══════════════════════════════════════════════════════════
   Resume painters — unit-js/45-painters.js reads this registry.
   What each screen saves (vars + the fields that hold state) and how a screen restored from the
   state document is drawn again on the fresh markup: init() is the screen's own ResetInitial,
   paint() replays the template's own terminal look (Mark / Finish / ShowPopup / ApplyView) or the
   retry look, without counting an attempt and without reporting. In-session re-entry still goes
   through resetScreenState(), which now resets only a pristine screen.
   Gates are state too (guide §8.8 rule 11): the widget (2), the comic (5), the flip cards (7) and
   the exercise reveal (10) save their flag and repaint the open Continue. The companion character
   is NOT here — it lives in the unit document (setUnitCharacter, 40-resume.js) and s0's own
   resetScreenState branch re-applies it.
   ═══════════════════════════════════════════════════════════ */
const PART_STATE_VARS = ['sectionResult'];
const PART_SCREENS = {
  2: {
    vars: ['s2WidgetCompleted'],
    pristine: () => !s2WidgetCompleted,
    init: () => {},
    paint: () => { s2WidgetCompleted = false; unlockS2Continue(); }   // the unlock is idempotent on the flag
  },
  5: {
    vars: ['cmcSequenceDone'],
    pristine: () => !cmcSequenceDone,
    init: cmcClearTimers,
    paint: cmcShowAll                                                  // finished state, no timers, Continue open
  },
  7: {
    vars: ['frcSeen', 'frcFace', 'frcDone', 'frcGhSeen'],
    pristine: () => !frcGhSeen && !frcSeen.some(Boolean),
    init: () => {},
    paint: () => {
      document.querySelectorAll('#s7 .frc-card').forEach(card => frcApplyFace(card, !!frcFace[+card.dataset.index]));
      if (frcGhSeen) frcGhStop();                                      // the learner found the gesture already
      frcCheckUnlock();
    }
  },
  8: {
    vars: ['s8scSelected', 's8scAttempts', 's8scDone', 's8scResult', 's8scLastWrong'],
    pristine: () => !s8scSelected && !s8scAttempts && !s8scDone,
    init: s8scResetInitial,
    paint: () => {
      const done = s8scDone, sel = s8scSelected;
      s8scDone = false;
      if (sel) s8scSelect(sel);                    // selection look + the retry gate
      if (done) {
        s8scOpt(sel).classList.add(s8scResult === 'correct' ? 'correct' : 'wrong');
        if (s8scResult !== 'correct') s8scOpt(S8SC.correctId).classList.add('correct');
        s8scFinish();
        s8scShowPopup(s8scResult);
      } else if (s8scAttempts && sel === s8scLastWrong) {
        s8scOpt(sel).classList.add('wrong');
      }
    }
  },
  10: {
    vars: ['s10Revealed'],
    pristine: () => !s10Revealed,
    init: () => {},
    paint: s10Enter                                                    // revealed → exercise shown, Continue open, no timer
  },
  11: {
    vars: ['s11Picked', 's11Done'],
    pristine: () => !s11Picked && !s11Done,
    init: () => {},
    paint: s11Apply                                                    // rebuilds marks + gate from the stored pick
  },
  15: {
    vars: ['s15Picked', 's15Done'],
    pristine: () => !s15Picked && !s15Done,
    init: () => {},
    paint: s15Apply
  },
  17: {
    vars: ['s17Picked', 's17Done'],
    pristine: () => !s17Picked && !s17Done,
    init: () => {},
    paint: s17Apply
  },
  21: {
    vars: ['s21hsSelected', 's21hsAttempts', 's21hsDone', 's21hsResult', 's21hsLastWrong'],
    pristine: () => !s21hsSelected && !s21hsAttempts && !s21hsDone,
    init: s21hsResetInitial,
    paint: () => {
      const done = s21hsDone, sel = s21hsSelected;
      s21hsDone = false;
      if (sel) s21hsSelect(sel);                   // selection look + the retry gate
      if (done) {
        s21hsMark(sel, s21hsResult === 'correct' ? 'correct' : 'wrong');
        if (s21hsResult !== 'correct') s21hsMark(S21HS.correctId, 'correct');
        s21hsFinish();
        s21hsShowPopup(s21hsResult);
      } else if (s21hsAttempts && sel === s21hsLastWrong) {
        s21hsMark(sel, 'wrong');
      }
    }
  },
  22: {
    vars: ['s22viqAttempts', 's22viqDone', 's22viqResult', 's22viqAnswer', 's22viqView', 's22viqLastWrongSig'],
    inputs: ['s22viq-input'],
    pristine: () => !s22viqAttempts && !s22viqDone && resumeFieldsEmpty(PART_SCREENS[22].inputs),
    init: s22viqResetInitial,
    paint: () => {
      if (s22viqDone) {
        s22viqDone = false;
        if (s22viqResult === 'correct') { s22viqMark('correct'); s22viqFinish(); s22viqShowPopup('correct'); }
        else { s22viqFinish(); s22viqApplyView(); }
        return;
      }
      s22viqOnInput();                             // check button through the live predicate (filled + retry gate)
      if (s22viqAttempts && s22viqSignature() === s22viqLastWrongSig) s22viqMark('wrong');
    }
  },
  23: {
    vars: ['s23viqAttempts', 's23viqDone', 's23viqResult', 's23viqAnswer', 's23viqView', 's23viqLastWrongSig'],
    inputs: ['s23viq-input'],
    pristine: () => !s23viqAttempts && !s23viqDone && resumeFieldsEmpty(PART_SCREENS[23].inputs),
    init: s23viqResetInitial,
    paint: () => {
      if (s23viqDone) {
        s23viqDone = false;
        if (s23viqResult === 'correct') { s23viqMark('correct'); s23viqFinish(); s23viqShowPopup('correct'); }
        else { s23viqFinish(); s23viqApplyView(); }
        return;
      }
      s23viqOnInput();
      if (s23viqAttempts && s23viqSignature() === s23viqLastWrongSig) s23viqMark('wrong');
    }
  }
};
