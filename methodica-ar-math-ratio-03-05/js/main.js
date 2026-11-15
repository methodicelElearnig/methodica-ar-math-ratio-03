'use strict';

/* ═══════════════════════════════════════════════════════════
   js/main.js  —  methodica-ar-math-ratio-03-05
   720 reference engine (canvas 1280×710) + Screen 00
   (TransitionScreen — Peak Question A launch).

   Source: the approved production script
   יעד 1.3_יחס וביטויים אלגבריים_להפקה_06092026.pptx, two levels above
   this part's folder. Slide identity re-derived THIS session from a
   fresh extraction of the .pptx (the file's mtime advances between
   sessions — it currently reads 2026-09-14 — so nothing was trusted
   from an earlier extraction): presentation.xml's <p:sldIdLst> r:id →
   ppt/_rels/presentation.xml.rels gives order 63 → slide63.xml and
   order 64 → slide64.xml (1:1 with no offset across all 83 slides —
   derived, not assumed). Presenter notes resolved through each slide's
   OWN .rels: slide64 → rId3 → notesSlides/notesSlide58.xml (part-name
   numbering is NOT slide numbering — pptx-script-extraction), whose
   cached slide-number field reads "64".

   PART BOUNDARY, from the deck's own component-divider slides:
     order 63 — "רכיב 5" / "שאלת שיא א" / "methodica-ar-math-ratio-03-05"
     order 73 — "רכיב 6" / "שאלת שיא ב" / "methodica-ar-math-ratio-03-06"
   Part 05 is therefore slides 64-72. Slide 63 is production-only and
   is not a screen.

   Screen roster (built so far):
     S0 — TransitionScreen, Peak Question A launch (slide 64).
     S1 — FullImageWithTextOverlay, bakery set-up (slide 65).
     S2 — ValueInputQuestion (Variant C), peak sub-part 1 (slide 66).
     S3 — FullImageWithTextOverlay, influencer clip (slide 67). Its
          authored video + watch-gate are NOT built: the clip does not
          exist yet. See the screen comment in index.html.
     S4 — SingleChoiceQuestion, peak sub-part 2 (slide 68), with the
          scroll-gated long feedback its own note asks for.
     S5 — FullImageWithTextOverlay, festival order (slide 69).
   Slides 70-72 are NOT built yet and are out of this step's scope.

   SOURCE RE-CHECKED for S1: the .pptx changed between the S0 pass and
   the S1 pass this session (7,732,625 → 7,717,974 bytes, 08:31 → 10:14),
   so the whole deck was re-extracted and slide identity re-derived by
   content rather than carried over. The producer's edits land in slides
   58-62 (Part 04's unbuilt tail) and notesSlide67 (slide 74, Part 06's
   launch screen). Part 05's own range is byte-identical, and the
   dividers still bound it at 64-72.

   S0 ported from the same-lomda sibling methodica-ar-math-ratio-03-04's
   own S0 — the identical screen role (component kickoff transition,
   slide 51 in the same deck) — which took it from -03 (slide 39), which
   took it from -02 (slide 29), per 720-templates →
   references/TransitionScreen.md + references/_global-components.md.
   Engine plumbing (scaleApp / goTo / resetScreenState / keyboard nav /
   setNavLabel / dev bridge) ported from methodica-ar-math-ratio-03-04,
   byte-identical apart from this part's own screen roster and the two
   zoom-modal hooks noted in goTo(), so every part of this lomda behaves
   the same. Cross-part Back uses this part's own js/partnav.js
   (PartNav).

   NO REPORTING LAYER. This file emits no SCORM, xAPI, postMessage or
   Ministry-of-Education outcome signal of any kind. Outcome reporting
   is owned externally by Vadim's integration layer and is outside
   lomda-production scope.
   ═══════════════════════════════════════════════════════════ */


/* ─── S5 image asset ────────────────────────────────────────
   The file slide 69's note names — "שם התמונה:<מאפייה טלפון>". Taken
   from the note, not guessed from the copy: the text talks about a food
   festival and פסטיבל אוכל.jpg is also supplied in this part, so a guess
   would have been easy and wrong. Verified against the deck's own
   embedded copy (image73.jpeg, 1281×855): same photograph, and the
   supplied file is the higher-resolution original (1535×1024), no
   watermark.

   The slide's srcRect crop is reproduced STRUCTURALLY in css/style.css
   (.fito-img--framed), not with object-position — see the block comment
   on this screen in index.html for why cover cannot express it here. */
const S5_IMAGE = {
  src: 'assets/img/מאפייה טלפון.jpg',
  alt: 'الخبازة سمر تقف في المخبز وتتحدث بالهاتف، وخلفها رفوف ممتلئة بأرغفة خبز طازجة'
};

/* S5 — FullImageWithTextOverlay. Stateless, same contract as S1 and S3. */
function s5MountImage() {
  const img = document.getElementById('s5-img');
  if (!img) return;
  if (!img.getAttribute('src') || !img.getAttribute('src').endsWith(S5_IMAGE.src)) {
    img.src = S5_IMAGE.src;
  }
  img.alt = S5_IMAGE.alt;
}

/* ─── Constants ─────────────────────────────────────────────
   NINE screens (S0-S8); S8 (slide 72) is the last one. goTo() bounds-checks
   against this, so the end-of-part crossing does not go through goTo():
   advanceScreen() hands S8 to crossToNextPart() (PartNav.forward, see
   NEXT_PART below). 720-templates → _global-components.md → "Advance
   gates" warns that a missing cross-part redirect looks like working
   code — this one is implemented. */
const TOTAL_SCREENS = 9;  // S0-S8 — S8 (slide 72) is the last screen of part 05


/* ─── S0 companion-character assets ─────────────────────────
   Explicit ID → asset mapping. The two clips were INVENTORIED ON DISK
   first and are referenced by their real filenames — the script names
   no media file for this slide, so nothing here is assumed from the
   deck. Files present in assets/video/ (and only these two):
       Character_1_Running.mp4   1,149,224 bytes
       Character_2_Running.mp4   1,139,716 bytes

   LOGICAL MAPPING confirmed by DECODING AND VIEWING frames from each
   clip (t = 0.0 / 1.0 / 2.0 / 3.0 / 3.95 s plus the true final frame),
   never by reading the filenames or guessing from colour — resolving a
   character by filename is forbidden regardless (720-templates →
   _global-components.md → Companion character system). Result matches
   methodica-ar-math-ratio-03-01's own CHARACTER_ASSETS (same lomda, same
   two companions):
       character-1 = the ORANGE/CORAL round one-eyed character
                     → Character_1_Running.mp4  (verified visually)
       character-2 = the TALL YELLOW character with the purple headband
                     → Character_2_Running.mp4  (verified visually)
   Both wear the same purple headband / purple tank / teal shorts /
   sneakers running kit, so the pair is a matched set for this screen.

   PRESENTER NOTE MATCH — "הדמות נראית בתחילת מרוץ" ("the character is
   seen at the start of a race"). Both clips are running-motion clips in
   athletic kit; clip 1 ends with the orange character planted upright
   and facing forward (a start-line ready pose), clip 2 ends with the
   yellow character in a forward running lunge. Both are meaningful
   final frames, which is what the end-state contract requires.

   TECHNICAL INVENTORY — probed, identical across the pair, no
   discrepancy to report:
       container/codec  MP4 / H.264 High, yuv420p progressive
       resolution       1280×720 both
       duration         4.01 s both
       frame rate       24 fps both
       audio            AAC-LC 48 kHz stereo 128 kb/s — PRESENT in both
   The audio track is irrelevant here and is not a defect: the element
   is `muted`, which is also what lets autoplay run at all.

   BACKGROUND vs THE LOMDA'S #FFFFFF — CORRECTED IN THE ASSET FILES
   THEMSELVES, 2026-09-14, on the producer's ruling. As supplied, neither
   clip was composited on true white: measured across all 96 frames of
   each, the flat background sat at RGB 252-254 and the frame boundary
   was 69.7% / 99.8% not-exactly-255. That read as a faint panel edge
   against this screen's #FFFFFF and was NOT approved.

   Both files were corrected and replaced under their existing names.
   NO CSS correction of any kind is used or permitted here — no filter,
   no mix-blend-mode, no mask, no overlay, no panel behind the video.
   Method and full before/after validation: docs/asset-corrections.md.
   In summary, on the decoded final encodes, all 96 frames:
       flat background below 255   55.3% -> 4.0%   (clip 1)
                                   99.4% -> 3.2%   (clip 2)
       plate-only frame boundary   69.7% -> 0.008% (clip 1)
       not exactly 255             99.8% -> 0.70%  (clip 2, and that
                                   residual is the character's own hair
                                   where it reaches the top frame edge,
                                   not plate — see the doc)
       character mean abs delta    1.336 / 1.387, against an encoder-only
                                   control floor of 1.334 / 1.377, i.e.
                                   the correction's own contribution is
                                   0.002 / 0.010 — nothing measurable.
   Chroma was never written, and the per-frame background level is now
   ~10x more temporally stable than the originals (no flicker).

   NO `fit` ENTRY — measured, not assumed, and deliberately so.
   TransitionScreen.md → "C" requires a shim to be derived from measured
   ink bounds of THIS clip pair, never by analogy. Ink bounds measured
   on decoded frames (luma threshold 30/255), as fractions of the 720px
   frame height. RE-MEASURED from the SHIPPED corrected files after the
   background correction and unchanged from the originals to within 1px
   on every sample — the correction moved no framing, so this decision
   still rests on the assets actually in the build:

       clip                    frame 0              final frame
       Character_1_Running     top .121  bot .910   top .160  bot .912
                               ink height .789      ink height .753
       Character_2_Running     top .015  bot .954   top .082  bot .947
                               ink height .939      ink height .865

   The raw final-frame ink-height gap is 13%, which on its own would sit
   above the template's "visibly noticeable" band — but the number is
   not a framing mismatch and must not be shimmed away:
     1. The two clips hold DIFFERENT POSES. Clip 1 ends upright; clip 2
        ends in a deep forward running lunge, which geometrically
        compresses its vertical ink. Ink height across different poses
        is not a like-for-like scale comparison, which is the only thing
        the shim rule is about.
     2. The residual runs in the same direction as the two companions'
        GENUINE build difference (character-1 is the short round one,
        character-2 the tall one) and is far SMALLER than it. The
        approved sibling methodica-ar-math-ratio-03-01 measured that
        authored difference from the standing PNGs and cross-checked it
        against the deck's own two-character scene on slide 4: the round
        character is 0.774-0.778 of the tall one's height, and it
        deliberately preserves that gap. Here the same ratio measures
        .753/.865 = 0.870 — the clips already show the pair CLOSER in
        height than they genuinely are, so scaling character-1 up would
        equalise a real height difference in the wrong direction.
   Each character therefore keeps its own natural proportions and its
   own authored framing: `object-fit: contain` on a 16:9 box showing a
   16:9 source is an exact fit, so neither clip is cropped or stretched.

   FRAMING DIFFERENCE — reported, deliberately preserved. The two clips
   do not frame their character identically: character-1's ink is
   horizontally centred (ink centre .500 of frame width) while
   character-2's sits left of centre (ink centre .439) and its hair
   reaches the very top of the frame (top ink as high as .004 during the
   run). Both stay fully inside the frame — nothing is clipped at any
   point in either clip — so this is the assets' own authored framing,
   not a defect, and it is left exactly as supplied.

   The learner only ever sees ONE of the two clips, so there is no
   side-by-side comparison for any of these differences to show up in. */
const S0_CHARACTER = {
  'character-1': {
    video: 'assets/video/Character_1_Running.mp4',
    alt: 'الشخصية البرتقالية المستديرة ذات العين الواحدة، بملابس رياضية عند بداية سباق'
  },
  'character-2': {
    video: 'assets/video/Character_2_Running.mp4',
    alt: 'الشخصية الصفراء الطويلة ذات عصابة الرأس البنفسجية، بملابس رياضية عند بداية سباق'
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
   prohibited. Both clips were decoded to their true final frame during
   build verification and each holds a meaningful pose (see the
   S0_CHARACTER note), so no `freezeAt` is needed either. The element
   carries no `loop` attribute, so playback simply ends and the last
   painted frame stays on screen. */

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


/* ─── S1 image asset ────────────────────────────────────────
   FullImageWithTextOverlay mounts src/alt from a map, never from the
   markup (720-templates → FullImageWithTextOverlay.md → "A"), so
   swapping the photo is a one-line change here.

   The file is the one slide 65's presenter note names —
   "שם התמונה:<מאפייה>" — and it was verified against the deck's own
   embedded copy (image70.jpeg, 1282×855) before use: same photograph,
   same 3:2 framing, no watermark on either, and the supplied
   assets/img/מאפייה.jpg is the higher-resolution original (1536×1024).
   The embedded copy is the producer's reference, not the runtime file.

   The note supplies no alt text, so the alt below is authored to
   describe what the photograph actually shows, matching how every other
   image alt in this lomda is written. */
const S1_IMAGE = {
  src: 'assets/img/מאפייה.jpg',
  alt: 'خبازة تقف خلف طاولة خشبية في المخبز وتعجن العجين بالطحين، وفي الخلفية رفوف عليها أرغفة خبز طازجة'
};

/* S1 — FullImageWithTextOverlay. Stateless: no answer, no attempt, no
   feedback to restore, so re-entry is just a re-mount (template
   contract → "D. Navigation and reuse"). Idempotent — the src is only
   written when it differs, so returning to the screen never re-fetches
   an image the browser already has decoded. */
function s1MountImage() {
  const img = document.getElementById('s1-img');
  if (!img) return;
  if (!img.getAttribute('src') || !img.getAttribute('src').endsWith(S1_IMAGE.src)) {
    img.src = S1_IMAGE.src;
  }
  img.alt = S1_IMAGE.alt;
}

/* S2's contextual photo — same asset-map pattern, same idempotence.
   S2_IMAGE is declared with the rest of the S2 block further down. */
function s2MountImage() {
  const img = document.getElementById('s2-img');
  if (!img) return;
  if (!img.getAttribute('src') || !img.getAttribute('src').endsWith(S2_IMAGE.src)) {
    img.src = S2_IMAGE.src;
  }
  img.alt = S2_IMAGE.alt;
}

/* ─── S3 instructional video ─────────────────────────────────
   Supplied 2026-09-14, replacing the interim still image and closing
   the part's oldest pending item. Inventoried before wiring, from the
   container and then from the decoded element in Edge:

       assets/video/Animate_smartphone_video_content.mp4
       1,791,844 bytes · isom/avc1 · H.264 video + AAC audio
       1280 × 720 (16:9) · 10.005 s · stereo 48 kHz
       470 KB of audio decoded on playback — a real track, not an
       empty one

   Content checked on real decoded frames at 0 / 1.5 / 3.5 / 5.5 / 7.5 /
   9.9 s: a hand holding a phone playing a baking influencer's clip
   ("איך תהפכו את הלחם לעשיר יותר בסיבים?"), which is exactly what this
   screen's own card describes.

   INSTRUCTIONAL, not decorative. It is learner-facing content, so the
   contract is figma-lomda-builder → SKILL.md → "Non-negotiable video
   rules": native controls, trimmed but never hidden; sound on; a
   removable play overlay; and a watch gate kept independent of the
   controls. The decorative-character contract in _global-components.md
   (no sound, autoplay, no controls) governs S0's running clips and
   deliberately does NOT govern this one.

   REDUCED MOTION: nothing here autoplays or animates on its own, so
   there is no motion to reduce and nothing to suppress — the clip only
   ever starts from the learner's own gesture. Equally, reduced motion
   must never mute or block it: it is content, not decoration. */
/* THE CROPPED DERIVATIVE, not the raw supply.

   The clip as delivered is 1280×720 and its bottom strip carries a mock
   app navigation bar whose Hebrew is malformed ("ררחסית", "מסריות").
   Hiding it in CSS only worked while the <video> stayed inside this
   screen's own frame — a native fullscreen, a picture-in-picture window
   or any other player would have shown it again. On the producer's
   instruction (2026-09-15) the strip is now gone from the ASSET:

     Animate_smartphone_video_content.mp4              1280×720  source — removed from the tree 2026-09-23 (in git)
     Animate_smartphone_video_content_cropped_1280x548.mp4       what this screen loads

   rows 0..548 kept (measured across the whole clip: the strip never
   rises above row 554 in any frame), audio stream-copied, duration,
   frame rate and A/V sync unchanged, H.264 High / AAC so it plays
   everywhere.

   ARABIC BUILD: the file under this name is the Arabic clip (a clean 16:9 phone frame with an Arabic app bar,
   14.04 s, sound), NOT a 1280×548 crop; the box's object-fit: cover trims it on screen as before, fullscreen
   shows the whole frame. Supplied 1920×1080 at 10 Mbps (18.2 MB); re-encoded 2026-11-15 to 1280×720, H.264 High
   CRF 23 (maxrate 2 Mb/s), AAC 128k 48 kHz stereo, +faststart: 846,530 bytes, SSIM 0.993 against the source at
   1280×720. The supplied file is in git history (cceb5dd). */
const S3_VIDEO = {
  src: 'assets/video/Animate_smartphone_video_content_cropped_1280x548.mp4',
  label: 'فيديو لمؤثر في مجال الخَبز حول إثراء الخبز بالألياف'
};

/* Latched once the clip has genuinely reached its end. Survives
   navigation within the loaded page, which is what keeps the screen
   resolved on return. */
let s3vidDone = false;

function s3Video() { return document.getElementById('s3-video'); }
function s3PlayBtn() { return document.getElementById('s3-play'); }
function s3Continue() { return document.getElementById('s3-continue'); }
function s3Stage() { return document.getElementById('s3-stage'); }
function s3FsBtn() { return document.getElementById('s3-fs'); }

function s3MountVideo() {
  const v = s3Video();
  if (!v) return;
  if (!v.getAttribute('src') || !v.getAttribute('src').endsWith(S3_VIDEO.src)) {
    v.src = S3_VIDEO.src;
  }
  v.setAttribute('aria-label', S3_VIDEO.label);
  /* Offer the control only where the API actually exists, so it can
     never be a button that does nothing. */
  const b = s3FsBtn();
  if (b) {
    const ok = !!(document.fullscreenEnabled || document.webkitFullscreenEnabled);
    b.classList.toggle('hidden', !ok);
  }
  s3SyncFullscreenBtn();
}

/* ─── S3 fullscreen ─────────────────────────────────────────
   The fullscreen target is the WRAPPER (#s3-stage), never the <video>.
   The clip is only half of this screen's content: the Hebrew card
   explains it, and a native per-element fullscreen would have taken the
   picture and left the words behind. So the native control is
   suppressed (`controlslist="… nofullscreen"` in the markup) and this
   pair of functions drives the wrapper instead.

   The card is NOT burnt into the video — it stays real HTML text, so it
   remains accessible and editable, and css/style.css sizes it in
   container units so it keeps its approved position over the picture at
   any fullscreen size. */
function s3FsElement() {
  return document.fullscreenElement || document.webkitFullscreenElement || null;
}
function s3IsFullscreen() {
  const st = s3Stage();
  return !!st && s3FsElement() === st;
}

/* Set while navigation is what closed fullscreen, so the exit handler
   does not yank focus back to a button on a screen the learner has just
   left. */
let s3fsNavExit = false;

function s3ExitFullscreen() {
  if (!s3FsElement()) return;
  const ex = document.exitFullscreen || document.webkitExitFullscreen;
  if (!ex) return;
  try {
    const p = ex.call(document);
    if (p && p.catch) p.catch(() => {});
  } catch (e) {}
}

function s3ToggleFullscreen() {
  const st = s3Stage();
  if (!st) return;
  if (s3IsFullscreen()) { s3ExitFullscreen(); return; }
  const req = st.requestFullscreen || st.webkitRequestFullscreen;
  if (!req) return;
  try {
    const p = req.call(st);
    if (p && p.catch) p.catch(() => {});
  } catch (e) {}
}

/* The control's accessible name and icon are derived from the DOCUMENT's
   state, never from what the last click intended — that is what keeps
   them truthful after an Escape or a browser-chrome exit. */
function s3SyncFullscreenBtn() {
  const b = s3FsBtn();
  if (!b) return;
  const on = s3IsFullscreen();
  b.classList.toggle('is-on', on);
  b.setAttribute('aria-pressed', on ? 'true' : 'false');
  b.setAttribute('aria-label', on ? 'إغلاق الشاشة الكاملة' : 'شاشة كاملة');
}

function s3OnFullscreenChange() {
  const fe = s3FsElement();
  /* Belt and braces for `nofullscreen`: if any UA route still managed to
     put the bare <video> into fullscreen, bounce straight back out
     rather than let the clip be shown without its explanation. */
  if (fe && fe === s3Video()) { s3ExitFullscreen(); return; }
  s3SyncFullscreenBtn();
  if (!fe) {
    const navExit = s3fsNavExit;
    s3fsNavExit = false;
    /* Predictable focus on exit: back to the control that owns the
       state — but only when the learner is still on this screen. */
    if (!navExit && currentScreen === 3) {
      const b = s3FsBtn();
      if (b && !b.classList.contains('hidden')) b.focus();
    }
  }
}
/* Registered once, at load — never per entry to the screen, which is
   how duplicate handlers (and duplicate reactions) get created. */
document.addEventListener('fullscreenchange', s3OnFullscreenChange);
document.addEventListener('webkitfullscreenchange', s3OnFullscreenChange);

/* Called by goTo() on every navigation: leaving the screen must not
   leave the learner staring at a fullscreen overlay of a screen they
   are no longer on. Pausing is goTo()'s own sweep. */
function s3ExitFullscreenOnNavigate() {
  if (!s3IsFullscreen()) return;
  s3fsNavExit = true;
  s3ExitFullscreen();
}

function s3VideoPlay() {
  const v = s3Video();
  if (!v) return;
  /* Sound is on, so this must run from the learner's gesture. The catch
     keeps a rejected play() from throwing into the console if a browser
     still refuses. */
  v.play().catch(() => {});
}

/* Bound to the element's own `play`, so it fires whether the learner
   used the overlay or clicked the native control bar. */
function s3VideoOnPlay() {
  const b = s3PlayBtn();
  if (b) b.classList.add('hidden');
}

function s3VideoOnEnded() {
  s3vidDone = true;
  const c = s3Continue();
  if (c) c.disabled = false;
  /* Resume: the watch gate is state (guide §8.8 rule 11) — save it the moment it is earned,
     not only on the next screen change (40-resume.js; a no-op while resume is off). */
  flushResumeSave();
}

/* Unwatched state: Continue held, play affordance offered, playback
   rewound. goTo() has already paused the element on the way out, so
   this never fights an in-flight play. */
function s3VideoResetInitial() {
  const v = s3Video(), b = s3PlayBtn(), c = s3Continue();
  if (v) { try { v.pause(); v.currentTime = 0; } catch (e) {} }
  if (b) b.classList.remove('hidden');
  if (c) c.disabled = true;
}

/* Already watched: Continue live immediately, no play affordance, and
   the clip left rewound and fully replayable through its own controls.
   Re-watching is optional — the gate is never charged twice, the same
   principle the producer set for the scroll gate. */
function s3VideoRestoreResolved() {
  const v = s3Video(), b = s3PlayBtn(), c = s3Continue();
  if (v) { try { v.pause(); v.currentTime = 0; } catch (e) {} }
  if (b) b.classList.add('hidden');
  if (c) c.disabled = false;
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
  // Feedback popup system). None exist in this part yet, but the sweep
  // is part of the shared engine shape every part carries.
  document.querySelectorAll('[id$="-popup"], [id$="-hint-overlay"]')
    .forEach(el => el.classList.add('hidden'));
  /* #recall-overlay does not match the sweep above; the RecallPopup
     contract requires its own close function here, never a bare toggle. */
  closeRecallPopup();
  /* Part 04's copy also calls closeAppletZoomOnNavigate() and
     closeImageZoomOnNavigate() here. Those two modals are screen-level
     components that no screen in this part uses yet, and their helpers
     do not exist in this file — carrying the CALLS without the
     functions would throw on the first navigation. They are omitted
     deliberately, and must be restored together with the modal markup,
     CSS and helpers if a later screen in this part adds an image or
     applet zoom (720-templates → _common-pitfalls.md → incomplete
     dependency closure). */
  /* A media wrapper must never survive the screen it belongs to. This
     runs BEFORE the pause sweep so the exit is already in flight while
     the clip is being stopped — the learner never hears audio from a
     screen they have left, fullscreen or not. */
  s3ExitFullscreenOnNavigate();
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
    /* Stateless screen — no answer, attempt or feedback to restore, so
       re-mounting the photo is the whole contract (720-templates →
       FullImageWithTextOverlay.md → "D"). No PartNav.mountBack() here:
       S1 has a real in-part predecessor (S0), so its Back is always
       available and is never trail-gated. */
    s1MountImage();
  }
  if (n === 8) {
    /* Stateless closing screen — nothing to restore beyond the photo. */
    s8MountImage();
  }
  if (n === 7) {
    /* Resume-state lock: a resolved sub-part never hands out a fresh
       attempt on return, and a resolved screen never comes back without
       its feedback. The photo is re-mounted on every entry, like every
       other image screen here. */
    s7MountImage();
    if (!s7viqDone) { if (screenPristine(7)) s7viqResetInitial(); }
    else s7viqRestoreResolved();
    syncPeakProgressNav(document.getElementById('s7'));
  }
  if (n === 6) {
    /* Resume-state lock: a resolved sub-part must never hand out a
       fresh attempt on return, and a resolved screen never comes back
       without its feedback. */
    if (!s6mcDone) { if (screenPristine(6)) s6mcResetInitial(); }
    else s6mcRestoreResolved();
    syncPeakProgressNav(document.getElementById('s6'));
  }
  if (n === 5) {
    /* Stateless — nothing to restore beyond the photo. */
    s5MountImage();
  }
  if (n === 4) {
    if (!s4scDone) { if (screenPristine(4)) s4scResetInitial(); }
    else s4scRestoreResolved();
    syncPeakProgressNav(document.getElementById("s4"));
  }
  if (n === 3) {
    /* The video screen. A screen already watched to the end stays
       resolved on return (Continue live, no play affordance); one that
       was never finished comes back to its unplayed state. */
    s3MountVideo();
    if (s3vidDone) s3VideoRestoreResolved();
    else           s3VideoResetInitial();
  }
  if (n === 2) {
    /* Resume-state lock: a resolved sub-part must never hand out a fresh
       attempt on return, and a resolved screen never comes back without
       its feedback (_global-components.md → "Reopening feedback on
       return to a resolved screen"). s2viqDone is the only flag
       involved — there is deliberately no second "popup was open on
       exit" flag. */
    s2MountImage();
    if (!s2viqDone) { if (screenPristine(2)) s2viqResetInitial(); }
    else s2viqRestoreResolved();
    syncPeakProgressNav(document.getElementById('s2'));
  }
}

/* Advance. BOTH screens built so far are ungated non-question screens —
   S0 a TransitionScreen, S1 a FullImageWithTextOverlay — so there is
   deliberately no condition here for either. S0 in particular does NOT
   wait for the character clip to finish playing. With TOTAL_SCREENS
   still 2 the goTo() off S1 is a bounds-checked no-op; see the
   TOTAL_SCREENS note for why that is the honest state of this part
   rather than a working end-of-part boundary. */
function advanceScreen() {
  /* S2 is a question screen and DOES gate: it must be resolved AND its
     feedback dismissed before the learner can move on. Both input paths
     resolve the same way — the engine binds ArrowLeft straight to
     advanceScreen(), so a gate living only in `button.disabled` would
     let a key press skip feedback the learner has not read, which is
     exactly the silent bypass _global-components.md → "Advance gates"
     exists to prevent. */
  if (currentScreen === 2 && !s2viqDone) return;
  /* S3's watch gate. Separate from the player's controls by
     figma-lomda-builder → "Video gating … is independent of controls":
     the learner may scrub, pause and replay freely, but the screen does
     not release until the element has actually fired `ended`. Guarding
     advanceScreen() — not just the button's `disabled` — is what stops
     the engine's ArrowLeft binding from skipping the video. */
  if (currentScreen === 3 && !s3vidDone) return;
  if (currentScreen === 4 && (!s4scDone || !s4scGateOk)) return;
  if (currentScreen === 6 && (!s6mcDone || !s6mcGateOk)) return;
  if (currentScreen === 7 && (!s7viqDone || !s7viqGateOk)) return;
  /* S8 is the last screen of the part: advancing off it leaves this
     part entirely rather than moving to screen 9. */
  if (currentScreen === 8) { crossToNextPart(); return; }
  goTo(currentScreen + 1);
}

/* Back returns to the previous valid predecessor in the learner's
   actual route. From this part's FIRST screen, Back crosses the part
   boundary backwards via PartNav instead of no-opping — with no
   recorded crossing (this part opened directly, e.g. in dev) that is
   itself a no-op, which is why the button stays hidden until
   PartNav.mountBack() confirms a real trail entry. Direct entry
   therefore cannot invent a previous destination. */
function goBack() {
  if (currentScreen === 0) { PartNav.back(); return; }
  goTo(currentScreen - 1);
}

document.addEventListener('keydown', e => {
  if (hasOwnArrowKeyMeaning(e.target)) return;
  /* Inside the fullscreen media wrapper the arrow keys belong to the
     player (seek/volume), not to screen navigation — and navigating out
     from underneath a fullscreen overlay is never what the learner
     asked for. Scoped to the fullscreen state only, so the approved
     normal-view behaviour of every screen is untouched. */
  if (s3IsFullscreen()) return;
  if (e.key === 'ArrowLeft')  advanceScreen();
  if (e.key === 'ArrowRight') goBack();
});


/* Cross-part navigation. init() must run AFTER this part's own initial
   screen setup: if this page was entered via Back, it jumps straight to
   the screen the learner left from. */
PartNav.init({
  part: '05',
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
/* 2026-09-24 fix: this listener was missing — a plain click/release
   left popupDragId set (pointerdown starts the drag but nothing ever
   cleared it on release), so the popup stayed glued to the cursor
   after the FIRST click, per the approved pattern already in
   03-01/02/03/04. */
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


   // px — see below


/* ─── Scroll gate: FIRST PRESENTATION ONLY ──────────────────
   Producer ruling, 2026-09-14:

     "the scroll gate must apply only when the learner receives that
      feedback outcome for the first time … switching from the revealed
      solution to My Answer and back must not re-arm the gate;
      navigating away and returning to the resolved screen must reopen
      the feedback with שנמשיך? enabled immediately; the feedback may
      reopen at the top, but rereading is optional."

   So the gate is a ONE-TIME read requirement per screen, not a toll
   charged on every render of the same feedback. The FIRST presentation
   is unchanged — it still holds the progression button until the
   learner reaches the end.

   This registry is the single owner of "has this screen's gate already
   been satisfied". It is deliberately shared rather than copied into
   each screen's own flags: S4, S6 and S7 run three byte-identical
   copies of the gate mechanics, and three copies of a second flag would
   be three chances to diverge. Each screen keeps its own `sNGateOk` —
   that is the LIVE "is the button released right now" state, which is
   a different question from "has this ever been read".

   Keyed per SCREEN, not per outcome, because a screen reaches exactly
   one terminal outcome per attempt cycle: `correct` and `wrong2` are
   mutually exclusive, and `retry` is never gated. A reset clears the
   key, so a screen returned to its unanswered state gates again.

   What deliberately does NOT change: `is-scroll` still tracks
   `cfg.scrollGate`, so the popup keeps its scrollable layout, its
   scrollbar and its typography on every presentation — rereading stays
   possible, it just stops being compulsory. */
const scrollGateSatisfied = Object.create(null);
function scrollGateNeeded(key, cfg) {
  return !!cfg.scrollGate && !scrollGateSatisfied[key];
}
function markScrollGateSatisfied(key) { scrollGateSatisfied[key] = true; }
function resetScrollGate(key)         { delete scrollGateSatisfied[key]; }


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

/* Generic Progress-Question renderer — ported VERBATIM from
   methodica-ar-math-ratio-03-04. It takes the question SET as a
   parameter and works for any count, so nothing here is specific to
   this part's four items (720-templates → _global-components.md →
   "Progress Question"). */
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

/* ═══════════════════════════════════════════════════════════
   PEAK QUESTION MODEL — ONE peak question, FOUR scored sub-parts
   720-templates → _global-components.md → "What one progress item
   represents" (Peak column, display) + _question-template-defaults.md →
   "Pass threshold formula (bundle scoring)" (scoring).

   ── AUTHORITY ──
   Producer ruling, 2026-09-14, and the deck says the same thing in two
   independent places:
     · slide 64's launch copy — "יש בה 4 סעיפים. עליכם להצליח ב - 3 לפחות."
     · slide 66's presenter note — "שאלה 1 מתוך 4"
   The Skill's own formula agrees independently: N = 4 sub-parts,
   threshold = ceil(2N/3) = ceil(8/3) = 3. Nothing here is invented.

   ── WHAT A DOT IS, IN THE PEAK STAGE ──
   This is the PEAK/climax stage, not basic/standard practice, and the
   two rows of that table are deliberately different:
     · one progress item = one LETTERED SUB-PART of the single peak
       question (not one question in a series);
     · an item resolves to ✓/✕ as soon as THAT sub-part is answered —
       "Peak — resolve immediately. A completed sub-part shows its
       verdict at once; do not wait for the rest of the peak question."
   That is the opposite of the never-resolve-early rule the sibling
   03-04 follows for its basic/standard series, and the difference is
   intentional. If a sub-part ever spans several screens, those screens
   are still ONE item.

   ── THE AGGREGATE ──
   The peak question's own result is SEPARATE from the four dots. It
   stays UNRESOLVED until all four sub-parts have resolved, and only
   then becomes:
       3 or 4 correct → peak question CORRECT
       0, 1 or 2      → peak question INCORRECT
   The four sub-part results are kept for diagnostics; the peak question
   contributes exactly ONE scored result. `sub-parts correct` is never
   itself a score.

   ── ROSTER — CONFIRMED ONLY FOR SUB-PART 1 ──
   Slide 66 is confirmed sub-part 1 by its own note ("שאלה 1 מתוך 4").
   Slides 68, 70 and 71 are still a WORKING HYPOTHESIS: they are the
   only other slides in this part carrying the progress-bar mock, but
   each must be confirmed from its own freshly extracted note before it
   is built — in particular whether 70 and 71 are two sub-parts or two
   sections of one. Entries below exist for all four so the bar renders
   at its authored length; only sub-part 1 has a resolver wired, and the
   others stay permanently unresolved until their screens are built.

   NOTHING IS STORED: every sub-part's state is derived on each sync
   from its screen's own sNDone / sNResult flags, so a second copy
   cannot drift. `resolved` and `scoredCorrect` stay separate fields —
   reaching a terminal state is not the same as being scored correct,
   and the final-wrong reveal never converts failure into success.

   NO REPORTING. The peak result is computed and published in-memory
   only; emitting it is Vadim's integration layer.
   ═══════════════════════════════════════════════════════════ */
const PEAK_THRESHOLD = 3;   /* = ceil(2*4/3); also the authored "3 לפחות" */

const PEAK_SUBPARTS = [
  { number: 1, slide: '66', screens: [2],
    built: true,
    resolved: () => (s2viqDone ? (s2viqResult === 'correct' ? 'correct' : 'incorrect') : null) },
  { number: 2, slide: '68', screens: [4],
    built: true,
    resolved: () => (s4scDone ? (s4scResult === 'correct' ? 'correct' : 'incorrect') : null) },
  /* AUTHORED, not inferred: slide 70's own presenter note (via
     slide70.xml.rels → notesSlide64.xml, cached slidenum "70") reads
     "שאלה 3 מתוך 4". That settles 70 as a sub-item of its own; it says
     nothing about 71, which is extracted on its own terms. */
  { number: 3, slide: '70', screens: [6],
    built: true,
    resolved: () => (s6mcDone ? (s6mcResult === 'correct' ? 'correct' : 'incorrect') : null) },
  /* AUTHORED, not inferred: slide 71's own presenter note (via
     slide71.xml.rels → notesSlide65.xml, cached slidenum "71") reads
     "שאלה 4 מתוך 4", and states the key outright (כוסמין 90 / חיטה 210).
     With slide 70's own note reading "שאלה 3 מתוך 4", every sub-part
     number in this model now comes from its own slide, and the peak
     AGGREGATE can finally resolve. */
  { number: 4, slide: '71', screens: [7],
    built: true,
    resolved: () => (s7viqDone ? (s7viqResult === 'correct' ? 'correct' : 'incorrect') : null) }
];

/* One entry per sub-part, in the shared component's own field shape. */
const peakProgress = {
  questions: PEAK_SUBPARTS.map(q => ({
    number: q.number, visited: false, state: 'not-answered', screens: q.screens.slice()
  })),
  results: []
};

/* A sub-part's DISPLAY state. Peak semantics: resolved → its verdict
   immediately; otherwise `current` while the learner is on one of its
   screens; otherwise not yet begun. */
function derivePeakState(q) {
  const verdict = q.resolved();
  if (verdict) return verdict;
  return q.screens.indexOf(currentScreen) !== -1 ? 'current' : 'not-answered';
}

/* The peak question as ONE scored unit. Unresolved until every sub-part
   has resolved; then correct iff at least PEAK_THRESHOLD are correct. */
function peakQuestionResult() {
  const r = PEAK_SUBPARTS.map(q => q.resolved());
  const correct = r.filter(x => x === 'correct').length;
  const allIn = r.every(x => x !== null);
  return {
    id: 'ratio-03-05-peak',
    subPartCount: PEAK_SUBPARTS.length,
    threshold: PEAK_THRESHOLD,
    subPartResults: r,                    /* diagnostics, never a score */
    subPartsResolved: r.filter(x => x !== null).length,
    subPartsCorrect: correct,             /* diagnostic, NOT a score */
    resolved: allIn,
    scoredCorrect: allIn ? correct >= PEAK_THRESHOLD : false
  };
}

/* Re-derives every sub-part from its screen's own result and renders the
   bar if the current screen carries one. Called on entry to each
   question screen and from each sub-part's finish handler, so returning
   to an earlier sub-part after resolving a later one shows both
   correctly. */
function syncPeakProgressNav(sectionEl) {
  peakProgress.questions.forEach((entry, i) => {
    const q = PEAK_SUBPARTS[i];
    entry.state = derivePeakState(q);
    if (q.screens.indexOf(currentScreen) !== -1) entry.visited = true;
    if (q.resolved()) entry.visited = true;
  });
  peakProgress.results = [peakQuestionResult()];
  const nav = sectionEl && sectionEl.querySelector('.progress-question');
  if (nav) updateProgressQuestion(nav, peakProgress);
}

/* ═══════════════════════════════════════════════════════════
   S2 — ValueInputQuestion, Variant C (two independent fields)
   Peak sub-part 1. Source: slide 66.

   ANSWER KEY — authored in slide 66's own presenter note, and matched
   by the slide's pre-filled indicator boxes and by its feedback body:
       spelt (קמח כוסמין), physically LEFT  → 3x
       wheat (קמח חיטה),  physically RIGHT → 7x
   The note also states explicitly that a capital X counts as correct.

   ACCEPTANCE — normalised, not pattern-matched loosely: strip all
   whitespace, lower-case, then compare to the key. That accepts "3x",
   "3X", " 3x " and rejects anything else, including "x3" (the factor
   must precede the variable, which is what the question asks for) and
   bare "3". A `*` between number and letter is also accepted, since it
   is the same expression typed the way a keyboard-first learner often
   writes it; nothing else is.

   GRANULARITY (ValueInputQuestion.md → "Scoring vs feedback
   granularity", Case A): the sub-part is ONE scored unit — correct only
   if BOTH fields are correct — but EVERY FIELD SHOWS ITS OWN VERDICT,
   so 1-right-1-wrong renders one green field and one red field while
   the sub-part scores incorrect. "Not scored per field" is not "not
   marked per field".

   ATTEMPTS — the canonical 2-attempt default. Attempt 1 wrong → red
   retry popup, fields marked per their own result, inputs stay editable.
   Attempt 2 wrong → final: the correct values are written into the
   fields and everything locks.

   FINAL-WRONG REVEAL — slide 66 authors NO View My Answer Toggle
   control, so this instance does not ship one. The reveal therefore
   happens on popup dismissal rather than immediately, which keeps the
   learner's own answer readable next to the verdict while the popup is
   open and still satisfies the template's "reveal the correct value and
   lock" requirement. Flagged to the producer.
   ═══════════════════════════════════════════════════════════ */
const S2_IMAGE = {
  src: 'assets/img/קמח.jpg',
  /* Alt text is the producer's own, from slide 66's note
     ("אלט טקסט לתמונה: מימין: קמח חיטה  משמאל: קמח כוסמין"), written out
     as a sentence rather than as the note's two label fragments. */
  alt: 'كيسا طحين بجانب بعضهما: على اليمين طحين قمح، وعلى اليسار طحين حنطة'
};

const S2_KEY = { spelt: '3x', wheat: '7x' };
const S2_MAX_ATTEMPTS = 2;

let s2viqAttempts = 0;
let s2viqDone     = false;
let s2viqResult   = null;   /* 'correct' | 'incorrect' */
/* Retry gate (_question-template-defaults.md → "Retry gate"), the same
   signature-snapshot shape as methodica-ar-math-ratio-03-01 S22: the
   normalised pair just marked wrong. Check stays disabled while the
   pair is unchanged and re-enables once either field differs. */
let s2viqLastWrongSig = null;
function s2viqSignature() {
  return ['spelt', 'wheat'].map(k => s2viqNorm(s2viqInput(k).value)).join('|');
}
/* There is deliberately NO gate flag here. Progression is gated on
   completion alone (see advanceScreen). The screen previously carried
   an `s2viqGateOk` that only opened when the learner dismissed the
   feedback by clicking outside it or pressing Escape — both of which
   _global-components.md → "No manual close" forbids outright, and
   neither of which is discoverable. The effect was that the visibly
   enabled "متابعة" did nothing on its first press. Removed 2026-09-14 on
   the producer's ruling; see docs/pending-assets.md item 10. */

/* ─── View My Answer Toggle state ───────────────────────────
   720-templates → _global-components.md → Feedback popup system →
   "View My Answer Toggle". This screen MUST carry it: all three
   applicability conditions hold — the final incorrect state reveals the
   solution, the reveal writes into the SAME input fields that held the
   learner's answer, and after it the learner can no longer see what they
   submitted. Slide 66 authoring no such control does not override a
   global interaction rule.

   s2viqAnswer — a snapshot of the learner's FINAL submitted RAW strings,
   taken in the final-wrong branch BEFORE the reveal overwrites them.
   It stores exactly what was typed ("3X", " 7x ", "3*x"), never a value
   reconstructed from the normalised comparison, so the answer view can
   restore it byte for byte. null = not toggleable (correct, or nothing
   submitted yet).
   s2viqView — 'solution' | 'answer'. Runtime only, never persisted, and
   reset to 'solution' on every fresh entry AND on every resolved
   re-entry: Back always lands on the solution view. */
let s2viqAnswer = null;
let s2viqView   = 'solution';

function s2viqInput(k) { return document.getElementById('s2viq-input-' + k); }
function s2viqField(k) { return document.getElementById('s2viq-field-' + k); }
function s2viqIco(k)   { return document.getElementById('s2viq-ico-' + k); }

/* Normalise a typed algebraic answer for comparison. Whitespace and
   case are irrelevant, and an explicit multiplication sign is the same
   expression. Everything else is left alone so the comparison stays
   strict. */
function s2viqNorm(v) {
  return String(v || '').replace(/\s+/g, '').replace(/\*/g, '').toLowerCase();
}
function s2viqFieldCorrect(k) {
  return s2viqNorm(s2viqInput(k).value) === S2_KEY[k];
}

/* Check enables only when BOTH fields carry something. */
/* Fires on a GENUINE value change only — typing, deleting, replacing,
   pasting, or any other real edit. The `input` event is what makes that
   distinction for free: focus, a click or Tab into the field that
   changes nothing, caret movement and text selection do not fire it,
   and neither does a programmatic `.value` write during screen entry or
   toggle restoration (those set the property without dispatching). That
   is why the dismissal hangs off this handler rather than off a pointer
   listener. */
function s2viqOnInput() {
  if (s2viqDone) return;
  /* Retry-feedback dismissal — approved path 2 of "No manual close":
     the feedback describes an attempt that is no longer current, so it
     goes away as a side effect of the learner getting on with the task.
     Never offered as a way to close the popup, and scoped to a
     non-final wrong attempt (s2viqAttempts > 0 && !s2viqDone). The same
     shape S4 and S6 already use in their own answer-change handlers.

     THE MARKS GO WITH IT. A verdict icon describes the value standing in
     its field, so the moment the learner changes a value the previous
     attempt's ✓/✗ stops describing anything on screen and must clear —
     both fields, not just the edited one, since the pair was judged
     together. Producer ruling 2026-09-14: "verdict styling is
     view-dependent and must always match the values currently visible,
     not the values from a previous view or submission state."
     Only a genuine value change reaches here; focus, a click or Tab that
     edits nothing, caret movement and text selection never fire `input`,
     so none of them clear the marks. */
  if (s2viqAttempts > 0) {
    document.getElementById('s2viq-popup').classList.add('hidden');
    s2viqClearMarks();
  }
  const filled = ['spelt', 'wheat'].every(k => s2viqInput(k).value.trim() !== '');
  document.getElementById('s2viq-check').disabled =
    !filled || s2viqSignature() === s2viqLastWrongSig;
}

const S2_ICO_CHECK = '<svg viewBox="0 0 12.7273 9.18763" width="16" height="16" fill="none" aria-hidden="true"><path d="M4.10942 9.18763C3.74814 9.18778 3.40167 9.04418 3.14642 8.78851L0.234956 5.87812C-0.0783188 5.56475 -0.0783188 5.05677 0.234956 4.74339C0.548332 4.43012 1.05631 4.43012 1.36969 4.74339L4.10942 7.48312L11.3576 0.234956C11.671 -0.0783188 12.1789 -0.0783188 12.4923 0.234956C12.8056 0.548332 12.8056 1.05631 12.4923 1.36969L5.07242 8.78851C4.81717 9.04418 4.47069 9.18778 4.10942 9.18763Z" fill="#609e12"/></svg>';
const S2_ICO_X = '<svg viewBox="0 0 10 10" width="14" height="14" fill="none" aria-hidden="true"><path d="M9.76736 0.232637C9.61836 0.0836796 9.4163 0 9.20561 0C8.99492 0 8.79286 0.0836796 8.64386 0.232637L5 3.87649L1.35615 0.232637C1.20714 0.0836796 1.00508 0 0.794391 0C0.583702 0 0.381639 0.0836796 0.232637 0.232637C0.0836796 0.381639 0 0.583702 0 0.794391C0 1.00508 0.0836796 1.20714 0.232637 1.35615L3.87649 5L0.232637 8.64386C0.0836796 8.79286 0 8.99492 0 9.20561C0 9.4163 0.0836796 9.61836 0.232637 9.76736C0.381639 9.91632 0.583702 10 0.794391 10C1.00508 10 1.20714 9.91632 1.35615 9.76736L5 6.12351L8.64386 9.76736C8.79286 9.91632 8.99492 10 9.20561 10C9.4163 10 9.61836 9.91632 9.76736 9.76736C9.91632 9.61836 10 9.4163 10 9.20561C10 8.99492 9.91632 8.79286 9.76736 8.64386L6.12351 5L9.76736 1.35615C9.91632 1.20714 10 1.00508 10 0.794391C10 0.583702 9.91632 0.381639 9.76736 0.232637Z" fill="#b20010"/></svg>';

/* Mark each field with ITS OWN verdict (Case A). */
function s2viqMarkFields() {
  ['spelt', 'wheat'].forEach(k => {
    const ok = s2viqFieldCorrect(k);
    const f = s2viqField(k);
    f.classList.toggle('viq-correct', ok);
    f.classList.toggle('viq-wrong', !ok);
    s2viqIco(k).innerHTML = ok ? S2_ICO_CHECK : S2_ICO_X;
  });
}
function s2viqClearMarks() {
  ['spelt', 'wheat'].forEach(k => {
    s2viqField(k).classList.remove('viq-correct', 'viq-wrong');
    s2viqIco(k).innerHTML = '';
  });
}

/* Authored feedback body, verbatim from slide 66. Both outcomes share the
   same pedagogical body (the deck authors them identically, id=15 and
   id=39); only the title differs. The authored blank paragraph is
   preserved as a real blank line, not as an invented margin. */
/* Client correction (Nitza, 2026-09-29): x is defined as a number of
   grams from the start of the task, so every later use of x as grams
   (S4 "להוסיף x גרם", S7 3x + 100 = 7x - 20) is well defined. Client QA
   wording, overriding the script. */
const S2VIQ_EXPLAIN = [
  'النسبة بين طحين الحنطة وطحين القمح في خليط العجين هي <bdi class="scq-expr" dir="ltr">3 : 7</bdi>.',
  '',
  'نوسّع النسبة بـ <bdi dir="ltr">x</bdi> مرات، بحيث إنّ  <bdi dir="ltr">x</bdi> هو عدد الغرامات، ونحصل على:',
  '<bdi dir="ltr">3x</bdi> غرام من طحين الحنطة،',
  'و-<bdi dir="ltr">7x</bdi> غرام من طحين القمح‎.'
];

/* Popup config, in the shared shape every question screen in this lomda
   uses. Backgrounds are this lomda's established pair (#edf8ed correct /
   #ffdbdc incorrect), already shipped by -02, -03 and -04; slide 66's own
   incorrect popup is filled #FFE5E5, the deck's approximation of the same
   role — the shared token wins, as it does for every other popup here.
   Titles and body are the slide's authored text, verbatim.

   `wrong2` is the FINAL-wrong outcome and the only one carrying
   `action`/`answerView`, so the View My Answer Toggle can never appear on
   the retry or correct paths (_global-components.md → View My Answer
   Toggle → "Only wrong2 ever carries action/answerView").

   titleSecondLine — canonical system copy owned by the template contract
   (_question-template-defaults.md → "Terminal-incorrect title carries a
   fixed second line"). It is APPENDED to the authored first line on the
   SOLUTION view only, and is explicitly cleared to null in `answerView`,
   which shows the learner's own answer rather than the solution. It has
   to be cleared rather than omitted, because the answer view is built by
   merging that object onto this base. */
const S2VIQ_POPUP = {
  /* CORRECTED 2026-09-22 (producer, on QA question: "למה המשוב המלא
     מופיע כבר אחרי ניסיון ראשון שגוי?"). This was the one screen in
     the whole lomda where `retry` (the non-final wrong attempt) carried
     the full worked-solution body and the wrong2 title instead of the
     canonical short nudge — every other screen across all six parts
     (30+ retry configs, checked) uses exactly
     `{ title: 'الإجابة غير صحيحة.', body: ['هل نحاول مرة أخرى؟'] }` here, and
     shows S2VIQ_EXPLAIN only on `correct` and on the FINAL wrong
     attempt (`wrong2`). The deck's own shape inventory for this slide
     authors only two feedback shapes (id=15 correct, id=39 incorrect),
     same as every other screen that correctly reserves the incorrect
     one for the final attempt — nothing here called for an exception.
     Fixed by dropping in the shared shape. */
  retry: { bg: '#ffdbdc', title: 'الإجابة غير صحيحة.', body: ['هل نحاول مرة أخرى؟'] },
  correct: {
    bg: '#edf8ed',
    title: 'إجابتكم صحيحة!',
    body: S2VIQ_EXPLAIN
  },
  wrong2: {
    bg: '#ffdbdc',
    title: 'هذا خطأ. لا تقلقوا، فالفهم يتم تدريجيًا‎.',
    titleSecondLine: 'الإجابة الصحيحة تَظهَر الآن.',
    body: S2VIQ_EXPLAIN,
    action: { label: 'الانتقال إلى إجابتي', view: 'answer' },
    answerView: {
      titleSecondLine: null,
      title: 'الإجابة غير صحيحة.',
      /* Fixed canonical copy owned by the template contract, never
         producer content and never a candidate for override. */
      body: ['يمكن مشاهدة الحل الصحيح مرة أخرى‎.'],
      action: { label: 'الإجابة الصحيحة', view: 'solution' }
    }
  }
};

function s2viqShowPopup(type, view) {
  const popup = document.getElementById('s2viq-popup');
  if (!popup) return;
  const base = S2VIQ_POPUP[type];
  const cfg  = (view === 'answer' && base.answerView)
    ? Object.assign({}, base, base.answerView)
    : base;

  popup.style.background = cfg.bg;
  /* Re-anchor in TWO cases, not one: on a fresh open, AND on every
     `wrong2` render including each toggle between its views — a popup
     that has been dragged converts to a fixed pixel `top`, and if that
     survives a view swap that changes its height the box shrinks from
     its pinned top edge and leaves a gap above the bottom bar
     (_global-components.md → View My Answer Toggle, approved product
     decision 2026-08-11). */
  if (popup.classList.contains('hidden') || type === 'wrong2') resetPopupPosition(popup);

  const titleEl = document.getElementById('s2viq-popup-title');
  titleEl.innerHTML = cfg.titleSecondLine
    ? (cfg.title + '<br>' + cfg.titleSecondLine)
    : cfg.title;

  document.getElementById('s2viq-popup-body').innerHTML = cfg.body.map(p =>
    p === '' ? '<p class="scq-gap" aria-hidden="true"></p>' : '<p>' + p + '</p>'
  ).join('');

  /* Shared helper — renders the button, or removes the row entirely when
     `cfg.action` is absent (retry / correct). */
  renderPopupAction('s2viq-popup', cfg.action, s2viqToggleView);

  popup.classList.remove('hidden');
  clampPopup(popup);
}

/* ─── View My Answer Toggle ─────────────────────────────────
   Single owner of the resolved-wrong board+popup state. Writes the
   displayed values for the current view into the live fields, re-marks
   each field from what is now displayed, re-locks (a re-render can
   recreate unlocked elements) and re-renders the popup in that view.

   Field marking follows the displayed values, which is what makes both
   views correct for free: the solution view shows 3x/7x so every field
   marks ✓, and the answer view shows the learner's own raw strings so
   each field marks by its own result — the mixed board the learner
   actually produced. */
function s2viqApplyView() {
  if (s2viqAnswer === null) return;
  ['spelt', 'wheat'].forEach(k => {
    s2viqInput(k).value = (s2viqView === 'answer') ? s2viqAnswer[k] : S2_KEY[k];
  });
  s2viqMarkFields();
  ['spelt', 'wheat'].forEach(k => { s2viqInput(k).disabled = true; });
  s2viqShowPopup('wrong2', s2viqView);
}

/* A PURE view switch. It never routes through s2viqCheck(), and it never
   touches attempts, done/complete state, the sub-part result, the
   progress dot, navigation state or any reporting call — the contract
   forbids all of those. The one thing it does mutate is the popup's
   position, which is its own explicit rule (see s2viqShowPopup). */
function s2viqToggleView() {
  if (s2viqAnswer === null) return;
  s2viqView = (s2viqView === 'solution') ? 'answer' : 'solution';
  s2viqApplyView();
}

function s2viqCheck() {
  if (s2viqDone) { advanceScreen(); return; }
  s2viqAttempts++;
  const allOk = ['spelt', 'wheat'].every(k => s2viqFieldCorrect(k));
  reportQ(['001', 'q1'], allOk, allOk || s2viqAttempts >= S2_MAX_ATTEMPTS, 'طحين الحنطة: ' + String(s2viqInput('spelt').value).trim() + ', طحين القمح: ' + String(s2viqInput('wheat').value).trim());

  if (allOk) {
    s2viqResult = 'correct';
    s2viqAnswer = null;            /* nothing to toggle to */
    s2viqMarkFields();
    s2viqFinish();
    s2viqShowPopup('correct');
    return;
  }

  if (s2viqAttempts < S2_MAX_ATTEMPTS) {
    /* Non-final wrong: fields carry their own marks, inputs stay
       editable, hint stays available, nothing is resolved, and NO
       toggle appears — `retry` carries no action. */
    s2viqMarkFields();
    s2viqLastWrongSig = s2viqSignature();
    document.getElementById('s2viq-check').disabled = true;
    S2VIQ_POPUP.retry.title = ['spelt', 'wheat'].some(k => s2viqFieldCorrect(k)) ? RETRY_TITLE_PARTIAL : RETRY_TITLE;
    s2viqShowPopup('retry');
    return;
  }

  /* FINAL wrong. Snapshot the learner's RAW submitted strings BEFORE the
     reveal overwrites the same fields — this is the only moment they
     still exist. Stored exactly as typed, so "3X" / " 7x " / "3*x" come
     back byte for byte rather than being rebuilt from the normalised
     comparison. */
  s2viqAnswer = { spelt: s2viqInput('spelt').value, wheat: s2viqInput('wheat').value };
  s2viqResult = 'incorrect';
  s2viqView   = 'solution';        /* the solution is shown IMMEDIATELY */
  s2viqFinish();
  s2viqApplyView();
}

/* Terminal state: lock the inputs, relabel Check as the progression
   action, hide the hint, and resolve this sub-part's dot IMMEDIATELY
   (peak semantics). The toggle never changes this button's label. */
function s2viqFinish() {
  s2viqDone = true;
  ['spelt', 'wheat'].forEach(k => { s2viqInput(k).disabled = true; });
  const check = document.getElementById('s2viq-check');
  check.disabled = false;
  setNavLabel(check, 'متابعة');
  hideHintButton('s2viq-hint');
  syncPeakProgressNav(document.getElementById('s2'));
}

/* Fresh, unanswered state — used on first entry and on re-entry to a
   screen the learner has not resolved. */
function s2viqResetInitial() {
  s2viqAttempts = 0;
  s2viqLastWrongSig = null;
  s2viqDone = false;
  s2viqResult = null;
  s2viqAnswer = null;
  s2viqView = 'solution';
  ['spelt', 'wheat'].forEach(k => {
    const el = s2viqInput(k);
    el.value = '';
    el.disabled = false;
  });
  s2viqClearMarks();
  const check = document.getElementById('s2viq-check');
  setNavLabel(check, 'هل إجابتي صحيحة؟');
  check.disabled = true;
  showHintButton('s2viq-hint');
  renderPopupAction('s2viq-popup', null);
  document.getElementById('s2viq-popup').classList.add('hidden');
  document.getElementById('s2viq-hint-overlay').classList.add('hidden');
}

/* Restore a RESOLVED screen on return: the resolved state must come back
   with its feedback, never as a fresh attempt (_global-components.md →
   "Reopening feedback on return to a resolved screen").
   NOTHING is re-armed here. The feedback reopens and "متابعة" is live
   immediately — there is no acknowledgement to repeat, and this screen
   has no scroll gate to re-charge (and must not be given one).
   Back ALWAYS lands on the SOLUTION view, never on whichever view the
   learner last had open — the toggle contract states that explicitly,
   and the learner's answer stays exactly one click away. */
function s2viqRestoreResolved() {
  const check = document.getElementById('s2viq-check');
  check.disabled = false;
  setNavLabel(check, 'متابعة');
  hideHintButton('s2viq-hint');
  if (s2viqResult === 'correct') {
    ['spelt', 'wheat'].forEach(k => { s2viqInput(k).disabled = true; });
    s2viqMarkFields();
    renderPopupAction('s2viq-popup', null);
    s2viqShowPopup('correct');
    return;
  }
  s2viqView = 'solution';
  s2viqApplyView();
}

function s2viqOpenHint()  { xapiRequestedHint('001', 'q1'); document.getElementById('s2viq-hint-overlay').classList.remove('hidden'); }
function s2viqCloseHint() { document.getElementById('s2viq-hint-overlay').classList.add('hidden'); }
function s2viqCloseHintOnBackdrop(e) { if (e.target === e.currentTarget) s2viqCloseHint(); }

/* RecallPopup — the "how to type x" keyboard guide (#recall-overlay).
   Base contract, _global-components.md → RecallPopup: one shared
   overlay, opens unconditionally, closes on its button / backdrop /
   Escape / every goTo(). Focus moves to the close control on open and
   back to the trigger on close, so keyboard users are never stranded. */
let recallReturnFocus = null;
function openRecallPopup() {
  const ov = document.getElementById('recall-overlay');
  if (!ov) return;
  recallReturnFocus = document.activeElement;
  ov.classList.remove('hidden');
  const close = ov.querySelector('.recall-popup-close');
  if (close) close.focus();
}
function closeRecallPopup() {
  const ov = document.getElementById('recall-overlay');
  if (!ov || ov.classList.contains('hidden')) return;
  ov.classList.add('hidden');
  if (recallReturnFocus && document.contains(recallReturnFocus)) recallReturnFocus.focus();
  recallReturnFocus = null;
}
function closeRecallPopupOnBackdrop(e) { if (e.target === e.currentTarget) closeRecallPopup(); }

/* Escape closes the HINT OVERLAY only.
   The feedback popup is NOT dismissible — not by this key, not by an
   outside click, not by any control. _global-components.md → "No manual
   close": "There is no close control, and no gesture, keystroke, or
   outside click dismisses it", and Escape-to-close and
   outside-click-to-close "must not exist … do not merely hide them".
   The outside-click listener that used to live here, and the popup half
   of this handler, were removed 2026-09-14 on the producer's ruling.
   The hint overlay is a different component: it has an authored close
   button ("חזרה לשאלה") and keeps all three of its dismissal paths. */
document.addEventListener('keydown', e => {
  if (e.key !== 'Escape') return;
  /* Whichever hint overlay is open (S2, S4, S6, S7), as in the sibling parts. */
  document.querySelectorAll('[id$="-hint-overlay"]').forEach(el => {
    if (!el.classList.contains('hidden')) el.classList.add('hidden');
  });
  closeRecallPopup();   // the keyboard guide keeps all three dismissal paths too
});

/* ═══════════════════════════════════════════════════════════
   S4 — SingleChoiceQuestion. Peak sub-part 2. Source: slide 68.

   State machine ported from the approved sibling
   methodica-ar-math-ratio-03-04's own S1 (the same template, the same
   scroll-gated long feedback), with this slide's own content.

   ANSWER KEY — o2 = "גלי", on three agreeing pieces of evidence:
   the slide's own indicator fill (three options #00B0F0, "גלי" alone
   #00B050 green), the feedback's own conclusion ("אך היחס לא נשמר" —
   the ratio is NOT preserved), and the stem's logic (Gali says it will
   not be preserved; the influencer says it will).

   OPTION ORDER is the slide's top-to-bottom GEOMETRY (y 249.4 / 324.4 /
   399.5 / 474.5), not XML document order, which lists them 11, 27, 32,
   37 and would have produced the wrong sequence.

   NO VIEW MY ANSWER TOGGLE — required, and correct: on a final wrong
   attempt this template marks the learner's pick red AND the correct
   option green at the same time, so both stay visible and
   distinguishable. That fails the toggle's third applicability
   condition, and _global-components.md names SingleChoiceQuestion as a
   standing exclusion.

   SCROLL GATE — authored by the slide's own note
   ("שימו לב שבמשוב ארוך, לאפשר גלילה."), applied to BOTH outcomes
   because both feedback boxes carry the same 645px-tall explanation.
   ═══════════════════════════════════════════════════════════ */
const S4SC = { correctId: 'o2', maxAttempts: 2 };

/* The shared pedagogical explanation. Both outcomes carry it verbatim
   in the deck; only the titles differ. '' is an authored blank
   paragraph and renders as one line of space, never an invented margin.

   The correct box additionally opens with a blank paragraph between its
   title and its body; that is PowerPoint spacing a title from a body
   inside ONE textbox. Here title and body are separate elements with
   the popup's own 4px gap, so that leading blank is deliberately not
   reproduced — it would open the body with a spurious empty line. */
const S4SC_EXPLAIN = [
  'كمية طحين الحنطة هي <bdi dir="ltr">3x</bdi> غرام،',
  'وكمية طحين القمح هي <bdi dir="ltr">7x</bdi> غرام.',
  '',
  'إذا أضفنا<bdi dir="ltr">x</bdi> غرام إلى كمية طحين الحنطة،',
  'نحصل على <bdi dir="ltr">4x</bdi> غرام.',
  '',
  'إذا طرحنا<bdi dir="ltr">x</bdi> غرام من طحين القمح،',
  'نحصل على <bdi dir="ltr">6x</bdi> غرام.',
  '',
  'النسبة الجديدة <bdi class="scq-expr" dir="ltr">4x : 6x</bdi> .',
  'بعد الاختزال، النسبة هي <bdi class="scq-expr" dir="ltr">2 : 3</bdi> .',
  'هذه النسبة مختلفة عن النسبة الأصلية (<bdi class="scq-expr" dir="ltr">3 : 7</bdi>).',
  '',
  'أي إن كمية الطحين الإجمالية بقيت <bdi dir="ltr">10x</bdi>،',
  'لكن النسبة لم تُحفظ‎.'
];

/* Titles are the slide's own authored text. Backgrounds are this
   lomda's established pair. `retry` carries the canonical generic
   copy — slide 68 authors no separate first-wrong feedback, only the
   two outcome boxes. */
const S4SC_POPUP = {
  retry:   { bg: '#ffdbdc', title: 'الإجابة غير صحيحة.', body: ['هل نحاول مرة أخرى؟'] },
  correct: { bg: '#edf8ed', title: 'صحيح‎!', body: S4SC_EXPLAIN, scrollGate: true },
  wrong2:  {
    bg: '#ffdbdc',
    title: 'هذا خطأ، لكن مجرد التعامل مع المسألة يطوّر التفكير‎!',
    /* The canonical fixed second line every terminal-incorrect SOLUTION
       view carries (_question-template-defaults.md → "Terminal-incorrect
       title carries a fixed second line"). Appended to, never replacing,
       the authored first line. Added 2026-09-14 on the producer's
       instruction, after the omission was found while building S6 — this
       screen's terminal-wrong state does reveal the correct answer
       (learner's pick red AND the key green, simultaneously), so the
       rule applies here exactly as it does on S2 and S6. It is the ONLY
       change to S4: answer key, body, attempts, markings, scroll gate,
       scoring and restoration are all untouched.
       `retry` deliberately does NOT get it — nothing is revealed yet. */
    titleSecondLine: 'الإجابة الصحيحة تَظهَر الآن.',
    body: S4SC_EXPLAIN,
    scrollGate: true
  }
};

let s4scSelected  = null;
let s4scAttempts  = 0;
let s4scDone      = false;
let s4scResult    = null;   /* 'correct' | 'wrong2' */
let s4scLastWrong = null;
let s4scGateOk    = true;   /* scroll gate satisfied */

function s4scOpt(id)  { return document.querySelector('#s4 .scq-opt[data-id="' + id + '"]'); }
function s4scOpts()   { return document.querySelectorAll('#s4 .scq-opt'); }

function s4scSelect(id) {
  if (s4scDone) return;
  s4scOpts().forEach(o => {
    o.classList.remove('selected');
    o.setAttribute('aria-checked', 'false');
  });
  const opt = s4scOpt(id);
  opt.classList.add('selected');
  opt.setAttribute('aria-checked', 'true');
  s4scSelected = id;

  /* Choosing again after a wrong attempt clears the previous verdict
     and closes its popup; re-picking the SAME wrong option leaves Check
     disabled, so a second attempt has to be a genuinely different
     answer. */
  if (s4scAttempts > 0) {
    document.getElementById('s4sc-popup').classList.add('hidden');
    s4scOpts().forEach(o => o.classList.remove('correct', 'wrong'));
  }
  const chk = document.getElementById('s4sc-check');
  if (chk) chk.disabled = (id === s4scLastWrong);
}

function s4scCheck() {
  if (s4scDone) { advanceScreen(); return; }
  if (!s4scSelected) return;

  s4scAttempts++;
  const correct = s4scSelected === S4SC.correctId;
  reportQ(['001', 'q2'], correct, correct || s4scAttempts >= S4SC.maxAttempts, xapiChoiceAnswer('001', 'q2', '#s4 .scq-opt', s4scSelected));

  if (correct) {
    s4scOpt(s4scSelected).classList.add('correct');
    s4scResult = 'correct';
    s4scFinish();
    s4scShowPopup('correct');
    return;
  }

  if (s4scAttempts < S4SC.maxAttempts) {
    /* Non-final wrong: mark the chosen option wrong; the correct one is
       NOT revealed yet. */
    s4scOpt(s4scSelected).classList.add('wrong');
    s4scLastWrong = s4scSelected;
    document.getElementById('s4sc-check').disabled = true;
    s4scShowPopup('retry');
    return;
  }

  /* Final wrong: chosen option red, correct option green, both shown at
     once — which is exactly why this template has no toggle. */
  s4scOpt(s4scSelected).classList.add('wrong');
  s4scOpt(S4SC.correctId).classList.add('correct');
  s4scResult = 'wrong2';
  s4scFinish();
  s4scShowPopup('wrong2');
}

function s4scFinish() {
  s4scDone = true;
  s4scOpts().forEach(o => { o.disabled = true; });
  const chk = document.getElementById('s4sc-check');
  /* SingleChoiceQuestion's own completion label is "هل نتابع؟" — NOT
     ValueInputQuestion's "متابعة"; the two templates do not share one,
     and slide 68 authors no override. s4scShowPopup() may immediately
     re-disable this button while the scroll gate is unsatisfied; that
     is the gate's one and only target. */
  if (chk) { setNavLabel(chk, 'هل نتابع؟'); chk.disabled = false; }
  hideHintButton('s4sc-hint');
  /* Peak semantics: this sub-part's dot resolves IMMEDIATELY. */
  syncPeakProgressNav(document.getElementById('s4'));
}

function s4scResetInitial() {
  s4scSelected  = null;
  s4scAttempts  = 0;
  s4scDone      = false;
  s4scResult    = null;
  s4scLastWrong = null;
  s4scGateOk    = true;
  /* A screen returned to its unanswered state gates again on the next
     terminal outcome — the one-time allowance belongs to a resolution,
     not to the session. */
  resetScrollGate('s4sc');
  s4scOpts().forEach(o => {
    o.classList.remove('selected', 'correct', 'wrong');
    o.setAttribute('aria-checked', 'false');
    o.disabled = false;
  });
  const chk = document.getElementById('s4sc-check');
  setNavLabel(chk, 'هل إجابتي صحيحة؟');
  chk.disabled = true;
  showHintButton('s4sc-hint');
  document.getElementById('s4sc-popup').classList.add('hidden');
  document.getElementById('s4sc-hint-overlay').classList.add('hidden');
}

/* Restore a RESOLVED screen on return: the resolved state comes back
   with its feedback, never as a fresh attempt.
   The scroll gate is NOT re-armed if it has already been satisfied once
   (producer ruling 2026-09-14 — see scrollGateNeeded). The feedback
   reopens at the top and stays scrollable, but the progression button
   is live immediately; rereading is optional. A learner who has never
   reached the end still meets the gate here. */
function s4scRestoreResolved() {
  s4scOpts().forEach(o => { o.disabled = true; });
  const chk = document.getElementById('s4sc-check');
  setNavLabel(chk, 'هل نتابع؟');
  chk.disabled = false;
  hideHintButton('s4sc-hint');
  s4scShowPopup(s4scResult === 'correct' ? 'correct' : 'wrong2');
}

function s4scShowPopup(type) {
  const popup = document.getElementById('s4sc-popup');
  if (!popup) return;
  const cfg      = S4SC_POPUP[type];
  const body     = document.getElementById('s4sc-popup-body');
  const scroller = document.getElementById('s4sc-popup-scroll');
  const chk      = document.getElementById('s4sc-check');

  popup.style.background = cfg.bg;
  if (popup.classList.contains('hidden')) resetPopupPosition(popup);
  /* The terminal-incorrect title is two bold lines, so the element has
     to render MARKUP — a textContent write would show a literal <br> or
     silently drop the break. Every other outcome passes no
     titleSecondLine and so renders exactly the single line it did
     before; the authored strings themselves carry no markup. */
  const s4scTitleEl = document.getElementById('s4sc-popup-title');
  s4scTitleEl.innerHTML = cfg.titleSecondLine
    ? (cfg.title + '<br>' + cfg.titleSecondLine)
    : cfg.title;
  body.innerHTML = cfg.body.map(p =>
    p === '' ? '<p class="scq-gap" aria-hidden="true"></p>' : '<p>' + p + '</p>'
  ).join('');

  /* Scroll gate — arm before showing, then reset and measure AFTER the
     popup is visible. Both have to happen after the unhide: a
     display:none element ignores a scrollTop write and reports
     scrollHeight/clientHeight as 0.
     `is-scroll` still tracks cfg.scrollGate on EVERY presentation, so
     the popup keeps its scrollable layout and scrollbar; only whether
     the gate is CHARGED depends on whether it has been read before. */
  popup.classList.toggle('is-scroll', !!cfg.scrollGate);
  const gateNeeded = scrollGateNeeded('s4sc', cfg);
  s4scGateOk = !gateNeeded;
  if (chk && !s4scGateOk) chk.disabled = true;
  if (cfg.scrollGate) scroller.setAttribute('tabindex', '0');   /* the only way a keyboard user can satisfy the gate */
  else                scroller.removeAttribute('tabindex');

  popup.classList.remove('hidden');
  clampPopup(popup);
  scroller.scrollTop = 0;
  if (gateNeeded) {
    placeScrollHint('s4sc-popup');
    s4scCheckScrollGate();
  } else {
    /* Already read once, or never gated: reopen at the top with the
       progression button live. Rereading stays possible, not required. */
    popup.classList.remove('needs-scroll');
    if (chk && s4scDone) chk.disabled = false;
  }
}

function s4scCheckScrollGate() {
  if (s4scGateOk) return;
  const popup    = document.getElementById('s4sc-popup');
  const scroller = document.getElementById('s4sc-popup-scroll');
  if (!popup || popup.classList.contains('hidden')) return;
  if (!atBottom(scroller)) { popup.classList.add('needs-scroll'); return; }
  /* Reaching the end satisfies the gate for this showing AND records it
     so no later presentation of the same feedback charges it again. */
  s4scGateOk = true;
  markScrollGateSatisfied('s4sc');
  popup.classList.remove('needs-scroll');
  const chk = document.getElementById('s4sc-check');
  if (chk && s4scDone) chk.disabled = false;
}

function s4scOpenHint()  { xapiRequestedHint('001', 'q2'); document.getElementById('s4sc-hint-overlay').classList.remove('hidden'); }
function s4scCloseHint() { document.getElementById('s4sc-hint-overlay').classList.add('hidden'); }
function s4scCloseHintOnBackdrop(e) { if (e.target === e.currentTarget) s4scCloseHint(); }

/* ═══════════════════════════════════════════════════════════
   S6 — MultipleChoiceQuestion (text variant)
   Peak sub-part 3. Source: slide 70.

   Spec: 720-templates → MultipleChoiceQuestion.md + _question-template-
   defaults.md + _global-components.md → Feedback popup / Long-feedback
   scroll gate. The state machine is the approved sibling
   methodica-ar-math-ratio-03-04's own S6 — the same template at the same
   option count — with this part's already-shipped scroll-gate mechanics
   on its popup.

   SUB-PART NUMBER IS AUTHORED. Slide 70's own presenter note (resolved
   through slide70.xml.rels → notesSlide64.xml) reads "שאלה 3 מתוך 4",
   so this screen alone is peak sub-part 3. It says nothing about slide
   71, which is extracted on its own terms.

   CORRECT SET = {o2, o3} — three independent, agreeing signals, set out
   in full on this screen's markup in index.html: the pill strokes
   (#00B050 ×2 against the stock #00B0F0 ×2), the marker squares (filled
   ×2, empty outline ×2 on the deck's own render), and the mathematics
   the feedback spells out. o1 and o4 are the distractors — o4 is the
   classic additive-preserves-ratio error, which the feedback's own
   closing line rejects in as many words.

   DOES NOT use the View My Answer Toggle: this template reveals the
   full correct set alongside the learner's own selections at the same
   time, so the toggle's "one slot, two possible values" applicability
   rule excludes it. MultipleChoiceQuestion.md names it a standing
   exclusion.

   ATTEMPT MODEL: the 2-attempt default. The note states no deviation,
   and the slide carries a hint plus exactly one correct and one
   final-wrong panel.
   ═══════════════════════════════════════════════════════════ */
const S6MC = { correct: new Set(['o2', 'o3']), maxAttempts: 2 };

let s6mcSelected  = new Set();
let s6mcAttempts  = 0;
let s6mcDone      = false;
let s6mcResult    = null;   /* 'correct' | 'wrong2' */
let s6mcLastWrong = null;   /* serialized Set of the immediately-preceding wrong attempt */
let s6mcGateOk    = true;   /* false only while a scroll-gated popup has not been read to the end */

/* Stable serialization for the retry gate — sorted, so membership and
   not click order is what gets compared. */
function s6mcSig(set) { return [...set].sort().join('|'); }

/* Shared explanation — the authored body of BOTH panels (id=26 correct,
   id=16 final-wrong), which are structurally identical in the deck (9
   paragraphs each, the same 3 authored blank lines, the same
   expressions) and differ only in title. That is the shared-explanation
   default of _question-template-defaults.md, not an assumption.
   Preserved verbatim, including the source's own empty paragraphs as
   real breaks. Every Latin or numeral group is an isolated,
   non-wrapping LTR unit written in the left-to-right order the deck's
   own export shows.

   ONE AUTHORED CHARACTER IS CHANGED, on the producer's instruction of
   2026-09-14: the source writes the second term as "(7x + K)" with a
   capital K, here and in statement 3's own label, while every other k
   is lower case. Ruled a mathematical-variable consistency error and
   corrected to "(7x + k)". Nothing else in the body is altered — no
   other equation, no ratio order, no wording. See
   docs/pending-assets.md item 3.

   ¶7 and ¶8 are two source paragraphs that finish ONE sentence
   ("…ولذلك النسبة الجديدة ستكون" / "(3x + k) : (7x + k)."), and ¶9 likewise
   continues it; the authored breaks are kept as authored. */
const S6MC_GAP = '';
const S6MC_EXPLAIN = [
  'نوسّع أولًا النسبة بـ <bdi class="scq-expr" dir="ltr">x</bdi> للحصول على الكمية الكلية من كل نوع طحين، ونحصل على النسبة <bdi class="scq-expr" dir="ltr">3x : 7x</bdi>.',
  S6MC_GAP,
  'والآن نضيف <bdi class="scq-expr" dir="ltr">k</bdi> ونحصل على أن كمية طحين الحنطة هي <bdi class="scq-expr" dir="ltr">3x + k</bdi>، وكمية طحين القمح هي <bdi class="scq-expr" dir="ltr">7x + k</bdi>.',
  S6MC_GAP,
  'ولذلك النسبة الجديدة ستكون',
  '<bdi class="scq-expr" dir="ltr">(3x + k) : (7x + k)</bdi>.',
  'هذه النسبة غير مكافِئة للنسبة <bdi class="scq-expr" dir="ltr">3x : 7x</bdi>.'
];

/* `scrollGate` on both long outcomes — slide 70's note asks for it in
   as many words and both source panels are 437.6 × 511.7, taller than
   the canvas. The mechanism unlocks immediately when the body does not
   actually overflow the gated box, so it can never strand the learner;
   whether it is genuinely needed at each viewport is MEASURED after the
   build and reported. No `action` / `answerView`: this template carries
   no View My Answer Toggle. */
const S6MC_POPUP = {
  retry:   { bg: '#ffdbdc', title: 'الإجابة غير صحيحة.', body: ['هل نحاول مرة أخرى؟'] },
  /* Authored title, verbatim (id=26 ¶1). */
  correct: { bg: '#edf8ed', title: 'إجابتكم صحيحة، كل الاحترام!', body: S6MC_EXPLAIN, scrollGate: true },
  /* Authored title, verbatim (id=16 ¶1), plus the canonical fixed
     second line that every terminal-incorrect SOLUTION view carries
     (_question-template-defaults.md → "Terminal-incorrect title carries
     a fixed second line"). Appended to, never replacing, the authored
     line. */
  wrong2:  {
    bg: '#ffdbdc',
    title: 'هذا خطأ. وهذه فرصة للتوقف وتوضيح الفكرة‎.',
    titleSecondLine: 'الإجابة الصحيحة تَظهَر الآن.',
    body: S6MC_EXPLAIN,
    scrollGate: true
  }
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
     dismisses its retry popup. */
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
  /* Gate on "at least one selected" — never on matching the NUMBER of
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
  reportQ(['001', 'q3'], correct, correct || s6mcAttempts >= S6MC.maxAttempts, xapiChoiceAnswer('001', 'q3', '#s6 .scq-opt', [...s6mcSelected]));

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
     fallback "هل نتابع؟" — slide 70 authors no progression copy of its
     own. NOT ValueInputQuestion's "متابعة". The button stays DISABLED
     here because both terminal outcomes are scroll-gated;
     s6mcCheckScrollGate() is what releases it. */
  if (chk) setNavLabel(chk, 'هل نتابع؟');
  hideHintButton('s6mc-hint');
  /* Peak sub-part 3 resolves IMMEDIATELY on the terminal attempt — the
     Peak stage's own dot semantics. The peak AGGREGATE stays unresolved
     until all four sub-parts have resolved; sub-part 4 is S7 (slide 71),
     so on this screen it still reads unresolved. */
  syncPeakProgressNav(document.getElementById('s6'));
}

function s6mcResetInitial() {
  s6mcSelected  = new Set();
  s6mcAttempts  = 0;
  s6mcDone      = false;
  s6mcResult    = null;
  s6mcLastWrong = null;
  s6mcGateOk    = true;
  /* A screen returned to its unanswered state gates again on the next
     terminal outcome — the one-time allowance belongs to a resolution,
     not to the session. */
  resetScrollGate('s6mc');
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

/* Return to a RESOLVED screen: the locked answer state comes back AND
   its feedback reopens (_global-components.md → "Reopening feedback on
   return to a resolved screen"). s6mcDone is the only flag involved —
   there is deliberately no second "popup was open on exit" flag. The
   option marks themselves survive on the DOM, because nothing cleared
   them. Re-showing the popup does NOT re-arm the scroll gate once it
   has been satisfied (producer ruling 2026-09-14 — see
   scrollGateNeeded): the feedback reopens at the top and stays
   scrollable, but the progression button is live immediately. */
function s6mcRestoreResolved() {
  s6mcOpts().forEach(o => { o.disabled = true; });
  const chk = document.getElementById('s6mc-check');
  setNavLabel(chk, 'هل نتابع؟');
  chk.disabled = false;
  hideHintButton('s6mc-hint');
  s6mcShowPopup(s6mcResult === 'correct' ? 'correct' : 'wrong2');
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

  /* The terminal-incorrect title is two bold lines, so the element has
     to render markup rather than plain text for the break to take
     effect at all. */
  const titleEl = document.getElementById('s6mc-popup-title');
  titleEl.innerHTML = cfg.titleSecondLine
    ? (cfg.title + '<br>' + cfg.titleSecondLine)
    : cfg.title;

  body.innerHTML = cfg.body.map(p =>
    p === '' ? '<p class="scq-gap" aria-hidden="true"></p>' : '<p>' + p + '</p>'
  ).join('');

  /* Scroll gate — arm before showing, then reset and measure AFTER the
     popup is visible. Both have to happen after the unhide: a
     display:none element ignores a scrollTop write and reports
     scrollHeight/clientHeight as 0.
     `is-scroll` still tracks cfg.scrollGate on EVERY presentation, so
     the popup keeps its scrollable layout and scrollbar; only whether
     the gate is CHARGED depends on whether it has been read before. */
  popup.classList.toggle('is-scroll', !!cfg.scrollGate);
  const gateNeeded = scrollGateNeeded('s6mc', cfg);
  s6mcGateOk = !gateNeeded;
  if (chk && !s6mcGateOk) chk.disabled = true;
  if (cfg.scrollGate) scroller.setAttribute('tabindex', '0');
  else                scroller.removeAttribute('tabindex');

  popup.classList.remove('hidden');
  clampPopup(popup);
  scroller.scrollTop = 0;
  if (gateNeeded) {
    placeScrollHint('s6mc-popup');
    s6mcCheckScrollGate();   /* unlocks at once when the content does not actually overflow */
  } else {
    /* Already read once, or never gated: reopen at the top with the
       progression button live. Rereading stays possible, not required. */
    popup.classList.remove('needs-scroll');
    if (chk && s6mcDone) chk.disabled = false;
  }
}

function s6mcCheckScrollGate() {
  if (s6mcGateOk) return;
  const scroller = document.getElementById('s6mc-popup-scroll');
  const popup    = document.getElementById('s6mc-popup');
  if (!scroller || !popup) return;
  if (!atBottom(scroller)) { popup.classList.add('needs-scroll'); return; }
  s6mcGateOk = true;
  /* Recorded so no later presentation of the same feedback charges it
     again. On this screen the body does not overflow, so the gate
     releases on the first render — which still counts as read. */
  markScrollGateSatisfied('s6mc');
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

function s6mcOpenHint()  { xapiRequestedHint('001', 'q3'); document.getElementById('s6mc-hint-overlay').classList.remove('hidden'); }
function s6mcCloseHint() { document.getElementById('s6mc-hint-overlay').classList.add('hidden'); }
function s6mcCloseHintOnBackdrop(e) { if (e.target === e.currentTarget) s6mcCloseHint(); }

/* ═══════════════════════════════════════════════════════════
   S7 — ValueInputQuestion (two numeric fields, inline)
   Peak sub-part 4 — the LAST scored sub-item. Source: slide 71.

   Spec: 720-templates → ValueInputQuestion.md + _question-template-
   defaults.md + _global-components.md → Feedback popup / View My Answer
   Toggle / Long-feedback scroll gate.

   This screen is the first in this part to need BOTH the toggle and the
   scroll gate on one popup. That combination is not new to the project:
   the approved sibling methodica-ar-math-ratio-03-04's own S5 is exactly
   it, and its ordering inside the show-popup function (render body →
   render action row → arm gate → unhide → measure) is followed here
   rather than re-derived.

   EVERYTHING SCORED IS AUTHORED. Slide 71's own note (via
   slide71.xml.rels → notesSlide65.xml, cached slidenum "71") gives the
   template, "שאלה 4 מתוך 4", and the key in as many words:
       כוסמין 90
       חיטה 210
   It is corroborated by the slide's own pre-filled indicator boxes
   (id=7 "90", id=31 "210") and by the feedback's own worked solution.
   With slide 70's note reading "שאלה 3 מתוך 4", PEAK_SUBPARTS is now
   complete and every sub-part number came from its own slide.

   ONE SCORED UNIT, TWO FIELDS (Variant C / Case A): the sub-part is
   correct only if BOTH fields are correct, but each field shows its own
   verdict, so a mixed board is possible and correct.

   ATTEMPT MODEL: the 2-attempt default; the note states no deviation.
   ═══════════════════════════════════════════════════════════ */
const S7_KEY = { spelt: 90, wheat: 210 };
const S7_MAX_ATTEMPTS = 2;

const S7_IMAGE = {
  /* The file slide 71's note names — "שם התמונה:< מאפייה הוספה>" — and
     the real filename on disk. 431 × 287, uncropped.

     REPLACED 2026-09-14 on the producer's instruction. The originally
     supplied 1536 × 1024 PNG carried baked-in callouts, one of which
     ("קמח חיטה +20 גרם") stated the OPPOSITE direction from the stem,
     the feedback and the key — the stem removes 20 g of wheat. The
     ruling was to use slide 71's own clean embedded picture
     (ppt/media/image74.jpeg) and to add no callouts, labels, arrows,
     chalkboard or anything else the slide does not have.

     It is transcoded JPEG → PNG so the filename, and therefore the
     note's own image name, is unchanged; PNG is lossless, so no second
     generation of JPEG loss is introduced. The superseded file is kept
     at docs/superseded-assets/ rather than destroyed.

     The deck stores it with cstate="email", i.e. PowerPoint's own
     compressed derivative, which is why it is only 431 × 287. That is
     1:1 with the 431.2 × 283.7 display box at scale 1 and upscales 1.5×
     at 1920×1080 — measured and reported, not hidden. */
  src: 'assets/img/מאפייה הוספה.jpg',
  alt: 'خبازة تقف بجانب وعاءين من الطحين على طاولة مخبز وتضيف طحينًا إلى أحدهما'
};
function s7MountImage() {
  const img = document.getElementById('s7-img');
  if (!img) return;
  if (!img.getAttribute('src') || !img.getAttribute('src').endsWith(S7_IMAGE.src)) {
    img.src = S7_IMAGE.src;
  }
  img.alt = S7_IMAGE.alt;
}

let s7viqAttempts = 0;
let s7viqDone     = false;
let s7viqResult   = null;   /* 'correct' | 'incorrect' */
let s7viqGateOk   = true;   /* false only while a scroll-gated popup is unread */
/* Retry gate — same shape as S2 above (and 03-01 S22). */
let s7viqLastWrongSig = null;
/* The learner's RAW submitted strings, snapshotted BEFORE the reveal
   overwrites the same inputs — the only moment they still exist. */
let s7viqAnswer   = null;
let s7viqView     = 'solution';

function s7viqInput(k) { return document.getElementById('s7viq-input-' + k); }
function s7viqField(k) { return document.getElementById('s7viq-field-' + k); }
function s7viqIco(k)   { return document.getElementById('s7viq-ico-' + k); }

/* Numeric normalisation: trim, drop a thousands comma. Everything else
   is left alone, so "90 " and "90" pass while "9 0", "90g" and "ninety"
   do not. The NaN guard matters — Number('') is 0, which would make an
   empty field read as a wrong-but-parsed answer rather than as empty. */
function s7viqNorm(v) { return String(v == null ? '' : v).trim().replace(/,/g, ''); }
function s7viqSignature() {
  return ['spelt', 'wheat'].map(k => s7viqNorm(s7viqInput(k).value)).join('|');
}
function s7viqFieldCorrect(k) {
  const raw = s7viqNorm(s7viqInput(k).value);
  if (raw === '') return false;
  const n = Number(raw);
  return Number.isFinite(n) && n === S7_KEY[k];
}

/* Check enables only when BOTH fields carry something.
   Fires on a GENUINE value change only — see the note on s2viqOnInput
   for why `input` is the right hook and focus/caret/selection are not. */
function s7viqOnInput() {
  /* type="text" (inputmode="numeric" only hints the on-screen keyboard,
     it does not block a physical keyboard or paste) so every keystroke
     AND paste is filtered here — same shape as every other numeric
     ValueInput screen in the lomda (_question-template-defaults.md →
     "Numeric input format"). Fix 2026-09-24: this screen had no filter
     at all, so a non-digit character (e.g. a Hebrew letter) could sit
     in the field. */
  ['spelt', 'wheat'].forEach(k => {
    const el = s7viqInput(k);
    const cleaned = el.value.replace(/[^0-9,.-]/g, '');
    if (cleaned !== el.value) el.value = cleaned;
  });
  if (s7viqDone) return;
  /* Retry-feedback dismissal — approved path 2 of "No manual close",
     scoped to a non-final wrong attempt. Added 2026-09-14: this screen
     had no dismissal path at all, so its retry feedback stayed on screen
     while the learner corrected.

     THE MARKS GO WITH IT, for the same reason as on S2: a verdict icon
     describes the value standing in its field, so a changed value must
     not keep the previous attempt's ✓/✗. Both fields clear, since the
     pair was judged together. Only a genuine value change reaches here. */
  if (s7viqAttempts > 0) {
    document.getElementById('s7viq-popup').classList.add('hidden');
    s7viqClearMarks();
  }
  const filled = ['spelt', 'wheat'].every(k => s7viqInput(k).value.trim() !== '');
  document.getElementById('s7viq-check').disabled =
    !filled || s7viqSignature() === s7viqLastWrongSig;
}

/* Mark each field with ITS OWN verdict (Case A). Reuses the shared icon
   markup already defined for #s2. */
function s7viqMarkFields() {
  ['spelt', 'wheat'].forEach(k => {
    const ok = s7viqFieldCorrect(k);
    const f = s7viqField(k);
    f.classList.toggle('viq-correct', ok);
    f.classList.toggle('viq-wrong', !ok);
    s7viqIco(k).innerHTML = ok ? S2_ICO_CHECK : S2_ICO_X;
  });
}
function s7viqClearMarks() {
  ['spelt', 'wheat'].forEach(k => {
    s7viqField(k).classList.remove('viq-correct', 'viq-wrong');
    s7viqIco(k).innerHTML = '';
  });
}

/* Authored feedback, verbatim from slide 71. Both panels carry the SAME
   worked solution, paragraph for paragraph (id=26 ¶3-¶17, id=28 ¶4-¶18),
   so this is the shared-explanation default rather than an assumption.
   The authored blank paragraphs are preserved as real blank lines, and
   each standalone equation keeps the centred line the source sets it on
   (algn="ctr" on its own paragraph) via the `@math:` prefix.

   ONE AUTHORED ODDITY REPRODUCED AS WRITTEN, recorded in
   docs/pending-assets.md:
     · "120 = 4x / :4" — the divide-both-sides step, written with the
       slash-colon notation the author uses.
   ONE AUTHORED LINE CORRECTED (producer, 2026-09-22, on reviewer Ofer's
   note; reverses the 2026-09-14 keep-as-authored ruling):
     · the source's "70 ⋅ 3 = 210" is now "7 ⋅ 30 = 210" — the ratio term
       times x, the two quantities that actually exist in this problem.
       70 ⋅ 3 was numerically true, so the key (210) never changed; only
       the derivation a learner is meant to follow did. Written in the
       reviewer's order (7 ⋅ 30), not the sibling line's (30 ⋅ 3), by the
       producer's instruction. Both panels share this array, so the
       correct and the final-wrong feedback changed together. */
const S7_GAP = '';
const S7VIQ_EXPLAIN = [
  'نضيف إلى طحين الحنطة 100 غرام‎:',
  '@math:<bdi class="scq-expr" dir="ltr">3x + 100</bdi>',
  'نطرح من طحين القمح 20 غرام:',
  '@math:<bdi class="scq-expr" dir="ltr">7x - 20</bdi>',
  S7_GAP,
  'نساوي بينهما، فنحصل على المعادلة ونحلها‎:',
  '@math:<bdi class="scq-expr" dir="ltr">3x + 100 = 7x - 20</bdi>',
  '@math:<bdi class="scq-expr" dir="ltr">120 = 4x / :4</bdi>',
  '@math:<bdi class="scq-expr" dir="ltr">x = 30</bdi>',
  S7_GAP,
  'نعوّض <bdi class="scq-expr" dir="ltr">x = 30</bdi> ونحصل على أن الكمية الأصلية من طحين الحنطة هي‎:',
  '@math:<bdi class="scq-expr" dir="ltr">30 ⋅ 3 = 90</bdi>',
  S7_GAP,
  'الكمية الأصلية من طحين القمح هي‎:',
  '@math:<bdi class="scq-expr" dir="ltr">7 ⋅ 30 = 210</bdi>'
];

/* The final-wrong panel opens with a SECOND bold line
   ("לא נורא, בואו נסביר:", id=28 ¶2) before the authored blank that
   separates title from body in both panels. It is carried as the first
   BODY paragraph in bold rather than as part of the title: that keeps
   the canonical terminal-wrong second line adjacent to the authored
   title, as the contract requires, and mirrors the correct panel's
   structure exactly. */
const S7VIQ_EXPLAIN_WRONG = ['@lead:لا بأس، دعونا نشرح‎:', S7_GAP].concat(S7VIQ_EXPLAIN);

/* `scrollGate` on the two long outcomes — the note asks for it in as
   many words and both source panels are 437.6 × 607.9, taller than the
   canvas. `action` / `answerView` implement the View My Answer Toggle,
   which this template REQUIRES on the final-wrong path because the
   reveal is destructive: it overwrites the very inputs the learner
   typed into. Only `wrong2` ever carries them, so the toggle can never
   appear on the retry or correct paths.

   The answer view is NOT gated: its body is one canonical line, so a
   gate there would have nothing to scroll. Switching back to the
   solution does not re-arm the gate either, once it has been satisfied
   (producer ruling 2026-09-14 — see scrollGateNeeded). The toggle still
   cannot be used to skip the explanation, because the action row that
   carries it is rendered at the END of the scroll column: reaching it
   at all means having scrolled through the body. */
const S7VIQ_POPUP = {
  /* No retry-specific pedagogical explanation is authored, so the
     canonical fallback is used — a missing explanation is not invented. */
  retry:   { bg: '#ffdbdc', title: 'الإجابة غير صحيحة.', body: ['هل نحاول مرة أخرى؟'] },
  /* Authored title, verbatim (id=26 ¶1). */
  correct: { bg: '#edf8ed', title: 'هذا صحيح‎!', body: S7VIQ_EXPLAIN, scrollGate: true },
  wrong2: {
    bg: '#ffdbdc',
    /* Authored title, verbatim (id=28 ¶1). */
    title: 'غير دقيق.',
    titleSecondLine: 'الإجابة الصحيحة تَظهَر الآن.',
    body: S7VIQ_EXPLAIN_WRONG,
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
      body: ['يمكن مشاهدة الحل الصحيح مرة أخرى‎.'],
      action: { label: 'الإجابة الصحيحة', view: 'solution' },
      scrollGate: false
    }
  }
};

function s7viqShowPopup(type, view) {
  const popup = document.getElementById('s7viq-popup');
  if (!popup) return;
  const base = S7VIQ_POPUP[type];
  const cfg  = (view === 'answer' && base.answerView)
    ? Object.assign({}, base, base.answerView)
    : base;
  const body     = document.getElementById('s7viq-popup-body');
  const scroller = document.getElementById('s7viq-popup-scroll');
  const chk      = document.getElementById('s7viq-check');

  popup.style.background = cfg.bg;
  /* Re-anchor in TWO cases: on a fresh open, AND on every `wrong2`
     render including each toggle between its views — a dragged popup
     converts to a fixed pixel `top`, and if that survives a view swap
     that changes its height the box shrinks from its pinned top edge
     and leaves a gap above the bottom bar. */
  if (popup.classList.contains('hidden') || type === 'wrong2') resetPopupPosition(popup);

  const titleEl = document.getElementById('s7viq-popup-title');
  titleEl.innerHTML = cfg.titleSecondLine
    ? (cfg.title + '<br>' + cfg.titleSecondLine)
    : cfg.title;

  body.innerHTML = cfg.body.map(p => {
    if (p === '')                  return '<p class="scq-gap" aria-hidden="true"></p>';
    if (p.indexOf('@math:') === 0) return '<p class="scq-math">' + p.slice(6) + '</p>';
    if (p.indexOf('@lead:') === 0) return '<p class="scq-lead">' + p.slice(6) + '</p>';
    return '<p>' + p + '</p>';
  }).join('');
  /* Renders the toggle button, or removes the row entirely when
     `cfg.action` is absent (retry / correct). */
  renderPopupAction('s7viq-popup', cfg.action, s7viqToggleView);

  /* Scroll gate — arm before showing, then reset and measure AFTER the
     popup is visible: a display:none element ignores a scrollTop write
     and reports scrollHeight/clientHeight as 0.
     `is-scroll` still tracks cfg.scrollGate on EVERY presentation, so
     the popup keeps its scrollable layout and scrollbar; only whether
     the gate is CHARGED depends on whether it has been read before.
     This is the screen the ruling was written against: the toggle
     re-renders `wrong2` on every view switch, so without the
     first-presentation rule a learner who had already read the
     explanation was re-locked each time they looked at their own
     answer and came back. */
  popup.classList.toggle('is-scroll', !!cfg.scrollGate);
  const gateNeeded = scrollGateNeeded('s7viq', cfg);
  s7viqGateOk = !gateNeeded;
  if (chk && !s7viqGateOk) chk.disabled = true;
  if (cfg.scrollGate) scroller.setAttribute('tabindex', '0');
  else                scroller.removeAttribute('tabindex');

  popup.classList.remove('hidden');
  clampPopup(popup);
  scroller.scrollTop = 0;
  if (gateNeeded) {
    placeScrollHint('s7viq-popup');
    s7viqCheckScrollGate();   /* unlocks at once when the content does not actually overflow */
  } else {
    /* Already read once, or never gated (the answer view): reopen at
       the top with the progression button live. Rereading stays
       possible, not required. */
    popup.classList.remove('needs-scroll');
    if (chk && s7viqDone) chk.disabled = false;
  }
}

function s7viqCheckScrollGate() {
  if (s7viqGateOk) return;
  const scroller = document.getElementById('s7viq-popup-scroll');
  const popup    = document.getElementById('s7viq-popup');
  if (!scroller || !popup) return;
  if (!atBottom(scroller)) { popup.classList.add('needs-scroll'); return; }
  s7viqGateOk = true;
  /* Recorded so no later presentation of the same feedback charges it
     again — neither a toggle round trip nor a return to the screen. */
  markScrollGateSatisfied('s7viq');
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

/* ─── View My Answer Toggle ─────────────────────────────────
   Single owner of the resolved-wrong board+popup state. Writes the
   displayed values for the current view into the live fields, re-marks
   each field from what is now displayed, re-locks (a re-render can
   recreate unlocked elements) and re-renders the popup in that view.

   Field marking follows the DISPLAYED values, which is what makes both
   views correct for free: the solution view shows 90/210 so both fields
   mark ✓, and the answer view shows the learner's own raw strings so
   each field marks by its own result. */
function s7viqApplyView() {
  if (s7viqAnswer === null) return;
  ['spelt', 'wheat'].forEach(k => {
    s7viqInput(k).value = (s7viqView === 'answer') ? s7viqAnswer[k] : String(S7_KEY[k]);
  });
  s7viqMarkFields();
  ['spelt', 'wheat'].forEach(k => { s7viqInput(k).disabled = true; });
  s7viqShowPopup('wrong2', s7viqView);
}

/* A PURE view switch. It never routes through s7viqCheck(), and it never
   touches attempts, done state, the sub-part result, the progress dot,
   navigation state or any reporting call — the contract forbids all of
   those. Nor does it re-arm the scroll gate: once the explanation has
   been read to the end, a view switch never re-locks the progression
   button (producer ruling 2026-09-14). */
function s7viqToggleView() {
  if (s7viqAnswer === null) return;
  s7viqView = (s7viqView === 'solution') ? 'answer' : 'solution';
  s7viqApplyView();
}

function s7viqCheck() {
  if (s7viqDone) { advanceScreen(); return; }
  if (!['spelt', 'wheat'].every(k => s7viqInput(k).value.trim() !== '')) return;
  s7viqAttempts++;
  const allOk = ['spelt', 'wheat'].every(k => s7viqFieldCorrect(k));
  /* xAPI (2026-11-11): the layer writes Kata's form from the metadata — performance, 'קמח כוסמין[.]90[,]קמח חיטה[.]210' (proposed to Kata, `other` there) */
  reportQ(['001', 'q4'], allOk, allOk || s7viqAttempts >= S7_MAX_ATTEMPTS,
    { steps: { 'طحين الحنطة': String(s7viqInput('spelt').value).trim(), 'طحين القمح': String(s7viqInput('wheat').value).trim() } });

  if (allOk) {
    s7viqResult = 'correct';
    s7viqAnswer = null;            /* nothing to toggle to */
    s7viqMarkFields();
    s7viqFinish();
    s7viqShowPopup('correct');
    return;
  }

  if (s7viqAttempts < S7_MAX_ATTEMPTS) {
    /* Non-final wrong: fields carry their own marks, inputs stay
       editable, hint stays available, nothing is resolved, and NO
       toggle appears — `retry` carries no action. */
    s7viqMarkFields();
    s7viqLastWrongSig = s7viqSignature();
    document.getElementById('s7viq-check').disabled = true;
    S7VIQ_POPUP.retry.title = ['spelt', 'wheat'].some(k => s7viqFieldCorrect(k)) ? RETRY_TITLE_PARTIAL : RETRY_TITLE;
    s7viqShowPopup('retry');
    return;
  }

  /* FINAL wrong. Snapshot the learner's RAW submitted strings BEFORE the
     reveal overwrites the same fields. Stored exactly as typed, so
     " 90 " comes back byte for byte rather than being rebuilt from the
     normalised comparison. */
  s7viqAnswer = { spelt: s7viqInput('spelt').value, wheat: s7viqInput('wheat').value };
  s7viqResult = 'incorrect';
  s7viqView   = 'solution';        /* the solution is shown IMMEDIATELY */
  s7viqFinish();
  s7viqApplyView();
}

/* Terminal state: lock the inputs, relabel Check as the progression
   action, hide the hint, and resolve this sub-part's dot IMMEDIATELY
   (peak semantics). The toggle never changes this button's label.
   The button stays DISABLED here when the outcome is scroll-gated;
   s7viqCheckScrollGate() is what releases it. */
function s7viqFinish() {
  s7viqDone = true;
  ['spelt', 'wheat'].forEach(k => { s7viqInput(k).disabled = true; });
  setNavLabel(document.getElementById('s7viq-check'), 'متابعة');
  hideHintButton('s7viq-hint');
  /* This is sub-part 4 — the one whose resolution finally lets the peak
     AGGREGATE resolve. Nothing extra is done for that here: the
     aggregate is derived, not stored, so syncing is enough. */
  syncPeakProgressNav(document.getElementById('s7'));
}

function s7viqResetInitial() {
  s7viqAttempts = 0;
  s7viqLastWrongSig = null;
  s7viqDone     = false;
  s7viqResult   = null;
  s7viqGateOk   = true;
  /* A screen returned to its unanswered state gates again on the next
     terminal outcome — the one-time allowance belongs to a resolution,
     not to the session. */
  resetScrollGate('s7viq');
  s7viqAnswer   = null;
  s7viqView     = 'solution';
  ['spelt', 'wheat'].forEach(k => {
    const el = s7viqInput(k);
    el.value = '';
    el.disabled = false;
  });
  s7viqClearMarks();
  const check = document.getElementById('s7viq-check');
  setNavLabel(check, 'هل إجابتي صحيحة؟');
  check.disabled = true;
  showHintButton('s7viq-hint');
  renderPopupAction('s7viq-popup', null);
  const popup = document.getElementById('s7viq-popup');
  popup.classList.add('hidden');
  popup.classList.remove('is-scroll', 'needs-scroll');
  document.getElementById('s7viq-hint-overlay').classList.add('hidden');
}

/* Restore a RESOLVED screen on return: the resolved state comes back
   with its feedback, never as a fresh attempt. The scroll gate is NOT
   re-armed once satisfied (producer ruling 2026-09-14): the feedback
   reopens at the top and stays scrollable, but שנמשיך? / המשך is live
   immediately.
   Back ALWAYS lands on the SOLUTION view, never on whichever view the
   learner last had open — the toggle contract states that explicitly,
   and the learner's answer stays exactly one click away. */
function s7viqRestoreResolved() {
  const check = document.getElementById('s7viq-check');
  setNavLabel(check, 'متابعة');
  hideHintButton('s7viq-hint');
  if (s7viqResult === 'correct') {
    ['spelt', 'wheat'].forEach(k => { s7viqInput(k).disabled = true; });
    s7viqMarkFields();
    renderPopupAction('s7viq-popup', null);
    s7viqShowPopup('correct');
    return;
  }
  s7viqView = 'solution';
  s7viqApplyView();
}

function s7viqOpenHint()  { xapiRequestedHint('001', 'q4'); document.getElementById('s7viq-hint-overlay').classList.remove('hidden'); }
function s7viqCloseHint() { document.getElementById('s7viq-hint-overlay').classList.add('hidden'); }
function s7viqCloseHintOnBackdrop(e) { if (e.target === e.currentTarget) s7viqCloseHint(); }

/* ─── S8 image asset ────────────────────────────────────────
   The file slide 72's note names — "שם התמונה:<פסטיבל אוכל>" — and the
   real filename on disk.

   REPLACED 2026-09-14 with the deck's own embedded copy
   (ppt/media/image75.jpeg, 1282 × 1026), transcoded JPEG → PNG so the
   path, and therefore the note's image name, is unchanged. The file
   previously here (1402 × 1122) is a superseded render of the same
   scene in which the woman at centre-right wears a sleeveless גופיה —
   exactly what this slide's own producer comment
   ("מאיה - רק לדאוג לתמונה בלי גופיה", status="resolved") asked to
   avoid. The comment is resolved in the deck's artwork and was not
   resolved in the asset folder. Same handling the producer ruled for
   S7's photo; the superseded file is kept under docs/superseded-assets/.
   Reported for confirmation — a corrected high-resolution PNG dropped in
   at this path needs no code change. */
const S8_IMAGE = {
  src: 'assets/img/פסטיבל אוכל.jpg',
  alt: 'كشك خبز يحمل اسم "خبزُ سمر" في مهرجان طعام في الهواء الطلق؛ سمر تقدّم شريحة خبز لإحدى الزائرات، والناس من حولها يتذوقون ويتحدثون'
};
function s8MountImage() {
  const img = document.getElementById('s8-img');
  if (!img) return;
  if (!img.getAttribute('src') || !img.getAttribute('src').endsWith(S8_IMAGE.src)) {
    img.src = S8_IMAGE.src;
  }
  img.alt = S8_IMAGE.alt;
}

/* ═══════════════════════════════════════════════════════════
   END OF PART — the crossing into part 06

   S8 (slide 72) is the last screen of this part: the deck's own
   component dividers bound it at order 63 (רכיב 5 / שאלת שיא א) and
   order 73 (רכיב 6 / שאלת שיא ב), so 64-72 is the whole part and 72 is
   its closing narrative screen. Advancing off it is therefore a real
   boundary for the first time, not the bounds-checked no-op every
   earlier screen's note described.

   GATED ON THE PART'S AUTHORED MASTERY RULE. Slide 72 itself carries no
   gate, but the rule is authored on this part's OPENING screen, S0
   (slide 64), in as many words:

       עכשיו מגיעה שאלת השיא
       יש בה 4 סעיפים. עליכם להצליח ב - 3 לפחות.

   — four sub-items, at least three correct. That is the same line
   docs/scoring-model.md has recorded since before the first question
   screen was built, and it is what PEAK_THRESHOLD = 3 already
   implements. END-OF-PART ROUTING IS A WHOLE-PART CONTRACT: it is
   derived from the part's own opening Transition Screen, never from the
   final slide alone. An earlier pass of this file read only slide 72,
   found no rule there, and wrongly crossed unconditionally; corrected on
   the producer's instruction, 2026-09-14.

   THE PEAK RESULT IS STILL NOT REPORTED. It gates routing here, and it
   is published in memory by peakQuestionResult() for the integration
   layer; emitting it — SCORM, xAPI, cmi.*, completion postMessage — is
   Vadim's layer and stays out of this build. Continue is never DISABLED
   by an outcome: a failing learner keeps an enabled control that simply
   does not advance.

   MECHANISM. PartNav.forward() (js/partnav.js), the project's existing
   cross-part contract, byte-identical to the copies in 03-01/02/03/04.
   It pushes this part and the screen being left onto the sessionStorage
   trail before navigating, which is what gives 03-06's S0 a Back button
   returning HERE, to S8, rather than to a hardcoded predecessor. No ad
   hoc window.location assignment. The selected character needs no
   special handling: it lives in localStorage under the lomda-wide
   CHARACTER_STORAGE_KEY that 03-06 reads.

   RE-CROSSING. If the learner already crossed and came back,
   PartNav.forwardOverride() reports the destination recorded for their
   journey and that destination is REPLAYED rather than re-derived
   (_global-components.md → "Re-crossing a previously-resolved boundary
   forward must replay the recorded destination"). The trail lives in
   sessionStorage, which a learner can edit, so a replayed id that is
   not this part's own destination is refused rather than followed.
   ═══════════════════════════════════════════════════════════ */
const NEXT_PART = '06';

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
  if (!peakQuestionResult().resolved) return;
  if (xapiEndPart(lastScreenButton(), NEXT_PART)) partCrossingInFlight = true;
}

/* ═══════════════════════════════════════════════════════════
   xAPI seams — read at call time by unit-js/20-xapi.js, 22-scoring.js and 50-loader.js.
   Ids must match metadata/methodica-ar-math-ratio-03-05.json byte-for-byte (the loader
   checks XAPI_COMP_ID against METADATA.id on every load; _test/verify-static.js checks the map).
   ═══════════════════════════════════════════════════════════ */
const XAPI_COMP_SLUG     = 'methodica-ar-math-ratio-03-05';
const XAPI_COMP_ID       = XAPI_ID_PREFIX + XAPI_COMP_SLUG + '/';
const XAPI_METADATA_FILE = '../metadata/' + XAPI_COMP_SLUG + '.json';
/* Items carrying a code-graded question (each declared question is reported by its screen). */
const XAPI_EVAL_ITEMS    = { '001': 1 };
/* screen → [item suffix, page within the item]; null = no catalogue item. Exactly TOTAL_SCREENS
   keys, 0-based (keep it a literal object — the static check reads it). */
const SCREEN_TO_SUBCONTENT = {
  0: ['001', 1],
  1: ['001', 2],
  2: ['001', 3],
  3: ['001', 4],
  4: ['001', 5],
  5: ['001', 6],
  6: ['001', 7],
  7: ['001', 8],
  8: ['001', 9]
};

/* The gate S0 promises: "3 of 4 sections" (peakQuestionResult). Score = sections correct / 4. */
function partGateResult() {
  const p = peakQuestionResult();
  return { success: !!p.scoredCorrect, score: { scaled: p.subPartsCorrect / p.subPartCount } };
}
function lastScreenButton() { return currentScreen === PART_LAST ? document.querySelector('#s8 .btn-continue') : null; }

/* ═══════════════════════════════════════════════════════════
   Resume painters — unit-js/45-painters.js reads this registry.
   What each screen saves (vars + the fields that hold state) and how a screen restored from the
   state document is drawn again on the fresh markup: init() is the screen's own ResetInitial,
   paint() replays the template's own terminal look (Mark / Select / Render / Finish /
   RestoreResolved) or the retry look, without counting an attempt and without reporting. In-session
   re-entry still goes through resetScreenState(), which now resets only a pristine screen.
   Gates are state (rule 11): the S3 watch gate and the scroll-gate registry are saved; the peak
   bar's visited flags are saved so a restored bar shows the route the learner actually took.
   ═══════════════════════════════════════════════════════════ */
const PART_STATE_VARS = ['peakProgress', 'scrollGateSatisfied'];
const PART_SCREENS = {
  2: {
    vars: ['s2viqAttempts', 's2viqDone', 's2viqResult', 's2viqAnswer', 's2viqView', 's2viqLastWrongSig'],
    inputs: ['s2viq-input-spelt', 's2viq-input-wheat'],
    pristine: () => !s2viqAttempts && !s2viqDone && resumeFieldsEmpty(PART_SCREENS[2].inputs),
    init: s2viqResetInitial,
    paint: () => {
      if (s2viqDone) { s2viqFinish(); s2viqRestoreResolved(); return; }
      s2viqOnInput();                              // Check enablement + the retry gate
      if (s2viqAttempts && s2viqSignature() === s2viqLastWrongSig) s2viqMarkFields();
    }
  },
  3: {
    vars: ['s3vidDone'],
    pristine: () => !s3vidDone,
    init: s3VideoResetInitial,
    paint: () => { s3MountVideo(); s3VideoRestoreResolved(); }
  },
  4: {
    vars: ['s4scSelected', 's4scAttempts', 's4scDone', 's4scResult', 's4scLastWrong', 's4scGateOk'],
    pristine: () => !s4scSelected && !s4scAttempts && !s4scDone,
    init: s4scResetInitial,
    paint: () => {
      const done = s4scDone, sel = s4scSelected;
      s4scDone = false;
      if (sel) s4scSelect(sel);                    // selection look + the retry gate
      if (done) {
        s4scOpt(sel).classList.add(s4scResult === 'correct' ? 'correct' : 'wrong');
        if (s4scResult !== 'correct') s4scOpt(S4SC.correctId).classList.add('correct');
        s4scFinish();
        s4scRestoreResolved();                     // popup + the scroll gate as saved
      } else if (s4scAttempts && sel === s4scLastWrong) {
        s4scOpt(sel).classList.add('wrong');
      }
    }
  },
  6: {
    vars: ['s6mcSelected', 's6mcAttempts', 's6mcDone', 's6mcResult', 's6mcLastWrong', 's6mcGateOk'],
    pristine: () => !s6mcSelected.size && !s6mcAttempts && !s6mcDone,
    init: s6mcResetInitial,
    paint: () => {
      const done = s6mcDone;
      s6mcDone = false;
      s6mcRender();                                // selection look + the retry gate
      if (done) {
        s6mcOpts().forEach(o => {
          o.classList.remove('selected', 'correct', 'wrong');
          if (S6MC.correct.has(o.dataset.id)) o.classList.add('correct');
          else if (s6mcResult !== 'correct' && s6mcSelected.has(o.dataset.id)) o.classList.add('wrong');
        });
        s6mcFinish();
        s6mcRestoreResolved();                     // popup + the scroll gate as saved
      } else if (s6mcAttempts && s6mcSig(s6mcSelected) === s6mcLastWrong) {
        s6mcOpts().forEach(o => {
          if (!s6mcSelected.has(o.dataset.id)) return;
          o.classList.remove('selected');
          o.classList.add(S6MC.correct.has(o.dataset.id) ? 'correct' : 'wrong');
        });
      }
    }
  },
  7: {
    vars: ['s7viqAttempts', 's7viqDone', 's7viqResult', 's7viqAnswer', 's7viqView', 's7viqLastWrongSig', 's7viqGateOk'],
    inputs: ['s7viq-input-spelt', 's7viq-input-wheat'],
    pristine: () => !s7viqAttempts && !s7viqDone && resumeFieldsEmpty(PART_SCREENS[7].inputs),
    init: s7viqResetInitial,
    paint: () => {
      if (s7viqDone) { s7viqFinish(); s7viqRestoreResolved(); return; }
      s7viqOnInput();                              // Check enablement + the retry gate
      if (s7viqAttempts && s7viqSignature() === s7viqLastWrongSig) s7viqMarkFields();
    }
  }
};
