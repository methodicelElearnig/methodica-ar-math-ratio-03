'use strict';

/* ═══════════════════════════════════════════════════════════
   js/main.js  —  methodica-math-ratio-03-06
   720 reference engine (canvas 1280×710) + Screen 00
   (TransitionScreen — Peak Question B launch).

   Source: the approved production script
   יעד 1.3_יחס וביטויים אלגבריים_להפקה_06092026.pptx, two levels above
   this part's folder. Slide identity re-derived THIS session from a
   fresh extraction of the .pptx (the file's mtime advances between
   sessions — it currently reads 2026-09-14 09:01 — so nothing was
   trusted from an earlier extraction): presentation.xml's <p:sldIdLst>
   r:id → ppt/_rels/presentation.xml.rels gives order 73 → slide73.xml
   and order 74 → slide74.xml (1:1 with no offset across all 83 slides —
   derived, not assumed). Presenter notes resolved through each slide's
   OWN .rels: slide74 → rId3 → notesSlides/notesSlide67.xml (part-name
   numbering is NOT slide numbering — 74 ≠ 67 — pptx-script-extraction),
   whose cached <a:fld type="slidenum"> reads "74".

   PART BOUNDARY, from the deck's own component-divider slides. All six
   dividers in the deck were located by content, not by offset:
     order  2 — "רכיב 1" / "הוק והקניה"              / ...-03-01
     order 28 — "רכיב 2" / "תרגול סטנדרטי"           / ...-03-02
     order 38 — "רכיב 3" / "תרגול בסיסי וסטנדרטי ב"  / ...-03-03
     order 50 — "רכיב 4" / "תרגול מתקדם"             / ...-03-04
     order 63 — "רכיב 5" / "שאלת שיא א"              / ...-03-05
     order 73 — "רכיב 6" / "שאלת שיא ב"              / ...-03-06
   Order 73 is the LAST divider, so this part runs to the deck's final
   slide: Part 06 is slides 74-83 — ten production slides. Slide 73 is
   production-only and is NOT a screen.

   Screen roster (built so far):
     S0 — TransitionScreen, Peak Question B launch (slide 74).
     S1 — FullImageWithTextOverlay, full-width-card variant, the
          "הקונטרול" game intro (slide 75).
     S2 — FullImageWithTextOverlay, the radar-alert screen, with the
          producer's own photo crop and a rising alert panel
          (slide 76).
     S3 — MultipleChoiceQuestion, text-pill variant — PEAK QUESTION 1
          of 4 (slide 77).
   Slides 78-83 are NOT built yet and are out of this step's scope.

   SOURCE RE-EXTRACTED for S1: the producer saved the deck again at
   2026-09-14 10:14 (md5 0eb1ccb8…, previously 3496629e…) mid-build. A
   slide-by-slide diff of the two versions shows the edit touched
   slides 58-62 and their notes only — Part 04's range. Every slide in
   Part 06's 73-83 range is byte-identical across both extractions, so
   S0 is unaffected. Nothing was carried over from the earlier
   extraction regardless; both slides were re-identified by content.

   S0 ported from the same-lomda sibling methodica-math-ratio-03-05's
   own S0 (slide 64). That is a stronger port than the usual
   same-role reuse: a full XML diff of slide64 vs slide74 this session
   shows the two slides are IDENTICAL in every shape, every text run,
   every EMU coordinate and every media relationship — they differ only
   in per-shape creationId GUIDs, the tags part, and the presenter
   note's pose line. -05 took the screen from -04 (slide 51), which took
   it from -03 (slide 39), which took it from -02 (slide 29), per
   720-templates → references/TransitionScreen.md +
   references/_global-components.md.
   Engine plumbing (scaleApp / goTo / resetScreenState / keyboard nav /
   setNavLabel / dev bridge) ported from methodica-math-ratio-03-05,
   byte-identical apart from this part's own screen roster and the
   PartNav part id, so every part of this lomda behaves the same.
   Cross-part Back uses this part's own js/partnav.js (PartNav).

   NO REPORTING LAYER. This file emits no SCORM, xAPI, postMessage or
   Ministry-of-Education outcome signal of any kind. Outcome reporting
   is owned externally by Vadim's integration layer and is outside
   lomda-production scope.
   ═══════════════════════════════════════════════════════════ */

/* ─── Constants ─────────────────────────────────────────────
   FOUR screens exist so far. goTo() bounds-checks against this, so
   advancing off S3 is currently a silent no-op — that is the honest
   state of a part whose next screen (slide 78, peak sub-item 2) is not
   built yet, NOT a working end-of-part boundary. 720-templates → _global-components.md
   → "Advance gates" warns explicitly that a missing cross-part
   redirect looks like working code; when the remaining screens of this
   part land, bump this. This is the LAST part of the lomda (order 73 is
   the deck's final divider), so the real end of the part — slide 83,
   "המשימה הושלמה!" — is the end of the journey, NOT another
   PartNav.forward(...) crossing. Do not invent a successor part here. */
/* The part is COMPLETE at ten screens: slides 74-83, S0 through S9.
   goTo() bounds-checks against this, so Continue on S9 is a deliberate
   no-op rather than a missing redirect — see NEXT_PART, which records
   that this lomda has no successor part to cross into. */
const TOTAL_SCREENS = 10;  // S0..S9 — slides 74..83, the whole part


/* ─── S0 companion-character assets ─────────────────────────
   Explicit ID → asset mapping. The two clips were INVENTORIED ON DISK
   first and are referenced by their real filenames — the script names
   no media file for this slide, so nothing here is assumed from the
   deck. Files in assets/video/ (only these two, neither carrying a
   doubled ".mp4.mp4" extension):
       Character_1_Climbing.mp4
       Character_2_Climbing.mp4

   LOGICAL MAPPING confirmed by DECODING AND VIEWING frames from each
   clip (t = 0 / 1 / 2 / 3 s, plus 4 and 5 s on the longer clip, plus
   each clip's true final frame), never by reading the filenames or
   guessing from colour — resolving a character by filename is forbidden
   regardless (720-templates → _global-components.md → Companion
   character system). Result matches methodica-math-ratio-03-01's own
   CHARACTER_ASSETS (same lomda, same two companions):
       character-1 = the ORANGE/CORAL round one-eyed character
                     → Character_1_Climbing.mp4  (verified visually)
       character-2 = the TALL YELLOW character with the purple headband
                     → Character_2_Climbing.mp4  (verified visually)
   The filenames and the logical IDs therefore agree; no swap is needed.
   Both clips stage the SAME scene — the character climbing a snow-capped
   rocky peak topped with a red summit flag, wearing a brown backpack and
   trailing a climbing rope — so the pair is a matched set for this
   screen.

   PRESENTER NOTE MATCH — "הדמות נראית מטפסת על הר" ("the character is
   seen climbing a mountain"), from slide74's OWN .rels → notesSlide67.
   Both clips are mountain-climbing clips, which is direct corroboration
   that the note is CURRENT and not stale residue: slide 74 is otherwise
   a byte-identical duplicate of slide 64, and the ONE thing the producer
   rewrote on the copy is exactly this pose line (slide 64's note reads
   "בתחילת מרוץ"). The supplied assets follow the rewrite.

   ───────────────────────────────────────────────────────────
   BACKGROUND CORRECTED IN THE ASSET FILES — 2026-09-14
   ───────────────────────────────────────────────────────────
   The clips were SUPPLIED on a near-white plate (RGB ≈ 253, i.e.
   #FDFDFD), which read as a faint rectangular panel against this
   screen's #FFFFFF. The producer did not approve that, and it may NOT
   be concealed with a CSS filter, blend mode, mask, overlay or white
   panel — so the correction was made in the video files themselves and
   they were replaced under their existing names. Nothing in this part's
   CSS compensates for anything: grep css/style.css for `filter`,
   `mix-blend-mode`, `mask` or `backdrop` and there is none.

   METHOD — the validated YUV-native, edge-connected, smooth-ramp
   correction first used on this lomda in methodica-math-ratio-03-01
   (scratchpad/vid/process_yuv2.js), applied unmodified:
     - operates natively on the decoded yuv420p plane, so there is no
       RGB round-trip;
     - the region is flood-filled from the FRAME BORDER, so anything
       enclosed by the artwork (snow highlights, the eye, teeth, shoes)
       can never be reached;
     - every candidate pixel must also be CHROMA-NEUTRAL (|U-128| and
       |V-128| ≤ 6), so the red flag, the orange/yellow characters and
       the brown pack/rope are excluded on colour alone;
     - the lift is a SMOOTH RAMP, not a binary fill:
           lift = (WHITE_Y - Y) · clamp(1 - (seed - Y)/RAMP, 0, 1)
       so a pixel at plate level goes exactly to white, a pixel RAMP
       units darker is untouched, and everything between tapers — the
       transfer is monotonic and continuous, which is what makes a
       contour/halo impossible rather than merely unlikely;
     - CHROMA IS NEVER WRITTEN. Reported chromaSamplesTouched = 0 on
       both clips.

   RANGE — verified, not assumed: both clips decode as LIMITED range
   (16-235). Checked by reading the same border pixels out of the yuv
   plane and the rgb24 decode — border median Y = 234 against an RGB
   mean of 253, which is the limited-range prediction (16 + 219·253/255
   = 233) and not the full-range one. WHITE_Y = 235 is therefore right;
   at full range it would have had to be 255 and the clip's highlights
   would have clipped.

   PARAMETERS — tolerance 12, ramp 14. Part 03-01 used ramp 10 on its
   own pair; that value was tested here first and REJECTED on evidence,
   not carried over: at ramp 10 the first six frames of clip 2 still
   held ~190,000 plate samples a level short of white (per-frame plate
   mean 254.03 against 255.00 elsewhere), i.e. a faint one-level flash
   over the clip's first quarter-second. Ramp 14 removes it. Ramps 18
   and 24 were also measured; they gain almost nothing further and raise
   the maximum lift to 6-7 units, so 14 is where this pair settles.
   Maximum lift actually applied: 5 luma units on both clips.

   ENCODE — the validated high-fidelity path, h264_nvenc -preset p7
   -tune hq -rc constqp -qp 12, yuv420p, audio stream COPIED (not
   re-encoded). Preservation verified on the shipped files:
       duration     6.016 s / 4.010 s   (video tracks 6.000 / 4.000 s)
       frames       144 / 96            — unchanged
       frame rate   24.000 fps          — unchanged
       resolution   1280×720            — unchanged
       pix fmt      yuv420p, progressive, untagged (as supplied)
       ftyp         isom / isomiso2avc1mp41 — unchanged
       audio        AAC-LC 48 kHz stereo, md5 of the extracted stream
                    IDENTICAL to the original on both clips
   Clip 2's container duration reads 4.010 s against the original's
   4.011 s — a 1 ms difference in how the muxer rounds the AUDIO track
   (4.0313 vs 4.0320 s). The video track is 96 frames / 4.000 s either
   way. Neither clip was trimmed, padded or looped, and the 6.016 vs
   4.011 s difference between the two clips is deliberately left alone.

   ENCODER-ONLY CONTROL — run FIRST, before attributing anything to the
   correction (methodica-math-ratio-03-01 nearly shipped a degraded
   asset for want of one). The UNPROCESSED source was pushed through the
   identical decode→encode path and measured the same way:
       character mean abs delta      clip 1        clip 2
         encoder-only control        0.974         0.960
         corrected                   0.978         0.966
         the correction's own share  0.004         0.006
   Four thousandths and six thousandths of a level: the correction
   changes the characters, mountain, flag, rope and pack by nothing
   measurable. Pixels changed by more than 8: 0.002% / 0.003%.

   RESULT — measured on the DECODED FINAL ENCODES, every frame
   (144 and 96), on a character-free plate strip:
                                          clip 1          clip 2
       plate mean, before              252.402         252.069
       plate mean, after               254.996         254.998
       plate pixels not exactly 255    99.69% → 0.21%  99.54% → 0.13%
       frame-boundary px not 255       99.65% → 3.11%  99.77% → 0.01%
       per-frame plate mean spread      0.749 → 0.008   0.740 → 0.003
   Clip 1's remaining 3.11% boundary residual is entirely on the BOTTOM
   edge (17,800 of 17,894 pixels; left 0, right 8, top 86) and is the
   mountain's own ground shadow, which reaches row 719 of 720 in the
   SUPPLIED ORIGINAL — a soft-threshold scan of the original puts its
   bottom extent at the frame edge. It is artwork, not plate, and is
   left alone. See tests/verify-plate-seam.mjs for the page-composited
   proof that no rectangular panel remains.

   FILE SIZE — a consequence worth knowing: re-encoding at qp 12
   enlarges these clips roughly tenfold (1.75 MB → 18.4 MB, 0.95 MB →
   9.1 MB), because the supplied files were encoded at ~2.2 Mb/s while
   qp 12 preserves the fur, rock and snow grain faithfully. qp 18
   through the same path was measured as an alternative and lands at
   14.1 MB / 7.1 MB — only ~23% smaller, so it does not really solve
   the problem. Flagged for the producer rather than decided here.

   Originals kept at scratchpad/vid/backup/ with their md5s recorded in
   docs/asset-corrections.md.
   ───────────────────────────────────────────────────────────

   TECHNICAL INVENTORY of the supplied originals — probed from the MP4
   box structure (ftyp / mvhd / mdhd / stsd / stts) and cross-checked
   against the decoded frames:
       container        MP4, ftyp isom (isom iso2 avc1 mp41) — both
       video codec      avc1 (H.264), 24-bit, progressive — both
       resolution       1280×720 — both    <- no discrepancy
       frame rate       24.000 fps — both  <- no discrepancy
       audio            mp4a (AAC) 48 kHz stereo 16-bit — both
       duration         Character_1_Climbing  6.016 s / 144 frames
                        Character_2_Climbing  4.011 s /  96 frames

   (!) DURATION DISCREPANCY — REPORTED, NOT "FIXED", and explicitly left
   as-is by producer ruling. The two clips are 2.005 s apart; the -05
   Running pair was matched at 4.011 s each. It changes nothing
   functionally: Continue is ungated and never waits for playback (see
   advanceScreen), the learner only ever sees ONE of the two clips, and
   both end on a meaningful pose. Do not trim or extend either clip to
   match unless the source explicitly requires synchronised duration.

   The audio track is irrelevant here and is not a defect: the element
   is `muted`, which is also what lets autoplay run at all.

   NO `fit` ENTRY — measured, not assumed, and deliberately so.
   TransitionScreen.md → "C" requires a shim to be derived from measured
   ink bounds of THIS clip pair, never by analogy. Measured on decoded
   frames, as fractions of the 720 px frame height. TWO separate
   measurements were taken, because a naive whole-frame ink scan does NOT
   isolate the character on these clips — unlike a cut-out character on a
   plain plate, each frame also contains the mountain prop and the flag:

     (a) whole-scene ink (luma threshold 30/255 — character + mountain
         + flag together):
           clip                  frame 0 top / bottom / height
           Character_1_Climbing   .0250 / .9889 / .9653
           Character_2_Climbing   .0806 / .9236 / .8444
         Stable to the pixel across every sampled frame in both clips,
         because the mountain itself does not move.

     (b) CHARACTER-ONLY ink (chroma > 120 to exclude the grey/white
         mountain and the brown pack/rope, restricted to x < 730 to
         exclude the red summit flag):
           clip                  frame 0 height   final frame height
           Character_1_Climbing      .7556             .7306
           Character_2_Climbing      .7194             .6472

   What that means, and why NO shim is applied:
     1. The mountain prop — the one object drawn identically in both
        clips — is rendered at DIFFERENT SCALE between them: .9653 of
        frame height in clip 1 vs .8444 in clip 2 (and .600 vs .5437 of
        frame width). Clip 2's entire scene is authored ~10-13% smaller
        in frame. That is the producer's own per-clip framing of a
        complete 16:9 scene, not the transparent-margin mismatch of a
        cut-out character that the shim rule exists to correct.
     2. Normalising each character against its OWN scene's mountain
        removes that framing difference and shows the two companions in
        their expected relationship, with the tall character no shorter
        than the round one:
            character-1  .7556 / .9653 = .783   (frame 0)
            character-2  .7194 / .8444 = .852   (frame 0)
            character-1  .7306 / .9653 = .757   (final frame)
            character-2  .6472 / .8444 = .766   (final frame)
     3. Producer ruling, 2026-09-14: preserve the clips' authored
        proportions, add no CSS fit shim, and do not independently scale
        either scene — the scene-scale difference is an asset
        characteristic, not a layout defect to compensate for in code.
     4. Poses also differ within each clip (the round character hugs the
        slope; the tall one steps up it), so ink height across frames is
        not a like-for-like scale comparison in the first place.
   Each character therefore keeps its own natural proportions and its own
   authored framing: `object-fit: contain` on a 16:9 box showing a 16:9
   source is an exact fit, so neither clip is cropped, letterboxed or
   stretched, and neither is scaled against the other.

   FINAL-FRAME HOLD — both clips were decoded at their true last frame
   and both hold a meaningful pose: clip 1 ends with the orange character
   at the summit, one arm raised beside the flag; clip 2 ends with the
   yellow character high on the slope looking up at the flag. Neither
   tail is blank or faded, so NO `freezeAt` is needed on either entry
   (720-templates → _global-components.md → End-state hierarchy step 1:
   let the clip end naturally and stay on its last painted frame).

   The learner only ever sees ONE of the two clips, so there is no
   side-by-side comparison for any of these differences to show up in. */
const S0_CHARACTER = {
  'character-1': {
    video: 'assets/video/Character_1_Climbing.mp4',
    alt: 'الشخصية البرتقالية المستديرة ذات العين الواحدة، تتسلق جبلًا ثلجيًا وعلم في القمة'
  },
  'character-2': {
    video: 'assets/video/Character_2_Climbing.mp4',
    alt: 'الشخصية الصفراء الطويلة ذات عصابة الرأس البنفسجية، تتسلق جبلًا ثلجيًا وعلم في القمة'
  }
};
const CHARACTER_IDS = Object.keys(S0_CHARACTER);


/* ─── S1 image ──────────────────────────────────────────────
   Mounted from here, never written into the markup, so swapping the
   asset is a one-line change in one place (FullImageWithTextOverlay.md
   → "A"). The filename is the one ON DISK, inventoried first; the
   slide's own note names the image as <קונטרול>, which matches
   assets/img/קונטרול.jpg byte-for-byte in intent — but the note is a
   producer label, not a path, so the real file was confirmed by
   listing the folder and opening the image, not by trusting the note.

   ASSET FACTS — the supplied PNG is 1536×1024 (1,847,888 bytes) and
   is the FULL game frame: it carries a star badge, a "1-2-3-trophy"
   progress strip and a gear across the top, and a control console
   with a yellow play button across the bottom. Slide 75's own
   `<a:srcRect t="23801" b="965"/>` crops the top band away. That crop
   is reproduced structurally in css/style.css → .fito-img-frame, not
   with object-position, because `cover` cannot express a srcRect and
   a plain `cover` would put the game chrome back on screen.

   The deck embeds its own downscaled copy of the same artwork
   (media/image76.jpeg, 1282×855, 108 KB) — that is PowerPoint's
   compressed paste, not the production asset. The 1536×1024 PNG on
   disk is what ships.

   The labels A / B / P are NOT in this image — verified by opening
   it — so they are built as live elements. See css/style.css →
   .fito-label.

   ALT TEXT is authored here, not taken from the deck: slide 75
   supplies no alt text and no description beyond the image name. It
   describes what the photo contributes to the copy (the two crossing
   flight paths and their meeting point), since the card refers to
   them by name. */
const S1_IMAGE = {
  src: 'assets/img/קונטרול.jpg',
  alt: 'برج مراقبة في مطار ليلًا. يظهر على شاشة الرادار مسارا طيران يتقاطعان — المسار‎ A ‎باللون الأزرق الفاتح والمسار‎ B ‎باللون الأخضر — ويلتقيان في النقطة‎ P ‎في مركز الشاشة'
};

/* ─── S2 images ─────────────────────────────────────────────
   Three images, all mounted from here rather than written into the
   markup (FullImageWithTextOverlay.md → "A"), and all resolved by
   listing assets/img/ and opening the files — the slide's note gives
   producer LABELS (<קונטרול>, <התראה>, <ארבע הזוויות>), not paths.

   `photo` is the very same file S1 uses. It is referenced again here
   rather than shared through one constant on purpose: the two screens
   show it under different crops, and a future asset swap on one screen
   should not silently move the other.

   ASSET FACTS
     קונטרול.jpg        1536×1024  — the full game frame; see S1_IMAGE
     התראה.png          1536×1024 RGBA, 66.5% opaque — the alert panel
     ארבע הזוויות.png   1536×1024 RGBA, 18.1% opaque — the diagram
   Both overlays are full-canvas artwork with transparent surrounds, and
   their opaque boxes match the deck's own downscaled copies to within
   a few thousandths of the frame, which is what lets them be placed at
   the slide's own pic boxes directly.

   ALT TEXT is authored here; slide 76 supplies none.
     - the panel is DECORATIVE and carries aria-hidden in the markup:
       its warning triangle repeats what the live title already says.
     - the four-angles diagram is the one instructional image on this
       screen, so it gets a real description.

   NOT WIRED — the alert SOUND the note asks for
   ("סאונד של התראה (אבל משהו חביב ולא מטרגר)"). No audio asset exists
   anywhere in this lomda; every part's assets/audio/ is empty. Adding
   an <audio> element pointing at a file that is not there would look
   wired and fail silently, so the requirement is reported instead. Drop
   the file in and add one entry here when it arrives. */
/* The alert sound slide 76 asks for — "סאונד של התראה (אבל משהו חביב
   ולא מטרגר)". Supplied by the producer on 2026-09-14 and the FIRST
   audio asset in this lomda; every part's assets/audio/ was empty
   before it.

   Mounted from here like every other asset on this screen. Playback is
   driven by s2PlayAlert() from resetScreenState(2), which is reached
   only through goTo() — so it fires on a genuine ENTRY to the screen
   and cannot be triggered by a resize, a responsive reflow, or the
   entrance animation being restarted. */
const S2_AUDIO = { id: 's2-alert-audio', src: 'assets/audio/subtle_alert.mp3' };

const S2_IMAGES = {
  photo:  { id: 's2-img',    src: 'assets/img/קונטרול.jpg',
            alt: 'برج مراقبة في مطار ليلًا. يظهر على شاشة الرادار مسارا طيران يتقاطعان' },
  panel:  { id: 's2-panel',  src: 'assets/img/התראה.png',  alt: '' },
  angles: { id: 's2-angles', src: 'assets/img/ארבע הזוויות.png',
            alt: 'مسارا الطيران يتقاطعان ويكوّنان أربع زوايا حول نقطة التقاطع' }
};

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


/* S1 — FullImageWithTextOverlay. No state machine either: the screen
   carries no question, no correctness and no gate, so this is a pure
   mount. Idempotent — the src is only touched when it actually differs,
   so returning to the screen never re-fetches an already-correct
   image. */
function s1MountImage() {
  const img = document.getElementById('s1-img');
  if (!img) return;
  if (img.getAttribute('src') !== S1_IMAGE.src) img.src = S1_IMAGE.src;
  if (img.getAttribute('alt') !== S1_IMAGE.alt) img.alt = S1_IMAGE.alt;
}

/* S2 — FullImageWithTextOverlay, radar alert. No state machine: no
   question, no correctness, no gate. Idempotent, like s1MountImage():
   each src is only touched when it actually differs, so returning to
   the screen never re-fetches. */
function s2MountImages() {
  for (const key of Object.keys(S2_IMAGES)) {
    const a = S2_IMAGES[key];
    const img = document.getElementById(a.id);
    if (!img) continue;
    if (img.getAttribute('src') !== a.src) img.src = a.src;
    if (img.getAttribute('alt') !== a.alt) img.alt = a.alt;
  }
}

/* Mount the alert sound's src once. Separate from playing it, so the
   file is fetched and decoded during the first entry rather than at the
   moment it is needed. */
function s2MountAudio() {
  const a = document.getElementById(S2_AUDIO.id);
  if (!a) return;
  if (a.getAttribute('src') !== S2_AUDIO.src) a.src = S2_AUDIO.src;
}

/* Play the alert once, from the top, as the entrance begins.

   OVERLAP is impossible by construction: there is ONE element, and it is
   paused and rewound before every play, so a fast back-and-forth cannot
   stack two instances. goTo() independently stops and rewinds it on the
   way OUT of the screen (see the media sweep there), so the rewind here
   is belt-and-braces for the case where this runs without a preceding
   exit — the first entry — rather than the only thing preventing an
   overlap.

   REDUCED MOTION DOES NOT MUTE IT. The CSS drops the entrance animation
   under prefers-reduced-motion, but that query is about motion, not
   sound, and this project has no reduced-sensory or audio-suppression
   rule of its own to honour — this is its first audio asset. Muting on
   a motion preference would be inventing a policy. If such a rule is
   ever established, this is the one place it belongs.

   A REJECTED play() IS NOT AN ERROR. Reaching this screen normally means
   clicking Continue, which gives the document sticky activation and lets
   playback start; a direct jump (dev harness, deep link) may have no
   activation and the browser will refuse. That refusal is caught so it
   cannot surface as an unhandled rejection, and it never blocks or
   delays anything — Continue is ungated and does not wait for audio. A
   genuine LOAD failure is a different thing and still surfaces, on the
   element's own error event and in the network log. */
function s2PlayAlert() {
  const a = document.getElementById(S2_AUDIO.id);
  if (!a) return;
  try {
    a.pause();
    a.currentTime = 0;
    const p = a.play();
    if (p && typeof p.catch === 'function') p.catch(function () {});
  } catch (e) { /* autoplay refused — never fatal, never blocking */ }
}

/* Replay the panel's entrance on every ENTRY to the screen, not just on
   first paint. A CSS animation runs once when the element is first
   rendered; coming back from S3 would otherwise show the alert already
   settled, with no rise at all. Removing the class, forcing a reflow
   and re-adding it is the standard restart, and it is confined to this
   screen's own elements. Under prefers-reduced-motion the CSS disables
   the animation entirely, so this is a no-op there. */
function s2ReplayEntrance() {
  const s2 = document.getElementById('s2');
  if (!s2) return;
  s2.querySelectorAll('.s2-scrim, .s2-panel, .s2-alert-text, .s2-angles')
    .forEach(function (el) {
      el.style.animation = 'none';
      void el.offsetWidth;          /* forced reflow — restarts the run */
      el.style.animation = '';
    });
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
  /* Nothing audible or moving may follow the learner off a screen:
     both clips are paused and any success sound is stopped and rewound
     on EVERY navigation, before the new screen is entered. */
  stopAllClips();
  stopSuccessSound();
  /* S3 added the image-zoom modal, so its close-on-navigate call is
     restored here together with the modal's markup, CSS and helpers —
     the complete dependency closure the earlier comment in this spot
     asked for (720-templates → _common-pitfalls.md). The APPLET zoom is
     still absent: no screen in this part embeds an applet, so
     closeAppletZoomOnNavigate() is deliberately not called and its
     helper deliberately not ported. Carrying the call without the
     function would throw on the first navigation. */
  closeImageZoomOnNavigate();
  /* MEDIA ON THE SCREEN BEING LEFT.

     Both kinds are stopped here, in the one place every route already
     passes through — Continue, Back, the arrow keys and any programmatic
     goTo() all land in this function — rather than by hanging a handler
     off each navigation control.

     They are NOT treated the same, and the difference is deliberate:

       video  pause only, never rewound. S0's companion clip has to HOLD
              its final meaningful frame; rewinding it here would reset
              it to the opening pose, which is the exact behaviour
              _global-components.md → "End-state hierarchy" prohibits.

       audio  pause AND rewind to 0. A sound effect that outlives its
              screen would play over the next one — S2's alert carrying
              on over S1 or S3 — so it is stopped immediately and left
              ready to start from the beginning on a genuine re-entry.
              This is also what keeps s2PlayAlert() free of any
              stop-the-previous-instance logic of its own.

     Neither gates or delays anything: both are synchronous property
     writes on elements that are about to be hidden, and the class swap
     below happens in the same frame.

     A CROSS-PART move needs nothing extra — PartNav.back()/forward()
     navigate the document, which tears the media down with the page. */
  const prev = document.querySelector('.screen.active');
  if (prev) {
    prev.querySelectorAll('video').forEach(v => { try { v.pause(); } catch (e) {} });
    prev.querySelectorAll('audio').forEach(a => {
      try { a.pause(); a.currentTime = 0; } catch (e) {}
    });
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
  /* Hidden clips get their source on screen entry (mountLateAssets, C4), BEFORE resetScreenState:
     an enter hook may call load() or seek. */
  document.querySelectorAll('.screen.active [data-src], .screen.active [data-poster]').forEach(function (el) {
    if (el.hasAttribute('data-poster')) { el.setAttribute('poster', el.getAttribute('data-poster')); el.removeAttribute('data-poster'); }
    if (el.hasAttribute('data-src'))    { el.setAttribute('src', el.getAttribute('data-src'));       el.removeAttribute('data-src'); }
  });
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
  /* S1's Back is NOT PartNav-mounted and is never hidden: this screen
     has a real in-part predecessor (S0), so the actual-predecessor
     contract resolves to a plain goTo(currentScreen - 1). Only a part's
     FIRST screen needs the cross-part trail. */
  if (n === 1) {
    s1MountImage();
  }
  /* S2's Back is the same plain in-part control as S1's. The entrance
     is replayed on every entry so returning to this screen shows the
     alert rise again rather than finding it already settled. */
  /* S3 restores rather than resets: a resolved question must still look
     resolved when the learner comes back to it. */
  if (n === 3) {
    s3mcEnter();
  }
  if (n === 6) { s6Enter(); }
  if (n === 7) { s7scEnter(); }
  if (n === 8) { s8Enter(); }
  if (n === 9) { s9Enter(); }
  /* S5 restores rather than resets, for the same reason as S3 and S4:
     a resolved question must still look resolved on return, and the
     scroll gate it already satisfied must not come back. */
  if (n === 5) {
    s5viqEnter();
  }
  /* S4 restores rather than resets, for the same reason as S3. */
  if (n === 4) {
    s4scEnter();
  }
  if (n === 2) {
    s2MountImages();
    s2MountAudio();
    s2ReplayEntrance();
    /* Sound and entrance start together — the note pairs them
       ("סאונד של התראה … והחלונית עולה"). Both are reached only from
       here, so both fire on a real entry and neither can be retriggered
       by a resize or an internal re-render. */
    s2PlayAlert();
  }
}

/* Advance. Neither built screen gates Continue, and both are ungated
   for the same documented reason rather than by omission:
     S0 is a TransitionScreen — Continue is its advance action, not a
        gate, and in particular it does NOT wait for the character clip
        to finish playing.
     S1 is a FullImageWithTextOverlay — a non-question content screen
        with nothing to resolve; slide 75's note authors no gate.
     S2 is the same template and likewise ungated; slide 76's note
        authors an entrance animation and a sound, but no gate, and
        Continue in particular does NOT wait for the panel to finish
        rising.
   Slides 80 and 82 DO author one ("כפתור המשך נעשה זמין רק לאחר שהלומד
   לוחץ על הכפתור הכתום"), so when those screens are built they need a
   real condition here — do not assume the whole part is ungated.
   With TOTAL_SCREENS still 3 the goTo() below is a bounds-checked no-op
   off S2; see the TOTAL_SCREENS note. */
/* ADVANCE GATE. It lives HERE, not on the Continue button, because
   the engine binds ArrowLeft directly to advanceScreen() — a gate
   sitting only in an onclick is silently bypassable by keyboard
   (_global-components.md → "Advance gates").

   This was absent until S4 existed, and harmlessly so: S3 was the last
   built screen, so goTo(4) off it was already a bounds-checked no-op
   and there was nothing to skip INTO. Building S4 turns that gap into
   a real bypass — ArrowLeft would carry the learner past an unanswered
   S3 — so the gate is added with the screen that makes it reachable.
   S3 is otherwise untouched.

   Gate on the RESOLVED flag, never on correctness
   (_global-components.md → "Gate on the resolved state, not on
   correctness"): a learner who spent both attempts and got it wrong
   has finished the question and must be able to leave. {prefix}GateOk
   is the separate, additive scroll-gate condition — it is already true
   whenever no scroll-gated popup is waiting to be read.

   S0/S1/S2 are deliberately absent: a TransitionScreen and two
   FullImageWithTextOverlay screens have nothing to resolve, and none
   of their notes authors a gate. */
function advanceScreen() {
  if (currentScreen === 3 && (!s3mcDone || !s3mcGateOk)) return;
  if (currentScreen === 4 && (!s4scDone || !s4scGateOk)) return;
  if (currentScreen === 5 && (!s5viqDone || !s5viqGateOk)) return;
  /* The clips gate on having actually ENDED, never on correctness —
     there is nothing to get right here, only something to watch. */
  if (currentScreen === 6 && !s6Done) return;
  if (currentScreen === 7 && (!s7scDone || !s7scGateOk)) return;
  if (currentScreen === 8 && !s8Done) return;
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
  if (e.key === 'ArrowLeft')  advanceScreen();
  if (e.key === 'ArrowRight') goBack();
});

/* Escape closes whichever hint overlay is open — never a feedback popup,
   which has no dismissal path at all (_global-components.md → "No manual
   close"). Same handler as the sibling parts. */
document.addEventListener('keydown', e => {
  if (e.key !== 'Escape') return;
  document.querySelectorAll('[id$="-hint-overlay"]').forEach(el => {
    if (!el.classList.contains('hidden')) el.classList.add('hidden');
  });
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
/* 2026-09-24 fix: both release listeners were missing — a plain
   click/release left popupDragId set (pointerdown starts the drag but
   nothing ever cleared it on release), so the popup stayed glued to
   the cursor after the FIRST click, per the approved pattern already
   in 03-01/02/03/04. */
window.addEventListener('pointerup',     () => { popupDragId = null; });
window.addEventListener('pointercancel', () => { popupDragId = null; });


/* ===== ported from part 04 — Progress Question nav renderer ===== */
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

/* ═══════════════════════════════════════════════════════════
   S3 — MultipleChoiceQuestion, text-pill variant (slide 77)
   PEAK QUESTION 1 of 4 — "שאלה 1 מתוך 4", from the slide's own note
   resolved through slide77.xml.rels → notesSlide70.xml (77 ≠ 70), whose
   cached slide-number field reads 77.

   Engine ported from the sibling methodica-math-ratio-03-04's own S6
   (slide 57), the same template and the same text-pill variant, with
   only this slide's content swapped in. The scroll-gate, drag, hint and
   popup mechanics come across unchanged so every question screen in the
   lomda behaves alike.

   DOES NOT use the View My Answer Toggle: this template leaves the
   learner's own selections and the full correct set visible at the same
   time, so the toggle's "one slot, two possible values" applicability
   rule excludes it. The producer confirmed this for S3 explicitly.
   ═══════════════════════════════════════════════════════════ */
/* ANSWER KEY — confirmed two independent ways, never inferred from the
   options' wording:
     1. the authored feedback body itself names them —
        "זיו צודקת כי…" and "שחר צודקת כי…", with no such sentence for
        עומר;
     2. the fresh PowerPoint render of slide 77 shows exactly those two
        checkboxes filled green and עומר's empty.
   עומר's claim is the distractor: he calls α and β vertical angles
   (קודקודיות) when they are in fact adjacent (צמודות).

   maxAttempts 2 is this lomda's own convention, not a guess — all 30
   questions across parts 01-05 use it, together with the retry popup
   below, which no slide authors. Slide 77 authors only the two OUTCOME
   popups, which is exactly the same shape as the sibling MCQ slides. */
const S3MC = { correct: new Set(['ziv', 'shachar']), maxAttempts: 2 };

let s3mcSelected  = new Set();
let s3mcAttempts  = 0;
let s3mcDone      = false;
let s3mcResult    = null;   // 'correct' | 'wrong2'
let s3mcLastWrong = null;   // serialized Set of the immediately-preceding wrong attempt
let s3mcGateOk    = true;   // false only while a scroll-gated popup has not been read to the end
/* The READING REQUIREMENT itself, as opposed to the gate's momentary
   state above. Set once the learner has actually reached the end of a
   scroll-gated feedback body, and cleared only by s3mcResetInitial().
   It is what makes re-reading OPTIONAL when the learner returns to an
   already-resolved screen: the obligation belongs to the outcome, not
   to each individual open of the popup, so it survives the navigation
   round trip that closes and reopens the popup.

   PRODUCER RULING (2026-09-14) — this deliberately DIVERGES from
   _global-components.md, which currently states the opposite twice:
   "If the reopened feedback is scroll-gated, the gate re-arms on
   reopen, like any other open" (Reopening feedback on return to a
   resolved screen), and scroll-gate rule 7 "The gate re-arms only the
   next time the popup is reopened … including a reopen triggered by
   Back-navigation onto a resolved screen". The ruling keeps the gate
   fully intact on the FIRST presentation and lifts it only after the
   learner has already met it. See docs/scroll-gate-restore.md. */
let s3mcGateDone  = false;

/* Stable serialization for the retry gate — sorted so membership, not
   click order, is what is compared. */
function s3mcSig(set) { return [...set].sort().join('|'); }

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
const S3MC_GAP = '';
/* Inline maths gets `dir="ltr"` inside a <bdi> so an expression cannot be
   re-ordered by the surrounding RTL paragraph. No fractions on this
   slide — the donor's frac helpers are deliberately not carried over. */
const X   = '<bdi class="scq-expr" dir="ltr">x</bdi>';
const A   = '<bdi class="scq-expr" dir="ltr">α</bdi>';
const B   = '<bdi class="scq-expr" dir="ltr">β</bdi>';

/* The explanation slide 77 authors, verbatim, paragraph for paragraph.
   Both outcome popups (id=20 "صحيح‎!" and id=22 "זו טעות…") carry the
   SAME body and differ only in their heading — so it is written once.
   The two empty source paragraphs (¶2 and ¶6) are reproduced as gaps,
   which is what they are: spacing, not content. */
const S3MC_EXPLAIN = [
  'مرام على صواب لأن النسبة بين مقدار الزاويتين ' + A + ' و-' + B + ' هي <bdi class="scq-expr" dir="ltr">3 : 2</bdi> .',
  'عندما يكون عامل التوسيع ' + X + ' نحصل على نسبة موسعة جديدة وهي <bdi class="scq-expr" dir="ltr">3x : 2x</bdi>.',   // client: the script had only "3x"
  'ولذلك في هذه الحالة، مجموع الزاويتين هو <bdi class="scq-expr" dir="ltr">2x + 3x</bdi>.',
  S3MC_GAP,
  'ربى على صواب لأن ' + A + ' و-' + B + ' هما زاويتان متجاورتان، ولذلك مجموعهما °180.',
  'أي يمكننا إيجاد ' + X + ' بواسطة:',
  '<bdi class="scq-expr" dir="ltr">2x + 3x = 180°</bdi>'
];

/* `scrollGate` on the two long outcomes — the note asks for it in as
   many words and both source panels are 438×653, taller than the
   canvas. The mechanism unlocks immediately when the body does not
   actually overflow the gated box (rule 6), so it can never strand the
   learner; whether it is genuinely needed here is MEASURED after the
   build and reported (the approved sibling 03-03 removed a vacuous gate
   on its own MCQ screen at the producer's request). No `action` /
   `answerView`: this template has no View My Answer Toggle. */
const S3MC_POPUP = {
  /* NOT authored by slide 77, and correctly so — no sibling MCQ slide
     authors it either. It is this lomda's own retry convention, used
     verbatim by all 30 questions in parts 01-05. */
  retry:   { bg: '#ffdbdc', title: 'الإجابة غير صحيحة.', body: ['هل نحاول مرة أخرى؟'] },
  /* Authored title, verbatim — slide 77 id=20 ¶1. */
  correct: { bg: '#edf8ed', title: 'صحيح‎!', body: S3MC_EXPLAIN, scrollGate: true },
  /* Authored title, verbatim — slide 77 id=22 ¶1 — PLUS the canonical
     second line.

     That line is NOT authored content and its absence from the slide is
     not evidence against it: _question-template-defaults.md owns it as
     system copy for any outcome that reveals the correct answer, and is
     explicit that it is "appended to, never replaces, the authored first
     line — an approved source that supplies its own final-incorrect
     title still gets that title, followed by this line". Slides 43, 52
     and 57 author no second line either, and parts 04 and 05 append it
     nine times between them.

     It applies to this outcome ONLY. The retry popup reveals nothing, so
     it correctly carries no second line, and neither does `correct`.
     The merge is done in s3mcShowPopup() with a real <br> so the two
     lines break rather than wrap, both rendering bold from the title
     element's own weight — the contract's two requirements. */
  wrong2:  { bg: '#ffdbdc', title: 'هذا خطأ، لكن مجرد التعامل مع المسألة يطوّر التفكير‎!',
             titleSecondLine: 'الإجابة الصحيحة تَظهَر الآن.',
             body: S3MC_EXPLAIN, scrollGate: true }
};

function s3mcOpts()  { return [...document.querySelectorAll('#s3 .scq-opt')]; }

/* Toggle — adds/removes ONLY this id. Never clears the other
   selections; this is the one behaviour that must not regress into
   SingleChoiceQuestion's exclusive select. */
function s3mcToggle(id) {
  if (s3mcDone) return;
  if (s3mcSelected.has(id)) s3mcSelected.delete(id);
  else                      s3mcSelected.add(id);
  /* A genuine change clears the previous attempt's per-option marks and
     dismisses its retry popup — approved path 2 of "No manual close". */
  if (s3mcAttempts > 0) {
    document.getElementById('s3mc-popup').classList.add('hidden');
    s3mcOpts().forEach(o => o.classList.remove('correct', 'wrong'));
  }
  s3mcRender();
}

function s3mcRender() {
  s3mcOpts().forEach(o => {
    const on = s3mcSelected.has(o.dataset.id);
    o.classList.toggle('selected', on);
    o.setAttribute('aria-checked', on ? 'true' : 'false');
  });
  const chk = document.getElementById('s3mc-check');
  if (!chk || s3mcDone) return;
  /* Gate on "at least one selected" — never on matching the number of
     correct answers; the learner may legitimately try a wrong count.
     Retry gate on top: the full sorted Set signature, so an off/on
     round-trip back to the same Set does NOT re-enable the button. */
  chk.disabled = s3mcSelected.size < 1 || s3mcSig(s3mcSelected) === s3mcLastWrong;
}

function s3mcCheck() {
  if (s3mcDone) { advanceScreen(); return; }
  if (s3mcSelected.size < 1) return;

  s3mcAttempts++;
  const correct = setsEqual(s3mcSelected, S3MC.correct);
  reportQ(['001', 'q1'], correct, correct || s3mcAttempts >= S3MC.maxAttempts, xapiChoiceAnswer('001', 'q1', '#s3 .scq-opt', [...s3mcSelected]));

  if (correct) {
    /* Fully correct: mark every option in the correct set green. */
    s3mcOpts().forEach(o => {
      o.classList.remove('selected', 'wrong');
      if (S3MC.correct.has(o.dataset.id)) o.classList.add('correct');
    });
    s3mcResult = 'correct';
    s3mcFinish();
    s3mcShowPopup('correct');
    return;
  }

  if (s3mcAttempts < S3MC.maxAttempts) {
    /* Non-final wrong — per-SELECTED-option feedback only. Every option
       the learner actually submitted gets a verdict; a correct option
       they did NOT select stays neutral and unrevealed. */
    s3mcOpts().forEach(o => {
      o.classList.remove('correct', 'wrong');
      if (!s3mcSelected.has(o.dataset.id)) return;
      o.classList.remove('selected');
      o.classList.add(S3MC.correct.has(o.dataset.id) ? 'correct' : 'wrong');
    });
    s3mcLastWrong = s3mcSig(s3mcSelected);
    document.getElementById('s3mc-check').disabled = true;
    S3MC_POPUP.retry.title = [...s3mcSelected].some(id => S3MC.correct.has(id)) ? RETRY_TITLE_PARTIAL : RETRY_TITLE;
    s3mcShowPopup('retry');
    return;
  }

  /* Final wrong — reveal the FULL correct set (including options never
     selected) and mark the learner's incorrect selections red. */
  s3mcOpts().forEach(o => {
    o.classList.remove('selected', 'correct', 'wrong');
    if (S3MC.correct.has(o.dataset.id)) o.classList.add('correct');
    else if (s3mcSelected.has(o.dataset.id)) o.classList.add('wrong');
  });
  s3mcResult = 'wrong2';
  s3mcFinish();
  s3mcShowPopup('wrong2');
}

function s3mcFinish() {
  s3mcDone = true;
  s3mcOpts().forEach(o => { o.disabled = true; });
  const chk = document.getElementById('s3mc-check');
  /* MultipleChoiceQuestion inherits SingleChoiceQuestion's completion
     fallback "هل نتابع؟" — slide 77 authors no progression copy of its
     own either. NOT ValueInputQuestion's "המשך". The button stays
     DISABLED here when the outcome is scroll-gated;
     s3mcCheckScrollGate() is what releases it. */
  if (chk) setNavLabel(chk, 'هل نتابع؟');
  hideHintButton('s3mc-hint');
  /* PEAK QUESTION 1 is resolved by this screen alone — it has no
     sections — so the dot is settled here and now, not deferred.
     Unlike part 04's S6, this screen DOES carry the bar, so the sync
     both updates the shared state and re-renders. */
  syncPeakProgressNav(document.getElementById('s3'));

}

function s3mcResetInitial() {
  s3mcSelected  = new Set();
  s3mcAttempts  = 0;
  s3mcDone      = false;
  s3mcResult    = null;
  s3mcLastWrong = null;
  s3mcGateOk    = true;
  s3mcGateDone  = false;
  s3mcOpts().forEach(o => {
    o.disabled = false;
    o.classList.remove('selected', 'correct', 'wrong');
    o.setAttribute('aria-checked', 'false');
  });
  const popup = document.getElementById('s3mc-popup');
  popup.classList.add('hidden');
  popup.classList.remove('is-scroll', 'needs-scroll');
  document.getElementById('s3mc-hint-overlay').classList.add('hidden');
  const chk = document.getElementById('s3mc-check');
  if (chk) { setNavLabel(chk, 'هل إجابتي صحيحة؟'); chk.disabled = true; }
  showHintButton('s3mc-hint');
}

/* Re-entering a RESOLVED question must bring the whole resolved
   presentation back, not just the marked options —
   _global-components.md → Feedback popup system → "Reopening feedback
   on return to a resolved screen": "A resolved screen never comes back
   without its feedback." MultipleChoiceQuestion.md repeats it for this
   template specifically.

   The state contract is `s3mcDone` ALONE. There is deliberately no
   second "popup was open on exit" flag: with no close button the popup
   can only disappear by navigating away, so such a flag could never be
   false and the spec calls it dead state.

   Nothing here re-scores or replays the submission path — s3mcCheck()
   is not involved. The outcome is read from the stored s3mcResult and
   simply re-rendered. The option markings (correct / wrong) are not
   re-applied because nothing ever removes them; only s3mcResetInitial()
   does, and this branch is the reason it is not called.

   s3mcShowPopup() re-anchors the popup to its default position (it is
   hidden at this point — goTo() sweeps every popup on every call — so
   its resetPopupPosition() branch runs) and re-arms the scroll gate
   like any other open, releasing it immediately via the shared
   atBottom() logic when the restored body does not actually overflow.

   Ported from part 05's s6mcRestoreResolved(), the same template's most
   recently approved implementation. */
function s3mcRestoreResolved() {
  s3mcOpts().forEach(o => { o.disabled = true; });
  const chk = document.getElementById('s3mc-check');
  if (chk) { setNavLabel(chk, 'هل نتابع؟'); chk.disabled = false; }
  hideHintButton('s3mc-hint');
  s3mcShowPopup(s3mcResult === 'correct' ? 'correct' : 'wrong2');
}

function s3mcShowPopup(type) {
  const popup = document.getElementById('s3mc-popup');
  if (!popup) return;
  const cfg      = S3MC_POPUP[type];
  const body     = document.getElementById('s3mc-popup-body');
  const scroller = document.getElementById('s3mc-popup-scroll');
  const chk      = document.getElementById('s3mc-check');

  popup.style.background = cfg.bg;
  if (popup.classList.contains('hidden')) resetPopupPosition(popup);

  /* Terminal-incorrect title carries the canonical fixed second line
     (_question-template-defaults.md) — appended to, never replacing,
     the authored title; only the outcome that reveals the answer. */
  const titleEl = document.getElementById('s3mc-popup-title');
  titleEl.innerHTML = cfg.titleSecondLine
    ? (cfg.title + '<br>' + cfg.titleSecondLine)
    : cfg.title;

  body.innerHTML = cfg.body.map(p => {
    if (p === '') return '<p class="scq-gap" aria-hidden="true"></p>';
    return '<p>' + p + '</p>';
  }).join('');
  /* The donor calls renderPopupAction(id, null, null) here to clear an
     action row. This screen has NO View My Answer Toggle — the producer
     confirmed it, and the template leaves the learner's selections and
     the correct set visible at the same time — so no action row is ever
     created and the helper is deliberately not ported. Calling it
     anyway threw a ReferenceError right before the popup was un-hidden,
     which showed up as "the popup never appears" with its title already
     correctly set. Caught by tests/verify-s3.mjs. */

  /* Scroll gate — arm before showing, then reset and measure AFTER the
     popup is visible (a display:none element ignores a scrollTop write
     and reports 0 for scrollHeight/clientHeight). */
  /* `is-scroll` is applied for every scroll-gated outcome, gated or
     not: it is what makes the body scrollable at all, and re-reading
     has to stay POSSIBLE on return even once it is no longer
     REQUIRED. Only the gate itself is conditional. */
  popup.classList.toggle('is-scroll', !!cfg.scrollGate);
  const gateApplies = !!cfg.scrollGate && !s3mcGateDone;
  s3mcGateOk = !gateApplies;
  if (chk && !s3mcGateOk) chk.disabled = true;
  if (cfg.scrollGate) scroller.setAttribute('tabindex', '0');
  else                scroller.removeAttribute('tabindex');

  popup.classList.remove('hidden');
  clampPopup(popup);
  /* Rule 4 — reset scroll on every open — still holds unconditionally.
     The popup reopens at the top of its content; what changed is only
     that reading on from there is now optional. */
  scroller.scrollTop = 0;
  if (gateApplies) {
    placeScrollHint('s3mc-popup');
    s3mcCheckScrollGate();     // unlocks immediately when the content does not actually overflow
  } else {
    /* Either the outcome was never gated, or the learner has already
       met the requirement and is coming back to a resolved screen. No
       hint bubble either way — the learner is never nagged to re-read
       something they have already read. */
    popup.classList.remove('needs-scroll');
    if (chk && s3mcDone) chk.disabled = false;
  }
}

function s3mcCheckScrollGate() {
  if (s3mcGateOk) return;
  const scroller = document.getElementById('s3mc-popup-scroll');
  const popup    = document.getElementById('s3mc-popup');
  if (!scroller || !popup) return;
  if (!atBottom(scroller)) { popup.classList.add('needs-scroll'); return; }
  s3mcGateOk = true;
  /* Reaching the end is what satisfies the requirement — whether the
     learner scrolled there or the content never overflowed in the
     first place (rule 6's immediate unlock). Either way they have been
     shown the whole body, so returning later must not ask again. */
  s3mcGateDone = true;
  const chk = document.getElementById('s3mc-check');
  if (chk && s3mcDone) chk.disabled = false;
  popup.classList.remove('needs-scroll');
}

let s3mcResizeRaf = null;
window.addEventListener('resize', () => {
  const popup = document.getElementById('s3mc-popup');
  if (!popup || popup.classList.contains('hidden')) return;
  placeScrollHint('s3mc-popup');
  if (s3mcGateOk) return;
  if (s3mcResizeRaf) cancelAnimationFrame(s3mcResizeRaf);
  s3mcResizeRaf = requestAnimationFrame(s3mcCheckScrollGate);
});

function s3mcOpenHint()  { xapiRequestedHint('001', 'q1'); document.getElementById('s3mc-hint-overlay').classList.remove('hidden'); }
function s3mcCloseHint() { document.getElementById('s3mc-hint-overlay').classList.add('hidden'); }
function s3mcCloseHintOnBackdrop(e) { if (e.target === e.currentTarget) s3mcCloseHint(); }

/* ═══════════════════════════════════════════════════════════════
   S4 — SingleChoiceQuestion (slide 78). PEAK QUESTION 2 of 4.

   SOURCE EVIDENCE — slide 78, resolved through <p:sldIdLst> rId82 ->
   slide78.xml, notes through the slide's OWN rels -> notesSlide71.xml
   (NOT notesSlide78; part-file numbering does not track slide
   numbering). Deck md5 0eb1ccb8, the same build all of S0-S3 came
   from. Presenter notes: Figma screen "SingleChoiceQuestionImage",
   "שאלה 2 מתוך 4", image "<קונטרול שאלה 1>" — the SAME control image
   as question 1, confirmed independently by slides 77 and 78 both
   referencing media/image79.png.

   ONE CHOICE, NOT MANY. This is the single-select sibling of S3's
   MultipleChoiceQuestion: the selection is one id, not a Set, the
   options carry role="radio" inside a role="radiogroup", and choosing
   an option replaces the previous choice rather than adding to it.
   Everything downstream of the selection model — attempts, retry
   popup, reveal, scroll gate, hint, peak dot — is the same contract
   S3 already carries.

   CORRECT ANSWER: 'ziv'. Established from the slide's own pre-filled
   radio fill (#00B050 green on זיו, #00B0F0 on the other three), and
   corroborated by the authored feedback body, which concludes that
   α is 1.5x β — exactly Ziv's claim.

   RATIO TERM ORDER — PRODUCER RULING 2026-09-14. The learner must SEE,
   left to right, `3 : 2` and `1.5 : 1`; never `2 : 3` or `1 : 1.5`.
   Success is judged on the final visible browser output, not on the
   DOM string or the order the runs appear in the PPTX XML. Inside
   <bdi dir="ltr"> the DOM order IS the visible order (measured: the
   glyphs of `1.5 : 1` land at 1@896 .@905.5 5@910.6 :@925.2 1@934.6),
   so the expanded ratio is authored here as '1.5 : 1' — the one place
   the build's string differs from the XML's run order, because the XML
   carries it as "1 : 1.5" and that would render the wrong way round.
   tests/measure-ratio-order.mjs asserts all of this. */
const S4SC = { correct: 'ziv', maxAttempts: 2 };

/* Authored feedback body — slide 78 shapes id=20 (correct) and id=22
   (terminal wrong), which carry an IDENTICAL body; only the titles
   differ (_question-template-defaults.md → "Feedback: shared
   explanation"). Blank strings are the authored blank paragraphs. */
const S4SC_EXPLAIN = [
  'إذا كانت النسبة بين الزاوية<bdi class="scq-expr" dir="ltr">α</bdi> والزاوية <bdi class="scq-expr" dir="ltr">β</bdi> هي <bdi class="scq-expr" dir="ltr">3 : 2</bdi>،',
  'يمكننا اختزالها بـ <bdi class="scq-expr" dir="ltr">2</bdi> والحصول على النسبة <bdi class="scq-expr" dir="ltr">1.5 : 1</bdi>.',
  '',
  'وهكذا يمكن أن نرى أن الزاوية <bdi class="scq-expr" dir="ltr">α</bdi> أكبر من الزاوية <bdi class="scq-expr" dir="ltr">β</bdi> بـ  <bdi class="scq-expr" dir="ltr">1.5</bdi> مرة.'
];

const S4SC_POPUP = {
  /* The retry popup is NOT authored by any slide — it is the
     established system copy every question screen in parts 01-05
     carries, and it reveals nothing, so it takes no second title line
     and no scroll gate. */
  retry:   { bg: '#ffdbdc', title: 'الإجابة غير صحيحة.', body: ['هل نحاول مرة أخرى؟'] },

  /* NO `scrollGate` on either outcome — a deliberate, measured decision,
     not an omission. The Long-feedback variant is explicitly "not the
     default for all feedback" and its own "When not to use" clause
     covers exactly this case: standard-length feedback keeps the
     default auto-height popup, with no internal scrollbar and no gate.

     Measured: this body renders 136px (correct) / 188px (terminal
     wrong) against the variant's fixed 474px footprint, so the gate
     never engaged and the popup carried 290px / 238px of blank space
     below the final line. The body cannot overflow at any supported
     viewport either — the canvas is a fixed design-pixel space that
     scales as a whole, and the only canvas that differs (1280x900) is
     TALLER, not shorter.

     Dropping the flag removes `is-scroll`, and with it the 474px fixed
     height, restoring the base .scq-popup's canonical height:auto,
     15px padding and 4px title/body gap. The gate MACHINERY is left
     fully intact — s4scCheckScrollGate(), s4scGateDone, the resize
     handler and the advance-gate condition all still work — so an
     outcome that genuinely overflows only has to set the flag. */
  correct: { bg: '#edf8ed', title: 'صحيح‎!', body: S4SC_EXPLAIN },
  /* Authored title, verbatim — slide 78 id=22 ¶1 — PLUS the canonical
     second line that _question-template-defaults.md appends to any
     outcome revealing the correct answer. Same treatment as S3. */
  wrong2:  { bg: '#ffdbdc', title: 'هذا خطأ، لكن مجرد التعامل مع المسألة يطوّر التفكير‎!',
             titleSecondLine: 'الإجابة الصحيحة تَظهَر الآن.',
             body: S4SC_EXPLAIN }
};

let s4scSelected  = null;   // ONE id, or null — the single-select difference
let s4scAttempts  = 0;
let s4scDone      = false;
let s4scResult    = null;   // 'correct' | 'wrong2'
let s4scLastWrong = null;   // the id of the immediately-preceding wrong attempt
let s4scGateOk    = true;
let s4scGateDone  = false;  // the reading requirement — see docs/scroll-gate-restore.md

function s4scOpt(id) { return document.querySelector('#s4 .scq-opt[data-id="' + id + '"]'); }
function s4scOpts()  { return [...document.querySelectorAll('#s4 .scq-opt')]; }

function s4scSelect(id) {
  if (s4scDone) return;
  s4scSelected = id;          // replaces, never accumulates
  s4scRender();
  /* Established retry-popup dismissal: once an attempt has been made,
     touching an option puts the learner back in correcting mode. */
  if (s4scAttempts > 0) {
    document.getElementById('s4sc-popup').classList.add('hidden');
    s4scOpts().forEach(o => o.classList.remove('correct', 'wrong'));
  }
}

function s4scRender() {
  s4scOpts().forEach(o => {
    const on = o.dataset.id === s4scSelected;
    o.classList.toggle('selected', on);
    o.setAttribute('aria-checked', on ? 'true' : 'false');
  });
  const chk = document.getElementById('s4sc-check');
  if (!chk || s4scDone) return;
  /* Gate on "something is chosen", plus the retry gate: re-submitting
     the identical wrong choice stays blocked until it changes. */
  chk.disabled = s4scSelected === null || s4scSelected === s4scLastWrong;
}

function s4scCheck() {
  if (s4scDone) { advanceScreen(); return; }
  if (s4scSelected === null) return;
  s4scAttempts++;
  reportQ(['001', 'q2'], s4scSelected === S4SC.correct, s4scSelected === S4SC.correct || s4scAttempts >= S4SC.maxAttempts, xapiChoiceAnswer('001', 'q2', '#s4 .scq-opt', s4scSelected));

  if (s4scSelected === S4SC.correct) {
    s4scResult = 'correct';
    s4scOpt(S4SC.correct).classList.add('correct');
    s4scFinish();
    s4scShowPopup('correct');
    return;
  }

  if (s4scAttempts < S4SC.maxAttempts) {
    /* FIRST WRONG — retry. The learner's own pick is marked wrong (red +
       ✗), as on every other single-choice screen in this lomda
       (SingleChoiceQuestion.md → state 4, Try Again); the CORRECT option
       is not revealed yet. Corrected 2026-09-23 on the producer's report
       — this screen and S4 shipped with no marking at all on the first
       attempt, unlike the template. */
    s4scOpt(s4scSelected).classList.add('wrong');
    s4scLastWrong = s4scSelected;
    const chk = document.getElementById('s4sc-check');
    if (chk) chk.disabled = true;
    s4scShowPopup('retry');
    return;
  }

  /* TERMINAL WRONG — the correct option is revealed alongside the
     learner's own wrong pick, both visible at once, which is why this
     screen needs no View My Answer Toggle. */
  s4scResult = 'wrong2';
  s4scOpt(s4scSelected).classList.add('wrong');
  s4scOpt(S4SC.correct).classList.add('correct');
  s4scFinish();
  s4scShowPopup('wrong2');
}

function s4scFinish() {
  s4scDone = true;
  s4scOpts().forEach(o => { o.disabled = true; });
  const chk = document.getElementById('s4sc-check');
  /* SingleChoiceQuestion's completion fallback — slide 78 authors no
     progression copy of its own, same as slide 77. The button stays
     disabled while a scroll-gated outcome is unread; the gate releases
     it. */
  if (chk) setNavLabel(chk, 'هل نتابع؟');
  hideHintButton('s4sc-hint');
  /* PEAK QUESTION 2 resolves on this screen alone — no sections — so
     the dot settles here, immediately. */
  syncPeakProgressNav(document.getElementById('s4'));
}

function s4scResetInitial() {
  s4scSelected  = null;
  s4scAttempts  = 0;
  s4scDone      = false;
  s4scResult    = null;
  s4scLastWrong = null;
  s4scGateOk    = true;
  s4scGateDone  = false;
  s4scOpts().forEach(o => {
    o.disabled = false;
    o.classList.remove('selected', 'correct', 'wrong');
    o.setAttribute('aria-checked', 'false');
  });
  const popup = document.getElementById('s4sc-popup');
  popup.classList.add('hidden');
  popup.classList.remove('is-scroll', 'needs-scroll');
  document.getElementById('s4sc-hint-overlay').classList.add('hidden');
  const chk = document.getElementById('s4sc-check');
  if (chk) { setNavLabel(chk, 'هل إجابتي صحيحة؟'); chk.disabled = true; }
  showHintButton('s4sc-hint');
}

/* A resolved screen never comes back without its feedback — see
   docs/scroll-gate-restore.md and _global-components.md → "Reopening
   feedback on return to a resolved screen". Nothing is re-scored:
   s4scCheck() is not involved. */
function s4scRestoreResolved() {
  s4scOpts().forEach(o => { o.disabled = true; });
  const chk = document.getElementById('s4sc-check');
  if (chk) { setNavLabel(chk, 'هل نتابع؟'); chk.disabled = false; }
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

  const titleEl = document.getElementById('s4sc-popup-title');
  titleEl.innerHTML = cfg.titleSecondLine
    ? (cfg.title + '<br>' + cfg.titleSecondLine)
    : cfg.title;

  body.innerHTML = cfg.body.map(p => {
    if (p === '') return '<p class="scq-gap" aria-hidden="true"></p>';
    return '<p>' + p + '</p>';
  }).join('');

  popup.classList.toggle('is-scroll', !!cfg.scrollGate);
  const gateApplies = !!cfg.scrollGate && !s4scGateDone;
  s4scGateOk = !gateApplies;
  if (chk && !s4scGateOk) chk.disabled = true;
  if (cfg.scrollGate) scroller.setAttribute('tabindex', '0');
  else                scroller.removeAttribute('tabindex');

  popup.classList.remove('hidden');
  clampPopup(popup);
  scroller.scrollTop = 0;
  if (gateApplies) {
    placeScrollHint('s4sc-popup');
    s4scCheckScrollGate();
  } else {
    popup.classList.remove('needs-scroll');
    if (chk && s4scDone) chk.disabled = false;
  }
}

function s4scCheckScrollGate() {
  if (s4scGateOk) return;
  const scroller = document.getElementById('s4sc-popup-scroll');
  const popup    = document.getElementById('s4sc-popup');
  if (!scroller || !popup) return;
  if (!atBottom(scroller)) { popup.classList.add('needs-scroll'); return; }
  s4scGateOk = true;
  s4scGateDone = true;
  const chk = document.getElementById('s4sc-check');
  if (chk && s4scDone) chk.disabled = false;
  popup.classList.remove('needs-scroll');
}

let s4scResizeRaf = null;
window.addEventListener('resize', () => {
  const popup = document.getElementById('s4sc-popup');
  if (!popup || popup.classList.contains('hidden')) return;
  placeScrollHint('s4sc-popup');
  if (s4scGateOk) return;
  if (s4scResizeRaf) cancelAnimationFrame(s4scResizeRaf);
  s4scResizeRaf = requestAnimationFrame(s4scCheckScrollGate);
});

function s4scOpenHint()  { xapiRequestedHint('001', 'q2'); document.getElementById('s4sc-hint-overlay').classList.remove('hidden'); }
function s4scCloseHint() { document.getElementById('s4sc-hint-overlay').classList.add('hidden'); }
function s4scCloseHintOnBackdrop(e) { if (e.target === e.currentTarget) s4scCloseHint(); }

function s4scEnter() {
  const sec = document.getElementById('s4');
  if (!s4scDone && screenPristine(4)) s4scResetInitial();
  s4scRender();
  if (s4scDone) s4scRestoreResolved();
  syncPeakProgressNav(sec);
  markPeakCurrent(sec, 2);
}


/* ═══════════════════════════════════════════════════════════
   SCREEN 05 — ValueInputQuestion (slide 79). PEAK QUESTION 3 of 4.

   Two slots judged as ONE answer, two attempts, then a terminal reveal
   — the established multi-field shape this lomda already ships in
   methodica-math-ratio-03-03 (its S7 and S10), ported rather than
   reinvented. The behaviours that shape guarantees and this screen
   inherits verbatim:

     - each box still shows its OWN verdict on the retry attempt, so the
       learner knows which half to fix, even though the score is one;
     - editing after a retry clears both marks and dismisses only the
       retry popup — nothing else;
     - FOCUS, caret movement and selection are not edits: only the
       `input` event clears marks, so tabbing through or clicking into a
       field after a verdict leaves the marks alone;
     - programmatic restoration and the solution toggle write values
       through s5viqMark()/applyView(), never through the input handler,
       so they cannot be mistaken for learner edits;
     - the displayed value and the verdict icon are written in the same
       place for every state, so they cannot disagree;
     - re-submitting the identical wrong pair is refused by the
       signature check rather than burning the second attempt.

   The DEGREE MARK is chrome (see css/style.css → .s5viq-unit): the
   learner types 108, the field shows 108°, and the compared value is
   the digits alone.
   ═══════════════════════════════════════════════════════════ */
const S5VIQ = { correct: { alpha: '108', beta: '72' }, maxAttempts: 2, slots: ['alpha', 'beta'] };

let s5viqAttempts  = 0;
let s5viqDone      = false;
let s5viqResult    = null;   // 'correct' | 'wrong2'
let s5viqLastWrong = null;   // signature of the immediately preceding wrong attempt
let s5viqAnswer    = null;   // per-slot snapshot for the toggle; null = nothing to toggle
let s5viqView      = 'solution';
/* Same two-flag gate model every question screen in this part uses:
   GateOk is the momentary state, GateDone is the reading requirement
   itself and survives a reopen (docs/scroll-gate-restore.md). */
let s5viqGateOk    = true;
let s5viqGateDone  = false;
let s5viqResizeRaf = null;

function s5viqInput(k) { return document.getElementById('s5viq-input-' + k); }
function s5viqField(k) { return document.getElementById('s5viq-field-' + k); }
function s5viqIco(k)   { return document.getElementById('s5viq-ico-' + k); }

/* Comparison form of a typed value (ValueInputQuestion.md → "Numeric input
   format"): grouping commas stripped, then the canonical number, so 072 and
   72 — or 080 and 80 — are the SAME answer. Used for the correctness check
   and the Retry-gate signature only; the raw typed text is what the View
   My Answer Toggle snapshots and shows back. */
function s5viqNormalize(raw) {
  const t = String(raw == null ? '' : raw).trim().replace(/,/g, '');
  const n = Number(t);
  return (t !== '' && isFinite(n)) ? String(n) : t;
}
function s5viqSig() {
  return JSON.stringify(S5VIQ.slots.map(k => (s5viqInput(k) ? s5viqNormalize(s5viqInput(k).value) : '')));
}
function s5viqFilled() {
  return S5VIQ.slots.every(k => s5viqInput(k) && s5viqInput(k).value.trim() !== '');
}

/* The ONLY learner-edit entry point. Bound to `input`, never to focus,
   click, keydown or selectionchange — which is what makes caret moves
   and selections non-destructive. */
/* type="text" (inputmode="numeric" only hints the on-screen keyboard, it
   does not block a physical keyboard or paste) so every keystroke AND
   paste is filtered here — same shape as every other numeric ValueInput
   screen in the lomda (_question-template-defaults.md → "Numeric input
   format"). Fix 2026-09-24: this screen had no filter at all, so a
   non-digit character (e.g. a Hebrew letter) could sit in the field. */
function s5viqOnInput() {
  S5VIQ.slots.forEach(k => {
    const el = s5viqInput(k);
    const cleaned = el.value.replace(/[^0-9,.-]/g, '');
    if (cleaned !== el.value) el.value = cleaned;
  });
  if (s5viqDone) return;
  if (s5viqAttempts > 0) {
    document.getElementById('s5viq-popup').classList.add('hidden');
    S5VIQ.slots.forEach(k => s5viqMark(k, null));
  }
  s5viqUpdateBar();
}

function s5viqUpdateBar() {
  const chk = document.getElementById('s5viq-check');
  if (!chk || s5viqDone) return;
  chk.disabled = !s5viqFilled() || s5viqSig() === s5viqLastWrong;
}

function s5viqMark(k, state) {
  const f = s5viqField(k), ico = s5viqIco(k), inp = s5viqInput(k);
  if (!f) return;
  f.classList.remove('viq-correct', 'viq-wrong');
  if (state === 'correct' || state === 'solution') f.classList.add('viq-correct');
  else if (state === 'wrong') f.classList.add('viq-wrong');
  if (ico) ico.innerHTML = !state ? '' : ((state === 'wrong') ? S5VIQ_ICON_WRONG : S5VIQ_ICON_CORRECT);
  /* the solution view writes the correct value into the same box the
     verdict is on, so value and icon can never disagree */
  if (state === 'solution' && inp) inp.value = S5VIQ.correct[k];
}

function s5viqLock() {
  S5VIQ.slots.forEach(k => {
    const i = s5viqInput(k);
    if (i) { i.readOnly = true; i.setAttribute('aria-readonly', 'true'); }
  });
}

function s5viqCheck() {
  if (s5viqDone) { advanceScreen(); return; }
  if (!s5viqFilled()) return;
  s5viqAttempts++;

  const got = {};
  S5VIQ.slots.forEach(k => { got[k] = s5viqInput(k).value.trim(); });
  const perSlot = S5VIQ.slots.map(k => ({ k, ok: s5viqNormalize(got[k]) === s5viqNormalize(S5VIQ.correct[k]) }));
  const allOk = perSlot.every(s => s.ok);
  /* xAPI (2026-11-11): the layer writes Kata's form from the metadata — performance, 'α[.]108[,]β[.]72' (proposed to Kata, `other` there) */
  reportQ(['001', 'q3'], allOk, allOk || s5viqAttempts >= S5VIQ.maxAttempts,
    { steps: { 'α': got.alpha, 'β': got.beta } });

  if (allOk) {
    S5VIQ.slots.forEach(k => s5viqMark(k, 'correct'));
    s5viqLock();
    s5viqAnswer = null;
    s5viqFinish('correct');
    s5viqShowPopup('correct');
    return;
  }

  if (s5viqAttempts < S5VIQ.maxAttempts) {
    perSlot.forEach(s => s5viqMark(s.k, s.ok ? 'correct' : 'wrong'));
    s5viqLastWrong = s5viqSig();
    document.getElementById('s5viq-check').disabled = true;
    S5VIQ_POPUP.retry.title = perSlot.some(s => s.ok) ? RETRY_TITLE_PARTIAL : RETRY_TITLE;
    s5viqShowPopup('retry');
    return;
  }

  /* Terminal wrong — snapshot the learner's own values BEFORE the
     solution overwrites the very inputs they live in. */
  s5viqAnswer = {};
  perSlot.forEach(s => { s5viqAnswer[s.k] = { value: got[s.k], correct: s.ok }; });
  s5viqView = 'solution';
  s5viqLock();
  s5viqFinish('wrong2');
  s5viqApplyView();
}

function s5viqApplyView() {
  if (s5viqAnswer === null) return;
  S5VIQ.slots.forEach(k => {
    const a = s5viqAnswer[k];
    if (a.correct) { s5viqMark(k, 'correct'); s5viqInput(k).value = a.value; return; }
    if (s5viqView === 'answer') { s5viqInput(k).value = a.value; s5viqMark(k, 'wrong'); }
    else s5viqMark(k, 'solution');
  });
  s5viqShowPopup('wrong2', s5viqView);
}

function s5viqToggleView() {
  if (s5viqAnswer === null) return;
  s5viqView = (s5viqView === 'solution') ? 'answer' : 'solution';
  s5viqApplyView();
}

function s5viqFinish(result) {
  s5viqDone = true;
  s5viqResult = result;
  const chk = document.getElementById('s5viq-check');
  if (chk) { setNavLabel(chk, 'هل نتابع؟'); chk.disabled = false; }
  hideHintButton('s5viq-hint');
  syncPeakProgressNav(document.getElementById('s5'));
}

function s5viqResetInitial() {
  s5viqAttempts = 0;
  s5viqDone = false;
  s5viqResult = null;
  s5viqLastWrong = null;
  s5viqAnswer = null;
  s5viqView = 'solution';
  s5viqGateOk = true;
  s5viqGateDone = false;
  S5VIQ.slots.forEach(k => {
    const i = s5viqInput(k);
    if (i) { i.value = ''; i.readOnly = false; i.removeAttribute('aria-readonly'); }
    s5viqMark(k, null);
  });
  const pop = document.getElementById('s5viq-popup');
  if (pop) pop.classList.add('hidden');
  const ho = document.getElementById('s5viq-hint-overlay');
  if (ho) ho.classList.add('hidden');
  const chk = document.getElementById('s5viq-check');
  if (chk) { setNavLabel(chk, 'هل إجابتي صحيحة؟'); chk.disabled = true; }
  showHintButton('s5viq-hint');
}

/* Shared explanation — the authored body of BOTH feedback panels
   (slide 79 id=26 and id=28), identical run for run, differing only in
   the title. Every mathematical fragment is an isolated LTR island so
   the visible order is the authored order. */
const S5VIQ_EXPLAIN = [
  'لإيجاد <bdi class="scq-expr" dir="ltr">α</bdi> و-<bdi class="scq-expr" dir="ltr">β</bdi> نحل المعادلة‎ التالية:',
  '<bdi class="scq-expr" dir="ltr">2x + 3x = 180°</bdi>',
  '<bdi class="scq-expr" dir="ltr">5x = 180°</bdi>',
  '<bdi class="scq-expr" dir="ltr">x = 36°</bdi>',
  '',
  'أي إن عامل التوسيع هو <bdi class="scq-expr" dir="ltr">36</bdi>، لذلك:',
  '<bdi class="scq-expr" dir="ltr">α = 3 ⋅ 36 = 108°</bdi>',
  '<bdi class="scq-expr" dir="ltr">β = 2 ⋅ 36 = 72°</bdi>'
];

const S5VIQ_POPUP = {
  retry:   { bg: '#ffdbdc', title: 'الإجابة غير صحيحة.', body: ['هل نحاول مرة أخرى؟'] },
  correct: { bg: '#edf8ed', title: 'هذا صحيح‎!', body: S5VIQ_EXPLAIN },
  /* Authored title, verbatim — slide 79 id=28 authors TWO bold lines —
     PLUS the canonical second line _question-template-defaults.md
     appends to any outcome that reveals the correct answer. This one
     does: the terminal view writes 108 and 72 into the fields. */
  wrong2:  { bg: '#ffdbdc', title: 'غير دقيق.<br>لا بأس، دعونا نشرح‎:',
             titleSecondLine: 'الإجابة الصحيحة تَظهَر الآن.',
             body: S5VIQ_EXPLAIN,
             action: { label: 'الانتقال إلى إجابتي', view: 'answer' },
             answerView: {
               title: 'الإجابة غير صحيحة.',
               titleSecondLine: null,
               body: ['يمكن مشاهدة الحل الصحيح مرة أخرى‎.'],
               action: { label: 'الإجابة الصحيحة', view: 'solution' }
             } }
};

function s5viqOpenHint()  { xapiRequestedHint('001', 'q3'); document.getElementById('s5viq-hint-overlay').classList.remove('hidden'); }
function s5viqCloseHint() { document.getElementById('s5viq-hint-overlay').classList.add('hidden'); }
function s5viqCloseHintOnBackdrop(e) { if (e.target === e.currentTarget) s5viqCloseHint(); }

/* Restoration — the same contract S3 and S4 ship: reopen the resolved
   feedback, keep the revealed values and the learner's marks, keep the
   fields locked and the hint withdrawn, keep the completion label, and
   re-score nothing. The gate is NOT re-armed: s5viqGateDone survives. */
function s5viqRestoreResolved() {
  s5viqLock();
  const chk = document.getElementById('s5viq-check');
  if (chk) { setNavLabel(chk, 'هل نتابع؟'); chk.disabled = false; }
  hideHintButton('s5viq-hint');
  if (s5viqResult === 'correct') s5viqShowPopup('correct');
  else s5viqApplyView();
}

function s5viqEnter() {
  const sec = document.getElementById('s5');
  if (!s5viqDone && screenPristine(5)) s5viqResetInitial();
  if (s5viqDone) s5viqRestoreResolved(); else s5viqUpdateBar();
  syncPeakProgressNav(sec);
  markPeakCurrent(sec, 3);
}


/* Feedback popup action button (View My Answer Toggle). `action` falsy
   → the row is removed, which is why the retry and correct outcomes
   never grow one. Ported verbatim from methodica-math-ratio-03-03. */
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

/* Verdict icons for a value field — Feedback/Correct #609E12 and
   Feedback/Incorrect #B20010, the same two this lomda uses everywhere. */
const S5VIQ_ICON_CORRECT = '<svg width="16" height="16" viewBox="0 0 14 12" fill="none" aria-hidden="true"><path d="M4.52036 12C4.12296 12.0002 3.74183 11.8126 3.46106 11.4787L0.258452 7.67744C-0.0861507 7.26814 -0.0861507 6.60467 0.258452 6.19536C0.603165 5.78619 1.16194 5.78619 1.50666 6.19536L4.52036 9.77374L12.4933 0.306878C12.8381 -0.102293 13.3968 -0.102293 13.7415 0.306878C14.0862 0.716179 14.0862 1.37966 13.7415 1.78896L5.57966 11.4787C5.29889 11.8126 4.91776 12.0002 4.52036 12Z" fill="#609E12"/></svg>';
const S5VIQ_ICON_WRONG   = '<svg width="16" height="16" viewBox="0 0 14 14" fill="none" aria-hidden="true"><path d="M13.6743 0.325691C13.4657 0.117151 13.1828 0 12.8879 0C12.5929 0 12.31 0.117151 12.1014 0.325691L7 5.42709L1.8986 0.325691C1.69 0.117151 1.40711 0 1.11215 0C0.817183 0 0.534294 0.117151 0.325691 0.325691C0.117151 0.534294 0 0.817183 0 1.11215C0 1.40711 0.117151 1.69 0.325691 1.8986L5.42709 7L0.325691 12.1014C0.117151 12.31 0 12.5929 0 12.8879C0 13.1828 0.117151 13.4657 0.325691 13.6743C0.534294 13.8829 0.817183 14 1.11215 14C1.40711 14 1.69 13.8829 1.8986 13.6743L7 8.57291L12.1014 13.6743C12.31 13.8829 12.5929 14 12.8879 14C13.1828 14 13.4657 13.8829 13.6743 13.6743C13.8829 13.4657 14 13.1828 14 12.8879C14 12.5929 13.8829 12.31 13.6743 12.1014L8.57291 7L13.6743 1.8986C13.8829 1.69 14 1.40711 14 1.11215C14 0.817183 13.8829 0.534294 13.6743 0.325691Z" fill="#B20010"/></svg>';

/* Popup renderer — the SAME shape s3mcShowPopup/s4scShowPopup use, so
   all three question screens in this part share one gate model. The
   only addition is the answerView merge and the action row that the
   ValueInputQuestion terminal reveal carries. */
function s5viqShowPopup(type, view) {
  const popup = document.getElementById('s5viq-popup');
  if (!popup) return;
  const base = S5VIQ_POPUP[type];
  const cfg  = (view === 'answer' && base.answerView) ? Object.assign({}, base, base.answerView) : base;
  const body     = document.getElementById('s5viq-popup-body');
  const scroller = document.getElementById('s5viq-popup-scroll');
  const chk      = document.getElementById('s5viq-check');

  popup.style.background = cfg.bg;
  if (popup.classList.contains('hidden')) resetPopupPosition(popup);

  const titleEl = document.getElementById('s5viq-popup-title');
  titleEl.innerHTML = cfg.titleSecondLine
    ? (cfg.title + '<br>' + cfg.titleSecondLine)
    : cfg.title;

  body.innerHTML = cfg.body.map(p => {
    if (p === '') return '<p class="scq-gap" aria-hidden="true"></p>';
    return '<p>' + p + '</p>';
  }).join('');

  renderPopupAction('s5viq-popup', cfg.action, s5viqToggleView);

  popup.classList.toggle('is-scroll', !!cfg.scrollGate);
  const gateApplies = !!cfg.scrollGate && !s5viqGateDone;
  s5viqGateOk = !gateApplies;
  if (chk && !s5viqGateOk) chk.disabled = true;
  if (cfg.scrollGate) scroller.setAttribute('tabindex', '0');
  else                scroller.removeAttribute('tabindex');

  popup.classList.remove('hidden');
  clampPopup(popup);
  scroller.scrollTop = 0;
  if (gateApplies) {
    placeScrollHint('s5viq-popup');
    s5viqCheckScrollGate();
  } else {
    popup.classList.remove('needs-scroll');
    if (chk && s5viqDone) chk.disabled = false;
  }
}

function s5viqCheckScrollGate() {
  if (s5viqGateOk) return;
  const scroller = document.getElementById('s5viq-popup-scroll');
  const popup = document.getElementById('s5viq-popup');
  if (!scroller || !popup) return;
  if (!atBottom(scroller)) { popup.classList.add('needs-scroll'); return; }
  /* Reached the end once — the requirement is met for this resolved
     outcome and does not come back on a later reopen. */
  s5viqGateOk = true;
  s5viqGateDone = true;
  popup.classList.remove('needs-scroll');
  const chk = document.getElementById('s5viq-check');
  if (chk && s5viqDone) chk.disabled = false;
}

window.addEventListener('resize', () => {
  const popup = document.getElementById('s5viq-popup');
  if (!popup || popup.classList.contains('hidden')) return;
  placeScrollHint('s5viq-popup');
  if (s5viqGateOk) return;
  if (s5viqResizeRaf) cancelAnimationFrame(s5viqResizeRaf);
  s5viqResizeRaf = requestAnimationFrame(s5viqCheckScrollGate);
});


/* ═══════════════════════════════════════════════════════════
   SCREENS 06 / 08 — airplane clips, and the shared success sound.

   SUCCESS SOUND. assets/audio holds exactly ONE success file, and the
   producer's ruling is that it is shared by all three success states.
   The rules that keeps it honest:
     - it fires on the TRANSITION into a success state, never on
       entering an unresolved screen and never on pressing Play;
     - each state fires at most once for the life of the page, tracked
       by key, so revisiting a completed screen is silent;
     - there is ONE element, so two states cannot overlap each other;
     - it starts only after `ended`, so it cannot overlap the clip's own
       audio — both clips carry a real AAC track;
     - navigation stops it (goTo → stopSuccessSound);
     - nothing here consults prefers-reduced-motion: that setting
       governs motion, not instructional or success audio.
   `subtle_alert.mp3` stays S2's alert and is never reused here.
   ═══════════════════════════════════════════════════════════ */
const SUCCESS_FIRED = new Set();

function playSuccessOnce(key) {
  if (SUCCESS_FIRED.has(key)) return;
  SUCCESS_FIRED.add(key);
  const a = document.getElementById('success-audio');
  if (!a) return;
  try {
    a.pause();
    a.currentTime = 0;
    const p = a.play();
    if (p && typeof p.catch === 'function') p.catch(function () {});
  } catch (e) { /* autoplay refused — never fatal, never blocking */ }
}

function stopSuccessSound() {
  const a = document.getElementById('success-audio');
  if (!a) return;
  try { a.pause(); a.currentTime = 0; } catch (e) { /* ignore */ }
}

/* Every clip in the part, stopped together on navigation so no audio or
   video can ever follow the learner onto another screen. */
function stopAllClips() {
  document.querySelectorAll('.fito-video').forEach(v => {
    try { v.pause(); } catch (e) { /* ignore */ }
  });
  ['s6', 's8'].forEach(id => {
    const sec = document.getElementById(id);
    if (sec) sec.classList.remove('is-playing');
  });
  s6Playing = false;
  s8Playing = false;
}

/* ── S6, slide 80 — the BLUE clip ── */
let s6Done = false, s6Playing = false;

function s6Play() {
  /* Duplicate-activation guard: a second press while playing, or any
     press after the clip has run, does nothing. */
  if (s6Playing || s6Done) return;
  const v = document.getElementById('s6-video');
  const sec = document.getElementById('s6');
  if (!v || !sec) return;
  s6Playing = true;
  sec.classList.add('is-playing');
  try {
    v.currentTime = 0;
    const p = v.play();
    if (p && typeof p.catch === 'function') {
      p.catch(function () { s6Playing = false; sec.classList.remove('is-playing'); });
    }
  } catch (e) { s6Playing = false; sec.classList.remove('is-playing'); }
}

function s6Ended() {
  s6Playing = false;
  s6Done = true;
  const sec = document.getElementById('s6');
  if (sec) sec.classList.add('is-done');
  const chk = document.getElementById('s6-check');
  if (chk) chk.disabled = false;
  playSuccessOnce('s6');
  /* Resume: the watch gate is state — saved the moment it is earned (guide §8.8 rule 11). */
  flushResumeSave();
}

function s6Enter() {
  const sec = document.getElementById('s6');
  const v = document.getElementById('s6-video');
  const chk = document.getElementById('s6-check');
  if (!sec || !v) return;
  s6Playing = false;
  if (s6Done) {
    /* Resolved: the finished frame stays on screen, Continue stays
       open, and nothing autoplays. */
    sec.classList.add('is-playing', 'is-done');
    try { v.pause(); if (v.duration) v.currentTime = Math.max(0, v.duration - 0.05); } catch (e) { /* ignore */ }
    if (chk) chk.disabled = false;
  } else {
    sec.classList.remove('is-playing', 'is-done');
    try { v.pause(); v.currentTime = 0; } catch (e) { /* ignore */ }
    if (chk) chk.disabled = true;
  }
}

/* ── S8, slide 82 — the GREEN clip ── */
let s8Done = false, s8Playing = false;

function s8Play() {
  if (s8Playing || s8Done) return;
  const v = document.getElementById('s8-video');
  const sec = document.getElementById('s8');
  if (!v || !sec) return;
  s8Playing = true;
  sec.classList.add('is-playing');
  try {
    v.currentTime = 0;
    const p = v.play();
    if (p && typeof p.catch === 'function') {
      p.catch(function () { s8Playing = false; sec.classList.remove('is-playing'); });
    }
  } catch (e) { s8Playing = false; sec.classList.remove('is-playing'); }
}

function s8Ended() {
  s8Playing = false;
  s8Done = true;
  const sec = document.getElementById('s8');
  if (sec) sec.classList.add('is-done');
  s8AnnounceDone(true);
  const chk = document.getElementById('s8-check');
  if (chk) chk.disabled = false;
  playSuccessOnce('s8');
  /* Resume: the watch gate is state — saved the moment it is earned (guide §8.8 rule 11). */
  flushResumeSave();
}

/* Completion announcement only. This used to also show the `הצלחה.png`
   panel over the finished clip; the producer ruled 2026-09-23 that that
   panel belongs to S9 alone, so S8's resolved state is the clip's own
   held final frame and nothing is drawn over it. The live region stays:
   the state still has to reach a screen reader in words. Called from
   exactly two places — the `ended` handler above and the resolved-state
   restoration below. */
function s8AnnounceDone(done) {
  const status = document.getElementById('s8-done-status');
  if (status) status.textContent = done ? 'مرّت الطائرة الثانية بأمان. يمكن المتابعة‎.' : '';
}

function s8Enter() {
  const sec = document.getElementById('s8');
  const v = document.getElementById('s8-video');
  const chk = document.getElementById('s8-check');
  if (!sec || !v) return;
  s8Playing = false;
  if (s8Done) {
    /* Resolved: the final frame is held. No sound — playSuccessOnce()
       has already fired for 's8' and is keyed, so a revisit is silent. */
    sec.classList.add('is-playing', 'is-done');
    try { v.pause(); if (v.duration) v.currentTime = Math.max(0, v.duration - 0.05); } catch (e) { /* ignore */ }
    s8AnnounceDone(true);
    if (chk) chk.disabled = false;
  } else {
    sec.classList.remove('is-playing', 'is-done');
    try { v.pause(); v.currentTime = 0; } catch (e) { /* ignore */ }
    s8AnnounceDone(false);
    if (chk) chk.disabled = true;
  }
}

/* ── S9, slide 83 — mission complete ──
   MASTERY. S0 authors the rule verbatim: "יש בה 4 סעיפים. עליכם להצליח
   ב - 3 לפחות." — at least 3 of the 4 peak sub-items correct. Only a
   RESOLVED-correct item counts; anything unresolved counts as not
   correct, which is what reading it off each question's own resolver
   gives for free.

   DESTINATION. There is none, and that is verified rather than assumed:
   this lomda's deck ends at slide 83, the repository holds parts 01-06
   and no 07, and this part's own header records the same finding. So
   NEXT_PART is null and no crossing is wired — inventing
   PartNav.forward('07') would point at a part that does not exist.
   The evaluation itself is real and is reported; the moment a
   successor part lands, setting NEXT_PART is the only change needed. */
const NEXT_PART = null;
let s9Crossed = false;

function s9MasteryCount() {
  return PEAK_QUESTIONS.reduce((n, q) => n + (q.resolve() === 'correct' ? 1 : 0), 0);
}
function s9MasteryMet() { return s9MasteryCount() >= 3; }

/* Platform routing (decision 2026-10-01, guide §6.4 / §6.5 / §8.8 rule 12): the last click
   reports the component 'completed' — pass AND fail — disables the button and stops; Kata routes
   on the statement (recommendedAfterFail on a fail). This supersedes the earlier producer rule
   that kept המשך enabled on a fail. The hop to the next part survives only in a ?dev=1
   walkthrough (xapiEndPart, unit-js/20-xapi.js). */
function s9Continue() {
  if (s9Crossed) return;
  s9Crossed = true;
  xapiEndPart(lastScreenButton(), NEXT_PART);   // last part of the lomda: NEXT_PART is null
}

function s9Enter() {
  const img = document.getElementById('s9-panel-art');
  if (img && !img.getAttribute('src')) img.src = 'assets/img/הצלחה.png';
  /* The completion sound fires on the learner's FIRST arrival at the
     completed state, keyed like the other two so a revisit is silent. */
  playSuccessOnce('s9');
}


/* ═══════════════════════════════════════════════════════════
   SCREEN 07 — SingleChoiceQuestion (slide 81). PEAK QUESTION 4 of 4.

   RATIO ORDER. The slide's XML string for the stem reads "3 : 1", but
   PowerPoint RENDERS it "1 : 3" — the digits are separate runs and the
   RTL paragraph orders them right-to-left. That was settled by
   measurement, not by reading the XML: an ink-column profile over the
   rendered island puts the lighter glyph (68 ink px, a "1") to the LEFT
   of the colon and the heavier (96 px, a "3") to its right. The deck's
   own feedback text agrees twice over — it writes "1:3" directly and
   concludes α = 45°, which only follows if α is ONE part of four.
   This build therefore authors 1 : 3 inside an LTR island, so the
   authored order and the visible order are the same string.

   THE MATHEMATICS. α : β = 1 : 3, so α is x and β is 3x and the sum is
   4x. α's share is 1x/4x = 1/4 of 180° = 45°. Omer's error is using β
   (3x) as the denominator instead of the SUM, giving 1/3 and 60°, so
   the answer is "לא".
   (2026-09-30: a client-round note asked for "כן" / 1x/3x / 60°. Checked
   and reverted - with 1 : 3, α + β = 180° gives α = 45°; 60° would need
   1 : 2. After talking to the client the feedback's first sentence was
   reworded to "הביטוי האלגברי המתאר את החלק של α מתוך סכום הזוויות הוא
   1x/4x", which names the sum; the answer stays "לא".)
   ═══════════════════════════════════════════════════════════ */
const S7SC = { correct: 'no', maxAttempts: 2 };

const S7SC_EXPLAIN = [
  'التعبير الجبري الذي يصف جزء الزاوية <bdi class="scq-expr" dir="ltr">α</bdi> من مجموع الزاويتين هو <span class="math-frac" role="img" aria-label="واحد إكس على أربعة إكس"><span class="math-frac__n">1<i>x</i></span><span class="math-frac__d">4<i>x</i></span></span>.',
  'وبعد الاختزال نرى أن <bdi class="scq-expr" dir="ltr">α</bdi> مساوية لـ <span class="math-frac" role="img" aria-label="واحد على أربعة"><span class="math-frac__n">1</span><span class="math-frac__d">4</span></span> من الـ <bdi class="scq-expr" dir="ltr">°180</bdi>،<br>ولذلك مقدارها هو <bdi class="scq-expr" dir="ltr">°45</bdi>.'
];

const S7SC_POPUP = {
  retry:   { bg: '#ffdbdc', title: 'الإجابة غير صحيحة.', body: ['هل نحاول مرة أخرى؟'] },
  correct: { bg: '#edf8ed', title: 'صحيح‎!', body: S7SC_EXPLAIN },
  /* Authored title, verbatim — slide 81 id=22 — plus the canonical
     second line _question-template-defaults.md appends whenever an
     outcome reveals the correct answer. */
  wrong2:  { bg: '#ffdbdc', title: 'هذا خطأ،<br>لكن مجرد التعامل مع المسألة يطوّر التفكير‎!',
             titleSecondLine: 'الإجابة الصحيحة تَظهَر الآن.',
             body: S7SC_EXPLAIN }
};


/* ── S7 behaviour ──
   Cloned from S4's verified block and re-scoped: same two attempts,
   same retry/terminal shape, same two-flag gate (GateOk momentary,
   GateDone the reading requirement that survives a reopen), same
   resolved-state restoration. Re-pointed to peak sub-item 4. */
let s7scSelected  = null;   // ONE id, or null — the single-select difference
let s7scAttempts  = 0;
let s7scDone      = false;
let s7scResult    = null;   // 'correct' | 'wrong2'
let s7scLastWrong = null;   // the id of the immediately-preceding wrong attempt
let s7scGateOk    = true;
let s7scGateDone  = false;  // the reading requirement — see docs/scroll-gate-restore.md

function s7scOpt(id) { return document.querySelector('#s7 .scq-opt[data-id="' + id + '"]'); }
function s7scOpts()  { return [...document.querySelectorAll('#s7 .scq-opt')]; }

function s7scSelect(id) {
  if (s7scDone) return;
  s7scSelected = id;          // replaces, never accumulates
  s7scRender();
  /* Established retry-popup dismissal: once an attempt has been made,
     touching an option puts the learner back in correcting mode. */
  if (s7scAttempts > 0) {
    document.getElementById('s7sc-popup').classList.add('hidden');
    s7scOpts().forEach(o => o.classList.remove('correct', 'wrong'));
  }
}

function s7scRender() {
  s7scOpts().forEach(o => {
    const on = o.dataset.id === s7scSelected;
    o.classList.toggle('selected', on);
    o.setAttribute('aria-checked', on ? 'true' : 'false');
  });
  const chk = document.getElementById('s7sc-check');
  if (!chk || s7scDone) return;
  /* Gate on "something is chosen", plus the retry gate: re-submitting
     the identical wrong choice stays blocked until it changes. */
  chk.disabled = s7scSelected === null || s7scSelected === s7scLastWrong;
}

function s7scCheck() {
  if (s7scDone) { advanceScreen(); return; }
  if (s7scSelected === null) return;
  s7scAttempts++;
  reportQ(['001', 'q4'], s7scSelected === S7SC.correct, s7scSelected === S7SC.correct || s7scAttempts >= S7SC.maxAttempts, xapiChoiceAnswer('001', 'q4', '#s7 .scq-opt', s7scSelected));

  if (s7scSelected === S7SC.correct) {
    s7scResult = 'correct';
    s7scOpt(S7SC.correct).classList.add('correct');
    s7scFinish();
    s7scShowPopup('correct');
    return;
  }

  if (s7scAttempts < S7SC.maxAttempts) {
    /* FIRST WRONG — retry. The learner's own pick is marked wrong (red +
       ✗), as on every other single-choice screen in this lomda
       (SingleChoiceQuestion.md → state 4, Try Again); the CORRECT option
       is not revealed yet. Corrected 2026-09-23 on the producer's report
       — this screen and S4 shipped with no marking at all on the first
       attempt, unlike the template. */
    s7scOpt(s7scSelected).classList.add('wrong');
    s7scLastWrong = s7scSelected;
    const chk = document.getElementById('s7sc-check');
    if (chk) chk.disabled = true;
    s7scShowPopup('retry');
    return;
  }

  /* TERMINAL WRONG — the correct option is revealed alongside the
     learner's own wrong pick, both visible at once, which is why this
     screen needs no View My Answer Toggle. */
  s7scResult = 'wrong2';
  s7scOpt(s7scSelected).classList.add('wrong');
  s7scOpt(S7SC.correct).classList.add('correct');
  s7scFinish();
  s7scShowPopup('wrong2');
}

function s7scFinish() {
  s7scDone = true;
  s7scOpts().forEach(o => { o.disabled = true; });
  const chk = document.getElementById('s7sc-check');
  /* SingleChoiceQuestion's completion fallback — slide 78 authors no
     progression copy of its own, same as slide 77. The button stays
     disabled while a scroll-gated outcome is unread; the gate releases
     it. */
  if (chk) setNavLabel(chk, 'هل نتابع؟');
  hideHintButton('s7sc-hint');
  /* PEAK QUESTION 2 resolves on this screen alone — no sections — so
     the dot settles here, immediately. */
  syncPeakProgressNav(document.getElementById('s7'));
}

function s7scResetInitial() {
  s7scSelected  = null;
  s7scAttempts  = 0;
  s7scDone      = false;
  s7scResult    = null;
  s7scLastWrong = null;
  s7scGateOk    = true;
  s7scGateDone  = false;
  s7scOpts().forEach(o => {
    o.disabled = false;
    o.classList.remove('selected', 'correct', 'wrong');
    o.setAttribute('aria-checked', 'false');
  });
  const popup = document.getElementById('s7sc-popup');
  popup.classList.add('hidden');
  popup.classList.remove('is-scroll', 'needs-scroll');
  document.getElementById('s7sc-hint-overlay').classList.add('hidden');
  const chk = document.getElementById('s7sc-check');
  if (chk) { setNavLabel(chk, 'هل إجابتي صحيحة؟'); chk.disabled = true; }
  showHintButton('s7sc-hint');
}

/* A resolved screen never comes back without its feedback — see
   docs/scroll-gate-restore.md and _global-components.md → "Reopening
   feedback on return to a resolved screen". Nothing is re-scored:
   s7scCheck() is not involved. */
function s7scRestoreResolved() {
  s7scOpts().forEach(o => { o.disabled = true; });
  const chk = document.getElementById('s7sc-check');
  if (chk) { setNavLabel(chk, 'هل نتابع؟'); chk.disabled = false; }
  hideHintButton('s7sc-hint');
  s7scShowPopup(s7scResult === 'correct' ? 'correct' : 'wrong2');
}

function s7scShowPopup(type) {
  const popup = document.getElementById('s7sc-popup');
  if (!popup) return;
  const cfg      = S7SC_POPUP[type];
  const body     = document.getElementById('s7sc-popup-body');
  const scroller = document.getElementById('s7sc-popup-scroll');
  const chk      = document.getElementById('s7sc-check');

  popup.style.background = cfg.bg;
  if (popup.classList.contains('hidden')) resetPopupPosition(popup);

  const titleEl = document.getElementById('s7sc-popup-title');
  titleEl.innerHTML = cfg.titleSecondLine
    ? (cfg.title + '<br>' + cfg.titleSecondLine)
    : cfg.title;

  body.innerHTML = cfg.body.map(p => {
    if (p === '') return '<p class="scq-gap" aria-hidden="true"></p>';
    return '<p>' + p + '</p>';
  }).join('');

  popup.classList.toggle('is-scroll', !!cfg.scrollGate);
  const gateApplies = !!cfg.scrollGate && !s7scGateDone;
  s7scGateOk = !gateApplies;
  if (chk && !s7scGateOk) chk.disabled = true;
  if (cfg.scrollGate) scroller.setAttribute('tabindex', '0');
  else                scroller.removeAttribute('tabindex');

  popup.classList.remove('hidden');
  clampPopup(popup);
  scroller.scrollTop = 0;
  if (gateApplies) {
    placeScrollHint('s7sc-popup');
    s7scCheckScrollGate();
  } else {
    popup.classList.remove('needs-scroll');
    if (chk && s7scDone) chk.disabled = false;
  }
}

function s7scCheckScrollGate() {
  if (s7scGateOk) return;
  const scroller = document.getElementById('s7sc-popup-scroll');
  const popup    = document.getElementById('s7sc-popup');
  if (!scroller || !popup) return;
  if (!atBottom(scroller)) { popup.classList.add('needs-scroll'); return; }
  s7scGateOk = true;
  s7scGateDone = true;
  const chk = document.getElementById('s7sc-check');
  if (chk && s7scDone) chk.disabled = false;
  popup.classList.remove('needs-scroll');
}

let s7scResizeRaf = null;
window.addEventListener('resize', () => {
  const popup = document.getElementById('s7sc-popup');
  if (!popup || popup.classList.contains('hidden')) return;
  placeScrollHint('s7sc-popup');
  if (s7scGateOk) return;
  if (s7scResizeRaf) cancelAnimationFrame(s7scResizeRaf);
  s7scResizeRaf = requestAnimationFrame(s7scCheckScrollGate);
});

function s7scOpenHint()  { xapiRequestedHint('001', 'q4'); document.getElementById('s7sc-hint-overlay').classList.remove('hidden'); }
function s7scCloseHint() { document.getElementById('s7sc-hint-overlay').classList.add('hidden'); }
function s7scCloseHintOnBackdrop(e) { if (e.target === e.currentTarget) s7scCloseHint(); }

function s7scEnter() {
  const sec = document.getElementById('s7');
  if (!s7scDone && screenPristine(7)) s7scResetInitial();
  s7scRender();
  if (s7scDone) s7scRestoreResolved();
  syncPeakProgressNav(sec);
  markPeakCurrent(sec, 4);
}


/* ─── S6 / S8 / S9 asset mounting ───────────────────────────
   Sources live HERE, not in the markup, so swapping an asset is a
   one-line change in one place — the same rule S1 and S2 follow.

   Every filename below was INVENTORIED ON DISK, not inferred. Note in
   particular that the blue clip is `Blue_airplane_moving_on_radar.mp4`
   — the doubled ".mp4.mp4" the file briefly carried has been corrected
   by the producer and nothing references it.

   S8 deliberately has NO still: slide 82's own background is
   `קונטרול2`, whose radar mirrors A and B against every other screen in
   this part AND against this screen's own clip. The clip is the whole
   media layer instead, and its poster is its OWN opening frame, so the
   mirrored artwork never appears — not as a background, not as a
   poster, not as a fallback. */
const LATE_ASSETS = {
  's9-img':       { src: 'assets/img/קונטרול.jpg',
                    alt: 'برج مراقبة في مطار ليلًا. يظهر على شاشة الرادار مسارا طيران يتقاطعان' },
  's9-panel-art': { src: 'assets/img/הצלחה.png' }
};
const LATE_CLIPS = {
  's6-video': { src: 'assets/video/Blue_airplane_moving_on_radar.mp4',
                poster: 'assets/img/blue-video-frame0.jpg' },
  's8-video': { src: 'assets/video/Green_airplane_moving_on_radar.mp4',
                poster: 'assets/img/green-video-frame0.jpg' }
};

function mountLateAssets() {
  Object.keys(LATE_ASSETS).forEach(id => {
    const el = document.getElementById(id);
    if (!el) return;
    const a = LATE_ASSETS[id];
    el.setAttribute('src', a.src);
    if (a.alt !== undefined) el.setAttribute('alt', a.alt);
  });
  /* The clips (and their posters) go in as data-src / data-poster, NOT src: both <video> carry
     preload="auto", so a src set here, at parse time, fetched ~1.9 MB of hidden video during the
     boot window, ahead of the restore (BOOT-AND-RESUME-HARDENING-2026-11-03.md §2.3 C4, 2026-11-14).
     goTo() moves them to src / poster when S6 or S8 becomes the active screen, before its
     resetScreenState — the enter hook and the resume painter (resumeHoldClipEnd) then find a src. */
  Object.keys(LATE_CLIPS).forEach(id => {
    const el = document.getElementById(id);
    if (!el) return;
    const c = LATE_CLIPS[id];
    el.setAttribute('data-src', c.src);
    if (c.poster) el.setAttribute('data-poster', c.poster);
  });
}
mountLateAssets();

/* ===== ported from part 04 — Image zoom ===== */
/* IMAGE ZOOM — the canonical component, one modal per part, ported
   verbatim from methodica-math-ratio-03-04. Required here because S3
   presents a primary learning image: _question-template-defaults.md →
   "Image zoom" makes it the default for any question template showing
   one, and the asset is genuinely worth zooming — 1536x1024 supplied,
   shown at 389x259. Slide 77 authors no zoom affordance of its own;
   this is the project default applied, and it is reported as such. */
let zoomReturnFocus = null;

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
  const cloneBtn = clone.querySelector('.img-zoom-btn');
  if (cloneBtn) cloneBtn.remove();
  clone.querySelectorAll('[id]').forEach(el => el.removeAttribute('id'));
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
       than the viewport is scaled UP so the zoom actually enlarges. */
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
  /* Focus restoration — back to the trigger, so a keyboard learner is
     never dropped at the top of the page. */
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
/* A screen change must never leave an enlarged image over the next
   screen — goTo() sweeps popups and hint overlays by id suffix, which
   this modal deliberately does not match, so it is closed explicitly. */
function closeImageZoomOnNavigate() {
  const modal = document.getElementById('img-zoom-modal');
  if (modal && !modal.classList.contains('hidden')) closeImageZoom();
}

/* ─── PEAK QUESTION progress model ──────────────────────────
   Part 06 is a peak question: FOUR scored sub-items, and the whole
   question counts as ONE result that is correct when at least 3 of the
   4 are. Producer ruling 2026-09-14, from slide 74's own line
   ("יש בה 4 סעיפים. עליכם להצליח ב - 3 לפחות") and confirmed against
   each question slide's own note, which numbers them "שאלה N מתוך 4".

   Slide → screen map, each resolved through its own .rels:
       slide 77  →  S3   MultipleChoiceQuestion
       slide 78  →  S4   SingleChoiceQuestionImage
       slide 79  →  S5   ValueInputQuestion
       slide 81  →  S7   SingleChoiceQuestionImage

   FOUR dots, not six. The slides carry a flattened progress-bar mock
   (media/image22.png) showing six items in a half-solved state, but
   that same PNG is pasted on 29 slides across the whole deck and every
   part builds its own live bar sized to its own question count (-02: 3,
   -03: 3, -04: 4, -05: 4). The mock is not traced.

   NO PART-LEVEL SCORE IS EMITTED HERE. The four individual results are
   kept for diagnostics and the aggregate stays UNRESOLVED until all
   four sub-items resolve; whoever assembles the part result reads them.
   Outcome reporting itself is owned externally by Vadim's integration
   layer and is out of scope for this lomda. */
const PEAK_QUESTIONS = [
  { n: 1, screens: [3], resolve: () => (s3mcDone ? (s3mcResult === 'correct' ? 'correct' : 'incorrect') : 'not-answered') },
  { n: 2, screens: [4], resolve: () => (s4scDone ? (s4scResult === 'correct' ? 'correct' : 'incorrect') : 'not-answered') },
  { n: 3, screens: [5], resolve: () => (s5viqDone ? (s5viqResult === 'correct' ? 'correct' : 'incorrect') : 'not-answered') },
  { n: 4, screens: [7], resolve: () => (s7scDone ? (s7scResult === 'correct' ? 'correct' : 'incorrect') : 'not-answered') }
];

const peakProgress = {
  questions: PEAK_QUESTIONS.map(q => ({ state: 'not-answered', visited: false, screens: q.screens })),
  results: []
};

/* Re-derives every sub-item from the screens' own state and renders the
   bar if the current screen carries one. Called on entry into each
   question screen and from each question's finish handler, so returning
   to an earlier sub-item after resolving a later one shows the right
   state on both. */
function syncPeakProgressNav(sectionEl) {
  PEAK_QUESTIONS.forEach((q, i) => {
    const item = peakProgress.questions[i];
    item.state   = q.resolve();
    item.visited = item.state !== 'not-answered';
    item.screens = q.screens;
  });
  peakProgress.results = peakProgress.questions.map(q => q.state);

  /* The aggregate the producer ruled on: at least 3 of 4 correct, and
     UNRESOLVED until every sub-item has resolved. Exposed for
     diagnostics only — nothing here reports it anywhere. */
  const resolved = peakProgress.results.filter(r => r !== 'not-answered').length;
  const correct  = peakProgress.results.filter(r => r === 'correct').length;
  window.lomdaState.peakSubItems = peakProgress.results.slice();
  window.lomdaState.peakResult =
    resolved < PEAK_QUESTIONS.length ? 'unresolved' : (correct >= 3 ? 'correct' : 'incorrect');

  const nav = sectionEl && sectionEl.querySelector('.progress-question');
  if (nav) updateProgressQuestion(nav, peakProgress);
}

/* Mark the CURRENT question's dot while the learner is on its screen —
   the state the bar shows for a sub-item that has been reached but not
   yet answered. Called after the sync so it wins over 'not-answered'. */
function markPeakCurrent(sectionEl, n) {
  const nav = sectionEl && sectionEl.querySelector('.progress-question');
  if (!nav) return;
  const item = nav.querySelector('[data-question="' + n + '"]');
  if (!item) return;
  const icon = item.querySelector('.progress-question__icon');
  if (peakProgress.questions[n - 1].state === 'not-answered') {
    icon.classList.add('progress-question__icon--current');
  }
}

/* S3 — MultipleChoiceQuestion (slide 77). Mounts the image, re-derives
   the progress bar and restores whatever state the learner left the
   screen in. NOT a reset: returning to a resolved question must show it
   resolved, which is why s3mcResetInitial() is called only while the
   question is pristine (nothing selected, nothing attempted). */
function s3mcEnter() {
  const sec = document.getElementById('s3');
  if (!s3mcDone && screenPristine(3)) s3mcResetInitial();
  s3mcRender();
  /* Resolved → restore the complete resolved presentation, feedback
     popup included. s3mcRender() has already re-applied the learner's
     own selection classes and returns early once s3mcDone, so it never
     touches the progression button here. */
  if (s3mcDone) s3mcRestoreResolved();
  syncPeakProgressNav(sec);
  markPeakCurrent(sec, 1);
}

/* Cross-part navigation. init() must run AFTER this part's own initial
   screen setup: if this page was entered via Back, it jumps straight to
   the screen the learner left from. */
PartNav.init({
  part: '06',
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
   Ids must match metadata/methodica-math-ratio-03-06.json byte-for-byte (the loader
   checks XAPI_COMP_ID against METADATA.id on every load; _test/verify-static.js checks the map).
   ═══════════════════════════════════════════════════════════ */
const XAPI_COMP_SLUG     = 'methodica-math-ratio-03-06';
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
  8: ['001', 9],
  9: ['001', 10]
};

/* The gate S0 promises: "3 of 4 sections" (s9MasteryMet). Score = sections correct / 4. */
function partGateResult() {
  return { success: s9MasteryMet(), score: { scaled: s9MasteryCount() / PEAK_QUESTIONS.length } };
}
function lastScreenButton() { return currentScreen === PART_LAST ? document.getElementById('s9-check') : null; }

/* ═══════════════════════════════════════════════════════════
   Resume painters — unit-js/45-painters.js reads this registry.
   What each screen saves (vars + the fields that hold state) and how a screen restored from the
   state document is drawn again on the fresh markup: init() is the screen's own ResetInitial,
   paint() replays the template's own terminal look (marks / Render / RestoreResolved / ApplyView)
   or the retry look, without counting an attempt and without reporting. In-session re-entry still
   goes through resetScreenState(), which now resets only a pristine screen.
   Gates are state (rule 11): the S6 / S8 watch gates and every question's scroll gate
   ({prefix}GateOk + GateDone, docs/scroll-gate-restore.md) are saved, and RestoreResolved /
   ShowPopup re-derive the momentary GateOk from GateDone exactly as an in-session return does.
   SUCCESS_FIRED is saved so a restored finished screen (or S9) never replays the success sound;
   s9Crossed is the S9 latch. The peak bar (peakProgress, lomdaState.peakSubItems/peakResult) is
   derived from the screens' own vars by syncPeakProgressNav() on every entry — not saved.
   ═══════════════════════════════════════════════════════════ */
const PART_STATE_VARS = ['SUCCESS_FIRED', 's9Crossed'];

/* The finished frame of a watched clip — what s6Enter / s8Enter seek to in-session. After a
   reload the metadata may not be in yet (duration NaN), so the seek waits for it once. */
function resumeHoldClipEnd(id, isDone) {
  const v = document.getElementById(id);
  if (!v || !isDone()) return;
  const seek = () => { try { if (isDone() && v.duration) { v.pause(); v.currentTime = Math.max(0, v.duration - 0.05); } } catch (e) {} };
  if (v.duration) seek(); else v.addEventListener('loadedmetadata', seek, { once: true });
}

const PART_SCREENS = {
  3: {
    vars: ['s3mcSelected', 's3mcAttempts', 's3mcDone', 's3mcResult', 's3mcLastWrong', 's3mcGateOk', 's3mcGateDone'],
    pristine: () => !s3mcSelected.size && !s3mcAttempts && !s3mcDone,
    init: s3mcResetInitial,
    paint: () => {
      if (s3mcDone) {
        s3mcOpts().forEach(o => {                  // the terminal marks s3mcCheck() left
          o.classList.remove('selected', 'correct', 'wrong');
          if (S3MC.correct.has(o.dataset.id)) o.classList.add('correct');
          else if (s3mcResult !== 'correct' && s3mcSelected.has(o.dataset.id)) o.classList.add('wrong');
        });
        s3mcRender();                              // as s3mcEnter: selection classes back on
        s3mcRestoreResolved();                     // lock, label, popup + the scroll gate as saved
        return;
      }
      if (s3mcAttempts && s3mcSig(s3mcSelected) === s3mcLastWrong) {
        s3mcOpts().forEach(o => {                  // the retry marks, on the submitted set only
          if (!s3mcSelected.has(o.dataset.id)) return;
          o.classList.remove('selected');
          o.classList.add(S3MC.correct.has(o.dataset.id) ? 'correct' : 'wrong');
        });
      }
      s3mcRender();                                // selection + Check enablement + the retry gate
    }
  },
  4: {
    vars: ['s4scSelected', 's4scAttempts', 's4scDone', 's4scResult', 's4scLastWrong', 's4scGateOk', 's4scGateDone'],
    pristine: () => s4scSelected === null && !s4scAttempts && !s4scDone,
    init: s4scResetInitial,
    paint: () => {
      if (s4scDone) {
        if (s4scResult !== 'correct' && s4scOpt(s4scSelected)) s4scOpt(s4scSelected).classList.add('wrong');
        s4scOpt(S4SC.correct).classList.add('correct');
        s4scRender();
        s4scRestoreResolved();
        return;
      }
      if (s4scAttempts && s4scSelected !== null && s4scSelected === s4scLastWrong) s4scOpt(s4scSelected).classList.add('wrong');
      s4scRender();
    }
  },
  5: {
    vars: ['s5viqAttempts', 's5viqDone', 's5viqResult', 's5viqLastWrong', 's5viqAnswer', 's5viqView', 's5viqGateOk', 's5viqGateDone'],
    inputs: ['s5viq-input-alpha', 's5viq-input-beta'],
    pristine: () => !s5viqAttempts && !s5viqDone && resumeFieldsEmpty(PART_SCREENS[5].inputs),
    init: s5viqResetInitial,
    paint: () => {
      if (s5viqDone) {
        /* correct: the marks first — s5viqRestoreResolved() only reopens the popup on that path;
           wrong2: s5viqApplyView() (via RestoreResolved) writes values + marks for the saved view. */
        if (s5viqResult === 'correct') S5VIQ.slots.forEach(k => s5viqMark(k, 'correct'));
        s5viqRestoreResolved();
        return;
      }
      if (s5viqAttempts && s5viqSig() === s5viqLastWrong) {
        S5VIQ.slots.forEach(k => s5viqMark(k,
          s5viqNormalize(s5viqInput(k).value) === s5viqNormalize(S5VIQ.correct[k]) ? 'correct' : 'wrong'));
      }
      s5viqUpdateBar();
    }
  },
  6: {
    vars: ['s6Done'],
    pristine: () => !s6Done,
    init: () => { s6Done = false; s6Enter(); },
    paint: () => { s6Enter(); resumeHoldClipEnd('s6-video', () => s6Done); }
  },
  7: {
    vars: ['s7scSelected', 's7scAttempts', 's7scDone', 's7scResult', 's7scLastWrong', 's7scGateOk', 's7scGateDone'],
    pristine: () => s7scSelected === null && !s7scAttempts && !s7scDone,
    init: s7scResetInitial,
    paint: () => {
      if (s7scDone) {
        if (s7scResult !== 'correct' && s7scOpt(s7scSelected)) s7scOpt(s7scSelected).classList.add('wrong');
        s7scOpt(S7SC.correct).classList.add('correct');
        s7scRender();
        s7scRestoreResolved();
        return;
      }
      if (s7scAttempts && s7scSelected !== null && s7scSelected === s7scLastWrong) s7scOpt(s7scSelected).classList.add('wrong');
      s7scRender();
    }
  },
  8: {
    vars: ['s8Done'],
    pristine: () => !s8Done,
    init: () => { s8Done = false; s8Enter(); },
    paint: () => { s8Enter(); resumeHoldClipEnd('s8-video', () => s8Done); }
  }
};
