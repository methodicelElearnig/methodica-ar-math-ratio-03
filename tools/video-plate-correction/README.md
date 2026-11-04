# Video background-plate correction

Lifts a near-white background plate in a character clip to true `#FFFFFF`, so
the clip blends into the module's `#fff` canvas instead of reading as a faint
rectangular panel.

These scripts live in the repo deliberately. Parts 03-05 and 03-06 recorded the
same method but kept the scripts only in a session scratchpad, which does not
survive; both sets of notes flag that as a problem. This is that fix.

## Outcome, 2026-09-22

The correction below fixed all ten defective clips and was validated. The
producer then reworked most character clips through Canva background removal,
which yields a true white plate at ~600 KB per clip but deletes the ground
shadow and exports at 30 fps. Seven clips still ship on this pipeline —
burnt_popcorn ×2 (2026-09-14), Character_2_sign (never defective),
Character_2_warm-up, Proud ×2 (cq 19), Character_2_Climbing (re-encoded cq 19).
Per-part status tables are at the top of each `docs/asset-corrections.md`.

## The defect

A clip supplied on a plate of RGB 251–253 rather than 255. Against the module's
`#FFFFFF` it shows as a rectangle the size of the video box. It may **not** be
concealed with a CSS filter, blend mode, mask or a white panel behind the video
— standing producer ruling, recorded in `methodica-math-ratio-03-05/docs/` and
`-06/docs/asset-corrections.md`. It is fixed in the asset file itself.

## Method — `process_yuv3.js`

```
node process_yuv3.js in.yuv out.yuv W H FRAMES [TOL] [FULL] [CHROMA]
```

- Operates natively on the decoded `yuv420p` plane. No RGB round trip.
- The region is an **edge-connected flood fill from the frame border**, so
  anything enclosed by the artwork (teeth, eyes, shoes, highlights) is
  unreachable.
- Every candidate must also be **chroma-neutral** (`|U-128|, |V-128| <= CHROMA`),
  which excludes coloured artwork on colour alone.
- **Chroma is never written.** `chromaSamplesTouched` must be 0; it is asserted
  in the script's own output.
- The lift is a **smooth two-point ramp**, not a threshold:

  ```
  lo = plate(x,y) - TOL      w = clamp((Y - lo) / (hi - lo), 0, 1)
  hi = plate(x,y) - FULL     Y += (WHITE_Y - Y) * w
  ```

  A pixel at or above `plate - FULL` lifts fully to white; one at or below
  `plate - TOL` is untouched; between them it tapers. Monotonic and continuous
  at **both** ends, so no contour can form where the correction stops.
- `plate(x,y)` is a **local** estimate — per 32×32 block, the median luma of
  masked (background) pixels only, diffused into character-covered blocks,
  box-blurred, bilinearly upsampled, and clamped so it can never sit more than
  6 units below white (which is what stops a block lying wholly inside a soft
  ground shadow from declaring that falloff to be "the plate"). On a flat plate
  this reproduces a global seed block for block.

`WHITE_Y = 235` — these assets are **limited range (16–235)**. Verify before
reusing: a border median of `Y = 233/234` against an RGB plate of 253 matches
the limited-range prediction (`16 + 219·253/255 = 233`), not the full-range one.
At full range `WHITE_Y` would need to be 255 and the highlights would clip.

### Why `FULL` exists

Parts 03-01/03-03/03-06 used a single-point ramp (`FULL = 0`, `TOL = RAMP`).
That measures every pixel's depth from the plate *median*, so a genuine
background pixel sitting 2 units below the median is deliberately under-lifted
and stops short of white. With `FULL = 3` every pixel within 3 units of the
local plate goes to white exactly, while `TOL = 6` keeps the fill from reaching
into artwork. Setting `FULL = 0` and `TOL = RAMP` recovers the old behaviour
exactly.

**Parameters used for the 2026-09-22 batch: `TOL 6, FULL 3, CHROMA 4`.**
They were not carried over blindly — `TOL 12 / RAMP 12` was tried first and
rejected on evidence: on `Character_1_Book` it leaked into the cream book pages
(1,243 px/frame changed by more than 8, correction's own share 0.231). Measured
on that clip, true background is `Y >= 232` with `U = V = 128` exactly, while
the page pixels that leaked are `Y <= 230` with chroma 4–6 off neutral — so the
tighter gate separates them cleanly. Re-measure if a new clip behaves
differently; do not assume these numbers.

## Encoding

```
ffmpeg -i in.mp4 -f rawvideo -pix_fmt yuv420p -vsync 0 src.yuv
node process_yuv3.js src.yuv out.yuv 1280 720 <frames> 6 3 4
ffmpeg -f rawvideo -pix_fmt yuv420p -s 1280x720 -r <fps> -i out.yuv -i in.mp4 \
  -map 0:v -map 1:a -c:a copy -c:v h264_nvenc \
  -preset p7 -tune hq -bf 3 -b_ref_mode middle -rc-lookahead 32 \
  -spatial-aq 1 -temporal-aq 1 -aq-strength 8 -multipass fullres \
  -rc vbr -cq 19 -b:v 0 \
  -pix_fmt yuv420p -profile:v high -movflags +faststart out.mp4
```

Drop `-i in.mp4 -map 1:a -c:a copy` for a clip with no audio track. Always
re-derive from the **original** file, not from a previously corrected encode,
so there is no generation loss.

`ffmpeg` on this machine: `C:/Users/sharons/AppData/Local/CapCut/Apps/9.3.0.3970/ffmpeg.exe`.
The build ships **no `ffprobe`** and **no software x264** — hardware encoders
only. Every ffmpeg on the machine was checked (three CapCut builds and
Playwright's); none has libx264. `h264_qsv` does not initialise here
("Error initializing a MFX session: unsupported"), `h264_mf` rejects a named
profile, and `h264_amf` is unusable on furry character animation at any qp
(part 03-05's notes have the measurements). So nvenc is the only option.

### The rate-control settings matter more than the quality level

Parts 03-01 through 03-06 encoded with `-rc constqp -qp 12` and **nvenc
defaults otherwise** — which means no B-frames, no B-pyramid, no lookahead, no
adaptive quantisation and no multipass. That is what made the corrected files
roughly 10x their originals, not the correction and not really the quality
target. Measured on two clips, re-encoding the *same* corrected frames with
those features enabled and a quality-targeted VBR:

| encode | Book clip | Waving clip | plate not-255 | character delta at display size |
|---|---|---|---|---|
| `constqp -qp 12`, nvenc defaults | 17.35 MB | 7.91 MB | 1.39% / 1.71% | — (reference) |
| tuned `constqp -qp 12` | 19.02 MB | 8.34 MB | 1.39% / 1.76% | — |
| **tuned `vbr -cq 19`** | **5.75 MB** | **3.09 MB** | 1.91% / 2.05% | 1.5–1.7, >8 on **0.03%** |
| tuned `vbr -cq 23` | 3.75 MB | 2.00 MB | 2.33% / 2.17% | 2.2–2.4, >8 on 0.7% |
| tuned `vbr -cq 26` | 2.71 MB | 1.44 MB | 2.64% / 2.42% | 2.9–3.1, >8 on 2.9% |
| tuned `vbr -cq 29` | 1.94 MB | 1.01 MB | 4.75% / 2.68% | 3.7–3.9, >8 on 7.8% |
| tuned `vbr -cq 32` | 1.43 MB | 0.74 MB | 5.16% / 5.16% | flicker 0.90 — **rejected** |

Note `constqp` gets *bigger* when tuned: constant-QP cannot spend the extra
efficiency, so it just adds B-frame overhead. The gain only appears with VBR.

**The white plate survives compression.** This is the non-obvious result: a flat
white field is nearly free to encode, so the bits go to the fur, and the plate
stays at a mean of 254.9+ even at cq 29. Raising cq does **not** bring the panel
back — it costs character detail. cq 32 is the exception and is rejected: the
per-frame plate mean starts to wander (flicker spread 0.05 → 0.90), i.e. the
background begins to shimmer, which is the original defect returning.

Measure quality **at display size**, not 1:1 — these clips are shown in a
640×360 box, so a 1280×720 pixel-peep overstates the loss by a lot.
`fidelity_at_display_size.js` does this.

## Always run the encoder-only control

Push the **unprocessed** source through the identical decode → encode path and
measure it the same way, before attributing anything to the correction.
Otherwise compression loss is mistaken for correction leakage. `fidelity.js`
takes the control as its third argument and reports the correction's own share
as the difference.

## Validation scripts

| script | what it answers |
|---|---|
| `scan.js <clips…>` | Is a clip on a true `#FFFFFF` plate? Border rings, whole-frame near-white counts, per-clip histogram. The detector. |
| `sides.js <clips…>` | Per-side border breakdown and longest contiguous non-255 run. Distinguishes a uniform plate edge (long runs on **all four** sides) from artwork touching the frame (localized runs, three sides clean). |
| `final_check.js <orig> <corrected>` | Plate-only purity. Rebuilds the unambiguous-background set from the original, erodes it 3px, reports what that same set looks like in the corrected encode, plus per-frame flicker. The headline metric. |
| `fidelity.js <orig> <cand> [control]` | Mean abs delta over character pixels only, and the correction's own share against the control. |
| `png.js <clip> <frame> <out.png> [plate]` | Renders a frame; `plate` mode recolours it — white = exactly `#FFFFFF`, red = 250–254, orange = 235–249, grey = darker. Makes a plate visible at a glance. |

A page-composited check is worth running too: serve the part over HTTP, load the
real `index.html`, draw the `<video>` to a canvas (same-origin, so untainted)
and read the four frame edges at first frame, mid-playback and the held final
frame. That is what confirms no CSS is quietly altering the result.
