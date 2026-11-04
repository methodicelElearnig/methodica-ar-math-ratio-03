// process_yuv3.js — lift a near-white background plate to true white, in YUV.
//
// Reimplementation of the method validated on this lomda (parts 03-01, 03-03,
// 03-04, 03-05, 03-06), with 03-05's spatially-varying plate estimate:
//
//   - operates natively on yuv420p; no RGB round trip
//   - the region is an edge-connected flood fill FROM THE FRAME BORDER, so
//     anything enclosed by the artwork (teeth, eyes, shoes, highlights) is
//     unreachable
//   - every candidate must also be chroma-neutral (|U-128|,|V-128| <= CHROMA),
//     which excludes coloured artwork on colour alone
//   - chroma is never written
//   - the lift is a smooth ramp, not a threshold:
//         lift = (WHITE_Y - Y) * clamp(1 - (plate(x,y) - Y)/RAMP, 0, 1)
//     monotonic and continuous, so no contour can form where it stops
//   - plate(x,y) is a local estimate (per 32x32 block, median of masked
//     pixels only, diffused, blurred, bilinearly upsampled, clamped to at most
//     PLATE_MAX_DEPTH below white) so a vignette cannot cause under-lift.
//     On a flat plate this reproduces the global-seed version block for block.
//
// usage: node process_yuv3.js in.yuv out.yuv W H FRAMES [TOL] [FULL] [CHROMA]

const fs = require('fs');

const [, , inFile, outFile, Ws, Hs, Fs, TOLs, FULLs, CHs] = process.argv;
const W = +Ws, H = +Hs, FRAMES = +Fs;
const TOL = TOLs ? +TOLs : 6;         // flood-fill depth below plate; w hits 0 here
const FULL = FULLs ? +FULLs : 3;      // at/above (plate-FULL) a pixel lifts fully
const CHROMA_ARG = CHs ? +CHs : 4;
const WHITE_Y = 235;                  // limited range (16-235) — verified per clip
const CHROMA = CHROMA_ARG;
const BLK = 32;
const PLATE_MAX_DEPTH = 6;

const ySize = W * H, cW = W >> 1, cH = H >> 1, cSize = cW * cH;
const frameSize = ySize + 2 * cSize;

const fin = fs.openSync(inFile, 'r');
const fout = fs.openSync(outFile, 'w');
const buf = Buffer.alloc(frameSize);

const mask = new Uint8Array(ySize);
const queue = new Int32Array(ySize);
const bW = Math.ceil(W / BLK), bH = Math.ceil(H / BLK);
const plateBlk = new Float32Array(bW * bH);
const plateBlk2 = new Float32Array(bW * bH);
const haveBlk = new Uint8Array(bW * bH);

let stats = { chromaTouched: 0, maxLift: 0, pxLifted: 0, frames: 0 };

function chromaNeutral(x, y) {
  const ci = (y >> 1) * cW + (x >> 1);
  const u = buf[ySize + ci], v = buf[ySize + cSize + ci];
  return Math.abs(u - 128) <= CHROMA && Math.abs(v - 128) <= CHROMA;
}

function medianOf(arr, n) {
  if (!n) return NaN;
  const a = arr.slice(0, n).sort((p, q) => p - q);
  return a[n >> 1];
}

for (let f = 0; f < FRAMES; f++) {
  const got = fs.readSync(fin, buf, 0, frameSize, null);
  if (got < frameSize) break;

  // --- global seed: median luma of the frame border -------------------------
  const border = new Uint8Array(2 * (W + H));
  let bn = 0;
  for (let x = 0; x < W; x++) { border[bn++] = buf[x]; border[bn++] = buf[(H - 1) * W + x]; }
  for (let y = 0; y < H; y++) { border[bn++] = buf[y * W]; border[bn++] = buf[y * W + W - 1]; }
  const seed = medianOf(border, bn);
  const floor = seed - TOL;

  // --- edge-connected flood fill from the frame border ----------------------
  mask.fill(0);
  let qh = 0, qt = 0;
  const push = (x, y) => {
    const i = y * W + x;
    if (mask[i]) return;
    if (buf[i] < floor) return;
    if (!chromaNeutral(x, y)) return;
    mask[i] = 1; queue[qt++] = i;
  };
  for (let x = 0; x < W; x++) { push(x, 0); push(x, H - 1); }
  for (let y = 0; y < H; y++) { push(0, y); push(W - 1, y); }
  while (qh < qt) {
    const i = queue[qh++], x = i % W, y = (i / W) | 0;
    if (x > 0) push(x - 1, y);
    if (x < W - 1) push(x + 1, y);
    if (y > 0) push(x, y - 1);
    if (y < H - 1) push(x, y + 1);
  }

  // --- local plate estimate: per-block median of MASKED pixels only ---------
  const samp = new Uint8Array(BLK * BLK);
  for (let by = 0; by < bH; by++) {
    for (let bx = 0; bx < bW; bx++) {
      let n = 0;
      const y1 = Math.min((by + 1) * BLK, H), x1 = Math.min((bx + 1) * BLK, W);
      for (let y = by * BLK; y < y1; y++) {
        for (let x = bx * BLK; x < x1; x++) {
          const i = y * W + x;
          if (mask[i]) samp[n++] = buf[i];
        }
      }
      const b = by * bW + bx;
      if (n >= 16) { plateBlk[b] = medianOf(samp, n); haveBlk[b] = 1; }
      else { plateBlk[b] = 0; haveBlk[b] = 0; }
    }
  }
  // diffuse into blocks the character covers, so they inherit a sane reference
  for (let pass = 0; pass < 64; pass++) {
    let changed = 0;
    for (let by = 0; by < bH; by++) for (let bx = 0; bx < bW; bx++) {
      const b = by * bW + bx;
      if (haveBlk[b]) continue;
      let s = 0, n = 0;
      for (let dy = -1; dy <= 1; dy++) for (let dx = -1; dx <= 1; dx++) {
        const nx = bx + dx, ny = by + dy;
        if (nx < 0 || ny < 0 || nx >= bW || ny >= bH) continue;
        const nb = ny * bW + nx;
        if (haveBlk[nb] === 1) { s += plateBlk[nb]; n++; }
      }
      if (n) { plateBlk[b] = s / n; haveBlk[b] = 2; changed++; }
    }
    for (let i = 0; i < haveBlk.length; i++) if (haveBlk[i] === 2) haveBlk[i] = 1;
    if (!changed) break;
  }
  for (let i = 0; i < plateBlk.length; i++) if (!haveBlk[i]) plateBlk[i] = seed;
  // 3x3 box blur
  for (let by = 0; by < bH; by++) for (let bx = 0; bx < bW; bx++) {
    let s = 0, n = 0;
    for (let dy = -1; dy <= 1; dy++) for (let dx = -1; dx <= 1; dx++) {
      const nx = bx + dx, ny = by + dy;
      if (nx < 0 || ny < 0 || nx >= bW || ny >= bH) continue;
      s += plateBlk[ny * bW + nx]; n++;
    }
    plateBlk2[by * bW + bx] = s / n;
  }
  // a block may never claim the plate sits more than PLATE_MAX_DEPTH below
  // white — that is what stops a block lying wholly inside a soft ground
  // shadow from declaring the shadow's falloff to be "the plate".
  for (let i = 0; i < plateBlk2.length; i++) {
    if (plateBlk2[i] < WHITE_Y - PLATE_MAX_DEPTH) plateBlk2[i] = WHITE_Y - PLATE_MAX_DEPTH;
    if (plateBlk2[i] > WHITE_Y) plateBlk2[i] = WHITE_Y;
  }

  const plateAt = (x, y) => { // bilinear upsample of the block grid
    const fx = x / BLK - 0.5, fy = y / BLK - 0.5;
    let x0 = Math.floor(fx), y0 = Math.floor(fy);
    const tx = fx - x0, ty = fy - y0;
    const cx0 = Math.max(0, Math.min(bW - 1, x0)), cx1 = Math.max(0, Math.min(bW - 1, x0 + 1));
    const cy0 = Math.max(0, Math.min(bH - 1, y0)), cy1 = Math.max(0, Math.min(bH - 1, y0 + 1));
    const p00 = plateBlk2[cy0 * bW + cx0], p10 = plateBlk2[cy0 * bW + cx1];
    const p01 = plateBlk2[cy1 * bW + cx0], p11 = plateBlk2[cy1 * bW + cx1];
    return p00 * (1 - tx) * (1 - ty) + p10 * tx * (1 - ty) + p01 * (1 - tx) * ty + p11 * tx * ty;
  };

  // --- the lift -------------------------------------------------------------
  for (let y = 0; y < H; y++) {
    for (let x = 0; x < W; x++) {
      const i = y * W + x;
      if (!mask[i]) continue;
      const Y = buf[i];
      if (Y >= WHITE_Y) { buf[i] = WHITE_Y; continue; }
      const p = plateAt(x, y);
      // two-point ramp: full lift at/above p-FULL, zero at/below p-TOL.
      // continuous and monotonic at both ends, so no contour can form; with
      // FULL = 0 and TOL = RAMP this is exactly the single-point ramp the
      // earlier parts used.
      const lo = p - TOL, hi = p - FULL;
      let w = (Y - lo) / (hi - lo);
      if (w <= 0) continue;
      if (w > 1) w = 1;
      const lift = (WHITE_Y - Y) * w;
      const nv = Math.round(Y + lift);
      if (nv !== Y) { stats.pxLifted++; if (nv - Y > stats.maxLift) stats.maxLift = nv - Y; }
      buf[i] = nv > WHITE_Y ? WHITE_Y : nv;
    }
  }

  fs.writeSync(fout, buf, 0, frameSize);
  stats.frames++;
}

fs.closeSync(fin); fs.closeSync(fout);
console.log(JSON.stringify({ ...stats, W, H, TOL, FULL, CHROMA, WHITE_Y }));
