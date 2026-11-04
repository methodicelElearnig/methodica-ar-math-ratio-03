// Background-plate scanner: reports whether a clip sits on true #FFFFFF.
// Decodes every frame to rgb24 over a pipe; no temp files.
const { spawnSync, spawn } = require('child_process');
const FF = 'C:/Users/sharons/AppData/Local/CapCut/Apps/9.3.0.3970/ffmpeg.exe';

function probe(file) {
  const r = spawnSync(FF, ['-hide_banner', '-i', file, '-f', 'null', '-'], { encoding: 'buffer', maxBuffer: 1 << 26 });
  const err = r.stderr.toString('latin1');
  const m = err.match(/Video:.*?, (\d+)x(\d+)/);
  const fps = err.match(/([\d.]+) fps/);
  const dur = err.match(/Duration: (\d+):(\d+):([\d.]+)/);
  if (!m) throw new Error('no video stream in ' + file + '\n' + err.slice(-800));
  return {
    w: +m[1], h: +m[2],
    fps: fps ? +fps[1] : null,
    dur: dur ? (+dur[1] * 3600 + +dur[2] * 60 + parseFloat(dur[3])) : null,
  };
}

function ringIndices(w, h, inset) {
  const idx = [];
  const x0 = inset, x1 = w - 1 - inset, y0 = inset, y1 = h - 1 - inset;
  if (x1 <= x0 || y1 <= y0) return idx;
  for (let x = x0; x <= x1; x++) { idx.push(y0 * w + x); idx.push(y1 * w + x); }
  for (let y = y0 + 1; y < y1; y++) { idx.push(y * w + x0); idx.push(y * w + x1); }
  return idx;
}

async function scan(file) {
  const info = probe(file);
  const { w, h } = info;
  const frameBytes = w * h * 3;
  const rings = [0, 2, 4, 8, 16].map(i => ({ inset: i, idx: ringIndices(w, h, i) }));
  const stats = {
    file, ...info, frames: 0,
    ring: rings.map(r => ({ inset: r.inset, n: 0, not255: 0, nearWhiteNot255: 0, min: 255, sum: 0 })),
    nearWhiteNot255: 0,   // whole frame: all channels 235..254 -> plate-like
    pureWhite: 0,
    totalPx: 0,
    perFrameBorderMean: [],
    hist: new Uint32Array(256), // border-ring luma-ish (min channel) histogram
  };

  const ff = spawn(FF, ['-hide_banner', '-loglevel', 'error', '-i', file,
    '-f', 'rawvideo', '-pix_fmt', 'rgb24', '-vsync', '0', '-'], { stdio: ['ignore', 'pipe', 'inherit'] });

  let buf = Buffer.alloc(0);
  const chunks = [];
  await new Promise((res, rej) => {
    ff.stdout.on('data', d => {
      chunks.push(d);
      let have = chunks.reduce((a, c) => a + c.length, 0);
      if (have >= frameBytes) {
        buf = Buffer.concat([buf, ...chunks.splice(0)]);
        while (buf.length >= frameBytes) {
          processFrame(buf.subarray(0, frameBytes));
          buf = buf.subarray(frameBytes);
        }
      }
    });
    ff.on('close', c => c === 0 || buf.length === 0 ? res() : res());
    ff.on('error', rej);
  });

  function processFrame(f) {
    stats.frames++;
    stats.totalPx += w * h;
    for (let i = 0; i < w * h; i++) {
      const r = f[i * 3], g = f[i * 3 + 1], b = f[i * 3 + 2];
      if (r === 255 && g === 255 && b === 255) { stats.pureWhite++; continue; }
      if (r >= 235 && g >= 235 && b >= 235) stats.nearWhiteNot255++;
    }
    let bsum = 0, bn = 0;
    for (const rg of stats.ring) {
      const idx = rings.find(x => x.inset === rg.inset).idx;
      for (const p of idx) {
        const r = f[p * 3], g = f[p * 3 + 1], b = f[p * 3 + 2];
        const mn = Math.min(r, g, b);
        rg.n++; rg.sum += (r + g + b) / 3;
        if (mn < rg.min) rg.min = mn;
        if (!(r === 255 && g === 255 && b === 255)) {
          rg.not255++;
          if (mn >= 235) rg.nearWhiteNot255++;
        }
        if (rg.inset === 0) { stats.hist[mn]++; bsum += (r + g + b) / 3; bn++; }
      }
    }
    stats.perFrameBorderMean.push(bsum / bn);
  }
  return stats;
}

(async () => {
  const out = [];
  for (const f of process.argv.slice(2)) {
    try { out.push(await scan(f)); }
    catch (e) { out.push({ file: f, error: String(e.message) }); }
  }
  // strip the big histogram into a compact top-values list
  for (const s of out) {
    if (s.hist) {
      const top = [...s.hist].map((c, v) => [v, c]).filter(x => x[1] > 0).sort((a, b) => b[1] - a[1]).slice(0, 6);
      s.borderTopValues = top;
      delete s.hist;
      const m = s.perFrameBorderMean;
      s.borderMeanMin = Math.min(...m); s.borderMeanMax = Math.max(...m);
      delete s.perFrameBorderMean;
    }
  }
  console.log(JSON.stringify(out, null, 1));
})();
