// How far the PPTX is from the HTML, measured rather than eyeballed.
//
// Chrome's picture of each slide is the reference; PowerPoint itself (over COM, so Windows
// with PowerPoint installed) draws the PPTX at the same 1280x720. Two numbers come out per
// slide: the share of pixels that differ, and for every text line how many px PowerPoint put
// it below (+) or above (-) Chrome. The second is what CALIBRATION in pptx.mjs is tuned from.
import { execFileSync } from "node:child_process";
import { mkdirSync, readFileSync, writeFileSync } from "node:fs";
import { basename, dirname, extname, join, resolve } from "node:path";
import { buildDeck, launchBrowser } from "./build.mjs";

/** Export every slide of a PPTX to slide-N.png with PowerPoint. Returns the paths. */
export function exportWithPowerPoint(pptxPath, outDir, prefix = "pptx") {
  const script = `
$ErrorActionPreference = 'Stop'
$app = New-Object -ComObject PowerPoint.Application
try {
  $deck = $app.Presentations.Open('${pptxPath.replace(/'/g, "''")}', -1, 0, 0)
  $i = 0
  foreach ($s in $deck.Slides) {
    $i++
    $s.Export((Join-Path '${outDir.replace(/'/g, "''")}' ('${prefix}-' + $i + '.png')), 'PNG', 1280, 720)
  }
  $deck.Close()
  Write-Output $i
} finally { $app.Quit() }`;
  const out = execFileSync("powershell.exe", ["-NoProfile", "-NonInteractive", "-Command", script],
    { encoding: "utf8" });
  const n = Number(out.trim().split(/\s+/).pop());
  return Array.from({ length: n }, (_, i) => join(outDir, `${prefix}-${i + 1}.png`));
}

// Runs in the page: both pictures -> differing share, a diff picture and each line's shift.
function compare({ a, b, lines, threshold }) {
  const load = (src) => new Promise((ok, no) => {
    const img = new Image();
    img.onload = () => ok(img);
    img.onerror = no;
    img.src = src;
  });
  const pixels = (img) => {
    const c = document.createElement("canvas");
    c.width = 1280; c.height = 720;
    const g = c.getContext("2d");
    g.fillStyle = "#fff"; g.fillRect(0, 0, 1280, 720);
    g.drawImage(img, 0, 0, 1280, 720);
    return g.getImageData(0, 0, 1280, 720).data;
  };
  return Promise.all([load(a), load(b)]).then(([ia, ib]) => {
    const pa = pixels(ia);
    const pb = pixels(ib);
    const c = document.createElement("canvas");
    c.width = 1280; c.height = 720;
    const g = c.getContext("2d");
    const out = g.createImageData(1280, 720);
    let differ = 0;
    for (let i = 0; i < pa.length; i += 4) {
      const d = Math.max(Math.abs(pa[i] - pb[i]), Math.abs(pa[i + 1] - pb[i + 1]), Math.abs(pa[i + 2] - pb[i + 2]));
      const grey = 255 - (pa[i] + pa[i + 1] + pa[i + 2]) / 12;
      if (d > threshold) {
        differ++;
        out.data[i] = 230; out.data[i + 1] = 30; out.data[i + 2] = 30;
      } else {
        out.data[i] = out.data[i + 1] = out.data[i + 2] = grey;
      }
      out.data[i + 3] = 255;
    }
    g.putImageData(out, 0, 0);
    // Ink per row inside a line's box, against the background at that box's left edge.
    const profile = (p, x0, x1, y0, y1) => {
      const rows = [];
      const bi = (Math.max(0, y0) * 1280 + Math.max(0, x0 - 2)) * 4;
      const bg = [p[bi], p[bi + 1], p[bi + 2]];
      for (let y = y0; y < y1; y++) {
        let ink = 0;
        if (y >= 0 && y < 720) {
          for (let x = Math.max(0, x0); x < Math.min(1280, x1); x++) {
            const k = (y * 1280 + x) * 4;
            ink += Math.max(Math.abs(p[k] - bg[0]), Math.abs(p[k + 1] - bg[1]), Math.abs(p[k + 2] - bg[2]));
          }
        }
        rows.push(ink);
      }
      return rows;
    };
    const R = 14;
    const shifts = lines.map((l) => {
      const x0 = Math.floor(l.x), x1 = Math.ceil(l.x + l.w);
      const y0 = Math.floor(l.top) - R, y1 = Math.ceil(l.bottom) + R;
      const ref = profile(pa, x0, x1, y0, y1);
      const got = profile(pb, x0, x1, y0, y1);
      const total = ref.reduce((s, v) => s + v, 0);
      if (!total) return { ...l, dy: null };
      let best = 0, bestErr = Infinity;
      for (let dy = -R; dy <= R; dy++) {
        let err = 0;
        for (let y = R; y < ref.length - R; y++) err += Math.abs(ref[y] - got[y + dy]);
        if (err < bestErr) { bestErr = err; best = dy; }
      }
      // Horizontal: the first and last inked column, Chrome against PowerPoint. Rows strictly
      // inside the line and a narrow margin: a wider window caught a blockquote's dashed rule
      // and a list's number and reported them as a 22 px shift.
      const cols = (p) => {
        let first = -1, last = -1;
        const yt = Math.ceil(l.top) + 1, yb = Math.floor(l.bottom) - 1;
        const bi = (yt * 1280 + Math.max(0, x0 - 2)) * 4;
        for (let x = Math.max(0, x0 - 6); x < Math.min(1280, x1 + 30); x++) {
          let ink = 0;
          for (let y = yt; y < yb; y++) {
            const k = (y * 1280 + x) * 4;
            ink += Math.abs(p[k] - p[bi]) + Math.abs(p[k + 1] - p[bi + 1]) + Math.abs(p[k + 2] - p[bi + 2]);
          }
          if (ink > 200) { if (first < 0) first = x; last = x; }
        }
        return [first, last];
      };
      const [ra, rb] = cols(pa);
      const [ga, gb] = cols(pb);
      return { ...l, dy: best, dx: ga - ra, dw: (gb - ga) - (rb - ra), fit: bestErr / total };
    });
    return { differ: differ / (1280 * 720), diff: c.toDataURL("image/png"), shifts };
  });
}

const dataUrl = (buf) => `data:image/png;base64,${buf.toString("base64")}`;

/** Build the deck, draw it both ways and compare. Writes chrome-N, pptx-N and diff-N PNGs. */
export async function fidelity(deckPath, { outDir, threshold = 48 } = {}) {
  const deck = resolve(deckPath);
  outDir = resolve(outDir || join(dirname(deck), `${basename(deck, extname(deck))}-fidelity`));
  mkdirSync(outDir, { recursive: true });
  const browser = await launchBrowser();
  try {
    const { written, laid } = await buildDeck(deck, { outDir, shots: true, browser });
    const pptPngs = exportWithPowerPoint(written.pptx, outDir);
    const page = await browser.newPage();
    const slides = [];
    for (const m of laid.slides) {
      const chromePng = laid.shots[m.index];
      writeFileSync(join(outDir, `chrome-${m.index + 1}.png`), chromePng);
      const lines = m.texts.flatMap((t) => t.lines.map((l) => ({
        text: l.segments.map((s) => s.runs.map((r) => r.text).join("")).join(" ").slice(0, 40),
        size: l.size,
        pitch: t.lineHeight,
        x: Math.min(...l.segments.map((s) => s.x)),
        w: Math.max(...l.segments.map((s) => s.x + s.w)) - Math.min(...l.segments.map((s) => s.x)),
        top: l.top, bottom: l.bottom, baseline: l.baseline,
      })));
      const r = await page.evaluate(compare, {
        a: dataUrl(chromePng), b: dataUrl(readFileSync(pptPngs[m.index])), lines, threshold,
      });
      writeFileSync(join(outDir, `diff-${m.index + 1}.png`), Buffer.from(r.diff.split(",")[1], "base64"));
      slides.push({ slide: m.index + 1, differ: r.differ, lines: r.shifts });
    }
    const measured = slides.flatMap((s) => s.lines).filter((l) => l.dy !== null);
    const mean = (xs) => (xs.length ? xs.reduce((a, b) => a + b, 0) / xs.length : 0);
    const summary = {
      outDir,
      slides: slides.map((s) => ({ slide: s.slide, differPct: +(s.differ * 100).toFixed(2) })),
      lines: measured.length,
      meanDy: +mean(measured.map((l) => l.dy)).toFixed(2),
      maxAbsDy: Math.max(0, ...measured.map((l) => Math.abs(l.dy))),
      meanDx: +mean(measured.map((l) => l.dx)).toFixed(2),
      meanDw: +mean(measured.map((l) => l.dw)).toFixed(2),
    };
    writeFileSync(join(outDir, "report.json"), JSON.stringify({ summary, slides }, null, 2));
    return { summary, slides };
  } finally {
    await browser.close();
  }
}
