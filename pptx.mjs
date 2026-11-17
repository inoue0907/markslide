// Measured slides -> PPTX (pptxgenjs).
//
// Every position here is one Chrome produced; nothing is laid out a second time. Text keeps
// Chrome's line breaks as explicit breaks with exact line spacing and no wrapping, so the
// only thing PowerPoint still decides is where a glyph sits inside its line - which is
// calibrated below rather than guessed.
import pptxgen from "pptxgenjs";

const PX_PER_IN = 96;
const PT_PER_PX = 0.75;
export const inch = (px) => px / PX_PER_IN;
export const pt = (px) => Math.round(px * PT_PER_PX * 100) / 100;

/**
 * Where PowerPoint puts the baseline of a line with exact spacing, measured from the top of
 * that line, in px: pitch x pitch + size x size + offset. Found by the fidelity test
 * (`markslide fidelity`), not derived: with Yu Gothic at 16-48 px and line pitches of 21-65 px,
 * 0.8 x pitch alone left every line 1.9 px too high, at every size and pitch alike, so the
 * correction is a constant rather than a share of either. Re-measured with Noto Sans JP:
 * every line still within 1 px, mean +0.2 px, so the constant held across the two faces.
 */
export const CALIBRATION = { pitch: 0.8, size: 0, offset: -1.9 };
export function baselineInLine(pitchPx, sizePx) {
  return CALIBRATION.pitch * pitchPx + CALIBRATION.size * sizePx + CALIBRATION.offset;
}

/**
 * The face PowerPoint should ask for: the one measure.js found Chrome drawing (`face`), else
 * the CSS list's head. Japanese and generic names become the face's Windows name.
 */
export function fontFace(cssFamily, face) {
  const first = face || String(cssFamily).split(",")[0].trim().replace(/^["']|["']$/g, "");
  const map = { "游ゴシック": "Yu Gothic", YuGothic: "Yu Gothic", monospace: "Consolas", "sans-serif": "Noto Sans JP" };
  return map[first] || first;
}

function runOptions(style) {
  const o = {
    fontFace: fontFace(style.font, style.face),
    fontSize: pt(style.size),
    bold: style.bold,
    italic: style.italic,
    color: style.color.hex,
  };
  if (style.color.alpha < 1) o.transparency = Math.round((1 - style.color.alpha) * 100);
  if (style.underline) o.underline = { style: "sng" };
  if (style.strike) o.strike = "sngStrike";
  if (style.va === "super") o.superscript = true;
  if (style.va === "sub") o.subscript = true;
  if (style.spacing) o.charSpacing = pt(style.spacing);
  return o;
}

const ALIGN = { left: "left", start: "left", center: "center", right: "right", end: "right", justify: "left" };

/**
 * One measured text block -> addText calls. A block whose every line is one stretch of text
 * becomes one text box (editable as a paragraph run); a line broken by an inline picture or
 * an inline-block's margin is placed stretch by stretch, since PowerPoint could not reproduce
 * the gap.
 */
export function textBoxes(block) {
  const align = ALIGN[block.align] || "left";
  const lines = block.lines;
  if (!lines.length) return [];
  const pitches = lines.map((l, i) =>
    i + 1 < lines.length ? lines[i + 1].baseline - l.baseline : block.lineHeight || (l.bottom - l.top));
  const simple = lines.every((l) => l.segments.length === 1);
  if (simple) {
    const left = Math.min(...lines.map((l) => l.segments[0].x));
    const right = Math.max(...lines.map((l) => l.segments[0].x + l.segments[0].w));
    let x = left;
    let w = right - left;
    if (align !== "left") { x = block.box.x; w = Math.max(block.box.w, w); }
    // Slack against PowerPoint's font metrics running a hair wider than Chrome's, added on
    // the side the alignment leaves free: +8 on the right of a right-aligned box pushed
    // every right-aligned line 8 px right of Chrome's.
    const slack = 8;
    x -= align === "right" ? slack : align === "center" ? slack / 2 : 0;
    const first = lines[0];
    const top = first.baseline - baselineInLine(pitches[0], first.size);
    const runs = [];
    lines.forEach((line, i) => {
      const r = line.segments[0].runs;
      r.forEach((run, j) => runs.push({
        text: run.text,
        options: {
          ...runOptions(run.style),
          align,
          lineSpacing: pt(pitches[i]),
          ...(j === r.length - 1 && i < lines.length - 1 ? { breakLine: true } : {}),
        },
      }));
    });
    const height = pitches.reduce((a, b) => a + b, 0);
    return [{ runs, x, y: top, w: w + slack, h: height, align }];
  }
  const out = [];
  lines.forEach((line, i) => {
    const top = line.baseline - baselineInLine(pitches[i], line.size);
    for (const seg of line.segments) {
      out.push({
        runs: seg.runs.map((run) => ({
          text: run.text,
          options: { ...runOptions(run.style), align: "left", lineSpacing: pt(pitches[i]) },
        })),
        x: seg.x, y: top, w: seg.w + 8, h: pitches[i], align: "left",
      });
    }
  });
  return out;
}

/**
 * Build the presentation. `slides` are measure.js's output; `assets` maps a raster id or an
 * image src to a data URL ("image/png;base64,..."); `notes[i]` is slide i's speaker notes.
 */
export function buildPptx(slides, { assets, notes = [], title = "markslide" }) {
  const pptx = new pptxgen();
  pptx.defineLayout({ name: "MARKSLIDE_16x9", width: inch(1280), height: inch(720) });
  pptx.layout = "MARKSLIDE_16x9";
  pptx.title = title;
  for (const m of slides) {
    const slide = pptx.addSlide();
    slide.background = { color: m.background.hex };
    for (const s of m.shapes) {
      if (s.kind === "rect") {
        const opts = {
          x: inch(s.x), y: inch(s.y), w: inch(s.w), h: inch(s.h),
          fill: { color: s.fill.hex, ...(s.fill.alpha < 1 ? { transparency: Math.round((1 - s.fill.alpha) * 100) } : {}) },
          line: { type: "none" },
        };
        if (s.radius > 0) slide.addShape(pptx.ShapeType.roundRect, { ...opts, rectRadius: inch(s.radius) });
        else slide.addShape(pptx.ShapeType.rect, opts);
      } else if (s.kind === "ellipse") {
        slide.addShape(pptx.ShapeType.ellipse, {
          x: inch(s.x), y: inch(s.y), w: inch(s.w), h: inch(s.h),
          fill: s.fill ? { color: s.fill.hex } : { type: "none" },
          line: s.line ? { color: s.line.color.hex, width: pt(s.line.width) } : { type: "none" },
        });
      } else if (s.kind === "line") {
        const x = Math.min(s.x1, s.x2);
        const y = Math.min(s.y1, s.y2);
        slide.addShape(pptx.ShapeType.line, {
          x: inch(x), y: inch(y), w: inch(Math.abs(s.x2 - s.x1)), h: inch(Math.abs(s.y2 - s.y1)),
          line: {
            color: s.color.hex,
            width: pt(s.width),
            dashType: s.dash,
            ...(s.color.alpha < 1 ? { transparency: Math.round((1 - s.color.alpha) * 100) } : {}),
          },
        });
      }
    }
    for (const p of m.pictures) {
      const data = assets.get(p.kind === "raster" ? p.id : p.src);
      if (!data) continue;
      slide.addImage({ data, x: inch(p.x), y: inch(p.y), w: inch(p.w), h: inch(p.h) });
    }
    for (const block of m.texts) {
      for (const box of textBoxes(block)) {
        slide.addText(box.runs, {
          x: inch(box.x), y: inch(box.y), w: inch(box.w), h: inch(box.h),
          margin: 0, valign: "top", wrap: false, fit: "none", align: box.align,
          paraSpaceBefore: 0, paraSpaceAfter: 0,
        });
      }
    }
    if (notes[m.index]) slide.addNotes(notes[m.index]);
  }
  return pptx;
}
