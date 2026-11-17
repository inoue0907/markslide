// Markdown -> one HTML document holding every slide, laid out for measuring.
//
// The document is the single source of truth: Chrome lays it out, markslide measures that
// layout for the PPTX, and the same DOM (after measure.js has made pseudo-elements real) is
// what the HTML output saves. Two renderings that never meet cannot be kept identical.
import { readdirSync, readFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";
import { Marp } from "@marp-team/marp-core";

const here = dirname(fileURLToPath(import.meta.url));
export const TEMPLATE_DIR = join(here, "templates");
export const SLIDE_W = 1280;
export const SLIDE_H = 720;

// A template is CSS in its body, but it is not a stylesheet: it is what the PPTX is drawn
// from too, so only the properties a PowerPoint slide can draw belong in it. Its own
// extension keeps it from being mistaken for (or linked as) an ordinary .css file.
export const TEMPLATE_EXT = ".mst";

export function templateFiles() {
  return readdirSync(TEMPLATE_DIR).filter((f) => f.endsWith(TEMPLATE_EXT)).sort();
}

export function loadTemplates() {
  return templateFiles().map((f) => readFileSync(join(TEMPLATE_DIR, f), "utf8"));
}

/** Render deck Markdown with marp-core and the markslide templates. */
export function renderMarkdown(markdown) {
  // inlineSVG off: each <section> is a plain 1280x720 box on the page, so its
  // getBoundingClientRect is the slide's own coordinate system with no scaling between.
  const marp = new Marp({ html: true, script: false, inlineSVG: false, math: "mathjax" });
  for (const css of loadTemplates()) marp.themeSet.add(css);
  const { html, css, comments } = marp.render(markdown);
  return { html, css, comments };
}

/** The page Chrome opens: slides stacked vertically, base href at the deck's folder. */
export function measurePage({ html, css }, deckDir) {
  const base = pathToFileURL(deckDir.endsWith("/") || deckDir.endsWith("\\") ? deckDir : deckDir + "/").href;
  return `<!doctype html>
<html lang="ja"><head><meta charset="utf-8"><base href="${base}">
<style>${css}</style>
<style>
  html, body { margin: 0; padding: 0; background: #888; }
  div.marpit > section { margin: 0 0 20px; }
</style></head>
<body>${html}</body></html>`;
}

/** The HTML output: the measured slides, a page that scales them to the window and prints one per sheet. */
export function viewerPage({ css, sections, title }) {
  return `<!doctype html>
<html lang="ja"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width">
<title>${escapeHtml(title)}</title>
<meta name="generator" content="markslide">
<style>${css}</style>
<style>
  html, body { margin: 0; background: #444; }
  .ms-deck { display: flex; flex-direction: column; align-items: center; gap: 16px; padding: 16px 0; }
  .ms-slide { width: min(100vw - 32px, (100vh - 32px) * 16 / 9); aspect-ratio: 16 / 9; position: relative; overflow: hidden; }
  .ms-slide > div.marpit { position: absolute; top: 0; left: 0; width: 1280px; height: 720px;
    transform-origin: 0 0; transform: scale(var(--ms-scale, 1)); }
  @page { size: 1280px 720px; margin: 0; }
  @media print {
    html, body { background: none; }
    .ms-deck { display: block; padding: 0; }
    .ms-slide { width: 1280px; height: 720px; break-after: page; }
    .ms-slide > div.marpit { transform: none; }
  }
</style></head>
<body><div class="ms-deck">
${sections.map((s) => `<div class="ms-slide"><div class="marpit">${s}</div></div>`).join("\n")}
</div>
<script>
  // Scale each 1280x720 slide to its frame. A script, not CSS, because a length cannot be
  // divided by a length in calc().
  const fit = () => document.querySelectorAll(".ms-slide").forEach((f) =>
    f.style.setProperty("--ms-scale", f.clientWidth / 1280));
  addEventListener("resize", fit); fit();
</script>
</body></html>`;
}

function escapeHtml(s) {
  return String(s).replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" })[c]);
}
