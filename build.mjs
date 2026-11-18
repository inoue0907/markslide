// deck.md -> deck.html + deck.pptx, both from one Chrome layout.
import { readFileSync, writeFileSync, mkdirSync, mkdtempSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { basename, dirname, extname, join, resolve } from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";
import { chromium } from "playwright";
import { renderMarkdown, measurePage, viewerPage, SLIDE_W, SLIDE_H } from "./render.mjs";
import { buildPptx } from "./pptx.mjs";

const here = dirname(fileURLToPath(import.meta.url));
const MIME = { ".png": "image/png", ".jpg": "image/jpeg", ".jpeg": "image/jpeg", ".gif": "image/gif",
  ".webp": "image/webp", ".svg": "image/svg+xml", ".bmp": "image/bmp" };

/** Chrome as installed, else Playwright's own: the installed one needs no download. */
export async function launchBrowser() {
  try {
    return await chromium.launch({ channel: "chrome" });
  } catch {
    return await chromium.launch();
  }
}

/**
 * Lay the deck out in Chrome and read it back.
 * Returns { slides, sections, css, assets, notes, shots } - shots only when asked.
 */
export async function layoutDeck(markdown, { deckDir, browser, shots = false }) {
  const rendered = renderMarkdown(markdown);
  const work = mkdtempSync(join(tmpdir(), "markslide-"));
  const pagePath = join(work, "measure.html");
  writeFileSync(pagePath, measurePage(rendered, deckDir));
  const page = await browser.newPage({ viewport: { width: SLIDE_W + 40, height: SLIDE_H }, deviceScaleFactor: 2 });
  try {
    await page.goto(pathToFileURL(pagePath).href, { waitUntil: "load" });
    await page.addScriptTag({ path: join(here, "measure.js") });
    const { slides, assets } = await measureOpenPage(page);
    const sectionsEl = await page.$$("div.marpit > section");
    let shotsOut = null;
    if (shots) {
      shotsOut = [];
      for (const el of sectionsEl) shotsOut.push(await el.screenshot({ type: "png", scale: "css" }));
    }
    const sections = await page.evaluate(() =>
      [...document.querySelectorAll("div.marpit > section")].map((s) => {
        const copy = s.cloneNode(true);
        copy.querySelectorAll("[data-ms-raster]").forEach((e) => e.removeAttribute("data-ms-raster"));
        copy.querySelectorAll("img").forEach((img, i) => img.setAttribute("src", s.querySelectorAll("img")[i].src));
        return copy.outerHTML;
      }));
    const inlined = sections.map((html) => html.replace(/src="(file:[^"]+)"/g, (all, url) => {
      const data = assets.get(url.replace(/&amp;/g, "&"));
      return data ? `src="data:${data}"` : all;
    }));
    const notes = rendered.comments.map((c) => c.join("\n\n"));
    return { slides, sections: inlined, css: rendered.css, assets, notes, shots: shotsOut };
  } finally {
    await page.close();
    rmSync(work, { recursive: true, force: true });
  }
}

/**
 * Measure the slides of a page already open in Chrome with measure.js loaded: markers made
 * real, every slide read back, rasters photographed and images read. Shared by the Marp path
 * and the HTML deck, which differ only in where their sections are.
 */
export async function measureOpenPage(page) {
  await page.evaluate(() => window.__markslide.prepare());
  await page.evaluate(async () => {
    await document.fonts.ready;
    await Promise.all([...document.images].map((i) => i.decode().catch(() => {})));
  });
  const slides = await page.evaluate(() => window.__markslide.measure());
  const assets = new Map();
  for (const m of slides) {
    for (const p of m.pictures) {
      if (p.kind === "raster") {
        const el = await page.$(`[data-ms-raster="${p.id}"]`);
        const png = await el.screenshot({ omitBackground: true, type: "png" });
        assets.set(p.id, `image/png;base64,${png.toString("base64")}`);
      } else if (!assets.has(p.src)) {
        const data = await readImage(p.src);
        if (data) assets.set(p.src, data);
      }
    }
  }
  return { slides, assets };
}

export async function readImage(src) {
  try {
    if (src.startsWith("data:")) return src.slice(5);
    if (src.startsWith("file:")) {
      const path = fileURLToPath(src);
      const mime = MIME[extname(path).toLowerCase()];
      if (!mime) return null;
      return `${mime};base64,${readFileSync(path).toString("base64")}`;
    }
    const res = await fetch(src);
    if (!res.ok) return null;
    const mime = res.headers.get("content-type")?.split(";")[0] || "image/png";
    return `${mime};base64,${Buffer.from(await res.arrayBuffer()).toString("base64")}`;
  } catch {
    return null;
  }
}

/** Build both outputs for one deck file. Returns the paths written and any warnings. */
export async function buildDeck(deckPath, { outDir, html = true, pptx = true, shots = false, browser } = {}) {
  const deck = resolve(deckPath);
  const name = basename(deck, extname(deck));
  outDir = resolve(outDir || dirname(deck));
  mkdirSync(outDir, { recursive: true });
  const own = !browser;
  browser = browser || (await launchBrowser());
  try {
    const laid = await layoutDeck(readFileSync(deck, "utf8"), { deckDir: dirname(deck), browser, shots });
    const written = {};
    if (html) {
      written.html = join(outDir, `${name}.html`);
      writeFileSync(written.html, viewerPage({ css: laid.css, sections: laid.sections, title: name }));
    }
    if (pptx) {
      written.pptx = join(outDir, `${name}.pptx`);
      const p = buildPptx(laid.slides, { assets: laid.assets, notes: laid.notes, title: name });
      writeFileSync(written.pptx, await p.write({ outputType: "nodebuffer" }));
    }
    const warnings = laid.slides.flatMap((s) => s.warnings.map((w) => `slide ${s.index + 1}: ${w}`));
    return { written, warnings, laid };
  } finally {
    if (own) await browser.close();
  }
}
