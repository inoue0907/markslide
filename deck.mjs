// The HTML deck: one deck.html is the source, the viewer, the editor and what the PPTX is
// measured from. The AI writes it a page at a time (`add`, then `check`); the person opens it
// in a browser, mends words and saves it back (runtime.js); `pptx` hands it to PowerPoint.
//
// The file is <main data-theme> with one <section id data-type> per slide, then a marker,
// then what markslide writes: the theme's CSS, the viewer's CSS, the images and rendered
// formulas as JSON, and the runtime. Everything after the marker is rewritten by `sync`;
// everything before it is the deck, kept in the one form runtime.js serializes to, so a
// save in the browser and a `sync` here write the same bytes for the same slides.
import { existsSync, readFileSync, writeFileSync, mkdtempSync } from "node:fs";
import { tmpdir } from "node:os";
import { basename, dirname, extname, join, resolve } from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";
import { renderMarkdown, SLIDE_W, SLIDE_H } from "./render.mjs";
import { launchBrowser, measureOpenPage, readImage } from "./build.mjs";
import { buildPptx } from "./pptx.mjs";
import * as T from "./types.mjs";

const here = dirname(fileURLToPath(import.meta.url));
export const MARKER = "<!-- markslide:runtime ここから下は markslide が書く。手で直さない -->";

// ---- the file ---------------------------------------------------------------------------------

/** [start, end) of <main> in the text, never reaching into the runtime below the marker. */
export function mainSpan(text) {
  const cut = text.indexOf("<!-- markslide:runtime");
  const head = cut < 0 ? text : text.slice(0, cut);
  const start = head.search(/<main[\s>]/);
  const end = head.lastIndexOf("</main>");
  return start < 0 || end < start ? null : [start, end + "</main>".length];
}

export function mainText(text) {
  const span = mainSpan(text);
  if (!span) throw new Error("<main data-theme> が見つかりません");
  return text.slice(...span);
}

function jsonBlock(text, id) {
  const m = new RegExp(`<script type="application/json" id="${id}">([\\s\\S]*?)</script>`).exec(text);
  try { return m ? JSON.parse(m[1]) : {}; } catch { return {}; }
}

const escHtml = (s) => String(s).replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" })[c]);
const jsonScript = (id, value) =>
  `<script type="application/json" id="${id}">${JSON.stringify(value).replace(/</g, "\\u003c")}</script>`;

/** Marp's theme CSS, re-aimed from Marp's wrapper at the deck's <main>. */
export function themeCss(theme) {
  const { html, css } = renderMarkdown(`---\ntheme: ${theme}\nmath: mathjax\n---\n\n$x$\n`);
  if (!html.includes(`data-theme="${theme}"`)) throw new Error(`テーマ「${theme}」がありません`);
  // The cover is section.lead in Marp; here it is its page type, so the theme sees either.
  return css.replaceAll("div.marpit", "main[data-theme]")
    .replaceAll("section.lead", 'section:is(.lead,[data-type="表紙"])')
    .replaceAll(":not(.lead)", ':not(.lead,[data-type="表紙"])');
}

const VIEW_CSS = `
html { background: #444; }
body { margin: 0; }
main[data-theme] { display: flex; flex-direction: column; align-items: center; gap: 16px; padding: 16px 0 64px; }
main[data-theme] > section { flex: none; zoom: var(--ms-zoom, 1); }
main[data-theme] > section > aside.notes { display: none; }
.ms-notes-box { display: none; }
html.ms-editing .ms-notes-box {
  display: block; box-sizing: border-box; width: calc(1280px * var(--ms-zoom, 1)); min-height: 3em;
  margin-top: -8px; padding: 8px 12px; background: #fffbe6; color: #333; font: 14px/1.6 sans-serif;
  white-space: pre-wrap; outline: none;
}
html.ms-editing .ms-notes-box:empty::before { content: attr(data-placeholder); color: #999; }
html.ms-editing main [contenteditable="true"] { outline: 1px dashed rgba(11, 139, 238, 0.5); outline-offset: 2px; }
html.ms-editing main [contenteditable="true"]:focus { outline: 2px solid #0b8bee; }
.ms-toolbar {
  position: fixed; top: 12px; right: 12px; z-index: 10; display: flex; gap: 8px; align-items: center;
  padding: 6px 10px; border-radius: 6px; background: rgba(0, 0, 0, 0.75); color: #fff; font: 13px sans-serif;
}
.ms-toolbar button { font: inherit; padding: 4px 10px; cursor: pointer; }
html.ms-measure main[data-theme] { display: block; padding: 0; }
html.ms-measure main[data-theme] > section { zoom: 1; margin: 0 0 20px; }
@media print {
  html { background: none; }
  main[data-theme] { display: block; padding: 0; }
  main[data-theme] > section { zoom: 1; break-after: page; }
  .ms-toolbar, .ms-notes-box { display: none !important; }
}
`.trim();

function runtimeScript() {
  const code = readFileSync(join(here, "measure.js"), "utf8") + "\n" + readFileSync(join(here, "runtime.js"), "utf8");
  return `<script id="ms-runtime">\n${code.replace(/<\/script/gi, "<\\/script")}</script>`;
}

/** The whole file around a deck. */
export function assemble({ title, main, theme, assets = {}, math = {} }) {
  return [
    "<!doctype html>", '<html lang="ja">', "<head>", '<meta charset="utf-8">',
    '<meta name="viewport" content="width=device-width">', `<title>${escHtml(title)}</title>`,
    "</head>", "<body>", main, MARKER,
    `<style id="ms-theme">${themeCss(theme)}</style>`,
    `<style id="ms-view">\n${VIEW_CSS}\n</style>`,
    jsonScript("ms-assets", assets), jsonScript("ms-math", math), runtimeScript(),
    "</body>", "</html>", "",
  ].join("\n");
}

// ---- formulas -----------------------------------------------------------------------------------

/** TeX -> MathJax SVG, through Marp so it is the same MathJax a Marp deck drew. Keys "I:"/"D:". */
export function renderMath(keys) {
  const piece = (k) => (k.startsWith("D:") ? `$$\n${k.slice(2)}\n$$` : `$${k.slice(2)}$`);
  const pick = (md) => renderMarkdown(`---\nmath: mathjax\n---\n\n${md}\n`).html
    .match(/<mjx-container[\s\S]*?<\/mjx-container>/g) || [];
  const out = {};
  if (!keys.length) return out;
  let all = pick(keys.map(piece).join("\n\n"));
  if (all.length !== keys.length) all = keys.map((k) => pick(piece(k))[0] || "");
  keys.forEach((k, i) => { if (all[i]) out[k] = all[i]; });
  return out;
}

// ---- Chrome -----------------------------------------------------------------------------------

async function withBrowser(browser, fn) {
  const own = !browser;
  browser = browser || (await launchBrowser());
  try { return await fn(browser); } finally { if (own) await browser.close(); }
}

/** The deck's <main> in the one form runtime.js writes, and the formulas it holds. */
async function normalize(main, browser) {
  const page = await browser.newPage();
  try {
    await page.setContent(`<!doctype html><html><head><meta charset="utf-8"></head><body>${main}</body></html>`);
    await page.evaluate(() => { window.__msNoBoot = true; });
    await page.addScriptTag({ path: join(here, "runtime.js") });
    return await page.evaluate(() => ({
      main: window.__ms.serialize(),
      maths: [...document.querySelectorAll("main[data-theme] .math")]
        .map((m) => (m.tagName === "DIV" ? "D:" : "I:") + m.textContent.trim()),
      assets: [...document.querySelectorAll("main[data-theme] img[data-asset]")].map((i) => i.getAttribute("data-asset")),
      theme: document.querySelector("main[data-theme]").getAttribute("data-theme"),
    }));
  } finally {
    await page.close();
  }
}

/** Rewrite the file in its normal form: deck tidied, formulas rendered, runtime current. */
export async function sync(path, { browser, text } = {}) {
  path = resolve(path);
  text = text ?? readFileSync(path, "utf8");
  return withBrowser(browser, async (b) => {
    const n = await normalize(mainText(text), b);
    const oldMath = jsonBlock(text, "ms-math");
    const oldAssets = jsonBlock(text, "ms-assets");
    const keys = [...new Set(n.maths)];
    const fresh = renderMath(keys.filter((k) => !oldMath[k]));
    const math = Object.fromEntries(keys.map((k) => [k, oldMath[k] || fresh[k]]).filter(([, v]) => v));
    const title = /<title>([\s\S]*?)<\/title>/.exec(text)?.[1] ?? basename(path, extname(path));
    const out = assemble({ title: unescHtml(title), main: n.main, theme: n.theme || "academic", assets: oldAssets, math });
    if (out !== text) writeFileSync(path, out);
    const missing = [...new Set(n.assets.filter((a) => !oldAssets[a]))];
    return { changed: out !== text, missing };
  });
}

const unescHtml = (s) => s.replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"').replace(/&amp;/g, "&");

/** Open the deck as markslide measures it: unzoomed, no editing chrome, fonts and images in. */
async function openDeck(path, browser) {
  const page = await browser.newPage({ viewport: { width: SLIDE_W + 40, height: SLIDE_H }, deviceScaleFactor: 2 });
  await page.addInitScript(() => { window.__msMeasure = true; });
  await page.goto(pathToFileURL(resolve(path)).href, { waitUntil: "load" });
  await page.waitForFunction(() => window.__msReady === true);
  await page.evaluate(() => Promise.all([...document.images].map((i) => i.decode().catch(() => {}))));
  return page;
}

// ---- commands -------------------------------------------------------------------------------

export async function newDeck(path, { title, theme = "academic", browser } = {}) {
  path = resolve(path);
  if (existsSync(path)) throw new Error(`${path} はもうあります`);
  const main = `<main data-theme="${theme}" data-paginate>\n${T.template("表紙", "s1")}\n</main>`;
  writeFileSync(path, assemble({ title: title || basename(path, extname(path)), main, theme }));
  return sync(path, { browser });
}

/** Insert a page skeleton after `after` (or at the end) and return its id and markup. */
export function addPage(path, { type, after } = {}) {
  if (!T.LIMITS[type]) throw new Error(`型「${type}」はありません（${T.TYPES.join("、")}）`);
  const text = readFileSync(path, "utf8");
  const main = mainText(text);
  const ids = [...main.matchAll(/<section[^>]*\sid="s(\d+)"/g)].map((m) => Number(m[1]));
  const id = `s${Math.max(0, ...ids) + 1}`;
  const page = T.template(type, id);
  let next;
  if (after) {
    const open = new RegExp(`<section[^>]*\\sid="${after}"`).exec(main);
    if (!open) throw new Error(`${after} という頁はありません`);
    const close = main.indexOf("</section>", open.index) + "</section>".length;
    next = main.slice(0, close) + "\n" + page + main.slice(close);
  } else {
    const end = main.lastIndexOf("</main>");
    next = main.slice(0, end) + page + "\n" + main.slice(end);
  }
  const [s, e] = mainSpan(text);
  writeFileSync(path, text.slice(0, s) + next + text.slice(e));
  return { id, page };
}

export function addAsset(path, name, image) {
  const text = readFileSync(path, "utf8");
  const mime = { ".png": "image/png", ".jpg": "image/jpeg", ".jpeg": "image/jpeg", ".gif": "image/gif",
    ".webp": "image/webp", ".svg": "image/svg+xml" }[extname(image).toLowerCase()];
  if (!mime) throw new Error(`${image} は画像として読めません`);
  const assets = jsonBlock(text, "ms-assets");
  assets[name] = `data:${mime};base64,${readFileSync(image).toString("base64")}`;
  const re = /<script type="application\/json" id="ms-assets">[\s\S]*?<\/script>/;
  writeFileSync(path, text.replace(re, () => jsonScript("ms-assets", assets)));
}

/** One slide's markup, and a picture of it when `png` names a file. */
export async function showPage(path, id, { png, browser } = {}) {
  return withBrowser(browser, async (b) => {
    const page = await openDeck(path, b);
    try {
      const html = await page.evaluate((id) => {
        const s = document.getElementById(id);
        return s && s.matches(window.__markslideSections) ? window.__ms.serializeSection(s) : null;
      }, id);
      if (html === null) throw new Error(`${id} という頁はありません`);
      if (png) await (await page.$(`[id="${id}"]`)).screenshot({ path: png, type: "png", scale: "css" });
      return { html, png };
    } finally {
      await page.close();
    }
  });
}

/** Every rule a page breaks, as one Japanese line each: `s3 背景: …`. Syncs first. */
export async function checkDeck(path, { browser, id } = {}) {
  return withBrowser(browser, async (b) => {
    const synced = await sync(path, { browser: b });
    const text = readFileSync(path, "utf8");
    const classes = new Set([...themeCss(/data-theme="([^"]+)"/.exec(mainText(text))?.[1] || "academic")
      .matchAll(/\.([A-Za-z_][\w-]*)/g)].map((m) => m[1]).concat(["notes", "math", "lead"]));
    const page = await openDeck(path, b);
    try {
      const found = await page.evaluate(checkInPage, {
        limits: T.LIMITS, allowed: [...T.ALLOWED_TAGS], classes: [...classes],
        assets: Object.keys(jsonBlock(text, "ms-assets")),
        max: { title: T.TITLE_MAX, item: T.ITEM_MAX, column: T.COLUMN_ITEM_MAX, strong: T.STRONG_MAX,
          cite: T.CITE_MAX, note: T.NOTE_MAX, symbols: T.SYMBOL_MAX },
      });
      const problems = found.filter((p) => !id || p.id === id).map((p) => `${p.id} ${p.type}: ${p.text}`);
      return { problems, pages: found.pages, missing: synced.missing };
    } finally {
      await page.close();
    }
  });
}

// Runs in the page. Counts as marp-slide-rules' check_rules.py counts a Markdown slide, and
// measures what that script could not: whether the words fit the page.
function checkInPage({ limits, allowed, classes, assets, max }) {
  const out = [];
  const allowedSet = new Set(allowed);
  const classSet = new Set(classes);
  const assetSet = new Set(assets);
  const RUNTIME = "[data-ms-runtime], .ms-marker, .ms-page";
  const sections = [...document.querySelectorAll(window.__markslideSections)];
  const seen = new Map();
  const visible = (el, drop = "") => {
    const c = el.cloneNode(true);
    c.querySelectorAll(`${RUNTIME}, sup${drop ? ", " + drop : ""}`).forEach((e) => e.remove());
    c.querySelectorAll(".math").forEach((m) => { m.textContent = "式"; });
    return c.textContent.replace(/\s/g, "");
  };
  sections.forEach((s, n) => {
    const id = s.id || `(${n + 1}枚目)`;
    const type = s.getAttribute("data-type") || "?";
    const say = (text) => out.push({ id, type, text });
    if (!s.id) say("id がありません");
    else if (seen.has(s.id)) say(`id が ${seen.get(s.id)} 枚目と同じです`);
    seen.set(s.id, n + 1);

    // Fit: everything but the bands the theme pins outside the padding.
    const r = s.getBoundingClientRect();
    const cs = getComputedStyle(s);
    const bottom = r.bottom - parseFloat(cs.paddingBottom);
    const right = r.right - parseFloat(cs.paddingRight);
    let low = -Infinity;
    let wide = -Infinity;
    for (const el of s.querySelectorAll("*")) {
      if (el.closest("blockquote, header, footer, aside, .ms-page, .ms-marker") || el.closest("svg") !== null && el.tagName !== "svg") continue;
      const b = el.getBoundingClientRect();
      if (!b.width && !b.height) continue;
      low = Math.max(low, b.bottom);
      wide = Math.max(wide, b.right);
    }
    if (low > bottom + 1) say(`下に ${Math.round(low - bottom)}px はみ出しています`);
    if (wide > right + 1) say(`右に ${Math.round(wide - right)}px はみ出しています`);
    const quote = s.querySelector(":scope > blockquote");
    if (quote && low > quote.getBoundingClientRect().top + 1) say("本文が出典の帯に重なっています");

    // Markup the theme cannot lay out, and the PPTX cannot follow.
    for (const el of s.querySelectorAll("*")) {
      if (el.closest(`${RUNTIME}, .math`)) continue;
      const tag = el.tagName.toLowerCase();
      if (!allowedSet.has(tag)) say(`<${tag}> は使えません`);
      if (el.hasAttribute("style")) say(`<${tag}> に style 属性があります（テーマのクラスを使う）`);
      for (const c of el.classList) if (!classSet.has(c)) say(`クラス「${c}」はテーマにありません`);
      const a = el.getAttribute("data-asset");
      if (a !== null && !assetSet.has(a)) say(`画像「${a}」がありません（markslide asset で入れる）`);
    }
    if (s.hasAttribute("style")) say("<section> に style 属性があります");

    const limit = limits[type];
    if (!limit) {
      say(s.hasAttribute("data-type") ? `型「${type}」は決まりにありません` : "data-type（頁の型）がありません");
      return;
    }
    const [bodyMax, itemsMax, itemMax, figsMax, eqsMax] = limit;
    const notes = s.querySelector(":scope > aside.notes");
    if (notes && notes.textContent.replace(/\s/g, "").length > max.note) {
      say(`話す原稿が ${notes.textContent.replace(/\s/g, "").length} 字です (上限 ${max.note})`);
    }
    const body = [...s.children].filter((c) => !c.matches(`${RUNTIME}, aside, blockquote, header, footer`));
    const title = body.find((c) => /^H[12]$/.test(c.tagName));
    if (title && type !== "表紙" && visible(title).length > max.title) {
      say(`タイトルが ${visible(title).length} 字です (上限 ${max.title})`);
    }
    const inBody = (sel) => body.flatMap((c) => [...(c.matches(sel) ? [c] : []), ...c.querySelectorAll(sel)]);
    const equations = inBody("div.math").length;
    if (equations > eqsMax) say(`式が ${equations} 本です (${type}の上限 ${eqsMax})`);
    const figures = inBody("img, video").length;
    if (figures > figsMax) say(`図が ${figures} 枚です (${type}の上限 ${figsMax})`);
    const symbols = inBody(".sym .math").length;
    if (symbols > max.symbols) say(`記号の一覧が ${symbols} 個です (上限 ${max.symbols})`);
    const cites = quote ? (quote.textContent.match(/\[\d+\]/g) || []).length : 0;
    if (cites > max.cite) say(`出どころの帯に引用が ${cites} 件あります (上限 ${max.cite})`);
    const strong = inBody("strong").filter((e) => !e.closest("h1, h2, h3, h4"));
    if (strong.length > max.strong) say(`赤い太字が ${strong.length} 個あります (上限 ${max.strong})`);
    for (const e of strong) if (/[。．]$/.test(e.textContent.trim())) say(`句点が強調の中にあります: ${e.textContent.trim().slice(0, 20)}…`);
    const items = inBody("li").filter((li) => !li.parentElement.closest("li"));
    if (items.length > itemsMax) say(`項目が ${items.length} 個です (${type}の上限 ${itemsMax})`);
    for (const li of items) {
      const t = visible(li, "ul, ol");
      const lim = itemMax || (li.closest(".flex") ? max.column : max.item);
      if (t.length > lim) say(`項目が ${t.length} 字です (上限 ${lim}): ${t.slice(0, 20)}…`);
    }
    if (bodyMax !== null) {
      const chars = body.filter((c) => c !== title && !c.matches(".sym")).map((c) => visible(c, ".sym").length)
        .reduce((a, b) => a + b, 0);
      if (chars > bodyMax) say(`本文が ${chars} 字です (${type}の上限 ${bodyMax})`);
    }
    // Digits read as a table squeezed into prose: the rules want them in a figure or a table.
    const walker = document.createTreeWalker(s, NodeFilter.SHOW_TEXT);
    if (type !== "表紙") for (let t = walker.nextNode(); t; t = walker.nextNode()) {
      const p = t.parentElement;
      if (p.closest(`${RUNTIME}, aside, blockquote, header, footer, sup, .math, table, ol, code, .cap`) || p.closest("h1, h2") === title && title) continue;
      if (/[0-9０-９]/.test(t.data)) { say(`本文に数字があります: ${t.data.trim().slice(0, 20)}…`); break; }
    }
  });
  out.pages = sections.length;
  return out;
}

/** The PPTX, measured from the deck as Chrome lays it out; notes from each page's aside. */
export async function deckToPptx(path, { out, browser } = {}) {
  path = resolve(path);
  out = resolve(out || join(dirname(path), `${basename(path, extname(path))}.pptx`));
  return withBrowser(browser, async (b) => {
    const page = await openDeck(path, b);
    try {
      const notes = await page.evaluate(() => [...document.querySelectorAll(window.__markslideSections)].map((s) => {
        const a = s.querySelector(":scope > aside.notes");
        const t = a ? a.textContent.trim() : "";
        if (a) a.remove();
        return t;
      }));
      const { slides, assets } = await measureOpenPage(page);
      const title = await page.title();
      const p = buildPptx(slides, { assets, notes, title });
      writeFileSync(out, await p.write({ outputType: "nodebuffer" }));
      const warnings = slides.flatMap((s) => s.warnings.map((w) => `slide ${s.index + 1}: ${w}`));
      return { written: out, warnings, slides: slides.length };
    } finally {
      await page.close();
    }
  });
}

// ---- from a Marp deck -----------------------------------------------------------------------

/** Formulas out of the Markdown before Marp sees them, so the deck holds TeX, not SVG. */
function stashMath(md) {
  const b64 = (s) => Buffer.from(s, "utf8").toString("base64");
  const parts = md.split(/(^```[\s\S]*?^```|`[^`\n]*`)/m);
  return parts.map((part, i) => (i % 2 ? part : part
    .replace(/^\$\$\s*\n([\s\S]*?)\n\s*\$\$\s*$/gm, (_, tex) => `\n<div class="math" data-ms-b64="${b64(tex.trim())}"></div>\n`)
    .replace(/(?<![\\$])\$(?!\s)([^$\n]+?)(?<!\s)\$/g, (_, tex) => `<span class="math" data-ms-b64="${b64(tex)}"></span>`)))
    .join("");
}

/** A Marp deck.md -> deck.html. Returns the path and anything that did not carry over. */
export async function importMarkdown(mdPath, { out, browser } = {}) {
  mdPath = resolve(mdPath);
  out = resolve(out || join(dirname(mdPath), `${basename(mdPath, extname(mdPath))}.html`));
  const source = readFileSync(mdPath, "utf8");
  const rendered = renderMarkdown(stashMath(source));
  const html = rendered.html.replace(/data-ms-b64="([^"]*)"><\/(span|div)>/g,
    (_, b, tag) => `>${escHtml(Buffer.from(b, "base64").toString("utf8"))}</${tag}>`);
  const comments = rendered.comments;
  return withBrowser(browser, async (b) => {
    const page = await b.newPage();
    try {
      const base = pathToFileURL(dirname(mdPath) + "/").href;
      await page.setContent(`<!doctype html><html><head><meta charset="utf-8"><base href="${base}"></head><body>${html}</body></html>`);
      const got = await page.evaluate(({ comments }) => {
        const warnings = [];
        const sections = [...document.querySelectorAll("div.marpit > section")];
        const first = sections[0];
        const main = document.createElement("main");
        main.setAttribute("data-theme", first?.getAttribute("data-theme") || "academic");
        const paginate = sections.some((s) => s.hasAttribute("data-marpit-pagination"));
        if (paginate) main.setAttribute("data-paginate", "");
        const header = first?.querySelector(":scope > header")?.textContent.trim();
        if (header) main.setAttribute("data-header", header);
        const images = new Map();
        sections.forEach((s, i) => {
          const sec = document.createElement("section");
          sec.id = `s${i + 1}`;
          const notes = [];
          let type = s.classList.contains("lead") ? "表紙" : "";
          for (const c of comments[i] || []) {
            const m = /^\s*型\s*[:：]\s*(\S+)\s*$/.exec(c);
            if (m) type = m[1]; else if (c.trim()) notes.push(c.trim());
          }
          if (type) sec.setAttribute("data-type", type);
          if (paginate && !s.hasAttribute("data-marpit-pagination")) sec.setAttribute("data-paginate", "false");
          for (const c of s.classList) if (c !== "lead") sec.classList.add(c);
          if (s.querySelector("style")) warnings.push(`${sec.id}: <style> は持ち込めないので外しました`);
          s.querySelectorAll(":scope > header, :scope > footer, style").forEach((e) => e.remove());
          for (const el of s.querySelectorAll("*")) {
            if (/^H\d$/.test(el.tagName)) el.removeAttribute("id");
            if (el.tagName === "IMG") {
              const src = el.src;
              if (!images.has(src)) images.set(src, `fig${images.size + 1}`);
              el.setAttribute("data-asset", images.get(src));
              el.removeAttribute("src");
              const w = /width:\s*([\d.]+)px/.exec(el.getAttribute("style") || "");
              const h = /height:\s*([\d.]+)px/.exec(el.getAttribute("style") || "");
              if (w) el.setAttribute("width", w[1]);
              if (h) el.setAttribute("height", h[1]);
              el.removeAttribute("style");
            }
            if (el.hasAttribute("style")) {
              warnings.push(`${sec.id}: <${el.tagName.toLowerCase()}> の style を外しました`);
              el.removeAttribute("style");
            }
          }
          sec.append(...s.childNodes);
          if (notes.length) {
            const a = document.createElement("aside");
            a.className = "notes";
            a.textContent = notes.join("\n\n");
            sec.append(a);
          }
          main.append(sec);
        });
        return { main: main.outerHTML, images: [...images].map(([src, name]) => ({ src, name })), warnings };
      }, { comments });
      const assets = {};
      for (const { src, name } of got.images) {
        const data = await readImage(src);
        if (data) assets[name] = `data:${data}`;
        else got.warnings.push(`画像 ${src} を読めませんでした`);
      }
      const title = /^title:\s*(.+)$/m.exec(source)?.[1]?.trim() || basename(out, extname(out));
      const theme = /data-theme="([^"]+)"/.exec(got.main)?.[1] || "academic";
      const text = assemble({ title, main: got.main, theme, assets });
      writeFileSync(out, text);
      await sync(out, { browser: b, text });
      return { written: out, warnings: got.warnings };
    } finally {
      await page.close();
    }
  });
}

export { T as types, mkdtempSync, tmpdir };
