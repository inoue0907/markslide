// Run: npm test. The build test drives Chrome (installed, else Playwright's) and is skipped
// when neither starts; the PowerPoint comparison is `npm run markslide -- fidelity <deck>`.
import { test } from "node:test"
import assert from "node:assert/strict"
import { mkdtempSync, readFileSync, rmSync } from "node:fs"
import { tmpdir } from "node:os"
import { join } from "node:path"
import { fileURLToPath } from "node:url"
import JSZip from "jszip"
import { inch, pt, fontFace, textBoxes, baselineInLine } from "../pptx.mjs"
import { renderMarkdown, templateFiles } from "../render.mjs"
import { buildDeck, launchBrowser } from "../build.mjs"

const style = { font: '"Yu Gothic", sans-serif', size: 26.6667, bold: false, italic: false,
  color: { hex: "333333", alpha: 1 }, underline: false, strike: false, va: null, spacing: 0 }
const line = (baseline, segments) => ({ baseline, top: baseline - 26, bottom: baseline + 8, size: 26.6667, segments })
const seg = (x, w, text) => ({ x, w, runs: [{ text, style }] })

test("slide px become inches at 96 dpi and points at 0.75", () => {
  assert.equal(inch(1280), 13.333333333333334)
  assert.equal(pt(26.6667), 20)
})

test("the CSS family list picks the face PowerPoint asks for", () => {
  assert.equal(fontFace('"Yu Gothic", "游ゴシック", sans-serif'), "Yu Gothic")
  assert.equal(fontFace('"游ゴシック"'), "Yu Gothic")
  assert.equal(fontFace("Consolas, monospace"), "Consolas")
  // The face Chrome drew wins over the list's head: Source Code Pro not installed drew Consolas.
  assert.equal(fontFace('"Source Code Pro", Consolas, monospace', "Consolas"), "Consolas")
})

test("a block of single-stretch lines is one text box with Chrome's breaks kept", () => {
  const block = { align: "left", box: { x: 40, w: 1200 }, lineHeight: 36,
    lines: [line(100, [seg(40, 300, "一行目")]), line(136, [seg(40, 200, "二行目")])] }
  const boxes = textBoxes(block)
  assert.equal(boxes.length, 1)
  assert.deepEqual(boxes[0].runs.map((r) => r.text), ["一行目", "二行目"])
  assert.equal(boxes[0].runs[0].options.breakLine, true)
  assert.equal(boxes[0].runs[1].options.breakLine, undefined)
  assert.equal(boxes[0].runs[0].options.lineSpacing, 27)
  assert.equal(boxes[0].y, 100 - baselineInLine(36, 26.6667))
})

test("a line split by an inline atom is placed stretch by stretch", () => {
  const block = { align: "left", box: { x: 40, w: 1200 }, lineHeight: 36,
    lines: [line(100, [seg(40, 120, "数式"), seg(260, 80, "も入る")])] }
  const boxes = textBoxes(block)
  assert.deepEqual(boxes.map((b) => b.x), [40, 260])
})

test("a right-aligned box takes its slack on the left, so the text's right edge stays put", () => {
  const block = { align: "right", box: { x: 600, w: 640 }, lineHeight: 36, lines: [line(100, [seg(1066, 174, "著者")])] }
  const [b] = textBoxes(block)
  assert.equal(b.x + b.w, 600 + 640)
})

test("templates are .mst files, not stylesheets", () => {
  assert.deepEqual(templateFiles(), ["academic.mst"])
})

test("the academic template is a markslide theme", () => {
  const { html } = renderMarkdown("---\nmarp: true\ntheme: academic\n---\n\n# 見出し\n")
  assert.match(html, /data-theme="academic"/)
})

test("a deck builds to HTML and a PPTX whose text is editable", async (t) => {
  let browser
  try { browser = await launchBrowser() } catch { return t.skip("no Chrome to lay the deck out") }
  const out = mkdtempSync(join(tmpdir(), "markslide-test-"))
  try {
    const { written, warnings } = await buildDeck(fileURLToPath(new URL("../samples/academic.md", import.meta.url)), { outDir: out, browser })
    assert.deepEqual(warnings, [])
    const html = readFileSync(written.html, "utf8")
    assert.equal(html.match(/class="ms-slide"/g).length, 4)
    assert.match(html, /src="data:image\/png;base64,/)
    const zip = await JSZip.loadAsync(readFileSync(written.pptx))
    const slide2 = await zip.file("ppt/slides/slide2.xml").async("string")
    assert.match(slide2, /Marp の HTML は Chrome で正確に描ける/)
    assert.match(slide2, /typeface="Noto Sans JP"/)
    // Bullets are Chrome's disc and circle as shapes, not a "•" glyph beside them.
    assert.match(slide2, /prstGeom prst="ellipse"/)
    assert.doesNotMatch(slide2, /<a:t>•<\/a:t>/)
    assert.ok(zip.file("ppt/notesSlides/notesSlide1.xml"))
  } finally {
    await browser.close()
    rmSync(out, { recursive: true, force: true })
  }
})
