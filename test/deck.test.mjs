// The HTML deck: one deck.html written a page at a time, mended and saved in the browser, and
// measured into the PPTX. Drives Chrome (installed, else Playwright's); skipped without one.
import { test } from "node:test"
import assert from "node:assert/strict"
import { mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import { join } from "node:path"
import { fileURLToPath, pathToFileURL } from "node:url"
import JSZip from "jszip"
import { launchBrowser } from "../build.mjs"
import * as deck from "../deck.mjs"

const sample = fileURLToPath(new URL("../samples/academic.md", import.meta.url))

async function withDeck(t, fn) {
  let browser
  try { browser = await launchBrowser() } catch { return t.skip("no Chrome to lay the deck out") }
  const dir = mkdtempSync(join(tmpdir(), "markslide-deck-"))
  try { await fn(browser, dir) } finally {
    await browser.close()
    rmSync(dir, { recursive: true, force: true })
  }
}

test("mainSpan never reaches the </main> inside the runtime below the marker", () => {
  const text = `<body>\n<main data-theme="a">\n</main>\n${deck.MARKER}\n<script>"</main>"</script>`
  const [s, e] = deck.mainSpan(text)
  assert.equal(text.slice(s, e), `<main data-theme="a">\n</main>`)
})

test("a page added after another lands right after it with the next free id", (t) => withDeck(t, async (browser, dir) => {
  const path = join(dir, "d.html")
  await deck.newDeck(path, { title: "t", browser })
  deck.addPage(path, { type: "背景" })
  const { id } = deck.addPage(path, { type: "目次", after: "s1" })
  assert.equal(id, "s3")
  const ids = [...deck.mainText(readFileSync(path, "utf8")).matchAll(/<section id="(s\d+)"/g)].map((m) => m[1])
  assert.deepEqual(ids, ["s1", "s3", "s2"])
  assert.throws(() => deck.addPage(path, { type: "雑談" }), /型「雑談」/)
}))

test("a Marp deck imports to pages with ids, types, notes and formulas kept as TeX", (t) => withDeck(t, async (browser, dir) => {
  const { written } = await deck.importMarkdown(sample, { out: join(dir, "a.html"), browser })
  const text = readFileSync(written, "utf8")
  const main = deck.mainText(text)
  assert.equal(main.match(/<section /g).length, 4)
  assert.match(main, /<main data-theme="academic" data-paginate data-header="markslide 試作">/)
  assert.match(main, /<section id="s1" data-type="表紙">/)
  assert.match(main, /<aside class="notes">表紙のノート<\/aside>/)
  assert.match(main, /<span class="math">E = mc\^2<\/span>/)
  assert.match(main, /<img alt="" data-asset="fig1" width="320">/)
  assert.ok(text.includes(String.raw`"I:E = mc^2":"\u003cmjx-container`))
  // Syncing a synced deck writes nothing: the CLI and the page share one serializer.
  assert.equal((await deck.sync(written, { browser })).changed, false)
}))

test("check measures what overflows and names markup the theme does not know", (t) => withDeck(t, async (browser, dir) => {
  const path = join(dir, "d.html")
  await deck.newDeck(path, { title: "t", browser })
  deck.addPage(path, { type: "背景" })
  const text = readFileSync(path, "utf8")
  const many = Array.from({ length: 12 }, () => "    <li>長い説明がここに入りますので場所をとります</li>").join("\n")
  writeFileSync(path, text.replace('<ul>\n    <li>（項目）</li>', `<ul class="nope">\n${many}\n    <li>（項目）</li>`)
    .replace("<h2>（主張を一文で）</h2>", '<h2 style="color:red">（主張を一文で）</h2>'))
  const { problems } = await deck.checkDeck(path, { browser, id: "s2" })
  assert.ok(problems.some((p) => /^s2 背景: 下に \d+px はみ出しています$/.test(p)), problems.join("\n"))
  assert.ok(problems.includes("s2 背景: <h2> に style 属性があります（テーマのクラスを使う）"))
  assert.ok(problems.includes("s2 背景: クラス「nope」はテーマにありません"))
  assert.ok(problems.includes("s2 背景: 項目が 15 個です (背景の上限 3)"))
}))

test("an untouched page saves byte for byte, and an edit changes only its own line", (t) => withDeck(t, async (browser, dir) => {
  const { written } = await deck.importMarkdown(sample, { out: join(dir, "a.html"), browser })
  const text = readFileSync(written, "utf8")
  const page = await browser.newPage()
  try {
    await page.goto(pathToFileURL(written).href)
    await page.waitForFunction(() => window.__msReady === true)
    const same = await page.evaluate(() => window.__ms.serialize())
    assert.equal(same, deck.mainText(text))
    const edited = await page.evaluate(() => {
      window.__ms.setEditing(true)
      document.querySelector("#s3 li").textContent = "Markdown を HTML にする"
      return window.__ms.serialize()
    })
    const [s, e] = deck.mainSpan(text)
    const saved = text.slice(0, s) + edited + text.slice(e)
    const diff = text.split("\n").filter((l, i) => l !== saved.split("\n")[i])
    assert.deepEqual(diff, ["    <li>Markdown を marp-core で HTML にする</li>"])
  } finally {
    await page.close()
  }
}))

test("the PPTX is measured from the HTML deck, with each page's notes", (t) => withDeck(t, async (browser, dir) => {
  const { written } = await deck.importMarkdown(sample, { out: join(dir, "a.html"), browser })
  const { written: pptx, slides } = await deck.deckToPptx(written, { browser })
  assert.equal(slides, 4)
  const zip = await JSZip.loadAsync(readFileSync(pptx))
  assert.match(await zip.file("ppt/slides/slide2.xml").async("string"), /Marp の HTML は Chrome で正確に描ける/)
  assert.match(await zip.file("ppt/notesSlides/notesSlide1.xml").async("string"), /表紙のノート/)
  // The notes are said in the notes pane, not drawn on the slide.
  assert.doesNotMatch(await zip.file("ppt/slides/slide1.xml").async("string"), /表紙のノート/)
}))

test("the deck is read as one column or one slide at a time, and the choice never reaches the file", (t) => withDeck(t, async (browser, dir) => {
  const { written } = await deck.importMarkdown(sample, { out: join(dir, "a.html"), browser })
  const text = readFileSync(written, "utf8")
  const page = await browser.newPage({ viewport: { width: 1000, height: 700 } })
  try {
    await page.goto(pathToFileURL(written).href)
    await page.waitForFunction(() => window.__msReady === true)
    const shown = () => page.evaluate(() => window.__ms.sections()
      .filter((s) => s.getBoundingClientRect().height > 0).map((s) => s.id))
    assert.deepEqual(await shown(), ["s1", "s2", "s3", "s4"])

    await page.click(".ms-toolbar [data-act=view]")
    assert.deepEqual(await shown(), ["s1"])
    // The one slide fits the window's height as well as its width.
    const box = await page.locator("#s1").boundingBox()
    assert.ok(box.y >= 0 && box.y + box.height <= 700, JSON.stringify(box))

    await page.keyboard.press("ArrowRight")
    assert.deepEqual(await shown(), ["s2"])
    await page.mouse.move(500, 350)
    await page.mouse.wheel(0, 120)
    assert.deepEqual(await shown(), ["s3"])
    await page.keyboard.press("End")
    assert.deepEqual(await shown(), ["s4"])
    assert.equal(await page.locator(".ms-count").textContent(), "4 / 4")
    assert.equal(await page.evaluate(() => window.__ms.serialize()), deck.mainText(text))

    // The choice is the browser's: a reload comes back to the same slide, one at a time.
    await page.reload()
    await page.waitForFunction(() => window.__msReady === true)
    assert.deepEqual(await shown(), ["s4"])

    // One at a time the slide fills the window, and the toolbar hides until the mouse moves.
    assert.ok(Math.abs((await page.locator("#s4").boundingBox()).width - 1000) < 2)
    assert.equal(await page.getAttribute(".ms-toolbar [data-act=view]", "aria-label"), "連続で見る")
    await page.waitForFunction(() => document.documentElement.classList.contains("ms-idle"), null, { timeout: 4000 })
    await page.mouse.move(300, 300)
    assert.equal(await page.evaluate(() => document.documentElement.classList.contains("ms-idle")), false)

    await page.click(".ms-toolbar [data-act=view]")
    assert.deepEqual(await shown(), ["s1", "s2", "s3", "s4"])
    await page.waitForFunction(() => {
      const r = document.getElementById("s4").getBoundingClientRect()
      return r.top < innerHeight && r.bottom > 0
    })
  } finally {
    await page.close()
  }
}))
