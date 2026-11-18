// The script a deck.html carries below its runtime marker. It runs in the person's browser
// (view, edit, save) and in markslide's Chrome (tidy, check, PPTX), so both read the deck
// through the same functions: what `markslide sync` writes is what the browser saves.
//
// The deck itself is <main data-theme> holding one <section id data-type> per slide. Every
// element this script adds for display (header, page number, list markers, rendered maths,
// image sources, the notes box) is taken off again by serialize(), so a save changes only
// the words the person changed and `git diff` shows exactly that.
// Plain script, no imports: it is pasted into the deck and evaluated in Chrome as text.

/* eslint-disable no-undef */
window.__markslideSections = "main[data-theme] > section";
window.__ms = (() => {
  const SECTIONS = window.__markslideSections;
  const RUNTIME = "[data-ms-runtime], .ms-marker, .ms-page";
  const BLOCKS = new Set(["SECTION", "DIV", "UL", "OL", "LI", "TABLE", "THEAD", "TBODY", "TR",
    "TH", "TD", "P", "H1", "H2", "H3", "H4", "H5", "H6", "BLOCKQUOTE", "ASIDE", "FIGURE",
    "FIGCAPTION", "IMG", "VIDEO", "HEADER", "FOOTER", "PRE"]);
  const RUNTIME_ATTRS = ["contenteditable", "spellcheck", "data-ms-tex", "data-ms-raster",
    "data-marpit-pagination", "data-marpit-pagination-total"];

  const json = (id) => {
    const el = document.getElementById(id);
    try { return el ? JSON.parse(el.textContent) : {}; } catch { return {}; }
  };

  // ---- serializing ------------------------------------------------------------------------
  const esc = (s) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")
    .replace(/ /g, "&nbsp;");
  const escAttr = (s) => s.replace(/&/g, "&amp;").replace(/"/g, "&quot;");

  /** The deck element as its author would write it: runtime additions off, one block a line. */
  function clean(main) {
    const copy = main.cloneNode(true);
    copy.querySelectorAll(RUNTIME).forEach((e) => e.remove());
    copy.querySelectorAll(".math[data-ms-tex]").forEach((m) => {
      m.textContent = m.getAttribute("data-ms-tex");
    });
    copy.querySelectorAll("img[data-asset]").forEach((i) => i.removeAttribute("src"));
    for (const el of [copy, ...copy.querySelectorAll("*")]) {
      for (const a of RUNTIME_ATTRS) el.removeAttribute(a);
    }
    return copy;
  }

  function openTag(el) {
    const attrs = [...el.attributes].map((a) => a.value === "" && a.name !== "alt"
      ? ` ${a.name}` : ` ${a.name}="${escAttr(a.value)}"`).join("");
    return `<${el.tagName.toLowerCase()}${attrs}>`;
  }

  // Inline content on one line: whitespace collapsed, so a re-indented file serializes the
  // same and a save never turns into a whitespace diff.
  function inline(nodes) {
    let out = "";
    for (const n of nodes) {
      if (n.nodeType === 3) out += esc(n.data);
      else if (n.nodeType === 1) out += n.outerHTML;
    }
    return out.replace(/\s+/g, " ").trim();
  }

  function block(el, depth) {
    const pad = "  ".repeat(depth);
    const tag = el.tagName.toLowerCase();
    if (tag === "img" || tag === "br") return pad + openTag(el);
    // Code keeps every space and line as written.
    if (tag === "pre") return pad + el.outerHTML;
    // Notes keep their line breaks: they are what is said, written as plain text.
    if (tag === "aside") return pad + openTag(el) + esc(el.textContent.trim()) + `</${tag}>`;
    const kids = [...el.childNodes].filter((n) => n.nodeType === 1 || n.nodeType === 3);
    if (!kids.some((n) => n.nodeType === 1 && BLOCKS.has(n.tagName))) {
      return pad + openTag(el) + inline(kids) + `</${tag}>`;
    }
    const lines = [pad + openTag(el)];
    let run = [];
    const flush = () => {
      const text = inline(run);
      if (text) lines.push(pad + "  " + text);
      run = [];
    };
    for (const n of kids) {
      if (n.nodeType === 1 && BLOCKS.has(n.tagName)) {
        flush();
        lines.push(block(n, depth + 1));
      } else run.push(n);
    }
    flush();
    lines.push(pad + `</${tag}>`);
    return lines.join("\n");
  }

  /** The text that stands for <main> in the file. */
  function serialize(main = document.querySelector("main[data-theme]")) {
    const copy = clean(main);
    const sections = [...copy.children].filter((c) => c.tagName === "SECTION");
    return [openTag(copy), ...sections.map((s) => block(s, 0)), "</main>"].join("\n");
  }

  // ---- display ------------------------------------------------------------------------------
  function sections() { return [...document.querySelectorAll(SECTIONS)]; }

  function hydrate() {
    const main = document.querySelector("main[data-theme]");
    if (!main) return;
    const assets = json("ms-assets");
    const maths = json("ms-math");
    for (const img of main.querySelectorAll("img[data-asset]")) {
      const src = assets[img.getAttribute("data-asset")];
      if (src) img.setAttribute("src", src);
    }
    for (const m of main.querySelectorAll(".math")) {
      if (m.hasAttribute("data-ms-tex")) continue;
      const tex = m.textContent.trim();
      const svg = maths[(m.tagName === "DIV" ? "D:" : "I:") + tex];
      if (!svg) continue;
      m.setAttribute("data-ms-tex", tex);
      m.setAttribute("contenteditable", "false");
      m.innerHTML = svg;
    }
    const all = sections();
    const header = main.getAttribute("data-header");
    const paginate = main.hasAttribute("data-paginate");
    all.forEach((s, i) => {
      if (header && !s.querySelector(":scope > header")) {
        const h = document.createElement("header");
        h.setAttribute("data-ms-runtime", "");
        h.textContent = header;
        s.prepend(h);
      }
      if (paginate && s.getAttribute("data-paginate") !== "false"
          && !s.querySelector(":scope > .ms-page")) {
        const p = document.createElement("span");
        p.className = "ms-page";
        p.textContent = `${i + 1}/${all.length}`;
        s.append(p);
      }
    });
  }

  // List markers are measure.js's real elements; laid out unzoomed, as the PPTX is.
  function markers() {
    if (!window.__markslide) return;
    const root = document.documentElement;
    const had = root.classList.contains("ms-measure");
    root.classList.add("ms-measure");
    window.__markslide.prepare();
    if (!had) root.classList.remove("ms-measure");
    document.querySelectorAll(".ms-marker").forEach((m) => m.setAttribute("contenteditable", "false"));
  }

  function fit() {
    const w = Math.min(1280, document.documentElement.clientWidth - 32);
    document.documentElement.style.setProperty("--ms-zoom", String(Math.max(0.2, w / 1280)));
  }

  // ---- editing --------------------------------------------------------------------------------
  // Only words: headings, paragraphs, items, cells and notes. A page, a figure or a formula
  // is the AI's to make, so nothing here can add or delete one.
  const EDITABLE = "h1, h2, h3, h4, h5, h6, p, li, th, td";
  let baseline = null;
  let dirty = false;
  let handle = null;
  let status = null;

  function setStatus(text) { if (status) status.textContent = text; }
  function markDirty() { dirty = true; setStatus("未保存の変更があります"); }

  function noteBoxes(on) {
    for (const s of sections()) {
      let box = s.nextElementSibling;
      if (!(box && box.classList.contains("ms-notes-box"))) {
        if (!on) continue;
        box = document.createElement("div");
        box.className = "ms-notes-box";
        box.setAttribute("data-ms-runtime", "");
        box.setAttribute("contenteditable", "plaintext-only");
        box.setAttribute("data-placeholder", "発表者ノート（ここに書く）");
        const aside = s.querySelector(":scope > aside.notes");
        box.textContent = aside ? aside.textContent.trim() : "";
        box.addEventListener("input", () => {
          let target = s.querySelector(":scope > aside.notes");
          if (!target) {
            target = document.createElement("aside");
            target.className = "notes";
            s.append(target);
          }
          target.textContent = box.innerText.replace(/\n$/, "");
          markDirty();
        });
        s.after(box);
      }
    }
  }

  function setEditing(on) {
    document.documentElement.classList.toggle("ms-editing", on);
    const main = document.querySelector("main[data-theme]");
    const roots = [...main.querySelectorAll(EDITABLE)].filter((el) =>
      !el.parentElement.closest(EDITABLE) && !el.closest(RUNTIME));
    for (const el of roots) {
      if (on) { el.setAttribute("contenteditable", "true"); el.setAttribute("spellcheck", "false"); }
      else { el.removeAttribute("contenteditable"); el.removeAttribute("spellcheck"); }
    }
    noteBoxes(on);
    const btn = document.querySelector(".ms-toolbar [data-act=edit]");
    if (btn) btn.textContent = on ? "編集を終える" : "編集";
  }

  function onKey(e) {
    const t = e.target;
    if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === "s") {
      e.preventDefault();
      save();
      return;
    }
    if (!(t instanceof Element) || !t.closest("main [contenteditable=true]")) return;
    // Enter would make a new paragraph or item, which is a structural change: a line break
    // inside the same one is Shift+Enter.
    if (e.key === "Enter" && !e.isComposing) {
      e.preventDefault();
      if (e.shiftKey) document.execCommand("insertLineBreak");
    }
  }

  function onPaste(e) {
    const t = e.target instanceof Element ? e.target : e.target.parentElement;
    if (!t || !t.closest("main [contenteditable=true]")) return;
    e.preventDefault();
    const text = (e.clipboardData.getData("text/plain") || "").replace(/\s*\n\s*/g, " ");
    document.execCommand("insertText", false, text);
  }

  // ---- saving ---------------------------------------------------------------------------------
  // Written split so this script never holds the marker itself: the first marker after the
  // deck is the real one, and an HTML comment opener inside a <script> is a parsing hazard.
  const MARKER = "<" + "!-- markslide:runtime ここから下は markslide が書く。手で直さない -->";

  /** [start, end) of the deck's <main> in the file's text: never into the runtime below it. */
  function mainSpan(text) {
    const cut = text.indexOf("<" + "!-- markslide:runtime");
    const head = cut < 0 ? text : text.slice(0, cut);
    const start = head.search(/<main[\s>]/);
    const end = head.lastIndexOf("</main>");
    return start < 0 || end < start ? null : [start, end + "</main>".length];
  }

  /** The whole file, from this page: used only where the file itself cannot be read back. */
  function wholeFile(mainText) {
    const tail = ["ms-theme", "ms-view", "ms-assets", "ms-math", "ms-runtime"]
      .map((id) => document.getElementById(id)).filter(Boolean).map((e) => e.outerHTML);
    return `<!doctype html>\n<html lang="ja">\n<head>\n<meta charset="utf-8">\n` +
      `<meta name="viewport" content="width=device-width">\n` +
      `<title>${esc(document.title)}</title>\n</head>\n<body>\n${mainText}\n` +
      `${MARKER}\n${tail.join("\n")}\n</body>\n</html>\n`;
  }

  function mainOf(text) {
    const doc = new DOMParser().parseFromString(text, "text/html");
    const main = doc.querySelector("main[data-theme]");
    return main ? serialize(main) : null;
  }

  async function save() {
    const main = serialize();
    if (!window.showSaveFilePicker) {
      // No File System Access (Firefox, Safari): hand the file over as a download instead.
      const a = document.createElement("a");
      a.href = URL.createObjectURL(new Blob([wholeFile(main)], { type: "text/html" }));
      a.download = decodeURIComponent(location.pathname.split("/").pop() || "deck.html");
      a.click();
      setStatus("ダウンロードしました。元のファイルと置き換えてください");
      return;
    }
    try {
      if (!handle) {
        handle = await window.showSaveFilePicker({
          suggestedName: decodeURIComponent(location.pathname.split("/").pop() || "deck.html"),
          types: [{ description: "HTML", accept: { "text/html": [".html"] } }],
        });
      }
      const disk = await (await handle.getFile()).text();
      let next;
      if (!disk.trim()) {
        next = wholeFile(main);
      } else {
        // Someone (the AI) wrote the file after it was opened here: overwriting would undo
        // their page. Stop, and let the person reload and redo the words.
        if (mainOf(disk) !== baseline) {
          setStatus("開いた後にファイルが変わっています。保存を止めました。再読み込みしてから直してください");
          alert("このファイルは開いた後に書き換えられています。\n上書きすると、その変更が消えるので保存を止めました。\n再読み込みしてから、もう一度直してください。");
          return;
        }
        const span = mainSpan(disk);
        next = disk.slice(0, span[0]) + main + disk.slice(span[1]);
      }
      const w = await handle.createWritable();
      await w.write(next);
      await w.close();
      baseline = main;
      dirty = false;
      setStatus(`保存しました（${new Date().toLocaleTimeString()}）`);
    } catch (err) {
      if (err && err.name === "AbortError") return;
      setStatus(`保存できませんでした: ${err && err.message}`);
    }
  }

  function toolbar() {
    const bar = document.createElement("div");
    bar.className = "ms-toolbar";
    bar.setAttribute("data-ms-runtime", "");
    bar.innerHTML = '<button type="button" data-act="edit">編集</button>' +
      '<button type="button" data-act="save">保存 (Ctrl+S)</button><span class="ms-status"></span>';
    bar.querySelector("[data-act=edit]").onclick = () =>
      setEditing(!document.documentElement.classList.contains("ms-editing"));
    bar.querySelector("[data-act=save]").onclick = save;
    status = bar.querySelector(".ms-status");
    document.body.append(bar);
  }

  function boot() {
    // markslide's own Chrome asks for the deck unzoomed and without the editing chrome.
    if (window.__msMeasure) document.documentElement.classList.add("ms-measure");
    hydrate();
    fit();
    addEventListener("resize", fit);
    const go = () => {
      markers();
      baseline = serialize();
      if (!document.documentElement.classList.contains("ms-measure")) {
        toolbar();
        document.addEventListener("keydown", onKey, true);
        document.addEventListener("paste", onPaste, true);
        document.querySelector("main[data-theme]").addEventListener("input", markDirty);
        addEventListener("beforeunload", (e) => { if (dirty) e.preventDefault(); });
      }
      window.__msReady = true;
    };
    document.fonts.ready.then(go, go);
  }

  /** One slide as the file writes it (for `markslide show`). */
  function serializeSection(section) {
    const main = document.createElement("main");
    main.append(section.cloneNode(true));
    return block(clean(main).firstElementChild, 0);
  }

  return { serialize, serializeSection, hydrate, sections, save, setEditing, mainOf, mainSpan,
    MARKER, boot };
})();

if (window.__msNoBoot) { /* markslide only wants serialize() */ }
else if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", () => window.__ms.boot());
else window.__ms.boot();
