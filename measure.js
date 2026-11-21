// Runs inside the page (page.evaluate). Two jobs, in this order:
//   1. prepare(): turn what cannot be measured into elements that can - list markers and the
//      page number are pseudo-elements, whose boxes no DOM API returns.
//   2. measure(): read every slide back as shapes, pictures and lines of text, in slide px.
// Plain script, no imports: it is sent to the browser as a function's source text.

/* eslint-disable no-undef */
window.__markslide = (() => {
  const BULLETS = ["•", "◦", "▪"];

  // A bullet is drawn as a shape, the way Chrome paints a disc/circle/square ::marker (its
  // list_marker.cc): a "•" glyph sat 9px right of and 3px under Chrome's 10px disc, which
  // was all that still told a markslide slide from the Marp one. A is the font's integer
  // ascent; the marker box is A tall, ending on the first line's baseline.
  function bullet(li, marker, kind) {
    const cs = getComputedStyle(li);
    const ctx = document.createElement("canvas").getContext("2d");
    ctx.font = `${cs.fontStyle} ${cs.fontWeight} ${cs.fontSize} ${cs.fontFamily}`;
    const A = Math.round(ctx.measureText("x").fontBoundingBoxAscent);
    const two = Math.floor((2 * A) / 3);
    const size = Math.floor((two + 1) / 2);
    const host = li.firstElementChild !== marker ? li : marker.nextElementSibling;
    const line = host && host !== li && getComputedStyle(host).display === "block" ? host : li;
    const probe = document.createElement("span");
    probe.style.cssText = "display:inline-block;width:0;height:0;vertical-align:baseline";
    line === li ? marker.after(probe) : line.prepend(probe);
    const baseline = probe.getBoundingClientRect().bottom;
    probe.remove();
    const scale = li.getBoundingClientRect().width / li.offsetWidth || 1;
    const top = (baseline - li.getBoundingClientRect().top) / scale - A + Math.floor((3 * (A - two)) / 2);
    const round = kind !== "▪";
    // Chrome strokes the circle on the rect's edge, half outside it; a border lies inside.
    const grow = kind === "◦" ? 0.5 : 0;
    marker.style.cssText = [
      "right:auto", "padding:0", "box-sizing:border-box",
      `left:${-(two + 7) - grow}px`, `top:${top - grow}px`,
      `width:${size + 2 * grow}px`, `height:${size + 2 * grow}px`,
      round ? `border-radius:${size / 2 + grow}px` : "",
      kind === "◦" ? "border:1px solid currentColor" : "background:currentColor",
    ].filter(Boolean).join(";");
  }

  function prepare() {
    for (const section of document.querySelectorAll(window.__markslideSections || "div.marpit > section")) {
      for (const li of section.querySelectorAll("li")) {
        if (li.querySelector(":scope > .ms-marker")) continue;
        const list = li.parentElement;
        const marker = document.createElement("span");
        marker.className = "ms-marker";
        marker.setAttribute("aria-hidden", "true");
        if (list && list.tagName === "OL") {
          const items = [...list.children].filter((c) => c.tagName === "LI");
          const start = Number(list.getAttribute("start") || 1);
          marker.textContent = `${start + items.indexOf(li)}. `; // Chrome's own suffix: ". "
          li.prepend(marker);
        } else {
          let depth = 0;
          for (let p = li.parentElement; p && p !== section; p = p.parentElement) {
            if (p.tagName === "UL" || p.tagName === "OL") depth++;
          }
          li.prepend(marker);
          bullet(li, marker, BULLETS[(depth - 1) % BULLETS.length]);
        }
      }
      const n = section.getAttribute("data-marpit-pagination");
      if (n && !section.querySelector(":scope > .ms-page")) {
        const total = section.getAttribute("data-marpit-pagination-total");
        const page = document.createElement("span");
        page.className = "ms-page";
        page.textContent = total ? `${n}/${total}` : n;
        section.append(page);
      }
    }
  }

  // ---- colours -------------------------------------------------------------------------
  function parseColor(s) {
    const m = /^rgba?\(\s*([\d.]+)[ ,]+([\d.]+)[ ,]+([\d.]+)(?:\s*[,/]\s*([\d.]+%?))?\s*\)$/.exec(s || "");
    if (!m) return null;
    let a = m[4] === undefined ? 1 : m[4].endsWith("%") ? parseFloat(m[4]) / 100 : parseFloat(m[4]);
    const hex = [m[1], m[2], m[3]].map((v) => Math.round(+v).toString(16).padStart(2, "0")).join("");
    return { hex: hex.toUpperCase(), alpha: a };
  }

  // ---- what PowerPoint cannot draw -------------------------------------------------------
  function describe(el) {
    const cls = typeof el.className === "string" && el.className ? "." + el.className.trim().split(/\s+/).join(".") : "";
    return el.tagName.toLowerCase() + cls;
  }
  function unsupported(el, cs, warn) {
    const checks = [
      ["background-image", cs.backgroundImage !== "none"],
      ["box-shadow", cs.boxShadow !== "none"],
      ["text-shadow", cs.textShadow !== "none"],
      ["transform", cs.transform !== "none"],
      ["filter", cs.filter !== "none"],
      ["opacity", parseFloat(cs.opacity) < 1],
      ["-webkit-text-stroke", parseFloat(cs.webkitTextStrokeWidth) > 0],
      ["writing-mode", cs.writingMode !== "horizontal-tb"],
    ];
    for (const [prop, bad] of checks) if (bad) warn(`${describe(el)}: ${prop}`);
  }

  // ---- fonts -------------------------------------------------------------------------------
  // The baseline of a run, as a fraction of its font size above the bottom of its glyph box.
  // A Range's rect is the font's content area (ascent + descent), so the baseline is
  // top + ascent; the ascent is measured once per font by an empty inline-block, whose
  // bottom sits on the baseline.
  const ascentCache = new Map();
  function ascentRatio(cs) {
    const key = `${cs.fontFamily}|${cs.fontWeight}|${cs.fontStyle}`;
    if (ascentCache.has(key)) return ascentCache.get(key);
    const probe = document.createElement("div");
    probe.style.cssText = `position:absolute;left:-9999px;top:0;font-family:${cs.fontFamily};` +
      `font-weight:${cs.fontWeight};font-style:${cs.fontStyle};font-size:100px;line-height:normal;white-space:nowrap`;
    const text = document.createElement("span");
    text.textContent = "Hg漢";
    const mark = document.createElement("span");
    mark.style.cssText = "display:inline-block;width:1px;height:0;vertical-align:baseline";
    probe.append(text, mark);
    document.body.append(probe);
    const r = document.createRange();
    r.selectNodeContents(text);
    const box = r.getBoundingClientRect();
    const base = mark.getBoundingClientRect().bottom;
    probe.remove();
    const ratio = { ascent: (base - box.top) / 100, descent: (box.bottom - base) / 100 };
    ascentCache.set(key, ratio);
    return ratio;
  }

  // The face Chrome actually drew a font-family list with: the first named family installed
  // here. PowerPoint is told that one, never the list's head - "Source Code Pro" absent from
  // Windows drew Consolas in Chrome while PowerPoint substituted a face of its own.
  const GENERIC = { "sans-serif": "Noto Sans JP", serif: "Yu Mincho", monospace: "Consolas" };
  const installed = new Map();
  function hasFamily(name) {
    if (installed.has(name)) return installed.get(name);
    const c = document.createElement("canvas").getContext("2d");
    const sample = "mmmmmmmmmmlli1WW漢字かなカナ";
    const found = ["monospace", "serif", "sans-serif"].some((g) => {
      c.font = `72px ${g}`;
      const plain = c.measureText(sample).width;
      c.font = `72px "${name}", ${g}`;
      return c.measureText(sample).width !== plain;
    });
    installed.set(name, found);
    return found;
  }
  const faceCache = new Map();
  function drawnFace(list) {
    if (faceCache.has(list)) return faceCache.get(list);
    let face = null;
    for (const raw of list.split(",")) {
      const name = raw.trim().replace(/^["']|["']$/g, "");
      if (GENERIC[name]) { face = GENERIC[name]; break; }
      if (hasFamily(name)) { face = name; break; }
    }
    face = face || "Noto Sans JP";
    faceCache.set(list, face);
    return face;
  }

  // ---- measuring ---------------------------------------------------------------------------
  const RASTER = "mjx-container, .katex, svg, canvas, video, iframe, object";

  function measureSection(section, index) {
    const origin = section.getBoundingClientRect();
    const X = (v) => v - origin.left;
    const Y = (v) => v - origin.top;
    const warnings = [];
    const warn = (w) => { if (!warnings.includes(w)) warnings.push(w); };
    const shapes = [];
    const pictures = [];
    let rasterId = 0;

    const sectionStyle = getComputedStyle(section);
    const background = parseColor(sectionStyle.backgroundColor) || { hex: "FFFFFF", alpha: 1 };
    unsupported(section, sectionStyle, warn);

    const boxes = (el, cs) => {
      if (cs.display === "inline") return [...el.getClientRects()];
      return [el.getBoundingClientRect()];
    };

    function visitBox(el, cs) {
      const fill = parseColor(cs.backgroundColor);
      const radius = parseFloat(cs.borderTopLeftRadius) || 0;
      for (const r of boxes(el, cs)) {
        if (r.width <= 0 || r.height <= 0) continue;
        // A box rounded all the way is an ellipse (a list bullet), its border one outline.
        if (el !== section && radius > 0 && radius * 2 >= Math.min(r.width, r.height) - 0.01) {
          const bw = parseFloat(cs.borderTopWidth) || 0;
          const bc = cs.borderTopStyle !== "none" ? parseColor(cs.borderTopColor) : null;
          const line = bw && bc && bc.alpha > 0 ? { width: bw, color: bc } : null;
          if ((fill && fill.alpha > 0) || line) {
            const inset = line ? bw / 2 : 0;
            shapes.push({
              kind: "ellipse", x: X(r.left + inset), y: Y(r.top + inset),
              w: r.width - 2 * inset, h: r.height - 2 * inset,
              fill: fill && fill.alpha > 0 ? fill : null, line,
            });
          }
          continue;
        }
        if (fill && fill.alpha > 0 && el !== section) {
          shapes.push({ kind: "rect", x: X(r.left), y: Y(r.top), w: r.width, h: r.height, fill, radius });
        }
        const sides = [
          ["Top", r.left, r.top, r.right, r.top, 1],
          ["Bottom", r.left, r.bottom, r.right, r.bottom, -1],
          ["Left", r.left, r.top, r.left, r.bottom, 1],
          ["Right", r.right, r.top, r.right, r.bottom, -1],
        ];
        for (const [side, x1, y1, x2, y2, inward] of sides) {
          const width = parseFloat(cs[`border${side}Width`]) || 0;
          const style = cs[`border${side}Style`];
          const color = parseColor(cs[`border${side}Color`]);
          if (!width || style === "none" || style === "hidden" || !color || color.alpha === 0) continue;
          // A CSS border lies inside the box; a PPTX line is centred on its path.
          const off = (width / 2) * inward;
          const horizontal = y1 === y2;
          shapes.push({
            kind: "line",
            x1: X(x1), y1: Y(horizontal ? y1 + off : y1),
            x2: X(x2), y2: Y(horizontal ? y2 + off : y2),
            ...(horizontal ? {} : { x1: X(x1 + off), x2: X(x2 + off) }),
            width, color, dash: style === "dashed" ? "dash" : style === "dotted" ? "sysDot" : "solid",
          });
        }
      }
    }

    // Shapes and pictures, in document order (which is paint order for this template).
    const visit = (el) => {
      const cs = getComputedStyle(el);
      if (cs.display === "none" || cs.visibility === "hidden") return;
      if (el !== section) unsupported(el, cs, warn);
      if (el !== section && el.matches(RASTER)) {
        const r = el.getBoundingClientRect();
        if (r.width > 0 && r.height > 0) {
          const id = `${index}-${rasterId++}`;
          el.setAttribute("data-ms-raster", id);
          pictures.push({ kind: "raster", id, x: X(r.left), y: Y(r.top), w: r.width, h: r.height });
        }
        return;
      }
      visitBox(el, cs);
      if (el.tagName === "IMG") {
        const r = el.getBoundingClientRect();
        if (r.width > 0 && r.height > 0) {
          const cropped = !["fill", "contain", "scale-down"].includes(cs.objectFit)
            && Math.abs(r.width / r.height - el.naturalWidth / el.naturalHeight) > 0.01;
          if (cropped) {
            const id = `${index}-${rasterId++}`;
            el.setAttribute("data-ms-raster", id);
            pictures.push({ kind: "raster", id, x: X(r.left), y: Y(r.top), w: r.width, h: r.height });
          } else {
            pictures.push({ kind: "image", src: el.currentSrc || el.src, x: X(r.left), y: Y(r.top), w: r.width, h: r.height });
          }
        }
      }
      for (const child of el.children) visit(child);
    };
    visit(section);

    return { index, background, shapes, pictures, texts: measureText(section, X, Y), warnings };
  }

  function containerOf(node, section) {
    for (let el = node.parentElement; el && el !== section; el = el.parentElement) {
      const d = getComputedStyle(el).display;
      if (d !== "inline" && d !== "contents") return el;
    }
    return section;
  }

  function runStyle(el) {
    const cs = getComputedStyle(el);
    let va = "baseline";
    for (let p = el; p && getComputedStyle(p).display === "inline"; p = p.parentElement) {
      const v = getComputedStyle(p).verticalAlign;
      if (v === "super" || v === "sub") { va = v; break; }
    }
    const deco = cs.textDecorationLine || "";
    const color = parseColor(cs.color) || { hex: "000000", alpha: 1 };
    return {
      font: cs.fontFamily,
      face: drawnFace(cs.fontFamily),
      size: parseFloat(cs.fontSize),
      bold: (parseInt(cs.fontWeight, 10) || 400) >= 600,
      italic: cs.fontStyle === "italic" || cs.fontStyle.startsWith("oblique"),
      color,
      underline: deco.includes("underline"),
      strike: deco.includes("line-through"),
      va,
      spacing: parseFloat(cs.letterSpacing) || 0,
      ...ascentRatio(cs),
    };
  }
  const styleKeys = new Map();
  function sharedStyle(el) {
    // One object per distinct style, so runs can be merged by identity.
    const st = runStyle(el);
    const key = JSON.stringify(st);
    if (!styleKeys.has(key)) styleKeys.set(key, st);
    return styleKeys.get(key);
  }

  function measureText(section, X, Y) {
    const blocks = new Map();
    const tw = document.createTreeWalker(section, NodeFilter.SHOW_TEXT);
    const range = document.createRange();
    for (let node = tw.nextNode(); node; node = tw.nextNode()) {
      const parent = node.parentElement;
      if (!parent || parent.closest(`${RASTER}, [data-ms-raster], style, script`)) continue;
      const pcs = getComputedStyle(parent);
      if (pcs.visibility === "hidden" || !node.data.length) continue;
      if (parent.closest("[style*='display: none'], [hidden]")) continue;
      const container = containerOf(node, section);
      if (!blocks.has(container)) blocks.set(container, []);
      const chars = blocks.get(container);
      const style = sharedStyle(parent);
      const preserve = /^pre/.test(pcs.whiteSpace) || pcs.whiteSpace === "break-spaces";
      const text = node.data;
      for (let i = 0; i < text.length; ) {
        const cp = text.codePointAt(i);
        const len = cp > 0xffff ? 2 : 1;
        let ch = text.slice(i, i + len);
        range.setStart(node, i);
        range.setEnd(node, i + len);
        i += len;
        const rects = range.getClientRects();
        if (!rects.length) continue;
        const r = rects[0];
        if (ch === "\n" && !preserve) ch = " ";
        if (ch === "\n" || ch === "\r") continue;
        if (r.width <= 0.01 && !/\s/.test(ch)) continue;
        if (/\s/.test(ch) && !preserve) {
          if (r.width <= 0.01) continue;
          ch = " ";
        }
        chars.push({ ch, left: X(r.left), right: X(r.right), top: Y(r.top), bottom: Y(r.bottom), style });
      }
    }

    const texts = [];
    for (const [container, chars] of blocks) {
      if (!chars.length) continue;
      const lines = [];
      let line = null;
      let prev = null;
      for (const c of chars) {
        const wrapped = prev && ((c.left < prev.left - 1 && c.top > prev.top + 1) || c.top >= prev.bottom - 1);
        if (!line || wrapped) {
          line = { segments: [] };
          lines.push(line);
        }
        let seg = line.segments[line.segments.length - 1];
        if (!seg || c.left - seg.right > 3 || c.left < seg.right - 3 - c.style.spacing) {
          seg = { x: c.left, right: c.right, chars: [] };
          line.segments.push(seg);
        }
        seg.chars.push(c);
        seg.right = Math.max(seg.right, c.right);
        prev = c;
      }
      const cs = getComputedStyle(container);
      const cr = container.getBoundingClientRect();
      const padL = parseFloat(cs.paddingLeft) + parseFloat(cs.borderLeftWidth);
      const padR = parseFloat(cs.paddingRight) + parseFloat(cs.borderRightWidth);
      texts.push({
        kind: "text",
        tag: container.tagName.toLowerCase(),
        align: cs.textAlign,
        box: { x: X(cr.left) + padL, y: Y(cr.top), w: cr.width - padL - padR, h: cr.height },
        lineHeight: parseFloat(cs.lineHeight) || null,
        lines: lines.map(finishLine),
      });
    }
    return texts;
  }

  function finishLine(line) {
    const segments = line.segments.map((seg) => {
      // Trim the space a wrapped line ends on; Chrome hangs it past the edge.
      while (seg.chars.length > 1 && seg.chars[seg.chars.length - 1].ch === " ") seg.chars.pop();
      const runs = [];
      for (const c of seg.chars) {
        const last = runs[runs.length - 1];
        if (last && last.style === c.style) last.text += c.ch;
        else runs.push({ text: c.ch, style: c.style });
      }
      const right = Math.max(...seg.chars.map((c) => c.right));
      return { x: seg.x, w: right - seg.x, runs };
    });
    // The line's baseline: from its biggest character that is not raised or lowered.
    let main = null;
    for (const s of line.segments) for (const c of s.chars) {
      if (c.style.va !== "baseline") continue;
      if (!main || c.style.size > main.style.size) main = c;
    }
    main = main || line.segments[0].chars[0];
    const baseline = main.top + main.style.ascent * main.style.size;
    const top = Math.min(...line.segments.flatMap((s) => s.chars.map((c) => c.top)));
    const bottom = Math.max(...line.segments.flatMap((s) => s.chars.map((c) => c.bottom)));
    return { baseline, top, bottom, size: main.style.size, segments };
  }

  function measure() {
    return [...document.querySelectorAll(window.__markslideSections || "div.marpit > section")].map(measureSection);
  }

  return { prepare, measure };
})();
