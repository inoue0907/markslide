#!/usr/bin/env node
// markslide: an HTML deck written one page at a time, viewed and mended in the browser, and
// handed to PowerPoint as an editable PPTX that looks the same.
//
//   markslide new deck.html [--title T] [--theme academic]
//   markslide add deck.html --type 背景 [--after s3]
//   markslide check deck.html [--id s3]
//   markslide show deck.html s3 [--png]
//   markslide asset deck.html fig1 figure.png
//   markslide sync deck.html
//   markslide pptx deck.html [-o deck.pptx]
//   markslide import deck.md [-o deck.html]
//
// The older Marp path stays: markslide deck.md [-o outdir] [--no-html] [--no-pptx], and
// markslide fidelity deck.md [-o outdir] (needs PowerPoint).
import { join } from "node:path";
import { mkdtempSync } from "node:fs";
import { tmpdir } from "node:os";
import { parseArgs } from "node:util";

const { values, positionals } = parseArgs({
  allowPositionals: true,
  options: {
    out: { type: "string", short: "o" },
    "no-html": { type: "boolean" },
    "no-pptx": { type: "boolean" },
    title: { type: "string" },
    theme: { type: "string" },
    type: { type: "string" },
    after: { type: "string" },
    id: { type: "string" },
    png: { type: "boolean" },
    help: { type: "boolean", short: "h" },
  },
});

const usage = `usage: markslide new <deck.html> [--title T] [--theme academic]
       markslide add <deck.html> --type <型> [--after sN]
       markslide check <deck.html> [--id sN]
       markslide show <deck.html> <sN> [--png]
       markslide asset <deck.html> <name> <image>
       markslide sync <deck.html>
       markslide pptx <deck.html> [-o deck.pptx]
       markslide import <deck.md> [-o deck.html]
       markslide <deck.md> [-o outdir] [--no-html] [--no-pptx]
       markslide fidelity <deck.md> [-o outdir]`;

const [verb, ...rest] = positionals;
if (values.help || !verb) {
  console.log(usage);
  process.exit(values.help ? 0 : 2);
}
const need = (n) => {
  if (rest.length < n) {
    console.error(usage);
    process.exit(2);
  }
};

try {
  const deck = await import("./deck.mjs");
  switch (verb) {
    case "new": {
      need(1);
      await deck.newDeck(rest[0], { title: values.title, theme: values.theme });
      console.log(rest[0]);
      break;
    }
    case "add": {
      need(1);
      const { id, page } = deck.addPage(rest[0], { type: values.type, after: values.after });
      await deck.sync(rest[0]);
      console.log(`${id} を足しました。中身を書き換えてから check してください:\n${page}`);
      break;
    }
    case "check": {
      need(1);
      const { problems, pages } = await deck.checkDeck(rest[0], { id: values.id });
      for (const p of problems) console.log(p);
      if (!problems.length) console.log(`問題なし（${pages} 枚）`);
      process.exit(problems.length ? 1 : 0);
      break;
    }
    case "show": {
      need(2);
      const png = values.png ? join(mkdtempSync(join(tmpdir(), "markslide-")), `${rest[1]}.png`) : null;
      const { html } = await deck.showPage(rest[0], rest[1], { png });
      console.log(html);
      if (png) console.log(png);
      break;
    }
    case "asset": {
      need(3);
      deck.addAsset(rest[0], rest[1], rest[2]);
      console.log(`画像「${rest[1]}」を入れました`);
      break;
    }
    case "sync": {
      need(1);
      const { missing } = await deck.sync(rest[0]);
      for (const m of missing) console.log(`画像「${m}」がまだありません`);
      break;
    }
    case "pptx": {
      need(1);
      const { written, warnings } = await deck.deckToPptx(rest[0], { out: values.out });
      for (const w of warnings) console.error(`warning: ${w} (PowerPoint cannot draw this; the PPTX will differ)`);
      console.log(written);
      break;
    }
    case "import": {
      need(1);
      const { written, warnings } = await deck.importMarkdown(rest[0], { out: values.out });
      for (const w of warnings) console.error(`warning: ${w}`);
      console.log(written);
      break;
    }
    case "fidelity": {
      need(1);
      const { fidelity } = await import("./fidelity.mjs");
      const report = await fidelity(rest[0], { outDir: values.out });
      console.log(JSON.stringify(report.summary, null, 2));
      break;
    }
    default: {
      const { buildDeck } = await import("./build.mjs");
      const { written, warnings } = await buildDeck(verb, {
        outDir: values.out, html: !values["no-html"], pptx: !values["no-pptx"],
      });
      for (const w of warnings) console.error(`warning: ${w} (PowerPoint cannot draw this; the PPTX will differ)`);
      for (const path of Object.values(written)) console.log(path);
    }
  }
} catch (e) {
  console.error(`markslide: ${e.message}`);
  process.exit(1);
}
