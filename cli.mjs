#!/usr/bin/env node
// markslide: a Marp deck -> HTML and PPTX that look the same.
//
//   node markslide/cli.mjs deck.md [-o outdir] [--no-html] [--no-pptx]
//   node markslide/cli.mjs fidelity deck.md [-o outdir]   (needs PowerPoint)
import { parseArgs } from "node:util";
import { buildDeck } from "./build.mjs";

const { values, positionals } = parseArgs({
  allowPositionals: true,
  options: {
    out: { type: "string", short: "o" },
    "no-html": { type: "boolean" },
    "no-pptx": { type: "boolean" },
    help: { type: "boolean", short: "h" },
  },
});

const usage = "usage: markslide <deck.md> [-o outdir] [--no-html] [--no-pptx]\n" +
  "       markslide fidelity <deck.md> [-o outdir]";

if (values.help || !positionals.length) {
  console.log(usage);
  process.exit(values.help ? 0 : 2);
}

if (positionals[0] === "fidelity") {
  const { fidelity } = await import("./fidelity.mjs");
  const report = await fidelity(positionals[1], { outDir: values.out });
  console.log(JSON.stringify(report.summary, null, 2));
  process.exit(0);
}

const { written, warnings } = await buildDeck(positionals[0], {
  outDir: values.out,
  html: !values["no-html"],
  pptx: !values["no-pptx"],
});
for (const w of warnings) console.error(`warning: ${w} (PowerPoint cannot draw this; the PPTX will differ)`);
for (const path of Object.values(written)) console.log(path);
