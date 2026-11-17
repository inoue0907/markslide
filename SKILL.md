---
name: markslide
description: Marp の Markdown（.md）から、見た目のそろった HTML と編集できる PowerPoint（.pptx）を書き出す。スライドを作ったあとに HTML か PPTX で渡すとき、「pptx にして」「PowerPoint で欲しい」「スライドを HTML で見せて」と頼まれたとき、テンプレート（templates/*.mst）を足すときや直すときに使う。
---

# markslide

Marp の Markdown を、この SKILL.md の隣にある `cli.mjs` で HTML と PPTX にします。
PPTX の図形と文字は、HTML を Chrome で並べた位置に置くので、2つの見た目がそろいます。
PPTX は画像を貼っただけのものではなく、PowerPoint で文字を直せます。

以下の `$SKILL` は、この SKILL.md があるフォルダの絶対パスです。

## 準備（初回と更新のあと）

`$SKILL/node_modules` が無ければ、依存を入れます。

```bash
npm install --prefix $SKILL
```

Chrome は PC に入っているものを使います。入っていなければ、`npx --prefix $SKILL playwright install chromium` で入れます。

## 書き出す

```bash
node $SKILL/cli.mjs slides/deck.md -o slides/out/
# → slides/out/deck.html と slides/out/deck.pptx
```

- `--no-html` と `--no-pptx` で、片方だけにできます。
- 画像は Markdown からの相対パスで書きます。HTML には data URI で埋め込まれるので、HTML は1ファイルで開けます。
- 発表者ノート（スライドの下の `<!-- -->`）は、PPTX のノートに入ります。HTML には入りません。
- PPTX で再現できない CSS（`box-shadow`、`transform` など）を使うと、書き出し時に警告が出ます。

## テンプレート

`$SKILL/templates/` の `.mst` はすべて読み込まれ、Markdown の `theme:` で選びます。
今入っているのは `academic`（kaisugi/marp-theme-academic をもとにしたもの）です。
中身は CSS ですが、PPTX に写せる範囲でしか書けません。範囲は `$SKILL/README.md` にあります。
数式と SVG は画像として貼られるため、PowerPoint では直せません。

## 渡す

HTML と PPTX は、パスを書くだけでは相手に届きません。`mcp__bridge__attach_file` で添付します。
レビューを受けるときの出し方は `mach:review-documents` にあります。

## 見た目を PowerPoint と比べる（Windows、PowerPoint が必要）

```bash
node $SKILL/cli.mjs fidelity slides/deck.md -o slides/check/
```

PowerPoint が描いた絵と Chrome の絵を並べ、ずれの大きいスライドを出します。
