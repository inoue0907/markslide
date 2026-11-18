---
name: markslide
description: スライドを1つの HTML ファイル（deck.html）として1ページずつ書き、はみ出しと字数を調べ、見た目のそろった編集できる PowerPoint（.pptx）に書き出す。スライドを新しく作るとき、1ページ足すとき・直すとき、「pptx にして」「PowerPoint で欲しい」と頼まれたとき、Marp の .md を移すとき、テンプレート（templates/*.mst）を足すときや直すときに使う。
---

# markslide

スライドを1つの HTML ファイルとして持ち、この SKILL.md の隣にある `cli.mjs` で調べて PPTX にします。
PPTX の図形と文字は、HTML を Chrome で並べた位置に置くので、2つの見た目がそろいます。
PPTX は画像を貼っただけのものではなく、PowerPoint で文字を直せます。

以下の `$SKILL` は、この SKILL.md があるフォルダの絶対パスです。

## 準備（初回と更新のあと）

`$SKILL/node_modules` が無ければ、依存を入れます。

```bash
npm install --prefix $SKILL
```

Chrome は PC に入っているものを使います。入っていなければ、`npx --prefix $SKILL playwright install chromium` で入れます。

## 1ページずつ書く

資料は `deck.html` 1つです。前半の `<main>` の中だけを書きます。`<!-- markslide:runtime` から下は markslide が書くので触りません。

```bash
node $SKILL/cli.mjs new slides/deck.html --title "発表"   # 表紙だけの資料
node $SKILL/cli.mjs add slides/deck.html --type 背景      # 骨組みを1ページ足し、その中身を出す
node $SKILL/cli.mjs check slides/deck.html --id s2        # そのページを調べる。問題が無くなるまで直す
```

- `add` が出した骨組みの「（…）」を、Edit で本文に書き換えます。型は 表紙・目次・背景・課題・手法・結果・まとめ・今後 です。
- 1ページ書いたら `check --id sN` を通します。はみ出しは px で、字数と項目数は上限つきで出ます。出たら削るか分けます。
- 見た目はテーマのクラス（`flex`・`cap`・`sym`・`note` など）で決めます。`style` 属性と `<style>` は使えません。
- 数式は `<span class="math">TeX</span>`、行立ての式は `<div class="math">TeX</div>` です。
- 画像は `asset deck.html fig1 path/to.png` で中に入れ、`<img data-asset="fig1" width="400">` で呼びます。
- 発表者ノートは各ページの `<aside class="notes">` です。
- 1ページの中身と絵は `show deck.html s3 --png` で見られます。その絵を読むのは、配置を確かめたいときだけにします。
- Marp の `.md` は `import deck.md` で `deck.html` に移せます。`<!-- 型: 背景 -->` の注釈は `data-type` になります。

人は deck.html をブラウザで開き、「編集」で文字を直して Ctrl+S で上書きします。人が直した後に書くときは、ファイルを読み直してから Edit します。

## 書き出す

```bash
node $SKILL/cli.mjs pptx slides/deck.html            # → slides/deck.pptx
node $SKILL/cli.mjs slides/deck.md -o slides/out/    # Marp の .md から直接（従来の使い方）
```

- 発表者ノートは PPTX のノートに入ります。
- PPTX で再現できない CSS（`box-shadow`、`transform` など）があると、書き出し時に警告が出ます。

## テンプレート

`$SKILL/templates/` の `.mst` はすべて読み込まれます。
HTML の資料では `<main data-theme>`、Markdown では `theme:` で選びます。
今入っているのは `academic`（kaisugi/marp-theme-academic をもとにしたもの）です。
中身は CSS ですが、PPTX に写せる範囲でしか書けません。範囲は `$SKILL/README.md` にあります。
数式と SVG は画像として貼られるため、PowerPoint では直せません。

## 渡す

deck.html と PPTX は、パスを書くだけでは相手に届きません。`mcp__bridge__attach_file` で添付します。
レビューを受けるときの出し方は `mach:review-documents` にあります。

## 見た目を PowerPoint と比べる（Windows、PowerPoint が必要）

```bash
node $SKILL/cli.mjs fidelity slides/deck.md -o slides/check/
```

PowerPoint が描いた絵と Chrome の絵を並べ、ずれの大きいスライドを出します。
