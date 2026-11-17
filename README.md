# markslide

Marp の Markdown から、見た目のそろった HTML と PPTX を書き出す。

```bash
npm install
node cli.mjs deck.md -o out/            # out/deck.html と out/deck.pptx
node cli.mjs fidelity deck.md -o check/ # PowerPoint で開いた絵と比べる（Windows、PowerPoint 必須）
npm test
```

## 仕組み

1. marp-core が Markdown をテンプレート（`templates/*.mst`）で HTML にする。
2. Chrome（Playwright）がそのページを並べる。この配置だけを正とする。
3. `measure.js` が並んだ結果を読む。図形・画像・文字の行を、1文字ずつの位置まで取る。
4. HTML は、測ったページそのものを書き出す。PPTX は、測った位置に図形・画像・テキストボックスを置く（`pptx.mjs`）。

PPTX の文字は、Chrome が折り返した位置で改行を入れ、折り返しを切り、行間を固定値にする。こうすると PowerPoint に任せるのは行の中で文字を置く位置だけになる。その位置は `fidelity` の結果から `CALIBRATION` に決めてある。数式・SVG・切り抜いた画像は、Chrome で撮った画像として貼る。

## テンプレートで使える CSS

テンプレートの拡張子は `.mst`（markslide template）にする。中身は CSS で書くが、PPTX もこれをもとに作るので、普通の CSS とは書ける範囲が違う。`templates/` に置いた `.mst` はすべて読み込まれる。

PPTX に写せるのは次の範囲に限る。

- 背景色・枠線（実線・破線・点線）・角丸
- 文字の色・太さ・斜体・下線・取り消し線・上付き・下付き・字間
- 位置とサイズ（flex や grid、absolute も可。Chrome が決めた結果を読むため）
- `::before`／`::after` は使わない。リストの記号とページ番号は `measure.js` が本物の要素に直してから測る

次の指定は PPTX で再現できない。書き出し時に警告が出る。

`background-image`、`box-shadow`、`text-shadow`、`transform`、`filter`、`opacity` < 1、`-webkit-text-stroke`、`writing-mode`

## academic テンプレート

[kaisugi/marp-theme-academic](https://github.com/kaisugi/marp-theme-academic)（MIT）の academic テーマをもとにした。元のテーマとの違いは次のとおり。

- フォントは元と同じ Noto Sans JP と Source Code Pro を使う。ただし Google Fonts からは読まず、PC に入っているものを使う。PowerPoint はウェブフォントを読めないからだ。
- Source Code Pro が入っていない PC では、コードを Consolas で描く。PPTX にも Chrome が実際に使ったフォント名を書く。
- gaia テーマを読み込まず、単体で完結させた。リストの余白は gaia が任せているブラウザの既定値と同じにした。
- 擬似要素による飾りを外した。リストの記号は実在の要素として描く。黒丸と白丸は Chrome と同じ大きさと位置の図形にし、番号は Chrome と同じ「1. 」の文字にした。
- よく使う追加の指定も、テンプレートの末尾に入れた。見出しの下線、赤い太字、下端の引用の帯、`.note`・`.cap`・`.mid`・`.sym`、段組みの幅がそれにあたる。資料の側にはこの `<style>` が要らない。スライドごとの `<style scoped>` はこれまでどおり使える。

詳しくは `templates/academic.mst` の冒頭に書いてある。

## 今の精度

`samples/academic.md`（4枚・40行）で、PPTX を PowerPoint で描いた絵と Chrome の絵を比べた結果は次のとおり。

- 縦のずれは全行で 1px 以内。
- 横のずれも 1px 以内。
- 行の幅の差は 2px 以内。
- 絵が食い違う画素はスライドの 0.8〜2.0%。差の大半は、字形の描き方（アンチエイリアス）の違いによる。

本物の Marp（gaia と元の academic.css）で描いた HTML とも比べた。Noto Sans JP と Source Code Pro が入った PC では、4枚とも食い違う画素が 0.01% 以下だった。Source Code Pro がない PC ではコードが Consolas になり、コードの幅がずれる。Source Code Pro は Adobe の GitHub の Releases から TTF を取り、自分のユーザーにだけ入れれば管理者権限は要らない。

数式と画像を含む8枚の資料でも、本物の Marp で描いた絵と並べた。8枚のうち5枚は画素まで一致し、残り3枚も1〜2pxのずれだった。PowerPoint で描いた絵とは、77行すべてで縦のずれが 2px 以内だった。

## ライセンス

MIT（`LICENSE`）。`templates/academic.mst` は kaisugi/marp-theme-academic（MIT, Copyright (c) 2022 Kaito Sugimoto）をもとにしている。
