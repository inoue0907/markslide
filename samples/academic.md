---
marp: true
theme: academic
paginate: true
math: mathjax
header: markslide 試作
---

<!-- _class: lead -->

# markslide の試作
## HTML と PPTX を同じ見た目で書き出す

markslide 2026-11

<!-- 表紙のノート -->

---

# 背景

- Marp の HTML は Chrome で正確に描ける
- PPTX は **画像の貼り付け** になり編集できない
  - 文字を直したいときに困る
  - 入れ子の項目も *そのまま* 出したい
- そこで Chrome の配置を測って PPTX に写す

> 出典: Example, A. (2026). A made-up reference.

---

# 手順

1. Markdown を marp-core で HTML にする
2. Chrome で配置を決める
3. 測った位置で PPTX を組み立てる

インラインコード `measure.js` や数式 $E = mc^2$ も入る。

---

# 表と図

<div class="flex">
<div>

| 項目 | HTML | PPTX |
|---|---|---|
| 文字 | 〇 | 〇 |
| 表 | 〇 | 図形 |
| 数式 | 〇 | 画像 |

</div>
<div>

![w:320](figure.png)

</div>
</div>

<!-- 表と図のノート -->
