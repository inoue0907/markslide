// Page types: what `markslide add` writes for each, and what `markslide check` holds it to.
// The limits are marp-slide-rules' (its check_rules.py), so a deck written one page at a time
// here passes the same review a Marp deck did.

/** [body characters, items, characters per item, figures, equations]; null = no limit. */
export const LIMITS = {
  表紙: [null, 0, null, 1, 0],
  目次: [120, 6, 20, 0, 0],
  背景: [150, 3, null, 1, 0],
  課題: [200, 4, null, 1, 0],
  手法: [150, 4, null, 0, 2],
  結果: [120, 3, null, 2, 0],
  まとめ: [250, 6, null, 0, 0],
  今後: [150, 6, 43, 0, 0],
};
export const TYPES = Object.keys(LIMITS);

export const TITLE_MAX = 30;
export const ITEM_MAX = 50;
export const COLUMN_ITEM_MAX = 45;
export const STRONG_MAX = 2;
export const CITE_MAX = 3;
export const NOTE_MAX = 600;
export const SYMBOL_MAX = 6;

// The elements a slide may be written with. Anything else (a <style>, a <script>, an inline
// style, a <font>) is a layout the theme does not know, which the PPTX cannot follow.
export const ALLOWED_TAGS = new Set([
  "section", "h1", "h2", "h3", "h4", "p", "ul", "ol", "li", "div", "span", "strong", "em",
  "b", "i", "code", "pre", "br", "sup", "sub", "a", "mark", "table", "thead", "tbody", "tr",
  "th", "td", "img", "blockquote", "aside", "header", "footer", "figure", "figcaption", "small",
]);

/** The skeleton `add` writes: the words are placeholders the AI replaces, the shape stays. */
export function template(type, id) {
  const head = `<section id="${id}" data-type="${type}">`;
  const notes = `  <aside class="notes">（話す内容）</aside>\n</section>`;
  const items = (n) => Array.from({ length: n }, () => `    <li>（項目）</li>`).join("\n");
  switch (type) {
    case "表紙":
      return `${head}\n  <h1>（タイトル）</h1>\n  <h2>（サブタイトル）</h2>\n  <p>（発表者・日付）</p>\n${notes}`;
    case "目次":
      return `${head}\n  <h2>目次</h2>\n  <ol>\n${items(4)}\n  </ol>\n${notes}`;
    case "手法":
      return `${head}\n  <h2>（主張を一文で）</h2>\n  <div class="math">（TeX）</div>\n` +
        `  <div class="sym"><p>（記号の説明）</p></div>\n  <ul>\n${items(2)}\n  </ul>\n${notes}`;
    case "結果":
      return `${head}\n  <h2>（結果を一文で）</h2>\n  <div class="flex">\n    <div>\n` +
        `      <img data-asset="（図の名前）" alt="">\n      <p class="cap">（図の説明）</p>\n` +
        `    </div>\n    <div>\n      <ul>\n    ${items(2).replace(/\n/g, "\n    ")}\n      </ul>\n` +
        `    </div>\n  </div>\n${notes}`;
    default:
      return `${head}\n  <h2>（主張を一文で）</h2>\n  <ul>\n${items(3)}\n  </ul>\n${notes}`;
  }
}
