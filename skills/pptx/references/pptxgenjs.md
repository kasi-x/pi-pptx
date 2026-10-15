# pptxgenjs 4.x — practical guide

Verified against **pptxgenjs 4.0.1** on 2026-10-15 (behaviors marked "verified" were probed on this machine; "OOXML rule" facts come from the shape/chart schemas).

## Minimal program

```js
const pptxgen = require("pptxgenjs");

const p = new pptxgen();
p.layout = "LAYOUT_WIDE";              // set BEFORE adding slides
const s = p.addSlide();
s.background = { color: "16324F" };
s.addText("Headline states the takeaway", { x: 0.7, y: 2.5, w: 11.9, h: 1.2, fontSize: 40, bold: true, color: "FFFFFF", margin: 0 });
s.addNotes("Talk track for this slide.");
p.writeFile({ fileName: "deck.pptx" }).then((f) => console.log("wrote:", f));
```

One `new pptxgen()` per output file; `writeFile()` is the only exit point. Full working starter: `examples/deck.js`.

## Canvas & coordinates

- Coordinates are **inches** from the top-left, or percent strings (`"5%"`). Off-canvas content is written into the XML but never shown — nothing is clamped.
- `LAYOUT_16x9` (default) is **10 × 5.625 in** — verified: without setting `layout`, the canvas is 9144000 × 5143500 EMU. `LAYOUT_WIDE` is 13.333 × 7.5 in.
- Set `p.layout` (and `p.defineSlideMaster` if used) before adding slides.

## Text

```js
s.addText(
  [{ text: "Bold lead. ", options: { bold: true } }, { text: "Rest of the sentence." }],
  { x: 0.7, y: 1.6, w: 6.0, h: 2.0, fontFace: "Calibri", fontSize: 16, color: "1F2933",
    align: "left", valign: "top", margin: 0, paraSpaceAfter: 8, lineSpacingMultiple: 1.15 },
);
```

- Paragraphs = array of `{ text, options }`; `breakLine: true` starts a new line, `bullet: true` a native bullet (both **verified**: emits `<a:buChar>`; a literal `•` in the text renders doubled).
- `charSpacing: 6` → `spc="600"` in the XML (**verified**); `letterSpacing` is **ignored** (**verified**).
- Text boxes carry built-in padding; set `margin: 0` whenever text must align with shapes/lines (**verified**: writes `lIns="0"`; without it the padding is applied at render time).
- `fit: "shrink"` / `"resize"` exists for overflow, but reduce text instead — shrink-to-fit usually ships unreadable font sizes.
- Colors: plain hex without `#`. A leading `#` is tolerated (**verified**: `"#FF0000"` is written as `val="FF0000"`), but don't rely on it. **No alpha in the hex** — use `transparency: 0-100` (fills/images) or `opacity: 0-1` (shadows); each is ignored where it doesn't apply.
- Ellipsis/quotes: plain Unicode is fine.

## Shapes

- `s.addShape(p.ShapeType.rect, {...})`, `.roundRect`, `.ellipse`, `.line`, `.chevron`, `.arc`, … `rectRadius` only has an effect on rounded rectangles (**verified**: a plain `rect` with `rectRadius` stays `prst="rect"`).
- Fill/line/shadow are separate option objects:
  `fill: { color: "2A9D8F", transparency: 15 }`, `line: { color: "16324F", width: 1 }`,
  `shadow: { type: "outer", color: "000000", blur: 6, offset: 3, angle: 90, opacity: 0.35 }`.
- **Option objects are mutated in place** (**verified**: after one `addShape` call the object gained `line: {type:"none"}` and `objectName`) — build a fresh object per call, never share one `shadow`/options object between two `add*` calls.
- Shadow `offset` is documented for positive values; keep offsets ≥ 0 and cast upward shadows with `angle: 270` (negative offsets still wrote and rendered here, but PowerPoint behavior is unverified — don't risk it).

## Images

```js
s.addImage({ path: "logo.png", x: 11.0, y: 0.5, w: 1.5, h: 0.5, transparency: 20 });
s.addImage({ data: "image/png;base64," + buf.toString("base64"), x: 0.7, y: 1.2, w: 4, h: 3, sizing: { type: "contain", w: 4, h: 3 } });
```

- Local path or `data:` URI (the `image/png;base64,` / `image/jpeg;base64,` prefix is required); `sizing` controls crop vs fit.
- Icons: there is no SVG/icon support — rasterize icons yourself (e.g. render an icon to PNG ≥ 256 px) and insert as an image, or draw simple icons from shapes.
- Prefer real image files in the project; base64 keeps a single-file workflow but bloats the JS.

## Tables

```js
s.addTable(
  [
    [{ text: "Workstream", options: { bold: true, fill: { color: "16324F" }, color: "FFFFFF" } }, /* … */],
    [{ text: "Capacity upgrade", options: { color: "1F2933" } }, /* … */],
  ],
  { x: 0.7, y: 1.7, w: 11.9, colW: [6, 3, 2.9], rowH: 0.55, fontSize: 14,
    border: { type: "solid", color: "E5E9EC", pt: 0.5 }, valign: "middle", fontFace: "Calibri" },
);
```

Style per cell (fill/bold/color/align); keep row heights uniform and let text wrap — set `rowH` large enough for the longest cell.

## Charts

Native charts only — never screenshots. `s.addChart(p.ChartType.bar | .line | .pie | .doughnut | .area | .radar | .scatter | .bubble, data, opts)`; a combo is an array of series specs.

Verified-good configuration for a single series:

```js
s.addChart(p.ChartType.line, [{ name: "requests", labels: ["Q1","Q2","Q3","Q4"], values: [120,180,260,410] }], {
  x: 0.7, y: 1.6, w: 7.6, h: 5.2,
  showTitle: false, showLegend: false, showValue: false,               // bare defaults otherwise
  chartColors: ["2A9D8F"], lineDataSymbol: "circle",
  catAxisLabelColor: "5C6B73", valAxisLabelColor: "5C6B73",
  catAxisLabelFontSize: 12, valAxisLabelFontSize: 12,
  valGridLine: { color: "E5E9EC", size: 1 }, catGridLine: { style: "none" },
});
```

- Defaults render bare (no title, no labels, dated palette) — always set title/labels/colors explicitly, or turn them off deliberately.
- Data labels on **stacked** bar/column charts: `dataLabelPosition` must be `"ctr"`, `"inEnd"`, or `"inBase"`. `"outEnd"` is illegal for stacked axes in the OOXML schema — pptxgenjs writes it anyway (**verified**: `dLblPos val="outEnd"` lands in `chart1.xml` without error), and PowerPoint then refuses the chart. Nothing local catches this — check your chart types.
- Combo charts with a secondary axis need explicit `valAxes` **and** `catAxes` arrays (two entries each) in the chart options; without them the chart references axis ids it never declares.
- Add the embedded-workbook caveat: chart data lives in the chart part; editing chart XML by hand is rarely worth it — regenerate that slide instead.

## Notes & metadata

- `s.addNotes("...")` — plain text, once per slide; never put talk track into slide text boxes.
- `p.title`, `p.subject`, `p.author` set document metadata.

## Fonts & QA reliability

Pick fonts whose local preview is metric-true — otherwise the render loop lies about overflow. Measured substitutions on this machine (LibreOffice + fontconfig):

| Font in deck | Substituted with | Metric-compatible? |
|---|---|---|
| Calibri | Carlito | ✅ |
| Cambria | Caladea | ✅ |
| Arial | Arimo | ✅ |
| Times New Roman | Tinos | ✅ |
| Courier New | Cousine | ✅ |
| Aptos | Noto Sans | ❌ (widths differ) |
| BIZ UDGothic | BIZ UDGothic | ✅ (installed) |
| Noto Sans JP | Noto Sans JP | ✅ (installed) |
| Yu Gothic | Noto Sans | ❌ |
| Meiryo | IPAexGothic | ❌ |

Practical rules: use Calibri/Cambria/Arial/Times/Courier New for Latin body text; for Japanese decks use `BIZ UDGothic`/`Noto Sans JP`; when using anything else, leave ~10% width slack and treat QA text-fit as approximate.

## Known gotchas (quick list)

1. `layout` after the first `addSlide` is ignored.
2. Shared option objects get mutated → build fresh objects.
3. `#`-colors tolerated, alpha-in-hex not; `transparency` vs `opacity` are not interchangeable.
4. `letterSpacing` is a no-op (`charSpacing` is real).
5. `rectRadius` only on rounded rects.
6. Literal `•` doubles with `bullet: true`.
7. `outEnd` labels on stacked bars → PowerPoint-rejected chart.
8. Charts render bare without explicit style options.
9. Off-canvas shapes are silently invisible.
10. One `pptxgenjs` instance per file.

## Upstream

- PptxGenJS — README & API docs: https://github.com/gitbrent/PptxGenJS
- Issue tracker (check here before working around a "bug"): https://github.com/gitbrent/PptxGenJS/issues
