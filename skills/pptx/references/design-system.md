# Design system for generated decks

Fix the tokens **before** the first `addSlide()`: palette, type scale, grid and one layout family. A deck that has no design system shows up as one-layout repetition, drifting font sizes and colors picked per slide.

Numbers marked `[measured]` were probed on this machine on 2026-10-15 (pptxgenjs 4.0.1, LibreOffice + poppler for renders, python-pptx 1.0.2 for inspection); the rest are computed from the WCAG contrast formula or from the grid in §3-§4. Hex values are written **without `#`**, the form pptxgenjs wants. Geometry is in inches and assumes `LAYOUT_WIDE` (13.333 x 7.5); the 10 x 5.625 canvas is handled in §2 and §4.

## 1. Color

### Thresholds (WCAG 2.1)

| Content | Minimum ratio | Applies to |
|---|---|---|
| Body text under 18 pt | **4.5:1** | bullet text, table cells, axes, captions, source lines |
| Large text: ≥ 18 pt, or ≥ 14 pt **bold** | **3:1** | slide titles, key numbers, section heads |
| Non-text graphics needed to understand the slide | **3:1** | chart series vs background, icon glyphs, rules that carry meaning, status dots [W4] |
| (Optional, AAA) | 7:1 | high-stakes print/large-room decks |

Ratios are thresholds, not guidelines: 4.499:1 fails [W1]. Slide text is rendered by PowerPoint, so measure the colors you set, never the rendered pixels.

### 60-30-10 (area budget)

Three color jobs per slide, split by *area of visible marks*, not by importance:

| Role | Share of colored area | Tokens | How it shows up |
|---|---|---|---|
| Dominant | ~60% | `surface` (+ tints at 10-20% of a block, plus photographs) | slide background, full-bleed image, tinted bands |
| Support | ~30% | `ink`, `muted`, tint fills, borders | all body text, labels, table zebra, chart baselines, card fills |
| Accent | **<= 10%** | `accent` | one number, one phrase, one series, one status dot |

A `LAYOUT_WIDE` slide is **100 in2**, so the accent budget is 8-10 in2 — roughly a 3.4 x 2.9 in block, one 40 pt number, or a five-word phrase; on the 10 x 5.625 canvas the whole slide is 56 in2, so the budget is 5-6 in2. Count *marks*, not empty space: a 12.3 x 7.5 in white background is 100% dominant by definition, and one accent used twice (title keyword + chart series) already spends the budget.

Two failure modes to check after rendering: an accent that appears on a third element has stopped being emphasis, and a deck whose slides each have a *different* accent has no accent at all — the accent is a deck-level token, picked once from the palette table below.

### Worked example

Relative luminance `L` of a color channel `c` in 0-255: `c/255` -> linear (`/12.92` if `<= 0.04045`, else `((v+0.055)/1.055)^2.4`, with `v = c/255`), then `L = 0.2126 R + 0.7152 G + 0.0722 B`; the ratio is `(L_lighter + 0.05) / (L_darker + 0.05)` [W2].

```
1F2933 on FFFFFF   (the default ink on the default background)
  R=31 -> 0.121569 -> 0.013702
  G=41 -> 0.160784 -> 0.022174
  B=51 -> 0.200000 -> 0.033105
  L = 0.2126(0.013702) + 0.7152(0.022174) + 0.0722(0.033105) = 0.021162
  ratio = (1.0 + 0.05) / (0.021162 + 0.05) = 14.76:1   -> pass (body text)
```

Recompute any pair yourself instead of trusting a palette:

```python
def lum(h):                      # h = "RRGGBB"
    v = [int(h[i:i + 2], 16) / 255 for i in (0, 2, 4)]
    v = [x / 12.92 if x <= 0.04045 else ((x + 0.055) / 1.055) ** 2.4 for x in v]
    return 0.2126 * v[0] + 0.7152 * v[1] + 0.0722 * v[2]

def ratio(a, b):                 # symmetric: order does not matter
    lo, hi = sorted((lum(a), lum(b)))
    return (hi + 0.05) / (lo + 0.05)
```

### Ready-to-use palettes

Four tokens each: `surface` (background), `ink` (body), `muted` (labels, axes, captions), `accent` (one or two elements per slide, max). Ratios are against the palette's own surface; every value passes 4.5:1 for body text unless flagged.

| # | Palette | Surface | Ink | Muted | Accent | Ratios on surface |
|---|---|---|---|---|---|---|
| 1 | Slate & Teal (consulting default) | `FFFFFF` | `1F2933` | `5C6B73` | `0F766E` | ink 14.8 · muted 5.5 · accent 5.5 · tint `D9E2E6`: ink 11.2, muted 4.2 (**muted fails on tint**) |
| 2 | Navy (dark cover / section) | `10263B` | `FFFFFF` | `C3D0DB` | `5EEAD4` | ink 15.4 · muted 9.8 · accent 10.4 |
| 3 | Academic high-contrast | `FFFFFF` | `111111` | `4A4A4A` | `0B4F9E` | ink 18.9 · muted 8.9 · accent 8.0 |
| 4 | Product stage (dark) | `0B0F14` | `F8FAFC` | `C3CBD5` | `FFB703` | ink 18.4 · muted 11.7 · accent 11.0 |
| 5 | Internal status | `FFFFFF` | `1F2933` | `5C6B73` | `B3261E` | ink 14.8 · muted 5.5 · accent 6.5 |
| 6 | Warm paper | `FAF7F2` | `2B2622` | `6B5F55` | `9A3412` | ink 14.0 · muted 5.8 · accent 6.8 |
| 7 | Cool gray & indigo | `F5F7FA` | `1B2430` | `5A6675` | `3730A3` | ink 14.6 · muted 5.5 · accent 9.3 |
| 8 | Deep green (sustainability, dark) | `0E3B2E` | `F2FBF6` | `BFD9CB` | `F4C95D` | ink 11.8 · muted 8.3 · accent 7.9 |
| 9 | Ink & amber (finance / ops, dark) | `14181F` | `F5F5F4` | `BFBFBD` | `F59E0B` | ink 16.3 · muted 9.7 · accent 8.3 |
| 10 | Mono & magenta | `FFFFFF` | `232323` | `6E6E6E` | `C2185B` | ink 15.7 · muted 5.1 · accent 5.9 |

Rules that keep these honest:

- **Two-tone accent.** An accent that carries text must clear 4.5:1; an accent used only as a fill, chart line or status dot needs 3:1 and can be brighter. Example: `2A9D8F` on white is **3.32:1** (fine as a series color, too weak for a caption); `264653` on white is **10.08:1** (fine for text).
- **Chart series, light decks:** `264653` 10.1 · `2A9D8F` 3.3 · `E76F51` 3.1 (fills/lines only; keep data labels in `ink`). **Dark decks (`0B0F14`):** `5EEAD4` 13.0 · `FFB703` 11.0 · `93C5FD` 10.7.
- **Tints carry ink, not muted.** On the `D9E2E6` tint above, muted text falls to 4.2:1 — use `ink` on tint blocks and keep muted for white/near-white surfaces.
- **Status colors need a text cue.** `1B7F4B` 5.0 · `8A5A00` 5.9 · `B3261E` 6.5 on white all pass text contrast, but their pairwise ratios are only **1.07:1 to 1.30:1** — indistinguishable in grayscale and for red-green color vision. Always pair the color with a word ("At risk") or a glyph.

### Light/dark strategy (sandwich decks)

A "sandwich" deck uses the dark palette for the cover, each section divider and the closing slide, and the light palette for every content slide. Rules:

- One switch per section, only at a section boundary; never a dark slide in the middle of a light run.
- The dark surface keeps its own ink/muted/accent — do not reuse light-slide text colors on it (contrast collapses: `1F2933` on `10263B` is 1.04:1).
- Reuse one accent across both surfaces. `5EEAD4`/`FFB703`-style accents above are chosen to pass on their dark surface.
- Budget: a 20-slide deck = 1 cover + 2-3 dividers + closing, so ~75-80% of slides are light.

### The mid-gray dead band

Backgrounds between `767676` and `8E8E8E` are the only grays where **neither** white text nor ink `1F2933` reaches 4.5:1 (`787878`: white 4.42, ink 3.34; `888888`: white 3.54, ink 4.16). For large text (3:1) the band disappears: at `808080` white is 3.95:1 and ink 3.74:1 — so mid-gray is acceptable behind a big number, never behind a paragraph. Pick surfaces at the extremes: `F*` tints, or `1*`-`4*` darks.

### Text over images

Never leave the image responsible for contrast. Composite the picture with a solid scrim rectangle:

- `transparency: t` on a fill writes `<a:alpha val="(100-t)*1000"/>` [measured] and the QA render composites it exactly as `255 * t / 100` gray on a white pixel.
- Worst case is a white area in the photo, so white text needs a **black scrim at `transparency <= 46`** for 4.5:1, `<= 58` for 3:1 (large text only). Encode that as a literal: body text over photos uses `transparency: 45`, never 50+ (at 50 the worst case is 3.95:1).
- Safer alternative: place text on an opaque band or on the image's own dark region; PowerPoint's contrast checker does not evaluate text in a transparent box over an image [M1], so nothing warns you later.

## 2. Typography

### Scale

| Role | `LAYOUT_WIDE` 13.333 x 7.5 | `LAYOUT_16x9` 10 x 5.625 |
|---|---|---|
| Cover title | 44-54 | 33-40 |
| Slide title (max 2 lines) | 36-44 | 27-33 |
| Section head / kicker | 20-24 | 15-18 |
| Key number / stat | 40-54 | 30-40 |
| Lead-in (subtitle, callout) | 18-20 | 15-16 |
| Body | 16-18 (never below 14) | 14-16 (never below 14) |
| Label, chart axis, data label | 12-14 | 11-12 |
| Caption, source, footnote | 10-12 | 10-11 |

Keep the ratios, not the points: title ≈ 2-2.5x body, label ≈ 0.7-0.8x body. The 10-inch canvas is the same slide projected onto the same screen, so its points run ~75% of the wide values — except body, which never goes below **14 pt** (the skill's house-rule floor; the generic "body 14-16 pt" in `SKILL.md` is the 10-inch column). One deck uses 3-4 sizes; more reads as accident.

### Safe fonts (metric-compatible)

| Deck font | Local preview substitutes | Metric-true? |
|---|---|---|
| Calibri | Carlito | yes |
| Cambria | Caladea | yes |
| Arial | Arimo | yes |
| Times New Roman | Tinos | yes |
| Courier New | Cousine | yes |
| Aptos | Noto Sans | no (widths differ) |
| BIZ UDGothic / Noto Sans JP | (installed) | yes [measured] |
| Yu Gothic / Meiryo | Noto Sans / IPAexGothic | no — leave ~10% width slack |

Latin body: Calibri, Arial (or Cambria/Times for academic prose). Japanese: BIZ UDGothic or Noto Sans JP. Anything else makes the QA render lie about wrapping, so treat the overflow check as approximate.

### Line length and box arithmetic

Measured average advance for mixed-case English prose [measured]: **0.0058 x fontSize inches per character** (Calibri 14 pt: 0.0810; 16 pt: 0.0920; 18 pt: 0.1039; Cambria 16 pt: 0.0917 — within 1% of each other). So:

```
chars per line ~ 172 * box_width_in / fontSize_pt
```

| Box width | 14 pt | 16 pt | 18 pt |
|---|---|---|---|
| 5.0 in (narrow column) | 61 | 54 | 48 |
| 6.0 in | 74 | 65 | 57 |
| 8.0 in | 98 | 86 | 76 |
| 12.3 in (full width) | 151 | 132 | 118 |

Aim for **45-75 characters** for slide body text (prose norms are 45-90 [B1]); a full-width body line at 16 pt is 132 characters and is unreadable from the back — that is what two columns are for. Titles: <= 2 lines, <= 50 characters. Japanese: one full-width character = one em = `fontSize/72` in, so a 6 in column holds only **27 characters at 16 pt** [measured: a 37-character JP line at 16 pt measured 7.97-8.08 in in BIZ UDGothic / Noto Sans JP].

Box height, so overflow never surprises the QA pass [measured]: line pitch is **1.20 x fontSize** (16 pt -> 19.19 pt = 0.267 in) and `lineSpacingMultiple: 1.15` multiplies that (22.05 pt = 1.38 em). `paraSpaceAfter` adds its value in **points** between paragraphs (8 -> pitch 27.18 pt at 16 pt).

```
box_height_in = (lines * 1.20 * fontSize_pt * lineSpacingMultiple + (paras-1) * paraSpaceAfter) / 72
```

### Emphasis

- **Bold** for the one phrase that carries the takeaway; color the same phrase in the accent only if it is the slide's single accent use. Do not do both on every bullet.
- Italics: for Latin publication titles and terms of art. Japanese text does not italicize well — use bold, or bracket the term (「」).
- ALL CAPS: only for 3-6 word labels, with `charSpacing: 6` (the real option name; `letterSpacing` is ignored [P1]). No all-caps sentences or headings — Microsoft's accessible-font guidance is to avoid all capitals and excessive italics or underlining [M1].
- Numbers: set key numbers 2-3x body size, in ink or accent, with the label in muted below. Keep one decimal format per deck.
- Never shrink to fit (`fit: "shrink"`): reduce words or split the slide instead.

## 3. Spacing and alignment

- **Outer margin 0.5 in** on all four sides (both canvases). Content box on the wide canvas: `x 0.5 -> 12.83`, `y 0.5 -> 7.0`. The QA rule flags any margin under 0.5 in and any gap under 0.3 in [S1].
- **Block gaps 0.3-0.5 in**: 0.3 in is the floor between sibling blocks, 0.4 in between rows of a grid, 0.5 in between major regions (title band / body / footer).
- **Vertical bands** (wide canvas): title band `y 0.5, h 0.9` (a 36 pt line is 1.2 x 36 / 72 = 0.6 in, so a two-line title needs `h 1.2` and pushes body to `y 2.1`); body `y 1.8 -> 6.4`; footer/source line `y 6.7, h 0.3` — every band bottom respects the 0.5 in margin.
- **Six-column grid**, wide canvas: column 1.89 in, gutter 0.2 in (6 x 1.89 + 5 x 0.2 = 12.34). Practical widths: 2 columns 5.92 in, 3 columns 3.84 in, 4 columns 2.78 in, with 0.4 in gutters.
- **Text aligned with shapes:** set `margin: 0`. Default text inset is 0.1 in (`lIns`) — unmanaged it offsets text 0.098 in from the shape it should align with [measured]. `margin` is in **points**, not inches: `margin: 36` = 0.5 in, `margin: 0.5` = 0.5 pt [measured].
- **Vertical placement:** pptxgenjs defaults `valign` to `"middle"` [P1], which silently centers text inside a taller box. Set `valign: "top"` for anything that must line up with a shape above it: with `valign: "top"` + `margin: 0` the rendered glyph tops land 0.6-1.5 pt below the box top (14-32 pt) [measured], i.e. box edges are the visual edges. Use `valign: "middle"` when centering text beside a shape, and then size the box to the shape.
- **Equalizing paired elements:** give paired cards the same `h` and the same `y`; equalize stat callouts by `valign: "middle"` + equal `h`, not by nudging `y`. Build each option object from a factory function — pptxgenjs mutates option objects on the text path [P1], so sharing one object between two calls makes the second call inherit the first's computed fields.
- **Whitespace budget:** aim for ~20% of the slide area empty; a slide with no gap >= 0.3 in between blocks reads as a wall.
- **QA hooks:** a design system is what makes the QA pass checkable — margins >= 0.5 in, gaps >= 0.3 in, box heights from the §2 pitch formula, and one text/background pair per palette from §1. Fix violations in the generator, then re-render (`SKILL.md` QA loop, `scripts/render.sh`).

## 4. Layout patterns

All coordinates are `LAYOUT_WIDE` with 0.5 in margins (`x 0.5, w 12.33`). For the 10 in canvas multiply every length by **0.75** (margins stay 0.5; text sizes follow §2).

| Pattern | Geometry recipe |
|---|---|
| Two-column | Title band; left `x 0.5 w 5.92`; right `x 6.92 w 5.91`; both `y 1.8 h 4.6`. Use when the left column is text and the right is evidence. |
| 2x2 grid | Columns 5.92 wide (gutter 0.5); rows `y 1.8 h 2.6` and `y 4.8 h 2.2` (gutter 0.4). One idea per cell, cell header 16-18 pt bold, body 14 pt. |
| Stat callout row | 3-4 equal cells at `y 2.4 h 1.6`: `w = (12.33 - (n-1) * 0.4) / n` -> 3.84 (n=3), 2.78 (n=4); `x = 0.5 + i * (w + 0.4)`. Number 40-54 pt in accent, label 14 pt muted, optional one-line source 10 pt. |
| Half-bleed image + overlay | Image `x 6.5 w 6.83 h 7.5 y 0` (bleeds to the right/top/bottom edges); text panel `x 0.5 w 5.5 y 1.8`. Text panel never overlaps the image; if it must, use the scrim rule from §1. |
| Full-bleed image, one line | Image `x 0 w 13.33 h 7.5`; black scrim `transparency: 45` over it; single line 32-40 pt white at `y 3.2 h 1.2`, left-aligned at `x 0.5`. One sentence only. |
| Timeline | Axis line `x 0.5 w 12.33 y 3.9`; n nodes centered in equal cells, `cell = 12.33 / n`, `node_center_x = 0.5 + (i + 0.5) * cell`; date label above (`y 3.2, h 0.4`, 12 pt muted), event text below (`y 4.2, h 1.4`, 14 pt, max 20 words per node). Cap at 5-6 nodes. |
| Comparison columns | 2-3 columns from the grid; header band `h 0.6` filled with the tint (ink text), then rows on a shared pitch: `y = 2.6 + r * 0.8`. Same row count on both sides or the comparison dies. |
| Section divider (sandwich) | Dark surface, accent kicker 16 pt at `y 2.6`, section title 40-44 pt at `y 3.0 h 1.4`, nothing else. |

Vary the pattern with the message: claims get two-column + evidence, processes get a timeline, comparisons get columns, single numbers get a stat row. Two consecutive slides with the same geometry is the most common AI-deck tell.

## 5. Genre notes

Full structure, length and tone per genre: `references/deck-types.md`. Styling deltas only:

- **Consulting / strategy** — palettes 1 or 7. One accent, gray support, no decorative imagery; 20% whitespace; source line and slide number on every slide; action titles in sentence case. Enable `p.addSection()` for the section spine.
- **Academic** — palette 3 (near-black on white, maximum contrast). Figure-first: the figure is the slide, text is a caption. Body >= 18 pt, <= 6-7 lines, figures labelled at >= 16 pt. 3 colors total and never red-vs-green alone; series colors from §1 with distinct lightness.
- **Product / pitch** — palette 4 or 2, dark stage. One idea per slide at >= 30 pt; the accent is used once (the number or the keyword), everything else is ink on surface. Show the product, not a feature list.
- **Internal status** — palette 5 with the status trio (`1B7F4B` / `8A5A00` / `B3261E`) plus a text cue per item; one status call at the top of slide 1. Keep it short and factual; no photograph backgrounds.

## 6. Accessibility

**Alt text (`altText`).** Verified on pptxgenjs 4.0.1 by writing a deck and inspecting `ppt/slides/slideN.xml`:

- The option name is `altText` and it works on **`addImage()`** and **`addChart()`** only, landing on the shape's `<p:cNvPr descr="...">`. `alt`, `description`, `title` and `decorative` emit no `descr` at all (verified with all four set on one image) [measured].
- **An image without `altText` gets its file path as the description** (`descr="/home/you/project/figure.png"`) — a screen reader reads your filesystem. `altText: ""` falls back to the path too, so there is no "decorative" switch in the API.
- `addShape()`, `addText()` and `addTable()` ignore `altText` entirely (their `<p:cNvPr>` gets no `descr`). For decorative-vs-meaningful control on those, post-edit the XML: `<p:cNvPr id="4" name="Shape 0" descr="Accent block for the stat callout">`. A hand-added `descr` survives unpack/pack and is readable via python-pptx `shape._element._nvXxPr.cNvPr.get("descr")` [measured]; `&`, `"`, `<` are escaped by pptxgenjs itself.
- Writing rules: describe content and purpose, not appearance; no "image of"/"chart of"; <= 1-2 short sentences; never restate the title [M1]. For a chart, give the takeaway ("Revenue doubles Q1->Q2; Q3 flat"), not the type.
- Tables are read cell by cell, so keep them simple, never merged, with a header row [M1]; a table needs no alt text, a chart and every figure do.

**Reading order.** A screen reader reads slide objects in shape-tree order, which is the order `add*()` calls were made [M1] — pptxgenjs writes `cNvPr` ids in insertion order [measured]. So add title, then body, then evidence, then decoration; reorder afterwards in PowerPoint's Reading Order pane if you must.

**Slide titles.** Give every slide a unique, descriptive title; the checker looks for a title *placeholder*, which plain `addText()` does not produce. Verified recipe: declare it on the master and fill it by name:

```js
p.defineSlideMaster({
  title: "BODY",
  objects: [{ placeholder: { options: { name: "title", type: "title", x: 0.5, y: 0.4, w: 12.33, h: 0.9 }, text: "" } }],
});
const s = p.addSlide({ masterName: "BODY" });
s.addText("Renewals offset a weak Q1", { placeholder: "title", x: 0.5, y: 0.4, w: 12.33, h: 0.9,
  fontFace: "Calibri", fontSize: 36, color: "1F2933", margin: 0 });   // -> <p:ph type="title"/>
```

The layout then carries `<p:ph idx="100" type="title"/>` and python-pptx reports the shape as `TITLE` placeholder [measured]. `[INFERENCE]` PowerPoint's Accessibility Checker accepts that as the slide title — it cannot be run on this machine.

**Color-vision safety.** Never use color as the only signal [W3]: status chips carry a word, chart series carry distinct lightness (3:1 between series that must be told apart) plus a legend and data labels, and "before/after" pairs differ in value, not just hue. Check the deck by rendering it and looking at the grays: 1.07:1 status colors disappear, the palettes in §1 do not.

**Minimum sizes and density.** Audience-read text >= 14 pt (>= 18 pt for dyslexia-friendly or large-room decks [M1]); sources/footnotes 10-12 pt and only for provenance. Line pitch 1.2 x fontSize keeps adjacent lines from touching; 45-75 characters per line (§2).

## 7. Do not

- **Accent stripes, title underlines, header/footer bars, edge bands.** They carry no information, they eat the only accent on the slide, and they break the moment a title wraps to two lines (the mark is sized for one). Emphasis goes in the words: bold the keyword, or set the title in the accent color.
- **Centered body text.** Centering forces the eye to find a new start point on every line; reserve centering for the cover title and single-line callouts. Body and bullets are left-aligned, ragged right.
- **One layout for the whole deck.** Identical geometry on every slide turns structure into noise and hides the slides that matter. Change the pattern at every change of message type (§4).
- **Color-only status or series coding.** Traffic-light colors are 1.07-1.30:1 apart in grayscale — a printed copy or a red-green color-vision viewer sees one gray.
- **Text over an unmanaged photo.** A 50% scrim over a bright image gives white body text 3.95:1 — under the floor. Scrim at 45 or use a solid band.
- **Decorative icon-card grids and progress bars.** Three boxes with generic icons say "this deck had nothing to show"; spend the space on the number, the chart or the sentence.
- **Shrink-to-fit and font size 12 body text.** `fit: "shrink"` cannot report that it shrank 40 pt to 9 pt; the room notices. Cut words, split the slide.
- **Mid-gray surfaces behind paragraphs** (`767676`-`8E8E8E`): no text color reaches 4.5:1 on them (§1).
- **Emoji as icons** and mixed accent colors per slide. Emoji render differently on every platform and carry no accessible name; a second accent destroys the 60-30-10 balance that makes the first one read as emphasis.

## Sources

- [W1] W3C — Understanding SC 1.4.3 Contrast (Minimum): 4.5:1 / 3:1 thresholds, ≥18 pt or ≥14 pt bold, no rounding — https://www.w3.org/WAI/WCAG21/Understanding/contrast-minimum.html
- [W2] W3C — Technique G18 (relative-luminance and contrast-ratio formula, sRGB threshold 0.04045) — https://www.w3.org/WAI/WCAG21/Techniques/general/G18
- [W3] W3C — Understanding SC 1.4.1 Use of Color (hue + ≥3:1 lightness difference counts as a non-color distinction) — https://www.w3.org/WAI/WCAG21/Understanding/use-of-color.html
- [W4] W3C — Understanding SC 1.4.11 Non-text Contrast (3:1 for graphics required to understand content) — https://www.w3.org/WAI/WCAG21/Understanding/non-text-contrast.html
- [M1] Microsoft — Make your PowerPoint presentations accessible to people with disabilities (alt text on shapes/charts, reading order, 18 pt sans-serif, table structure, contrast checker limits) — https://support.microsoft.com/en-us/accessibility/powerpoint/make-your-powerpoint-presentations-accessible-to-people-with-disabilities
- [B1] Butterick — Practical Typography: Line length (45-90 characters, 2-3 alphabets) — https://practicaltypography.com/line-length.html
- [S1] siril9/presentation-skill (MIT) — visual QA prompt: gaps < 0.3", edge margins < 0.5", low-contrast text/icons, wrapped-title decoration mismatch — https://raw.githubusercontent.com/siril9/presentation-skill/main/references/visual_qa_prompt.md
- [P1] PptxGenJS 4.0.1 behavior, verified locally — see `references/pptxgenjs.md` (`margin`/`valign`/`charSpacing` semantics, option mutation)
- Measurements, scrim compositing, `altText` and title-placeholder probes: run on this machine (pptxgenjs 4.0.1, LibreOffice, poppler, python-pptx 1.0.2) — research probes under `research/probes/` in the skill repo.
