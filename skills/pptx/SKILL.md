---
name: pptx
description: "Create, edit, read, and visually QA PowerPoint (.pptx/.potx) files. Use whenever a .pptx/.potx file is involved as input or output, or when the user mentions a deck, slides, a presentation, or preparing materials (プレゼン, スライド, 発表資料, 資料作成). Covers: building decks from scratch with pptxgenjs, editing existing decks or templates via OOXML, extracting text, rendering slides to images for visual review, and audience-specific playbooks (consulting / academic / product / internal)."
license: MIT
compatibility: "node + pptxgenjs (npm) for creation; python3 + python-pptx for inspect; markitdown[pptx] for text dumps; LibreOffice (soffice) + poppler (pdftoppm) for rendering"
---

# PPTX toolkit

A `.pptx` is a zip archive of XML parts. Every task below is either "emit one with a script" or "edit the XML inside one". Route by task first:

| Task | Approach |
|---|---|
| Create a new deck | `pptxgenjs` script — how-to in `references/pptxgenjs.md`, starter in `examples/deck.js` |
| Edit / fill an existing deck or template | unpack → edit slides + bookkeeping → repack — `references/ooxml-workflow.md` |
| Read content | `markitdown deck.pptx`, or `scripts/inspect_deck.py deck.pptx` |
| Visual QA | `scripts/render.sh deck.pptx`, then look at every image |
| Choose structure & style | `references/deck-types.md` (consulting / academic / product / internal / training) |

Paths are relative to **this skill's directory** (the harness prints it when the skill loads; `scripts/` and `references/` live next to this file). Load only the reference you need.

## Environment check (once per session)

```bash
node -e "require.resolve('pptxgenjs')"   # creation; install if missing: npm install pptxgenjs
python3 -c "import pptx"                  # scripts/inspect_deck.py
command -v soffice pdftoppm markitdown    # rendering / text dump (optional but recommended)
```

If `require('pptxgenjs')` fails: `npm install pptxgenjs` in the working project (or `NODE_PATH` to a `node_modules` that has it).

## Create a deck (pptxgenjs)

1. Pick the genre and outline first: `references/deck-types.md`.
2. Fix the design system **before** writing slides: one dominant color + support + one accent; white or brand background; titles 36–44pt, section heads 20–24pt, body 14–16pt, captions 10–12pt.
3. Write the generator — copy `examples/deck.js` as the skeleton.
4. Generate → QA → fix loop (see QA below).

Top footguns (full list, verified notes and chart recipes in `references/pptxgenjs.md`):

- Set `pres.layout` (`LAYOUT_WIDE` = 13.33″×7.5″) **before** adding slides; off-canvas content is silently lost, not clamped.
- Colors are hex strings **without `#`** and without alpha digits — both corrupt the file; use `transparency` (fills/images) or `opacity` (shadows) for see-through.
- pptxgenjs **mutates option objects in place**; build a fresh object per `add*` call and never share one `shadow` object between calls.
- Charts: anything PowerPoint can chart natively goes through `addChart()` — never a screenshot of a chart. Set titles, data labels, palette and quiet axes explicitly or charts render bare.
- One `new pptxgen()` per output file; `writeFile()` at the end.
- Notes go to `slide.addNotes()`; bullets via `bullet: true` (never a literal `•`); the real letter-spacing option is `charSpacing`.

## Edit a deck or template

```bash
scripts/unpack.sh deck.pptx unpacked/
# edit ppt/slides/slideN.xml; reorder/delete via <p:sldIdLst> in ppt/presentation.xml
scripts/dup_slide.py unpacked/ ppt/slides/slide2.xml --after ppt/slides/slide2.xml
scripts/pack.sh unpacked/ out.pptx
```

Rules that prevent most corruption and rework:

- Do all **structural** work (add / duplicate / delete / reorder) before editing content; duplicated slides clone whatever content existed at copy time.
- `dup_slide.py` writes every registration a new slide needs. A duplicated slide **shares** its chart/media parts with the source (editing one changes the other).
- Keep `<p:presentation>` child order untouched; never round-trip OOXML through `xml.etree.ElementTree` (namespace prefixes get rewritten) — use `defusedxml.minidom`, `lxml`, or surgical edits.
- One `<a:p>` per paragraph/list item; add `xml:space="preserve"` when text has leading/trailing spaces; keep sibling `<a:rPr>`/`<a:pPr>` to preserve formatting. With python-pptx, assign `run.text`, never `text_frame.text`.
- `.potx` templates are the same zip format — keep the extension on the output.
- Details and a checked edit script: `references/ooxml-workflow.md`.

## Read

- `markitdown deck.pptx` — one text block per slide (good for content checks).
- `scripts/inspect_deck.py deck.pptx` — opens the file, prints slide count, shapes/charts/tables/pictures, notes and a text preview.

## QA (required before delivering)

1. `scripts/render.sh out.pptx` → look at **every** image. Check in this order: text overflow/clipping (most common), overlaps, gaps < 0.3″, margins < 0.5″, low contrast, misalignment, leftover placeholder text.
2. Deterministic overflow spot-check: `pdftotext -bbox-layout <outdir>/out.pdf -` and scan for word boxes crossing the page edges (page box minus ~36 pt ≈ 0.5″).
3. `markitdown out.pptx | grep -iE "lorem|ipsum|TODO|xxx|\[insert"` — template leftovers.
4. Fix in the generator or source XML, re-render only changed slides, repeat until clean.

## Dependencies

| Need | For | Install |
|---|---|---|
| node ≥ 18 + `pptxgenjs` ^4 | deck creation | `npm install pptxgenjs` in the project |
| python3 + `python-pptx` | `scripts/inspect_deck.py` | `pip install python-pptx` (or a tool runner: `uv tool`) |
| `markitdown[pptx]` | text dumps | `uv tool install "markitdown[pptx]"` |
| LibreOffice + poppler (`soffice`, `pdftoppm`) | `scripts/render.sh` | distro packages: `libreoffice poppler-utils` |

## House rules

- Native-first: charts, tables, and diagrams PowerPoint can build stay native parts, not images.
- Body text ≥ 14pt and high contrast — the back row must be able to read it.
- Give every slide a visual element; vary the layout between slides instead of repeating one template.
- No decorative header/footer bars, title underlines, or edge stripes — use whitespace, tints, shadows, and images to structure content.
- Japanese decks: use fonts that exist on the target machine (`BIZ UDGothic`, `Noto Sans JP`, `Yu Gothic`); the LibreOffice preview substitutes fonts, so leave ~10% width slack on JP text boxes.
