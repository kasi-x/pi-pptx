# Native equations (LaTeX → OMML)

`scripts/inject_omml.py` replaces `{{MATH:id}}` placeholders in a generated
`.pptx` with **native, editable PowerPoint equations** (OMML wrapped in
`a14:m`). It needs `pandoc` at runtime; the Python side uses only the standard
library (`xml.etree.ElementTree` + `zipfile`).

For maximum portability (LibreOffice, Google Slides, older PowerPoint), use the
image fallback described below instead.

## Workflow

1. Put a placeholder in the outline text that `outline_to_pptx.js` renders.
   PptxGenJS may split the marker across several `<a:r>` runs — the injector
   concatenates all run text in a paragraph before matching, so that is fine.

   ```json
   {
     "meta": { "title": "Physics recap" },
     "slides": [
       {
         "type": "bullets",
         "title": "Two classic results",
         "bullets": [
           "Mass–energy equivalence: {{MATH:eq1}}",
           "Kinetic energy: {{MATH:eq2}}"
         ]
       }
     ]
   }
   ```

2. Generate the draft deck:

   ```bash
   node scripts/outline_to_pptx.js outline.json draft.pptx
   ```

3. Write `formulas.json` (path relative to the outline or anywhere you like):

   ```json
   [
     { "id": "eq1", "latex": "E=mc^2", "mode": "display" },
     { "id": "eq2", "latex": "\\frac{1}{2}mv^2", "mode": "inline" }
   ]
   ```

   Remember that JSON strings need escaped backslashes (`\\frac`, `\\alpha`).

4. Inject and keep the result:

   ```bash
   python3 scripts/inject_omml.py draft.pptx formulas.json final.pptx
   # {"success": 2, "failed": 0, "skipped": 0}
   ```

## `formulas.json` schema

Top level: a JSON array of objects.

| Field | Required | Values |
|---|---|---|
| `id` | yes | `[A-Za-z0-9_-]+`; must match the `{{MATH:id}}` marker in the deck |
| `latex` | yes | non-empty LaTeX math source (no `$` delimiters needed) |
| `mode` | no | `"display"` (default, wraps in `$$…$$`) or `"inline"` (wraps in `$…$`) |

Invalid entries are reported on stderr, counted as `failed`, and their
placeholders are left untouched. Unknown `{{MATH:id}}` markers (no matching
entry) are also left untouched.

## Exit status and stats

On success the command prints one JSON object with three integer fields:

- `success` — number of placeholder occurrences replaced (an id can appear on
  several slides or several times in one slide).
- `failed` — formulas that were invalid or could not be converted to OMML;
  their placeholders remain in the deck.
- `skipped` — formulas that converted successfully but had no matching
  placeholder in any `ppt/slides/slideN.xml`.

Partial success is normal: conversion failures do not fail the command (exit
`0`), so inspect stderr and re-run after fixing the formula. A missing or
unusable `pandoc` is fatal (exit `1`) with an install hint.

## Requirements and caveats

```bash
brew install pandoc        # macOS; apt-get install pandoc on Debian/Ubuntu
```

- The injector adds `xmlns:a14`, `xmlns:m`, `xmlns:mc` and appends `a14` to
  `mc:Ignorable` on each slide it changes. No extra Python packages needed.
- Keep one `{{MATH:id}}` inside a single paragraph; markers split across two
  paragraphs (or two text boxes) are not matched.
- Native OMML is supported by PowerPoint 2010 and later. LibreOffice and
  Google Slides may ignore `a14:m` and render nothing — use the image fallback
  for those targets.
- pandoc does **not** fully support LaTeX → OMML for every construct. Custom
  macros, some `amsmath` environments, TikZ, and unusual packages may fail or
  convert imperfectly. Prefer standard math syntax, check stderr, and visually
  QA every equation in PowerPoint before delivery.

## Image fallback (no native equation)

When native equations are unavailable or unsupported, render the formula to a
PNG and place it on a `figure` slide:

```json
{
  "slides": [
    {
      "type": "bullets",
      "title": "Mass–energy",
      "bullets": ["A classic result"],
      "assets": [
        { "type": "latex", "id": "eq1", "mode": "display", "code": "E=mc^2" }
      ]
    },
    { "type": "figure", "title": "Mass–energy", "image": "eq1", "width": 5.0 }
  ]
}
```

```bash
python3 scripts/render_assets.py outline.json assets/
node scripts/outline_to_pptx.js outline.json out.pptx
```

This path needs `pandoc` plus a LaTeX engine (`pdflatex`/`latex`) and a PDF
rasterizer (`pdftocairo`/`pdftoppm`); see `references/assets.md`. The assets
are pictures, so they stay readable everywhere but are not editable equations.
