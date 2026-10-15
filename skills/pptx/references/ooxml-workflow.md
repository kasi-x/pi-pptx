# Editing OOXML: existing decks and templates

Use this when the task is to modify an existing `.pptx`, or to fill a template — building from scratch with pptxgenjs is easier only when there is no house style to match.

## Unpack / pack

```bash
scripts/unpack.sh deck.pptx unpacked/
# ... edit ...
scripts/pack.sh unpacked/ out.pptx
```

`pack.sh` zips from **inside** the directory and deletes the target first, so removed parts do not survive in the output.

## Structure map (what lives where)

| Part | Role |
|---|---|
| `ppt/presentation.xml` | deck root; `<p:sldIdLst>` defines **which slides and in what order**; `<p:sldSz>` canvas |
| `ppt/slides/slideN.xml` | one slide's shapes/text |
| `ppt/slides/_rels/slideN.xml.rels` | each slide's relationships (images, charts, layouts, notes) |
| `ppt/slideLayouts/`, `ppt/slideMasters/` | theme + placeholders that slides inherit |
| `ppt/media/`, `ppt/charts/`, `ppt/notesSlides/` | images, chart parts (+ embedded xlsx), speaker notes |
| `[Content_Types].xml` | one `Override` entry per slide/chart/notes part |
| `ppt/_rels/presentation.xml.rels` | `rId` → slide part mapping referenced by `<p:sldId>` |

## Structural operations (always before content edits)

- **Duplicate a slide** — `scripts/dup_slide.py` (copies the part, its rels, and writes the content-type / presentation-rels / `<p:sldIdLst>` registrations):

  ```bash
  scripts/dup_slide.py unpacked/ ppt/slides/slide2.xml --after ppt/slides/slide2.xml
  scripts/dup_slide.py deck.pptx ppt/slides/slide2.xml --out out.pptx   # direct on a file
  ```

  A duplicated slide **references the source's** chart/media parts — editing one slide's chart changes the copy. Replace the part with a fresh one if they must diverge.
- **Delete a slide** — remove its `<p:sldId>` entry from `<p:sldIdLst>`. PowerPoint then ignores the orphan part; optionally delete `ppt/slides/slideN.xml`, its `.rels`, the `[Content_Types].xml` override, and the presentation rel to shrink the file.
- **Reorder** — the order of `<p:sldId>` entries *is* the slide order.
- Keep ids unique: `<p:sldId id>` ≥ 256, and never reuse an existing one.

## Content edits

- One `<a:p>` per paragraph. For list items, set bullets via `<a:buChar>` / `<a:buAutoNum>` / `<a:buNone>` on `<a:pPr>` — never a literal `•` in the text (renders doubled).
- Preserve formatting by copying sibling `<a:pPr>` / `<a:rPr>` elements from an existing run and editing only the text nodes.
- Add `xml:space="preserve"` to `<a:t>` when text has leading/trailing spaces.
- Never round-trip the XML through `xml.etree.ElementTree` without preserving namespace declarations — it rewrites prefixes/attributes and corrupts decks. Use `defusedxml.minidom` / `lxml`, surgical string edits, or the pattern `dup_slide.py` uses (collect the `start-ns` declarations with `iterparse` and re-register every prefix before writing).
- Never reorder the children of `<p:presentation>` or renumber relationship ids.
- With **python-pptx**: assign `run.text` (never `text_frame.text`, which collapses styling to one run). python-pptx cannot duplicate slides or parse most template art (SVG/EMF).
- **New media**: put the file in `ppt/media/`, add a relationship in the slide's `.rels`, and add a content-type default/override if the extension is new.
- **Charts** live in `ppt/charts/chartN.xml` with an embedded workbook; heavy data changes are usually faster and safer by regenerating that slide (or the chart) with pptxgenjs and re-inserting it.

## Template workflow (fill a `.potx` / house deck)

1. Render the template first — `scripts/render.sh template.pptx` works on `.potx` too — and pick the layouts/slides to copy.
2. Map each content section onto a template slide; use `scripts/dup_slide.py` to stamp out one copy per section.
3. Replace placeholder text only inside the placeholder shapes; when a repeated element no longer applies (e.g. a 4th team member with no data), delete the whole group (image + text boxes), not just the text.
4. Keep the `.potx` extension on the output.
5. Check for leftovers: `markitdown out.pptx | grep -iE "lorem|ipsum|TODO|xxx|\[insert|placeholder"`.

## Verify after every structural change

```bash
scripts/inspect_deck.py out.pptx        # opens + counts
scripts/render.sh out.pptx              # render every slide and look at them
```
