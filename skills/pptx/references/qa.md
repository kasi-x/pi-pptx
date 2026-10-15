# QA playbook for decks

The generator (or the source XML) is the source of truth: fix defects there, never in the rendered output. And assume defects exist — the author of a deck sees their intent, not the pixels. Run the loop.

## The loop, in order

1. **Static pass** — `python3 scripts/lint_deck.py out.pptx` (geometry, fonts, overflow heuristics, placeholder text, image resolution). Cheap; catches most mechanical faults before anything is rendered.
2. **Content pass** — `markitdown out.pptx` — read it like a stranger: missing sections, wrong order, typos; then `markitdown out.pptx | grep -iE "lorem|ipsum|TODO|xxx|\[insert"` for template leftovers.
3. **Render pass** — `scripts/render.sh out.pptx` and look at **every** image (or hand the folder to a fresh subagent — prompt below). Rendered defects that no static check sees: text clipped by its own box, one element stacked under another, decoration positioned for a title that wrapped to two lines.
4. **Rendered-text pass** — deterministic overflow check without eyes:

   ```bash
   pdftotext -bbox-layout render/out.pdf - | grep -E 'xMin="(|-)' # or parse the boxes:
   ```
   Compare word boxes against the page box minus ~36 pt (0.5 in); anything crossing an edge is clipped or spilling.
5. **Fix → re-render only what changed → repeat.** Cap the loop: after ~3 iterations, cut content instead of shrinking type.

## Visual QA with a fresh subagent (copy-paste prompt)

```
Review these rendered slide images of a deck: <paths>.
Assume there are defects. For EACH slide, report findings as
blocker / major / minor, with slide number and, where possible, the
element and measured problem. Check, in this order:
- text overflow or clipping; text touching its box edge
- overlaps (text through shapes, elements stacked, decorations colliding)
- margins < 0.5 in from any edge; block gaps < 0.3 in, or uneven gaps
- low contrast (text or icons vs their background); meaning that is
  carried only by color
- misalignment: columns or rows that should share an edge but do not;
  inconsistent left edges; uneven baselines
- leftover placeholder/ghost text; duplicated decorations
- title/content mismatch (the title claims something the slide does not show)
- figure legibility: smallest text in charts/tables at real size
End with the top 3 fixes ranked by user-visible impact. Do not answer
"looks fine" — list what you checked per slide.
```

Use a subagent (or a second model pass) because the deck's author cannot see its own defects.

## Numeric thresholds (acceptance)

| Item | Rule |
|---|---|
| Outer margin | ≥ 0.5 in on every edge (full-bleed art excepted) |
| Block gaps | 0.3–0.5 in, consistent within a slide |
| Body text | ≥ 14 pt projected, ≥ 16–18 pt presented; captions ≥ 10–12 pt |
| Contrast | ≥ 4.5:1 body text, ≥ 3:1 for ≥ 24 pt text and icons (WCAG 2.1) |
| Charts | bars start at 0; value axis not truncated; ≤ 5 pie slices; one message per chart |
| Images | effective ≥ 150 dpi (pixels ÷ display inches) |
| Slide budget | ~1 slide/minute presented; exec summary 1 slide (consulting) |

Contrast references: WCAG 2.1 SC 1.4.3/1.4.11. The gap/margin thresholds follow common deck-QA practice (open-source examples: siril9/presentation-skill, EveryInc/hands-on-deck).

## Incremental review (did my fix break something else?)

- Lint diffs: `lint_deck.py baseline.pptx > a.txt; lint_deck.py out.pptx > b.txt; diff a.txt b.txt` — fix only new/worsened findings, keep known leftovers documented.
- Pixel diffs between two renders: `magick compare -metric RMSE old-1.jpg new-1.jpg null:` per slide — non-zero on slides you touched, zero elsewhere.

## Deliverable checklist

- [ ] Deck opens in python-pptx (`inspect_deck.py`), lints clean (or: every remaining finding explained), and renders with slide-count == image-count.
- [ ] Placeholders gone; speaker notes carry the talk track (`slide.addNotes`) — not slide text.
- [ ] Fonts from the safe list, or width slack left for substituted fonts (see `design-system.md`).
- [ ] The generating script is saved next to the deck (regenerable, deterministic); seed/inputs noted.
- [ ] If a PDF/PNG export is part of the delivery, it was exported **after** the last fix.

## Genre-specific acceptance (short)

- **Consulting:** answer-first — executive summary holds the full answer; titles are action titles; deck readable alone.
- **Academic:** one finding per slide; figures legible at print size; backup slides ready for Q&A.
- **Product/pitch:** one idea per slide; the problem sells, not the tech; competition slide present; demo shows real output.
- **Internal:** status visible on slide 1; every risk has a response; owners and dates on actions.
