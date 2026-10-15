# Changelog

All notable changes to this project are documented here. The format follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/), and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [0.2.0] - 2026-10-15

### Added

- `references/charts.md` — chart-type chooser, tested pptxgenjs recipes (incl. combo with secondary axis, waterfall), option tables, failure modes.
- `references/design-system.md` — color/typography/spacing system, palettes with contrast checks, layout patterns, accessibility, anti-cliche list.
- `references/qa.md` — QA loop, numeric thresholds, copy-paste visual-QA subagent prompt, deliverable checklist, genre acceptance.
- `scripts/lint_deck.py` — geometry/typography linter (margins, overlaps, tiny gaps, fonts, overflow heuristic, image DPI, placeholder text; optional rendered-PDF overflow pass).
- `scripts/outline_to_pptx.js` + `examples/outline.sample.json` — deterministic outline→deck renderer (title/section/bullets/stats/chart/table/closing; 3 palettes).
- `examples/master-deck.js` — branded slide master with title/body placeholders, sections, slide numbers.

### Changed

- `SKILL.md` routing now covers charts, design system, QA playbook and the linter; house rules add regenerable-generator discipline.
- `npm test` now also lints the generated example deck.

## [0.1.0] - 2026-10-15

### Added

- Initial release: `SKILL.md`, `references/pptxgenjs.md`, `references/ooxml-workflow.md`, `references/deck-types.md`.
- Scripts: `unpack.sh`, `pack.sh`, `dup_slide.py`, `render.sh`, `inspect_deck.py`.
- `examples/deck.js` starter; pi-package manifest; MIT license.
