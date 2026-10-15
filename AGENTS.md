# pi-pptx — AGENTS.md

## Overview

Original, MIT-licensed PPTX skill (Agent Skills format) for pi and OMP, packaged as a pi package: create decks with pptxgenjs, edit PPTX/OOXML, render + visually QA slides, and follow genre playbooks (consulting / academic / product / internal). The live skill copy for OMP/Claude Code is symlinked from `~/.agents/skills/pptx`.

## Commands

```bash
npm test                                            # generate example deck, inspect + lint it
bash skills/pptx/scripts/render.sh example-deck.pptx   # visual QA renders (then LOOK at them)
python3 skills/pptx/scripts/dup_slide.py --help     # slide duplication
python3 skills/pptx/scripts/lint_deck.py --help     # geometry/typography lint
pi install ~/dev/pi-pptx                            # (re)register with pi
```

Publish (author's call, not automated): `gh repo create kasi-x/pi-pptx --public --source . --push`, then optionally `npm publish`.

## Layout

- `skills/pptx/SKILL.md` — routing + guardrails (keep ≤ ~150 lines)
- `skills/pptx/references/` — pptxgenjs, charts, design system, OOXML workflow, QA playbook, deck types (with sources)
- `skills/pptx/scripts/` — unpack/pack, dup_slide, render, inspect, lint (python3 stdlib + python-pptx or POSIX bash), outline_to_pptx (node + pptxgenjs)
- `skills/pptx/examples/` — deck.js, master-deck.js, outline.sample.json (must keep running via `npm test`)
- `research/` — local-only working notes (gitignored, not shipped)

## Constraints

- **Original content only.** Never copy text or scripts from Anthropic's source-available pptx skill (its license forbids copying, derivative works, and redistribution). Public-documentation facts are fine — cite the source in the reference.
- Keep runtime deps thin: deck creation uses pptxgenjs (MIT); QA scripts use python3 stdlib / python-pptx / LibreOffice / poppler.
- After changing scripts or the example: run `npm test` and the render loop, and look at the images.
- No emoji in shipped files.
- Tainted sources (do not copy text/code from): `MiniMax-AI/skills`, `anyideaz/pptx-skills`, `msxabc/powerpoint-skills`, `Piotr-Miller/anthropic-skills` — they contain text from Anthropic's source-available skill; `tfriedel/claude-office-skills` and the claudekit `slides` skill have no license; blog posts are author-owned (re-verify facts, rewrite prose).
- Before claiming novelty, re-check GitHub for `pptx` / `pptxgenjs` / `slide deck skill` prior art (see README "Prior art & alternatives"; note `siril9/presentation-skill` claims the `pptx-skill` alias).
