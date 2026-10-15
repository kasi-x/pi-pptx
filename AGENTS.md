# pi-pptx — AGENTS.md

## Overview

Original, MIT-licensed PPTX skill (Agent Skills format) for pi and OMP, packaged as a pi package: create decks with pptxgenjs, edit PPTX/OOXML, render + visually QA slides, and follow genre playbooks (consulting / academic / product / internal). The live skill copy for OMP/Claude Code is symlinked from `~/.agents/skills/pptx`.

## Commands

```bash
npm test                                            # generate example deck + inspect it
bash skills/pptx/scripts/render.sh example-deck.pptx   # visual QA renders (then LOOK at them)
python3 skills/pptx/scripts/dup_slide.py --help     # slide duplication
pi install ~/dev/pi-pptx                            # (re)register with pi
```

Publish (author's call, not automated): `gh repo create kasi-x/pi-pptx --public --source . --push`, then optionally `npm publish`.

## Layout

- `skills/pptx/SKILL.md` — routing + guardrails (keep ≤ ~150 lines)
- `skills/pptx/references/` — pptxgenjs guide, OOXML workflow, deck types (with sources)
- `skills/pptx/scripts/` — unpack/pack, dup_slide, render, inspect (python3 stdlib or POSIX only)
- `skills/pptx/examples/deck.js` — starter generator (must keep running via `npm test`)
- `research/` — local-only working notes (gitignored, not shipped)

## Constraints

- **Original content only.** Never copy text or scripts from Anthropic's source-available pptx skill (its license forbids copying, derivative works, and redistribution). Public-documentation facts are fine — cite the source in the reference.
- Keep runtime deps thin: deck creation uses pptxgenjs (MIT); QA scripts use python3 stdlib / python-pptx / LibreOffice / poppler.
- After changing scripts or the example: run `npm test` and the render loop, and look at the images.
- No emoji in shipped files.
