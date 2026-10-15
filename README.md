# pi-pptx

An original, MIT-licensed **PPTX skill** for coding agents that speak the [Agent Skills](https://agentskills.io/specification) convention — built and verified for [pi](https://github.com/badlogic/pi-mono) and OMP, and usable by Claude Code via a symlink.

It creates decks with **pptxgenjs**, edits existing decks/templates at the **OOXML** level, renders slides to images for **visual QA**, and ships **genre playbooks** (consulting / academic / product / internal / training) with sources.

> Independent implementation. Inspired by publicly documented workflows and by Anthropic's source-available `pptx` skill (reference use only — no text or code from it is reused). Not affiliated with Anthropic.

## Install

**pi** (package install — npm source once published, git, or a local path):

```bash
pi install /path/to/pi-pptx                 # local checkout
pi install github:kasi-x/pi-pptx            # git source (after pushing)
pi install npm:pi-pptx                      # npm source (after publishing)
```

**OMP / other Agent Skills hosts** (global skills directory):

```bash
ln -s /path/to/pi-pptx/skills/pptx ~/.agents/skills/pptx
```

**Claude Code** — the same `~/.agents/skills/pptx` entry works through the usual symlink pattern:

```bash
ln -s ../../.agents/skills/pptx ~/.claude/skills/pptx
```

## Requirements

| Need | For | Install |
|---|---|---|
| node ≥ 18 + `pptxgenjs` ^4 | deck creation | `npm install pptxgenjs` (per project) |
| python3 + `python-pptx` | `scripts/inspect_deck.py` | `pip install python-pptx` |
| `markitdown[pptx]` | text dumps | `uv tool install "markitdown[pptx]"` |
| LibreOffice + poppler | `scripts/render.sh` (visual QA) | distro: `libreoffice poppler-utils` |

Quick check: `node -e "require.resolve('pptxgenjs')" && python3 -c "import pptx" && command -v soffice pdftoppm`.

## Usage

- **pi / OMP**: `/skill:pptx 顧客向け提案資料を作って` — or just mention decks/slides/プレゼン and the skill is pulled in by its description.
- **Any host**: read `skills/pptx/SKILL.md`; it routes to `references/` (pptxgenjs guide, OOXML workflow, deck types) and `scripts/` (unpack / pack / dup_slide / render / inspect).

## Layout

```text
skills/pptx/
├── SKILL.md                     routing + guardrails
├── references/
│   ├── pptxgenjs.md             API guide + verified gotchas
│   ├── ooxml-workflow.md        edit/template workflow
│   └── deck-types.md            consulting / academic / product / internal playbooks
├── scripts/
│   ├── unpack.sh / pack.sh      zip round-trip for OOXML edits
│   ├── dup_slide.py             duplicate a slide with full bookkeeping
│   ├── render.sh                deck → PDF → slide JPEGs
│   └── inspect_deck.py          open check + structure/text summary
└── examples/deck.js             starter generator (4 slides, chart, table, notes)
```

## Verify

```bash
npm test                                    # generates example-deck.pptx + inspects it
bash skills/pptx/scripts/render.sh example-deck.pptx   # then LOOK at the images
```

## License

MIT — see [LICENSE](LICENSE). Third-party inspirations and further reading are listed in the source sections of `skills/pptx/references/deck-types.md`; the pptxgenjs reference links its upstream docs.
