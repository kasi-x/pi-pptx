# Diagram and formula assets

Use this when a deck needs a diagram or equation that native pptxgenjs shapes/charts
cannot express. Everything is optional: `render_assets.py` uses only the Python
standard library, detects external renderers at runtime, and skips only the asset
types whose tool is missing. Other assets and the deck build continue normally.

There is **no OMML/equation slide type**. Equations are rendered as images and
placed on a `figure` slide. Native charts stay native (`addChart()`) — do not
turn a bar/line chart into a picture.

## Workflow

```bash
# 1. render every slide.assets[] entry into an assets directory
python3 scripts/render_assets.py outline.json assets/         # --dpi 200 by default

# 2. build the deck; figure slides pick up images via the outline path or manifest
node scripts/outline_to_pptx.js outline.json out.pptx
```

Paths are relative to this skill directory. From the repository root the same
renderer is available as `python3 scripts/render_assets.py ...`.

## Outline schema

Add an `assets` array to any slide. Each entry needs a `type`, a filename-safe
`id`, and inline `code` (or `codeFile`, a path relative to the outline):

```json
{
  "meta": { "title": "System overview" },
  "slides": [
    {
      "type": "bullets",
      "title": "Context",
      "bullets": ["Input flows through three services"],
      "assets": [
        { "type": "mermaid", "id": "flow1", "code": "graph TD; A-->B; A-->C; B-->D" },
        { "type": "dot", "id": "arch1", "code": "digraph G { A -> B [label=\"rpc\"]; }" },
        { "type": "latex", "id": "eq1", "mode": "display", "code": "E=mc^2" }
      ]
    },
    {
      "type": "figure",
      "title": "Request flow",
      "image": "assets/flow1.svg",
      "caption": "Requests fan out after ingress.",
      "width": 8.0,
      "notes": "Paused at service B while the cache is cold."
    },
    { "type": "figure", "title": "Mass-energy", "image": "eq1", "width": 5.0 }
  ]
}
```

| Type | Tool | Output |
|---|---|---|
| `mermaid` | `mmdc` (mermaid-cli) | `<id>.svg` + `<id>.png` |
| `dot` | `dot` (Graphviz) | `<id>.svg` + `<id>.png` |
| `latex` | `pandoc` + `pdflatex`/`latex` + poppler | `<id>.png` (+ `<id>.svg` with `pdftocairo`) |

`mode` applies only to `latex`: `display` (default) wraps code in `$$…$$`,
`inline` in `$…$`. `--dpi N` (72–600, default 200) is the target density at a
nominal 8-inch display width; the PNGs are rendered large enough that an 8-inch
figure stays above the 96 dpi lint threshold.

## Outputs and manifest

`assets/<id>.svg`, `assets/<id>.png` and `assets/manifest.json` are written next
to each other. The manifest is an object keyed by asset id, so it is safe to
merge with later runs:

```json
{
  "flow1": {
    "type": "mermaid",
    "path": "flow1.svg",
    "png": "flow1.png",
    "dpi": 200,
    "size": { "w": 360, "h": 231 }
  }
}
```

`path` is the primary (vector) file, `size` its canvas in px, `dpi` the requested
raster density. Entries for skipped assets are not added.

## `figure` slides

`image` is required. It may be a path relative to the outline (for example
`assets/flow1.svg`) or a manifest id (`flow1`). Missing images fail with a clear
error. When both SVG and PNG exist, the PNG is embedded — pptxgenjs cannot
rasterize SVG previews in Node, so raw SVG slides show a broken preview in some
viewers; a bare SVG path is still embedded as `image/svg+xml`.

`width` is the requested width in inches (default 8.0, clamped to the 0.7-inch
page margins and the available body height). The aspect ratio comes from the PNG
header or SVG `viewBox`; `caption` is centered below the image and also becomes
the image alt text. Every palette works, and the consulting kicker/source/page
chrome is added automatically when `"palette": "consulting"`.

## Optional dependencies

```bash
brew install graphviz                        # dot: architecture / call graphs
npm install -g @mermaid-js/mermaid-cli       # mmdc: flowcharts, sequences, gantt
brew install pandoc                          # LaTeX math -> LaTeX source
brew install mactex                          # pdflatex (full TeX distribution)
brew install poppler                         # pdftocairo / pdftoppm -> PNG/SVG
```

Without `mmdc`/`dot`, those assets are skipped with a warning. Without `pandoc`,
the renderer warns that OMML is unavailable and that **this environment can only
use images for equations** (数式は画像化のみ); generate the formula elsewhere
and paste it with a `figure` slide. The same skip happens when a LaTeX engine or
PDF rasterizer is missing. A failed render command also warns and lets the
remaining assets continue.
