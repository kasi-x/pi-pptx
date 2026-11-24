#!/usr/bin/env python3
"""Render diagram and formula assets declared in an outline JSON.

Usage:
    python3 render_assets.py outline.json assets/ [--dpi 200]

Every slide may carry an "assets" array:
    {"type": "mermaid", "id": "flow1", "code": "graph TD; A-->B"}
    {"type": "dot",     "id": "arch1", "code": "digraph G { A -> B; }"}
    {"type": "latex",   "id": "eq1", "code": "E=mc^2", "mode": "display"}

Outputs (when the matching tool is installed) are <assets>/<id>.svg and
<assets>/<id>.png, plus <assets>/manifest.json with path/dpi/size. Missing
optional tools only skip that asset type with a warning; other assets still
render. Python standard library only.
"""

from __future__ import annotations

import argparse
import json
import re
import shutil
import struct
import subprocess
import sys
import tempfile
from pathlib import Path

ID_RE = re.compile(r"^[A-Za-z0-9][A-Za-z0-9_-]*$")
LATEX_WRAPPER = r"""\documentclass[12pt]{article}
\usepackage{amsmath,amssymb}
\usepackage[paperwidth=8in,paperheight=2in,margin=0.2in]{geometry}
\pagestyle{empty}
\begin{document}
\noindent
@@MATH@@
\end{document}
"""


def warn(message: str) -> None:
    print(f"warning: {message}", file=sys.stderr)


def fail(message: str) -> None:
    print(f"error: {message}", file=sys.stderr)
    raise SystemExit(1)


def run(command: list[str], timeout: int = 180, input_text: str | None = None) -> str:
    try:
        result = subprocess.run(
            command, input=input_text, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
            text=True, timeout=timeout,
        )
    except (OSError, subprocess.TimeoutExpired) as exc:
        raise RuntimeError(f"{command[0]}: {exc}") from exc
    if result.returncode != 0:
        detail = (result.stderr or result.stdout).strip().replace("\n", " ")[:400]
        raise RuntimeError(f"{command[0]} exited {result.returncode}: {detail}")
    return result.stdout


def png_size(path: Path) -> tuple[int, int] | None:
    with path.open("rb") as handle:
        header = handle.read(24)
    if len(header) < 24 or header[:8] != b"\x89PNG\r\n\x1a\n":
        return None
    return struct.unpack(">II", header[16:24])


def svg_size(path: Path) -> tuple[int, int] | None:
    text = path.read_text(encoding="utf-8", errors="replace")[:8192]
    viewbox = re.search(r'viewBox\s*=\s*"\s*(-?[\d.]+)[ ,]+(-?[\d.]+)[ ,]+(-?[\d.]+)[ ,]+(-?[\d.]+)', text)
    if viewbox:
        w, h = float(viewbox.group(3)), float(viewbox.group(4))
        if w > 0 and h > 0:
            return (round(w), round(h))
    width = re.search(r'\bwidth\s*=\s*"([\d.]+)(?:px)?"', text)
    height = re.search(r'\bheight\s*=\s*"([\d.]+)(?:px)?"', text)
    if width and height and float(width.group(1)) > 0 and float(height.group(1)) > 0:
        return (round(float(width.group(1))), round(float(height.group(1))))
    return None


def image_size(path: Path) -> tuple[int, int] | None:
    return svg_size(path) if path.suffix.lower() == ".svg" else png_size(path)


def code_of(asset: dict, outline_path: Path) -> str:
    if "code" in asset:
        code = asset["code"]
        if not isinstance(code, str) or not code.strip():
            fail(f"asset \"{asset.get('id')}\": code must be a non-empty string")
        return code.strip()
    file_key = asset.get("codeFile", asset.get("code_file"))
    if isinstance(file_key, str) and file_key.strip():
        source = Path(file_key)
        if not source.is_absolute():
            source = outline_path.parent / source
        try:
            return source.read_text(encoding="utf-8").strip()
        except OSError as exc:
            fail(f"asset \"{asset.get('id')}\": cannot read codeFile {file_key}: {exc}")
    fail(f"asset \"{asset.get('id')}\": needs code or codeFile")


def collect_assets(outline: dict, outline_path: Path) -> list[tuple[int, dict, str]]:
    slides = outline.get("slides")
    if not isinstance(slides, list) or not slides:
        fail("outline has no slides[]")
    found: list[tuple[int, dict, str]] = []
    for number, slide in enumerate(slides, 1):
        raw = slide.get("assets", []) if isinstance(slide, dict) else []
        if isinstance(raw, dict):
            raw = [raw]
        if not isinstance(raw, list):
            fail(f"slide {number}: assets must be an array or object")
        for asset in raw:
            if not isinstance(asset, dict):
                fail(f"slide {number}: each asset must be an object")
            kind = asset.get("type")
            if kind not in ("mermaid", "dot", "latex"):
                fail(f"slide {number}: unknown asset type {kind!r} (mermaid|dot|latex)")
            asset_id = asset.get("id")
            if not isinstance(asset_id, str) or not ID_RE.match(asset_id):
                fail(f"slide {number}: asset id must be a filename-safe string (letters, digits, -, _)")
            mode = asset.get("mode", "display")
            if kind == "latex" and mode not in ("display", "inline"):
                fail(f"asset \"{asset_id}\": mode must be display or inline")
            found.append((number, asset, code_of(asset, outline_path)))
    return found


def render_mermaid(asset_id: str, code: str, out: Path, tmp: Path, dpi: int) -> list[Path]:
    source = tmp / f"{asset_id}.mmd"
    source.write_text(code + "\n", encoding="utf-8")
    # mmdc's PNG is 2x its --size; dpi is expressed at a nominal 8in slide width.
    size = str(dpi * 4)
    products = [out / f"{asset_id}.svg", out / f"{asset_id}.png"]
    run(["mmdc", "-i", str(source), "-o", str(products[0]), "-b", "white", "--size", size, "-q"])
    run(["mmdc", "-i", str(source), "-o", str(products[1]), "-b", "white", "--size", size, "-s", "2", "-q"])
    return products


def render_dot(asset_id: str, code: str, out: Path, tmp: Path, dpi: int) -> list[Path]:
    source = tmp / f"{asset_id}.dot"
    source.write_text(code + "\n", encoding="utf-8")
    products = [out / f"{asset_id}.svg", out / f"{asset_id}.png"]
    run(["dot", "-Tsvg", "-o", str(products[0]), str(source)])
    run(["dot", f"-Gdpi={dpi * 8}", "-Tpng", "-o", str(products[1]), str(source)])
    return products


def render_latex(asset_id: str, code: str, out: Path, tmp: Path, dpi: int, mode: str) -> list[Path]:
    latex_engine = shutil.which("pdflatex") or shutil.which("latex")
    if not latex_engine:
        warn(f"skip latex \"{asset_id}\": no LaTeX engine (pdflatex/latex) found")
        return []
    to_png = shutil.which("pdftocairo") or shutil.which("pdftoppm")
    if not to_png:
        warn(f"skip latex \"{asset_id}\": no PDF rasterizer (poppler) found")
        return []
    fragment = f"${code}$" if mode == "inline" else f"$${code}$$"
    fragment_tex = run(["pandoc", "-f", "markdown+raw_tex", "-t", "latex", "--wrap=none"], timeout=60, input_text=fragment + "\n").strip()
    wrapper = tmp / f"{asset_id}.tex"
    wrapper.write_text(LATEX_WRAPPER.replace("@@MATH@@", fragment_tex), encoding="utf-8")
    run([latex_engine, "-interaction=nonstopmode", "-halt-on-error", "-output-directory", str(tmp), str(wrapper)], timeout=120)
    pdf = tmp / f"{asset_id}.pdf"
    if not pdf.exists():
        raise RuntimeError(f"{latex_engine}: no PDF produced")
    products: list[Path] = []
    png = out / f"{asset_id}.png"
    if shutil.which("pdftocairo"):
        run(["pdftocairo", "-png", "-r", str(dpi), "-singlefile", str(pdf), str(png.with_suffix(""))])
    else:
        run(["pdftoppm", "-png", "-r", str(dpi), "-singlefile", str(pdf), str(png.with_suffix(""))])
    products.append(png)
    svg = out / f"{asset_id}.svg"
    if shutil.which("pdftocairo"):
        run(["pdftocairo", "-svg", str(pdf), str(svg)])
        products.append(svg)
    else:
        warn(f"latex \"{asset_id}\": pdftocairo missing, SVG output skipped (PNG is available)")
    return products


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Render mermaid/dot/latex assets from an outline JSON into an assets directory.",
    )
    parser.add_argument("outline", help="outline JSON with optional slide.assets[] entries")
    parser.add_argument("assets_dir", help="output directory for <id>.svg/.png and manifest.json")
    parser.add_argument("--dpi", type=int, default=200, help="raster resolution (default: 200)")
    args = parser.parse_args()
    if args.dpi < 72 or args.dpi > 600:
        fail("--dpi must be between 72 and 600")

    outline_path = Path(args.outline).resolve()
    try:
        outline = json.loads(outline_path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        fail(f"cannot read outline {args.outline}: {exc}")
    if not isinstance(outline, dict):
        fail("outline must be a JSON object")
    assets = collect_assets(outline, outline_path)
    out = Path(args.assets_dir)
    if not out.is_absolute():
        out = (Path.cwd() / out).resolve()
    out.mkdir(parents=True, exist_ok=True)
    manifest_path = out / "manifest.json"
    manifest: dict = {}
    if manifest_path.exists():
        try:
            loaded = json.loads(manifest_path.read_text(encoding="utf-8"))
            if isinstance(loaded, dict):
                manifest = loaded
            else:
                warn("existing manifest.json is not an object; rebuilding it")
        except (OSError, json.JSONDecodeError) as exc:
            warn(f"cannot read existing manifest.json: {exc}; rebuilding it")
    if not assets:
        print("no assets declared in outline")
    else:
        with tempfile.TemporaryDirectory(prefix="pi-pptx-assets-") as temp:
            tmp = Path(temp)
            for number, asset, code in assets:
                asset_id = asset["id"]
                kind = asset["type"]
                if shutil.which("mmdc" if kind == "mermaid" else "dot" if kind == "dot" else "pandoc") is None:
                    if kind == "latex":
                        warn(f"skip latex \"{asset_id}\": pandoc not found - OMML is unavailable; この環境では数式は画像化のみ (equations are image-only)")
                    else:
                        warn(f"skip {kind} \"{asset_id}\" (slide {number}): {('mmdc' if kind == 'mermaid' else 'dot')} not found")
                    continue
                try:
                    if kind == "mermaid":
                        products = render_mermaid(asset_id, code, out, tmp, args.dpi)
                    elif kind == "dot":
                        products = render_dot(asset_id, code, out, tmp, args.dpi)
                    else:
                        products = render_latex(asset_id, code, out, tmp, args.dpi, asset.get("mode", "display"))
                except RuntimeError as exc:
                    warn(f"skip {kind} \"{asset_id}\" (slide {number}): {exc}")
                    continue
                usable = [p for p in products if p.exists() and image_size(p)]
                if not usable:
                    warn(f"skip {kind} \"{asset_id}\" (slide {number}): renderer produced no usable image")
                    continue
                primary = next((p for p in usable if p.suffix.lower() == ".svg"), usable[0])
                png = next((p for p in usable if p.suffix.lower() == ".png"), None)
                entry = {
                    "type": kind,
                    "path": primary.name,
                    "dpi": args.dpi,
                    "size": {"w": image_size(primary)[0], "h": image_size(primary)[1]},
                }
                if png:
                    entry["png"] = png.name
                manifest[asset_id] = entry
                print(f"rendered {kind} \"{asset_id}\" -> {primary}")
    manifest_path.write_text(json.dumps(manifest, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")
    print(f"wrote: {manifest_path} ({len(manifest)} assets)")


if __name__ == "__main__":
    main()
