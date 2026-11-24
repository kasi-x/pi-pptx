#!/usr/bin/env python3
"""Lightweight .pptx -> self-contained HTML preview (no LibreOffice needed).

Reads every slide with python-pptx and writes one HTML file with a 16:9 box
per slide: text, tables, embedded images (base64 data URIs), chart data and
speaker notes. This is for checking content and rough layout when rendering
is unavailable -- it is not a substitute for `render.sh` + visual QA.

usage:
  preview_deck.py deck.pptx [--out out.html]

Default output is <deck-stem>.preview.html next to the deck.

Exit codes: 0 ok, 1 cannot open/write the deck, 2 python-pptx missing.
"""

from __future__ import annotations

import argparse
import base64
import html
import sys
from pathlib import Path
from typing import Any

CANVAS_W = 960  # px, 16:9
CANVAS_H = 540
Transform = tuple[float, float, float, float]  # (sx, tx, sy, ty); abs = raw * s + t
IDENTITY: Transform = (1.0, 0.0, 1.0, 0.0)

CSS = """
body { margin: 0; padding: 24px; background: #f2f3f5; color: #1b1b1b;
       font-family: -apple-system, "Helvetica Neue", Arial, "Noto Sans JP", sans-serif; }
h2 { font-size: 15px; margin: 20px auto 6px; max-width: 960px; color: #555; }
.slide { position: relative; box-sizing: border-box; width: 960px; height: 540px;
         margin: 0 auto; background: #fff; border: 1px solid #c9cdd2; overflow: hidden; }
.box { position: absolute; box-sizing: border-box; padding: 2px 4px; font-size: 16px;
       line-height: 1.25; overflow: visible; }
.title { font-weight: 700; }
.p { margin: 0 0 0.35em; }
.box table { border-collapse: collapse; width: 100%; font-size: inherit; }
.box td, .box th { border: 1px solid #9aa0a6; padding: 2px 5px; text-align: left; }
.box th { background: #eef1f4; }
.box img { display: block; width: 100%; height: 100%; object-fit: fill; }
.chart { border: 1px dashed #888; background: #fafafa; }
.notes { max-width: 960px; margin: 6px auto 18px; padding: 8px 12px; background: #fff;
         border-left: 3px solid #8a8f98; font-size: 13px; white-space: pre-wrap; color: #444; }
.notes b { display: block; margin-bottom: 4px; color: #222; }
@media print { body { background: #fff; padding: 0; } .slide { border: none; page-break-after: always; } }
"""


def esc(text: Any) -> str:
    return html.escape(str(text), quote=True)


def px(value: int, scale: float) -> str:
    return f"{value * scale:.1f}px"


def group_transform(shape: Any) -> Transform | None:
    """Map a group's child coordinate space onto the slide, if available."""
    try:
        xfrm = shape._element.grpSpPr.xfrm
        if xfrm is None or xfrm.chExt.cx == 0 or xfrm.chExt.cy == 0:
            return None
        sx = shape.width / xfrm.chExt.cx
        sy = shape.height / xfrm.chExt.cy
        return (sx, shape.left - xfrm.chOff.x * sx, sy, shape.top - xfrm.chOff.y * sy)
    except (AttributeError, TypeError, ZeroDivisionError):
        return None


def compose(parent: Transform, child: Transform) -> Transform:
    psx, ptx, psy, pty = parent
    csx, ctx, csy, cty = child
    return (csx * psx, ctx * psx + ptx, csy * psy, cty * psy + pty)


def paragraph_html(paragraph: Any, is_title: bool, scale: float) -> str:
    text = "".join(run.text for run in paragraph.runs)
    if not text:
        return ""
    p_pr = paragraph._p.pPr
    bullet = ""
    if p_pr is not None and not is_title:
        if p_pr.find("{http://schemas.openxmlformats.org/drawingml/2006/main}buChar") is not None:
            bullet = "• "
        elif p_pr.find("{http://schemas.openxmlformats.org/drawingml/2006/main}buAutoNum") is not None:
            bullet = "1. "
        elif paragraph.level > 0 and p_pr.find(
            "{http://schemas.openxmlformats.org/drawingml/2006/main}buNone"
        ) is None:
            bullet = "• "
    size = next((r.font.size.pt for r in paragraph.runs if r.font.size is not None), None)
    bold = any(r.font.bold for r in paragraph.runs)
    italic = any(r.font.italic for r in paragraph.runs)
    style = []
    if size:
        style.append(f"font-size:{max(8.0, size * scale):.1f}px")
    if bold or is_title:
        style.append("font-weight:700")
    if italic:
        style.append("font-style:italic")
    if paragraph.level:
        style.append(f"margin-left:{paragraph.level * 18}px")
    if paragraph.alignment is not None:
        style.append(f"text-align:{str(paragraph.alignment).split()[0].lower()}")
    attr = f' style="{esc(";".join(style))}"' if style else ""
    return f'<div class="p"{attr}>{esc(bullet + text)}</div>'


def text_html(shape: Any, scale: float, is_title: bool) -> str:
    body = "".join(paragraph_html(p, is_title, scale) for p in shape.text_frame.paragraphs)
    return f'<div class="box{" title" if is_title else ""}">{body}</div>'


def table_html(shape: Any, scale: float) -> str:
    rows = []
    for ridx, row in enumerate(shape.table.rows):
        cells = []
        for cell in row.cells:
            tag = "th" if ridx == 0 else "td"
            cells.append(f"<{tag}>{esc(cell.text)}</{tag}>")
        rows.append("<tr>" + "".join(cells) + "</tr>")
    return (
        f'<div class="box"><table style="font-size:{max(9.0, 14 * scale):.1f}px">'
        + "".join(rows) + "</table></div>"
    )


def image_html(shape: Any) -> str | None:
    try:
        blob = shape.image.blob
        content_type = shape.image.content_type
    except (AttributeError, ValueError, KeyError):
        return None
    uri = "data:" + esc(content_type) + ";base64," + base64.b64encode(blob).decode("ascii")
    return f'<div class="box"><img src="{uri}" alt="image"></div>'


def chart_html(shape: Any, scale: float) -> str:
    """Chart data as a small table: content check, not a chart rendering."""
    lines: list[str] = []
    try:
        chart = shape.chart
        if chart.has_title:
            lines.append(f"<b>{esc(chart.chart_title.text_frame.text)}</b>")
        try:
            categories = [str(c) for c in chart.plots[0].categories]
        except (AttributeError, IndexError, ValueError):
            categories = []
        series = list(chart.series)
        lines.append("<table><tr><th></th>" + "".join(f"<th>{esc(s.name)}</th>" for s in series) + "</tr>")
        for idx, category in enumerate(categories):
            values = "".join(
                f"<td>{esc(list(s.values)[idx]) if idx < len(list(s.values)) else ''}</td>"
                for s in series
            )
            lines.append(f"<tr><th>{esc(category)}</th>{values}</tr>")
        lines.append("</table>")
    except Exception as exc:  # noqa: BLE001 - chart internals vary; keep the preview usable
        lines.append(f"<div>[chart data unavailable: {esc(exc)}]</div>")
    return (
        f'<div class="box chart"><div style="font-size:{max(9.0, 12 * scale):.1f}px">'
        + "".join(lines) + "</div></div>"
    )


def render_shape(shape: Any, tf: Transform, scale_x: float, scale_y: float,
                 scale_font: float, is_title: bool) -> str:
    sx, tx, sy, ty = tf
    try:
        left, top = shape.left * sx + tx, shape.top * sy + ty
        width, height = shape.width * sx, shape.height * sy
    except TypeError:
        return ""  # shape without resolved geometry
    position = (f"left:{px(int(left), scale_x)};top:{px(int(top), scale_y)};"
                f"width:{px(int(width), scale_x)};height:{px(int(height), scale_y)}")

    if getattr(shape, "shape_type", None) == 6:  # GROUP
        child_tf = group_transform(shape)
        if child_tf is None:
            return ""
        return "".join(
            render_shape(child, compose(tf, child_tf), scale_x, scale_y, scale_font, False)
            for child in shape.shapes
        )
    inner: str | None = None
    if getattr(shape, "has_table", False):
        inner = table_html(shape, scale_font)
    elif getattr(shape, "has_chart", False):
        inner = chart_html(shape, scale_font)
    elif getattr(shape, "shape_type", None) == 13:  # PICTURE
        inner = image_html(shape)
    if inner is None and getattr(shape, "has_text_frame", False) and shape.text_frame.text.strip():
        inner = text_html(shape, scale_font, is_title)
    if inner is None:
        return ""
    return inner.replace("<div class=", f'<div style="{esc(position)}" class=', 1)


def render_slide(slide: Any, number: int, prs: Any, scale_x: float, scale_y: float,
                 scale_font: float) -> str:
    try:
        title_shape = slide.shapes.title
    except (AttributeError, ValueError, KeyError):
        title_shape = None
    shapes = "".join(
        render_shape(shape, IDENTITY, scale_x, scale_y, scale_font, shape is title_shape)
        for shape in slide.shapes
    )
    notes = ""
    if slide.has_notes_slide:
        note_text = slide.notes_slide.notes_text_frame.text.strip()
        if note_text:
            notes = (f'<div class="notes"><b>Slide {number} notes</b>'
                     f"{esc(note_text)}</div>")
    return (
        f"<h2>Slide {number}</h2>"
        f'<div class="slide" aria-label="slide {number}">{shapes}</div>{notes}'
    )


def build_html(prs: Any) -> str:
    scale_x = CANVAS_W / prs.slide_width
    scale_y = CANVAS_H / prs.slide_height
    scale_font = CANVAS_W / (prs.slide_width / 12700)  # px per pt
    slides = "".join(
        render_slide(slide, number, prs, scale_x, scale_y, scale_font)
        for number, slide in enumerate(prs.slides, 1)
    )
    return (
        "<!doctype html>\n<html><head><meta charset=\"utf-8\">"
        '<meta name="viewport" content="width=device-width, initial-scale=1">'
        f"<title>Deck preview ({len(prs.slides)} slides)</title>"
        f"<style>{CSS}</style></head><body>{slides}</body></html>\n"
    )


def main() -> int:
    ap = argparse.ArgumentParser(
        prog="preview_deck.py",
        description="Write a self-contained HTML content preview of a .pptx deck.",
    )
    ap.add_argument("deck", help="path to a .pptx/.potx file")
    ap.add_argument("--out", metavar="FILE", help="output HTML (default: <deck-stem>.preview.html)")
    args = ap.parse_args()

    try:
        from pptx import Presentation  # type: ignore[import-not-found]
    except ImportError:
        print("preview: python-pptx is required (pip install python-pptx)", file=sys.stderr)
        return 2

    deck = Path(args.deck)
    out = Path(args.out) if args.out else deck.with_suffix(".preview.html")
    try:
        prs = Presentation(str(deck))
    except Exception as exc:  # noqa: BLE001 - report any parse failure verbatim
        print(f"preview: cannot open {deck}: {exc}", file=sys.stderr)
        return 1
    try:
        out.write_text(build_html(prs), encoding="utf-8")
    except OSError as exc:
        print(f"preview: cannot write {out}: {exc}", file=sys.stderr)
        return 1
    print(f"{deck}: wrote {len(prs.slides)} slides to {out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
