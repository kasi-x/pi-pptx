#!/usr/bin/env python3
"""Geometry and typography linter for .pptx decks.

`render.sh` shows you what a deck looks like; this script says which shapes are
wrong before you look. It walks every shape on every slide (groups included,
via the group transform) and reports the layout damage that is cheap to detect
and expensive to miss: off-canvas text, zero-size extents, margin and gap
violations, overlapping text, disallowed fonts, likely text overflow,
low-resolution images, leftover placeholder copy and empty text frames. With
`--pdf` it also reads the rendered PDF (`pdftotext -bbox-layout`) and flags word
boxes that cross the content margins (ignoring words that sit inside a slide
number/date/footer placeholder), plus a page/slide count mismatch.

Requires python-pptx; --pdf additionally requires poppler's pdftotext.

usage:
  lint_deck.py deck.pptx [--json] [--strict] [--pdf deck.pdf]
                         [--margin 0.5] [--gap 0.3] [--fonts a,b,c]

Exit codes:
  0  no findings (warnings only, without --strict)
  1  errors found, or warnings with --strict
  2  python-pptx missing / bad usage
  3  the deck cannot be opened or parsed
  4  --pdf given but pdftotext is missing or the PDF cannot be read

Limits (deliberate, keep the output trustworthy):
  * boxes are measured axis-aligned; shape rotation is not modelled.
  * table geometry uses the row/column sums when they exceed the graphic
    frame's <a:ext> (pptxgenjs writes a frame that is shorter than its rows).
  * fonts come from text frames and table cells; chart-internal fonts are not
    scanned, and runs without an explicit typeface fall back to the theme's
    minor latin font, which is then checked as the effective font.
  * the overflow check is a heuristic (character counts, not shaping) and is
    skipped for empty frames (nothing can overflow) and for slide-number/date/
    footer placeholders, which render from a field and sit in the margin.
  * a shape whose position/size cannot be resolved (a placeholder inheriting
    from a layout/master that has no geometry either) is skipped by the
    geometry checks and counted as shapes_without_geometry in --json.
"""

from __future__ import annotations

import argparse
import html
import json
import math
import re
import subprocess
import sys
import unicodedata
import zipfile
from collections.abc import Iterator
from dataclasses import dataclass
from typing import Any

EMU_PER_IN = 914400
PT_PER_IN = 72.0
EMU_PER_PT = EMU_PER_IN / PT_PER_IN

MIN_EXTENT_EMU = 10000  # 0.011 in: anything smaller is a zero/undersized box
FULL_BLEED = 0.95  # spanning >=95% of an axis is deliberate bleed, not a mistake
TOUCH_TOL_EMU = 9144  # 0.01 in: text closer than this to an edge is "touching"
OVERLAP_MIN_EMU2 = int(0.02 * EMU_PER_IN * EMU_PER_IN)  # 0.02 in^2
LOW_DPI = 96.0
LINE_SPACING = 1.2  # PowerPoint single spacing
DEFAULT_INSET_LR = 91440  # 0.1 in
DEFAULT_INSET_TB = 45720  # 0.05 in
DEFAULT_FONT_SIZE_PT = 18.0
EMPTY_MIN_W_EMU = int(0.5 * EMU_PER_IN)
EMPTY_MIN_H_EMU = int(0.25 * EMU_PER_IN)
ALIGN_TOL_EMU = 91440  # 0.1 in tolerance for "the two boxes line up"
GAP_OVERLAP_RATIO = 0.5  # perpendicular overlap needed for a gap to matter

PLACEHOLDER_RE = re.compile(
    r"lorem|ipsum|TODO|xxx{3,}|\[insert|placeholder|click to add", re.IGNORECASE
)
DEFAULT_FONTS = (
    "Calibri",
    "Cambria",
    "Arial",
    "Times New Roman",
    "Courier New",
    "BIZ UDGothic",
    "Noto Sans JP",
)

SEVERITY_RANK = {"error": 0, "warn": 1}

Affine = tuple[float, float, float, float]  # (sx, tx, sy, ty): abs = raw * s + t
IDENTITY: Affine = (1.0, 0.0, 1.0, 0.0)


@dataclass
class Finding:
    severity: str
    code: str
    slide: int
    shape: str
    message: str
    fix: str
    values: dict[str, Any]

    def as_line(self) -> str:
        where = f"slide {self.slide}" if self.slide else "deck"
        return f'{self.severity.upper():<5} {where:<9} "{self.shape}" [{self.code}]: {self.message} -- fix: {self.fix}'

    def as_dict(self) -> dict[str, Any]:
        return {
            "severity": self.severity,
            "code": self.code,
            "slide": self.slide,
            "shape": self.shape,
            "message": self.message,
            "fix": self.fix,
            "values": self.values,
        }


@dataclass
class Frame:
    """One text frame (a shape's, or a single table cell's) with usable extents."""

    where: str
    paragraphs: list[tuple[str, float, str | None]]  # text, size pt, font name
    width_emu: int  # usable (insets removed)
    height_emu: int


@dataclass
class Item:
    slide: int
    shape_id: int
    name: str
    left: int
    top: int
    width: int
    height: int
    text: str
    frames: list[Frame]
    group: tuple[int, ...]  # ancestor group ids, outermost first
    ancestors: frozenset[int]
    has_text_frame: bool
    is_text: bool
    is_table: bool
    is_picture: bool
    is_chart: bool
    geom_ok: bool  # False when the shape inherits an unresolvable position/size
    is_chrome: bool  # slide-number/date/footer placeholder: lives in the margin by design
    shape: Any

    @property
    def right(self) -> int:
        return self.left + self.width

    @property
    def bottom(self) -> int:
        return self.top + self.height


@dataclass
class Ctx:
    slide_w: int
    slide_h: int
    margin: int
    gap: int
    fonts: set[str]  # casefolded allowlist
    theme_font: str | None = None


def inches(emu: float) -> float:
    return emu / EMU_PER_IN


def span_axis(start: int, end: int) -> str:
    return f"{inches(start):.2f}..{inches(end):.2f} in ({start}..{end} EMU)"


def char_units(text: str) -> float:
    """Text width in em units: full width for CJK, half width for latin."""
    return sum(1.0 if unicodedata.east_asian_width(ch) in ("W", "F") else 0.5 for ch in text)


# --------------------------------------------------------------------------- #
# model
# --------------------------------------------------------------------------- #


def theme_minor_font(path: str) -> str | None:
    """Latin typeface of the theme's minor font, read straight from the zip."""
    try:
        with zipfile.ZipFile(path) as zf:
            names = [n for n in zf.namelist() if re.fullmatch(r"ppt/theme/theme\d+\.xml", n)]
            if not names:
                return None
            xml = zf.read(min(names)).decode("utf-8", "replace")
    except (OSError, zipfile.BadZipFile, KeyError):
        return None
    match = re.search(r"<a:minorFont>.*?<a:latin typeface=\"([^\"]+)\"", xml, re.DOTALL)
    return match.group(1) if match else None


def _int(value: Any) -> int:
    return int(value) if value is not None else 0


def _is_chrome_placeholder(shape: Any) -> bool:
    """True for slide-number/date/footer placeholders.

    They render from the placeholder (often as a field, which is why the frame
    looks empty) and conventionally sit inside the page margin, so margin,
    overflow and empty-box warnings would always be noise.
    """
    from pptx.enum.shapes import PP_PLACEHOLDER

    try:
        if not shape.is_placeholder:
            return False
        return shape.placeholder_format.type in (
            PP_PLACEHOLDER.SLIDE_NUMBER,
            PP_PLACEHOLDER.DATE,
            PP_PLACEHOLDER.FOOTER,
        )
    except (AttributeError, ValueError):
        return False


def _group_parent(shape: Any) -> Affine:
    """(sx, tx, sy, ty) mapping the group's child space onto its parent space."""
    from pptx.oxml.ns import qn

    xfrm = None
    grp = shape.element.find(qn("p:grpSpPr"))
    if grp is not None:
        xfrm = grp.find(qn("a:xfrm"))

    def attr(tag: str, key: str) -> int:
        if xfrm is None:
            return 0
        node = xfrm.find(qn(tag))
        return _int(node.get(key)) if node is not None else 0

    off_x, off_y = attr("a:off", "x"), attr("a:off", "y")
    ext_x, ext_y = attr("a:ext", "cx"), attr("a:ext", "cy")
    ch_x, ch_y = attr("a:chOff", "x"), attr("a:chOff", "y")
    ch_ext_x, ch_ext_y = attr("a:chExt", "cx"), attr("a:chExt", "cy")
    kx = ext_x / ch_ext_x if ch_ext_x else 1.0
    ky = ext_y / ch_ext_y if ch_ext_y else 1.0
    return (kx, off_x - ch_x * kx, ky, off_y - ch_y * ky)


def _compose(parent: Affine, child: Affine) -> Affine:
    sx, tx, sy, ty = parent
    kx, ox, ky, oy = child
    return (kx * sx, ox * sx + tx, ky * sy, oy * sy + ty)


def _inset(value: Any, fallback: int) -> int:
    return int(value) if value is not None else fallback


def _frame_info(
    text_frame: Any, width: int, height: int, where: str, insets: tuple[int, int, int, int] | None = None
) -> Frame:
    """insets (l, r, t, b) in EMU; None means "read them from the text frame"."""
    if insets is None:
        insets = (
            _inset(text_frame.margin_left, DEFAULT_INSET_LR),
            _inset(text_frame.margin_right, DEFAULT_INSET_LR),
            _inset(text_frame.margin_top, DEFAULT_INSET_TB),
            _inset(text_frame.margin_bottom, DEFAULT_INSET_TB),
        )
    l_ins, r_ins, t_ins, b_ins = insets
    paragraphs: list[tuple[str, float, str | None]] = []
    for para in text_frame.paragraphs:
        # paragraph.text includes runs *and* fields/breaks, so a slide-number or
        # date field is not mistaken for an empty frame
        text = para.text
        size: float | None = None
        font: str | None = None
        for run in para.runs:
            if run.font.size is not None:
                size = max(size or 0.0, run.font.size.pt)
            if run.font.name and not font:
                font = run.font.name
        if size is None:
            size = para.font.size.pt if para.font.size is not None else DEFAULT_FONT_SIZE_PT
        if font is None:
            font = para.font.name
        paragraphs.append((text, float(size), font))
    return Frame(where, paragraphs, max(0, width - l_ins - r_ins), max(0, height - t_ins - b_ins))


def _make_item(
    shape: Any, slide_no: int, tf: Affine, group: tuple[int, ...], ancestors: frozenset[int]
) -> Item:
    sx, tx, sy, ty = tf
    # python-pptx resolves inherited placeholder geometry from the layout and
    # master; when even that fails the values are None and no geometry check
    # may run (0 EMU would look like a broken shape).
    geom_ok = all(v is not None for v in (shape.left, shape.top, shape.width, shape.height))
    width = round(_int(shape.width) * sx)
    height = round(_int(shape.height) * sy)
    left = round(_int(shape.left) * sx + tx)
    top = round(_int(shape.top) * sy + ty)

    is_table = bool(getattr(shape, "has_table", False))
    if is_table:  # pptxgenjs: the frame <a:ext> is often shorter than its rows
        table = shape.table
        cols = sum(_int(col.width) for col in table.columns)
        rows = sum(_int(row.height) for row in table.rows)
        width = max(width, round(cols * sx))
        height = max(height, round(rows * sy))

    frames: list[Frame] = []
    if getattr(shape, "has_text_frame", False):
        frames.append(_frame_info(shape.text_frame, width, height, "text frame"))
    if is_table:
        table = shape.table
        for ri, row in enumerate(table.rows, 1):
            for ci, cell in enumerate(row.cells, 1):
                try:
                    col_w = _int(table.columns[ci - 1].width)
                except IndexError:
                    col_w = 0
                cell_insets = (
                    _inset(cell.margin_left, DEFAULT_INSET_LR),
                    _inset(cell.margin_right, DEFAULT_INSET_LR),
                    _inset(cell.margin_top, DEFAULT_INSET_TB),
                    _inset(cell.margin_bottom, DEFAULT_INSET_TB),
                )
                frames.append(
                    _frame_info(
                        cell.text_frame,
                        col_w,
                        _int(row.height),
                        f"cell r{ri}c{ci}",
                        insets=cell_insets,
                    )
                )

    text = "\n".join("\n".join(p[0] for p in frame.paragraphs) for frame in frames).strip()
    is_picture = shape.shape_type is not None and int(shape.shape_type) == 13
    is_chart = bool(getattr(shape, "has_chart", False))
    return Item(
        slide=slide_no,
        shape_id=_int(shape.shape_id),
        name=shape.name,
        left=left,
        top=top,
        width=width,
        height=height,
        text=text,
        frames=frames,
        group=group,
        ancestors=ancestors,
        has_text_frame=bool(getattr(shape, "has_text_frame", False)),
        is_text=bool(text.strip()),
        is_table=is_table,
        is_picture=is_picture,
        is_chart=is_chart,
        geom_ok=geom_ok,
        is_chrome=_is_chrome_placeholder(shape),
        shape=shape,
    )


def collect(
    shapes: Any,
    slide_no: int,
    tf: Affine,
    group: tuple[int, ...] = (),
    ancestors: frozenset[int] = frozenset(),
) -> Iterator[Item]:
    """Flatten a slide (or a group) into items with absolute EMU geometry."""
    from pptx.enum.shapes import MSO_SHAPE_TYPE

    for shape in shapes:
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            child_tf = _compose(tf, _group_parent(shape))
            yield from collect(
                shape.shapes, slide_no, child_tf, group + (shape.shape_id,), ancestors | {shape.shape_id}
            )
        else:
            yield _make_item(shape, slide_no, tf, group, ancestors)


# --------------------------------------------------------------------------- #
# checks
# --------------------------------------------------------------------------- #


def _add(out: list[Finding], severity: str, code: str, item: Item, message: str, fix: str, values: dict[str, Any]) -> None:
    out.append(Finding(severity, code, item.slide, item.name, message, fix, values))


def check_extents(items: list[Item], out: list[Finding]) -> None:
    for it in items:
        if not it.geom_ok:
            continue
        if it.width >= MIN_EXTENT_EMU and it.height >= MIN_EXTENT_EMU:
            continue
        _add(
            out,
            "error",
            "tiny-extent",
            it,
            f"extent {it.width} x {it.height} EMU ({inches(it.width):.3f} x {inches(it.height):.3f} in) "
            f"is below the {MIN_EXTENT_EMU} EMU minimum",
            "delete the shape or give it a real width/height (a few inches, not fractions of one)",
            {"width_emu": it.width, "height_emu": it.height},
        )


def check_bounds(items: list[Item], ctx: Ctx, out: list[Finding]) -> None:
    """Off-canvas (error) and margin (warn) checks, one finding per shape."""
    for it in items:
        if not it.geom_ok:
            continue
        full_w = it.width >= FULL_BLEED * ctx.slide_w
        full_h = it.height >= FULL_BLEED * ctx.slide_h

        cross_x = it.left < -TOUCH_TOL_EMU or it.right > ctx.slide_w + TOUCH_TOL_EMU
        cross_y = it.top < -TOUCH_TOL_EMU or it.bottom > ctx.slide_h + TOUCH_TOL_EMU
        if it.is_text:  # text also fails when it merely touches an edge
            cross_x = cross_x or it.left <= TOUCH_TOL_EMU or it.right >= ctx.slide_w - TOUCH_TOL_EMU
            cross_y = cross_y or it.top <= TOUCH_TOL_EMU or it.bottom >= ctx.slide_h - TOUCH_TOL_EMU

        bleed_x = cross_x and full_w
        bleed_y = cross_y and full_h
        if (cross_x and not bleed_x) or (cross_y and not bleed_y):
            sides = []
            if cross_x and not bleed_x:
                sides.append(f"x {span_axis(it.left, it.right)}")
            if cross_y and not bleed_y:
                sides.append(f"y {span_axis(it.top, it.bottom)}")
            kind = "text shape touching/crossing the slide edge" if it.is_text else "shape outside the canvas"
            _add(
                out,
                "error",
                "off-canvas",
                it,
                f"{kind}: {'; '.join(sides)}; canvas is {span_axis(0, ctx.slide_w)} x {span_axis(0, ctx.slide_h)}",
                "move it inside the slide, or make it a genuine full-bleed shape (>=95% of that axis)",
                {
                    "left_emu": it.left,
                    "top_emu": it.top,
                    "right_emu": it.right,
                    "bottom_emu": it.bottom,
                },
            )
            continue  # already reported; don't stack a margin warning on top

        if full_w or full_h:  # a bleed shape is deliberate: it sits on the edge by design
            continue
        if it.is_chrome:  # slide number / date / footer live in the margin by design
            continue
        violations = []
        if it.left < ctx.margin:
            violations.append(f"left {inches(it.left):.2f} in < {inches(ctx.margin):.2f} in")
        if it.right > ctx.slide_w - ctx.margin:
            violations.append(f"right {inches(ctx.slide_w - it.right):.2f} in < {inches(ctx.margin):.2f} in")
        if it.top < ctx.margin:
            violations.append(f"top {inches(it.top):.2f} in < {inches(ctx.margin):.2f} in")
        if it.bottom > ctx.slide_h - ctx.margin:
            violations.append(f"bottom {inches(ctx.slide_h - it.bottom):.2f} in < {inches(ctx.margin):.2f} in")
        if violations:
            _add(
                out,
                "warn",
                "margin",
                it,
                f"closer to the edge than --margin {inches(ctx.margin):.2f} in: {'; '.join(violations)} "
                f"(shape spans x {span_axis(it.left, it.right)}, y {span_axis(it.top, it.bottom)})",
                "pull the shape in on the grid (0.5 in margins are the house default) or raise --margin for this deck",
                {"left_emu": it.left, "top_emu": it.top, "right_emu": it.right, "bottom_emu": it.bottom},
            )


def _related(a: Item, b: Item) -> bool:
    if a.group and b.group and a.group[-1] == b.group[-1]:
        return True  # siblings inside one group
    return a.shape_id in b.ancestors or b.shape_id in a.ancestors


def check_overlap(items: list[Item], out: list[Finding]) -> None:
    for i, a in enumerate(items):
        if not a.geom_ok:
            continue
        for b in items[i + 1 :]:
            if not b.geom_ok:
                continue
            if not a.is_text and not b.is_text:
                continue
            text_item = a if a.is_text else b
            if _related(a, b):
                continue
            ox = min(a.right, b.right) - max(a.left, b.left)
            oy = min(a.bottom, b.bottom) - max(a.top, b.top)
            if ox <= 0 or oy <= 0:
                continue
            area = ox * oy
            if area < OVERLAP_MIN_EMU2:
                continue
            other = b.name if text_item is a else a.name
            _add(
                out,
                "warn",
                "overlap",
                text_item,
                f'overlaps "{other}" by {inches(ox):.2f} x {inches(oy):.2f} in '
                f"({area / (EMU_PER_IN * EMU_PER_IN):.3f} in^2 = {area} EMU^2); text box "
                f"x {span_axis(text_item.left, text_item.right)}, y {span_axis(text_item.top, text_item.bottom)}",
                "move one of the two boxes, or drop the non-text shape behind the text on purpose (send to back)",
                {
                    "text_shape": text_item.name,
                    "other_shape": other,
                    "ox_emu": ox,
                    "oy_emu": oy,
                    "area_emu2": area,
                },
            )


def _axis_overlap(a_lo: int, a_hi: int, b_lo: int, b_hi: int) -> int:
    return min(a_hi, b_hi) - max(a_lo, b_lo)


def check_gaps(items: list[Item], ctx: Ctx, out: list[Finding]) -> None:
    texts = [it for it in items if it.is_text and it.geom_ok]
    for i, a in enumerate(texts):
        for b in texts[i + 1 :]:
            if _related(a, b):
                continue
            # vertical neighbours (stacked): separated in y, aligned in x
            gap_y = max(a.top - b.bottom, b.top - a.bottom)
            if 0 < gap_y < ctx.gap:
                ov = _axis_overlap(a.left, a.right, b.left, b.right)
                aligned = ov >= GAP_OVERLAP_RATIO * min(a.width, b.width) or abs(a.left - b.left) <= ALIGN_TOL_EMU
                if aligned:
                    _add(
                        out,
                        "warn",
                        "tight-gap",
                        a,
                        f'vertical gap to "{b.name}" is {inches(gap_y):.2f} in ({gap_y} EMU), under --gap '
                        f"{inches(ctx.gap):.2f} in; boxes y {span_axis(a.top, a.bottom)} / {span_axis(b.top, b.bottom)}",
                        "snap the boxes to the spacing grid (0.3 in) or intentionally tighten it to 0",
                        {"other_shape": b.name, "gap_emu": gap_y, "axis": "vertical"},
                    )
                    continue
            # horizontal neighbours (side by side): separated in x, aligned in y
            gap_x = max(a.left - b.right, b.left - a.right)
            if 0 < gap_x < ctx.gap:
                ov = _axis_overlap(a.top, a.bottom, b.top, b.bottom)
                aligned = ov >= GAP_OVERLAP_RATIO * min(a.height, b.height) or abs(a.top - b.top) <= ALIGN_TOL_EMU
                if aligned:
                    _add(
                        out,
                        "warn",
                        "tight-gap",
                        a,
                        f'horizontal gap to "{b.name}" is {inches(gap_x):.2f} in ({gap_x} EMU), under --gap '
                        f"{inches(ctx.gap):.2f} in; boxes x {span_axis(a.left, a.right)} / {span_axis(b.left, b.right)}",
                        "snap the boxes to the spacing grid (0.3 in) or intentionally tighten it to 0",
                        {"other_shape": b.name, "gap_emu": gap_x, "axis": "horizontal"},
                    )


def _estimate_height_pt(frame: Frame) -> tuple[float, int]:
    """(estimated text height in pt, total lines) for one frame."""
    total_pt = 0.0
    total_lines = 0
    usable_pt = frame.width_emu / EMU_PER_PT
    for text, size_pt, _font in frame.paragraphs:
        if not text.strip():
            lines = 1
        else:
            per_line = max(1.0, usable_pt / max(1.0, size_pt))
            lines = max(1, math.ceil(char_units(text) / per_line))
        total_lines += lines
        total_pt += lines * size_pt * LINE_SPACING
    return total_pt, total_lines


def check_text_quality(items: list[Item], ctx: Ctx, used_fonts: set[str], out: list[Finding]) -> None:
    for it in items:
        seen_fonts: set[str] = set()
        for frame in it.frames:
            for text, size_pt, font in frame.paragraphs:
                name = font or ctx.theme_font
                if name:
                    used_fonts.add(name)
                if name and name.casefold() not in ctx.fonts and name.casefold() not in seen_fonts:
                    seen_fonts.add(name.casefold())
                    _add(
                        out,
                        "warn",
                        "font",
                        it,
                        f"font {name!r} in {frame.where} is not in the allowlist "
                        f"({len(ctx.fonts)} faces); first use: {text[:40]!r}",
                        f"switch the run to an allowed face or pass --fonts {name} if this deck is meant to use it",
                        {"font": name, "where": frame.where, "size_pt": size_pt},
                    )
            if not any(text.strip() for text, _size, _font in frame.paragraphs):
                continue  # empty frame: nothing to overflow (decorative rects, chrome)
            estimated_pt, lines = _estimate_height_pt(frame)
            usable_pt = frame.height_emu / EMU_PER_PT
            if not it.is_chrome and usable_pt > 0 and estimated_pt > usable_pt * 1.05:
                _add(
                    out,
                    "warn",
                    "text-overflow",
                    it,
                    f"{frame.where} needs ~{lines} line(s) = {estimated_pt:.1f} pt ({estimated_pt / PT_PER_IN:.2f} in) "
                    f"but has {usable_pt:.1f} pt ({usable_pt / PT_PER_IN:.2f} in) of usable height "
                    f"(box {inches(it.width):.2f} x {inches(it.height):.2f} in)",
                    "enlarge the box, cut words, or lower the font size -- LibreOffice wraps differently than PowerPoint",
                    {
                        "estimated_pt": round(estimated_pt, 1),
                        "usable_pt": round(usable_pt, 1),
                        "lines": lines,
                        "where": frame.where,
                    },
                )
        if (
            it.geom_ok
            and not it.is_chrome
            and it.has_text_frame
            and not it.is_text
            and it.width >= EMPTY_MIN_W_EMU
            and it.height >= EMPTY_MIN_H_EMU
        ):
            _add(
                out,
                "warn",
                "empty-text-frame",
                it,
                f"text frame is empty but the box is {inches(it.width):.2f} x {inches(it.height):.2f} in "
                f"({it.width} x {it.height} EMU)",
                "delete the leftover box, or fill it (an empty box still swallows clicks and shows the caret)",
                {"width_emu": it.width, "height_emu": it.height},
            )
        for match in PLACEHOLDER_RE.finditer(it.text):
            _add(
                out,
                "warn",
                "placeholder-text",
                it,
                f"placeholder copy {match.group(0)!r} in {it.text[max(0, match.start() - 20) : match.end() + 20]!r}",
                "replace the lorem/TODO text with the real content before delivering the deck",
                {"match": match.group(0)},
            )
            break


def _image_pixels(shape: Any) -> tuple[int, int] | None:
    """(px_w, px_h) of a picture, or None when the image cannot be measured."""
    try:
        px_w, px_h = shape.image.size
    except Exception:  # noqa: BLE001 - unreadable (wmf/emf) or missing image part
        return None
    return (px_w, px_h) if px_w and px_h else None


def check_images(items: list[Item], out: list[Finding]) -> None:
    for it in items:
        if not it.geom_ok or not it.is_picture or not it.width or not it.height:
            continue
        pixels = _image_pixels(it.shape)
        if pixels is None:
            continue
        px_w, px_h = pixels
        dpi_x = px_w / inches(it.width)
        dpi_y = px_h / inches(it.height)
        dpi = min(dpi_x, dpi_y)
        if dpi < LOW_DPI:
            _add(
                out,
                "warn",
                "low-res-image",
                it,
                f"image is {px_w} x {px_h} px displayed at {inches(it.width):.2f} x {inches(it.height):.2f} in "
                f"= {dpi:.0f} dpi (min {LOW_DPI:.0f})",
                "re-export the image at >= 2x its display size (or display it smaller)",
                {"px_w": px_w, "px_h": px_h, "dpi": round(dpi, 1)},
            )


# --------------------------------------------------------------------------- #
# rendered PDF
# --------------------------------------------------------------------------- #


class PdfError(RuntimeError):
    pass


PAGE_RE = re.compile(r'<page width="([\d.]+)" height="([\d.]+)">(.*?)</page>', re.DOTALL)
WORD_RE = re.compile(
    r'<word xMin="(-?[\d.]+)" yMin="(-?[\d.]+)" xMax="(-?[\d.]+)" yMax="(-?[\d.]+)">(.*?)</word>',
    re.DOTALL,
)


def _run_pdftotext(pdf: str) -> str:
    try:
        proc = subprocess.run(
            ["pdftotext", "-bbox-layout", pdf, "-"],
            check=False,
            capture_output=True,
            text=True,
        )
    except FileNotFoundError as exc:
        raise PdfError("pdftotext (poppler-utils) not found") from exc
    if proc.returncode != 0:
        raise PdfError(f"pdftotext failed on {pdf}: {proc.stderr.strip() or proc.returncode}")
    return proc.stdout


def _emu_pt(emu: float) -> float:
    return emu / EMU_PER_PT


def check_pdf(
    pdf: str,
    slide_count: int,
    margin_in: float,
    chrome: dict[int, list[tuple[float, float, float, float]]],
    out: list[Finding],
) -> int:
    """Word boxes crossing the content margins; returns the page count.

    `chrome` maps slide number -> page-space rectangles (pt) of slide-number/
    date/footer placeholders; words inside them are page chrome, not content.
    """
    text = _run_pdftotext(pdf)
    pages = PAGE_RE.findall(text)
    margin_pt = margin_in * PT_PER_IN
    for page_no, (pw, ph, body) in enumerate(pages, 1):
        page_w, page_h = float(pw), float(ph)
        chrome_rects = chrome.get(page_no, [])
        worst: tuple[float, str, str, dict[str, float]] | None = None
        offenders = 0
        for x_min, y_min, x_max, y_max, raw in WORD_RE.findall(body):
            wx0, wy0, wx1, wy1 = float(x_min), float(y_min), float(x_max), float(y_max)
            if any(
                rx0 - 1.0 <= wx0 and wx1 <= rx1 + 1.0 and ry0 - 1.0 <= wy0 and wy1 <= ry1 + 1.0
                for rx0, ry0, rx1, ry1 in chrome_rects
            ):
                continue
            over = []
            if wx0 < margin_pt:
                over.append(f"left edge {wx0 / PT_PER_IN:.2f} in (< {margin_in:.2f} in)")
            if wx1 > page_w - margin_pt:
                over.append(f"right edge {wx1 / PT_PER_IN:.2f} in (> {page_w / PT_PER_IN - margin_in:.2f} in)")
            if wy0 < margin_pt:
                over.append(f"top edge {wy0 / PT_PER_IN:.2f} in (< {margin_in:.2f} in)")
            if wy1 > page_h - margin_pt:
                over.append(f"bottom edge {wy1 / PT_PER_IN:.2f} in (> {page_h / PT_PER_IN - margin_in:.2f} in)")
            if not over:
                continue
            offenders += 1
            excursion = max(margin_pt - wx0, wx1 - (page_w - margin_pt), margin_pt - wy0, wy1 - (page_h - margin_pt))
            if worst is None or excursion > worst[0]:
                word = html.unescape(re.sub(r"<[^>]+>", "", raw)).strip()
                worst = (
                    excursion,
                    word,
                    "; ".join(over),
                    {"x_min_pt": wx0, "y_min_pt": wy0, "x_max_pt": wx1, "y_max_pt": wy1},
                )
        if worst is not None:
            excursion, word, sides, box = worst
            out.append(
                Finding(
                    "warn",
                    "pdf-margin",
                    page_no,
                    "text",
                    f"{offenders} word box(es) cross the {margin_in:.2f} in content margin; worst {word!r} at {sides} "
                    f"-- {excursion:.1f} pt ({excursion / PT_PER_IN:.2f} in) past the margin, "
                    f"page {page_w:.1f} x {page_h:.1f} pt",
                    "move the offending text inward, or lower --margin if the bleed is intentional",
                    {"page": page_no, "words": offenders, "worst": word, **box},
                )
            )
    if len(pages) != slide_count:
        out.append(
            Finding(
                "error",
                "pdf-page-count",
                0,
                "deck",
                f"rendered PDF has {len(pages)} page(s) but the deck has {slide_count} slide(s)",
                "re-render the PDF from the current deck (soffice --headless --convert-to pdf) and check for hidden slides",
                {"pages": len(pages), "slides": slide_count},
            )
        )
    return len(pages)


# --------------------------------------------------------------------------- #


def main() -> int:
    ap = argparse.ArgumentParser(
        prog="lint_deck.py",
        description="Geometry/typography linter for .pptx decks (python-pptx).",
    )
    ap.add_argument("deck", help="path to a .pptx/.potx file")
    ap.add_argument("--json", action="store_true", help="machine-readable findings on stdout")
    ap.add_argument("--strict", action="store_true", help="warnings fail the run (exit 1)")
    ap.add_argument("--pdf", metavar="PDF", help="rendered PDF of the same deck (checks text margins + page count)")
    ap.add_argument(
        "--margin",
        type=float,
        default=0.5,
        metavar="IN",
        help="minimum distance from every slide edge, inches (default 0.5)",
    )
    ap.add_argument(
        "--gap",
        type=float,
        default=0.3,
        metavar="IN",
        help="minimum gap between neighbouring text boxes, inches (default 0.3)",
    )
    ap.add_argument(
        "--fonts",
        metavar="LIST",
        help="comma-separated font allowlist (default: " + ",".join(DEFAULT_FONTS) + ")",
    )
    args = ap.parse_args()

    if args.margin < 0 or args.gap < 0:
        print("lint: --margin and --gap must be >= 0", file=sys.stderr)
        return 2

    try:
        from pptx import Presentation  # type: ignore[import-not-found]
    except ImportError:
        print("lint: python-pptx is required (pip install python-pptx)", file=sys.stderr)
        return 2

    try:
        prs = Presentation(args.deck)
    except Exception as exc:  # noqa: BLE001 - report any parse failure verbatim
        print(f"lint: cannot open {args.deck}: {exc}", file=sys.stderr)
        return 3

    allowed = [f.strip() for f in args.fonts.split(",") if f.strip()] if args.fonts else list(DEFAULT_FONTS)
    ctx = Ctx(
        slide_w=_int(prs.slide_width),
        slide_h=_int(prs.slide_height),
        margin=round(args.margin * EMU_PER_IN),
        gap=round(args.gap * EMU_PER_IN),
        fonts={f.casefold() for f in allowed},
        theme_font=theme_minor_font(args.deck),
    )

    findings: list[Finding] = []
    used_fonts: set[str] = set()
    chrome_boxes: dict[int, list[tuple[float, float, float, float]]] = {}
    slides = list(prs.slides)
    stats = {"slides": len(slides), "shapes": 0, "charts": 0, "images": 0, "tables": 0, "no_geometry": 0}

    for slide_no, slide in enumerate(slides, 1):
        items = list(collect(slide.shapes, slide_no, IDENTITY))
        stats["shapes"] += len(items)
        stats["charts"] += sum(1 for it in items if it.is_chart)
        stats["images"] += sum(1 for it in items if it.is_picture)
        stats["tables"] += sum(1 for it in items if it.is_table)
        stats["no_geometry"] += sum(1 for it in items if not it.geom_ok)
        chrome_boxes[slide_no] = [
            (_emu_pt(it.left), _emu_pt(it.top), _emu_pt(it.right), _emu_pt(it.bottom))
            for it in items
            if it.is_chrome and it.geom_ok
        ]
        check_extents(items, findings)
        check_bounds(items, ctx, findings)
        check_overlap(items, findings)
        check_gaps(items, ctx, findings)
        check_text_quality(items, ctx, used_fonts, findings)
        check_images(items, findings)

    pages = None
    if args.pdf:
        try:
            pages = check_pdf(args.pdf, len(slides), args.margin, chrome_boxes, findings)
        except PdfError as exc:
            print(f"lint: {exc}", file=sys.stderr)
            return 4

    findings.sort(key=lambda f: (f.slide, SEVERITY_RANK[f.severity], f.code, f.shape))
    errors = sum(1 for f in findings if f.severity == "error")
    warnings = sum(1 for f in findings if f.severity == "warn")

    if args.json:
        print(
            json.dumps(
                {
                    "file": args.deck,
                    "slides": stats["slides"],
                    "shapes": stats["shapes"],
                    "fonts": sorted(used_fonts),
                    "charts": stats["charts"],
                    "images": stats["images"],
                    "tables": stats["tables"],
                    "shapes_without_geometry": stats["no_geometry"],
                    "pdf_pages": pages,
                    "margin_in": args.margin,
                    "gap_in": args.gap,
                    "errors": errors,
                    "warnings": warnings,
                    "findings": [f.as_dict() for f in findings],
                },
                indent=2,
            )
        )
        return 1 if errors or (args.strict and warnings) else 0

    for finding in findings:
        print(finding.as_line())
    canvas = f"{inches(ctx.slide_w):.2f} x {inches(ctx.slide_h):.2f} in"
    font_list = ", ".join(sorted(used_fonts)) if used_fonts else "none"
    print(
        f"{args.deck}: {stats['slides']} slides, {stats['shapes']} shapes, {stats['charts']} charts, "
        f"{stats['images']} images, {stats['tables']} tables on {canvas}; fonts: {font_list}; "
        f"{errors} error(s), {warnings} warning(s)" + (f"; pdf pages: {pages}" if pages is not None else "")
    )
    return 1 if errors or (args.strict and warnings) else 0


if __name__ == "__main__":
    sys.exit(main())
