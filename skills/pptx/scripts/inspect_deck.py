#!/usr/bin/env python3
"""Inspect a .pptx: sanity-open, slide count, per-slide shape/text summary.

Requires python-pptx (pip install python-pptx).

usage: inspect.py deck.pptx [--max-chars N] [--quiet]

Exit codes: 0 ok, 2 python-pptx missing, 1 parse failure.
"""

from __future__ import annotations

import argparse
import sys


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("deck", help="path to a .pptx/.potx file")
    ap.add_argument("--max-chars", type=int, default=240, help="text preview length per slide")
    ap.add_argument("--quiet", action="store_true", help="only print the header line")
    args = ap.parse_args()

    try:
        from pptx import Presentation  # type: ignore[import-not-found]
    except ImportError:
        print("inspect: python-pptx is required (pip install python-pptx)", file=sys.stderr)
        return 2

    try:
        prs = Presentation(args.deck)
    except Exception as exc:  # noqa: BLE001 - report any parse failure verbatim
        print(f"inspect: cannot open {args.deck}: {exc}", file=sys.stderr)
        return 1

    slides = list(prs.slides)
    print(f"{args.deck}: OK - {len(slides)} slides, canvas {prs.slide_width} x {prs.slide_height} EMU")

    if args.quiet:
        return 0

    for idx, slide in enumerate(slides, 1):
        texts: list[str] = []
        shapes = charts = tables = pictures = 0
        for shape in slide.shapes:
            shapes += 1
            if getattr(shape, "has_chart", False):
                charts += 1
            if getattr(shape, "has_table", False):
                tables += 1
            if shape.shape_type == 13:  # PICTURE
                pictures += 1
            if getattr(shape, "has_text_frame", False):
                text = " ".join(shape.text_frame.text.split())
                if text:
                    texts.append(text)
        notes = ""
        if slide.has_notes_slide:
            note_text = " ".join(slide.notes_slide.notes_text_frame.text.split())
            notes = f" notes={note_text!r}" if note_text else ""
        preview = " | ".join(texts)[: args.max_chars]
        counters = f"{shapes} shapes"
        if charts or tables or pictures:
            counters += f" (charts={charts} tables={tables} pics={pictures})"
        print(f"  slide {idx}: {counters}{notes}")
        if preview:
            print(f"    {preview}")

    return 0


if __name__ == "__main__":
    sys.exit(main())
