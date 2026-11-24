#!/usr/bin/env python3
"""OOXML package-structure validator for .pptx decks (stdlib only).

`lint_deck.py` judges layout; this script judges whether the package is put
together correctly. It reads the zip/XML directly, with no XSD, lxml or
python-pptx dependency, and checks:

  * zip integrity, `[Content_Types].xml`, and a declared content type for
    every part;
  * well-formed XML in every XML/rels part;
  * .rels Id uniqueness and resolvable internal targets (external targets
    are skipped);
  * <p:sldIdLst> slide registration vs. the slide parts present (no missing
    and no orphan slides);
  * media/chart/other relationship targets referenced from each slide;
  * a final sanity open with python-pptx, when that package is installed.

usage:
  validate_deck.py deck.pptx [--json]

Exit codes:
  0  no errors (warnings do not fail)
  1  structural errors found
  2  bad usage
  3  not a readable ZIP package
"""

from __future__ import annotations

import argparse
import json
import posixpath
import re
import sys
import urllib.parse
import xml.etree.ElementTree as ET
import zipfile
from dataclasses import dataclass, field
from typing import Any

CONTENT_TYPES = "[Content_Types].xml"
PRESENTATION = "ppt/presentation.xml"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
SLIDE_RE = re.compile(r"^ppt/slides/slide\d+\.xml$")


@dataclass
class Finding:
    severity: str  # "error" | "warning"
    code: str
    part: str
    message: str
    fix: str
    values: dict[str, Any] = field(default_factory=dict)

    def as_line(self) -> str:
        return (
            f'{self.severity.upper():<7} "{self.part}" [{self.code}]: '
            f"{self.message} -- fix: {self.fix}"
        )

    def as_dict(self) -> dict[str, Any]:
        return {
            "severity": self.severity,
            "code": self.code,
            "part": self.part,
            "message": self.message,
            "fix": self.fix,
            "values": self.values,
        }


def add(
    findings: list[Finding],
    severity: str,
    code: str,
    part: str,
    message: str,
    fix: str,
    **values: Any,
) -> None:
    findings.append(Finding(severity, code, part, message, fix, values))


def parse_part(
    zf: zipfile.ZipFile,
    name: str,
    findings: list[Finding],
    cache: dict[str, ET.Element | None] | None = None,
) -> ET.Element | None:
    """Parse one XML part, recording a finding on failure (once per part)."""
    if cache is not None and name in cache:
        return cache[name]
    root: ET.Element | None = None
    try:
        root = ET.fromstring(zf.read(name))
    except (KeyError, OSError) as exc:
        add(findings, "error", "PART_UNREADABLE", name, f"cannot read part: {exc}",
            "repack the deck so the part is stored correctly")
    except ET.ParseError as exc:
        add(findings, "error", "XML_NOT_WELL_FORMED", name, f"invalid XML: {exc}",
            "fix the XML in the source/generator, then repack")
    if cache is not None:
        cache[name] = root
    return root


def resolve_target(owner: str, target: str) -> str | None:
    """Resolve a package-relative relationship target to a zip part name."""
    path = urllib.parse.unquote(urllib.parse.urlsplit(target).path)
    if path.startswith("/"):
        return path.lstrip("/") or None
    base = posixpath.dirname(owner) if owner else ""
    resolved = posixpath.normpath(posixpath.join(base, path))
    if resolved.startswith("../") or resolved in ("..", "."):
        return None
    return resolved


def check_content_types(
    zf: zipfile.ZipFile,
    names: list[str],
    findings: list[Finding],
    cache: dict[str, ET.Element | None] | None = None,
) -> None:
    root = parse_part(zf, CONTENT_TYPES, findings, cache) if CONTENT_TYPES in names else None
    if root is None:
        if CONTENT_TYPES not in names:
            add(findings, "error", "MISSING_CONTENT_TYPES", CONTENT_TYPES,
                "the package has no content-type declarations",
                "repack with [Content_Types].xml at the package root")
        return

    overrides: set[str] = set()
    defaults: set[str] = set()
    for element in root:
        if element.tag == f"{{{CT_NS}}}Override":
            name = urllib.parse.unquote(element.get("PartName", "")).lstrip("/")
            overrides.add(name)
        elif element.tag == f"{{{CT_NS}}}Default":
            defaults.add(element.get("Extension", "").casefold())

    for name in names:
        if name == CONTENT_TYPES:
            continue
        ext = name.rsplit(".", 1)[-1].casefold() if "." in name else ""
        if name not in overrides and ext not in defaults:
            add(findings, "error", "UNDECLARED_CONTENT_TYPE", name,
                f"no Override or Default content type for extension {ext!r}",
                "declare the part in [Content_Types].xml", part_name=name, extension=ext)


def collect_relationships(
    zf: zipfile.ZipFile,
    names: list[str],
    findings: list[Finding],
    cache: dict[str, ET.Element | None] | None = None,
) -> dict[str, dict[str, dict[str, str]]]:
    """Map each source part to its {rel id: {target, mode, type}} relationships."""
    name_set = set(names)
    by_owner: dict[str, dict[str, dict[str, str]]] = {}
    for name in names:
        if not name.endswith(".rels"):
            continue
        if name == "_rels/.rels":
            owner = ""
        else:
            owner_dir = posixpath.dirname(posixpath.dirname(name))
            owner = posixpath.join(owner_dir, posixpath.basename(name)[:-5])
        root = parse_part(zf, name, findings, cache)
        if root is None:
            continue
        rels: dict[str, dict[str, str]] = {}
        by_owner[owner] = rels
        for element in root:
            if element.tag != f"{{{REL_NS}}}Relationship":
                continue
            rid = element.get("Id", "")
            target = element.get("Target", "")
            mode = element.get("TargetMode", "Internal")
            rel_type = element.get("Type", "")
            if not rid:
                add(findings, "error", "REL_NO_ID", name,
                    "relationship without Id", "give every relationship a unique Id")
                continue
            if rid in rels:
                add(findings, "error", "REL_DUPLICATE_ID", name,
                    f"duplicate relationship Id {rid!r}",
                    "renumber the relationships and update referencing XML", rid=rid)
            if not target:
                add(findings, "error", "REL_NO_TARGET", name,
                    f"relationship {rid!r} has no Target", "add the missing Target", rid=rid)
            if mode != "External":
                resolved = resolve_target(owner, target)
                if resolved is None:
                    add(findings, "error", "REL_OUTSIDE_PACKAGE", name,
                        f"relationship {rid!r} resolves outside the package: {target!r}",
                        "use a package-relative target or mark it External",
                        rid=rid, target=target)
                elif resolved not in name_set:
                    add(findings, "error", "REL_MISSING_TARGET", name,
                        f"relationship {rid!r} targets missing part {resolved!r}",
                        "add the part or remove the relationship",
                        rid=rid, target=target, resolved=resolved)
            rels[rid] = {"target": target, "mode": mode, "type": rel_type}
    return by_owner


def check_slide_registration(
    zf: zipfile.ZipFile,
    names: list[str],
    rels: dict[str, dict[str, dict[str, str]]],
    findings: list[Finding],
    cache: dict[str, ET.Element | None] | None = None,
) -> int:
    """Check presentation.xml's slide list; returns the registered slide count."""
    if PRESENTATION not in names:
        add(findings, "error", "MISSING_PRESENTATION", PRESENTATION,
            "the package has no presentation part", "repack with ppt/presentation.xml")
        return 0
    root = parse_part(zf, PRESENTATION, findings, cache)
    if root is None:
        return 0

    listed: list[str] = []
    seen_rids: set[str] = set()
    present_rels = rels.get(PRESENTATION, {})
    for sld_id in root.iter(f"{{{P_NS}}}sldId"):
        rid = sld_id.get(f"{{{R_NS}}}id", "")
        if not rid:
            add(findings, "error", "SLIDE_ID_NO_REL", PRESENTATION,
                "<p:sldId> without an r:id", "restore the relationship Id")
            continue
        if rid in seen_rids:
            add(findings, "error", "SLIDE_ID_DUPLICATE", PRESENTATION,
                f"duplicate slide relationship Id {rid!r}",
                "remove the duplicate <p:sldId>", rid=rid)
        seen_rids.add(rid)
        rel = present_rels.get(rid)
        if rel is None:
            add(findings, "error", "SLIDE_REL_MISSING", PRESENTATION,
                f"slide Id {rid!r} has no relationship",
                "register the slide in ppt/_rels/presentation.xml.rels", rid=rid)
            continue
        target = resolve_target(PRESENTATION, rel["target"])
        if target is None or not SLIDE_RE.match(target or ""):
            add(findings, "error", "SLIDE_TARGET_INVALID", PRESENTATION,
                f"slide Id {rid!r} points to {rel['target']!r}, not a slide part",
                "point the relationship at ppt/slides/slideN.xml", rid=rid, target=rel["target"])
            continue
        if target not in names:
            add(findings, "error", "SLIDE_PART_MISSING", target,
                f"registered slide {rid!r} has no part", "add the slide part", rid=rid)
        listed.append(target)

    existing = sorted(n for n in names if SLIDE_RE.match(n))
    for part in existing:
        if part not in listed:
            add(findings, "error", "SLIDE_ORPHAN", part,
                "slide part is not listed in <p:sldIdLst>",
                "register it in presentation.xml + its rels, or delete it")
    return len(listed)


def check_slide_references(
    zf: zipfile.ZipFile,
    names: list[str],
    rels: dict[str, dict[str, dict[str, str]]],
    findings: list[Finding],
    cache: dict[str, ET.Element | None] | None = None,
) -> None:
    """Check every relationship reference in a slide (images, charts, media...)."""
    name_set = set(names)
    for slide in sorted(n for n in names if SLIDE_RE.match(n)):
        root = parse_part(zf, slide, findings, cache)
        if root is None:
            continue
        slide_rels = rels.get(slide, {})
        for element in root.iter():
            for attr, rid in element.attrib.items():
                if not attr.startswith("{" + R_NS + "}") or not rid:
                    continue
                rel = slide_rels.get(rid)
                if rel is None:
                    add(findings, "error", "SLIDE_REFERENCE_UNRESOLVED", slide,
                        f"{element.tag.rsplit('}', 1)[-1]} references unknown rel {rid!r}",
                        "add the relationship in the slide's .rels part",
                        rid=rid, element=element.tag)
                    continue
                if rel["mode"] == "External":
                    continue
                target = resolve_target(slide, rel["target"])
                if target is None or target not in name_set:
                    add(findings, "error", "SLIDE_PART_REFERENCE_MISSING", slide,
                        f"{element.tag.rsplit('}', 1)[-1]} references missing part "
                        f"{rel['target']!r}",
                        "restore the media/chart part or remove the reference",
                        rid=rid, target=rel["target"], element=element.tag)


def check_python_pptx(deck: str, slides: int, findings: list[Finding]) -> None:
    """Best-effort final open; a missing python-pptx is only a warning."""
    try:
        from pptx import Presentation  # type: ignore[import-not-found]
    except ImportError:
        add(findings, "warning", "PYTHON_PPTX_UNAVAILABLE", deck,
            "python-pptx is not installed; final open check skipped",
            "pip install python-pptx to enable the extra check")
        return
    try:
        opened = Presentation(deck)
    except Exception as exc:  # noqa: BLE001 - report any parse failure verbatim
        add(findings, "error", "PYTHON_PPTX_OPEN_FAILED", deck,
            f"python-pptx cannot open the deck: {exc}",
            "fix the package errors above, or check python-pptx's error for details")
        return
    count = len(opened.slides)
    if count != slides:
        add(findings, "error", "SLIDE_COUNT_MISMATCH", deck,
            f"python-pptx sees {count} slides, <p:sldIdLst> lists {slides}",
            "reconcile presentation.xml with python-pptx's view",
            pptx_slides=count, listed_slides=slides)


def validate(deck: str) -> tuple[int, int, int, list[Finding]]:
    """Return (registered slides, parts, errors, findings)."""
    findings: list[Finding] = []
    with zipfile.ZipFile(deck) as zf:
        bad = zf.testzip()
        if bad:
            add(findings, "error", "ZIP_CRC", bad, "zip entry fails its CRC check",
                "repack the deck; the stored part is corrupt")
        names = [n for n in zf.namelist() if not n.endswith("/")]
        parts = len(names)
        cache: dict[str, ET.Element | None] = {}
        for name in names:
            if name.endswith(".xml") or name.endswith(".rels"):
                parse_part(zf, name, findings, cache)
        check_content_types(zf, names, findings, cache)
        rels = collect_relationships(zf, names, findings, cache)
        slides = check_slide_registration(zf, names, rels, findings, cache)
        check_slide_references(zf, names, rels, findings, cache)
    check_python_pptx(deck, slides, findings)
    errors = sum(1 for f in findings if f.severity == "error")
    return slides, parts, errors, findings


def main() -> int:
    ap = argparse.ArgumentParser(
        prog="validate_deck.py",
        description="Validate .pptx package structure and relationships (stdlib).",
    )
    ap.add_argument("deck", help="path to a .pptx/.potx file")
    ap.add_argument("--json", action="store_true", help="machine-readable findings on stdout")
    args = ap.parse_args()

    if not zipfile.is_zipfile(args.deck):
        print(f"validate: not a ZIP package: {args.deck}", file=sys.stderr)
        return 3
    try:
        slides, parts, errors, findings = validate(args.deck)
    except (OSError, zipfile.BadZipFile) as exc:
        print(f"validate: cannot read {args.deck}: {exc}", file=sys.stderr)
        return 3

    warnings = sum(1 for f in findings if f.severity == "warning")
    if args.json:
        print(json.dumps({
            "file": args.deck,
            "slides": slides,
            "parts": parts,
            "errors": errors,
            "warnings": warnings,
            "findings": [f.as_dict() for f in findings],
        }, indent=2))
    else:
        for finding in findings:
            print(finding.as_line())
        if errors:
            print(f"{args.deck}: {slides} slides, {parts} parts, {errors} error(s), "
                  f"{warnings} warning(s)")
        else:
            print(f"{args.deck}: OK - {slides} slides, {parts} parts, 0 errors"
                  + (f", {warnings} warning(s)" if warnings else ""))
    return 1 if errors else 0


if __name__ == "__main__":
    sys.exit(main())
