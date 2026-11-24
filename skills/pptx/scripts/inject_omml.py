#!/usr/bin/env python3
"""Inject native OMML equations into {{MATH:id}} placeholders in a .pptx.

Usage:
    python3 inject_omml.py input.pptx formulas.json output.pptx

LaTeX is converted to OMML with pandoc (-f latex -t docx); the m:oMathPara /
m:oMath element is extracted from the generated docx and wrapped in a14:m
inside the slide XML. Standard library only (ElementTree + zipfile).
"""

from __future__ import annotations

import argparse
import copy
import json
import re
import shutil
import subprocess
import sys
import tempfile
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006"
A14_NS = "http://schemas.microsoft.com/office/drawing/2010/main"
XML_NS = "http://www.w3.org/XML/1998/namespace"

# Register the prefixes PowerPoint expects before any parsing/serialization.
# Extra prefixes found in the actual documents are registered dynamically.
BASE_PREFIXES = {
    "a": A_NS,
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "m": M_NS,
    "mc": MC_NS,
    "a14": A14_NS,
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "w14": "http://schemas.microsoft.com/office/word/2010/wordml",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "cp": "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
    "dc": "http://purl.org/dc/elements/1.1/",
    "dcterms": "http://purl.org/dc/terms/",
    "xsi": "http://www.w3.org/2001/XMLSchema-instance",
}
for _prefix, _uri in BASE_PREFIXES.items():
    ET.register_namespace(_prefix, _uri)

PLACEHOLDER_RE = re.compile(r"\{\{MATH:([A-Za-z0-9_-]+)\}\}")
ID_RE = re.compile(r"[A-Za-z0-9_-]+")
NS_DECL_RE = re.compile(r"""xmlns:([A-Za-z_][\w.-]*)\s*=\s*(["'])(.*?)\2""")
DEFAULT_NS_RE = re.compile(r"""(?<![\w:])xmlns\s*=\s*(["'])(.*?)\1""")
IGNORABLE_RE = re.compile(r"""\bmc:Ignorable\s*=\s*(["'])(.*?)\1""")
XML_SPACE = f"{{{XML_NS}}}space"
PANDOC = "pandoc"


def warn(message: str) -> None:
    print(f"warning: {message}", file=sys.stderr)


def fail(message: str) -> None:
    print(f"error: {message}", file=sys.stderr)
    raise SystemExit(1)


def require_pandoc() -> None:
    """pandoc is a hard requirement: without it the command exits 1."""
    global PANDOC
    found = shutil.which("pandoc")
    if found is None:
        fail("pandoc not found; install it with 'brew install pandoc' (or your package manager)")
    PANDOC = found
    try:
        subprocess.run([PANDOC, "--version"], capture_output=True, timeout=10, check=True)
    except (OSError, subprocess.TimeoutExpired, subprocess.CalledProcessError) as exc:
        fail(f"pandoc is not usable ({exc}); install it with 'brew install pandoc'")


def register_prefixes(xml_text: str) -> None:
    """Keep original namespace prefixes instead of ElementTree's ns0/ns1."""
    for match in NS_DECL_RE.finditer(xml_text):
        prefix, uri = match.group(1), match.group(3)
        if prefix == "xml" or re.fullmatch(r"ns\d+", prefix):
            continue
        try:
            ET.register_namespace(prefix, uri)
        except ValueError:
            pass
    default = DEFAULT_NS_RE.search(xml_text[:4096])
    if default:
        try:
            ET.register_namespace("", default.group(2))
        except ValueError:
            pass


def find_root_tag(xml_text: str) -> tuple[int, int]:
    """Locate the first element's start tag, skipping declarations/comments."""
    pos = 0
    while True:
        pos = xml_text.find("<", pos)
        if pos < 0:
            raise ValueError("no root element")
        if xml_text.startswith("<!--", pos):
            end = xml_text.find("-->", pos + 4)
            pos = end + 3 if end >= 0 else pos + 4
        elif xml_text.startswith("<?", pos):
            end = xml_text.find("?>", pos + 2)
            pos = end + 2 if end >= 0 else pos + 2
        elif xml_text.startswith("<!", pos):
            end = xml_text.find(">", pos + 2)
            pos = end + 1 if end >= 0 else pos + 2
        elif xml_text.startswith("</", pos):
            pos += 2
        else:
            break
    quote = ""
    for idx in range(pos, len(xml_text)):
        char = xml_text[idx]
        if quote:
            if char == quote:
                quote = ""
        elif char in ('"', "'"):
            quote = char
        elif char == ">":
            return pos, idx + 1
    raise ValueError("unterminated root start tag")


def xml_declaration(xml_text: str) -> str:
    match = re.match(r"\s*<\?xml\b[^?]*\?>", xml_text, re.IGNORECASE | re.DOTALL)
    return match.group(0).strip() if match else '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'


def append_attr(tag: str, attribute: str) -> str:
    stripped = tag.rstrip()
    if stripped.endswith("/>"):
        return stripped[:-2].rstrip() + " " + attribute + "/>"
    return stripped[:-1].rstrip() + " " + attribute + ">"


def augment_root_tag(tag: str, serialized_tag: str = "") -> str:
    """Keep original declarations, add math namespaces, mark a14 as ignorable."""
    result = tag
    # ElementTree puts every namespace used by the new OMML subtree on its root.
    # Carry over anything the original slide root did not already declare.
    for match in NS_DECL_RE.finditer(serialized_tag):
        prefix, uri = match.group(1), match.group(3)
        if prefix in ("mc", "a14", "m"):
            continue
        if not re.search(rf"\bxmlns:{re.escape(prefix)}\s*=", result):
            result = append_attr(result, f'xmlns:{prefix}="{uri}"')
    default = DEFAULT_NS_RE.search(serialized_tag)
    if default and not DEFAULT_NS_RE.search(result):
        result = append_attr(result, f'xmlns="{default.group(2)}"')
    for prefix, uri in (("mc", MC_NS), ("a14", A14_NS), ("m", M_NS)):
        if not re.search(rf"\bxmlns:{prefix}\s*=", result):
            result = append_attr(result, f'xmlns:{prefix}="{uri}"')
    ignorable = IGNORABLE_RE.search(result)
    if ignorable:
        tokens = ignorable.group(2).split()
        if "a14" not in tokens:
            tokens.append("a14")
            result = result[: ignorable.start(2)] + " ".join(tokens) + result[ignorable.end(2) :]
    else:
        result = append_attr(result, 'mc:Ignorable="a14"')
    return result


def latex_to_omml(latex: str, mode: str) -> ET.Element | None:
    """Convert one formula via pandoc and return its m:oMathPara/m:oMath."""
    delimiters = "$$" if mode == "display" else "$"
    with tempfile.TemporaryDirectory(prefix="pi-pptx-omml-") as tmp:
        tex_path = Path(tmp) / "formula.tex"
        docx_path = Path(tmp) / "formula.docx"
        tex_path.write_text(f"{delimiters}{latex}{delimiters}\n", encoding="utf-8")
        try:
            result = subprocess.run(
                [PANDOC, str(tex_path), "-f", "latex", "-t", "docx", "-o", str(docx_path)],
                capture_output=True,
                text=True,
                timeout=60,
            )
        except (OSError, subprocess.TimeoutExpired) as exc:
            warn(f"pandoc failed ({exc})")
            return None
        if result.returncode != 0:
            detail = (result.stderr or result.stdout).strip().replace("\n", " ")[:300]
            warn(f"pandoc exited {result.returncode}: {detail}")
            return None
        try:
            with zipfile.ZipFile(docx_path) as archive:
                doc_xml = archive.read("word/document.xml")
            register_prefixes(doc_xml.decode("utf-8", errors="replace"))
            document = ET.fromstring(doc_xml)
            for xpath in (f".//{{{M_NS}}}oMathPara", f".//{{{M_NS}}}oMath"):
                math = document.find(xpath)
                if math is not None:
                    return math
            warn("pandoc produced no OMML element")
        except (KeyError, OSError, zipfile.BadZipFile, ET.ParseError) as exc:
            warn(f"cannot read pandoc docx output ({exc})")
    return None


def make_text_run(text: str, rpr: ET.Element | None) -> ET.Element:
    run = ET.Element(f"{{{A_NS}}}r")
    if rpr is not None:
        run.append(copy.deepcopy(rpr))
    text_elem = ET.SubElement(run, f"{{{A_NS}}}t")
    text_elem.text = text
    text_elem.set(XML_SPACE, "preserve")
    return run


def math_wrapper(omml: ET.Element) -> ET.Element:
    wrapper = ET.Element(f"{{{A14_NS}}}m")
    wrapper.append(copy.deepcopy(omml))
    return wrapper


def merge_intervals(intervals: list[tuple[int, int]]) -> list[tuple[int, int]]:
    merged: list[list[int]] = []
    for start, end in sorted(intervals):
        if merged and start <= merged[-1][1]:
            merged[-1][1] = max(merged[-1][1], end)
        else:
            merged.append([start, end])
    return [(start, end) for start, end in merged]


def nodes_for_run(
    text: str,
    removed: list[tuple[int, int]],
    insertions: list[tuple[int, ET.Element]],
    rpr: ET.Element | None,
) -> list[ET.Element]:
    """Split one run's text around removed placeholders and insert math nodes."""
    events: dict[int, list[ET.Element]] = {}
    for position, element in insertions:
        events.setdefault(position, []).append(element)
    points = {0, len(text)}
    for start, end in removed:
        points.update((start, end))
    points.update(events)
    ordered = sorted(points)
    nodes: list[ET.Element] = []
    for index, start in enumerate(ordered[:-1]):
        end = ordered[index + 1]
        if not any(rs <= start < re for rs, re in removed) and end > start:
            nodes.append(make_text_run(text[start:end], rpr))
        nodes.extend(events.get(start, []))
    nodes.extend(events.get(ordered[-1], []))
    return nodes


def change_paragraph(paragraph: ET.Element, omml_cache: dict[str, ET.Element]) -> tuple[int, set[str]]:
    """Replace placeholders in one paragraph, including runs split by pptxgenjs."""
    children = list(paragraph)
    runs = []
    for index, child in enumerate(children):
        if child.tag != f"{{{A_NS}}}r":
            continue
        text_elem = child.find(f"{{{A_NS}}}t")
        if text_elem is None or not text_elem.text:
            continue
        runs.append({"index": index, "elem": child, "text": text_elem.text})
    if not runs:
        return 0, set()

    full_text = "".join(run["text"] for run in runs)
    matches = list(PLACEHOLDER_RE.finditer(full_text))
    if not matches:
        return 0, set()

    ranges = []
    offset = 0
    for run in runs:
        ranges.append((offset, offset + len(run["text"]), run))
        offset += len(run["text"])

    removed: dict[int, list[tuple[int, int]]] = {id(run["elem"]): [] for run in runs}
    insertions: dict[int, list[tuple[int, ET.Element]]] = {id(run["elem"]): [] for run in runs}
    replaced = 0
    for match in matches:
        omml = omml_cache.get(match.group(1))
        if omml is None:
            continue  # failed conversion: leave the placeholder text untouched
        replaced += 1
        first_run = True
        for start, end, run in ranges:
            if end <= match.start() or start >= match.end():
                continue
            local_start = max(match.start(), start) - start
            local_end = min(match.end(), end) - start
            key = id(run["elem"])
            removed[key].append((local_start, local_end))
            if first_run:
                insertions[key].append((local_start, math_wrapper(omml)))
                first_run = False

    replacements: dict[int, tuple[int, ET.Element, list[ET.Element]]] = {}
    for run in runs:
        key = id(run["elem"])
        intervals = merge_intervals(removed[key])
        inserted = insertions[key]
        if not intervals and not inserted:
            continue
        rpr = run["elem"].find(f"{{{A_NS}}}rPr")
        replacements[key] = (
            run["index"],
            run["elem"],
            nodes_for_run(run["text"], intervals, inserted, rpr),
        )

    # Work backwards so earlier child indices stay valid.
    for index, elem, nodes in sorted(replacements.values(), key=lambda item: item[0], reverse=True):
        paragraph.remove(elem)
        for node in nodes:
            paragraph.insert(index, node)

    return replaced, {match.group(1) for match in matches}


def inject_into_slide(
    xml_bytes: bytes, omml_cache: dict[str, ET.Element]
) -> tuple[bytes, int, set[str]]:
    text = xml_bytes.decode("utf-8")
    register_prefixes(text)
    start, end = find_root_tag(text)
    original_root = text[start:end]
    declaration = xml_declaration(text)

    root = ET.fromstring(xml_bytes)
    replaced = 0
    found_ids: set[str] = set()
    for paragraph in root.iter(f"{{{A_NS}}}p"):
        count, found = change_paragraph(paragraph, omml_cache)
        replaced += count
        found_ids.update(found)
    if replaced == 0:
        return xml_bytes, replaced, found_ids

    serialized = ET.tostring(root, encoding="unicode", xml_declaration=False)
    new_start, new_end = find_root_tag(serialized)
    output = (
        declaration
        + serialized[:new_start]
        + augment_root_tag(original_root, serialized[new_start:new_end])
        + serialized[new_end:]
    )
    return output.encode("utf-8"), replaced, found_ids


def load_formulas(path: str) -> tuple[dict[str, dict[str, str]], int]:
    try:
        data = json.loads(Path(path).read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        fail(f"cannot read formulas.json '{path}': {exc}")
    if not isinstance(data, list):
        fail("formulas.json must be a JSON array of formula objects")

    formulas: dict[str, dict[str, str]] = {}
    invalid = 0
    for number, item in enumerate(data, 1):
        if not isinstance(item, dict):
            warn(f"formulas.json[{number}]: expected an object")
            invalid += 1
            continue
        fid, latex = item.get("id"), item.get("latex")
        mode = item.get("mode", "display")
        if not isinstance(fid, str) or not ID_RE.fullmatch(fid):
            warn(f"formulas.json[{number}]: id must match [A-Za-z0-9_-]+")
            invalid += 1
        elif not isinstance(latex, str) or not latex.strip():
            warn(f"formulas.json[{number}]: latex must be a non-empty string")
            invalid += 1
        elif mode not in ("display", "inline"):
            warn(f"formulas.json[{number}]: mode must be 'display' or 'inline'")
            invalid += 1
        elif fid in formulas:
            warn(f"formulas.json[{number}]: duplicate id '{fid}' ignored")
            invalid += 1
        else:
            formulas[fid] = {"latex": latex, "mode": mode}
    return formulas, invalid


def inject_pptx(input_path: str, formulas_path: str, output_path: str) -> dict[str, int]:
    formulas, invalid = load_formulas(formulas_path)
    if not Path(input_path).is_file():
        fail(f"input deck not found: {input_path}")

    try:
        with zipfile.ZipFile(input_path) as archive:
            entries = [(copy.copy(info), archive.read(info)) for info in archive.infolist()]
    except (OSError, zipfile.BadZipFile) as exc:
        fail(f"cannot read input deck '{input_path}': {exc}")

    omml_cache: dict[str, ET.Element] = {}
    failed = invalid
    for fid, formula in formulas.items():
        omml = latex_to_omml(formula["latex"], formula["mode"])
        if omml is None:
            warn(f"{fid}: conversion failed; placeholder left as-is")
            failed += 1
        else:
            omml_cache[fid] = omml

    success = 0
    found_ids: set[str] = set()
    updated = [(info, data) for info, data in entries]
    for index, (info, data) in enumerate(updated):
        name = info.filename
        if not (name.startswith("ppt/slides/slide") and name.endswith(".xml")):
            continue
        if b"{{MATH:" not in data:
            continue
        new_data, count, found = inject_into_slide(data, omml_cache)
        found_ids.update(found)
        if new_data != data:
            updated[index] = (info, new_data)
            success += count

    skipped = sum(1 for fid in omml_cache if fid not in found_ids)
    try:
        with zipfile.ZipFile(output_path, "w") as archive:
            for info, data in updated:
                archive.writestr(info, data)
    except OSError as exc:
        fail(f"cannot write output deck '{output_path}': {exc}")

    return {"success": success, "failed": failed, "skipped": skipped}


def main(argv: list[str] | None = None) -> None:
    parser = argparse.ArgumentParser(
        description="Replace {{MATH:id}} placeholders in a .pptx with native OMML equations.",
    )
    parser.add_argument("input_pptx", help="input .pptx generated from an outline with {{MATH:id}} text")
    parser.add_argument("formulas_json", help="JSON array of {id, latex, mode?} formula objects")
    parser.add_argument("output_pptx", help="output .pptx with native equations")
    args = parser.parse_args(argv)

    require_pandoc()
    stats = inject_pptx(args.input_pptx, args.formulas_json, args.output_pptx)
    print(json.dumps(stats, ensure_ascii=False))


if __name__ == "__main__":
    main()
