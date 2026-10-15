#!/usr/bin/env python3
"""Duplicate a slide inside a PPTX package (python3 standard library only).

Usage
-----
  # in-place: work on an already unpacked package directory
  python3 dup_slide.py <unpacked_dir> <slide-relative-path> [--after <slide-relative-path>]

  # one-shot: unpack a .pptx, duplicate, repack to a new file
  python3 dup_slide.py <deck.pptx> <slide-relative-path> --out <out.pptx> [--after <slide-relative-path>]

Arguments
---------
  package     Unpacked package directory (one containing ``[Content_Types].xml``),
              or a .pptx file when ``--out`` is given.
  slide       Slide part to duplicate, package-relative, e.g. ``ppt/slides/slide2.xml``.
              A bare ``slide2.xml`` is also accepted (resolved under ``ppt/slides``).
  --after     Insert the copy directly after the entry of this slide in
              ``<p:sldIdLst>``. Default: append at the end of the slide list.
  --out       Write the modified package to this .pptx instead of editing in place.

What it does (OOXML bookkeeping)
--------------------------------
  * copies ``ppt/slides/slideN.xml`` to the next free ``slideM.xml`` in the same
    directory (part names are unique package-wide, so the number is chosen
    against every ``slide???.xml`` part in the package);
  * copies ``ppt/slides/_rels/slideN.xml.rels`` to ``slideM.xml.rels`` verbatim,
    so images / notesSlides / charts become *shared* parts (no second copy is made
    and existing rIds inside the slide XML stay valid);
  * adds an ``<Override>`` for the new part to ``[Content_Types].xml`` (reusing the
    source slide's content type);
  * adds a ``<Relationship Type=.../slide>`` with a fresh unique ``rId`` to
    ``ppt/_rels/presentation.xml.rels``;
  * inserts ``<p:sldId id=... r:id=.../>`` into ``<p:sldIdLst>`` in
    ``ppt/presentation.xml`` with an unused id (>= 256 and greater than every
    existing id).

Output
------
Prints the package-relative path of the created slide part on stdout.
Usage problems exit 2 (argparse); any package inconsistency (missing parts,
unknown slide, already existing target) exits 1 with a message on stderr.
"""

import argparse
import os
import posixpath
import re
import shutil
import sys
import tempfile
import xml.etree.ElementTree as ET
import zipfile

PML_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
CONTENT_TYPES_PART = "[Content_Types].xml"
PRESENTATION_PART = "ppt/presentation.xml"
PRESENTATION_RELS_PART = "ppt/_rels/presentation.xml.rels"
SLIDE_RELTYPE = REL_NS + "/slide"
SLIDE_CONTENT_TYPE = (
    "application/vnd.openxmlformats-officedocument.presentationml.slide+xml"
)
SLIDE_NAME_RE = re.compile(r"slide(\d+)\.xml")
RID_RE = re.compile(r"rId(\d+)")
# Sub-elements of p:presentation that must follow p:sldIdLst (schema order).
AFTER_SLDIDLST = (
    "sldSz",
    "notesSz",
    "smartTags",
    "embeddedFontLst",
    "custShowLst",
    "photoAlbum",
    "custDataLst",
    "kinsoku",
    "defaultTextStyle",
    "modifyVerifier",
    "extLst",
)


class DupError(Exception):
    """Package inconsistency / unusable input (reported on stderr)."""


# --------------------------------------------------------------------------- xml


def _namespace_map(path):
    """Collect {prefix: uri} from the namespace declarations of ``path``."""
    nsmap = {}
    for _, decl in ET.iterparse(path, events=("start-ns",)):
        prefix, uri = decl
        nsmap[prefix] = uri
    return nsmap


def load(path):
    """Parse ``path``; return (ElementTree, namespace map of the original file)."""
    if not os.path.isfile(path):
        raise DupError("missing required part: %s" % path)
    nsmap = _namespace_map(path)
    try:
        tree = ET.parse(path)
    except ET.ParseError as exc:
        raise DupError("cannot parse %s: %s" % (path, exc))
    return tree, nsmap


def save(path, tree, nsmap):
    """Write ``tree`` back, re-registering the original prefixes first.

    ElementTree's namespace registry is global, so it is refreshed right before
    every write to keep ``p:``/``r:``/default prefixes exactly as they were.
    """
    for prefix, uri in nsmap.items():
        try:
            ET.register_namespace(prefix, uri)
        except ValueError:  # reserved prefix (e.g. "ns0") - leave as is
            pass
    tree.write(path, encoding="UTF-8", xml_declaration=True)


def local(tag):
    """Local name of a possibly namespaced tag."""
    return tag.rsplit("}", 1)[-1]


def qname(root, name):
    """Qualify ``name`` with the (default) namespace of ``root``, if any."""
    tag = root.tag
    if tag.startswith("{"):
        return "%s%s" % (tag[: tag.index("}") + 1], name)
    return name


def children(root, name):
    return [el for el in root if local(el.tag) == name]


def part_path(root, *parts):
    """Join package-relative part path(s) onto a filesystem root."""
    joined = "/".join(parts).replace("\\", "/")
    return os.path.join(root, *joined.split("/"))


# ------------------------------------------------------------------ package math


def resolve_slide_part(root, arg):
    """Turn a user-supplied slide path into a package-relative part path."""
    raw = arg.replace("\\", "/").lstrip("/")
    candidates = [posixpath.normpath(raw)]
    if "/" not in raw:
        candidates.append(posixpath.join("ppt", "slides", raw))
    for cand in candidates:
        if cand.startswith(".."):
            break
        if os.path.isfile(os.path.join(root, *cand.split("/"))):
            if not SLIDE_NAME_RE.fullmatch(posixpath.basename(cand)):
                raise DupError("not a slide part: %s" % arg)
            return cand
    raise DupError("unknown slide: %s" % arg)


def next_free_slide_part(root, src_part):
    """Next unused ``slideN.xml`` name, taken package-wide, in src's directory."""
    used = set()
    for dirpath, _dirnames, filenames in os.walk(root):
        for name in filenames:
            match = SLIDE_NAME_RE.fullmatch(name)
            if match:
                used.add(int(match.group(1)))
    number = max(used, default=0) + 1
    directory = posixpath.dirname(src_part)
    name = "slide%d.xml" % number
    return posixpath.join(directory, name) if directory else name


def resolve_target(base_dir, target):
    """Resolve a relationship target to a package-relative part path."""
    target = target.replace("\\", "/")
    if target.startswith("/"):
        return posixpath.normpath(target.lstrip("/"))
    return posixpath.normpath(posixpath.join(base_dir, target))


def unique_rid(existing_ids):
    highest = 0
    for rid in existing_ids:
        match = RID_RE.fullmatch(rid or "")
        if match:
            highest = max(highest, int(match.group(1)))
    number = highest + 1
    while "rId%d" % number in existing_ids:
        number += 1
    return "rId%d" % number


# --------------------------------------------------------------------- core work


def duplicate(root, slide_arg, after_arg=None):
    """Duplicate ``slide_arg`` inside the unpacked package ``root``.

    Returns the package-relative path of the created slide part.
    """
    root = os.path.abspath(root)
    ct_path = os.path.join(root, CONTENT_TYPES_PART)
    pres_path = part_path(root, PRESENTATION_PART)
    rels_path = part_path(root, PRESENTATION_RELS_PART)
    if not os.path.isfile(ct_path):
        raise DupError("not an unpacked package (no %s): %s" % (CONTENT_TYPES_PART, root))

    src_part = resolve_slide_part(root, slide_arg)
    src_xml_abs = part_path(root, src_part)
    src_rels_part = posixpath.join(
        posixpath.dirname(src_part),
        "_rels",
        posixpath.basename(src_part) + ".rels",
    )
    src_rels_abs = part_path(root, src_rels_part)
    after_part = resolve_slide_part(root, after_arg) if after_arg else None

    # --- parse the three package indexes (this also validates their presence)
    ct_tree, ct_nsmap = load(ct_path)
    rels_tree, rels_nsmap = load(rels_path)
    pres_tree, pres_nsmap = load(pres_path)
    ct_root = ct_tree.getroot()
    rels_root = rels_tree.getroot()
    pres_root = pres_tree.getroot()

    # --- pick the new part name
    dup_part = next_free_slide_part(root, src_part)
    dup_abs = part_path(root, dup_part)
    dup_rels_part = posixpath.join(
        posixpath.dirname(dup_part), "_rels", posixpath.basename(dup_part) + ".rels"
    )
    if os.path.lexists(dup_abs):
        raise DupError("duplicate target: %s already exists" % dup_part)

    # --- [Content_Types].xml: reuse the source slide's content type
    override_tag = qname(ct_root, "Override")
    content_type = SLIDE_CONTENT_TYPE
    for override in children(ct_root, "Override"):
        part_name = (override.get("PartName") or "").lstrip("/")
        if part_name == dup_part:
            raise DupError("duplicate target: %s already declared in %s" % (part_name, CONTENT_TYPES_PART))
        if part_name == src_part:
            content_type = override.get("ContentType") or content_type

    # --- ppt/_rels/presentation.xml.rels: fresh relationship
    slide_rels = {}  # part path -> rId
    existing_ids = set()
    for rel in children(rels_root, "Relationship"):
        rel_id = rel.get("Id")
        existing_ids.add(rel_id)
        if rel.get("Type") == SLIDE_RELTYPE and (rel.get("TargetMode") or "Internal") == "Internal":
            slide_rels[resolve_target("ppt", rel.get("Target") or "")] = rel_id
    if dup_part in slide_rels:
        raise DupError("duplicate target: %s already targeted by %s" % (dup_part, PRESENTATION_RELS_PART))
    if src_part not in slide_rels:
        raise DupError("unknown slide: %s is not referenced by %s" % (src_part, PRESENTATION_RELS_PART))
    src_rid = slide_rels[src_part]
    dup_rid = unique_rid(existing_ids)

    # --- ppt/presentation.xml: sldIdLst entry
    sld_id_tag = qname(pres_root, "sldId")
    r_id_attr = "{%s}id" % REL_NS
    sld_id_lst = None
    for el in pres_root:
        if local(el.tag) == "sldIdLst":
            sld_id_lst = el
            break
    by_rid = {}
    highest_id = 255  # ids must be >= 256
    if sld_id_lst is not None:
        for el in children(sld_id_lst, "sldId"):
            raw_id = el.get("id")
            if raw_id is None or not raw_id.isdigit():
                raise DupError("invalid <p:sldId id=%r> in %s" % (raw_id, PRESENTATION_PART))
            highest_id = max(highest_id, int(raw_id))
            by_rid[el.get(r_id_attr)] = el
    if src_rid not in by_rid:
        raise DupError("inconsistency: %s has no <p:sldId> entry (r:id=%s)" % (src_part, src_rid))
    anchor = None
    if after_part is not None:
        after_rid = slide_rels.get(after_part)
        if after_rid is None:
            raise DupError("unknown slide: --after %s is not referenced by %s" % (after_part, PRESENTATION_RELS_PART))
        anchor = by_rid.get(after_rid)
        if anchor is None:
            raise DupError("inconsistency: --after %s has no <p:sldId> entry" % after_part)

    # everything below only mutates files
    shutil.copyfile(src_xml_abs, dup_abs)
    if os.path.isfile(src_rels_abs):
        shutil.copyfile(src_rels_abs, part_path(root, dup_rels_part))

    override = ET.Element(override_tag)
    override.set("PartName", "/" + dup_part)
    override.set("ContentType", content_type)
    ct_root.append(override)
    save(ct_path, ct_tree, ct_nsmap)

    relationship = ET.Element(qname(rels_root, "Relationship"))
    relationship.set("Id", dup_rid)
    relationship.set("Type", SLIDE_RELTYPE)
    relationship.set("Target", posixpath.relpath(dup_part, "ppt"))
    rels_root.append(relationship)
    save(rels_path, rels_tree, rels_nsmap)

    new_sld_id = ET.Element(sld_id_tag)
    new_sld_id.set("id", str(highest_id + 1))
    new_sld_id.set(r_id_attr, dup_rid)
    if sld_id_lst is None:
        sld_id_lst = ET.Element(qname(pres_root, "sldIdLst"))
        index = len(pres_root)
        for i, el in enumerate(pres_root):
            if local(el.tag) in AFTER_SLDIDLST:
                index = i
                break
        pres_root.insert(index, sld_id_lst)
    if anchor is None:
        sld_id_lst.append(new_sld_id)
    else:
        sld_id_lst.insert(list(sld_id_lst).index(anchor) + 1, new_sld_id)
    save(pres_path, pres_tree, pres_nsmap)

    return dup_part


# ------------------------------------------------------------------ .pptx plumbing


def extract(pptx_path, dest_dir):
    try:
        with zipfile.ZipFile(pptx_path) as archive:
            archive.extractall(dest_dir)
    except (zipfile.BadZipFile, OSError) as exc:
        raise DupError("cannot read %s: %s" % (pptx_path, exc))


def repack(dir_path, pptx_path):
    parts = []
    for dirpath, _dirnames, filenames in os.walk(dir_path):
        for name in filenames:
            parts.append(os.path.relpath(os.path.join(dirpath, name), dir_path))
    parts = [p.replace(os.sep, "/") for p in parts]
    parts.sort(key=lambda p: (p != CONTENT_TYPES_PART, p))  # [Content_Types].xml first
    directory = os.path.dirname(os.path.abspath(pptx_path)) or "."
    handle, tmp_path = tempfile.mkstemp(prefix=".dup_slide_", suffix=".pptx", dir=directory)
    os.close(handle)
    try:
        with zipfile.ZipFile(tmp_path, "w", zipfile.ZIP_DEFLATED) as archive:
            for part in parts:
                archive.write(os.path.join(dir_path, *part.split("/")), part)
        os.replace(tmp_path, pptx_path)
    except BaseException:
        if os.path.exists(tmp_path):
            os.unlink(tmp_path)
        raise


# ---------------------------------------------------------------------------- cli


def build_parser():
    from_text = __doc__.split("Usage\n-----\n", 1)[-1].split("\n\n", 1)[0]
    parser = argparse.ArgumentParser(
        prog="dup_slide.py",
        description="Duplicate a slide inside a PPTX (unpacked package or .pptx).",
        epilog=from_text,
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument("package", help="unpacked package directory, or .pptx file with --out")
    parser.add_argument("slide", help="slide part to duplicate, e.g. ppt/slides/slide2.xml")
    parser.add_argument("--after", metavar="SLIDE", default=None,
                        help="insert the copy after this slide's <p:sldId> entry")
    parser.add_argument("--out", metavar="OUT", default=None,
                        help="write to OUT (.pptx) instead of editing the directory in place")
    return parser


def main(argv=None):
    args = build_parser().parse_args(argv)
    if args.out:
        if not os.path.isfile(args.package):
            raise DupError("--out needs a .pptx file as first argument: %s" % args.package)
        if os.path.abspath(args.package) == os.path.abspath(args.out):
            raise DupError("--out must differ from the input file")
        workdir = tempfile.mkdtemp(prefix="dup_slide_")
        try:
            extract(args.package, workdir)
            part = duplicate(workdir, args.slide, args.after)
            repack(workdir, args.out)
        finally:
            shutil.rmtree(workdir, ignore_errors=True)
    else:
        if os.path.isfile(args.package) and args.package.lower().endswith(".pptx"):
            raise DupError("%s is a .pptx file; pass --out <out.pptx> for file mode" % args.package)
        if not os.path.isdir(args.package):
            raise DupError("not a directory: %s" % args.package)
        part = duplicate(args.package, args.slide, args.after)
    sys.stdout.write(part + "\n")
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except DupError as exc:
        sys.stderr.write("dup_slide: error: %s\n" % exc)
        sys.exit(1)
