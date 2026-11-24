#!/usr/bin/env python3
"""Bind a visual-review verdict to the exact deck and render set it judged.

A review can go stale: the deck is edited after the images were reviewed and
nobody looks again. This tool writes a receipt that stores the SHA-256 of the
deck, of every slide-*.jpg/png in the render directory, and the review JSON
payload itself. `--verify` re-hashes the current files and fails if anything
changed, so a stale verdict cannot silently pass QA.

usage:
  visual_review_receipt.py deck.pptx renders/ review.json [-o receipt.json]
  visual_review_receipt.py --verify receipt.json

review.json:
  {"schema_version": "visual_judgment_v1",
   "reviewer": {"type": "human"|"model", "name": "..."},
   "verdict": "pass"|"fail", "findings": [...]}

Exit codes: 0 pass/created, 1 verification mismatch, 2 bad input.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import re
import sys
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

RECEIPT_VERSION = "visual_review_receipt_v1"
REVIEW_VERSION = "visual_judgment_v1"
RENDER_RE = re.compile(r"^slide-\d+\.(?:jpe?g|png)$", re.IGNORECASE)
CHUNK = 1 << 20


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as fh:
        for chunk in iter(lambda: fh.read(CHUNK), b""):
            digest.update(chunk)
    return digest.hexdigest()


def load_json(path: Path, label: str) -> dict[str, Any]:
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeDecodeError, json.JSONDecodeError) as exc:
        raise ValueError(f"cannot read {label} {path}: {exc}") from exc
    if not isinstance(payload, dict):
        raise ValueError(f"{label} {path} must be a JSON object")
    return payload


def check_review(payload: dict[str, Any]) -> dict[str, Any]:
    version = payload.get("schema_version")
    if version != REVIEW_VERSION:
        raise ValueError(f"review schema_version must be {REVIEW_VERSION!r}, got {version!r}")
    reviewer = payload.get("reviewer")
    if not isinstance(reviewer, dict) or reviewer.get("type") not in ("human", "model") \
            or not isinstance(reviewer.get("name"), str) or not reviewer["name"].strip():
        raise ValueError("review.reviewer must be {type: 'human'|'model', name: '<name>'}")
    if payload.get("verdict") not in ("pass", "fail"):
        raise ValueError("review.verdict must be 'pass' or 'fail'")
    if not isinstance(payload.get("findings"), list):
        raise ValueError("review.findings must be a list")
    return payload


def render_records(renders_dir: Path) -> list[dict[str, str]]:
    """One {name, sha256} record per render.sh slide image, in slide order."""
    if not renders_dir.is_dir():
        raise ValueError(f"render directory not found: {renders_dir}")
    images = [p for p in renders_dir.iterdir() if p.is_file() and RENDER_RE.match(p.name)]
    if not images:
        raise ValueError(f"no slide-*.jpg/png images in {renders_dir}")
    images.sort(key=lambda p: (int(re.search(r"\d+", p.name).group()), p.name))
    return [{"name": p.name, "sha256": sha256(p)} for p in images]


def create(deck: Path, renders_dir: Path, review_path: Path, out: Path) -> int:
    deck = deck.expanduser().resolve()
    renders_dir = renders_dir.expanduser().resolve()
    review_path = review_path.expanduser().resolve()
    out = out.expanduser().resolve()
    if not deck.is_file():
        raise ValueError(f"deck not found: {deck}")
    review = check_review(load_json(review_path, "review"))
    receipt = {
        "schema_version": RECEIPT_VERSION,
        "generated_at": datetime.now(timezone.utc).isoformat(timespec="seconds"),
        "deck": {"path": str(deck), "name": deck.name, "sha256": sha256(deck)},
        "renders_dir": str(renders_dir),
        "renders": render_records(renders_dir),
        "review": review,
    }
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(json.dumps(receipt, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")
    print(f"{deck.name}: receipt written to {out} "
          f"({len(receipt['renders'])} renders, verdict={review['verdict']})")
    return 0


def _locate(stored: str, fallback_name: str, receipt_path: Path, is_dir: bool = False) -> Path:
    path = Path(stored).expanduser()
    if (path.is_dir() if is_dir else path.is_file()):
        return path
    near = receipt_path.parent / fallback_name
    return near if (near.is_dir() if is_dir else near.is_file()) else path


def verify(receipt_path: Path) -> int:
    receipt_path = receipt_path.expanduser().resolve()
    receipt = load_json(receipt_path, "receipt")
    if receipt.get("schema_version") != RECEIPT_VERSION:
        raise ValueError(f"unsupported receipt schema_version: {receipt.get('schema_version')!r}")
    deck = receipt.get("deck")
    if not isinstance(deck, dict) or not deck.get("sha256") or not deck.get("name"):
        raise ValueError("receipt has no deck name/hash")
    stored_renders = receipt.get("renders")
    if not isinstance(stored_renders, list):
        raise ValueError("receipt has no render records")

    problems: list[str] = []
    path = _locate(str(deck.get("path", "")), str(deck["name"]), receipt_path)
    if not path.is_file():
        problems.append(f"deck missing: {path}")
    elif sha256(path) != deck["sha256"]:
        problems.append(f"deck changed: {path}")

    render_dir = _locate(str(receipt.get("renders_dir", "")),
                         str(Path(str(receipt.get("renders_dir", ""))).name),
                         receipt_path, is_dir=True)
    try:
        current = render_records(render_dir)
    except ValueError as exc:
        current = []
        problems.append(str(exc))
    expected = {r["name"]: r["sha256"] for r in stored_renders if isinstance(r, dict)}
    actual = {r["name"]: r["sha256"] for r in current}
    for name in sorted(set(expected) | set(actual)):
        if name not in actual:
            problems.append(f"render missing: {name}")
        elif name not in expected:
            problems.append(f"render added: {name}")
        elif actual[name] != expected[name]:
            problems.append(f"render changed: {name}")

    review = receipt.get("review")
    verdict = review.get("verdict") if isinstance(review, dict) else "unknown"
    reviewer = review.get("reviewer", {}) if isinstance(review, dict) else {}
    reviewer_name = reviewer.get("name", "unknown") if isinstance(reviewer, dict) else "unknown"
    print(f"receipt: {receipt_path}")
    print(f"  deck: {deck['name']}  renders: {len(stored_renders)}  "
          f"reviewer: {reviewer_name}  verdict: {verdict}")
    if problems:
        for problem in problems:
            print(f"  MISMATCH {problem}")
        print("FAIL - the reviewed deck/renders do not match the current files; re-review required")
        return 1
    print("PASS - deck and renders match the receipt")
    return 0


def main() -> int:
    ap = argparse.ArgumentParser(
        prog="visual_review_receipt.py",
        description="Create or verify a hash-bound visual-review receipt.",
    )
    ap.add_argument("deck", nargs="?", help="deck that was reviewed")
    ap.add_argument("renders", nargs="?", help="render.sh output directory (slide-*.jpg/png)")
    ap.add_argument("review", nargs="?", help="review JSON payload")
    ap.add_argument("-o", "--out", default="receipt.json", metavar="FILE",
                    help="receipt to write (default: receipt.json)")
    ap.add_argument("--verify", metavar="FILE", help="verify this receipt against current files")
    args = ap.parse_args()

    try:
        if args.verify:
            return verify(Path(args.verify))
        if not (args.deck and args.renders and args.review):
            ap.error("create mode requires deck, renders and review arguments")
        return create(Path(args.deck), Path(args.renders), Path(args.review), Path(args.out))
    except (OSError, ValueError) as exc:
        print(f"receipt: {exc}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    sys.exit(main())
