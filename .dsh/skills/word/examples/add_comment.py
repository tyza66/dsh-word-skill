#!/usr/bin/env python3
"""Add a comment to a .docx, anchored to a text match.

Uses python-docx's native comment API (`Document.add_comment`), which creates
the comments part and anchors the reference range automatically (requires
python-docx >= 1.1).

Usage:
    python3 examples/add_comment.py input.docx output.docx "要评注的文字" "这是我的批注" --author "dsh"
"""

from __future__ import annotations

import argparse
from pathlib import Path

from docx import Document


def main() -> None:
    ap = argparse.ArgumentParser(description="Add a comment to a .docx")
    ap.add_argument("input", type=Path)
    ap.add_argument("output", type=Path)
    ap.add_argument("match")
    ap.add_argument("comment")
    ap.add_argument("--author", type=str, default="dsh")
    args = ap.parse_args()

    doc = Document(str(args.input))

    target_run = None
    for p in doc.paragraphs:
        for r in p.runs:
            if args.match in r.text:
                target_run = r
                break
        if target_run is not None:
            break
    if target_run is None:
        raise SystemExit(f"text not found: {args.match!r}")

    comment = doc.add_comment(target_run, text=args.comment, author=args.author)
    doc.save(str(args.output))
    print(f"added comment #{comment.comment_id}; saved {args.output}")


if __name__ == "__main__":
    main()
