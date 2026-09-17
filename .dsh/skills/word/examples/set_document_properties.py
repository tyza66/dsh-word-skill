#!/usr/bin/env python3
"""Read or set document core properties (title, author, subject, keywords, ...).

Usage:
    python3 examples/set_document_properties.py input.docx --show
    python3 examples/set_document_properties.py input.docx output.docx \
        --title "论文标题" --author "张三" --subject "主题" --keywords "关键词1,关键词2"
"""

from __future__ import annotations

import argparse
from pathlib import Path

from docx import Document


def show(cp) -> None:
    for name in ("title", "author", "subject", "keywords", "category", "comments", "last_modified_by"):
        val = getattr(cp, name, None)
        if val:
            print(f"{name}: {val}")


def main() -> None:
    ap = argparse.ArgumentParser(description="Read or set document core properties")
    ap.add_argument("input", type=Path)
    ap.add_argument("output", type=Path, nargs="?", default=None)
    ap.add_argument("--show", action="store_true", help="print current properties and exit")
    ap.add_argument("--title", type=str, default=None)
    ap.add_argument("--author", type=str, default=None)
    ap.add_argument("--subject", type=str, default=None)
    ap.add_argument("--keywords", type=str, default=None)
    ap.add_argument("--category", type=str, default=None)
    ap.add_argument("--comments", type=str, default=None)
    ap.add_argument("--last-modified-by", type=str, default=None)
    args = ap.parse_args()

    doc = Document(str(args.input))
    cp = doc.core_properties

    if args.show:
        show(cp)
        return

    if args.output is None:
        raise SystemExit("provide output path (or use --show)")

    if args.title is not None:
        cp.title = args.title
    if args.author is not None:
        cp.author = args.author
    if args.subject is not None:
        cp.subject = args.subject
    if args.keywords is not None:
        cp.keywords = args.keywords
    if args.category is not None:
        cp.category = args.category
    if args.comments is not None:
        cp.comments = args.comments
    if args.last_modified_by is not None:
        cp.last_modified_by = args.last_modified_by

    doc.save(str(args.output))
    print(f"saved {args.output}")


if __name__ == "__main__":
    main()
