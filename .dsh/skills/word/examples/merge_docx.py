#!/usr/bin/env python3
"""Merge multiple .docx into one by appending body content.

Copies paragraphs and tables (text, styles, and table structure). Note:
relationship-backed content such as images and hyperlinks is not carried across
documents — rebuild those after merging if needed.

Usage:
    python3 examples/merge_docx.py output.docx part1.docx part2.docx part3.docx
"""

from __future__ import annotations

import argparse
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn


def append_document(target: Document, source_path: str) -> None:
    src = Document(source_path)
    body = target.element.body
    for element in list(src.element.body):
        if element.tag == qn("w:sectPr"):
            continue  # keep the target's final section properties
        body.append(element)


def main() -> None:
    ap = argparse.ArgumentParser(description="Merge multiple .docx into one")
    ap.add_argument("output", type=Path)
    ap.add_argument("inputs", type=Path, nargs="+")
    args = ap.parse_args()

    if len(args.inputs) < 1:
        raise SystemExit("at least one input required")

    target = Document(str(args.inputs[0]))
    for extra in args.inputs[1:]:
        append_document(target, str(extra))

    target.save(str(args.output))
    print(f"merged {len(args.inputs)} file(s); saved {args.output}")


if __name__ == "__main__":
    main()
