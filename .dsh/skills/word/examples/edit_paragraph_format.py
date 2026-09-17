#!/usr/bin/env python3
"""Set paragraph-level formatting: style, alignment, spacing, indent, line spacing.

Applies to all paragraphs, or only those containing --match text.

Usage:
    python3 examples/edit_paragraph_format.py input.docx output.docx \
        --align justify --space-before 6 --space-after 6 --line-spacing 1.5
    python3 examples/edit_paragraph_format.py input.docx output.docx \
        --style "Heading 1" --match "Chapter"
"""

from __future__ import annotations

import argparse
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt

ALIGN = {
    "left": WD_ALIGN_PARAGRAPH.LEFT,
    "center": WD_ALIGN_PARAGRAPH.CENTER,
    "right": WD_ALIGN_PARAGRAPH.RIGHT,
    "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
}


def main() -> None:
    ap = argparse.ArgumentParser(description="Edit paragraph-level formatting")
    ap.add_argument("input", type=Path)
    ap.add_argument("output", type=Path)
    ap.add_argument("--match", type=str, default=None, help="only paragraphs containing this text")
    ap.add_argument("--style", type=str, default=None, help="apply a named paragraph style")
    ap.add_argument("--align", type=str, choices=list(ALIGN), default=None)
    ap.add_argument("--space-before", type=float, default=None, help="points before paragraph")
    ap.add_argument("--space-after", type=float, default=None, help="points after paragraph")
    ap.add_argument("--line-spacing", type=float, default=None, help="e.g. 1.5 or 2.0")
    ap.add_argument("--first-line-indent", type=float, default=None, help="points")
    ap.add_argument("--left-indent", type=float, default=None, help="points")
    args = ap.parse_args()

    doc = Document(str(args.input))
    changed = 0
    for p in doc.paragraphs:
        if args.match and args.match not in p.text:
            continue
        pf = p.paragraph_format
        if args.style:
            p.style = doc.styles[args.style]
        if args.align:
            pf.alignment = ALIGN[args.align]
        if args.space_before is not None:
            pf.space_before = Pt(args.space_before)
        if args.space_after is not None:
            pf.space_after = Pt(args.space_after)
        if args.line_spacing is not None:
            pf.line_spacing = args.line_spacing
        if args.first_line_indent is not None:
            pf.first_line_indent = Pt(args.first_line_indent)
        if args.left_indent is not None:
            pf.left_indent = Pt(args.left_indent)
        changed += 1

    doc.save(str(args.output))
    print(f"changed {changed} paragraph(s); saved {args.output}")


if __name__ == "__main__":
    main()
