#!/usr/bin/env python3
"""Set fonts on runs, including East Asian (CJK) fonts.

Word stores Latin (w:ascii / w:hAnsi) and East Asian (w:eastAsia) fonts
separately; setting run.font.name alone does not change Chinese text. This
example sets both.

Usage:
    python3 examples/edit_fonts.py input.docx output.docx \
        --ascii "Times New Roman" --east "黑体" --size 12 --bold --color 000000
    python3 examples/edit_fonts.py input.docx output.docx --east "宋体" --match "正文"
"""

from __future__ import annotations

import argparse
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor


def set_run_font(run, ascii_font=None, east_asia=None, size=None, bold=None, italic=None, color=None) -> None:
    if ascii_font:
        run.font.name = ascii_font
    if east_asia:
        rPr = run._element.get_or_add_rPr()
        rFonts = rPr.get_or_add_rFonts()
        rFonts.set(qn("w:eastAsia"), east_asia)
    if size is not None:
        run.font.size = Pt(size)
    if bold is not None:
        run.bold = bold
    if italic is not None:
        run.italic = italic
    if color:
        run.font.color.rgb = RGBColor.from_string(color)


def main() -> None:
    ap = argparse.ArgumentParser(description="Set fonts on runs")
    ap.add_argument("input", type=Path)
    ap.add_argument("output", type=Path)
    ap.add_argument("--ascii", type=str, default=None, help="Latin font name")
    ap.add_argument("--east", type=str, default=None, help="East Asian (CJK) font name")
    ap.add_argument("--size", type=float, default=None, help="point size")
    ap.add_argument("--bold", type=bool, default=None, help="True/False")
    ap.add_argument("--italic", type=bool, default=None, help="True/False")
    ap.add_argument("--color", type=str, default=None, help="hex RGB, e.g. 000000")
    ap.add_argument("--match", type=str, default=None, help="only runs whose paragraph contains this text")
    args = ap.parse_args()

    doc = Document(str(args.input))
    changed = 0
    for p in doc.paragraphs:
        if args.match and args.match not in p.text:
            continue
        for run in p.runs:
            set_run_font(run, args.ascii, args.east, args.size, args.bold, args.italic, args.color)
            changed += 1

    doc.save(str(args.output))
    print(f"changed {changed} run(s); saved {args.output}")


if __name__ == "__main__":
    main()
