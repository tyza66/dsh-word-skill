#!/usr/bin/env python3
"""Dump run-level formatting to diagnose visual inconsistencies.

Shows per-run font, East Asian font, size, bold/italic/underline, and color.
Direct run overrides are the usual cause of "why does this paragraph look
different from the rest".

Usage:
    python3 examples/read_formatting.py input.docx
    python3 examples/read_formatting.py input.docx --max-runs 100
"""

from __future__ import annotations

import argparse
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn


def run_info(run) -> str:
    f = run.font
    rPr = run._element.rPr
    east = None
    if rPr is not None and rPr.rFonts is not None:
        east = rPr.rFonts.get(qn("w:eastAsia"))
    color = f.color.rgb if (f.color and f.color.rgb) else None
    parts = []
    if f.name:
        parts.append(f"font={f.name}")
    if east:
        parts.append(f"eastAsia={east}")
    if f.size:
        parts.append(f"size={f.size.pt}pt")
    if run.bold is not None:
        parts.append(f"bold={run.bold}")
    if run.italic is not None:
        parts.append(f"italic={run.italic}")
    if run.underline:
        parts.append("underline")
    if color:
        parts.append(f"color=#{color}")
    return " ".join(parts) or "(inherited)"


def main() -> None:
    ap = argparse.ArgumentParser(description="Dump run-level formatting")
    ap.add_argument("docx", type=Path)
    ap.add_argument("--max-runs", type=int, default=0, help="limit output (0 = all)")
    args = ap.parse_args()

    doc = Document(str(args.docx))
    count = 0
    for i, p in enumerate(doc.paragraphs, start=1):
        for run in p.runs:
            if args.max_runs and count >= args.max_runs:
                return
            text = run.text.strip()
            if not text:
                continue
            style = getattr(p.style, "name", "")
            print(f"p{i} [{style}] {text[:50]!r} -> {run_info(run)}")
            count += 1
    if count == 0:
        print("(no runs with text found)")


if __name__ == "__main__":
    main()
