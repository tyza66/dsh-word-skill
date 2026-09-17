#!/usr/bin/env python3
"""Print the heading outline of a .docx as an indented tree.

Useful for understanding document structure, generating a table of contents,
or locating a section before editing it.

Usage:
    python3 examples/read_outline.py input.docx
    python3 examples/read_outline.py input.docx --levels 2
"""

from __future__ import annotations

import argparse
from pathlib import Path

from docx import Document


def heading_level(style_name):
    if not style_name:
        return None
    s = style_name.strip().lower()
    if not s.startswith("heading"):
        return None
    parts = s.split()
    if len(parts) < 2:
        return None
    try:
        return int(parts[1])
    except Exception:
        return None


def main() -> None:
    ap = argparse.ArgumentParser(description="Print the heading outline of a .docx")
    ap.add_argument("docx", type=Path)
    ap.add_argument("--levels", type=int, default=0, help="max heading depth (0 = all)")
    args = ap.parse_args()

    doc = Document(str(args.docx))
    found = 0
    for p in doc.paragraphs:
        lvl = heading_level(getattr(p.style, "name", ""))
        if lvl is None:
            continue
        if args.levels and lvl > args.levels:
            continue
        text = (p.text or "").strip()
        print("  " * (lvl - 1) + f"[H{lvl}] {text}")
        found += 1
    if not found:
        print("(no heading-styled paragraphs found)")


if __name__ == "__main__":
    main()
