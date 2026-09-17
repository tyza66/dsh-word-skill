#!/usr/bin/env python3
"""List available styles, or (re)assign paragraph styles.

Common uses:
- map "looks like a heading" paragraphs onto real Heading styles so the TOC
  and navigation work;
- switch body text onto a named style.

Usage:
    python3 examples/edit_styles.py input.docx --list
    python3 examples/edit_styles.py input.docx output.docx --match "Chapter" --style "Heading 1"
    python3 examples/edit_styles.py input.docx output.docx --style "Normal" --not-style "Title"
"""

from __future__ import annotations

import argparse
from pathlib import Path

from docx import Document


def main() -> None:
    ap = argparse.ArgumentParser(description="List or reassign paragraph styles")
    ap.add_argument("input", type=Path)
    ap.add_argument("output", type=Path, nargs="?", default=None)
    ap.add_argument("--list", action="store_true", help="print available styles and exit")
    ap.add_argument("--match", type=str, default=None, help="only paragraphs containing this text")
    ap.add_argument("--style", type=str, default=None, help="target paragraph style name")
    ap.add_argument("--not-style", type=str, default=None, help="skip paragraphs already in this style")
    args = ap.parse_args()

    doc = Document(str(args.input))

    if args.list:
        for s in doc.styles:
            print(f"{s.type}: {s.name}")
        return

    if not args.output or not args.style:
        raise SystemExit("provide output and --style (or use --list)")

    changed = 0
    for p in doc.paragraphs:
        cur = getattr(p.style, "name", "")
        if args.match and args.match not in p.text:
            continue
        if args.not_style and cur == args.not_style:
            continue
        p.style = doc.styles[args.style]
        changed += 1

    doc.save(str(args.output))
    print(f"changed {changed} paragraph(s); saved {args.output}")


if __name__ == "__main__":
    main()
