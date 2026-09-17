#!/usr/bin/env python3
"""Set page geometry: size, orientation, and margins.

Usage:
    python3 examples/set_page_setup.py input.docx output.docx --size A4 --orientation landscape
    python3 examples/set_page_setup.py input.docx output.docx --margins-cm 2.0
"""

from __future__ import annotations

import argparse
from pathlib import Path

from docx import Document
from docx.enum.section import WD_ORIENT
from docx.shared import Cm

PAGE_SIZES_CM = {
    "A4": (21.0, 29.7),
    "Letter": (21.59, 27.94),
    "A5": (14.8, 21.0),
    "B5": (17.6, 25.0),
}


def main() -> None:
    ap = argparse.ArgumentParser(description="Set page geometry")
    ap.add_argument("input", type=Path)
    ap.add_argument("output", type=Path)
    ap.add_argument("--size", type=str, choices=list(PAGE_SIZES_CM), default=None)
    ap.add_argument("--orientation", type=str, choices=["portrait", "landscape"], default=None)
    ap.add_argument("--margins-cm", type=float, default=None, help="set all four margins")
    args = ap.parse_args()

    doc = Document(str(args.input))
    for section in doc.sections:
        if args.size:
            w, h = PAGE_SIZES_CM[args.size]
            if args.orientation == "landscape":
                section.orientation = WD_ORIENT.LANDSCAPE
                section.page_width = Cm(h)
                section.page_height = Cm(w)
            else:
                section.orientation = WD_ORIENT.PORTRAIT
                section.page_width = Cm(w)
                section.page_height = Cm(h)
        elif args.orientation:
            section.orientation = (
                WD_ORIENT.LANDSCAPE if args.orientation == "landscape" else WD_ORIENT.PORTRAIT
            )
        if args.margins_cm is not None:
            section.left_margin = Cm(args.margins_cm)
            section.right_margin = Cm(args.margins_cm)
            section.top_margin = Cm(args.margins_cm)
            section.bottom_margin = Cm(args.margins_cm)

    doc.save(str(args.output))
    print(f"saved {args.output}")


if __name__ == "__main__":
    main()
