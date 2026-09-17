#!/usr/bin/env python3
"""Read and dump the structured content of a .docx file.

Usage:
    python examples/read_content.py input.docx
    python examples/read_content.py input.docx --max-paragraphs 50
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from docx import Document


def read_content(docx_path: str, max_paragraphs: int = 0) -> None:
    doc = Document(docx_path)

    print(f"=== BODY PARAGRAPHS ({len(doc.paragraphs)}) ===")
    for i, p in enumerate(doc.paragraphs, start=1):
        if max_paragraphs and i > max_paragraphs:
            print(f"... ({len(doc.paragraphs) - max_paragraphs} more)")
            break
        style = getattr(p.style, "name", "")
        text = p.text.strip()
        if text:
            print(f"[{style}] {text}")

    print(f"\n=== TABLES ({len(doc.tables)}) ===")
    for ti, table in enumerate(doc.tables, start=1):
        print(f"\n--- Table {ti}: {len(table.rows)} rows x {len(table.columns)} cols ---")
        for row in table.rows:
            cells = [cell.text.strip().replace("\n", " | ") for cell in row.cells]
            print(" || ".join(cells))

    print(f"\n=== SECTIONS ({len(doc.sections)}) ===")
    for si, section in enumerate(doc.sections, start=1):
        print(
            f"Section {si}: page={section.page_width}x{section.page_height}, "
            f"margins L={section.left_margin} R={section.right_margin}"
        )
        for p in section.header.paragraphs:
            if p.text.strip():
                print(f"  header: {p.text.strip()}")
        for p in section.footer.paragraphs:
            if p.text.strip():
                print(f"  footer: {p.text.strip()}")


def main() -> None:
    ap = argparse.ArgumentParser(description="Read .docx content")
    ap.add_argument("docx", type=Path)
    ap.add_argument("--max_paragraphs", type=int, default=0)
    args = ap.parse_args()
    read_content(str(args.docx), args.max_paragraphs)


if __name__ == "__main__":
    main()
