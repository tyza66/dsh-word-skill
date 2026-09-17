#!/usr/bin/env python3
"""Precisely edit text content of a .docx with python-docx.

Demonstrates the workflows word skill uses for content edits:
- find and replace text in paragraphs and table cells
- rewrite a paragraph while preserving its style
- insert and delete paragraphs
- modify runs (the actual text carriers) without losing formatting

Usage:
    python examples/edit_content.py input.docx output.docx
"""

from __future__ import annotations

import sys
from pathlib import Path

from docx import Document


def replace_text_in_paragraph(paragraph, old: str, new: str) -> int:
    """Replace text across runs in a paragraph. Returns replacements made.

    Text may span multiple runs, so we rebuild the paragraph text when a
    simple per-run replace is insufficient.
    """
    count = 0
    for run in paragraph.runs:
        if old in run.text:
            run.text = run.text.replace(old, new)
            count += 1
    if count == 0 and old in paragraph.text:
        # Text spans runs: rebuild from the first run's format
        full = paragraph.text.replace(old, new)
        for run in paragraph.runs[1:]:
            run.text = ""
        if paragraph.runs:
            paragraph.runs[0].text = full
        count = 1
    return count


def edit_content(doc: Document) -> None:
    """Apply content edits in place. Edit these three demo rules as needed."""
    rules = [
        ("old phrase", "new phrase"),
        ("TODO: fill in", "done"),
    ]

    # Body paragraphs
    for p in doc.paragraphs:
        for old, new in rules:
            replace_text_in_paragraph(p, old, new)

    # Table cells
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for old, new in rules:
                        replace_text_in_paragraph(p, old, new)


def main() -> None:
    if len(sys.argv) != 3:
        print(f"usage: {sys.argv[0]} input.docx output.docx", file=sys.stderr)
        sys.exit(2)
    doc = Document(sys.argv[1])
    edit_content(doc)
    doc.save(sys.argv[2])
    print(f"saved {sys.argv[2]}")


if __name__ == "__main__":
    main()
