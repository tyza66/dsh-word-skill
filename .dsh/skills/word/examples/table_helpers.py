#!/usr/bin/env python3
"""Table helpers: add rows, merge cells, repeat header row, shade and align cells.

Run directly to apply a batch of common table edits to table #1, or import the
functions into your own script.

Usage:
    python3 examples/table_helpers.py input.docx output.docx
"""

from __future__ import annotations

import argparse
from pathlib import Path

from docx import Document
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_ROW_HEIGHT_RULE
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn


def add_row(table, values) -> None:
    """Append a row filled with `values`."""
    row = table.add_row()
    for i, v in enumerate(values):
        if i < len(row.cells):
            row.cells[i].text = v
    return row


def merge_cells(table, r1, c1, r2, c2):
    """Merge a rectangular block into one cell."""
    return table.cell(r1, c1).merge(table.cell(r2, c2))


def repeat_header_row(table, repeat: bool = True) -> None:
    """Make the first row repeat on every page."""
    row = table.rows[0]
    trPr = row._tr.get_or_add_trPr()
    el = trPr.find(qn("w:tblHeader"))
    if el is None:
        el = OxmlElement("w:tblHeader")
        trPr.append(el)
    el.set(qn("w:val"), "true" if repeat else "false")


def shade_cell(cell, fill: str) -> None:
    """Set cell background color (hex, e.g. 'D9D9D9')."""
    tcPr = cell._tc.get_or_add_tcPr()
    shd = tcPr.find(qn("w:shd"))
    if shd is None:
        shd = OxmlElement("w:shd")
        tcPr.append(shd)
    shd.set(qn("w:fill"), fill)


def set_cell(cell, text=None, align=None, v_align=None, fill=None) -> None:
    """Set text, horizontal/vertical alignment, and fill on a cell."""
    if text is not None:
        cell.text = text
    if align is not None:
        for p in cell.paragraphs:
            p.alignment = align
    if v_align is not None:
        cell.vertical_alignment = v_align
    if fill is not None:
        shade_cell(cell, fill)


def main() -> None:
    ap = argparse.ArgumentParser(description="Apply common table edits to table #1")
    ap.add_argument("input", type=Path)
    ap.add_argument("output", type=Path)
    args = ap.parse_args()

    doc = Document(str(args.input))
    if not doc.tables:
        raise SystemExit("no tables in document")

    table = doc.tables[0]
    # Demo: repeat header, shade header row, center header text.
    repeat_header_row(table)
    for cell in table.rows[0].cells:
        set_cell(cell, align=WD_ALIGN_PARAGRAPH.CENTER, v_align=WD_CELL_VERTICAL_ALIGNMENT.CENTER, fill="D9D9D9")

    doc.save(str(args.output))
    print(f"edited table #1 ({len(table.rows)} rows x {len(table.columns)} cols); saved {args.output}")


if __name__ == "__main__":
    main()
