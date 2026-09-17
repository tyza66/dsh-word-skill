#!/usr/bin/env python3
"""Precisely edit table geometry and cell formatting in a .docx.

Demonstrates:
- apply exact column widths from content-derived weights
- force full content width with equal columns
- set cell padding/margins on every cell
- apply borders to the whole table at once

Usage:
    python examples/edit_tables.py input.docx output.docx
"""

from __future__ import annotations

import sys
from pathlib import Path

# Allow `from scripts.table_geometry import ...` regardless of the cwd.
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Twips

from scripts.table_geometry import apply_cell_widths, column_widths_from_weights

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def set_cell_margins(table, top=80, bottom=80, start=120, end=120) -> None:
    """Set uniform cell margins on every cell (values in twips/DXA)."""
    for row in table.rows:
        for cell in row.cells:
            tcPr = cell._tc.get_or_add_tcPr()
            tcMar = tcPr.find(qn("w:tcMar"))
            if tcMar is None:
                tcMar = OxmlElement("w:tcMar")
                tcPr.append(tcMar)
            for tag, val in (("top", top), ("bottom", bottom), ("start", start), ("end", end)):
                el = tcMar.find(qn(f"w:{tag}"))
                if el is None:
                    el = OxmlElement(f"w:{tag}")
                    tcMar.append(el)
                el.set(qn("w:w"), str(val))
                el.set(qn("w:type"), "dxa")


def set_table_borders(table, color="D9D9D9", size=4) -> None:
    """Apply visible borders to the whole table (size in eighths of a point)."""
    tblPr = table._tbl.tblPr
    borders = tblPr.find(qn("w:tblBorders"))
    if borders is None:
        borders = OxmlElement("w:tblBorders")
        tblPr.append(borders)
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = borders.find(qn(f"w:{edge}"))
        if el is None:
            el = OxmlElement(f"w:{edge}")
            borders.append(el)
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), str(size))
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), color)


def edit_first_table(doc: Document) -> None:
    if not doc.tables:
        return
    table = doc.tables[0]
    n_cols = len(table.columns)
    weights = [2.0] + [1.0] * (n_cols - 1) if n_cols > 1 else [1.0]
    widths = column_widths_from_weights(weights)
    apply_cell_widths(table, widths)
    set_cell_margins(table)
    set_table_borders(table)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER


def main() -> None:
    if len(sys.argv) != 3:
        print(f"usage: {sys.argv[0]} input.docx output.docx", file=sys.stderr)
        sys.exit(2)
    doc = Document(sys.argv[1])
    edit_first_table(doc)
    doc.save(sys.argv[2])
    print(f"saved {sys.argv[2]}")


if __name__ == "__main__":
    main()
