#!/usr/bin/env python3
"""Exact Word table geometry helpers for python-docx.

python-docx can set visible table widths, but consistent Word/LibreOffice/
Google Docs rendering requires the same width to be present in three OOXML
places:
- table width: w:tblPr/w:tblW
- table grid:   w:tblGrid/w:gridCol
- every cell:   w:tcPr/w:tcW

Use apply_cell_widths() / set_table_width() after all rows are created.
"""

from __future__ import annotations

from typing import Iterable, Sequence

from docx.oxml import OxmlElement
from docx.oxml.ns import qn

DEFAULT_CONTENT_WIDTH_DXA = 9360
DEFAULT_CELL_MARGINS_DXA = {"top": 80, "bottom": 80, "start": 120, "end": 120}
DEFAULT_TABLE_INDENT_DXA = DEFAULT_CELL_MARGINS_DXA["start"]

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def column_widths_from_weights(
    weights: Sequence[float], total_width_dxa: int = DEFAULT_CONTENT_WIDTH_DXA
) -> list[int]:
    """Allocate integer DXA widths whose sum is exactly total_width_dxa."""
    if not weights:
        raise ValueError("weights must not be empty")
    if any(w <= 0 for w in weights):
        raise ValueError("all weights must be positive")

    total_weight = float(sum(weights))
    widths = [int(round(total_width_dxa * (w / total_weight))) for w in weights]
    widths[-1] += total_width_dxa - sum(widths)
    if any(w <= 0 for w in widths):
        raise ValueError(f"invalid computed widths: {widths}")
    return widths


def exact_column_widths(
    widths_dxa: Iterable[int], total_width_dxa: int = DEFAULT_CONTENT_WIDTH_DXA
) -> list[int]:
    """Adjust widths so their sum exactly matches total_width_dxa."""
    widths = [int(w) for w in widths_dxa]
    if not widths:
        raise ValueError("widths_dxa must not be empty")
    if any(w <= 0 for w in widths):
        raise ValueError("all column widths must be positive")
    widths[-1] += total_width_dxa - sum(widths)
    if any(w <= 0 for w in widths):
        raise ValueError(f"invalid adjusted widths: {widths}")
    return widths


def _set_cell_width(cell, width_dxa: int) -> None:
    tcPr = cell._tc.get_or_add_tcPr()
    tcW = tcPr.find(qn("w:tcW"))
    if tcW is None:
        tcW = OxmlElement("w:tcW")
        tcPr.append(tcW)
    tcW.set(qn("w:w"), str(width_dxa))
    tcW.set(qn("w:type"), "dxa")


def _set_table_width_dxa(table, width_dxa: int) -> None:
    tblPr = table._tbl.tblPr
    tblW = tblPr.find(qn("w:tblW"))
    if tblW is None:
        tblW = OxmlElement("w:tblW")
        tblPr.append(tblW)
    tblW.set(qn("w:w"), str(width_dxa))
    tblW.set(qn("w:type"), "dxa")


def _set_table_grid(table, widths: Sequence[int]) -> None:
    tbl = table._tbl
    grid = tbl.find(qn("w:tblGrid"))
    if grid is not None:
        tbl.remove(grid)
    grid = OxmlElement("w:tblGrid")
    for w in widths:
        gridCol = OxmlElement("w:gridCol")
        gridCol.set(qn("w:w"), str(w))
        grid.append(gridCol)
    tblPr = tbl.tblPr
    tbl.addnext(grid)


def apply_cell_widths(table, widths_dxa: Sequence[int]) -> None:
    """Apply exact column widths to every cell and the table grid."""
    widths = list(widths_dxa)
    for row in table.rows:
        for idx, cell in enumerate(row.cells):
            if idx < len(widths):
                _set_cell_width(cell, widths[idx])
    _set_table_width_dxa(table, sum(widths))
    _set_table_grid(table, widths)


def set_table_full_width(table, total_width_dxa: int = DEFAULT_CONTENT_WIDTH_DXA) -> None:
    """Force the table to span the full content width with equal columns."""
    n_cols = len(table.columns)
    widths = column_widths_from_weights([1.0] * n_cols, total_width_dxa)
    apply_cell_widths(table, widths)
