#!/usr/bin/env python3
"""Search a term across body, tables, headers, and footers, with context.

Usage:
    python3 examples/search_docx.py input.docx "search term"
    python3 examples/search_docx.py input.docx "search term" --context 20 --case-sensitive
"""

from __future__ import annotations

import argparse
from pathlib import Path

from docx import Document


def search_paragraphs(paragraphs, term, label, context, case_sensitive) -> int:
    hits = 0
    for i, p in enumerate(paragraphs):
        haystack = p.text if case_sensitive else p.text.lower()
        needle = term if case_sensitive else term.lower()
        if needle not in haystack:
            continue
        hits += 1
        idx = haystack.index(needle)
        lo = max(0, idx - context)
        hi = min(len(haystack), idx + len(needle) + context)
        snippet = ("…" if lo > 0 else "") + haystack[lo:hi] + ("…" if hi < len(haystack) else "")
        print(f"[{label} #{i + 1}] {snippet}")
    return hits


def main() -> None:
    ap = argparse.ArgumentParser(description="Search a term across a .docx")
    ap.add_argument("docx", type=Path)
    ap.add_argument("term")
    ap.add_argument("--context", type=int, default=40, help="characters of context each side")
    ap.add_argument("--case-sensitive", action="store_true")
    args = ap.parse_args()

    if not args.term:
        raise SystemExit("empty search term")

    doc = Document(str(args.docx))
    total = 0
    total += search_paragraphs(doc.paragraphs, args.term, "body", args.context, args.case_sensitive)
    for ti, table in enumerate(doc.tables, start=1):
        for ri, row in enumerate(table.rows, start=1):
            for ci, cell in enumerate(row.cells, start=1):
                total += search_paragraphs(
                    cell.paragraphs, args.term, f"table{ti}.r{ri}.c{ci}", args.context, args.case_sensitive
                )
    for si, section in enumerate(doc.sections, start=1):
        total += search_paragraphs(
            section.header.paragraphs, args.term, f"header{si}", args.context, args.case_sensitive
        )
        total += search_paragraphs(
            section.footer.paragraphs, args.term, f"footer{si}", args.context, args.case_sensitive
        )
    print(f"\n{total} occurrence(s) of {args.term!r}")


if __name__ == "__main__":
    main()
