#!/usr/bin/env python3
"""Edit headers and footers: set text and add a page-number field.

Usage:
    python3 examples/edit_headers_footers.py input.docx output.docx \
        --header "我的文档" --footer "第 {page} 页"
    python3 examples/edit_headers_footers.py input.docx output.docx --page-number-footer
"""

from __future__ import annotations

import argparse
from pathlib import Path

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn


def set_text(paragraph, text: str) -> None:
    """Set paragraph text, expanding {page} into a live PAGE field."""
    if "{page}" in text:
        before, after = text.split("{page}", 1)
        if before:
            paragraph.add_run(before)
        _add_page_field(paragraph)
        if after:
            paragraph.add_run(after)
    else:
        paragraph.text = text


def _add_page_field(paragraph) -> None:
    run = paragraph.add_run()
    fld_begin = OxmlElement("w:fldChar")
    fld_begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = "PAGE"
    fld_end = OxmlElement("w:fldChar")
    fld_end.set(qn("w:fldCharType"), "end")
    run._r.append(fld_begin)
    run._r.append(instr)
    run._r.append(fld_end)


def main() -> None:
    ap = argparse.ArgumentParser(description="Edit headers and footers")
    ap.add_argument("input", type=Path)
    ap.add_argument("output", type=Path)
    ap.add_argument("--header", type=str, default=None)
    ap.add_argument("--footer", type=str, default=None)
    ap.add_argument("--page-number-footer", action="store_true", help="footer = page number only")
    args = ap.parse_args()

    doc = Document(str(args.input))
    for section in doc.sections:
        if args.header is not None:
            header = section.header
            header.is_linked_to_previous = False
            set_text(header.paragraphs[0], args.header)
        if args.page_number_footer:
            footer = section.footer
            footer.is_linked_to_previous = False
            footer.paragraphs[0].text = ""
            _add_page_field(footer.paragraphs[0])
        elif args.footer is not None:
            footer = section.footer
            footer.is_linked_to_previous = False
            set_text(footer.paragraphs[0], args.footer)

    doc.save(str(args.output))
    print(f"saved {args.output}")


if __name__ == "__main__":
    main()
