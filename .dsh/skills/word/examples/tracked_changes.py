#!/usr/bin/env python3
"""Add a tracked replacement (w:ins) to a .docx via raw OOXML patch.

python-docx does not expose tracked changes. This is the pattern the
word skill documents for editors: wrap the replacement at the OOXML
level so Word shows it as a revision.

Usage:
    python examples/tracked_changes.py input.docx output.docx
"""

from __future__ import annotations

import shutil
import sys
import zipfile
from datetime import datetime, timezone
from pathlib import Path

from lxml import etree

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W = f"{{{W_NS}}}"


def _make_ins_run(text: str, author: str, date: str, run_template) -> etree.Element:
    """Build a w:ins element containing a copy of the run template."""
    ins = etree.Element(f"{W}ins")
    ins.set(f"{W}id", "0")
    ins.set(f"{W}author", author)
    ins.set(f"{W}date", date)
    # Copy formatting from the template run
    rPr = run_template.find(f"{W}rPr")
    if rPr is not None:
        new_rPr = etree.fromstring(etree.tostring(rPr))
        ins.append(new_rPr)
    t = etree.SubElement(ins, f"{W}t")
    t.text = text
    return ins


def add_tracked_replacement(docx_path: str, output_path: str, old: str, new: str, author: str = "dsh") -> None:
    """Replace `old` with `new` in the first matching paragraph, tracked."""
    with zipfile.ZipFile(docx_path, "r") as zin:
        names = zin.namelist()
        document_xml = zin.read("word/document.xml")

    root = etree.fromstring(document_xml)
    replaced = False
    for p in root.iter(f"{W}p"):
        for r in p.findall(f"{W}r"):
            t = r.find(f"{W}t")
            if t is not None and t.text and old in t.text:
                t.text = t.text.replace(old, new, 1)
                # Wrap the run in w:ins for tracked change
                ins = _make_ins_run(
                    t.text, author, datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"), r
                )
                p.replace(r, ins)
                replaced = True
                break
        if replaced:
            break

    if not replaced:
        raise SystemExit(f"text not found: {old!r}")

    shutil.copy2(docx_path, output_path)
    tmp = output_path + ".tmp"
    with zipfile.ZipFile(docx_path, "r") as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "word/document.xml":
                data = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
            zout.writestr(item, data)
    Path(tmp).replace(output_path)


def main() -> None:
    if len(sys.argv) != 4:
        print(f"usage: {sys.argv[0]} input.docx output.docx 'old text'", file=sys.stderr)
        sys.exit(2)
    add_tracked_replacement(sys.argv[1], sys.argv[2], sys.argv[3], f"[tracked] {sys.argv[3]}")
    print(f"saved {sys.argv[2]}")


if __name__ == "__main__":
    main()
