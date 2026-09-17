#!/usr/bin/env python3
"""Extract every embedded image from a .docx to a directory.

Usage:
    python3 examples/extract_images.py input.docx output_dir/
"""

from __future__ import annotations

import argparse
import os
from pathlib import Path

from docx import Document


def extract_images(docx_path: str, out_dir: str) -> int:
    doc = Document(docx_path)
    os.makedirs(out_dir, exist_ok=True)
    seen: set[str] = set()
    count = 0
    for part in doc.part.package.iter_parts():
        if not part.content_type.startswith("image/"):
            continue
        base = Path(part.partname).name or f"image{count}"
        name = base
        i = 1
        while name in seen:
            stem, ext = os.path.splitext(base)
            name = f"{stem}_{i}{ext}"
            i += 1
        seen.add(name)
        with open(os.path.join(out_dir, name), "wb") as f:
            f.write(part.blob)
        count += 1
    return count


def main() -> None:
    ap = argparse.ArgumentParser(description="Extract images from a .docx")
    ap.add_argument("docx", type=Path)
    ap.add_argument("out_dir", type=Path)
    args = ap.parse_args()

    n = extract_images(str(args.docx), str(args.out_dir))
    print(f"extracted {n} image(s) to {args.out_dir}")


if __name__ == "__main__":
    main()
