#!/usr/bin/env python3
"""Render DOCX-like files to PNG images via LibreOffice headless.

Generic, self-contained version for word skill. Converts input to PDF
with LibreOffice, then rasterizes the PDF to page-N.png images with pdf2image.
The rendered PNGs are the visual ground truth for verifying edits.

Dependencies:
    python-docx (not directly used, but expected in the same env)
    pdf2image     (pip install pdf2image)
    Pillow       (pip install Pillow)
    lxml         (pip install lxml)
    LibreOffice   (on PATH, or set SOFFICE_BIN)

Usage:
    python scripts/render_docx.py input.docx --output_dir out/
    python scripts/render_docx.py input.docx --output_dir out/ --width 1600 --height 2000
    python scripts/render_docx.py input.docx --output_dir out/ --emit_pdf --verbose
"""

from __future__ import annotations

import argparse
import glob
import os
import re
import shutil
import subprocess
import sys
import tempfile
import xml.etree.ElementTree as ET
from os import makedirs, replace
from os.path import abspath, basename, exists, expanduser, join, splitext
from pathlib import Path
from typing import Sequence, cast
from zipfile import ZipFile, ZipInfo

from pdf2image import convert_from_path, pdfinfo_from_path

TWIPS_PER_INCH: int = 1440


def _resolve_soffice() -> str:
    """Find a LibreOffice binary.

    Priority order:
      1. SOFFICE_BIN environment variable
      2. soffice on PATH
      3. soffice.exe on PATH (Windows)

    Returns the absolute path to the binary.
    """
    env_bin = os.environ.get("SOFFICE_BIN", "").strip()
    if env_bin and exists(env_bin):
        return os.path.abspath(env_bin)

    executable_name = "soffice.exe" if sys.platform == "win32" else "soffice"
    soffice = shutil.which(executable_name)
    if soffice is None:
        raise FileNotFoundError(
            f"LibreOffice ({executable_name}) not found on PATH. "
            "Install LibreOffice or set SOFFICE_BIN."
        )
    return os.path.abspath(soffice)


def _read_ooxml_member(zf: ZipFile, member_name: str) -> bytes:
    """Read an OOXML zip member, tolerating Windows-style backslash member names."""
    try:
        return zf.read(member_name)
    except KeyError:
        backslash_name = member_name.replace("/", "\\")
        if backslash_name != member_name:
            return zf.read(backslash_name)
        raise


def _zipinfo_with_filename(info: ZipInfo, filename: str) -> ZipInfo:
    out = ZipInfo(filename=filename, date_time=info.date_time)
    out.comment = info.comment
    out.extra = info.extra
    out.internal_attr = info.internal_attr
    out.external_attr = info.external_attr
    out.create_system = info.create_system
    out.compress_type = info.compress_type
    return out


def make_renderable_docx_copy(input_path: str, verbose: bool = False):
    """Return a DOCX path that common OOXML readers can open.

    Some generators write zip entries such as ``word\\document.xml`` and
    ``_rels\\.rels``. Word may still open those, but headless conversion tools
    expect canonical OOXML member names with ``/``.
    """
    if not input_path.lower().endswith((".docx", ".docm", ".dotx", ".dotm")):
        return input_path, None

    try:
        with ZipFile(input_path, "r") as zf:
            infos = zf.infolist()
            normalized_names = [info.filename.replace("\\", "/") for info in infos]
            should_repair = any(
                info.filename != normalized_name
                for info, normalized_name in zip(infos, normalized_names)
            )
            if not should_repair or "word/document.xml" not in normalized_names:
                return input_path, None

            temp_dir = tempfile.TemporaryDirectory(prefix="render_docx_ooxml_")
            repaired_path = join(temp_dir.name, basename(input_path))
            seen: set[str] = set()
            with ZipFile(repaired_path, "w") as zout:
                for info, normalized_name in zip(infos, normalized_names):
                    if normalized_name in seen:
                        continue
                    seen.add(normalized_name)
                    zout.writestr(
                        _zipinfo_with_filename(info, normalized_name),
                        zf.read(info.filename),
                    )
    except Exception:
        return input_path, None

    if verbose:
        print(f"[render] repaired: {input_path} -> {repaired_path}")
    return repaired_path, temp_dir


def calc_dpi_via_ooxml_docx(input_path: str, max_w_px: int, max_h_px: int) -> int:
    """Calculate DPI from OOXML page size (w:pgSz in twips)."""
    with ZipFile(input_path, "r") as zf:
        xml = _read_ooxml_member(zf, "word/document.xml")
    root = ET.fromstring(xml)
    ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}

    sect_pr = root.find(".//w:sectPr", ns)
    if sect_pr is None:
        raise RuntimeError("Section properties not found")
    pg_sz = sect_pr.find("w:pgSz", ns)
    if pg_sz is None:
        raise RuntimeError("Page size not found in section properties")

    w_ns = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
    w_twips_str = pg_sz.get(f"{w_ns}w") or pg_sz.get("w")
    h_twips_str = pg_sz.get(f"{w_ns}h") or pg_sz.get("h")

    if not w_twips_str or not h_twips_str:
        raise RuntimeError("Page size attributes missing in pgSz")

    width_in = int(w_twips_str) / TWIPS_PER_INCH
    height_in = int(h_twips_str) / TWIPS_PER_INCH
    if width_in <= 0 or height_in <= 0:
        raise RuntimeError("Invalid page size values")

    return round(min(max_w_px / width_in, max_h_px / height_in))


def _build_lo_env(user_profile: str) -> dict:
    env = os.environ.copy()
    env["HOME"] = user_profile
    env.setdefault("XDG_CONFIG_HOME", join(user_profile, "xdg_config"))
    env.setdefault("XDG_CACHE_HOME", join(user_profile, "xdg_cache"))
    os.makedirs(env["XDG_CONFIG_HOME"], exist_ok=True)
    os.makedirs(env["XDG_CACHE_HOME"], exist_ok=True)
    return env


def _run_cmd(cmd: list[str], env: dict, verbose: bool) -> subprocess.CompletedProcess:
    proc = subprocess.run(cmd, check=False, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True, env=env)
    if verbose:
        print("[render] $ " + " ".join(cmd))
        if proc.stdout:
            print(proc.stdout)
        if proc.stderr:
            print(proc.stderr)
    return proc


def convert_to_pdf(doc_path: str, user_profile: str, convert_tmp_dir: str, stem: str, verbose: bool) -> tuple[str, str]:
    """Convert input to PDF. Returns (pdf_path, debug_log)."""
    soffice = _resolve_soffice()
    env = _build_lo_env(user_profile)
    profile_argument = f"-env:UserInstallation={Path(user_profile).resolve().as_uri()}"
    logs: list[str] = []

    def _nonempty(path: str) -> bool:
        try:
            return exists(path) and os.path.getsize(path) > 0
        except Exception:
            return exists(path)

    pdf_path = join(convert_tmp_dir, f"{stem}.pdf")
    cmd_pdf = [soffice, profile_argument, "--invisible", "--headless", "--norestore", "--convert-to", "pdf", "--outdir", convert_tmp_dir, doc_path]
    proc = _run_cmd(cmd_pdf, env=env, verbose=verbose)
    logs.append(f"EXIT: {proc.returncode}\nSTDOUT:\n{proc.stdout}\nSTDERR:\n{proc.stderr}")

    if _nonempty(pdf_path):
        return pdf_path, "\n".join(logs)

    pdf_glob = glob.glob(join(convert_tmp_dir, "*.pdf"))
    if pdf_glob:
        cand = min(pdf_glob)
        if _nonempty(cand):
            return cand, "\n".join(logs)

    # Fallback: DOCX -> ODT -> PDF
    cmd_odt = [soffice, profile_argument, "--invisible", "--headless", "--norestore", "--convert-to", "odt", "--outdir", convert_tmp_dir, doc_path]
    _run_cmd(cmd_odt, env=env, verbose=verbose)
    odt_path = join(convert_tmp_dir, f"{stem}.odt")
    if exists(odt_path):
        cmd_odt_pdf = [soffice, profile_argument, "--invisible", "--headless", "--norestore", "--convert-to", "pdf", "--outdir", convert_tmp_dir, odt_path]
        _run_cmd(cmd_odt_pdf, env=env, verbose=verbose)
        if _nonempty(pdf_path):
            return pdf_path, "\n".join(logs)

    return "", "\n".join(logs)


def calc_dpi_via_pdf(input_path: str, max_w_px: int, max_h_px: int, verbose: bool) -> int:
    """Convert input to PDF and compute DPI from its page size."""
    with tempfile.TemporaryDirectory(prefix="soffice_profile_") as user_profile:
        with tempfile.TemporaryDirectory(prefix="soffice_convert_") as convert_tmp_dir:
            stem = splitext(basename(input_path))[0]
            pdf_path, debug = convert_to_pdf(input_path, user_profile, convert_tmp_dir, stem, verbose=verbose)
            if not (pdf_path and exists(pdf_path)):
                raise RuntimeError("Failed to convert input to PDF.\n" + debug)

            info = pdfinfo_from_path(pdf_path)
            size_val = info.get("Page size")
            if not size_val:
                raise RuntimeError("Failed to read PDF page size")
            m = re.search(r"(\d+(?:\.\d+)?)\s*x\s*(\d+(?:\.\d+)?)\s*pts", str(size_val))
            if not m:
                raise RuntimeError("Unrecognized PDF page size format.")
            width_in = float(m.group(1)) / 72.0
            height_in = float(m.group(2)) / 72.0
            return round(min(max_w_px / width_in, max_h_px / height_in))


def rasterize(doc_path: str, out_dir: str, dpi: int, verbose: bool, emit_pdf: bool) -> Sequence[str]:
    """Rasterize DOCX to images placed in out_dir, return their paths."""
    makedirs(out_dir, exist_ok=True)
    doc_path = abspath(doc_path)
    stem = splitext(basename(doc_path))[0]

    with tempfile.TemporaryDirectory(prefix="soffice_profile_") as user_profile:
        with tempfile.TemporaryDirectory(prefix="soffice_convert_") as convert_tmp_dir:
            pdf_path, debug = convert_to_pdf(doc_path, user_profile, convert_tmp_dir, stem, verbose=verbose)
            if not pdf_path or not exists(pdf_path):
                raise RuntimeError("Failed to produce PDF.\n" + debug)

            if emit_pdf:
                dst_pdf = join(out_dir, f"{stem}.pdf")
                tmp_pdf = dst_pdf + ".tmp"
                shutil.copy2(pdf_path, tmp_pdf)
                replace(tmp_pdf, dst_pdf)

            paths_raw = cast(
                list[str],
                convert_from_path(pdf_path, dpi=dpi, fmt="png", thread_count=8, output_folder=out_dir, paths_only=True, output_file="page"),
            )

    pages: list[tuple[int, str]] = []
    for src_path in paths_raw:
        base = splitext(basename(src_path))[0]
        page_num = int(base.split("-")[-1])
        dst_path = join(out_dir, f"page-{page_num}.png")
        replace(src_path, dst_path)
        pages.append((page_num, dst_path))
    pages.sort(key=lambda t: t[0])
    return [path for _, path in pages]


def main() -> None:
    parser = argparse.ArgumentParser(description="Render DOCX to PNG images")
    parser.add_argument("input_path", type=str)
    parser.add_argument("--output_dir", type=str, default=None)
    parser.add_argument("--width", type=int, default=1600)
    parser.add_argument("--height", type=int, default=2000)
    parser.add_argument("--dpi", type=int, default=None)
    parser.add_argument("--emit_pdf", action="store_true")
    parser.add_argument("--verbose", action="store_true")
    args = parser.parse_args()

    input_path = abspath(expanduser(args.input_path))
    out_dir = abspath(expanduser(args.output_dir)) if args.output_dir else splitext(input_path)[0]

    render_input_path, repair_temp_dir = make_renderable_docx_copy(input_path, verbose=args.verbose)
    try:
        if args.dpi is not None:
            dpi = int(args.dpi)
        else:
            try:
                if render_input_path.lower().endswith((".docx", ".docm", ".dotx", ".dotm")):
                    dpi = calc_dpi_via_ooxml_docx(render_input_path, args.width, args.height)
                else:
                    raise RuntimeError("not a DOCX container")
            except Exception:
                dpi = calc_dpi_via_pdf(render_input_path, args.width, args.height, verbose=args.verbose)

        rasterize(render_input_path, out_dir, dpi, verbose=args.verbose, emit_pdf=args.emit_pdf)
        print(f"Pages rendered to {out_dir}")
    finally:
        if repair_temp_dir is not None:
            repair_temp_dir.cleanup()


if __name__ == "__main__":
    main()
