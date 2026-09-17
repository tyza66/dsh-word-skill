---
name: word
description: Read a Word (.docx) document's content or make precise edits to its content and formatting. Use when the user wants to inspect, summarize, extract, or search text in a Word file, or asks to modify wording, restructure sections, adjust heading styles, tweak tables, fix spacing, headers, footers, or otherwise refine a document's content and visual layout. Do not use for spreadsheets, slides, or PDFs.
---

# Word Skill

A single entrypoint for two Word capabilities: **reading** `.docx` content and **precisely editing** both content and formatting.

## Capabilities

- **Read** a `.docx`: extract and return structured text (summaries, answers, targeted lookups).
- **Edit precisely**: change wording, restructure sections, adjust styles/headings, fix tables, spacing, headers, footers, and other layout details — then verify the result visually.

## Runtime and dependencies

Work is done with Python through the `bash` tool — there is no separate "documents" service to route to. Resolve the interpreter with `python3` (or `python`), never a global package assumption.

Required Python packages for editing and rendering:

```bash
python3 -m pip install python-docx lxml pdf2image
```

- `python-docx` — paragraphs, runs, styles, tables, headers/footers.
- `lxml` — raw OOXML patches (tracked changes, comments, hyperlinks, fields).
- `pdf2image` + `Pillow` — rasterizing the rendered PDF into `page-<N>.png`.

Rendering also requires **LibreOffice** on `PATH` (or set `SOFFICE_BIN`). If it is absent, install it, or skip the render-verify step and state clearly that visual verification was not run.

## Golden path

1. **Read first.** For any non-trivial edit, dump the current content with `examples/read_content.py` (and `scripts/word_audit.py` when formatting matters) so you edit with full context, not assumptions.
2. **Author the edit.** Use `python-docx` for paragraphs, runs, styles, tables, and headers/footers. For tracked changes, comments, hyperlinks, or fields, patch OOXML directly (see `examples/tracked_changes.py`).
3. **Render and verify.** Run `scripts/render_docx.py` to produce `page-<N>.png`, then inspect every page at 100% zoom (see *Inspecting rendered pages* below for how to actually look at them in this environment). Fix anything off and repeat until flawless. Word's own view is not ground truth — the rendered PNGs are.
4. **Deliver the `.docx` only**, unless the user explicitly asks for intermediates.

## Inspecting rendered pages (vision)

DSH does not bundle a built-in document reader, so how you *look at* the rendered `page-<N>.png` files depends on what the current session can see. Pick the first option that applies:

1. **The model has native vision or an image-input tool** (e.g. a `read_image` tool): read each `page-<N>.png` directly and check the layout at 100% zoom.
2. **Text-only model, but a vision/OCR skill or plugin is available** (e.g. a `see` skill that proxies an external vision model or local OCR): route the PNGs through it and act on the returned markdown. If the session catalog lists such a skill, load it and use it instead of claiming you cannot see images.
3. **No vision path at all:** fall back to text-only QA — `scripts/word_audit.py`, XML/OOXML inspection, and LibreOffice stderr — and tell the user plainly that visual verification was not performed, rather than asserting the layout is correct.

Never refuse a request to look at a document with "I can't see images": check for an image-input tool first, then a vision skill, before degrading to text-only checks.

## Precision editing

When the task is specifically about formatting accuracy, read `references/format-precision.md` before authoring. It captures the non-obvious checks (cell padding, border consistency, run-vs-paragraph spacing, tracked-changes visibility) that separate a visually correct document from a hand-tweaked one.

## Scripts and examples

This skill ships generic, self-contained tooling. Prefer these over ad-hoc scripts. Run them from the skill directory (`.dsh/skills/word/`) so relative imports resolve.

Scripts:

- `scripts/render_docx.py` — render any `.docx` to `page-<N>.png` via LibreOffice headless; the visual ground truth for verify. Supports `--dpi`, `--width`, `--height`, `--emit_pdf`, `--verbose`.
- `scripts/word_audit.py` — report heading level jumps, numbering without Heading styles, direct run-level formatting overrides, and font usage.
- `scripts/table_geometry.py` — write exact column widths to `tblW`, `tblGrid`, and every `tcW` so Word, LibreOffice, and Google Docs render the table identically.

Examples (each runnable as `python3 examples/<name>.py input.docx output.docx`):

- `examples/read_content.py` — dump paragraphs, tables, headers/footers, and section geometry.
- `examples/edit_content.py` — find-and-replace across runs and table cells, preserving style.
- `examples/edit_tables.py` — set content-derived column weights, cell margins, and visible borders on a table.
- `examples/tracked_changes.py` — add a tracked replacement via raw OOXML for cases `python-docx` cannot express.

## Final response

- Create/edit: cite the final `.docx` exactly once with a plain output citation and summarize representative changes.
- Q&A / read-only: cite the needed page(s) once each; do not re-export.
