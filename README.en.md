# DSH Word Skill

A [DSH](https://github.com/deepseek-ai/dsh) skill that unifies two capabilities into a single entrypoint: **reading** Word documents and **precisely editing** their content and formatting.

## Capabilities

- **Read** a `.docx`: extract and return structured text (summaries, answers, targeted lookups).
- **Edit precisely**: change wording, restructure sections, adjust styles/headings, fix tables, spacing, headers, footers, and other layout details, then verify the result visually.
- **Built-in scripts and examples**: ships generic, self-contained Python tooling for render verification, auditing, and precise table editing.

## Structure

```text
dsh-word-skill/
|-- .dsh/
|   `-- skills/
|       `-- word/
|           |-- SKILL.md                Entry point: scope, tool contract, golden path
|           |-- README.md               Skill overview
|           |-- scripts/
|           |   |-- render_docx.py      Generic DOCX renderer (LibreOffice headless, page-N.png)
|           |   |-- word_audit.py       Generic document audit (heading hierarchy, numbering, overrides)
|           |   `-- table_geometry.py   Generic table geometry (tblW/tblGrid/tcW exact widths)
|           |-- examples/
|           |   |-- read_content.py     Dump paragraphs, tables, headers/footers, section geometry
|           |   |-- edit_content.py     Precise find-and-replace across runs and table cells
|           |   |-- edit_tables.py      Set content-derived column weights, cell margins, borders
|           |   `-- tracked_changes.py  Add a tracked replacement via raw OOXML patch
|           `-- references/
|               `-- format-precision.md Formatting-accuracy checklist (spacing, borders, alignment, tables)
```

## Installation

Place `.dsh/skills/word/` under `.dsh/skills/` in your project root (the directory containing `.git`). DSH discovers `SKILL.md` automatically — no restart required.

## How It Works

This skill is an orchestrator. It does not duplicate lower-level machinery; instead, the agent does the work through the `bash` tool by running Python directly:

- **Reading**: `python-docx` extracts paragraphs, tables, headers/footers, and other structured content.
- **Precise editing**: `python-docx` handles routine paragraphs/styles/tables, and raw OOXML patches handle tracked changes, comments, hyperlinks, fields, and other advanced constructs.
- **Render verification**: LibreOffice renders the document to PNG for page-by-page review. DSH ships no built-in document reader, so prefer the model's own vision or an image-input tool when looking at the PNGs; text-only models should route through a vision/OCR skill (e.g. `see`); with neither available, fall back to text-only auditing and say so plainly.

## Usage

Describe your Word task in natural language inside DSH, for example:

- "Read this docx and summarize it"
- "Make the chapter 3 headings bold"
- "Fix the column widths and borders in this table"

The skill activates automatically and follows the golden path: read for context, author the edit, render to PNG and inspect every page, then iterate until the document is flawless. Only the final `.docx` is delivered.

Scripts and examples run directly in a Python environment (install dependencies first):

```bash
pip install python-docx lxml pdf2image

# Render verification
python3 .dsh/skills/word/scripts/render_docx.py input.docx --output_dir out/

# Pre-edit audit
python3 .dsh/skills/word/scripts/word_audit.py input.docx

# Precise table editing
python3 .dsh/skills/word/examples/edit_tables.py input.docx output.docx
```

Rendering requires [LibreOffice](https://www.libreoffice.org/) (or set `SOFFICE_BIN`).

## Scope

- Intended for `.docx` content reading and precise editing.
- Not for creating spreadsheets, slides, or PDFs.
- Not for hand-drawn or illustrative artifacts.

## License

[MIT](./LICENSE)
