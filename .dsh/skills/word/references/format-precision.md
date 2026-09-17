# Format-precision checklist

A formatting edit is "precise" only when the rendered document matches intent on every page. Use these checks — drawn from the render-verify discipline — before you declare an edit done.

## Render first, judge second

- Run `render_docx.py` → `page-<N>.png`, inspect **every** page at 100% zoom. Text extraction and XML inspection miss clipping, overlap, glyph drops, and spacing drift. How you look at the PNGs depends on the session: use an image-input tool if the model has one, else a vision/OCR skill (e.g. `see`), else fall back to text-only QA — see *Inspecting rendered pages* in `SKILL.md`.
- LibreOffice stderr noise (e.g. `Unknown IO error`) is common even on success. Trust the PNGs, not the logs.
- After **any** OOXML patch or layout-sensitive change, re-render and re-inspect.

## Common invisible defects

| Defect | What to check in the PNG |
| --- | --- |
| Cell text pinned to top/left | Generous, even internal padding on all four sides |
| Descenders/ascenders clipped | Line spacing and cell height must let wrapped content expand; never fix row heights |
| Overfull right edge | Wrap text, widen columns, or (last) reduce type slightly — in that order |
| Border inconsistency | Explicit outer/internal cell borders (e.g. light gray `#D9D9D9`); no inherited theme gaps |
| Run vs. paragraph spacing drift | Space sits on paragraph properties, not stray run-level spacing or blank "spacer" paragraphs |
| Table pushed to next page leaving a gap | Move preceding prose with it, scale the visual, or split with repeated headers |
| Tracked changes not visible | PNG review shows them; for comments, also inspect `comments.xml` + anchors + rels + content-types (headless export often drops comments) |

## Defaults that age badly

- Equal-width columns when content widths differ — set widths intentionally.
- All body cells defaulting to top-left — center short values/dates/status, left-align narrative.
- Header fill forced to one color — pick light/dark gray/blue per document, keep related tables consistent, contrast text color with fill.
- Direct overrides of the `Title` style or paragraph borders under titles — keep titles black, no underlines, no decorative lines.

## Tight-content priority order

When a cell, line, or box feels cramped: **wrap → adjust width/padding → reduce font slightly → abbreviate or use a two-line header**. Never start by shrinking type.

## Deliverables

- The requested final `.docx` only. Rendered PNGs and any optional PDF are internal QA — do not deliver unless asked.
- Do not leak tool decision text ("internal working draft", source caveats) into the document; flag those to the user in prose instead.
