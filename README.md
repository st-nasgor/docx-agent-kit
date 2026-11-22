# docx-agent-kit

One importable Python toolkit that fills the gaps in `python-docx`, packaged as
an agent skill for [OpenCode](https://opencode.ai) and Claude Code so coding
agents use it automatically.

Search terms this project answers:
`python-docx hyperlink` · `python-docx field PAGE DATE TOC` ·
`python-docx bullet numbered list numbering` · `python-docx table borders shading` ·
`python-docx merge cells` · `python-docx image placeholder replace` ·
`python-docx template {{placeholders}} repeat block` · `python-docx merge split runs` ·
`python-docx accept tracked changes strip comments` · `python-docx file looks like Word not python-docx` ·
`python-docx human metadata author` · `docx template renderer static site editor` · `docx to pdf LibreOffice`

It covers the things `python-docx` has no public API for: hyperlinks, Word
fields, real bullet/numbered lists, table borders and shading, image swaps,
template placeholders with repeatable blocks, tracked-change acceptance, comment
removal, and header/footer/section switches.

- Package: `docx_agent_kit`
- Skill ID: `docx-agent-kit`
- Python: 3.10+
- License: MIT (`LICENSE`)

## Install

The repository can live anywhere — a code folder, a temp directory, wherever
you keep things. The installer is what decides where the *skill* goes: it
detects OpenCode or Claude Code and copies the skill into the skills directory
that agent already reads, then pip-installs the package from that installed
copy.

| Target | Default directory |
|---|---|
| OpenCode (global) | `~/.config/opencode/skills` (`%USERPROFILE%\.config\opencode\skills` on Windows) |
| Claude Code (global) | `~/.claude/skills` |
| Project-local | pass `--skills-dir .opencode/skills` or `.claude/skills` |

You can force a target with `--agent opencode` / `--agent claude`, or point at
any directory with `--skills-dir`.

### For AI agents (copy this prompt)

```text
Install https://github.com/st-nasgor/docx-agent-kit for my agent.

1. Clone or download the repo anywhere convenient. Do not treat that folder
   as the install location.
2. Windows: run `powershell -ExecutionPolicy Bypass -File install.ps1`.
   macOS/Linux: run `./install.sh`.
   The installer detects OpenCode or Claude Code and installs the skill into
   that agent's skills directory (`~/.config/opencode/skills` or
   `~/.claude/skills`). Use `--skills-dir` for a project-local install.
3. Verify with `python -m docx_agent_kit inspect <some.docx>` and confirm the
   skill named docx-agent-kit is available in the agent.
```

### Full document stack (Claude's office skills)

The kit works alone, but it is designed to sit alongside Anthropic's document
skills (`docx`, `pdf`, `pptx`, `xlsx` from
[anthropics/skills](https://github.com/anthropics/skills)), which cover PDF
forms/merging, slide decks and spreadsheet recalculation. One flag installs
the whole stack:

```powershell
powershell -ExecutionPolicy Bypass -File install.ps1 -WithOfficeSkills
```

```bash
./install.sh --with-office-skills
```

This copies the four skills into the same agent skills directory, installs
their Python dependencies (`pypdf`, `pdfplumber`, `Pillow`, `python-pptx`,
`openpyxl`, `pandas`, `markitdown`), and applies two patches from
`patches/` in this repo:

- `soffice.py` — Anthropic ships a Linux-sandbox version (bare `soffice` on
  PATH, `LD_PRELOAD` socket shim needing gcc/`AF_UNIX`) that crashes on
  Windows; the patch locates `soffice.com` under Program Files and skips the
  shim.
- `recalc.py` — same hard-coded `soffice` fix, so xlsx formula recalculation
  works.

The patched files are full copies of the current upstream versions; re-run
the installer after an upstream update if the skills change significantly.
Requires `git` on PATH (for the anthropics/skills clone).

The installer replaces a previous copy and is safe to re-run. Use `--link` if
you want the skills directory to point at the repository instead of copying
(handy while developing the kit).

### Manual install

```bash
git clone https://github.com/st-nasgor/docx-agent-kit.git
cd docx-agent-kit
./install.sh                 # macOS / Linux
```

```powershell
powershell -ExecutionPolicy Bypass -File install.ps1   # Windows
```

OpenCode discovers `~/.config/opencode/skills` automatically and advertises the
skill by its `description`; no config entry is needed. Claude Code reads
`~/.claude/skills` the same way.

If you installed by copying and later pull changes, re-run the installer to
refresh the installed skill.

## Quick start

```python
import docx_agent_kit as dk
from docx import Document

doc = Document()

p = doc.add_paragraph("Project site: ")
dk.add_hyperlink(p, "https://example.com", "example.com")

dk.add_field(doc.add_paragraph("Page "), "PAGE", result="1")

bullets = dk.new_list(doc, "bullet")
dk.add_item(doc, "First point", bullets)
dk.add_item(doc, "Nested point", bullets, level=1)

table = doc.add_table(rows=2, cols=2)
dk.set_table_borders(table, color="BFBFBF")
dk.set_cell_shading(table.cell(0, 0), "1F3864")
dk.set_col_widths(table, [2.0, 4.0])

dk.update_fields(doc)
doc.save("out.docx")
```

## Feature map

| Need | Call |
|---|---|
| External link | `dk.add_hyperlink(paragraph, url, text)` |
| Internal link / bookmark | `dk.add_internal_hyperlink`, `dk.add_bookmark` |
| Page / page count | `dk.add_page_number`, `dk.add_num_pages` |
| Date field | `dk.add_date_field(paragraph, "MMMM d, yyyy")` |
| Any field | `dk.add_field(paragraph, "PAGE \\* MERGEFORMAT", result="1")` |
| Table of contents | `dk.add_toc(paragraph)` |
| Refresh fields on open | `dk.update_fields(document)` |
| Bullets / numbers | `dk.new_list(doc, "bullet")` + `dk.add_item` |
| Table borders | `dk.set_table_borders(table, color="BFBFBF")` |
| Cell fill | `dk.set_cell_shading(cell, "FFF2CC")` |
| Merge cells | `dk.merge_cells(table, 0, 0, 0, 1)` |
| Fixed widths | `dk.set_col_widths(table, [1.4, 2.6, 2.5])` |
| Repeat header row | `dk.set_repeat_header(table.rows[0])` |
| Clone row per item | `dk.repeat_table_row(table, row_index, items)` |
| Inline image | `dk.add_picture(paragraph, path, width=2.0)` |
| Swap image bytes | `dk.replace_inline_image(doc, 0, new_path)` |
| Merge split runs | `dk.merge_runs(doc)` |
| Accept revisions | `dk.accept_all_changes(doc)` |
| Remove comments | `dk.strip_comments(doc)` |
| First-page / even-odd headers | `dk.set_different_first_page`, `dk.set_even_odd_headers` |
| Restart page numbers | `dk.set_page_number_start(section, 1)` |
| New section (safe) | `sec = dk.new_section(doc)` — clears inherited titlePg |
| Footer page number | `dk.set_footer_page_number(section)` |
| Page number format | `dk.set_page_number_start(section, 1, fmt="lowerRoman")` |
| Paragraph starts a page | `dk.page_break_before(paragraph)` |

## Templates

Placeholders live in a normal `.docx`; the file itself stays the template, so
untouched formatting survives exactly.

```text
{{client_name}}                        text slot
{{#line_items}} ... {{/line_items}}    repeat once per dict in a list
{{image:hero}}                         paragraph becomes a picture
```

```python
dk.render("template.docx", {
    "client_name": "Acme",
    "hero": ("hero.png", 6.0),                 # path or (path, inches)
    "line_items": [{"desc": "A", "qty": 2}],
}, "out.docx")
```

Discover the slots as a JSON schema for a form or UI:

```python
dk.manifest_from_template("template.docx")
# {"template": "...", "fields": [{"key": ..., "type": ...}, ...]}

dk.inspect("file.docx")
# hyperlinks, fields, revisions, comments, numbering, headers, TOC flag
```

## CLI

```bash
python -m docx_agent_kit inspect report.docx
python -m docx_agent_kit manifest template.docx
python -m docx_agent_kit render template.docx payload.json -o out.docx
```

## Verify

```bash
python tests/selftest.py
```

The suite builds a document that exercises every helper, validates it against
the OOXML schema, converts it with LibreOffice, and checks the rendered PDF
text. It writes artifacts to the system temp directory.

For validation you can also use the `docx` OpenCode skill:

```bash
python ~/.config/opencode/skills/docx/scripts/office/validate.py out.docx
```

## Limits

- Word computes PAGE/TOC values when it opens the file; `update_fields`
  requests a refresh.
- Keep image swaps in the same format (PNG to PNG, JPEG to JPEG).
- Floating/anchored images, charts, SmartArt and equations are still raw XML;
  clone the original element instead of rebuilding it.
- The kit does not try to reproduce a document byte-for-byte. Keep the source
  `.docx` as the template and replace text/images, rather than regenerating
  layout from scratch.

## Layout

```text
docx_agent_kit/
  links.py      hyperlinks, bookmarks
  fields.py     PAGE, NUMPAGES, DATE, TOC, REF, updateFields
  lists.py      bullet and numbered numbering definitions
  tables.py     borders, shading, widths, merges, repeat headers
  images.py     add and replace inline images
  blocks.py     placeholders, repeat blocks, repeat table rows
  docops.py     merge runs, accept changes, strip comments, section flags
  template.py   render, manifest, inspect
  __main__.py   CLI
tests/selftest.py
patches/        Windows fixes for Anthropic's office skills (soffice.py, recalc.py)
install.ps1     Windows installer (add -WithOfficeSkills for the full stack)
install.sh      macOS/Linux installer (add --with-office-skills)
```
