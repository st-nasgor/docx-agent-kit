# docx-agent-kit

**The Word features `python-docx` left out, as plain functions.**

`python-docx` writes paragraphs and simple tables. Hyperlinks, page numbers, real numbered lists, bordered tables: each one means editing OOXML by hand. This package wraps those jobs in one importable toolkit and ships as an agent skill, so OpenCode and Claude Code use it instead of writing XML.

```python
import docx_agent_kit as dk
```

Python 3.10+ · MIT · Dependencies: `python-docx`, `lxml`

## What it adds

- **Links**: external hyperlinks, internal links to bookmarks, bookmarks
- **Fields**: `PAGE`, `NUMPAGES`, `DATE`, `TOC`, `REF`, and a switch that asks Word to refresh them on open
- **Lists**: bullet and numbered lists with numbering definitions and nesting levels
- **Tables**: borders for a table or one cell, shading, merges, fixed column widths, repeating header rows, cell margins
- **Images**: insert inline pictures, replace a picture in place, list what is embedded
- **Templates**: fill `{{slots}}`, repeat `{{#blocks}}`, swap `{{image:name}}`, all in a real `.docx` so formatting you did not touch survives
- **Clean-up**: accept tracked changes, strip comments, merge runs that Word split mid-sentence
- **Sections**: first-page and odd/even headers, page-number restarts and formats, new sections without inherited flags
- **Metadata**: the file stops reporting `python-docx` as the author

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

## Install

```powershell
# Windows
powershell -ExecutionPolicy Bypass -File install.ps1
```

```bash
# macOS / Linux
./install.sh
```

The installer detects OpenCode or Claude Code, copies the skill into the skills directory that agent reads, and installs the package into the active Python. Start a new agent session and the agent finds the skill.

Full flags and manual installs: [Install reference](#install-reference).

## Templates

Placeholders live in a normal `.docx`, and the file stays the template, so formatting you did not touch survives.

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

`render` repairs the settings part, refreshes fields on open, and cleans the document metadata (see [Metadata](#metadata)). With no `out.docx` it writes `<name>-filled.docx` beside the template. `strict=True` raises if a `{{token}}` is left unfilled. Pass `accept_changes=True` or `remove_comments=True` to clean a draft on the way through.

`dk.manifest_from_template` returns the slot names as a JSON schema, ready for a form or a UI:

```python
dk.manifest_from_template("template.docx")
# {"template": "...", "fields": [{"key": ..., "type": ...}, ...]}

dk.inspect("file.docx")
# paragraphs, tables, hyperlinks, fields, revisions, comments, numbering, headers, TOC flag
```

## Command line

The package runs as a module and installs a `docx-agent-kit` script:

```bash
python -m docx_agent_kit inspect report.docx
python -m docx_agent_kit manifest template.docx
python -m docx_agent_kit render template.docx payload.json -o out.docx
```

`render` takes `-` as the payload path to read JSON from stdin. All three commands print JSON.

## Feature map

### Links

| Need | Call |
|---|---|
| External link | `dk.add_hyperlink(paragraph, url, text)` |
| Link to a bookmark | `dk.add_internal_hyperlink(paragraph, anchor, text)` |
| Create a bookmark | `dk.add_bookmark(paragraph, name, bookmark_id, text=None)` |
| Count links | `dk.count_hyperlinks(document)` |

### Fields

| Need | Call |
|---|---|
| Any field | `dk.add_field(paragraph, "PAGE \\* MERGEFORMAT", result="1")` |
| Page number | `dk.add_page_number(paragraph)` |
| Page count | `dk.add_num_pages(paragraph)` |
| Date | `dk.add_date_field(paragraph, "MMMM d, yyyy")` |
| Table of contents | `dk.add_toc(paragraph, levels="1-3")` |
| Reference to a bookmark | `dk.add_ref(paragraph, bookmark)` |
| Refresh fields on open | `dk.update_fields(document)` |
| Read field codes | `dk.field_instructions(document)` |

### Lists

| Need | Call |
|---|---|
| New list definition | `dk.new_list(document, "bullet")` or `"number"` |
| Add an item | `dk.add_item(container, "Text", num_id, level=0)` |
| Build a list inline | `dk.bullet_list(document, ["One", "Two"])` |
| Apply an existing list | `dk.apply_list(paragraph, num_id, level=0)` |

### Tables

| Need | Call |
|---|---|
| Table borders | `dk.set_table_borders(table, color="BFBFBF")` |
| Per-cell borders | `dk.set_cell_borders(cell, color="BFBFBF", edges=...)` |
| Cell fill | `dk.set_cell_shading(cell, "FFF2CC")` |
| Merge cells | `dk.merge_cells(table, 0, 0, 0, 1)` |
| Fixed widths | `dk.set_col_widths(table, [1.4, 2.6, 2.5])` |
| Repeat header row | `dk.set_repeat_header(table.rows[0])` |
| Keep row on one page | `dk.set_cant_split(table.rows[0])` |
| Cell padding | `dk.set_cell_margins(table, top=40, start=80, bottom=40, end=80)` |
| Clone a row per item | `dk.repeat_table_row(table, row_index, items)` |

### Images

| Need | Call |
|---|---|
| Inline image | `dk.add_picture(paragraph, path, width=2.0)` |
| Swap image bytes | `dk.replace_inline_image(document, 0, new_path)` |
| Replace by part name | `dk.replace_image_by_partname(document, "image1.png", path)` |
| List embedded images | `dk.list_images(document)` |

### Clean-up

| Need | Call |
|---|---|
| Merge split runs | `dk.merge_runs(document)` |
| Accept revisions | `dk.accept_all_changes(document)` |
| Remove comments | `dk.strip_comments(document)` |
| Do both | `dk.strip_comments_and_changes(document)` |
| Clean metadata | `dk.humanize(document, author=None)` |
| Fix settings defects | `dk.repair_settings(document)` |

### Sections and headers

| Need | Call |
|---|---|
| New section, no inherited flags | `sec = dk.new_section(document)` |
| First-page header | `dk.set_different_first_page(section)` |
| Odd/even headers | `dk.set_even_odd_headers(document)` |
| Restart / reformat page numbers | `dk.set_page_number_start(section, 1, fmt="lowerRoman")` |
| Footer page number | `dk.set_footer_page_number(section)` |
| Paragraph starts a page | `dk.page_break_before(paragraph)` |

### Placeholders and blocks

| Need | Call |
|---|---|
| Fill `{{slot}}` tokens | `dk.replace_placeholders(document, {"client": "Acme"})` |
| Fill repeat blocks | `dk.fill_blocks(document, {"items": [...]})` |
| Replace `{{image:hero}}` | `dk.fill_image_placeholders(document, {"hero": path})` |
| List tokens, no filling | `dk.find_tokens(document)` |
| Replace text in a subtree | `dk.replace_text(root, {"a": "b"})` |

### Templates and introspection

| Need | Call |
|---|---|
| Fill a template | `dk.render(template, payload, out=None, strict=False, update=True, ...)` |
| Field schema from a template | `dk.manifest_from_template(path)` |
| Report features used by a file | `dk.inspect(path)` |

## Metadata

`python-docx` stamps every document it saves: author `python-docx`, comment `generated by python-docx`, a 2013 creation date, revision 1. `dk.humanize()` overwrites those, sets the author to the current user (or the name you pass), resets timestamps, and removes `keywords`, `description`, `category`, `contentStatus` and `lastPrinted`.

`dk.render()` calls it for you, so rendered files do not name the tool that made them. Call `dk.humanize(doc, author="Your Name")` on documents you build by hand.

## Verify

`python tests/selftest.py` builds a document that exercises every helper, validates it against the OOXML schema, converts it with LibreOffice, and checks the text of the rendered PDF. Artifacts go to the system temp directory.

With Anthropic's `docx` skill installed, you can validate one file at a time:

```bash
python ~/.config/opencode/skills/docx/scripts/office/validate.py out.docx
```

## Limits

- Word computes `PAGE` and `TOC` values when it opens the file. `update_fields` asks for the refresh; a LibreOffice PDF conversion resolves them too.
- Keep image swaps in the same format (PNG to PNG, JPEG to JPEG).
- Floating and anchored images, charts, SmartArt and equations stay raw XML. Clone the original element instead of rebuilding it.
- The kit does not reproduce a document byte-for-byte. Keep the source `.docx` as the template and replace text or images instead of regenerating layout.

## Full document stack (optional)

The kit works alone, but pairs with Anthropic's document skills (`docx`, `pdf`, `pptx`, `xlsx` from [anthropics/skills](https://github.com/anthropics/skills)), which handle PDF forms, slide decks and spreadsheet recalculation. One flag installs the whole stack:

```powershell
powershell -ExecutionPolicy Bypass -File install.ps1 -WithOfficeSkills
```

```bash
./install.sh --with-office-skills
```

This copies the four skills into the same directory, installs their Python dependencies (`pypdf`, `pdfplumber`, `Pillow`, `python-pptx`, `openpyxl`, `pandas`, `markitdown`), and applies two patches from `patches/`:

- `soffice.py`: Anthropic ships a Linux-sandbox version (bare `soffice` on PATH, an `LD_PRELOAD` socket shim that needs gcc and `AF_UNIX`) and it crashes on Windows. The patch finds `soffice.com` under Program Files and skips the shim.
- `recalc.py`: the same fix for the hard-coded `soffice` path, so xlsx formula recalculation works.

Patched files are full copies of the current upstream versions, so re-run the installer after an upstream update. Needs `git` on PATH to clone anthropics/skills.

## Install reference

The repository can live anywhere. The installer decides where the skill goes: it detects OpenCode or Claude Code, copies the skill into the skills directory that agent reads, and pip-installs the package from that copy.

| Target | Default directory |
|---|---|
| OpenCode (global) | `~/.config/opencode/skills` (`%USERPROFILE%\.config\opencode\skills` on Windows) |
| Claude Code (global) | `~/.claude/skills` |
| Project-local | pass `--skills-dir .opencode/skills` or `.claude/skills` |

| Flag | Effect |
|---|---|
| `--agent auto\|opencode\|claude` | Force a target instead of detecting one (`-Agent` on Windows) |
| `--skills-dir PATH` | Install into any skills directory, e.g. a project-local one |
| `--link` | Symlink the repository instead of copying (useful while developing the kit) |
| `--with-office-skills` | Also install Anthropic's four document skills (`-WithOfficeSkills`) |

The installer replaces an earlier copy and is safe to re-run. If you installed by copying and pull changes later, re-run it.

### Manual install

```bash
git clone https://github.com/st-nasgor/docx-agent-kit.git
cd docx-agent-kit
./install.sh                 # macOS / Linux
```

```powershell
powershell -ExecutionPolicy Bypass -File install.ps1   # Windows
```

OpenCode reads `~/.config/opencode/skills` and Claude Code reads `~/.claude/skills`, both without a config entry, because each agent advertises a skill by its `description`.

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

## Project layout

```text
docx_agent_kit/
  links.py      hyperlinks, bookmarks
  fields.py     PAGE, NUMPAGES, DATE, TOC, REF, updateFields
  lists.py      bullet and numbered numbering definitions
  tables.py     borders, shading, widths, merges, repeat headers
  images.py     add and replace inline images
  blocks.py     placeholders, repeat blocks, repeat table rows
  docops.py     merge runs, accept changes, strip comments, section flags, metadata
  template.py   render, manifest, inspect
  __main__.py   CLI
tests/selftest.py
patches/        Windows fixes for Anthropic's office skills (soffice.py, recalc.py)
install.ps1     Windows installer (add -WithOfficeSkills for the full stack)
install.sh      macOS/Linux installer (add --with-office-skills)
```

## Search keywords

`python-docx hyperlink` · `python-docx field PAGE DATE TOC` ·
`python-docx bullet numbered list numbering` · `python-docx table borders shading` ·
`python-docx merge cells` · `python-docx image placeholder replace` ·
`python-docx template {{placeholders}} repeat block` · `python-docx merge split runs` ·
`python-docx accept tracked changes strip comments` · `python-docx file looks like Word not python-docx` ·
`python-docx human metadata author` · `docx template renderer static site editor` · `docx to pdf LibreOffice`

## License

MIT. See `LICENSE`.
