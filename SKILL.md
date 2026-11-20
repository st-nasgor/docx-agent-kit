---
name: docx-agent-kit
description: >
  Reusable python-docx toolkit for Word documents. Use whenever a task needs
  hyperlinks, Word fields (PAGE, NUMPAGES, DATE, TOC, REF), real bullet or
  numbered lists, table borders/shading/merges, image replacement, template
  placeholders, repeatable blocks, tracked-change acceptance, comment
  removal, or header/footer/section switches. Prefer this over hand-writing
  OOXML; it is installed in the active Python as `docx_agent_kit`.
---

# docx-agent-kit

A single importable pack that fills python-docx's API gaps. Installed
editable from this directory, so edits here take effect immediately.

```python
import docx_agent_kit as dk
from docx import Document
```

## When to use

Any `.docx` work in Python. `python-docx` alone cannot do hyperlinks,
fields, numbering, merged borders, image swaps, or template loops; docx_agent_kit
wraps those as functions. Use the `docx` skill for its validator and
LibreOffice helpers, and this pack for construction.

## API map

| Need | Call |
|---|---|
| External link | `dk.add_hyperlink(paragraph, url, text, font=, size=)` |
| Internal link / bookmark | `dk.add_internal_hyperlink(p, anchor, text)`, `dk.add_bookmark(p, name, id)` |
| Page / page count | `dk.add_page_number(paragraph)`, `dk.add_num_pages(paragraph)` |
| Date field | `dk.add_date_field(paragraph, "MMMM d, yyyy")` |
| Any field | `dk.add_field(paragraph, "PAGE \\* MERGEFORMAT", result="1")` |
| TOC | `dk.add_toc(paragraph, levels="1-3")` |
| Refresh on open | `dk.update_fields(document)` (also repairs the zoom bug) |
| New section (safe) | `sec = dk.new_section(doc)` — clears inherited titlePg |
| Footer page number | `dk.set_footer_page_number(section)` |
| Page number restart / format | `dk.set_page_number_start(section, 1, fmt="lowerRoman")` |
| Paragraph starts a page | `dk.page_break_before(paragraph)` |
| Bullets / numbers | `num = dk.new_list(doc, "bullet"); dk.add_item(doc, "Text", num, level=0)` |
| Table borders | `dk.set_table_borders(table, color="BFBFBF")` |
| Cell fill | `dk.set_cell_shading(cell, "FFF2CC")` |
| Merge | `dk.merge_cells(table, 0, 0, 0, 1)` |
| Fixed widths | `dk.set_col_widths(table, [1.4, 2.6, 2.5])` |
| Repeat header row | `dk.set_repeat_header(table.rows[0])` |
| Clone a row per item | `dk.repeat_table_row(table, row_index, items)` |
| Inline image | `dk.add_picture(paragraph, path, width=2.0)` |
| Swap image bytes | `dk.replace_inline_image(doc, 0, new_path)` |
| Merge split runs | `dk.merge_runs(doc)` |
| Accept revisions | `dk.accept_all_changes(doc)` |
| Remove comments | `dk.strip_comments(doc)` |
| First-page / even-odd headers | `dk.set_different_first_page(section)`, `dk.set_even_odd_headers(doc)` |
| Restart page numbers | `dk.set_page_number_start(section, 1)` |

## Templates

Conventions recognised by `dk.render`:

```text
{{client_name}}                 text slot
{{#line_items}} ... {{/line_items}}   repeat once per item (dict)
{{image:hero}}                  becomes a picture
```

```python
dk.render("template.docx", {
    "client_name": "Acme",
    "hero": ("hero.png", 6.0),            # path or (path, inches)
    "line_items": [{"desc": "A", "qty": 2}, {"desc": "B", "qty": 5}],
}, "out.docx")
```

Discover slots and build a form schema:

```python
dk.manifest_from_template("template.docx")
# {"template": ..., "fields": [{"key", "type", "block", "fields"}]}

dk.inspect("file.docx")
# hyperlinks, fields, revisions, comments, numbering, headers, TOC flag
```

## CLI

```bash
python -m docx_agent_kit inspect report.docx
python -m docx_agent_kit manifest template.docx
python -m docx_agent_kit render template.docx payload.json -o out.docx
```

## Rules of thumb

1. Prefer editing a copy of a real `.docx` template over rebuilding a
   document from scratch. Unchanged parts keep exact formatting.
2. Placeholders may be split across runs; `replace_placeholders` handles
   that. Run `dk.merge_runs` first when inspecting text.
3. Word computes PAGE/TOC fields on open. `update_fields` requests a
   refresh; a PDF preview via LibreOffice also resolves them.
4. Keep image replacements in the same format (PNGâ†’PNG, JPEGâ†’JPEG).
5. Anchored/floating images, charts and SmartArt are still raw XML; clone
   the original element rather than recreating it.

## Verify

```bash
python tests/selftest.py          # 35 feature checks
python "$HOME/.config/opencode/skills/docx/scripts/office/validate.py" out.docx
```

The self test writes to `%LOCALAPPDATA%\Temp\opencode\docx-agent-kit-test`.
