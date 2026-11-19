"""docx_agent_kit: one importable toolbox for the python-docx feature gaps.

    import docx_agent_kit as dk
    from docx import Document

    doc = Document("template.docx")
    dk.merge_runs(doc)
    dk.replace_placeholders(doc, {"client": "Acme"})
    dk.add_hyperlink(doc.paragraphs[0], "https://example.com", "site")
    dk.add_field(doc.paragraphs[1], "PAGE")
    bullets = dk.new_list(doc, "bullet")
    dk.add_item(doc, "First point", bullets)
    dk.update_fields(doc)
    doc.save("out.docx")
"""

from .blocks import (
    fill_blocks,
    fill_image_placeholders,
    find_tokens,
    replace_placeholders,
    replace_text,
    repeat_table_row,
)
from .docops import (
    accept_all_changes,
    merge_runs,
    repair_settings,
    set_different_first_page,
    set_even_odd_headers,
    set_page_number_start,
    strip_comments,
    strip_comments_and_changes,
)
from .fields import (
    add_date_field,
    add_field,
    add_num_pages,
    add_page_number,
    add_ref,
    add_toc,
    field_instructions,
    update_fields,
)
from .images import add_picture, list_images, replace_image_by_partname, replace_inline_image
from .links import add_bookmark, add_hyperlink, add_internal_hyperlink, count_hyperlinks
from .lists import add_item, apply_list, bullet_list, new_list
from .tables import (
    merge_cells,
    set_cant_split,
    set_cell_borders,
    set_cell_margins,
    set_cell_shading,
    set_col_widths,
    set_repeat_header,
    set_table_borders,
)
from .template import inspect, manifest_from_template, render

__all__ = [
    # placeholders / blocks
    "replace_text", "replace_placeholders", "fill_blocks", "repeat_table_row",
    "fill_image_placeholders", "find_tokens",
    # links
    "add_hyperlink", "add_internal_hyperlink", "add_bookmark", "count_hyperlinks",
    # fields
    "add_field", "add_page_number", "add_num_pages", "add_date_field",
    "add_toc", "add_ref", "update_fields", "field_instructions",
    # lists
    "new_list", "apply_list", "add_item", "bullet_list",
    # tables
    "set_table_borders", "set_cell_borders", "set_cell_shading",
    "set_col_widths", "merge_cells", "set_repeat_header", "set_cant_split",
    "set_cell_margins",
    # images
    "add_picture", "list_images", "replace_inline_image", "replace_image_by_partname",
    # document ops
    "merge_runs", "accept_all_changes", "strip_comments",
    "strip_comments_and_changes", "set_different_first_page",
    "set_even_odd_headers", "set_page_number_start", "repair_settings",
    # template
    "render", "manifest_from_template", "inspect",
]

__version__ = "1.0.0"
