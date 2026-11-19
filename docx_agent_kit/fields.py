"""Word fields: PAGE, NUMPAGES, DATE, TOC, REF and arbitrary instructions."""

from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from .oxml import insert_ordered, SETTINGS_AFTER_UPDATEFIELDS


def _field_char(kind):
    node = OxmlElement("w:fldChar")
    node.set(qn("w:fldCharType"), kind)
    return node


def _instr_text(instruction):
    node = OxmlElement("w:instrText")
    node.set(qn("xml:space"), "preserve")
    node.text = f" {instruction.strip()} "
    return node


def add_field(paragraph, instruction, result="", font=None, size=None,
              bold=None, italic=None, color=None):
    """Insert a complete field (begin / instrText / separate / result / end).

    Word or LibreOffice must recalculate the value; call ``update_fields``
    on the document so Word refreshes on open.
    """
    begin_run = paragraph.add_run()
    begin_run._r.append(_field_char("begin"))

    instr_run = paragraph.add_run()
    instr_run._r.append(_instr_text(instruction))

    sep_run = paragraph.add_run()
    sep_run._r.append(_field_char("separate"))

    result_run = paragraph.add_run(result)
    if font:
        result_run.font.name = font
    if size:
        result_run.font.size = size
    if bold is not None:
        result_run.bold = bold
    if italic is not None:
        result_run.italic = italic
    if color is not None:
        result_run.font.color.rgb = color

    end_run = paragraph.add_run()
    end_run._r.append(_field_char("end"))
    return result_run


def add_page_number(paragraph, result="1", **kwargs):
    return add_field(paragraph, "PAGE", result, **kwargs)


def add_num_pages(paragraph, result="1", **kwargs):
    return add_field(paragraph, "NUMPAGES", result, **kwargs)


def add_date_field(paragraph, format_code="MMMM d, yyyy", result="", **kwargs):
    return add_field(paragraph, f'DATE \\@ "{format_code}"', result, **kwargs)


def add_toc(paragraph, levels="1-3", result="Right-click and choose Update Field"):
    instruction = f'TOC \\o "{levels}" \\h \\z \\u'
    return add_field(paragraph, instruction, result)


def add_ref(paragraph, bookmark, result="", **kwargs):
    return add_field(paragraph, f"REF {bookmark} \\h", result, **kwargs)


def update_fields(document, value=True):
    """Set <w:updateFields> so Word refreshes fields when the file opens."""
    from .docops import repair_settings

    repair_settings(document)
    settings = document.settings.element
    existing = settings.findall(qn("w:updateFields"))
    for node in existing:
        node.set(qn("w:val"), "true" if value else "false")
        return
    node = OxmlElement("w:updateFields")
    node.set(qn("w:val"), "true" if value else "false")
    insert_ordered(settings, node, SETTINGS_AFTER_UPDATEFIELDS)


def count_fields(container):
    return len(container.findall(".//" + qn("w:fldChar")))


def field_instructions(container):
    return [node.text.strip() for node in container.iter(qn("w:instrText"))
            if node.text]
