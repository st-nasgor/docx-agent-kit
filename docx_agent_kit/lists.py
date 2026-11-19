"""Bullets and numbered lists without hand-editing numbering.xml.

The template's numbering part is reused when present; otherwise one is
created. Each call to :func:`new_list` allocates a fresh numId, so lists
created independently restart at 1.
"""

from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls, qn

from .oxml import insert_ordered, PPR_AFTER_NUMPR

BULLET_CHARS = ["\u2022", "\u25e6", "\u25aa"]
BULLET_FONT = "Arial"


def _numbering(document):
    try:
        return document.part.numbering_part
    except Exception:
        from docx.parts.numbering import NumberingPart

        part = NumberingPart.new()
        document.part.relate_to(part, RT.NUMBERING)
        return part


def _max_id(root, tag, attr):
    values = [int(node.get(qn(attr))) for node in root.findall(qn(tag))
              if (node.get(qn(attr)) or "").lstrip("-").isdigit()]
    return max(values) if values else -1


def _level_xml(kind, level, char=None):
    left = 720 * (level + 1)
    hanging = 360
    if kind == "bullet":
        fmt = "bullet"
        text = char or BULLET_CHARS[level % len(BULLET_CHARS)]
        rpr = (f'<w:rPr><w:rFonts w:ascii="{BULLET_FONT}" '
               f'w:hAnsi="{BULLET_FONT}" w:hint="default"/></w:rPr>')
    else:
        fmt = "decimal"
        text = f"%{level + 1}."
        rpr = ""
    return (
        f'<w:lvl w:ilvl="{level}">'
        f'<w:start w:val="1"/>'
        f'<w:numFmt w:val="{fmt}"/>'
        f'<w:lvlText w:val="{text}"/>'
        f'<w:lvlJc w:val="left"/>'
        f'<w:pPr><w:ind w:left="{left}" w:hanging="{hanging}"/></w:pPr>'
        f'{rpr}'
        f'</w:lvl>'
    )


def new_list(document, kind="bullet", levels=3, chars=None):
    """Create a bullet ('bullet') or numbered ('number') list definition.

    Returns the numId to pass to :func:`apply_list` / :func:`add_item`.
    """
    if kind not in ("bullet", "number"):
        raise ValueError("kind must be 'bullet' or 'number'")
    numbering = _numbering(document)
    root = numbering.element
    abstract_id = _max_id(root, "w:abstractNum", "w:abstractNumId") + 1
    num_id = _max_id(root, "w:num", "w:numId") + 1

    lvls = "".join(
        _level_xml(kind, level, chars[level] if chars else None)
        for level in range(levels)
    )
    abstract = parse_xml(
        f'<w:abstractNum {nsdecls("w")} w:abstractNumId="{abstract_id}">'
        f'<w:multiLevelType w:val="hybridMultilevel"/>{lvls}</w:abstractNum>'
    )
    first_num = root.find(qn("w:num"))
    if first_num is not None:
        first_num.addprevious(abstract)
    else:
        root.append(abstract)

    num = parse_xml(
        f'<w:num {nsdecls("w")} w:numId="{num_id}">'
        f'<w:abstractNumId w:val="{abstract_id}"/></w:num>'
    )
    root.append(num)
    return num_id


def apply_list(paragraph, num_id, level=0):
    """Attach an existing list definition to a paragraph."""
    pPr = paragraph._p.get_or_add_pPr()
    numPr = parse_xml(
        f'<w:numPr {nsdecls("w")}>'
        f'<w:ilvl w:val="{level}"/><w:numId w:val="{num_id}"/></w:numPr>'
    )
    insert_ordered(pPr, numPr, PPR_AFTER_NUMPR)
    return paragraph


def _list_style(document):
    names = {style.name for style in document.styles}
    return "List Paragraph" if "List Paragraph" in names else None


def add_item(container, text, num_id, level=0, style=None):
    """Add a paragraph that belongs to *num_id* at *level*."""
    document = container if hasattr(container, "styles") else container.part.document
    style = style or _list_style(document)
    paragraph = container.add_paragraph(text, style=style) if style else container.add_paragraph(text)
    return apply_list(paragraph, num_id, level)


def bullet_list(document, items, kind="bullet", levels=3):
    """Convenience: create a list and fill it with strings or (level, text)."""
    num_id = new_list(document, kind=kind, levels=levels)
    paragraphs = []
    for item in items:
        level, text = item if isinstance(item, tuple) else (0, item)
        paragraphs.append(add_item(document, text, num_id, level))
    return num_id, paragraphs
