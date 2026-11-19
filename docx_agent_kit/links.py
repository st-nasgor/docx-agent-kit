"""Hyperlinks and bookmarks that python-docx does not expose publicly."""

from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from .oxml import make, insert_ordered


def _run_props(color=None, underline=True, bold=None, italic=None, font=None,
               size=None, strike=None):
    rPr = OxmlElement("w:rPr")
    if font:
        rFonts = make("w:rFonts", **{
            "w:ascii": font, "w:hAnsi": font, "w:cs": font, "w:eastAsia": font,
        })
        rPr.append(rFonts)
    if bold:
        rPr.append(OxmlElement("w:b"))
    if italic:
        rPr.append(OxmlElement("w:i"))
    if strike:
        rPr.append(OxmlElement("w:strike"))
    if color:
        rPr.append(make("w:color", **{"w:val": color.lstrip("#")}))
    if size:
        half_points = int(round(float(size) * 2))
        rPr.append(make("w:sz", **{"w:val": half_points}))
        rPr.append(make("w:szCs", **{"w:val": half_points}))
    if underline:
        value = "single" if underline is True else underline
        rPr.append(make("w:u", **{"w:val": value}))
    return rPr


def add_hyperlink(paragraph, url, text, color="0563C1", underline=True,
                  bold=None, italic=None, font=None, size=None):
    """Append an external hyperlink run to *paragraph*.

    python-docx has no public hyperlink API, so this builds the
    <w:hyperlink> element and its external relationship by hand.
    """
    r_id = paragraph.part.relate_to(url, RT.HYPERLINK, is_external=True)
    hyperlink = make("w:hyperlink", **{"r:id": r_id})
    run = OxmlElement("w:r")
    run.append(_run_props(color, underline, bold, italic, font, size))
    node = OxmlElement("w:t")
    node.set(qn("xml:space"), "preserve")
    node.text = text
    run.append(node)
    hyperlink.append(run)
    paragraph._p.append(hyperlink)
    return hyperlink


def add_bookmark(paragraph, name, bookmark_id, text=None):
    """Wrap the paragraph's existing text (or append *text*) in a bookmark."""
    start = make("w:bookmarkStart", **{"w:id": bookmark_id, "w:name": name})
    end = make("w:bookmarkEnd", **{"w:id": bookmark_id})
    paragraph._p.insert(0, start)
    if text is not None:
        paragraph.add_run(text)
    paragraph._p.append(end)
    return bookmark_id


def add_internal_hyperlink(paragraph, anchor, text, color="0563C1",
                           underline=True, bold=None, italic=None, font=None,
                           size=None):
    """Link to an existing bookmark name inside the same document."""
    hyperlink = make("w:hyperlink", **{"w:anchor": anchor})
    run = OxmlElement("w:r")
    run.append(_run_props(color, underline, bold, italic, font, size))
    node = OxmlElement("w:t")
    node.set(qn("xml:space"), "preserve")
    node.text = text
    run.append(node)
    hyperlink.append(run)
    paragraph._p.append(hyperlink)
    return hyperlink


def count_hyperlinks(container):
    return len(container.findall(".//" + qn("w:hyperlink")))
