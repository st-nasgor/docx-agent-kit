"""Document clean-up and section switches.

Covers the things the docx skill says to handle by rezipping XML:
merge split runs, accept tracked changes, strip comments, and set
section-level flags.
"""

from lxml import etree
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from .oxml import (
    make,
    remove,
    insert_ordered,
    SETTINGS_AFTER_EVENODD,
    SECTPR_TITLE_PAGE,
)

CHANGE_TAGS = (
    "w:rPrChange", "w:pPrChange", "w:tblPrChange", "w:trPrChange",
    "w:tcPrChange", "w:sectPrChange", "w:numberingChange",
)


def _unwrap(element):
    parent = element.getparent()
    if parent is None:
        return
    index = parent.index(element)
    for child in list(element):
        parent.insert(index, child)
        index += 1
    parent.remove(element)


def _text_only(run):
    allowed = {qn("w:rPr"), qn("w:t")}
    return all(child.tag in allowed for child in run)


def _rpr_key(run):
    rPr = run.find(qn("w:rPr"))
    return b"" if rPr is None else etree.tostring(rPr)


def merge_runs(document):
    """Coalesce adjacent identically-formatted runs so text is findable."""
    merged = 0
    for paragraph in document.element.body.iter(qn("w:p")):
        previous = None
        for run in list(paragraph.findall(qn("w:r"))):
            if not _text_only(run):
                previous = None
                continue
            if previous is not None and _rpr_key(previous) == _rpr_key(run):
                target = previous.find(qn("w:t"))
                source = run.find(qn("w:t"))
                if source is not None:
                    if target is None:
                        target = OxmlElement("w:t")
                        previous.append(target)
                    target.text = (target.text or "") + (source.text or "")
                    target.set(qn("xml:space"), "preserve")
                paragraph.remove(run)
                merged += 1
            else:
                previous = run
    return merged


def accept_all_changes(document):
    """Accept insertions and drop deletions, moves and formatting revisions."""
    root = document.element.body
    for tag in ("w:ins", "w:moveTo"):
        for element in list(root.iter(qn(tag))):
            _unwrap(element)
    for tag in ("w:del", "w:moveFrom"):
        for element in list(root.iter(qn(tag))):
            remove(element)
    for tag in ("w:moveFromRangeStart", "w:moveFromRangeEnd",
                "w:moveToRangeStart", "w:moveToRangeEnd"):
        for element in list(root.iter(qn(tag))):
            remove(element)
    for tag in CHANGE_TAGS:
        for element in list(root.iter(qn(tag))):
            remove(element)


def strip_comments(document):
    """Delete comment anchors and the comment parts' relationships."""
    root = document.element.body
    for tag in ("w:commentRangeStart", "w:commentRangeEnd"):
        for element in list(root.iter(qn(tag))):
            remove(element)
    for reference in list(root.iter(qn("w:commentReference"))):
        parent = reference.getparent()
        remove(parent if parent is not None and parent.tag == qn("w:r") else reference)

    dropped = []
    for r_id, rel in list(document.part.rels.items()):
        if "comments" in rel.reltype.lower():
            document.part.drop_rel(r_id)
            dropped.append(rel.reltype)
    for part in list(document.part.package.parts):
        if "comments" in str(part.partname).lower():
            try:
                document.part.package.parts.remove(part)
            except (ValueError, AttributeError):
                pass
    return dropped


def set_different_first_page(section, value=True):
    sectPr = section._sectPr
    existing = sectPr.find(qn("w:titlePg"))
    if value and existing is None:
        insert_ordered(sectPr, OxmlElement("w:titlePg"), SECTPR_TITLE_PAGE)
    elif not value and existing is not None:
        sectPr.remove(existing)


def set_even_odd_headers(document, value=True):
    settings = document.settings.element
    existing = settings.find(qn("w:evenAndOddHeaders"))
    if value and existing is None:
        node = OxmlElement("w:evenAndOddHeaders")
        insert_ordered(settings, node, SETTINGS_AFTER_EVENODD)
    elif not value and existing is not None:
        settings.remove(existing)


def repair_settings(document):
    """Fix known schema defects in python-docx's default settings part."""
    settings = document.settings.element
    for zoom in settings.findall(qn("w:zoom")):
        if zoom.get(qn("w:percent")) is None:
            zoom.set(qn("w:percent"), "100")
    return settings


def set_page_number_start(section, start):
    """Restart page numbering at *start* for this section."""
    sectPr = section._sectPr
    existing = sectPr.find(qn("w:pgNumType"))
    if existing is not None:
        sectPr.remove(existing)
    node = make("w:pgNumType", **{"w:start": start})
    insert_ordered(sectPr, node, [
        "w:cols", "w:formProt", "w:vAlign", "w:noEndnote", "w:titlePg",
        "w:textDirection", "w:bidi", "w:rtlGutter", "w:docGrid",
        "w:printerSettings", "w:sectPrChange",
    ])


def strip_comments_and_changes(document):
    """Convenience: accept revisions, then remove comments."""
    accept_all_changes(document)
    return strip_comments(document)
