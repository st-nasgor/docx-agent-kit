"""Table helpers: borders, shading, widths, merges, repeating headers."""

from docx.shared import Inches
from docx.oxml import OxmlElement
from docx.oxml.ns import qn, nsdecls

from .oxml import make, insert_ordered, TBLPR_AFTER_BORDERS, TCPR_AFTER_SHD, TCPR_AFTER_BORDERS

EDGES = ("top", "left", "bottom", "right", "insideH", "insideV")


def _border_element(edge, style="single", size=4, color="BFBFBF"):
    return make(
        f"w:{edge}",
        **{"w:val": style, "w:sz": size, "w:space": 0, "w:color": color.lstrip("#")},
    )


def set_table_borders(table, style="single", size=4, color="BFBFBF",
                      edges=EDGES):
    """Apply borders to every edge; pass style=None to remove them."""
    tblPr = table._tbl.tblPr
    existing = tblPr.find(qn("w:tblBorders"))
    if existing is not None:
        tblPr.remove(existing)
    if style is None:
        return
    borders = OxmlElement("w:tblBorders")
    for edge in EDGES:
        if edge in edges:
            borders.append(_border_element(edge, style, size, color))
        else:
            borders.append(make(f"w:{edge}", **{"w:val": "nil"}))
    insert_ordered(tblPr, borders, TBLPR_AFTER_BORDERS)


def set_cell_shading(cell, fill, pattern="clear"):
    """Fill a cell with a hex color such as 'FFF2CC' or 'auto'."""
    tcPr = cell._tc.get_or_add_tcPr()
    existing = tcPr.find(qn("w:shd"))
    if existing is not None:
        tcPr.remove(existing)
    shd = make(
        "w:shd",
        **{"w:val": pattern, "w:color": "auto", "w:fill": (fill or "auto").lstrip("#")},
    )
    insert_ordered(tcPr, shd, TCPR_AFTER_SHD)


def set_cell_borders(cell, style="single", size=4, color="BFBFBF", edges=EDGES):
    tcPr = cell._tc.get_or_add_tcPr()
    existing = tcPr.find(qn("w:tcBorders"))
    if existing is not None:
        tcPr.remove(existing)
    if style is None:
        return
    borders = OxmlElement("w:tcBorders")
    for edge in EDGES:
        if edge in edges and edge not in ("insideH", "insideV"):
            borders.append(_border_element(edge, style, size, color))
    insert_ordered(tcPr, borders, TCPR_AFTER_BORDERS)


def set_col_widths(table, widths_inches):
    """Set fixed column widths (inches) on the table and every cell."""
    table.autofit = False
    tblW = table._tbl.tblPr.find(qn("w:tblW"))
    total = sum(widths_inches)
    if tblW is None:
        tblW = OxmlElement("w:tblW")
        table._tbl.tblPr.append(tblW)
    tblW.set(qn("w:type"), "dxa")
    tblW.set(qn("w:w"), str(int(total * 1440)))
    for row in table.rows:
        for index, width in enumerate(widths_inches):
            if index < len(row.cells):
                row.cells[index].width = Inches(width)
    for index, width in enumerate(widths_inches):
        if index < len(table.columns):
            table.columns[index].width = Inches(width)


def merge_cells(table, start_row, start_col, end_row, end_col):
    return table.cell(start_row, start_col).merge(table.cell(end_row, end_col))


def set_repeat_header(row):
    """Mark a row so Word repeats it at the top of each page."""
    trPr = row._tr.get_or_add_trPr()
    if trPr.find(qn("w:tblHeader")) is None:
        trPr.append(OxmlElement("w:tblHeader"))


def set_cant_split(row):
    trPr = row._tr.get_or_add_trPr()
    if trPr.find(qn("w:cantSplit")) is None:
        trPr.append(OxmlElement("w:cantSplit"))


def set_cell_margins(table, top=40, start=80, bottom=40, end=80):
    """Cell padding in twentieths of a point (dxa)."""
    tblPr = table._tbl.tblPr
    existing = tblPr.find(qn("w:tblCellMar"))
    if existing is not None:
        tblPr.remove(existing)
    mar = parse_cell_margins(top, start, bottom, end)
    insert_ordered(tblPr, mar, ["w:tblLook", "w:tblCaption", "w:tblDescription", "w:tblPrChange"])


def parse_cell_margins(top, start, bottom, end):
    from docx.oxml import parse_xml

    return parse_xml(
        f'<w:tblCellMar {nsdecls("w")}>'
        f'<w:top w:w="{top}" w:type="dxa"/>'
        f'<w:left w:w="{start}" w:type="dxa"/>'
        f'<w:bottom w:w="{bottom}" w:type="dxa"/>'
        f'<w:right w:w="{end}" w:type="dxa"/>'
        f"</w:tblCellMar>"
    )
