"""Low-level OOXML helpers shared by the docx_agent_kit modules."""

from docx.oxml import OxmlElement
from docx.oxml.ns import qn


def make(tag, **attrs):
    """Create an OOXML element. Keys may be 'w:val' or 'r:id'."""
    element = OxmlElement(tag)
    for key, value in attrs.items():
        if value is None:
            continue
        element.set(qn(key), str(value))
    return element


def insert_ordered(parent, child, successors):
    """Insert *child* before the first existing successor tag (schema order)."""
    tags = {qn(tag) for tag in successors}
    for existing in parent:
        if existing.tag in tags:
            existing.addprevious(child)
            return child
    parent.append(child)
    return child


def remove(element):
    if element is not None and element.getparent() is not None:
        element.getparent().remove(element)


def text_of(element):
    """Concatenate every w:t under *element*, in document order."""
    return "".join(node.text or "" for node in element.iter(qn("w:t")))


def set_text(node, text):
    node.text = text
    node.set(qn("xml:space"), "preserve")


# --- schema successor tables -------------------------------------------------

PPR_AFTER_NUMPR = [
    "w:suppressLineNumbers", "w:pBdr", "w:shd", "w:tabs", "w:suppressAutoHyphens",
    "w:kinsoku", "w:wordWrap", "w:overflowPunct", "w:topLinePunct",
    "w:autoSpaceDE", "w:autoSpaceDN", "w:bidi", "w:adjustRightInd",
    "w:snapToGrid", "w:spacing", "w:ind", "w:contextualSpacing",
    "w:mirrorIndents", "w:suppressOverlap", "w:jc", "w:textDirection",
    "w:textAlignment", "w:textboxTightWrap", "w:outlineLvl", "w:divId",
    "w:cnfStyle", "w:rPr", "w:sectPr", "w:pPrChange",
]

PPR_AFTER_PBDR = [
    "w:shd", "w:tabs", "w:suppressAutoHyphens", "w:kinsoku", "w:wordWrap",
    "w:overflowPunct", "w:topLinePunct", "w:autoSpaceDE", "w:autoSpaceDN",
    "w:bidi", "w:adjustRightInd", "w:snapToGrid", "w:spacing", "w:ind",
    "w:contextualSpacing", "w:mirrorIndents", "w:suppressOverlap", "w:jc",
    "w:textDirection", "w:textAlignment", "w:textboxTightWrap", "w:outlineLvl",
    "w:divId", "w:cnfStyle", "w:rPr", "w:sectPr", "w:pPrChange",
]

TBLPR_AFTER_BORDERS = [
    "w:shd", "w:tblLayout", "w:tblCellMar", "w:tblLook", "w:tblCaption",
    "w:tblDescription", "w:tblPrChange",
]

TCPR_AFTER_SHD = [
    "w:noWrap", "w:tcMar", "w:textDirection", "w:tcFitText", "w:vAlign",
    "w:hideMark", "w:headers", "w:cellIns", "w:cellDel", "w:cellMerge",
    "w:tcPrChange",
]

TCPR_AFTER_BORDERS = [
    "w:shd", "w:noWrap", "w:tcMar", "w:textDirection", "w:tcFitText",
    "w:vAlign", "w:hideMark", "w:headers", "w:cellIns", "w:cellDel",
    "w:cellMerge", "w:tcPrChange",
]

SETTINGS_AFTER_EVENODD = [
    "w:bookFoldRevPrinting", "w:bookFoldPrinting", "w:bookFoldPrintingSheets",
    "w:drawingGridHorizontalSpacing", "w:drawingGridVerticalSpacing",
    "w:displayHorizontalDrawingGridEvery", "w:displayVerticalDrawingGridEvery",
    "w:doNotUseMarginsForDrawingGridOrigin", "w:drawingGridHorizontalOrigin",
    "w:drawingGridVerticalOrigin", "w:doNotShadeFormData",
    "w:noPunctuationKerning", "w:characterSpacingControl", "w:printTwoOnOne",
    "w:strictFirstAndLastChars", "w:noLineBreaksAfter", "w:noLineBreaksBefore",
    "w:savePreviewPicture", "w:doNotValidateAgainstSchema", "w:saveInvalidXml",
    "w:ignoreMixedContent", "w:alwaysShowPlaceholderText",
    "w:doNotDemarcateInvalidXml", "w:saveXmlDataOnly", "w:useXSLTWhenSaving",
    "w:saveThroughXslt", "w:showXMLTags", "w:alwaysMergeEmptyNamespace",
    "w:updateFields", "w:hdrShapeDefaults", "w:footnotePr", "w:endnotePr",
    "w:compat", "w:docVars", "w:rsids", "w:mathPr", "w:uiCompat97To2003",
    "w:attachedSchema", "w:themeFontLang", "w:clrSchemeMapping",
    "w:doNotIncludeSubdocsInStats", "w:doNotAutoCompressPictures",
    "w:forceUpgrade", "w:captions", "w:readModeInkLockDown",
    "w:smartTagType", "w:schemaLibrary", "w:shapeDefaults",
    "w:doNotEmbedSmartTags", "w:decimalSymbol", "w:listSeparator",
]

SETTINGS_AFTER_UPDATEFIELDS = [
    "w:hdrShapeDefaults", "w:footnotePr", "w:endnotePr", "w:compat",
    "w:docVars", "w:rsids", "w:mathPr", "w:uiCompat97To2003",
    "w:attachedSchema", "w:themeFontLang", "w:clrSchemeMapping",
    "w:doNotIncludeSubdocsInStats", "w:doNotAutoCompressPictures",
    "w:forceUpgrade", "w:captions", "w:readModeInkLockDown",
    "w:smartTagType", "w:schemaLibrary", "w:shapeDefaults",
    "w:doNotEmbedSmartTags", "w:decimalSymbol", "w:listSeparator",
]

SECTPR_TITLE_PAGE = [
    "w:textDirection", "w:bidi", "w:rtlGutter", "w:docGrid",
    "w:printerSettings", "w:sectPrChange",
]
