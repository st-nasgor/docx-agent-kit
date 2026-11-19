"""Picture helpers: add inline images and swap an existing image's bytes."""

from pathlib import Path

from docx.oxml.ns import qn
from docx.shared import Inches


def add_picture(paragraph, path, width=None, height=None):
    run = paragraph.add_run()
    run.add_picture(path, width=Inches(width) if width else None,
                    height=Inches(height) if height else None)
    return run


def list_images(document):
    """Return [{'index', 'width', 'height', 'content_type', 'partname'}]."""
    result = []
    for index, shape in enumerate(document.inline_shapes):
        blip = shape._inline.graphic.graphicData.pic.blipFill.blip
        r_id = blip.get(qn("r:embed"))
        part = document.part.related_parts.get(r_id)
        result.append({
            "index": index,
            "width": shape.width,
            "height": shape.height,
            "content_type": getattr(part, "content_type", None),
            "partname": str(getattr(part, "partname", "")),
        })
    return result


def replace_inline_image(document, index, path):
    """Replace the bytes of an inline image, keeping its displayed size.

    The new file must use the same image format (PNG for PNG, JPEG for
    JPEG); changing format requires rebuilding the relationship.
    """
    shape = document.inline_shapes[index]
    blip = shape._inline.graphic.graphicData.pic.blipFill.blip
    r_id = blip.get(qn("r:embed"))
    part = document.part.related_parts[r_id]
    part._blob = Path(path).read_bytes()
    return part


def replace_image_by_partname(document, partname_suffix, path):
    """Replace an image part by a suffix such as 'image1.png'."""
    for part in document.part.package.parts:
        if str(part.partname).endswith(partname_suffix):
            part._blob = Path(path).read_bytes()
            return part
    raise KeyError(f"no image part ending with {partname_suffix!r}")
