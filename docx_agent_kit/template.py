"""Template rendering, manifest discovery and a document inspector."""

from pathlib import Path

from docx import Document
from docx.oxml.ns import qn

from . import blocks
from .docops import accept_all_changes, humanize, repair_settings, strip_comments
from .fields import update_fields


def render(template_path, payload, out_path=None, strict=False,
           update=True, accept_changes=False, remove_comments=False):
    """Fill a .docx template from a dict and save the result.

    ``payload`` keys map to ``{{key}}`` tokens; list values fill
    ``{{#key}}...{{/key}}`` blocks; image keys map to paths.
    """
    document = Document(template_path)
    repair_settings(document)
    if accept_changes:
        accept_all_changes(document)
    if remove_comments:
        strip_comments(document)

    list_values = {k: v for k, v in payload.items() if isinstance(v, list)}
    scalar_values = {
        k: v for k, v in payload.items()
        if not isinstance(v, (list, dict)) and not isinstance(v, bytes)
    }

    if list_values:
        blocks.fill_blocks(document, list_values)
    blocks.replace_placeholders(document, scalar_values, strict=False,
                                include_headers=True)
    blocks.fill_image_placeholders(document, payload)
    if update:
        update_fields(document, True)

    if strict:
        remaining = set()
        for kind, key in blocks.find_tokens(document.element.body):
            if kind == "text" and key not in scalar_values:
                remaining.add(key)
        if remaining:
            raise KeyError(f"unfilled template tokens: {sorted(remaining)}")

    out_path = out_path or str(Path(template_path).with_name(
        Path(template_path).stem + "-filled.docx"))
    humanize(document)
    document.save(out_path)
    return out_path


def manifest_from_template(template_path):
    """Build a field schema from the tokens present in a template."""
    document = Document(template_path)
    fields = []
    seen = set()

    def add(kind, key, **extra):
        if key in seen:
            return
        seen.add(key)
        entry = {"key": key, "type": "image" if kind == "image" else "text"}
        entry.update(extra)
        fields.append(entry)

    children = list(document.element.body)
    block_ranges = []
    index = 0
    while index < len(children):
        text = blocks.text_of(children[index]).strip()
        match = blocks.BLOCK_START_RE.match(text)
        if not match:
            index += 1
            continue
        name = match.group(1)
        depth = 1
        end = index + 1
        while end < len(children):
            inner = blocks.text_of(children[end]).strip()
            if inner == f"{{{{#{name}}}}}":
                depth += 1
            elif inner == f"{{{{/{name}}}}}":
                depth -= 1
                if depth == 0:
                    break
            end += 1
        sub = []
        for element in children[index + 1:end]:
            for kind, key in blocks.find_tokens(element):
                if kind == "text" and not any(f["key"] == key for f in sub):
                    sub.append({"key": key, "type": "text"})
        fields.append({"key": name, "type": "repeat", "block": name, "fields": sub})
        seen.add(name)
        block_ranges.append((index, end))
        index = end + 1

    covered = set()
    for start, end in block_ranges:
        covered.update(range(start, end + 1))
    for position, element in enumerate(children):
        if position in covered:
            continue
        for kind, key in blocks.find_tokens(element):
            if kind == "image":
                add(kind, key)
            elif kind == "text":
                add(kind, key)
    for part in document.part.package.parts:
        from docx.opc.constants import CONTENT_TYPE as CT
        if part.content_type in (CT.WML_HEADER, CT.WML_FOOTER):
            for kind, key in blocks.find_tokens(part.element):
                add(kind, key)

    return {"template": str(template_path), "fields": fields}


def inspect(template_path):
    """Report which advanced features a document uses."""
    document = Document(template_path)
    body = document.element.body
    numbering = None
    try:
        numbering = document.part.numbering_part.element
    except Exception:
        pass
    report = {
        "paragraphs": len(document.paragraphs),
        "tables": len(document.tables),
        "sections": len(document.sections),
        "inline_images": len(document.inline_shapes),
        "hyperlinks": len(body.findall(".//" + qn("w:hyperlink"))),
        "fields": len(body.findall(".//" + qn("w:fldChar"))),
        "bookmarks": len(body.findall(".//" + qn("w:bookmarkStart"))),
        "comments": len(body.findall(".//" + qn("w:commentReference"))),
        "revisions": len(body.findall(".//" + qn("w:ins")))
                     + len(body.findall(".//" + qn("w:del"))),
        "numbering_definitions": 0 if numbering is None
                                 else len(numbering.findall(qn("w:num"))),
        "headers_footers": sum(
            1 for part in document.part.package.parts
            if part.content_type.endswith(("header+xml", "footer+xml"))
        ),
        "has_toc_field": any(
            "TOC" in (node.text or "")
            for node in body.iter(qn("w:instrText"))
        ),
    }
    return report
