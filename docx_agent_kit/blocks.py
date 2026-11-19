"""Placeholder replacement, repeatable blocks and repeatable table rows.

Template conventions (all optional):

    {{client_name}}          simple text slot
    {{#line_items}} ... {{/line_items}}   repeat the paragraphs in between
    {{image:hero}}           paragraph that becomes a picture

Tokens may be split across several Word runs; replacement works at the
XML text-node level and keeps the surrounding formatting.
"""

import re
from copy import deepcopy

from docx.oxml.ns import qn
from docx.shared import Inches

from .oxml import remove, set_text, text_of

TOKEN_RE = re.compile(r"\{\{([^{}]+)\}\}")
BLOCK_START_RE = re.compile(r"^\{\{\s*#\s*([\w.\-]+)\s*\}\}$")
BLOCK_END_RE = re.compile(r"^\{\{\s*/\s*([\w.\-]+)\s*\}\}$")
IMAGE_RE = re.compile(r"^\{\{\s*image:([\w.\-]+)\s*\}\}$")


def _token(inner):
    value = inner.strip()
    if value.startswith("#") or value.startswith("/"):
        return None, None
    if value.startswith("image:"):
        return "image", value[6:].strip()
    return "text", value


def replace_text(root, mapping, strict=False):
    """Replace ``{{key}}`` tokens under *root*; return the set of missing keys."""
    nodes = list(root.iter(qn("w:t")))
    if not nodes:
        return set()
    full = "".join(node.text or "" for node in nodes)
    matches = list(TOKEN_RE.finditer(full))
    if not matches:
        return set()

    spans = []
    position = 0
    for node in nodes:
        length = len(node.text or "")
        spans.append((node, position, position + length))
        position += length

    def locate(index):
        for node, start, end in spans:
            if start <= index < end:
                return node, start
        node, start, _ = spans[-1]
        return node, start

    missing = set()
    for match in reversed(matches):
        kind, key = _token(match.group(1))
        if kind != "text":
            continue
        if key not in mapping:
            missing.add(key)
            continue
        value = "" if mapping[key] is None else str(mapping[key])
        first, first_start = locate(match.start())
        last, last_start = locate(match.end() - 1)
        if first is last:
            text = first.text or ""
            set_text(
                first,
                text[: match.start() - first_start] + value + text[match.end() - first_start:],
            )
        else:
            text = first.text or ""
            set_text(first, text[: match.start() - first_start] + value)
            started = False
            for node, _, _ in spans:
                if node is first:
                    started = True
                elif node is last:
                    break
                elif started:
                    set_text(node, "")
            text = last.text or ""
            set_text(last, text[match.end() - last_start:])

    if strict and missing:
        raise KeyError(f"missing template values: {sorted(missing)}")
    return missing


def replace_placeholders(document, mapping, strict=False, include_headers=True):
    """Replace text tokens in the body and (optionally) headers/footers."""
    missing = replace_text(document.element.body, mapping, strict=False)
    if include_headers:
        from docx.opc.constants import CONTENT_TYPE as CT

        for part in document.part.package.parts:
            if part.content_type in (CT.WML_HEADER, CT.WML_FOOTER):
                missing |= replace_text(part.element, mapping, strict=False)
    if strict and missing:
        raise KeyError(f"missing template values: {sorted(missing)}")
    return missing


def fill_blocks(document, mapping):
    """Expand ``{{#name}}...{{/name}}`` regions once per item in mapping[name]."""
    body = document.element.body
    children = list(body)
    index = 0
    while index < len(children):
        start_match = BLOCK_START_RE.match(text_of(children[index]).strip())
        if not start_match:
            index += 1
            continue
        name = start_match.group(1)
        depth = 1
        end_index = index + 1
        while end_index < len(children):
            text = text_of(children[end_index]).strip()
            if text == f"{{{{#{name}}}}}":
                depth += 1
            elif text == f"{{{{/{name}}}}}":
                depth -= 1
                if depth == 0:
                    break
            end_index += 1
        if end_index >= len(children):
            index += 1
            continue

        start_element = children[index]
        end_element = children[end_index]
        template_elements = children[index + 1:end_index]
        items = mapping.get(name) or []
        for item in items:
            for template in template_elements:
                clone = deepcopy(template)
                replace_text(clone, item if isinstance(item, dict) else {"": item})
                start_element.addprevious(clone)
        for element in template_elements:
            body.remove(element)
        body.remove(start_element)
        body.remove(end_element)
        children = list(body)
        index = 0


def repeat_table_row(table, row_index, items, keep_template=False):
    """Clone a table row once per item, replacing its ``{{tokens}}``."""
    template = table.rows[row_index]._tr
    anchor = template
    for item in items:
        clone = deepcopy(template)
        replace_text(clone, item)
        anchor.addnext(clone)
        anchor = clone
    if not keep_template:
        remove(template)


def fill_image_placeholders(document, mapping, default_width=None):
    """Replace paragraphs containing only ``{{image:key}}`` with a picture.

    mapping values may be a path, or ``(path, width_inches)``.
    """
    replaced = []
    for paragraph in document.paragraphs:
        match = IMAGE_RE.match(paragraph.text.strip())
        if not match:
            continue
        key = match.group(1)
        if key not in mapping:
            continue
        value = mapping[key]
        path, width = (value if isinstance(value, (tuple, list)) else (value, None))
        width = width or default_width
        for run in list(paragraph.runs):
            run._r.getparent().remove(run._r)
        run = paragraph.add_run()
        run.add_picture(path, width=Inches(width) if width else None)
        replaced.append(key)
    return replaced


def find_tokens(root):
    """Return every token found under *root* as (kind, key) pairs."""
    found = []
    for match in TOKEN_RE.finditer(text_of(root)):
        found.append(_token(match.group(1)))
    return [pair for pair in found if pair[0]]


def iter_paragraphs(root):
    return root.iter(qn("w:p"))
