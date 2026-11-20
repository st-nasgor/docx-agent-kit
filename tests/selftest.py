"""End-to-end self test for docx_agent_kit. Run: python tests/selftest.py"""

import json
import os
import shutil
import struct
import subprocess
import sys
import zlib
from pathlib import Path

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor

import docx_agent_kit as dk

ROOT = Path(__file__).resolve().parent.parent
WORK = Path(os.environ.get("LOCALAPPDATA", ROOT)) / "Temp" / "opencode" / "docx-agent-kit-test"
SKILL = Path.home() / ".config" / "opencode" / "skills" / "docx"
SOFFICE_CANDIDATES = [
    Path(os.environ.get("ProgramFiles", "C:/Program Files")) / "LibreOffice/program/soffice.com",
    Path(os.environ.get("ProgramFiles(x86)", "C:/Program Files (x86)")) / "LibreOffice/program/soffice.com",
]

checks = []


def check(name, condition, detail=""):
    checks.append((name, bool(condition), detail))
    print(("PASS " if condition else "FAIL ") + name + (f"  [{detail}]" if detail and not condition else ""))
    if not condition:
        raise AssertionError(name)


def make_png(path, width=160, height=90, rgb=(45, 90, 160)):
    def chunk(tag, data):
        return (struct.pack(">I", len(data)) + tag + data
                + struct.pack(">I", zlib.crc32(tag + data) & 0xFFFFFFFF))
    raw = b"".join(b"\x00" + bytes(rgb) * width for _ in range(height))
    payload = (b"\x89PNG\r\n\x1a\n"
               + chunk(b"IHDR", struct.pack(">IIBBBBB", width, height, 8, 2, 0, 0, 0))
               + chunk(b"IDAT", zlib.compress(raw))
               + chunk(b"IEND", b""))
    path.write_bytes(payload)
    return path


def wrap(paragraph, start_tag, run):
    element = OxmlElement(start_tag)
    element.set(qn("w:id"), "1")
    element.set(qn("w:author"), "tester")
    element.set(qn("w:date"), "2026-01-01T00:00:00Z")
    run._r.addprevious(element)
    element.append(run._r)
    return element


def main():
    if WORK.exists():
        shutil.rmtree(WORK)
    WORK.mkdir(parents=True)
    image = make_png(WORK / "hero.png")
    image2 = make_png(WORK / "hero2.png", rgb=(200, 60, 40))

    doc = Document()
    doc.styles["Normal"].font.name = "Calibri"
    doc.styles["Normal"].font.size = Pt(10.5)

    # 1. split-run placeholder replacement
    paragraph = doc.add_paragraph()
    paragraph.add_run("Hello ")
    paragraph.add_run("{{cli")
    paragraph.add_run("ent}}")
    dk.replace_placeholders(doc, {"client": "Acme"})
    check("placeholder across runs", paragraph.text == "Hello Acme", paragraph.text)

    # 2. hyperlink
    link_paragraph = doc.add_paragraph("Site: ")
    dk.add_hyperlink(link_paragraph, "https://example.com", "example.com")
    body = doc.element.body
    check("hyperlink element", len(body.findall(".//" + qn("w:hyperlink"))) == 1)
    r_ids = [rel for rel in doc.part.rels.values() if rel.reltype.endswith("/hyperlink")]
    check("hyperlink relationship", len(r_ids) == 1 and r_ids[0].is_external)

    # 3. fields (body DATE, footer PAGE)
    date_run = dk.add_date_field(doc.add_paragraph("Generated: "), result="")
    dk.update_fields(doc)
    instructions = dk.field_instructions(doc.element.body)
    check("date field", any("DATE" in item for item in instructions), instructions)
    check("updateFields flag", doc.settings.element.find(qn("w:updateFields")) is not None)

    footer = doc.sections[0].footer.paragraphs[0]
    footer.add_run("Page ")
    dk.add_page_number(footer, result="1")
    check("footer page field",
          any("PAGE" in item for item in dk.field_instructions(footer._p)))

    # 4. lists
    bullets = dk.new_list(doc, "bullet", levels=3)
    dk.add_item(doc, "First bullet", bullets, 0)
    dk.add_item(doc, "Nested bullet", bullets, 1)
    numbers = dk.new_list(doc, "number")
    dk.add_item(doc, "Step one", numbers, 0)
    dk.add_item(doc, "Step two", numbers, 0)
    numbering = doc.part.numbering_part.element
    num_prs = body.findall(".//" + qn("w:numPr"))
    check("list paragraphs", len(num_prs) == 4, len(num_prs))
    check("new numbering definitions",
          len(numbering.findall(qn("w:num"))) >= 11,
          len(numbering.findall(qn("w:num"))))

    # 5. table
    table = doc.add_table(rows=4, cols=3)
    table.style = "Table Grid"
    dk.set_col_widths(table, [1.4, 2.6, 2.5])
    dk.set_table_borders(table, color="9E9E9E", size=4)
    dk.set_cell_shading(table.cell(0, 0), "1F3864")
    dk.set_cell_shading(table.cell(0, 1), "1F3864")
    dk.set_cell_shading(table.cell(0, 2), "1F3864")
    dk.set_cell_margins(table)
    dk.set_repeat_header(table.rows[0])
    dk.merge_cells(table, 1, 0, 1, 1)
    for col in range(3):
        table.cell(0, col).text = "Header %d" % (col + 1)
    table.cell(1, 0).text = "Merged"
    dk.repeat_table_row(table, 3, [{"col": "row A"}, {"col": "row B"}], keep_template=True)
    check("table borders", table._tbl.tblPr.find(qn("w:tblBorders")) is not None)
    check("cell shading", table.cell(0, 0)._tc.tcPr.find(qn("w:shd")) is not None)
    check("merged cells", table.cell(1, 0)._tc.grid_span == 2,
          table.cell(1, 0)._tc.grid_span)
    check("repeat header", table.rows[0]._tr.trPr.find(qn("w:tblHeader")) is not None)

    # 6. image placeholder
    image_paragraph = doc.add_paragraph("{{image:hero}}")
    dk.fill_image_placeholders(doc, {"hero": (str(image), 2.2)})
    check("image placeholder", len(doc.inline_shapes) == 1 and not image_paragraph.text)
    original_bytes = (WORK / "hero.png").read_bytes()
    dk.replace_inline_image(doc, 0, image2)
    r_id = doc.inline_shapes[0]._inline.graphic.graphicData.pic.blipFill.blip.get(qn("r:embed"))
    check("image bytes swapped",
          doc.part.related_parts[r_id].blob == (WORK / "hero2.png").read_bytes())

    # 7. repeat blocks
    doc.add_paragraph("{{#items}}")
    doc.add_paragraph("Item: {{name}} - {{desc}}")
    doc.add_paragraph("{{/items}}")
    dk.fill_blocks(doc, {"items": [
        {"name": "Alpha", "desc": "first"},
        {"name": "Beta", "desc": "second"},
    ]})
    texts = [p.text for p in doc.paragraphs]
    check("block expansion", "Item: Alpha - first" in texts and "Item: Beta - second" in texts)
    check("block markers removed", "{{#items}}" not in texts and "{{/items}}" not in texts)

    # 8. header token replacement
    header = doc.sections[0].header.paragraphs[0]
    header.text = "{{doc_title}}"
    dk.set_different_first_page(doc.sections[0], True)
    dk.set_even_odd_headers(doc, True)
    dk.set_page_number_start(doc.sections[0], 3)
    dk.replace_placeholders(doc, {"doc_title": "Quarterly Report"})
    check("header token", header.text == "Quarterly Report", header.text)
    check("titlePg set", doc.sections[0]._sectPr.find(qn("w:titlePg")) is not None)
    check("even/odd set", doc.settings.element.find(qn("w:evenAndOddHeaders")) is not None)
    check("page start set", doc.sections[0]._sectPr.find(qn("w:pgNumType")) is not None)
    dk.set_page_number_start(doc.sections[0], 3, fmt="lowerRoman")
    pgnum = doc.sections[0]._sectPr.find(qn("w:pgNumType"))
    check("page fmt set", pgnum is not None and pgnum.get(qn("w:fmt")) == "lowerRoman")

    # 8b. section helpers: new_section clears inherited titlePg, footer PAGE field
    sec_doc = Document()
    dk.set_different_first_page(sec_doc.sections[0], True)
    sec2 = dk.new_section(sec_doc)
    check("new_section titlePg cleared", sec2._sectPr.find(qn("w:titlePg")) is None)
    dk.set_footer_page_number(sec2)
    check("footer unlinked", sec2.footer.is_linked_to_previous is False)
    check("footer page field", "PAGE" in sec2.footer.paragraphs[0]._p.xml)
    centered = sec2.footer.paragraphs[0].alignment
    check("footer centered", centered is not None and int(centered) == 1, centered)
    break_para = sec_doc.add_paragraph("Next chapter")
    dk.page_break_before(break_para)
    check("page_break_before", break_para.paragraph_format.page_break_before is True)

    # 9. merge runs
    merge_paragraph = doc.add_paragraph()
    merge_paragraph.add_run("Split ")
    merge_paragraph.add_run("across ")
    merge_paragraph.add_run("runs.")
    merged = dk.merge_runs(doc)
    check("merge runs", merged >= 2 and merge_paragraph.text == "Split across runs.", merged)

    # 10. revisions and comments
    revision_paragraph = doc.add_paragraph("Keep ")
    wrap(revision_paragraph, "w:ins", revision_paragraph.add_run("inserted"))
    wrap(revision_paragraph, "w:del", revision_paragraph.add_run("deleted"))
    comment_run = revision_paragraph.add_run()
    comment_run._r.append(OxmlElement("w:commentReference"))
    revision_paragraph._p.insert(0, OxmlElement("w:commentRangeStart"))
    dk.accept_all_changes(doc)
    dk.strip_comments(doc)
    check("tracked changes accepted", revision_paragraph.text == "Keep inserted",
          revision_paragraph.text)
    check("comment markers stripped",
          not body.findall(".//" + qn("w:commentReference")))

    # 11. manifest + render round trip
    template = Document()
    template.add_paragraph("Title: {{title}}")
    template.add_paragraph("Note: {{note}}")
    template.add_paragraph("{{image:logo}}")
    template.add_paragraph("{{#rows}}")
    template.add_paragraph("- {{label}}")
    template.add_paragraph("{{/rows}}")
    template_path = WORK / "template.docx"
    template.save(template_path)
    manifest = dk.manifest_from_template(template_path)
    kinds = {field["key"]: field["type"] for field in manifest["fields"]}
    check("manifest text field", kinds.get("title") == "text", kinds)
    check("manifest image field", kinds.get("logo") == "image", kinds)
    check("manifest repeat field", kinds.get("rows") == "repeat", kinds)
    top_level = [field["key"] for field in manifest["fields"]]
    check("manifest has no leaked block fields", top_level.count("label") == 0, top_level)
    check("manifest field count", len(top_level) == 4, top_level)

    out = dk.render(template_path, {
        "title": "When Pigs Fly",
        "note": "Filled by docx_agent_kit",
        "logo": (str(image), 1.5),
        "rows": [{"label": "one"}, {"label": "two"}],
    }, WORK / "rendered.docx")
    rendered = Document(out)
    rendered_text = "\n".join(p.text for p in rendered.paragraphs)
    check("render text", "Title: When Pigs Fly" in rendered_text, rendered_text)
    check("render block", "- one" in rendered_text and "- two" in rendered_text)
    check("render image", len(rendered.inline_shapes) == 1)

    # save the feature document and validate
    feature_path = WORK / "features.docx"
    doc.save(feature_path)
    report = dk.inspect(feature_path)
    check("inspect hyperlinks", report["hyperlinks"] >= 1, report)
    check("inspect fields", report["fields"] >= 3, report["fields"])

    validate = SKILL / "scripts" / "office" / "validate.py"
    if validate.exists():
        result = subprocess.run([sys.executable, str(validate), str(feature_path)],
                                capture_output=True, text=True)
        check("ooxml validation", "PASSED" in result.stdout,
              (result.stdout + result.stderr)[-500:])

    soffice = next((path for path in SOFFICE_CANDIDATES if path.exists()), None)
    if soffice:
        subprocess.run([str(soffice), "--headless",
                        "-env:UserInstallation=" + (WORK / "lo").as_uri(),
                        "--convert-to", "pdf", "--outdir", str(WORK),
                        str(feature_path)], capture_output=True, text=True)
        pdf = feature_path.with_suffix(".pdf")
        check("pdf conversion", pdf.exists())
        if pdf.exists():
            import pymupdf

            pdf_doc = pymupdf.open(pdf)
            text = "".join(page.get_text() for page in pdf_doc)
            check("pdf pages", pdf_doc.page_count >= 1, pdf_doc.page_count)
            check("pdf contains replacement", "Acme" in text)
            pdf_doc.close()
    else:
        print("SKIP pdf conversion (LibreOffice not found)")

    passed = sum(1 for _, ok, _ in checks if ok)
    print(f"\n{passed}/{len(checks)} checks passed")
    print(json.dumps(report, indent=2, default=str))
    return 0 if passed == len(checks) else 1


if __name__ == "__main__":
    raise SystemExit(main())
