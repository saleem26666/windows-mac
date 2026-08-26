"""Build Word (.docx) and PDF guides from the English and Urdu markdown files."""

from __future__ import annotations

import re
import sys
import time
from pathlib import Path

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

ROOT = Path(__file__).resolve().parent

NAVY = RGBColor(0x1B, 0x3A, 0x4B)
INK = RGBColor(0x1F, 0x29, 0x37)
MUTED = RGBColor(0x4B, 0x55, 0x63)
ACCENT = RGBColor(0xC4, 0x56, 0x1A)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
RULE = "D6D3CD"
HEAD_FILL = "1B3A4B"
ALT_FILL = "F4F1EA"

INLINE_RE = re.compile(r"(\*\*[^*]+\*\*|`[^`]+`)")


def set_run_font(run, latin, cs=None, size=11, bold=False, color=None, italic=False):
    run.bold = bold
    run.italic = italic
    run.font.size = Pt(size)
    run.font.name = latin
    if color is not None:
        run.font.color.rgb = color
    r_pr = run._element.get_or_add_rPr()
    r_fonts = r_pr.find(qn("w:rFonts"))
    if r_fonts is None:
        r_fonts = OxmlElement("w:rFonts")
        r_pr.insert(0, r_fonts)
    complex_name = cs or latin
    r_fonts.set(qn("w:ascii"), latin)
    r_fonts.set(qn("w:hAnsi"), latin)
    r_fonts.set(qn("w:eastAsia"), latin)
    r_fonts.set(qn("w:cs"), complex_name)


def shade_cell(cell, fill):
    tc_pr = cell._tc.get_or_add_tcPr()
    for child in list(tc_pr):
        if child.tag == qn("w:shd"):
            tc_pr.remove(child)
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), fill)
    shd.set(qn("w:val"), "clear")
    tc_pr.append(shd)


def set_cell_borders(cell, color=RULE):
    tc_pr = cell._tc.get_or_add_tcPr()
    for child in list(tc_pr):
        if child.tag == qn("w:tcBorders"):
            tc_pr.remove(child)
    borders = OxmlElement("w:tcBorders")
    for edge in ("top", "left", "bottom", "right"):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), "4")
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), color)
        borders.append(el)
    tc_pr.append(borders)


def set_cell_margins(cell, top=70, bottom=70, left=90, right=90):
    tc_pr = cell._tc.get_or_add_tcPr()
    for child in list(tc_pr):
        if child.tag == qn("w:tcMar"):
            tc_pr.remove(child)
    tc_mar = OxmlElement("w:tcMar")
    for name, val in (("top", top), ("left", left), ("bottom", bottom), ("right", right)):
        node = OxmlElement(f"w:{name}")
        node.set(qn("w:w"), str(val))
        node.set(qn("w:type"), "dxa")
        tc_mar.append(node)
    tc_pr.append(tc_mar)


def prevent_row_split(row):
    tr_pr = row._tr.get_or_add_trPr()
    tr_pr.append(OxmlElement("w:cantSplit"))


def set_paragraph_bidi(paragraph, rtl):
    p_pr = paragraph._p.get_or_add_pPr()
    for child in list(p_pr):
        if child.tag == qn("w:bidi"):
            p_pr.remove(child)
    if not rtl:
        return
    bidi = OxmlElement("w:bidi")
    bidi.set(qn("w:val"), "1")
    p_pr.append(bidi)
    paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    jc = p_pr.find(qn("w:jc"))
    if jc is None:
        jc = OxmlElement("w:jc")
        p_pr.append(jc)
    jc.set(qn("w:val"), "right")


def add_page_field(paragraph, latin, cs, rtl):
    run = paragraph.add_run()
    set_run_font(run, latin, cs, 9, color=MUTED)
    r = run._r
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = " PAGE "
    separate = OxmlElement("w:fldChar")
    separate.set(qn("w:fldCharType"), "separate")
    text = OxmlElement("w:t")
    text.text = "1"
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    r.append(begin)
    r.append(instr)
    r.append(separate)
    r.append(text)
    r.append(end)


def add_horizontal_line(paragraph):
    p_pr = paragraph._p.get_or_add_pPr()
    for child in list(p_pr):
        if child.tag == qn("w:pBdr"):
            p_pr.remove(child)
    p_bdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "12")
    bottom.set(qn("w:space"), "1")
    bottom.set(qn("w:color"), HEAD_FILL)
    p_bdr.append(bottom)
    p_pr.append(p_bdr)


def add_inline(paragraph, text, latin, cs, size, rtl, color=INK, bold=False, italic=False):
    parts = INLINE_RE.split(text or "")
    if not parts:
        run = paragraph.add_run(text or "")
        set_run_font(run, latin, cs, size, bold=bold, color=color, italic=italic)
        return
    for part in parts:
        if not part:
            continue
        if part.startswith("**") and part.endswith("**") and len(part) >= 4:
            run = paragraph.add_run(part[2:-2])
            set_run_font(run, latin, cs, size, bold=True, color=color, italic=italic)
        elif part.startswith("`") and part.endswith("`") and len(part) >= 2:
            run = paragraph.add_run(part[1:-1])
            set_run_font(run, "Consolas", cs, max(size - 0.5, 9), color=NAVY)
        else:
            run = paragraph.add_run(part)
            set_run_font(run, latin, cs, size, bold=bold, color=color, italic=italic)


def is_table_separator(line: str) -> bool:
    stripped = line.strip().strip("|")
    return bool(stripped) and "-" in stripped and bool(re.match(r"^[\s:|-]+$", stripped))


def split_row(line: str) -> list[str]:
    line = line.strip()
    if line.startswith("|"):
        line = line[1:]
    if line.endswith("|"):
        line = line[:-1]
    return [c.strip() for c in line.split("|")]


def parse_markdown(md: str) -> list[tuple]:
    md = re.sub(r"</?div[^>]*>", "", md)
    md = md.replace("\r\n", "\n").strip()
    lines = md.split("\n")
    blocks: list[tuple] = []
    i = 0
    n = len(lines)
    while i < n:
        stripped = lines[i].strip()
        if not stripped:
            i += 1
            continue
        if stripped.startswith("#"):
            hashes = len(stripped) - len(stripped.lstrip("#"))
            blocks.append(("h", min(hashes, 3), stripped[hashes:].strip()))
            i += 1
            continue
        if stripped in ("---", "***", "___"):
            blocks.append(("hr",))
            i += 1
            continue
        if stripped.startswith("|") and i + 1 < n and is_table_separator(lines[i + 1]):
            headers = split_row(stripped)
            i += 2
            rows = []
            while i < n and lines[i].strip().startswith("|"):
                rows.append(split_row(lines[i]))
                i += 1
            blocks.append(("table", headers, rows))
            continue
        if re.match(r"^[-*] \[[ xX]\] ", stripped):
            items = []
            while i < n and re.match(r"^[-*] \[[ xX]\] ", lines[i].strip()):
                raw = lines[i].strip()
                items.append((raw[3] in "xX", raw[6:].strip()))
                i += 1
            blocks.append(("tasks", items))
            continue
        if re.match(r"^[-*] ", stripped):
            items = []
            while i < n and re.match(r"^[-*] ", lines[i].strip()):
                items.append(lines[i].strip()[2:].strip())
                i += 1
            blocks.append(("ul", items))
            continue
        if re.match(r"^\d+\.\s+", stripped):
            items = []
            while i < n and re.match(r"^\d+\.\s+", lines[i].strip()):
                items.append(re.sub(r"^\d+\.\s+", "", lines[i].strip()))
                i += 1
            blocks.append(("ol", items))
            continue
        para_lines = [stripped]
        i += 1
        while i < n:
            ns = lines[i].strip()
            if not ns:
                break
            if ns.startswith("#") or ns.startswith("|") or ns in ("---", "***", "___"):
                break
            if re.match(r"^[-*] ", ns) or re.match(r"^\d+\.\s+", ns) or re.match(r"^[-*] \[[ xX]\] ", ns):
                break
            para_lines.append(ns)
            i += 1
        blocks.append(("p", " ".join(para_lines)))
    return blocks


def configure_document(rtl: bool) -> Document:
    doc = Document()
    section = doc.sections[0]
    section.page_width = Cm(21.0)
    section.page_height = Cm(29.7)
    section.top_margin = Cm(1.8)
    section.bottom_margin = Cm(1.8)
    section.left_margin = Cm(1.8)
    section.right_margin = Cm(1.8)

    latin = "Calibri"
    cs = "Arial Unicode MS" if rtl else "Calibri"
    normal = doc.styles["Normal"]
    normal.font.name = cs if rtl else latin
    normal.font.size = Pt(12 if rtl else 11)
    normal.font.color.rgb = INK
    normal.paragraph_format.space_after = Pt(8)
    normal.paragraph_format.line_spacing = 1.15
    if rtl:
        defaults = doc.styles.element.find(qn("w:docDefaults"))
        if defaults is not None:
            p_def = defaults.find(qn("w:pPrDefault"))
            if p_def is None:
                p_def = OxmlElement("w:pPrDefault")
                defaults.append(p_def)
            p_pr = p_def.find(qn("w:pPr"))
            if p_pr is None:
                p_pr = OxmlElement("w:pPr")
                p_def.append(p_pr)
            bidi = OxmlElement("w:bidi")
            bidi.set(qn("w:val"), "1")
            p_pr.append(bidi)

    header = section.header
    header.is_linked_to_previous = False
    hp = header.paragraphs[0]
    hp.text = ""
    set_paragraph_bidi(hp, rtl)
    hp.alignment = WD_ALIGN_PARAGRAPH.RIGHT if rtl else WD_ALIGN_PARAGRAPH.LEFT
    label = "ونڈوز سے میک · M1 رہنمائی" if rtl else "Windows to Mac · M1 user guide"
    add_inline(hp, label, latin, cs, 9, rtl, color=MUTED)
    add_horizontal_line(hp)

    footer = section.footer
    footer.is_linked_to_previous = False
    fp = footer.paragraphs[0]
    fp.text = ""
    set_paragraph_bidi(fp, rtl)
    fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    add_inline(fp, "صفحہ " if rtl else "Page ", latin, cs, 9, rtl, color=MUTED)
    add_page_field(fp, latin, cs, rtl)
    return doc


def add_heading_block(doc, level, text, latin, cs, rtl):
    p = doc.add_paragraph()
    set_paragraph_bidi(p, rtl)
    sizes = {1: 26, 2: 16, 3: 13}
    before = {1: 4, 2: 16, 3: 12}
    p.paragraph_format.space_before = Pt(before[level])
    p.paragraph_format.space_after = Pt(8 if level > 1 else 4)
    add_inline(p, text, latin, cs, sizes[level], rtl, color=NAVY if level <= 2 else INK, bold=True)
    if level == 1:
        sub = doc.add_paragraph()
        set_paragraph_bidi(sub, rtl)
        sub.paragraph_format.space_before = Pt(0)
        sub.paragraph_format.space_after = Pt(10)
        subtitle = (
            "ونڈوز سے Apple silicon میک پر آنے والوں کے لیے مکمل رہنمائی"
            if rtl
            else "A complete guide for switching from Windows to an Apple silicon Mac"
        )
        add_inline(sub, subtitle, latin, cs, 12, rtl, color=ACCENT, italic=True)
        rule = doc.add_paragraph()
        set_paragraph_bidi(rule, rtl)
        rule.paragraph_format.space_before = Pt(0)
        rule.paragraph_format.space_after = Pt(14)
        add_horizontal_line(rule)
    elif level == 2:
        add_horizontal_line(p)


def add_paragraph_block(doc, text, latin, cs, rtl, size):
    p = doc.add_paragraph()
    set_paragraph_bidi(p, rtl)
    p.paragraph_format.space_after = Pt(8)
    add_inline(p, text, latin, cs, size, rtl)


def add_list_block(doc, items, latin, cs, rtl, size, ordered):
    for idx, item in enumerate(items, start=1):
        p = doc.add_paragraph()
        set_paragraph_bidi(p, rtl)
        p.paragraph_format.left_indent = Cm(0.55)
        p.paragraph_format.space_after = Pt(4)
        mark = f"{idx}. " if ordered else "• "
        run = p.add_run(mark)
        set_run_font(run, latin, cs, size, bold=True, color=NAVY)
        add_inline(p, item, latin, cs, size, rtl)


def add_task_block(doc, items, latin, cs, rtl, size):
    for checked, item in items:
        p = doc.add_paragraph()
        set_paragraph_bidi(p, rtl)
        p.paragraph_format.left_indent = Cm(0.5)
        p.paragraph_format.space_after = Pt(4)
        run = p.add_run("☑ " if checked else "☐ ")
        set_run_font(run, latin, cs, size + 1, color=NAVY)
        add_inline(p, item, latin, cs, size, rtl)


def add_table_block(doc, headers, rows, latin, cs, rtl, size):
    cols = len(headers)
    table = doc.add_table(rows=1 + len(rows), cols=cols)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = True
    tbl_pr = table._tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), "4")
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), RULE)
        borders.append(el)
    tbl_pr.append(borders)

    usable = Cm(17.4)
    if cols == 2:
        widths = [usable * 0.38, usable * 0.62]
    elif cols == 3:
        widths = [usable * 0.30, usable * 0.30, usable * 0.40]
    elif cols == 4:
        widths = [usable * 0.22, usable * 0.24, usable * 0.24, usable * 0.30]
    else:
        widths = [usable / cols] * cols

    def fill_cell(cell, text, header=False, alt=False):
        cell.text = ""
        set_cell_margins(cell)
        set_cell_borders(cell)
        if header:
            shade_cell(cell, HEAD_FILL)
        elif alt:
            shade_cell(cell, ALT_FILL)
        p = cell.paragraphs[0]
        set_paragraph_bidi(p, rtl)
        p.paragraph_format.space_after = Pt(0)
        add_inline(
            p,
            text,
            latin,
            cs,
            size - 0.5,
            rtl,
            color=WHITE if header else INK,
            bold=header,
        )

    for c, header in enumerate(headers):
        fill_cell(table.rows[0].cells[c], header, header=True)
        table.rows[0].cells[c].width = widths[c]
    prevent_row_split(table.rows[0])

    for r, row in enumerate(rows):
        prevent_row_split(table.rows[r + 1])
        for c in range(cols):
            value = row[c] if c < len(row) else ""
            fill_cell(table.rows[r + 1].cells[c], value, alt=r % 2 == 1)
            table.rows[r + 1].cells[c].width = widths[c]

    spacer = doc.add_paragraph()
    spacer.paragraph_format.space_after = Pt(10)


def build_docx(md_path: Path, out_path: Path, rtl: bool, title: str) -> Path:
    latin = "Calibri"
    cs = "Arial Unicode MS" if rtl else "Calibri"
    size = 12 if rtl else 11
    doc = configure_document(rtl)
    doc.core_properties.title = title
    doc.core_properties.author = "Windows to Mac Guide"
    doc.core_properties.subject = "Windows to Mac (M1) user guide"

    for block in parse_markdown(md_path.read_text(encoding="utf-8")):
        kind = block[0]
        if kind == "h":
            add_heading_block(doc, block[1], block[2], latin, cs, rtl)
        elif kind == "p":
            add_paragraph_block(doc, block[1], latin, cs, rtl, size)
        elif kind == "ul":
            add_list_block(doc, block[1], latin, cs, rtl, size, ordered=False)
        elif kind == "ol":
            add_list_block(doc, block[1], latin, cs, rtl, size, ordered=True)
        elif kind == "tasks":
            add_task_block(doc, block[1], latin, cs, rtl, size)
        elif kind == "table":
            add_table_block(doc, block[1], block[2], latin, cs, rtl, size)

    out_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(out_path)
    return out_path


def word_to_pdf(docx_path: Path, pdf_path: Path) -> None:
    import win32com.client

    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    try:
        doc = word.Documents.Open(str(docx_path), ReadOnly=True)
        try:
            doc.ExportAsFixedFormat(
                OutputFileName=str(pdf_path),
                ExportFormat=17,
                OpenAfterExport=False,
                OptimizeFor=0,
                BitmapMissingFonts=True,
                DocStructureTags=True,
                CreateBookmarks=1,
            )
        finally:
            doc.Close(False)
    finally:
        word.Quit()
        time.sleep(0.4)


def main() -> int:
    jobs = [
        {
            "md": ROOT / "guide-english.md",
            "docx": ROOT / "Windows-to-Mac-Guide-English.docx",
            "pdf": ROOT / "Windows-to-Mac-Guide-English.pdf",
            "rtl": False,
            "title": "Windows to Mac user guide (M1)",
        },
        {
            "md": ROOT / "guide-urdu.md",
            "docx": ROOT / "Windows-to-Mac-Guide-Urdu.docx",
            "pdf": ROOT / "Windows-to-Mac-Guide-Urdu.pdf",
            "rtl": True,
            "title": "ونڈوز سے میک: مکمل رہنمائی (M1)",
        },
    ]
    for job in jobs:
        print(f"Building {job['docx'].name} ...")
        build_docx(job["md"], job["docx"], job["rtl"], job["title"])
        print(f"Exporting {job['pdf'].name} ...")
        word_to_pdf(job["docx"], job["pdf"])
        print(f"  DOCX {job['docx'].stat().st_size // 1024} KB")
        print(f"  PDF  {job['pdf'].stat().st_size // 1024} KB")
    print("Done.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
