#!/usr/bin/env python3
from __future__ import annotations

import argparse
from pathlib import Path

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

RED = RGBColor(245, 0, 0)
BLACK = RGBColor(0, 0, 0)
WHITE = RGBColor(255, 255, 255)


def set_cell_shading(cell, fill_hex: str) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill_hex)
    tc_pr.append(shd)


def set_paragraph_bottom_border(paragraph, color_hex: str = "F50000", size: str = "8") -> None:
    p = paragraph._p
    p_pr = p.get_or_add_pPr()
    p_bdr = p_pr.find(qn("w:pBdr"))
    if p_bdr is None:
        p_bdr = OxmlElement("w:pBdr")
        p_pr.append(p_bdr)

    bottom = p_bdr.find(qn("w:bottom"))
    if bottom is None:
        bottom = OxmlElement("w:bottom")
        p_bdr.append(bottom)

    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), size)
    bottom.set(qn("w:space"), "1")
    bottom.set(qn("w:color"), color_hex)


def clear_header_paragraphs(header) -> None:
    for p in list(header.paragraphs):
        p._element.getparent().remove(p._element)


def style_named_styles(doc: Document) -> None:
    normal = doc.styles["Normal"]
    normal_font = normal.font
    normal_font.name = "Calibri"
    normal_font.size = Pt(11)
    normal_font.color.rgb = BLACK

    heading_specs = {
        "Heading 1": (Pt(20), True),
        "Heading 2": (Pt(15), True),
        "Heading 3": (Pt(13), True),
    }

    for style_name, (size, bold) in heading_specs.items():
        if style_name in doc.styles:
            st = doc.styles[style_name]
            st_font = st.font
            st_font.name = "Calibri"
            st_font.size = size
            st_font.bold = bold
            st_font.color.rgb = RED

    if "Title" in doc.styles:
        title = doc.styles["Title"]
        title_font = title.font
        title_font.name = "Calibri"
        title_font.size = Pt(28)
        title_font.bold = True
        title_font.color.rgb = RED


def style_paragraphs(doc: Document) -> None:
    first_content_paragraph = None

    for p in doc.paragraphs:
        text = (p.text or "").strip()
        pf = p.paragraph_format
        pf.space_after = Pt(8)
        pf.space_before = Pt(0)
        pf.line_spacing = 1.15

        if not text:
            continue

        if first_content_paragraph is None:
            first_content_paragraph = p

        # Enforce clean black body text by default
        for r in p.runs:
            if not r.font.name:
                r.font.name = "Calibri"
            if not r.font.size:
                r.font.size = Pt(11)
            if r.font.color is None or r.font.color.rgb is None:
                r.font.color.rgb = BLACK

        # Heuristic: keep headings consistently red
        style_name = p.style.name if p.style else ""
        if style_name.startswith("Heading"):
            for r in p.runs:
                r.font.color.rgb = RED

    # First non-empty paragraph as title if not already a heading/title
    if first_content_paragraph is not None:
        style_name = first_content_paragraph.style.name if first_content_paragraph.style else ""
        if style_name not in {"Title", "Heading 1"}:
            first_content_paragraph.style = doc.styles["Title"] if "Title" in doc.styles else doc.styles["Heading 1"]
            first_content_paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT


def style_tables(doc: Document) -> None:
    for table in doc.tables:
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        if not table.rows:
            continue

        # Header row in corporate red with white text
        header_row = table.rows[0]
        for cell in header_row.cells:
            set_cell_shading(cell, "F50000")
            for p in cell.paragraphs:
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                for r in p.runs:
                    r.font.bold = True
                    r.font.color.rgb = WHITE
                    r.font.name = "Calibri"
                    r.font.size = Pt(10.5)

        # Body rows in neutral business style
        for row in table.rows[1:]:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for r in p.runs:
                        if r.font.color is None or r.font.color.rgb is None:
                            r.font.color.rgb = BLACK
                        if not r.font.name:
                            r.font.name = "Calibri"
                        if not r.font.size:
                            r.font.size = Pt(10.5)


def style_sections_and_header(doc: Document, logo_path: Path | None, org_name: str) -> None:
    for section in doc.sections:
        section.top_margin = Cm(2.1)
        section.bottom_margin = Cm(2.1)
        section.left_margin = Cm(2.1)
        section.right_margin = Cm(2.1)

        section.different_first_page_header_footer = True

        for header in [section.header, section.first_page_header]:
            clear_header_paragraphs(header)
            p = header.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.space_after = Pt(6)

            if logo_path and logo_path.exists():
                run = p.add_run()
                run.add_picture(str(logo_path), width=Cm(5.5))
            else:
                run = p.add_run(org_name)
                run.font.name = "Calibri"
                run.font.size = Pt(11)
                run.font.bold = True
                run.font.color.rgb = BLACK

            set_paragraph_bottom_border(p, color_hex="F50000", size="12")



def ensure_custom_styles(doc: Document) -> None:
    # Adds a dedicated style that can be used by follow-up automation.
    if "GW Accent" not in [s.name for s in doc.styles]:
        accent = doc.styles.add_style("GW Accent", WD_STYLE_TYPE.PARAGRAPH)
        accent.base_style = doc.styles["Normal"]
        accent.font.name = "Calibri"
        accent.font.size = Pt(11)
        accent.font.bold = True
        accent.font.color.rgb = RED


def main() -> None:
    parser = argparse.ArgumentParser(description="Apply Gruenewald business styling to DOCX whitepapers.")
    parser.add_argument("input_docx", type=Path)
    parser.add_argument("output_docx", type=Path)
    parser.add_argument("--logo", type=Path, default=None, help="Path to logo image (png/jpg).")
    parser.add_argument("--org-name", type=str, default="GRUENEWALD GmbH")
    args = parser.parse_args()

    doc = Document(str(args.input_docx))

    style_named_styles(doc)
    ensure_custom_styles(doc)
    style_paragraphs(doc)
    style_tables(doc)
    style_sections_and_header(doc, args.logo, args.org_name)

    args.output_docx.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(args.output_docx))


if __name__ == "__main__":
    main()
