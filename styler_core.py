from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Iterator

from docx import Document
from docx.document import Document as DocxDocumentType
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.shared import Cm, Pt, RGBColor
from docx.table import Table as DocxTable
from docx.text.paragraph import Paragraph as DocxParagraph
from reportlab.lib import colors
from reportlab.lib.colors import HexColor
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.platypus import Image, PageBreak, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle


@dataclass
class TemplateConfig:
    margin_cm: float
    body_size: float
    h1_size: float
    h2_size: float
    h3_size: float
    title_size: float
    line_spacing: float
    header_line_size: str
    logo_width_cm: float


TEMPLATES = {
    "minimal": TemplateConfig(2.4, 10.8, 19, 14, 12.5, 25, 1.13, "8", 4.8),
    "executive": TemplateConfig(2.1, 11, 20, 15, 13, 28, 1.15, "12", 5.5),
    "bold": TemplateConfig(1.8, 11.2, 22, 16, 14, 30, 1.2, "16", 6.0),
}


def normalize_color_string(value: str) -> str:
    text = value.strip().lower().replace(" ", "")
    if text.startswith("#"):
        text = text[1:]
    if "," in text:
        parts = text.split(",")
        if len(parts) != 3:
            raise ValueError(f"Invalid RGB color: {value}")
        rgb = [max(0, min(255, int(p))) for p in parts]
        return f"{rgb[0]:02X}{rgb[1]:02X}{rgb[2]:02X}"
    if len(text) == 6 and all(c in "0123456789abcdef" for c in text):
        return text.upper()
    raise ValueError(f"Invalid color: {value} (use #F50000 or 245,0,0)")


def hex_to_rgb(hex_color: str) -> RGBColor:
    return RGBColor(int(hex_color[0:2], 16), int(hex_color[2:4], 16), int(hex_color[4:6], 16))


def set_cell_shading(cell, fill_hex: str) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill_hex)
    tc_pr.append(shd)


def set_paragraph_bottom_border(paragraph, color_hex: str, size: str) -> None:
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


def style_named_styles(doc: Document, font_name: str, text_rgb: RGBColor, primary_rgb: RGBColor, cfg: TemplateConfig) -> None:
    normal = doc.styles["Normal"]
    normal_font = normal.font
    normal_font.name = font_name
    normal_font.size = Pt(cfg.body_size)
    normal_font.color.rgb = text_rgb

    heading_specs = {
        "Heading 1": (Pt(cfg.h1_size), True),
        "Heading 2": (Pt(cfg.h2_size), True),
        "Heading 3": (Pt(cfg.h3_size), True),
    }

    for style_name, (size, bold) in heading_specs.items():
        if style_name in doc.styles:
            st = doc.styles[style_name]
            st_font = st.font
            st_font.name = font_name
            st_font.size = size
            st_font.bold = bold
            st_font.color.rgb = primary_rgb

    if "Title" in doc.styles:
        title = doc.styles["Title"]
        title_font = title.font
        title_font.name = font_name
        title_font.size = Pt(cfg.title_size)
        title_font.bold = True
        title_font.color.rgb = primary_rgb


def style_paragraphs(doc: Document, font_name: str, text_rgb: RGBColor, primary_rgb: RGBColor, cfg: TemplateConfig) -> None:
    first_content_paragraph = None

    for p in doc.paragraphs:
        text = (p.text or "").strip()
        pf = p.paragraph_format
        pf.space_after = Pt(8)
        pf.space_before = Pt(0)
        pf.line_spacing = cfg.line_spacing

        if not text:
            continue

        if first_content_paragraph is None:
            first_content_paragraph = p

        for r in p.runs:
            if not r.font.name:
                r.font.name = font_name
            if not r.font.size:
                r.font.size = Pt(cfg.body_size)
            if r.font.color is None or r.font.color.rgb is None:
                r.font.color.rgb = text_rgb

        style_name = p.style.name if p.style else ""
        if style_name.startswith("Heading"):
            for r in p.runs:
                r.font.color.rgb = primary_rgb

    if first_content_paragraph is not None:
        style_name = first_content_paragraph.style.name if first_content_paragraph.style else ""
        if style_name not in {"Title", "Heading 1"}:
            first_content_paragraph.style = doc.styles["Title"] if "Title" in doc.styles else doc.styles["Heading 1"]
            first_content_paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT


def style_tables(doc: Document, font_name: str, text_rgb: RGBColor, primary_hex: str) -> None:
    white = RGBColor(255, 255, 255)
    for table in doc.tables:
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        if not table.rows:
            continue

        header_row = table.rows[0]
        for cell in header_row.cells:
            set_cell_shading(cell, primary_hex)
            for p in cell.paragraphs:
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                for r in p.runs:
                    r.font.bold = True
                    r.font.color.rgb = white
                    r.font.name = font_name
                    r.font.size = Pt(10.5)

        for row in table.rows[1:]:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for r in p.runs:
                        if r.font.color is None or r.font.color.rgb is None:
                            r.font.color.rgb = text_rgb
                        if not r.font.name:
                            r.font.name = font_name
                        if not r.font.size:
                            r.font.size = Pt(10.5)


def style_sections_and_header(
    doc: Document,
    logo_path: Path | None,
    org_name: str,
    font_name: str,
    text_rgb: RGBColor,
    primary_hex: str,
    cfg: TemplateConfig,
) -> None:
    for section in doc.sections:
        section.top_margin = Cm(cfg.margin_cm)
        section.bottom_margin = Cm(cfg.margin_cm)
        section.left_margin = Cm(cfg.margin_cm)
        section.right_margin = Cm(cfg.margin_cm)

        section.different_first_page_header_footer = True

        for header in [section.header, section.first_page_header]:
            clear_header_paragraphs(header)
            p = header.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.space_after = Pt(6)

            if logo_path and logo_path.exists():
                run = p.add_run()
                run.add_picture(str(logo_path), width=Cm(cfg.logo_width_cm))
            else:
                run = p.add_run(org_name)
                run.font.name = font_name
                run.font.size = Pt(cfg.body_size)
                run.font.bold = True
                run.font.color.rgb = text_rgb

            set_paragraph_bottom_border(p, color_hex=primary_hex, size=cfg.header_line_size)


def ensure_custom_styles(doc: Document, font_name: str, primary_rgb: RGBColor, body_size: float) -> None:
    if "Beautifier Accent" not in [s.name for s in doc.styles]:
        accent = doc.styles.add_style("Beautifier Accent", WD_STYLE_TYPE.PARAGRAPH)
        accent.base_style = doc.styles["Normal"]
        accent.font.name = font_name
        accent.font.size = Pt(body_size)
        accent.font.bold = True
        accent.font.color.rgb = primary_rgb


def apply_style(
    input_docx: Path,
    output_docx: Path,
    logo: Path | None = None,
    org_name: str = "Your Organization",
    font: str = "Calibri",
    template: str = "executive",
    primary_color: str = "#F50000",
    text_color: str = "#111111",
) -> None:
    if template not in TEMPLATES:
        raise ValueError(f"Unsupported template: {template}")

    primary_hex = normalize_color_string(primary_color)
    text_hex = normalize_color_string(text_color)
    primary_rgb = hex_to_rgb(primary_hex)
    text_rgb = hex_to_rgb(text_hex)
    cfg = TEMPLATES[template]

    doc = Document(str(input_docx))

    style_named_styles(doc, font, text_rgb, primary_rgb, cfg)
    ensure_custom_styles(doc, font, primary_rgb, cfg.body_size)
    style_paragraphs(doc, font, text_rgb, primary_rgb, cfg)
    style_tables(doc, font, text_rgb, primary_hex)
    style_sections_and_header(doc, logo, org_name, font, text_rgb, primary_hex, cfg)

    output_docx.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(output_docx))


def _iter_block_items(parent: DocxDocumentType) -> Iterator[DocxParagraph | DocxTable]:
    for child in parent.element.body.iterchildren():
        if isinstance(child, CT_P):
            yield DocxParagraph(child, parent)
        elif isinstance(child, CT_Tbl):
            yield DocxTable(child, parent)


def _extract_doc_title(doc: Document) -> str:
    for p in doc.paragraphs:
        text = (p.text or "").strip()
        if text:
            return text
    return "Styled Document"


def _build_pdf_styles(font: str, primary_hex: str, text_hex: str, cfg: TemplateConfig):
    primary = HexColor(f"#{primary_hex}")
    text = HexColor(f"#{text_hex}")
    sheet = getSampleStyleSheet()
    return {
        "h1": ParagraphStyle(
            "h1",
            parent=sheet["Heading1"],
            fontName=font,
            fontSize=cfg.h1_size,
            leading=cfg.h1_size + 4,
            textColor=primary,
            spaceBefore=8,
            spaceAfter=8,
        ),
        "h2": ParagraphStyle(
            "h2",
            parent=sheet["Heading2"],
            fontName=font,
            fontSize=cfg.h2_size,
            leading=cfg.h2_size + 3,
            textColor=primary,
            spaceBefore=7,
            spaceAfter=6,
        ),
        "h3": ParagraphStyle(
            "h3",
            parent=sheet["Heading3"],
            fontName=font,
            fontSize=cfg.h3_size,
            leading=cfg.h3_size + 3,
            textColor=primary,
            spaceBefore=6,
            spaceAfter=5,
        ),
        "body": ParagraphStyle(
            "body",
            parent=sheet["BodyText"],
            fontName=font,
            fontSize=cfg.body_size,
            leading=cfg.body_size * 1.45,
            textColor=text,
            spaceAfter=5,
        ),
        "cover_title": ParagraphStyle(
            "cover_title",
            parent=sheet["Title"],
            fontName=font,
            fontSize=cfg.title_size + 4,
            leading=cfg.title_size + 8,
            textColor=primary,
            alignment=1,
            spaceAfter=12,
        ),
        "cover_sub": ParagraphStyle(
            "cover_sub",
            parent=sheet["Normal"],
            fontName=font,
            fontSize=cfg.body_size + 1,
            textColor=text,
            alignment=1,
        ),
        "footer": ParagraphStyle(
            "footer",
            parent=sheet["Normal"],
            fontName=font,
            fontSize=8.5,
            textColor=colors.HexColor("#666666"),
        ),
    }


def apply_style_pdf(
    input_docx: Path,
    output_pdf: Path,
    logo: Path | None = None,
    org_name: str = "Your Organization",
    font: str = "Helvetica",
    template: str = "executive",
    primary_color: str = "#F50000",
    text_color: str = "#111111",
) -> None:
    if template not in TEMPLATES:
        raise ValueError(f"Unsupported template: {template}")

    primary_hex = normalize_color_string(primary_color)
    text_hex = normalize_color_string(text_color)
    cfg = TEMPLATES[template]

    doc = Document(str(input_docx))
    title = _extract_doc_title(doc)
    styles = _build_pdf_styles(font, primary_hex, text_hex, cfg)

    output_pdf.parent.mkdir(parents=True, exist_ok=True)

    margin = cfg.margin_cm * cm
    rl_doc = SimpleDocTemplate(
        str(output_pdf),
        pagesize=A4,
        leftMargin=margin,
        rightMargin=margin,
        topMargin=margin,
        bottomMargin=margin,
    )

    primary = HexColor(f"#{primary_hex}")
    story = []

    # Cover
    story.append(Spacer(1, 1.4 * cm))
    if logo and logo.exists():
        story.append(Image(str(logo), width=cfg.logo_width_cm * cm, preserveAspectRatio=True, hAlign="CENTER"))
        story.append(Spacer(1, 0.8 * cm))
    story.append(Paragraph(title, styles["cover_title"]))
    story.append(Paragraph(org_name, styles["cover_sub"]))
    story.append(Spacer(1, 0.8 * cm))
    story.append(
        Table(
            [[" "]],
            colWidths=[rl_doc.width],
            rowHeights=[0.25 * cm],
            style=[("BACKGROUND", (0, 0), (-1, -1), primary), ("LINEBELOW", (0, 0), (-1, -1), 0, primary)],
        )
    )
    story.append(PageBreak())

    # Content
    for item in _iter_block_items(doc):
        if isinstance(item, DocxParagraph):
            text = (item.text or "").strip()
            if not text:
                continue
            style_name = (item.style.name or "").lower()
            if "heading 1" in style_name:
                story.append(Paragraph(text, styles["h1"]))
            elif "heading 2" in style_name:
                story.append(Paragraph(text, styles["h2"]))
            elif "heading 3" in style_name:
                story.append(Paragraph(text, styles["h3"]))
            else:
                story.append(Paragraph(text.replace("\n", "<br/>"), styles["body"]))
        else:
            rows = []
            for row in item.rows:
                rows.append([" ".join(cell.text.split()) for cell in row.cells])
            if not rows:
                continue
            col_count = max(len(r) for r in rows)
            table = Table(rows, colWidths=[rl_doc.width / col_count] * col_count, repeatRows=1)
            table.setStyle(
                TableStyle(
                    [
                        ("BACKGROUND", (0, 0), (-1, 0), primary),
                        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                        ("FONTNAME", (0, 0), (-1, -1), font),
                        ("FONTSIZE", (0, 0), (-1, 0), 10),
                        ("FONTSIZE", (0, 1), (-1, -1), 9),
                        ("GRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#BBBBBB")),
                        ("ALIGN", (0, 0), (-1, 0), "LEFT"),
                        ("VALIGN", (0, 0), (-1, -1), "TOP"),
                        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.HexColor("#FAFAFA"), colors.white]),
                        ("LEFTPADDING", (0, 0), (-1, -1), 6),
                        ("RIGHTPADDING", (0, 0), (-1, -1), 6),
                        ("TOPPADDING", (0, 0), (-1, -1), 4),
                        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
                    ]
                )
            )
            story.append(Spacer(1, 0.2 * cm))
            story.append(table)
            story.append(Spacer(1, 0.4 * cm))

    def draw_frame(canvas, _doc):
        canvas.saveState()
        canvas.setStrokeColor(primary)
        canvas.setLineWidth(1)
        canvas.line(_doc.leftMargin, A4[1] - _doc.topMargin + 8, A4[0] - _doc.rightMargin, A4[1] - _doc.topMargin + 8)
        canvas.setFont("Helvetica", 8)
        canvas.setFillColor(colors.HexColor("#666666"))
        canvas.drawString(_doc.leftMargin, _doc.bottomMargin - 16, org_name)
        canvas.drawRightString(A4[0] - _doc.rightMargin, _doc.bottomMargin - 16, f"Page {canvas.getPageNumber()}")
        canvas.restoreState()

    rl_doc.build(story, onFirstPage=draw_frame, onLaterPages=draw_frame)
