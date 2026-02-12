from __future__ import annotations

import base64
import tempfile
from pathlib import Path

import fitz  # PyMuPDF
from flask import Flask, jsonify, request

from styler_core import apply_style_pdf

app = Flask(__name__)


def _parse_bool(value: str | None, default: bool) -> bool:
    if value is None:
        return default
    return value.strip().lower() not in {"0", "false", "no", "off"}


def _render_pdf_pages_to_base64_png(pdf_path: Path, max_pages: int = 3, max_width_px: int = 1100) -> tuple[int, list[dict]]:
    doc = fitz.open(pdf_path)
    pages = []
    page_count = 0
    try:
        page_count = doc.page_count
        n = min(max_pages, doc.page_count)
        for i in range(n):
            page = doc.load_page(i)
            # Scale to a consistent pixel width.
            rect = page.rect
            scale = max_width_px / max(rect.width, 1)
            mat = fitz.Matrix(scale, scale)
            pix = page.get_pixmap(matrix=mat, alpha=False)
            png_bytes = pix.tobytes("png")
            pages.append(
                {
                    "page": i + 1,
                    "pngBase64": base64.b64encode(png_bytes).decode("ascii"),
                    "width": pix.width,
                    "height": pix.height,
                }
            )
    finally:
        doc.close()
    return (page_count, pages)


@app.route("/", methods=["POST"])
@app.route("/api/preview", methods=["POST"])
def preview():
    doc_file = request.files.get("document")
    if doc_file is None:
        return jsonify({"error": "Missing 'document' upload (.docx required)."}), 400

    template = request.form.get("template", "executive")
    pdf_theme = request.form.get("pdfTheme", "consulting")
    primary_color = request.form.get("primaryColor", "auto")
    text_color = request.form.get("textColor", "auto")
    org_name = request.form.get("orgName", "Your Organization")

    line_spacing = float(request.form.get("lineSpacing", "1.55"))
    reading_width_ch = int(request.form.get("readingWidthCh", "72"))
    include_summary_page = _parse_bool(request.form.get("includeSummaryPage"), True)

    logo_file = request.files.get("logo")

    with tempfile.TemporaryDirectory() as tmp_dir:
        tmp = Path(tmp_dir)
        input_docx = tmp / "input.docx"
        output_pdf = tmp / "preview.pdf"
        logo_path = None

        doc_file.save(str(input_docx))

        if logo_file and logo_file.filename:
            suffix = Path(logo_file.filename).suffix or ".png"
            logo_path = tmp / f"logo{suffix}"
            logo_file.save(str(logo_path))

        try:
            apply_style_pdf(
                input_docx=input_docx,
                output_pdf=output_pdf,
                logo=logo_path,
                org_name=org_name,
                font="auto",
                pdf_theme=pdf_theme,
                template=template,
                primary_color=primary_color,
                text_color=text_color,
                reading_width_ch=reading_width_ch,
                line_spacing=line_spacing,
                include_summary_page=include_summary_page,
            )
        except Exception as exc:  # noqa: BLE001
            return jsonify({"error": str(exc)}), 400

        try:
            page_count, pages = _render_pdf_pages_to_base64_png(output_pdf, max_pages=3, max_width_px=1100)
        except Exception as exc:  # noqa: BLE001
            return jsonify({"error": f"Preview render failed: {exc}"}), 500

        return jsonify(
            {
                "pageCount": page_count,
                "pages": pages,
            }
        )
