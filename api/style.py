from __future__ import annotations

import io
import tempfile
from pathlib import Path

from flask import Flask, jsonify, request, send_file

from styler_core import apply_style, apply_style_pdf

app = Flask(__name__)


def _safe_out_name(name: str | None, output_format: str) -> str:
    default = "document_styled.pdf" if output_format == "pdf" else "document_styled.docx"
    if not name:
        return default

    cleaned = name.strip().replace("/", "_").replace("\\", "_")
    ext = ".pdf" if output_format == "pdf" else ".docx"
    if not cleaned.lower().endswith(ext):
        cleaned += ext
    return cleaned


@app.route("/", methods=["POST"])
@app.route("/api/style", methods=["POST"])
def style_document():
    doc_file = request.files.get("document")
    if doc_file is None:
        return jsonify({"error": "Missing 'document' upload (.docx required)."}), 400

    template = request.form.get("template", "executive")
    primary_color = request.form.get("primaryColor", "#F50000")
    text_color = request.form.get("textColor", "#111111")
    org_name = request.form.get("orgName", "Your Organization")
    font = request.form.get("font", "auto")
    line_spacing = float(request.form.get("lineSpacing", "1.55"))
    reading_width_ch = int(request.form.get("readingWidthCh", "72"))
    include_summary_page = request.form.get("includeSummaryPage", "true").lower() != "false"
    output_format = request.form.get("outputFormat", "docx").lower().strip()
    pdf_theme = request.form.get("pdfTheme", "consulting")
    if output_format not in {"docx", "pdf"}:
        return jsonify({"error": "outputFormat must be 'docx' or 'pdf'."}), 400

    download_name = _safe_out_name(request.form.get("outputName"), output_format)
    logo_file = request.files.get("logo")

    with tempfile.TemporaryDirectory() as tmp_dir:
        tmp = Path(tmp_dir)
        input_docx = tmp / "input.docx"
        logo_path = None
        output_file = tmp / ("output.pdf" if output_format == "pdf" else "output.docx")

        doc_file.save(str(input_docx))

        if logo_file and logo_file.filename:
            suffix = Path(logo_file.filename).suffix or ".png"
            logo_path = tmp / f"logo{suffix}"
            logo_file.save(str(logo_path))

        try:
            if output_format == "pdf":
                apply_style_pdf(
                    input_docx=input_docx,
                    output_pdf=output_file,
                    logo=logo_path,
                    org_name=org_name,
                    font=font,
                    pdf_theme=pdf_theme,
                    template=template,
                    primary_color=primary_color,
                    text_color=text_color,
                    reading_width_ch=reading_width_ch,
                    line_spacing=line_spacing,
                    include_summary_page=include_summary_page,
                )
            else:
                apply_style(
                    input_docx=input_docx,
                    output_docx=output_file,
                    logo=logo_path,
                    org_name=org_name,
                    font=font,
                    template=template,
                    primary_color=primary_color,
                    text_color=text_color,
                    line_spacing=line_spacing,
                )
        except Exception as exc:  # noqa: BLE001
            return jsonify({"error": str(exc)}), 400

        payload = io.BytesIO(output_file.read_bytes())
        payload.seek(0)
        mimetype = (
            "application/pdf"
            if output_format == "pdf"
            else "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
        return send_file(payload, as_attachment=True, download_name=download_name, mimetype=mimetype)
