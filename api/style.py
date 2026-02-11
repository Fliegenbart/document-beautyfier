from __future__ import annotations

import io
import tempfile
from pathlib import Path

from flask import Flask, jsonify, request, send_file

from styler_core import apply_style

app = Flask(__name__)


def _safe_out_name(name: str | None) -> str:
    if not name:
        return "document_styled.docx"
    cleaned = name.strip().replace("/", "_").replace("\\\\", "_")
    if not cleaned.lower().endswith(".docx"):
        cleaned += ".docx"
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
    font = request.form.get("font", "Calibri")
    download_name = _safe_out_name(request.form.get("outputName"))

    logo_file = request.files.get("logo")

    with tempfile.TemporaryDirectory() as tmp_dir:
        tmp = Path(tmp_dir)
        input_docx = tmp / "input.docx"
        output_docx = tmp / "output.docx"
        logo_path = None

        doc_file.save(str(input_docx))

        if logo_file and logo_file.filename:
            suffix = Path(logo_file.filename).suffix or ".png"
            logo_path = tmp / f"logo{suffix}"
            logo_file.save(str(logo_path))

        try:
            apply_style(
                input_docx=input_docx,
                output_docx=output_docx,
                logo=logo_path,
                org_name=org_name,
                font=font,
                template=template,
                primary_color=primary_color,
                text_color=text_color,
            )
        except Exception as exc:  # noqa: BLE001
            return jsonify({"error": str(exc)}), 400

        payload = io.BytesIO(output_docx.read_bytes())
        payload.seek(0)
        return send_file(
            payload,
            as_attachment=True,
            download_name=download_name,
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )
