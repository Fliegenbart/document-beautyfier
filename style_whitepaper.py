#!/usr/bin/env python3
from __future__ import annotations

import argparse
from pathlib import Path

from styler_core import TEMPLATES, apply_style, apply_style_pdf


def main() -> None:
    parser = argparse.ArgumentParser(description="Apply configurable business styling to DOCX whitepapers.")
    parser.add_argument("input_docx", type=Path)
    parser.add_argument("output_file", type=Path, help="Output .docx or .pdf file")
    parser.add_argument("--logo", type=Path, default=None, help="Path to logo image (png/jpg).")
    parser.add_argument("--org-name", type=str, default="Your Organization")
    parser.add_argument("--font", type=str, default="Helvetica")
    parser.add_argument("--template", choices=sorted(TEMPLATES.keys()), default="executive")
    parser.add_argument("--primary-color", type=str, default="#F50000", help="Hex or RGB string, e.g. #F50000 or 245,0,0")
    parser.add_argument("--text-color", type=str, default="#111111", help="Hex or RGB string, e.g. #111111 or 17,17,17")
    parser.add_argument("--output-format", choices=["docx", "pdf"], default=None)
    args = parser.parse_args()

    output_format = args.output_format
    if output_format is None:
        output_format = "pdf" if args.output_file.suffix.lower() == ".pdf" else "docx"

    if output_format == "pdf":
        apply_style_pdf(
            input_docx=args.input_docx,
            output_pdf=args.output_file,
            logo=args.logo,
            org_name=args.org_name,
            font=args.font,
            template=args.template,
            primary_color=args.primary_color,
            text_color=args.text_color,
        )
    else:
        apply_style(
            input_docx=args.input_docx,
            output_docx=args.output_file,
            logo=args.logo,
            org_name=args.org_name,
            font=args.font,
            template=args.template,
            primary_color=args.primary_color,
            text_color=args.text_color,
        )


if __name__ == "__main__":
    main()
