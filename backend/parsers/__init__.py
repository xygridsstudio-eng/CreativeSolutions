"""
Format-specific parsers. Every parser takes raw file bytes and a filename
and returns the same shape apps/contentcheck/js/parser.js's CCParser.parseFile
already produces client-side, so the frontend pipeline (normalizer -> comparer
-> report -> excel) needs zero changes to consume backend-parsed documents:

{
  "docType": "docx" | "pptx" | "pdf" | "txt",
  "fileName": str,
  "sections": [
    {
      "title": str,
      "page": int | None,
      "slide": int | None,
      "blocks": [
        {"type": "heading"|"paragraph"|"list"|"notes"|"tableRow",
         "text": str, "cells": list[str] | None,
         "tableId": str | None, "rowIndex": int | None}
      ]
    }
  ]
}
"""
from .docx_parser import parse_docx
from .pptx_parser import parse_pptx
from .pdf_parser import parse_pdf
from .txt_parser import parse_txt

PARSERS = {
    "docx": parse_docx,
    "pptx": parse_pptx,
    "pdf": parse_pdf,
    "txt": parse_txt,
}


def parse_file(ext: str, data: bytes, filename: str) -> dict:
    parser = PARSERS.get(ext)
    if parser is None:
        raise ValueError(f"Unsupported file type: .{ext}. Supported types: docx, pptx, pdf, txt.")
    return parser(data, filename)
