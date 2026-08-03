import re


def parse_txt(data: bytes, filename: str) -> dict:
    text = data.decode("utf-8", errors="replace")
    raw_paragraphs = re.split(r"\r?\n\s*\r?\n", text)  # blank-line separated
    blocks = []
    for p in raw_paragraphs:
        cleaned = re.sub(r"\r?\n", " ", p).strip()
        if cleaned:
            blocks.append({"type": "paragraph", "text": cleaned, "cells": None, "tableId": None, "rowIndex": None})

    section = {"title": "Document", "page": None, "slide": None, "blocks": blocks}
    return {"docType": "txt", "fileName": filename, "sections": [section]}
