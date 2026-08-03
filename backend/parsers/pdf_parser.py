import io

import pdfplumber


def _in_any_bbox(obj: dict, bboxes: list[tuple]) -> bool:
    for (x0, top, x1, bottom) in bboxes:
        if obj["x0"] >= x0 - 1 and obj["x1"] <= x1 + 1 and obj["top"] >= top - 1 and obj["bottom"] <= bottom + 1:
            return True
    return False


def _group_words_into_paragraphs(words: list[dict]) -> list[str]:
    """Group words into lines by rounded y-position, then lines into
    paragraphs on a large vertical gap — same heuristic as the JS PDF
    parser, but starting from pdfplumber's word extraction (tolerance-based
    clustering) rather than raw text-item coordinates.
    """
    lines: dict[int, list[dict]] = {}
    for w in words:
        y = round(w["top"])
        lines.setdefault(y, []).append(w)

    paragraphs = []
    buffer: list[str] = []
    last_bottom = None
    for y in sorted(lines.keys()):
        line_words = sorted(lines[y], key=lambda w: w["x0"])
        line_text = " ".join(w["text"] for w in line_words).strip()
        if not line_text:
            continue
        top = min(w["top"] for w in line_words)
        bottom = max(w["bottom"] for w in line_words)
        if last_bottom is not None and top - last_bottom > 8:
            if buffer:
                paragraphs.append(" ".join(buffer).strip())
            buffer = []
        buffer.append(line_text)
        last_bottom = bottom
    if buffer:
        paragraphs.append(" ".join(buffer).strip())
    return [p for p in paragraphs if p]


def parse_pdf(data: bytes, filename: str) -> dict:
    sections = []
    with pdfplumber.open(io.BytesIO(data)) as pdf:
        for page_num, page in enumerate(pdf.pages, start=1):
            section = {"title": f"Page {page_num}", "page": page_num, "slide": None, "blocks": []}

            tables = page.find_tables()
            table_bboxes = [t.bbox for t in tables]

            for t_idx, table in enumerate(tables, start=1):
                table_id = f"page{page_num}_table_{t_idx}"
                for row_index, row in enumerate(table.extract()):
                    cells = [(c or "").strip() for c in row]
                    if any(cells):
                        section["blocks"].append({
                            "type": "tableRow",
                            "text": " | ".join(cells),
                            "cells": cells,
                            "tableId": table_id,
                            "rowIndex": row_index,
                        })

            # Non-table words only, so paragraph text and table text never overlap.
            non_table_words = [w for w in page.extract_words() if not _in_any_bbox(w, table_bboxes)]
            for text in _group_words_into_paragraphs(non_table_words):
                section["blocks"].append({"type": "paragraph", "text": text, "cells": None, "tableId": None, "rowIndex": None})

            sections.append(section)

    return {"docType": "pdf", "fileName": filename, "sections": sections}
