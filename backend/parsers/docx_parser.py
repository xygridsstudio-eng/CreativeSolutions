import io
import re

from docx import Document
from docx.oxml.ns import qn
from docx.table import Table
from docx.text.paragraph import Paragraph

HEADING_RE = re.compile(r"^Heading\s*\d+$", re.IGNORECASE)


def _is_heading(paragraph: Paragraph) -> bool:
    name = (paragraph.style.name or "") if paragraph.style else ""
    return bool(HEADING_RE.match(name)) or name.lower() == "title"


def _is_list_item(paragraph: Paragraph) -> bool:
    ppr = paragraph._p.pPr
    return ppr is not None and ppr.numPr is not None


def parse_docx(data: bytes, filename: str) -> dict:
    document = Document(io.BytesIO(data))
    body = document.element.body

    sections = []
    current = {"title": "Document", "page": None, "slide": None, "blocks": []}
    sections.append(current)
    table_counter = 0

    for child in body.iterchildren():
        tag = child.tag

        if tag == qn("w:p"):
            paragraph = Paragraph(child, document)
            text = paragraph.text.strip()
            if not text:
                continue
            if _is_heading(paragraph):
                current = {"title": text, "page": None, "slide": None, "blocks": []}
                sections.append(current)
                current["blocks"].append({"type": "heading", "text": text, "cells": None, "tableId": None, "rowIndex": None})
            elif _is_list_item(paragraph):
                current["blocks"].append({"type": "list", "text": text, "cells": None, "tableId": None, "rowIndex": None})
            else:
                current["blocks"].append({"type": "paragraph", "text": text, "cells": None, "tableId": None, "rowIndex": None})

        elif tag == qn("w:tbl"):
            table_counter += 1
            table_id = f"docx_table_{table_counter}"
            table = Table(child, document)
            for row_index, row in enumerate(table.rows):
                cells = [cell.text.strip() for cell in row.cells]
                if any(cells):
                    current["blocks"].append({
                        "type": "tableRow",
                        "text": " | ".join(cells),
                        "cells": cells,
                        "tableId": table_id,
                        "rowIndex": row_index,
                    })

    non_empty = [s for s in sections if s["blocks"]]
    return {"docType": "docx", "fileName": filename, "sections": non_empty or sections}
