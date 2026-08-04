import io
import re

from lxml import etree
from pptx import Presentation

NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "dgm": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}


def _collapse(text: str) -> str:
    return re.sub(r"\s+", " ", text or "").strip()


def _paragraph_texts_from_shape(shape) -> list[str]:
    """One entry per paragraph, matching python-docx/pptx's own paragraph
    boundaries — NOT joined across paragraphs. This mirrors the fix in
    apps/contentcheck/js/parser.js where concatenating every run in a shape
    into one string glued unrelated bullets together (e.g. "VF"+"XYZ"+"XYZ"
    -> "VFXYZXYZ"), making them impossible to match against the other doc.
    """
    texts = []
    for paragraph in shape.text_frame.paragraphs:
        text = _collapse(paragraph.text)
        if text:
            texts.append(text)
    return texts


def _extract_diagram_texts(shape) -> list[str]:
    """SmartArt diagrams aren't exposed by python-pptx at all — the actual
    text lives in a separate ppt/diagrams/dataN.xml part referenced by a
    relationship id on the graphicFrame. Without this, content redesigned
    into a SmartArt diagram would silently disappear from parsing, same gap
    that existed in the JS parser before extractDiagramTexts was added.
    """
    graphic_data = shape._element.find(f".//{{{NS['a']}}}graphicData")
    if graphic_data is None or "diagram" not in (graphic_data.get("uri") or ""):
        return []
    rel_ids_el = graphic_data.find(f".//{{{NS['dgm']}}}relIds")
    if rel_ids_el is None:
        return []
    dm_rel_id = rel_ids_el.get(f"{{{NS['r']}}}dm")
    if not dm_rel_id:
        return []

    slide_part = shape.part
    if dm_rel_id not in slide_part.rels:
        return []
    diagram_blob = slide_part.rels[dm_rel_id].target_part.blob
    diagram_xml = etree.fromstring(diagram_blob)

    texts = []
    for pt in diagram_xml.iter(f"{{{NS['dgm']}}}pt"):
        pt_type = pt.get("type")
        if pt_type and pt_type != "node":
            continue  # parTrans/sibTrans/doc are structural connectors, not content
        for p in pt.iter(f"{{{NS['a']}}}p"):
            text = _collapse("".join(t.text or "" for t in p.iter(f"{{{NS['a']}}}t")))
            if text:
                texts.append(text)
    return texts


def _extract_chart_texts(chart) -> list[str]:
    """Unlike SmartArt, python-pptx has a proper Chart API — no need to hand
    -walk the chart XML part. Without this, edits to a chart's underlying
    numbers (the whole point of a chart) are completely invisible to the
    comparison, same gap that existed in the JS parser before
    extractChartTexts was added there.
    """
    texts = []
    if chart.has_title and chart.chart_title.text_frame is not None:
        title = _collapse(chart.chart_title.text_frame.text)
        if title:
            texts.append(f"Chart title: {title}")

    try:
        categories = [str(c) for c in list(chart.plots[0].categories)] if chart.plots else []
    except (IndexError, ValueError):
        categories = []

    def fmt_value(v):
        if v is None:
            return ""
        # python-pptx always hands back floats; the JS parser reads the raw
        # cached string from the chart XML (e.g. "5000", not "5000.0") — match
        # that so the same chart parses to identical text via either path.
        return str(int(v)) if float(v).is_integer() else str(v)

    for series in chart.series:
        pairs = [
            f"{categories[i] if i < len(categories) else f'#{i}'}: {fmt_value(v)}"
            for i, v in enumerate(series.values)
        ]
        line = (f"{series.name} — " if series.name else "") + ", ".join(pairs)
        line = line.strip()
        if line:
            texts.append(line)
    return texts


def parse_pptx(data: bytes, filename: str) -> dict:
    prs = Presentation(io.BytesIO(data))
    sections = []

    for slide_num, slide in enumerate(prs.slides, start=1):
        section = {"title": f"Slide {slide_num}", "page": None, "slide": slide_num, "blocks": []}
        title_shape = slide.shapes.title
        table_counter = 0

        for shape in slide.shapes:
            if shape.has_text_frame:
                para_texts = _paragraph_texts_from_shape(shape)
                if para_texts:
                    if shape is title_shape:
                        text = " ".join(para_texts)
                        section["title"] = text
                        section["blocks"].append({"type": "heading", "text": text, "cells": None, "tableId": None, "rowIndex": None})
                    else:
                        block_type = "list" if len(para_texts) > 1 else "paragraph"
                        for text in para_texts:
                            section["blocks"].append({"type": block_type, "text": text, "cells": None, "tableId": None, "rowIndex": None})

            if shape.has_table:
                table_counter += 1
                table_id = f"slide{slide_num}_table_{table_counter}"
                for row_index, row in enumerate(shape.table.rows):
                    cells = [_collapse(cell.text) for cell in row.cells]
                    if any(cells):
                        section["blocks"].append({
                            "type": "tableRow",
                            "text": " | ".join(cells),
                            "cells": cells,
                            "tableId": table_id,
                            "rowIndex": row_index,
                        })

            # python-pptx's shape_type is documented to be None for SmartArt
            # ("This value is None when none of these [four] types apply, for
            # example when the shape contains SmartArt") — there's no reliable
            # shape_type to gate on, so just attempt extraction on every shape;
            # _extract_diagram_texts itself checks graphicData's uri and is a
            # no-op (returns []) for anything that isn't actually a diagram.
            for text in _extract_diagram_texts(shape):
                section["blocks"].append({"type": "list", "text": text, "cells": None, "tableId": None, "rowIndex": None})

            if shape.has_chart:
                for text in _extract_chart_texts(shape.chart):
                    section["blocks"].append({"type": "list", "text": text, "cells": None, "tableId": None, "rowIndex": None})

        if slide.has_notes_slide and slide.notes_slide.notes_text_frame is not None:
            for paragraph in slide.notes_slide.notes_text_frame.paragraphs:
                text = _collapse(paragraph.text)
                if text:
                    section["blocks"].append({"type": "notes", "text": text, "cells": None, "tableId": None, "rowIndex": None})

        sections.append(section)

    return {"docType": "pptx", "fileName": filename, "sections": sections}
