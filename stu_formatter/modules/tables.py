"""Table formatting: fonts, header rows, captions, borders."""

import logging
import re

from docx import Document
from docx.shared import Pt, Mm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from lxml import etree

from ..config import STUConfig

logger = logging.getLogger(__name__)

# Pattern for table captions: "Таблица 1 – Название" or "Таблица 1.1 – Название"
TABLE_CAPTION_RE = re.compile(
    r"^Таблица\s+(\d+(?:\.\d+)?|[А-Я]\.\d+)\s*[–\-—]\s*(.+)",
    re.IGNORECASE,
)

TABLE_CONTINUATION_RE = re.compile(
    r"^Продолжение\s+таблицы\s+(\d+(?:\.\d+)?|[А-Я]\.\d+)",
    re.IGNORECASE,
)

TABLE_END_RE = re.compile(
    r"^Окончание\s+таблицы\s+(\d+(?:\.\d+)?|[А-Я]\.\d+)",
    re.IGNORECASE,
)


def _format_table_caption(paragraph, config: STUConfig) -> None:
    """Format a table caption paragraph: left-aligned, no indent, TNR 14."""
    for run in paragraph.runs:
        run.font.name = config.font_name
        run.font.size = Pt(config.font_size_table_header_pt)

    pf = paragraph.paragraph_format
    pf.alignment = WD_ALIGN_PARAGRAPH.LEFT
    pf.first_line_indent = Mm(0)
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)


def _format_table_continuation(paragraph, config: STUConfig) -> None:
    """Format continuation/end caption: left-aligned, no indent."""
    for run in paragraph.runs:
        run.font.name = config.font_name
        run.font.size = Pt(config.font_size_table_header_pt)

    pf = paragraph.paragraph_format
    pf.alignment = WD_ALIGN_PARAGRAPH.LEFT
    pf.first_line_indent = Mm(0)


def _set_cell_font(cell, font_name: str, font_size_pt: int) -> None:
    """Set font for all paragraphs in a table cell."""
    for paragraph in cell.paragraphs:
        for run in paragraph.runs:
            run.font.name = font_name
            run.font.size = Pt(font_size_pt)


def _format_header_row(row, config: STUConfig) -> None:
    """Format the header row: bold, centered, proper font."""
    for cell in row.cells:
        for paragraph in cell.paragraphs:
            paragraph.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
            paragraph.paragraph_format.first_line_indent = Mm(0)
            for run in paragraph.runs:
                run.font.name = config.font_name
                run.font.size = Pt(config.font_size_table_pt)
                run.font.bold = True


def _set_table_borders(table) -> None:
    """Set table borders: lines on left, right, bottom; double line under header."""
    tbl = table._tbl
    tblPr = tbl.find(qn("w:tblPr"))
    if tblPr is None:
        tblPr = etree.SubElement(tbl, qn("w:tblPr"))

    # Remove existing borders
    existing_borders = tblPr.find(qn("w:tblBorders"))
    if existing_borders is not None:
        tblPr.remove(existing_borders)

    borders = etree.SubElement(tblPr, qn("w:tblBorders"))

    border_attrs = {
        qn("w:val"): "single",
        qn("w:sz"): "4",
        qn("w:space"): "0",
        qn("w:color"): "000000",
    }

    for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
        border = etree.SubElement(borders, qn(f"w:{side}"))
        for attr, val in border_attrs.items():
            border.set(attr, val)


def _set_header_row_repeat(table) -> None:
    """Mark the first row as a header row that repeats on each page."""
    if len(table.rows) == 0:
        return
    first_row = table.rows[0]
    trPr = first_row._tr.find(qn("w:trPr"))
    if trPr is None:
        trPr = etree.SubElement(first_row._tr, qn("w:trPr"))
    tblHeader = trPr.find(qn("w:tblHeader"))
    if tblHeader is None:
        tblHeader = etree.SubElement(trPr, qn("w:tblHeader"))


def apply_tables(doc: Document, config: STUConfig) -> None:
    """Format all tables and their captions in the document."""
    # Format table captions in paragraphs
    caption_count = 0
    for para in doc.paragraphs:
        text = para.text.strip()
        if TABLE_CAPTION_RE.match(text):
            _format_table_caption(para, config)
            caption_count += 1
        elif TABLE_CONTINUATION_RE.match(text) or TABLE_END_RE.match(text):
            _format_table_continuation(para, config)

    # Format table contents
    table_count = 0
    for table in doc.tables:
        # Set borders
        _set_table_borders(table)

        # Set header row repeat
        _set_header_row_repeat(table)

        # Format header row (first row)
        if len(table.rows) > 0:
            _format_header_row(table.rows[0], config)

        # Format data rows
        for row_idx, row in enumerate(table.rows):
            if row_idx == 0:
                continue  # Already formatted as header
            for cell in row.cells:
                _set_cell_font(cell, config.font_name, config.font_size_table_pt)
                for paragraph in cell.paragraphs:
                    paragraph.paragraph_format.first_line_indent = Mm(0)

        # Table alignment
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        table_count += 1

    logger.info("Tables formatted: %d table(s), %d caption(s).", table_count, caption_count)
