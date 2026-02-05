"""Font, size, line spacing, indentation, and alignment for body text."""

import logging
import re

from docx import Document
from docx.shared import Pt, Mm
from docx.enum.text import WD_ALIGN_PARAGRAPH

from ..config import STUConfig
from ..utils.xml_helpers import has_drawing, has_math_element

logger = logging.getLogger(__name__)

# Patterns for detecting headings and special paragraphs
HEADING_STYLE_PREFIXES = ("Heading", "heading", "Заголовок")
NUMBERED_HEADING_RE = re.compile(r"^\d+(\.\d+)*\s+\S")


def _is_heading_style(paragraph) -> bool:
    """Check if paragraph has a heading style."""
    style_name = paragraph.style.name if paragraph.style else ""
    return any(style_name.startswith(p) for p in HEADING_STYLE_PREFIXES)


def _is_toc_style(paragraph) -> bool:
    """Check if paragraph has a table-of-contents style."""
    style_name = paragraph.style.name if paragraph.style else ""
    return style_name.startswith("TOC") or style_name.startswith("toc")


def _is_table_or_figure_caption(text: str) -> bool:
    """Check if text looks like a table/figure caption."""
    stripped = text.strip()
    return (
        stripped.startswith("Таблица ")
        or stripped.startswith("Рисунок ")
        or stripped.startswith("Продолжение таблицы ")
        or stripped.startswith("Окончание таблицы ")
    )


def apply_body_formatting(doc: Document, config: STUConfig) -> None:
    """Apply font, size, spacing, indent, and alignment to all body paragraphs.

    Skips headings, table contents, figure captions, and formulas —
    those are handled by their respective modules.
    """
    count = 0
    for paragraph in doc.paragraphs:
        # Skip headings (handled by headings module)
        if _is_heading_style(paragraph):
            continue

        # Skip TOC entries
        if _is_toc_style(paragraph):
            continue

        # Skip table/figure captions (handled by their modules)
        text = paragraph.text.strip()
        if _is_table_or_figure_caption(text):
            continue

        # Apply formatting to runs
        for run in paragraph.runs:
            run.font.name = config.font_name
            run.font.size = Pt(config.font_size_pt)

        # Paragraph format
        pf = paragraph.paragraph_format

        # Alignment: justify for body text
        pf.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

        # Line spacing
        if config.line_spacing == 1.5:
            pf.line_spacing = 1.5
        else:
            pf.line_spacing = 1.0

        # First line indent (12.5 mm) for text paragraphs
        # Don't set indent for paragraphs that only contain images
        if not has_drawing(paragraph) or text:
            pf.first_line_indent = Mm(config.first_line_indent_mm)

        # Reset space before/after (will be set by specific modules where needed)
        pf.space_before = Pt(0)
        pf.space_after = Pt(0)

        count += 1

    logger.info("Body formatting applied to %d paragraphs.", count)
