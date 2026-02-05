"""Bibliography (СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ) formatting."""

import logging
import re

from docx import Document
from docx.shared import Pt, Mm
from docx.enum.text import WD_ALIGN_PARAGRAPH

from ..config import STUConfig

logger = logging.getLogger(__name__)

# The bibliography heading text (case-insensitive match)
BIB_HEADING = "список использованных источников"

# Pattern for numbered bibliography entries: "1 ", "2 ", "10 ", etc.
BIB_ENTRY_RE = re.compile(r"^(\d+)\s+")
# Also matches "1. ", "2. " style
BIB_ENTRY_DOT_RE = re.compile(r"^(\d+)\.\s+")


def _find_bibliography_section(doc: Document) -> tuple[int, int] | None:
    """Find the start and end paragraph indices of the bibliography section.

    Returns:
        Tuple of (start_index, end_index) exclusive of heading,
        or None if not found.
    """
    paragraphs = doc.paragraphs
    start_idx = None

    for i, para in enumerate(paragraphs):
        text = para.text.strip().lower()
        if text == BIB_HEADING:
            start_idx = i + 1  # entries start after the heading
            break

    if start_idx is None:
        return None

    # Find the end: next structural element heading or end of document
    from .headings import STRUCTURAL_ELEMENTS_LOWER, APPENDIX_RE

    end_idx = len(paragraphs)
    for i in range(start_idx, len(paragraphs)):
        text = paragraphs[i].text.strip()
        text_lower = text.lower()
        if text_lower in STRUCTURAL_ELEMENTS_LOWER and text_lower != BIB_HEADING:
            end_idx = i
            break
        if APPENDIX_RE.match(text):
            end_idx = i
            break

    return (start_idx, end_idx)


def apply_bibliography(doc: Document, config: STUConfig) -> None:
    """Format the bibliography section.

    - Each entry: TNR 14, first-line indent 12.5 mm, justified
    - Numbered with Arabic numerals
    - Continuous numbering
    """
    section = _find_bibliography_section(doc)
    if section is None:
        logger.info("No bibliography section found.")
        return

    start_idx, end_idx = section
    entry_count = 0

    for i in range(start_idx, end_idx):
        para = doc.paragraphs[i]
        text = para.text.strip()

        if not text:
            continue

        # Apply font formatting
        for run in para.runs:
            run.font.name = config.font_name
            run.font.size = Pt(config.font_size_pt)

        # Paragraph formatting
        pf = para.paragraph_format
        pf.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        pf.first_line_indent = Mm(config.first_line_indent_mm)
        pf.space_before = Pt(0)
        pf.space_after = Pt(0)

        # Check if it's a numbered entry
        if BIB_ENTRY_RE.match(text) or BIB_ENTRY_DOT_RE.match(text):
            entry_count += 1

    logger.info("Bibliography formatted: %d entries.", entry_count)
