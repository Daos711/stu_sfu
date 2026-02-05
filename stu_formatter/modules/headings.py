"""Headings: structural elements and numbered section/subsection headings."""

import logging
import re

from docx import Document
from docx.shared import Pt, Mm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn

from ..config import STUConfig
from ..utils.xml_helpers import paragraph_is_empty

logger = logging.getLogger(__name__)

# Structural elements that must be uppercase, bold, centered, on a new page
STRUCTURAL_ELEMENTS_LOWER = {
    "реферат", "аннотация", "содержание", "введение", "заключение",
    "список сокращений", "список использованных источников",
}

# Regex for numbered headings: "1 Title", "1.1 Title", "1.1.1 Title", "1.1.1.1 Title"
HEADING_NUM_RE = re.compile(r"^(\d+(?:\.\d+)*)\s+(.*)")

# Heading style name patterns
HEADING_STYLE_RE = re.compile(r"^(?:Heading|heading|Заголовок)\s*(\d+)$")

# Appendix heading pattern
APPENDIX_RE = re.compile(
    r"^ПРИЛОЖЕНИЕ\s+([А-Я])\b",
    re.IGNORECASE,
)


def _is_structural_element(text: str) -> bool:
    """Check if text matches a structural element name."""
    return text.strip().lower() in STRUCTURAL_ELEMENTS_LOWER


def _is_appendix_heading(text: str) -> bool:
    """Check if text is an appendix heading."""
    return bool(APPENDIX_RE.match(text.strip()))


def _get_heading_level(paragraph) -> int | None:
    """Determine heading level from style.

    Returns:
        1 for Heading 1, 2 for Heading 2, etc. None if not a heading style.
    """
    style_name = paragraph.style.name if paragraph.style else ""
    m = HEADING_STYLE_RE.match(style_name)
    if m:
        return int(m.group(1))
    return None


def _get_numbered_heading_level(text: str) -> int | None:
    """Determine heading level from numbering pattern.

    "1 Title" -> 1, "1.1 Title" -> 2, "1.1.1 Title" -> 3, etc.
    """
    m = HEADING_NUM_RE.match(text.strip())
    if m:
        num_part = m.group(1)
        return num_part.count(".") + 1
    return None


def _add_page_break_before(paragraph) -> None:
    """Set page break before this paragraph."""
    paragraph.paragraph_format.page_break_before = True


def _ensure_blank_line_after(doc: Document, paragraph, index: int) -> None:
    """Ensure there is a blank paragraph after the given paragraph.

    Inserts one if the next paragraph is not empty.
    """
    paragraphs = doc.paragraphs
    if index + 1 < len(paragraphs):
        next_p = paragraphs[index + 1]
        if not paragraph_is_empty(next_p):
            # Insert a blank paragraph after current one
            new_p = doc.element.makeelement(qn("w:p"), {})
            paragraph._element.addnext(new_p)
    else:
        # Last paragraph — add blank after
        new_p = doc.element.makeelement(qn("w:p"), {})
        paragraph._element.addnext(new_p)


def _ensure_blank_line_before(doc: Document, paragraph, index: int) -> None:
    """Ensure there is a blank paragraph before the given paragraph.

    Does not insert before the very first paragraph.
    """
    if index <= 0:
        return
    paragraphs = doc.paragraphs
    prev_p = paragraphs[index - 1]
    if not paragraph_is_empty(prev_p):
        new_p = doc.element.makeelement(qn("w:p"), {})
        paragraph._element.addprevious(new_p)


def _format_structural_heading(paragraph, config: STUConfig) -> None:
    """Format a structural element heading (ВВЕДЕНИЕ, etc.)."""
    # Uppercase the text
    for run in paragraph.runs:
        if run.text:
            run.text = run.text.upper()
        run.font.name = config.font_name
        run.font.size = Pt(config.font_size_pt)
        run.font.bold = True
        run.font.underline = False

    pf = paragraph.paragraph_format
    pf.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pf.first_line_indent = Mm(0)
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)

    # Remove trailing period
    if paragraph.runs and paragraph.runs[-1].text:
        last_run = paragraph.runs[-1]
        if last_run.text.endswith("."):
            last_run.text = last_run.text.rstrip(".")


def _format_section_heading(paragraph, config: STUConfig) -> None:
    """Format a numbered section/subsection heading."""
    for run in paragraph.runs:
        run.font.name = config.font_name
        run.font.size = Pt(config.font_size_pt)
        run.font.bold = True
        run.font.underline = False

    pf = paragraph.paragraph_format
    pf.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    pf.first_line_indent = Mm(config.first_line_indent_mm)
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)

    # Remove trailing period (unless two sentences)
    if paragraph.runs and paragraph.runs[-1].text:
        last_run = paragraph.runs[-1]
        text = last_run.text
        # Only strip trailing period if it's a single sentence
        if text.endswith(".") and text.count(".") <= text.count(" ") + 1:
            # Simple heuristic: strip trailing dot if it doesn't look like
            # the dot separating two sentences
            stripped = text.rstrip(".")
            if stripped:
                last_run.text = stripped


def _format_appendix_heading(paragraph, config: STUConfig) -> None:
    """Format an appendix heading."""
    for run in paragraph.runs:
        run.font.name = config.font_name
        run.font.size = Pt(config.font_size_pt)
        run.font.bold = True
        run.font.underline = False

    pf = paragraph.paragraph_format
    pf.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pf.first_line_indent = Mm(0)
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)


def apply_headings(doc: Document, config: STUConfig) -> None:
    """Detect and format all headings in the document.

    Handles:
    1. Structural elements (ВВЕДЕНИЕ, ЗАКЛЮЧЕНИЕ, etc.)
    2. Appendix headings (ПРИЛОЖЕНИЕ А, etc.)
    3. Numbered section headings (1, 1.1, 1.1.1, etc.)
    4. Style-based headings (Heading 1, Heading 2, etc.)
    """
    paragraphs = list(doc.paragraphs)
    structural_count = 0
    section_count = 0
    appendix_count = 0

    for i, para in enumerate(paragraphs):
        text = para.text.strip()
        if not text:
            continue

        # 1. Check for structural elements
        if _is_structural_element(text):
            _format_structural_heading(para, config)
            _add_page_break_before(para)
            _ensure_blank_line_after(doc, para, i)
            structural_count += 1
            continue

        # 2. Check for appendix headings
        if _is_appendix_heading(text):
            _format_appendix_heading(para, config)
            _add_page_break_before(para)
            _ensure_blank_line_after(doc, para, i)
            appendix_count += 1
            continue

        # 3. Check heading by style
        level = _get_heading_level(para)
        if level is not None:
            if level == 1:
                # Section heading (level 1) — may need page break
                _format_section_heading(para, config)
                _ensure_blank_line_before(doc, para, i)
                _ensure_blank_line_after(doc, para, i)
            else:
                # Subsection, point, subpoint
                _format_section_heading(para, config)
                _ensure_blank_line_before(doc, para, i)
                _ensure_blank_line_after(doc, para, i)
            section_count += 1
            continue

        # 4. Check heading by numbering pattern (fallback)
        num_level = _get_numbered_heading_level(text)
        if num_level is not None and num_level <= 4:
            _format_section_heading(para, config)
            _ensure_blank_line_before(doc, para, i)
            _ensure_blank_line_after(doc, para, i)
            section_count += 1
            continue

    logger.info(
        "Headings formatted: %d structural, %d section/subsection, %d appendix.",
        structural_count, section_count, appendix_count,
    )
