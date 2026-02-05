"""List formatting: replace bullets with dashes, proper indentation."""

import logging
import re

from docx import Document
from docx.shared import Pt, Mm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn

from ..config import STUConfig

logger = logging.getLogger(__name__)

# Common bullet characters to replace with dashes
BULLET_CHARS = {"•", "●", "○", "■", "□", "▪", "▫", "►", "▸", "‣", "⁃", "∙"}

# Pattern for lettered list items: а), б), в)
LETTER_LIST_RE = re.compile(r"^([а-яё])\)\s+")

# Pattern for numbered sub-list items: 1), 2), 3)
NUMBERED_SUB_LIST_RE = re.compile(r"^(\d+)\)\s+")

# List style name patterns
LIST_STYLE_PATTERNS = (
    "List Paragraph",
    "List Bullet",
    "List Number",
    "Маркированный",
    "Нумерованный",
)


def _is_list_style(paragraph) -> bool:
    """Check if paragraph has a list-related style."""
    style_name = paragraph.style.name if paragraph.style else ""
    return any(pattern in style_name for pattern in LIST_STYLE_PATTERNS)


def _has_numbering(paragraph) -> bool:
    """Check if paragraph has Word numbering (numPr)."""
    pPr = paragraph._element.find(qn("w:pPr"))
    if pPr is not None:
        numPr = pPr.find(qn("w:numPr"))
        return numPr is not None
    return False


def _starts_with_bullet(text: str) -> bool:
    """Check if text starts with a bullet character."""
    if not text:
        return False
    return text[0] in BULLET_CHARS


def _replace_bullet_with_dash(paragraph) -> bool:
    """Replace leading bullet character with a dash.

    Returns True if replacement was made.
    """
    if not paragraph.runs:
        return False

    first_run = paragraph.runs[0]
    text = first_run.text or ""

    if not text:
        return False

    # Replace leading bullet
    if text[0] in BULLET_CHARS:
        first_run.text = "– " + text[1:].lstrip(" ")
        return True

    return False


def _remove_word_numbering(paragraph) -> None:
    """Remove Word's built-in numbering from a paragraph."""
    pPr = paragraph._element.find(qn("w:pPr"))
    if pPr is not None:
        numPr = pPr.find(qn("w:numPr"))
        if numPr is not None:
            pPr.remove(numPr)


def apply_lists(doc: Document, config: STUConfig) -> None:
    """Format list items: replace bullets with dashes, set indent.

    Rules:
    - Each list item has first-line indent of 12.5 mm
    - Bullets replaced with em-dash (–)
    - Lettered items: а), б), в) — kept as-is
    - Numbered sub-items: 1), 2) — shifted right by ~2 characters
    """
    replaced_count = 0

    for para in doc.paragraphs:
        text = para.text.strip()
        if not text:
            continue

        is_list = _is_list_style(para) or _has_numbering(para) or _starts_with_bullet(text)

        if not is_list:
            continue

        # Replace bullets with dashes
        if _starts_with_bullet(text):
            if _replace_bullet_with_dash(para):
                replaced_count += 1

        # Handle Word numbering for bullet lists
        if _has_numbering(para) and _is_list_style(para):
            style_name = para.style.name if para.style else ""
            if "Bullet" in style_name or "Маркированный" in style_name:
                _remove_word_numbering(para)
                # Prepend dash if not already there
                if para.runs and not (para.runs[0].text or "").startswith("–"):
                    para.runs[0].text = "– " + (para.runs[0].text or "")
                    replaced_count += 1

        # Apply formatting
        for run in para.runs:
            run.font.name = config.font_name
            run.font.size = Pt(config.font_size_pt)

        pf = para.paragraph_format
        pf.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        pf.first_line_indent = Mm(config.first_line_indent_mm)
        pf.space_before = Pt(0)
        pf.space_after = Pt(0)

        # For numbered sub-items 1), 2) — add extra left indent
        refreshed_text = para.text.strip()
        if NUMBERED_SUB_LIST_RE.match(refreshed_text):
            pf.left_indent = Mm(config.first_line_indent_mm + 5)

    logger.info("Lists formatted: %d bullet(s) replaced with dashes.", replaced_count)
