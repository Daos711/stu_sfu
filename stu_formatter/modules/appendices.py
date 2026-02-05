"""Appendix formatting: headings, internal numbering."""

import logging
import re

from docx import Document
from docx.shared import Pt, Mm
from docx.enum.text import WD_ALIGN_PARAGRAPH

from ..config import STUConfig

logger = logging.getLogger(__name__)

# Appendix heading pattern: "ПРИЛОЖЕНИЕ А" or "Приложение А"
APPENDIX_HEADING_RE = re.compile(
    r"^ПРИЛОЖЕНИЕ\s+([А-Я])\b",
    re.IGNORECASE,
)

# Appendix sub-heading (after "ПРИЛОЖЕНИЕ А"): title on the next line
APPENDIX_TITLE_RE = re.compile(r"^[А-ЯA-Z]")


def _find_appendix_sections(doc: Document) -> list[dict]:
    """Find all appendix sections and their boundaries.

    Returns:
        List of dicts with keys: 'letter', 'heading_idx', 'start_idx', 'end_idx'
    """
    paragraphs = doc.paragraphs
    appendices = []

    for i, para in enumerate(paragraphs):
        text = para.text.strip()
        m = APPENDIX_HEADING_RE.match(text)
        if m:
            letter = m.group(1).upper()
            appendices.append({
                "letter": letter,
                "heading_idx": i,
                "start_idx": i + 1,
                "end_idx": len(paragraphs),  # will be adjusted
            })

    # Adjust end indices
    for j in range(len(appendices) - 1):
        appendices[j]["end_idx"] = appendices[j + 1]["heading_idx"]

    return appendices


def apply_appendices(doc: Document, config: STUConfig) -> None:
    """Format appendix headings and content.

    - Appendix heading: ПРИЛОЖЕНИЕ А — centered, bold, no indent, new page
    - Appendix title (next line): centered, bold, with first capital letter
    - Content within appendices follows standard body formatting
    """
    appendices = _find_appendix_sections(doc)

    if not appendices:
        logger.info("No appendices found.")
        return

    paragraphs = doc.paragraphs

    for app in appendices:
        heading_para = paragraphs[app["heading_idx"]]

        # Format heading: "ПРИЛОЖЕНИЕ А"
        for run in heading_para.runs:
            if run.text:
                run.text = run.text.upper()
            run.font.name = config.font_name
            run.font.size = Pt(config.font_size_pt)
            run.font.bold = True

        pf = heading_para.paragraph_format
        pf.alignment = WD_ALIGN_PARAGRAPH.CENTER
        pf.first_line_indent = Mm(0)
        pf.page_break_before = True

        # Check if next paragraph is the appendix title
        title_idx = app["start_idx"]
        if title_idx < app["end_idx"]:
            title_para = paragraphs[title_idx]
            title_text = title_para.text.strip()

            # If the next paragraph looks like a title (not a numbered entry,
            # not empty, not a table/figure caption)
            if (
                title_text
                and not title_text.startswith("Таблица ")
                and not title_text.startswith("Рисунок ")
                and not re.match(r"^\d", title_text)
            ):
                for run in title_para.runs:
                    run.font.name = config.font_name
                    run.font.size = Pt(config.font_size_pt)
                    run.font.bold = True

                title_pf = title_para.paragraph_format
                title_pf.alignment = WD_ALIGN_PARAGRAPH.CENTER
                title_pf.first_line_indent = Mm(0)

    logger.info("Appendices formatted: %d appendix/appendices.", len(appendices))
