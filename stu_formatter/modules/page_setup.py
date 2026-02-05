"""Page setup: margins, paper size, orientation."""

import logging

from docx import Document
from docx.shared import Mm
from docx.enum.section import WD_ORIENT

from ..config import STUConfig

logger = logging.getLogger(__name__)


def apply_page_setup(doc: Document, config: STUConfig) -> None:
    """Configure page margins, paper size, and orientation for all sections."""
    for section in doc.sections:
        # Paper size: A4
        section.page_width = Mm(210)
        section.page_height = Mm(297)

        if section.orientation == WD_ORIENT.LANDSCAPE:
            # Landscape: swap width/height, use landscape margins
            section.page_width = Mm(297)
            section.page_height = Mm(210)
            section.left_margin = Mm(config.margin_landscape_left_mm)
            section.right_margin = Mm(config.margin_landscape_right_mm)
            section.top_margin = Mm(config.margin_landscape_top_mm)
            section.bottom_margin = Mm(config.margin_landscape_bottom_mm)
            logger.info("Applied landscape margins to section.")
        else:
            # Portrait: standard margins
            section.left_margin = Mm(config.margin_left_mm)
            section.right_margin = Mm(config.margin_right_mm)
            section.top_margin = Mm(config.margin_top_mm)
            section.bottom_margin = Mm(config.margin_bottom_mm)

    logger.info("Page setup applied: A4, margins configured for %d section(s).", len(doc.sections))
