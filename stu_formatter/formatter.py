"""Main orchestrator: opens the document, runs all formatting modules, saves."""

import logging

from docx import Document

from .config import STUConfig
from .modules.page_setup import apply_page_setup
from .modules.fonts_and_paragraphs import apply_body_formatting
from .modules.page_numbers import apply_page_numbers
from .modules.headings import apply_headings
from .modules.formulas import apply_formulas
from .modules.tables import apply_tables
from .modules.figures import apply_figures
from .modules.lists import apply_lists
from .modules.bibliography import apply_bibliography
from .modules.appendices import apply_appendices

logger = logging.getLogger(__name__)


class STUFormatter:
    """Formats a .docx document according to STU 7.5-07-2021."""

    def __init__(self, config: STUConfig | None = None):
        self.config = config or STUConfig()

    def format(self, input_path: str, output_path: str) -> None:
        """Open, format, and save the document.

        Args:
            input_path: Path to the input .docx file.
            output_path: Path for the formatted output .docx file.
        """
        logger.info("Opening document: %s", input_path)
        doc = Document(input_path)

        # 1. Page setup (margins, paper size)
        logger.info("Step 1/10: Page setup...")
        apply_page_setup(doc, self.config)

        # 2. Body text formatting (font, size, spacing, indent, alignment)
        logger.info("Step 2/10: Body text formatting...")
        apply_body_formatting(doc, self.config)

        # 3. Headings (structural elements and numbered sections)
        logger.info("Step 3/10: Headings...")
        apply_headings(doc, self.config)

        # 4. Formulas
        logger.info("Step 4/10: Formulas...")
        apply_formulas(doc, self.config)

        # 5. Tables
        logger.info("Step 5/10: Tables...")
        apply_tables(doc, self.config)

        # 6. Figures
        logger.info("Step 6/10: Figures...")
        apply_figures(doc, self.config)

        # 7. Lists
        logger.info("Step 7/10: Lists...")
        apply_lists(doc, self.config)

        # 8. Bibliography
        logger.info("Step 8/10: Bibliography...")
        apply_bibliography(doc, self.config)

        # 9. Appendices
        logger.info("Step 9/10: Appendices...")
        apply_appendices(doc, self.config)

        # 10. Page numbers
        logger.info("Step 10/10: Page numbers...")
        apply_page_numbers(doc, self.config)

        # Save
        doc.save(output_path)
        logger.info("Document saved: %s", output_path)
