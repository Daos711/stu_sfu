"""Page break detection using LibreOffice for table splitting.

This module provides functionality to determine which pages tables
occupy by converting .docx to PDF via LibreOffice and analyzing the result.

Note: This is a Phase 4 feature. Currently provides stub/warning functionality.
"""

import logging
import shutil
import subprocess
import tempfile
from pathlib import Path

logger = logging.getLogger(__name__)


def is_libreoffice_available() -> bool:
    """Check if LibreOffice is installed and accessible."""
    return shutil.which("soffice") is not None


def convert_to_pdf(docx_path: str, output_dir: str | None = None) -> str | None:
    """Convert a .docx file to PDF using LibreOffice.

    Args:
        docx_path: Path to the .docx file.
        output_dir: Directory for the output PDF. Uses temp dir if None.

    Returns:
        Path to the generated PDF, or None on failure.
    """
    if not is_libreoffice_available():
        logger.warning("LibreOffice not found. Table page-break detection unavailable.")
        return None

    if output_dir is None:
        output_dir = tempfile.mkdtemp()

    try:
        result = subprocess.run(
            [
                "soffice",
                "--headless",
                "--convert-to", "pdf",
                "--outdir", output_dir,
                docx_path,
            ],
            capture_output=True,
            text=True,
            timeout=120,
        )
        if result.returncode != 0:
            logger.error("LibreOffice conversion failed: %s", result.stderr)
            return None

        pdf_name = Path(docx_path).stem + ".pdf"
        pdf_path = Path(output_dir) / pdf_name
        if pdf_path.exists():
            return str(pdf_path)

        logger.error("PDF file not found after conversion.")
        return None

    except subprocess.TimeoutExpired:
        logger.error("LibreOffice conversion timed out.")
        return None
    except Exception as e:
        logger.error("Error during PDF conversion: %s", e)
        return None


def detect_table_pages(docx_path: str) -> dict:
    """Detect which pages each table occupies.

    This is a placeholder for Phase 4 implementation.
    Full implementation would:
    1. Convert .docx to PDF
    2. Parse the PDF to find table positions
    3. Return a mapping of table index -> list of page numbers

    Returns:
        Empty dict (stub). Future: {table_index: [page_numbers]}
    """
    logger.info(
        "Table page-break detection is not yet implemented. "
        "Tables spanning multiple pages may need manual adjustment."
    )
    return {}
