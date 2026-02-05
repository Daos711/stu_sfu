"""Formula formatting: OOXML math elements, indentation, spacing, numbering."""

import logging
import re

from docx import Document
from docx.shared import Pt, Mm, Emu
from docx.oxml.ns import qn
from lxml import etree

from ..config import STUConfig
from ..utils.xml_helpers import (
    has_omathpara,
    has_inline_omath,
    paragraph_is_empty,
    find_elements,
    NAMESPACES,
)

logger = logging.getLogger(__name__)

# Pattern for formula references in text: (1), (1.1), (А.1)
FORMULA_NUM_RE = re.compile(r"\((\d+(?:\.\d+)?|[А-Я]\.\d+)\)")


def _is_standalone_formula(paragraph) -> bool:
    """Check if paragraph is a standalone formula (oMathPara, no significant text)."""
    if not has_omathpara(paragraph):
        return False

    # Check that the paragraph doesn't have significant non-math text
    text = ""
    for run in paragraph.runs:
        run_xml = run._element
        # Check if run is outside math
        parent = run_xml.getparent()
        in_math = False
        while parent is not None:
            tag = parent.tag
            if "oMath" in tag:
                in_math = True
                break
            parent = parent.getparent()
        if not in_math:
            text += run.text or ""

    # Allow formula number like "(1)" or "(1.1)" in the text
    cleaned = FORMULA_NUM_RE.sub("", text).strip()
    return len(cleaned) == 0


def _ensure_empty_para_before(doc: Document, paragraph, index: int) -> None:
    """Ensure blank paragraph before formula."""
    if index <= 0:
        return
    paragraphs = doc.paragraphs
    prev_p = paragraphs[index - 1]
    if not paragraph_is_empty(prev_p):
        new_p = doc.element.makeelement(qn("w:p"), {})
        paragraph._element.addprevious(new_p)


def _ensure_empty_para_after(doc: Document, paragraph, index: int) -> None:
    """Ensure blank paragraph after formula."""
    paragraphs = doc.paragraphs
    if index + 1 < len(paragraphs):
        next_p = paragraphs[index + 1]
        if not paragraph_is_empty(next_p):
            new_p = doc.element.makeelement(qn("w:p"), {})
            paragraph._element.addnext(new_p)
    else:
        new_p = doc.element.makeelement(qn("w:p"), {})
        paragraph._element.addnext(new_p)


def _set_formula_indent(paragraph, config: STUConfig) -> None:
    """Set first-line indent of 12.5 mm for formula paragraph."""
    pf = paragraph.paragraph_format
    pf.first_line_indent = Mm(config.first_line_indent_mm)
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)


def _add_right_tab_number(paragraph, number_text: str, config: STUConfig) -> None:
    """Add a right-aligned tab stop with formula number like (1).

    Places the number at the right margin of the page.
    """
    # Calculate right tab position:
    # Page width 210mm - left margin 30mm - right margin 10mm = 170mm text area
    tab_position = Mm(170 - config.first_line_indent_mm)

    pPr = paragraph._element.find(qn("w:pPr"))
    if pPr is None:
        pPr = etree.SubElement(paragraph._element, qn("w:pPr"))

    # Add tab definition
    tabs = pPr.find(qn("w:tabs"))
    if tabs is None:
        tabs = etree.SubElement(pPr, qn("w:tabs"))

    tab = etree.SubElement(tabs, qn("w:tab"))
    tab.set(qn("w:val"), "right")
    tab.set(qn("w:pos"), str(int(tab_position)))

    # Add tab character and number as a run at the end
    run_elem = etree.SubElement(paragraph._element, qn("w:r"))

    # Run properties
    rPr = etree.SubElement(run_elem, qn("w:rPr"))
    rFonts = etree.SubElement(rPr, qn("w:rFonts"))
    rFonts.set(qn("w:ascii"), config.font_name)
    rFonts.set(qn("w:hAnsi"), config.font_name)
    sz = etree.SubElement(rPr, qn("w:sz"))
    sz.set(qn("w:val"), str(config.font_size_pt * 2))

    # Tab character
    etree.SubElement(run_elem, qn("w:tab"))

    # Number text
    t = etree.SubElement(run_elem, qn("w:t"))
    t.set(qn("xml:space"), "preserve")
    t.text = number_text


def apply_formulas(doc: Document, config: STUConfig) -> None:
    """Find and format all standalone formulas in the document.

    - Sets first-line indent of 12.5 mm
    - Ensures blank lines above and below
    - Preserves existing formula numbering
    """
    paragraphs = list(doc.paragraphs)
    formula_count = 0

    for i, para in enumerate(paragraphs):
        if _is_standalone_formula(para):
            _set_formula_indent(para, config)

            # Set font for non-math runs
            for run in para.runs:
                run.font.name = config.font_name
                run.font.size = Pt(config.font_size_pt)

            _ensure_empty_para_before(doc, para, i)
            _ensure_empty_para_after(doc, para, i)
            formula_count += 1

    if formula_count > 0:
        logger.info("Formatted %d standalone formula(s).", formula_count)
    else:
        logger.info("No standalone formulas found.")
