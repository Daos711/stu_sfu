"""Page numbering: centered at bottom, Times New Roman 14, not on title page."""

import logging

from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn
from lxml import etree

from ..config import STUConfig

logger = logging.getLogger(__name__)


def _create_page_number_footer_xml(font_name: str, font_size: int) -> etree._Element:
    """Create XML for a centered page number footer.

    The footer contains a single paragraph with:
    - Center alignment
    - No first-line indent
    - A PAGE field showing the current page number
    - Font: Times New Roman 14pt
    """
    nsmap = {
        "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    }
    w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

    # Create paragraph
    p = etree.SubElement(etree.Element("dummy"), qn("w:p"))

    # Paragraph properties: centered, no indent
    pPr = etree.SubElement(p, qn("w:pPr"))
    jc = etree.SubElement(pPr, qn("w:jc"))
    jc.set(qn("w:val"), "center")
    ind = etree.SubElement(pPr, qn("w:ind"))
    ind.set(qn("w:firstLine"), "0")

    # Run properties for font
    rPr_para = etree.SubElement(pPr, qn("w:rPr"))
    rFonts = etree.SubElement(rPr_para, qn("w:rFonts"))
    rFonts.set(qn("w:ascii"), font_name)
    rFonts.set(qn("w:hAnsi"), font_name)
    rFonts.set(qn("w:cs"), font_name)
    sz = etree.SubElement(rPr_para, qn("w:sz"))
    sz.set(qn("w:val"), str(font_size * 2))  # half-points
    szCs = etree.SubElement(rPr_para, qn("w:szCs"))
    szCs.set(qn("w:val"), str(font_size * 2))

    # Run with field begin
    r1 = etree.SubElement(p, qn("w:r"))
    rPr1 = etree.SubElement(r1, qn("w:rPr"))
    rFonts1 = etree.SubElement(rPr1, qn("w:rFonts"))
    rFonts1.set(qn("w:ascii"), font_name)
    rFonts1.set(qn("w:hAnsi"), font_name)
    rFonts1.set(qn("w:cs"), font_name)
    sz1 = etree.SubElement(rPr1, qn("w:sz"))
    sz1.set(qn("w:val"), str(font_size * 2))
    szCs1 = etree.SubElement(rPr1, qn("w:szCs"))
    szCs1.set(qn("w:val"), str(font_size * 2))
    fldChar1 = etree.SubElement(r1, qn("w:fldChar"))
    fldChar1.set(qn("w:fldCharType"), "begin")

    # Run with field code PAGE
    r2 = etree.SubElement(p, qn("w:r"))
    rPr2 = etree.SubElement(r2, qn("w:rPr"))
    rFonts2 = etree.SubElement(rPr2, qn("w:rFonts"))
    rFonts2.set(qn("w:ascii"), font_name)
    rFonts2.set(qn("w:hAnsi"), font_name)
    rFonts2.set(qn("w:cs"), font_name)
    sz2 = etree.SubElement(rPr2, qn("w:sz"))
    sz2.set(qn("w:val"), str(font_size * 2))
    szCs2 = etree.SubElement(rPr2, qn("w:szCs"))
    szCs2.set(qn("w:val"), str(font_size * 2))
    instrText = etree.SubElement(r2, qn("w:instrText"))
    instrText.set(qn("xml:space"), "preserve")
    instrText.text = " PAGE "

    # Run with field end
    r3 = etree.SubElement(p, qn("w:r"))
    rPr3 = etree.SubElement(r3, qn("w:rPr"))
    rFonts3 = etree.SubElement(rPr3, qn("w:rFonts"))
    rFonts3.set(qn("w:ascii"), font_name)
    rFonts3.set(qn("w:hAnsi"), font_name)
    rFonts3.set(qn("w:cs"), font_name)
    sz3 = etree.SubElement(rPr3, qn("w:sz"))
    sz3.set(qn("w:val"), str(font_size * 2))
    szCs3 = etree.SubElement(rPr3, qn("w:szCs"))
    szCs3.set(qn("w:val"), str(font_size * 2))
    fldChar3 = etree.SubElement(r3, qn("w:fldChar"))
    fldChar3.set(qn("w:fldCharType"), "end")

    return p


def apply_page_numbers(doc: Document, config: STUConfig) -> None:
    """Add page numbers to the document footer.

    - Centered at bottom of page
    - Times New Roman, 14pt
    - Title page (first page) has no number but is counted
    """
    for i, section in enumerate(doc.sections):
        if i == 0:
            # First section: different first page (no number on title page)
            section.different_first_page_header_footer = True

            # Set up the default footer (pages 2+)
            footer = section.footer
            footer.is_linked_to_previous = False
            # Clear existing footer content
            for p in footer.paragraphs:
                p_elem = p._element
                p_elem.getparent().remove(p_elem)

            page_num_p = _create_page_number_footer_xml(
                config.font_name, config.font_size_pt
            )
            footer._element.append(page_num_p)

            # First page footer: empty (no page number on title page)
            first_footer = section.first_page_footer
            first_footer.is_linked_to_previous = False
            for p in first_footer.paragraphs:
                p_elem = p._element
                p_elem.getparent().remove(p_elem)
        else:
            # Subsequent sections: link footer to previous
            footer = section.footer
            footer.is_linked_to_previous = True

    logger.info("Page numbers configured: centered bottom, no number on title page.")
