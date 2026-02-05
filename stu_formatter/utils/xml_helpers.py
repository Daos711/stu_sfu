"""XML/OOXML helper utilities for working with .docx internals."""

from lxml import etree

# OOXML namespaces
NAMESPACES = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "pic": "http://schemas.openxmlformats.org/drawingml/2006/picture",
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "wps": "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
}


def qn(tag: str) -> str:
    """Convert a namespace-prefixed tag to Clark notation.

    Example: qn('w:p') -> '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p'
    """
    prefix, local = tag.split(":")
    uri = NAMESPACES[prefix]
    return f"{{{uri}}}{local}"


def find_elements(element, xpath: str):
    """Find elements using namespace-aware XPath."""
    return element.findall(xpath, NAMESPACES)


def find_element(element, xpath: str):
    """Find a single element using namespace-aware XPath."""
    return element.find(xpath, NAMESPACES)


def has_math_element(paragraph) -> bool:
    """Check if a paragraph contains math (formula) elements."""
    xml = paragraph._element
    math_paras = find_elements(xml, ".//m:oMathPara")
    math_inline = find_elements(xml, ".//m:oMath")
    return len(math_paras) > 0 or len(math_inline) > 0


def has_omathpara(paragraph) -> bool:
    """Check if a paragraph contains an oMathPara (standalone formula)."""
    xml = paragraph._element
    return len(find_elements(xml, ".//m:oMathPara")) > 0


def has_inline_omath(paragraph) -> bool:
    """Check if paragraph has inline oMath (but not oMathPara)."""
    xml = paragraph._element
    has_para = len(find_elements(xml, ".//m:oMathPara")) > 0
    has_inline = len(find_elements(xml, ".//m:oMath")) > 0
    return has_inline and not has_para


def has_drawing(paragraph) -> bool:
    """Check if a paragraph contains a drawing (image/shape)."""
    xml = paragraph._element
    drawings = find_elements(xml, ".//w:drawing")
    return len(drawings) > 0


def get_paragraph_text_without_math(paragraph) -> str:
    """Get text content of a paragraph, excluding math elements."""
    text_parts = []
    for run in paragraph.runs:
        run_xml = run._element
        # Skip runs that are inside math elements
        parent = run_xml.getparent()
        is_math = False
        while parent is not None:
            if parent.tag == qn("m:oMath") or parent.tag == qn("m:oMathPara"):
                is_math = True
                break
            parent = parent.getparent()
        if not is_math:
            text_parts.append(run.text or "")
    return "".join(text_parts).strip()


def paragraph_is_empty(paragraph) -> bool:
    """Check if a paragraph is effectively empty (no text, no images, no formulas)."""
    if paragraph.text and paragraph.text.strip():
        return False
    for run in paragraph.runs:
        if run.text and run.text.strip():
            return False
    if has_math_element(paragraph):
        return False
    if has_drawing(paragraph):
        return False
    return True
