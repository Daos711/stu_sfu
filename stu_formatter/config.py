"""Configuration for STU 7.5-07-2021 document formatting."""

from dataclasses import dataclass


@dataclass
class STUConfig:
    """Parameters for STU 7.5-07-2021 formatting standard."""

    # Font
    font_name: str = "Times New Roman"
    font_size_pt: int = 14
    font_size_table_pt: int = 12
    font_size_table_header_pt: int = 14
    font_size_figure_caption_pt: int = 14
    font_size_figure_notes_pt: int = 12

    # Line spacing: 1.0 = single, 1.5 = one-and-a-half
    line_spacing: float = 1.0

    # First line indent
    first_line_indent_mm: float = 12.5

    # Page margins (portrait orientation)
    margin_left_mm: float = 30.0
    margin_right_mm: float = 10.0
    margin_top_mm: float = 20.0
    margin_bottom_mm: float = 20.0

    # Page margins (landscape orientation)
    margin_landscape_left_mm: float = 20.0
    margin_landscape_right_mm: float = 20.0
    margin_landscape_top_mm: float = 30.0
    margin_landscape_bottom_mm: float = 10.0

    # Numbering: True = within sections (e.g. Table 1.1), False = continuous
    numbering_within_sections: bool = False

    # Structural elements (printed uppercase, bold, centered, new page)
    structural_elements: tuple = (
        "реферат",
        "аннотация",
        "содержание",
        "введение",
        "заключение",
        "список сокращений",
        "список использованных источников",
    )

    # Appendix letters (Russian alphabet excluding some letters)
    appendix_letters: tuple = (
        "А", "Б", "В", "Г", "Д", "Е", "Ж", "И", "К", "Л",
        "М", "Н", "П", "Р", "С", "Т", "У", "Ф", "Х", "Ц",
        "Ш", "Щ", "Э", "Ю", "Я",
    )
