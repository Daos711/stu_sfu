"""CLI entry point for STU 7.5-07-2021 document formatter."""

import argparse
import logging
import sys
from pathlib import Path

from .config import STUConfig
from .formatter import STUFormatter


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    """Parse command-line arguments."""
    parser = argparse.ArgumentParser(
        prog="stu_formatter",
        description=(
            "Format a .docx document according to "
            "STU 7.5-07-2021 (SFU university standard)."
        ),
    )

    parser.add_argument(
        "input",
        help="Path to the input .docx file.",
    )

    parser.add_argument(
        "-o", "--output",
        default=None,
        help=(
            "Path for the output .docx file. "
            "Default: <input_name>_formatted.docx"
        ),
    )

    parser.add_argument(
        "--spacing",
        type=float,
        choices=[1.0, 1.5],
        default=1.0,
        help="Line spacing: 1.0 (single, default) or 1.5 (one-and-a-half).",
    )

    parser.add_argument(
        "--section-numbering",
        action="store_true",
        default=False,
        help=(
            "Number tables, figures, and formulas within sections "
            "(e.g., Table 1.1 instead of Table 1)."
        ),
    )

    parser.add_argument(
        "-v", "--verbose",
        action="store_true",
        default=False,
        help="Enable verbose (debug) logging.",
    )

    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    """Main entry point."""
    args = parse_args(argv)

    # Configure logging
    log_level = logging.DEBUG if args.verbose else logging.INFO
    logging.basicConfig(
        level=log_level,
        format="%(levelname)s: %(message)s",
    )

    # Validate input file
    input_path = Path(args.input)
    if not input_path.exists():
        logging.error("Input file not found: %s", input_path)
        return 1

    if input_path.suffix.lower() != ".docx":
        logging.error("Input file must be a .docx file: %s", input_path)
        return 1

    # Determine output path
    if args.output:
        output_path = Path(args.output)
    else:
        output_path = input_path.with_name(
            f"{input_path.stem}_formatted{input_path.suffix}"
        )

    # Prevent overwriting the input file
    if output_path.resolve() == input_path.resolve():
        logging.error("Output path must differ from input path.")
        return 1

    # Build config
    config = STUConfig(
        line_spacing=args.spacing,
        numbering_within_sections=args.section_numbering,
    )

    # Format
    formatter = STUFormatter(config)
    try:
        formatter.format(str(input_path), str(output_path))
    except Exception:
        logging.exception("Formatting failed.")
        return 1

    logging.info("Done. Formatted file: %s", output_path)
    return 0


if __name__ == "__main__":
    sys.exit(main())
