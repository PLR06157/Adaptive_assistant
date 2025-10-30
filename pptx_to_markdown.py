#!/usr/bin/env python3
"""
Utility script to convert PPTX presentations to Markdown using the MarkItDown library.

Example usage:
    python3 pptx_to_markdown.py 1_Adaptive_Group_Profile.pptx --output 1_Adaptive_Group_Profile.md
    python3 pptx_to_markdown.py 2_Adaptive_Group_Strategic_Business_Guidance.pptx --output 2_Adaptive_Group_Strategic_Business_Guidance.md
    python3 pptx_to_markdown.py 3_Adaptive_Group_Transformation_Management_Office.pptx --output 3_Adaptive_Group_Transformation_Management_Office.md
    python3 pptx_to_markdown.py 4_Adaptive_Group_Service_Delivery_Enhancement.pptx --output 4_Adaptive_Group_Service_Delivery_Enhancement.md
    python3 pptx_to_markdown.py 5_Adaptive_Group_Enterprise_Automation_Hub.pptx --output 5_Adaptive_Group_Enterprise_Automation_Hub.md

If the --output argument is omitted, the Markdown is printed to stdout.
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Convert PPTX files to Markdown using MarkItDown."
    )
    parser.add_argument(
        "pptx",
        type=Path,
        help="Path to the PPTX file that should be converted.",
    )
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        help="Optional path for the generated Markdown file. Defaults to stdout.",
    )
    return parser.parse_args()


def check_dependencies() -> None:
    try:
        import markitdown  # noqa: F401
    except ModuleNotFoundError as exc:
        msg = (
            "The 'markitdown' package is required but not installed.\n"
            "Install it with: pip install markitdown"
        )
        raise SystemExit(msg) from exc


def convert_pptx_to_markdown(pptx_path: Path) -> str:
    from markitdown import MarkItDown  # type: ignore

    if not pptx_path.exists():
        raise FileNotFoundError(f"PPTX file not found: {pptx_path}")
    if pptx_path.suffix.lower() != ".pptx":
        raise ValueError(f"Expected a .pptx file, got: {pptx_path}")

    converter = MarkItDown()
    # MarkItDown expects a binary stream.
    with pptx_path.open("rb") as pptx_file:
        result = converter.convert(pptx_file)
    return result.text_content


def main() -> None:
    check_dependencies()
    args = parse_args()
    try:
        markdown_content = convert_pptx_to_markdown(args.pptx)
    except Exception as error:
        raise SystemExit(f"Conversion failed: {error}") from error

    if args.output:
        args.output.write_text(markdown_content, encoding="utf-8")
    else:
        sys.stdout.write(markdown_content)


if __name__ == "__main__":
    main()
