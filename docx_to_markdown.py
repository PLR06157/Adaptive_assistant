#!/usr/bin/env python3
"""
Utility script to convert DOCX documents to Markdown using the MarkItDown library.

Example usage:
    python3 docx_to_markdown.py PLAN.docx --output PLAN.md

If the --output argument is omitted, the Markdown is printed to stdout.
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path


def parse_args() -> argparse.Namespace:
    """Parsuje argumenty wiersza poleceń."""
    parser = argparse.ArgumentParser(
        description="Convert DOCX files to Markdown using MarkItDown."
    )
    parser.add_argument(
        "docx",
        type=Path,
        help="Path to the DOCX file that should be converted.",
    )
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        help="Optional path for the generated Markdown file. Defaults to stdout.",
    )
    return parser.parse_args()


def check_dependencies() -> None:
    """Sprawdza, czy pakiet 'markitdown' jest zainstalowany."""
    try:
        import markitdown  # noqa: F401
    except ModuleNotFoundError as exc:
        msg = (
            "The 'markitdown' package is required but not installed.\n"
            "Install it with: pip install markitdown[all]"
        )
        # Sugeruję instalację z [all], aby obsłużyć wszystkie formaty, w tym DOCX.
        raise SystemExit(msg) from exc


def convert_docx_to_markdown(docx_path: Path) -> str:
    """Wykonuje konwersję DOCX do Markdown."""
    from markitdown import MarkItDown  # type: ignore

    if not docx_path.exists():
        raise FileNotFoundError(f"DOCX file not found: {docx_path}")
    
    # Zmieniona walidacja, aby sprawdzać rozszerzenie .docx
    if docx_path.suffix.lower() not in [".docx", ".doc"]:
        raise ValueError(f"Expected a .docx or .doc file, got: {docx_path}")

    converter = MarkItDown()
    # MarkItDown oczekuje strumienia binarnego.
    with docx_path.open("rb") as docx_file:
        result = converter.convert(docx_file)
        
    return result.text_content


def main() -> None:
    """Główna funkcja programu."""
    check_dependencies()
    args = parse_args()
    try:
        markdown_content = convert_docx_to_markdown(args.docx)
    except Exception as error:
        # W przypadku błędu z konwersją, program się zakończy i wyświetli komunikat.
        raise SystemExit(f"Conversion failed: {error}") from error

    if args.output:
        # Zapisz zawartość Markdown do pliku
        args.output.write_text(markdown_content, encoding="utf-8")
        print(f"Sukces! Plik zapisany jako: {args.output}")
    else:
        # W przeciwnym razie, wydrukuj na standardowe wyjście
        sys.stdout.write(markdown_content)


if __name__ == "__main__":
    main()