"""
Merge multiple PDFs into one.
Uses PyMuPDF — no new dependencies needed.
"""

import argparse
import sys
from pathlib import Path

import fitz


def merge_pdfs(input_paths: list[Path], output_path: Path, password: str = "") -> None:
    merged = fitz.open()
    for path in input_paths:
        doc = fitz.open(path)
        if doc.is_encrypted:
            if not doc.authenticate(password):
                print(f"Error: wrong password for {path}", file=sys.stderr)
                sys.exit(1)
        merged.insert_pdf(doc)
        print(f"  Added: {path.name}  ({doc.page_count} page(s))")
        doc.close()
    merged.save(output_path)
    merged.close()
    print(f"\nSaved: {output_path}  ({merged.page_count if False else sum(1 for _ in fitz.open(output_path))} page(s) total)")


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Merge multiple PDF files into one.",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    parser.add_argument("pdfs", nargs="+", type=Path, help="PDF files to merge (order matters)")
    parser.add_argument("--output", "-o", type=Path, default=Path("merged.pdf"), help="Output PDF path")
    parser.add_argument("--password", default="", help="Password for encrypted PDFs")
    args = parser.parse_args()

    missing = [p for p in args.pdfs if not p.exists()]
    if missing:
        for p in missing:
            print(f"Error: file not found: {p}", file=sys.stderr)
        sys.exit(1)

    merge_pdfs(args.pdfs, args.output, password=args.password)


if __name__ == "__main__":
    main()
