"""
JPG (or any images) to PDF combiner.
Uses PyMuPDF — already installed for pdf_to_jpg.py.
"""

import argparse
import sys
from pathlib import Path

import fitz


def images_to_pdf(image_paths: list[Path], output_path: Path) -> None:
    pdf = fitz.open()
    for img_path in image_paths:
        img_doc = fitz.open(img_path)
        pdfbytes = img_doc.convert_to_pdf()
        img_doc.close()
        page_pdf = fitz.open("pdf", pdfbytes)
        pdf.insert_pdf(page_pdf)
    pdf.save(output_path)
    pdf.close()
    print(f"Saved: {output_path}  ({len(image_paths)} page(s))")


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Combine JPG/PNG images into a single PDF.",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    parser.add_argument(
        "images", nargs="+", type=Path,
        help="Image files to combine (order matters)",
    )
    parser.add_argument(
        "--output", "-o", type=Path, default=Path("output.pdf"),
        help="Output PDF path",
    )
    args = parser.parse_args()

    missing = [p for p in args.images if not p.exists()]
    if missing:
        for p in missing:
            print(f"Error: file not found: {p}", file=sys.stderr)
        sys.exit(1)

    images_to_pdf(args.images, args.output)


if __name__ == "__main__":
    main()
