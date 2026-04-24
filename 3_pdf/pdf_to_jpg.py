"""
PDF to JPG converter
Uses PyMuPDF (pymupdf) — fast C++ MuPDF backend, no external binaries needed.

Install:
    python3 -m pip install pymupdf
"""

import argparse
import sys
from pathlib import Path

import fitz  # pymupdf


def pdf_to_jpg(
    input_path: Path,
    output_dir: Path,
    dpi: int = 200,
    quality: int = 90,
    pages: list[int] | None = None,
    password: str = "",
) -> list[Path]:
    """
    Render each page of a PDF as a JPEG image.
    Returns list of written file paths.
    """
    doc = fitz.open(input_path)

    if doc.is_encrypted:
        if not doc.authenticate(password):
            print("Error: wrong password or PDF is encrypted.", file=sys.stderr)
            sys.exit(1)

    output_dir.mkdir(parents=True, exist_ok=True)
    zoom = dpi / 72  # MuPDF base resolution is 72 DPI
    matrix = fitz.Matrix(zoom, zoom)

    total = doc.page_count
    page_indices = pages if pages else list(range(total))

    # Validate requested pages
    invalid = [p for p in page_indices if not (0 <= p < total)]
    if invalid:
        bad = ", ".join(str(p + 1) for p in invalid)
        print(f"Error: page(s) {bad} out of range (PDF has {total} page(s)).", file=sys.stderr)
        sys.exit(1)

    stem = input_path.stem
    digits = len(str(total))
    written: list[Path] = []

    for idx in page_indices:
        page = doc[idx]
        pix = page.get_pixmap(matrix=matrix, alpha=False)
        out_path = output_dir / f"{stem}_p{str(idx + 1).zfill(digits)}.jpg"
        pix.save(str(out_path), jpg_quality=quality)
        written.append(out_path)
        print(f"  Saved: {out_path.name}  ({pix.width}x{pix.height})")

    doc.close()
    return written


def parse_pages(spec: str, total: int) -> list[int]:
    """Parse page spec like '1,3,5-8' into 0-based indices."""
    indices: list[int] = []
    for part in spec.split(","):
        part = part.strip()
        if "-" in part:
            start, end = part.split("-", 1)
            indices.extend(range(int(start) - 1, int(end)))
        else:
            indices.append(int(part) - 1)
    return indices


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Convert PDF pages to JPG images using PyMuPDF.",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    parser.add_argument("input", type=Path, help="Input PDF file")
    parser.add_argument(
        "--output-dir", "-o", type=Path, default=None,
        help="Output directory (default: same folder as input PDF)",
    )
    parser.add_argument(
        "--dpi", type=int, default=200,
        help="Render resolution in DPI (150=screen, 200=good, 300=print quality)",
    )
    parser.add_argument(
        "--quality", type=int, default=90, metavar="1-95",
        help="JPEG quality",
    )
    parser.add_argument(
        "--pages", type=str, default=None, metavar="SPEC",
        help="Pages to convert, e.g. '1', '1,3', '2-5', '1,4-6' (default: all)",
    )
    parser.add_argument(
        "--password", default="", help="Password for encrypted PDFs",
    )

    args = parser.parse_args()

    if not args.input.exists():
        print(f"Error: file not found: {args.input}", file=sys.stderr)
        sys.exit(1)

    if not 1 <= args.quality <= 95:
        print("Error: --quality must be between 1 and 95", file=sys.stderr)
        sys.exit(1)

    output_dir = args.output_dir or args.input.parent

    # Need page count for parse_pages — open briefly
    doc = fitz.open(args.input)
    total = doc.page_count
    doc.close()

    pages = parse_pages(args.pages, total) if args.pages else None

    print(f"Input:  {args.input}  ({total} page(s))")
    print(f"Output: {output_dir}/  @ {args.dpi} DPI, quality {args.quality}")

    written = pdf_to_jpg(
        args.input,
        output_dir,
        dpi=args.dpi,
        quality=args.quality,
        pages=pages,
        password=args.password,
    )

    print(f"\nDone — {len(written)} image(s) written.")


if __name__ == "__main__":
    main()
