"""
PDF Compressor
Uses pikepdf (built on QPDF) for secure, optimized PDF compression.
Supports stream compression, object deduplication, metadata stripping,
and optional image downsampling via Pillow.

Install:
    pip install pikepdf Pillow
"""

import argparse
import io
import os
import sys
from pathlib import Path

import pikepdf
from PIL import Image


def compress_images(pdf: pikepdf.Pdf, quality: int = 75, max_dpi: int = 150) -> int:
    """
    Re-compress embedded images as JPEG at reduced quality/DPI.
    Returns the number of images processed.
    """
    count = 0
    for page in pdf.pages:
        if "/Resources" not in page:
            continue
        resources = page["/Resources"]
        if "/XObject" not in resources:
            continue
        xobjects = resources["/XObject"]
        for key in list(xobjects.keys()):
            xobj = xobjects[key]
            try:
                if xobj.get("/Subtype") != "/Image":
                    continue
                # Skip images with soft masks or special color spaces
                if "/SMask" in xobj or "/Mask" in xobj:
                    continue
                colorspace = xobj.get("/ColorSpace")
                if colorspace in ("/DeviceCMYK", "/Separation", "/DeviceN"):
                    continue

                raw = xobj.read_raw_bytes()
                img = Image.open(io.BytesIO(raw))

                # Downsample if DPI metadata is present and exceeds max_dpi
                dpi_info = img.info.get("dpi")
                if dpi_info:
                    orig_dpi = max(dpi_info)
                    if orig_dpi > max_dpi:
                        scale = max_dpi / orig_dpi
                        new_size = (
                            max(1, int(img.width * scale)),
                            max(1, int(img.height * scale)),
                        )
                        img = img.resize(new_size, Image.LANCZOS)

                # Convert to RGB for JPEG (avoid palette/RGBA issues)
                if img.mode in ("RGBA", "P", "LA"):
                    img = img.convert("RGB")
                elif img.mode not in ("RGB", "L"):
                    img = img.convert("RGB")

                buf = io.BytesIO()
                img.save(buf, format="JPEG", quality=quality, optimize=True)
                buf.seek(0)
                compressed = buf.read()

                # Only replace if actually smaller
                if len(compressed) < len(raw):
                    xobj.write(
                        compressed,
                        filter=pikepdf.Name("/DCTDecode"),
                    )
                    count += 1

            except Exception:
                # Skip images that cannot be processed
                continue
    return count


def compress_pdf(
    input_path: Path,
    output_path: Path,
    *,
    compress_images_flag: bool = False,
    image_quality: int = 75,
    max_dpi: int = 150,
    strip_metadata: bool = True,
    password: str = "",
) -> dict:
    """
    Compress a PDF and write to output_path.
    Returns a dict with compression statistics.
    """
    open_kwargs = {}
    if password:
        open_kwargs["password"] = password

    with pikepdf.open(input_path, **open_kwargs) as pdf:
        images_compressed = 0

        if compress_images_flag:
            images_compressed = compress_images(pdf, quality=image_quality, max_dpi=max_dpi)

        if strip_metadata:
            with pdf.open_metadata() as meta:
                meta.clear()

        pdf.save(
            output_path,
            compress_streams=True,
            object_stream_mode=pikepdf.ObjectStreamMode.generate,
            recompress_flate=True,
            linearize=False,      # linearization adds size; omit unless needed for web
        )

    orig_size = input_path.stat().st_size
    new_size = output_path.stat().st_size
    saved = orig_size - new_size
    ratio = (saved / orig_size * 100) if orig_size else 0.0

    return {
        "original_bytes": orig_size,
        "compressed_bytes": new_size,
        "saved_bytes": saved,
        "reduction_pct": ratio,
        "images_recompressed": images_compressed,
    }


def format_size(n: int) -> str:
    for unit in ("B", "KB", "MB", "GB"):
        if n < 1024:
            return f"{n:.1f} {unit}"
        n /= 1024
    return f"{n:.1f} TB"


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Compress a PDF using pikepdf (QPDF backend).",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    parser.add_argument("input", type=Path, help="Input PDF file")
    parser.add_argument(
        "output",
        type=Path,
        nargs="?",
        help="Output PDF file (default: <input>_compressed.pdf)",
    )
    parser.add_argument(
        "--images",
        action="store_true",
        default=False,
        help="Re-compress embedded images as JPEG (lossy)",
    )
    parser.add_argument(
        "--quality",
        type=int,
        default=75,
        metavar="1-95",
        help="JPEG quality for image compression (requires --images)",
    )
    parser.add_argument(
        "--max-dpi",
        type=int,
        default=150,
        metavar="DPI",
        help="Downsample images above this DPI (requires --images)",
    )
    parser.add_argument(
        "--keep-metadata",
        action="store_true",
        default=False,
        help="Keep document metadata (author, title, etc.)",
    )
    parser.add_argument(
        "--password",
        default="",
        metavar="PWD",
        help="Password for encrypted PDFs",
    )

    args = parser.parse_args()

    if not args.input.exists():
        print(f"Error: file not found: {args.input}", file=sys.stderr)
        sys.exit(1)

    if not 1 <= args.quality <= 95:
        print("Error: --quality must be between 1 and 95", file=sys.stderr)
        sys.exit(1)

    output = args.output or args.input.with_stem(args.input.stem + "_compressed")

    if output.resolve() == args.input.resolve():
        print("Error: output path must differ from input", file=sys.stderr)
        sys.exit(1)

    print(f"Input:  {args.input}  ({format_size(args.input.stat().st_size)})")
    print(f"Output: {output}")

    stats = compress_pdf(
        args.input,
        output,
        compress_images_flag=args.images,
        image_quality=args.quality,
        max_dpi=args.max_dpi,
        strip_metadata=not args.keep_metadata,
        password=args.password,
    )

    sign = "-" if stats["saved_bytes"] >= 0 else "+"
    print(
        f"\nResult: {format_size(stats['compressed_bytes'])}  "
        f"({sign}{format_size(abs(stats['saved_bytes']))} / "
        f"{abs(stats['reduction_pct']):.1f}% {'reduction' if stats['saved_bytes'] >= 0 else 'increase'})"
    )
    if args.images:
        print(f"Images re-compressed: {stats['images_recompressed']}")


if __name__ == "__main__":
    main()
