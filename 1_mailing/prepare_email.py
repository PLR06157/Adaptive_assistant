"""
prepare_email.py - Pre-process HTML email templates for Outlook Windows compatibility.

Outlook on Windows uses Microsoft Word's rendering engine, which does not reliably
honour width="100%" or height:auto on <img> tags inside percentage-based <td> cells.
Images are often rendered at their natural pixel size and clipped, cutting off the tops
of people's heads in photo galleries.

This script fixes that by:
  1. Reading the actual pixel dimensions of every local image with Pillow.
  2. Calculating the correct display width from the parent <td> width and padding.
  3. Writing explicit width/height HTML attributes onto every <img> tag.
  4. Adding valign="top" to parent <td> elements so that, in the worst case,
     any remaining clipping cuts the bottom (bodies) rather than the top (heads).

The script is idempotent: running it multiple times always produces the correct result.

Usage:
    python3 mailing/prepare_email.py --template mailing/sets/.../template.html
    python3 mailing/prepare_email.py --template mailing/sets/.../template.html --email-width 600
    python3 mailing/prepare_email.py --template mailing/sets/.../template.html --dry-run
"""

from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path

from bs4 import BeautifulSoup, Tag
from PIL import Image

EMAIL_WIDTH_DEFAULT = 600


# ---------------------------------------------------------------------------
# Width / padding helpers
# ---------------------------------------------------------------------------

def _parse_px(value: str) -> int | None:
    """Return the integer pixel value from a string like '150' or '150px', or None."""
    value = value.strip().rstrip("px").strip()
    try:
        return int(value)
    except ValueError:
        return None


def _horizontal_padding_px(td: Tag) -> int:
    """
    Return the total horizontal padding (left + right) of a <td> in pixels.
    Handles shorthand padding in the style attribute:
      padding:1px            -> 2  (1 left + 1 right)
      padding:5px 10px       -> 20 (10 left + 10 right)
      padding:5px 10px 5px   -> 20 (10 left + 10 right)
      padding:5px 10px 5px 8px -> 18 (10 right + 8 left)
    """
    style = td.get("style", "")

    # padding-left / padding-right explicit properties
    left_match = re.search(r'padding-left\s*:\s*(\d+)px', style)
    right_match = re.search(r'padding-right\s*:\s*(\d+)px', style)
    if left_match or right_match:
        left = int(left_match.group(1)) if left_match else 0
        right = int(right_match.group(1)) if right_match else 0
        return left + right

    # padding shorthand
    short = re.search(r'(?<![a-z-])padding\s*:\s*([\d\s.px]+)', style)
    if short:
        parts = short.group(1).split()
        values = []
        for p in parts:
            v = _parse_px(p)
            if v is not None:
                values.append(v)
        if len(values) == 1:       # padding: A          -> all sides = A
            return values[0] * 2
        elif len(values) == 2:     # padding: V H        -> left/right = H
            return values[1] * 2
        elif len(values) == 3:     # padding: T H B      -> left/right = H
            return values[1] * 2
        elif len(values) >= 4:     # padding: T R B L    -> left + right = L + R
            return values[1] + values[3]

    # cellpadding attribute on the parent table
    parent_table = td.find_parent("table")
    if parent_table:
        cp = parent_table.get("cellpadding", "0")
        v = _parse_px(str(cp))
        if v:
            return v * 2

    return 0


def _resolve_table_width_px(table: Tag, email_width: int) -> int:
    """
    Return the effective pixel width of a <table> element.
    Recursively resolves percentage widths by walking up to ancestor tables.
    Falls back to email_width if no width can be determined.
    """
    raw = str(table.get("width", "")).strip()

    if not raw:
        parent_table = table.find_parent("table")
        if parent_table:
            return _resolve_table_width_px(parent_table, email_width)
        return email_width

    if raw.endswith("%"):
        pct = float(raw[:-1]) / 100.0
        parent_table = table.find_parent("table")
        if parent_table:
            return int(_resolve_table_width_px(parent_table, email_width) * pct)
        return int(email_width * pct)

    v = _parse_px(raw)
    return v if v is not None else email_width


def _td_content_width_px(td: Tag, email_width: int) -> int:
    """
    Return the display content width (in pixels) available inside a <td>.
    This is: td_total_width - horizontal_padding.
    """
    raw = str(td.get("width", "")).strip()
    parent_table = td.find_parent("table")
    table_width = _resolve_table_width_px(parent_table, email_width) if parent_table else email_width

    if not raw:
        # No explicit width: divide table width equally among sibling tds
        tr = td.parent
        if tr:
            siblings = [c for c in tr.children if isinstance(c, Tag) and c.name == "td"]
            if siblings:
                td_total = table_width // len(siblings)
                return max(1, td_total - _horizontal_padding_px(td))
        return table_width

    if raw.endswith("%"):
        pct = float(raw[:-1]) / 100.0
        td_total = int(table_width * pct)
    else:
        v = _parse_px(raw)
        td_total = v if v is not None else table_width

    return max(1, td_total - _horizontal_padding_px(td))


# ---------------------------------------------------------------------------
# Main processing
# ---------------------------------------------------------------------------

def fix_gallery_images(
    html: str,
    asset_root: Path,
    email_width: int = EMAIL_WIDTH_DEFAULT,
    resize: bool = False,
) -> str:
    """
    Parse *html*, set explicit pixel width/height on every local <img>,
    and add valign="top" to the parent <td>. Returns the updated HTML string.

    When *resize* is True, also physically resize each image file to its display
    dimensions using Pillow. This is the most reliable fix for Outlook Windows,
    which sometimes ignores HTML width/height attributes and renders images at
    their natural (full) size — causing cropping and row overlap.
    """
    soup = BeautifulSoup(html, "html.parser")
    fixed = 0

    for img in soup.find_all("img"):
        src = img.get("src", "")
        # Skip remote URLs, CID references, and data URIs
        if src.startswith(("http://", "https://", "cid:", "data:")):
            continue

        image_path = Path(src) if Path(src).is_absolute() else asset_root / src
        if not image_path.exists():
            print(f"  [skip] image not found: {image_path}")
            continue

        try:
            with Image.open(image_path) as im:
                natural_w, natural_h = im.size
                img_format = im.format or image_path.suffix.lstrip(".").upper()
        except Exception as exc:
            print(f"  [skip] cannot read {image_path}: {exc}")
            continue

        if natural_w == 0 or natural_h == 0:
            continue

        # Check if the HTML already has explicit pixel dimensions (e.g. icons).
        # If so, skip HTML stamping — but still resize the file if --resize is set.
        existing_w = str(img.get("width", "")).strip()
        existing_h = str(img.get("height", "")).strip()
        already_stamped = (
            existing_w and not existing_w.endswith("%") and existing_w != "100%"
            and existing_h and not existing_h.endswith("%") and existing_h != "auto"
        )

        if already_stamped:
            try:
                target_w, target_h = int(existing_w), int(existing_h)
            except ValueError:
                print(f"  [skip] {image_path.name}: cannot parse existing dimensions ({existing_w}x{existing_h})")
                continue

            # Always apply td-level Outlook fixes (older Outlook needs height + mso styles).
            parent_td = img.find_parent("td")
            if parent_td:
                _fix_td_for_outlook(parent_td, target_h)

            if resize:
                if (natural_w, natural_h) != (target_w, target_h):
                    _resize_image(image_path, target_w, target_h, img_format)
                    print(f"  {image_path.name}: resized {natural_w}x{natural_h} -> {target_w}x{target_h}px (td fixed)")
                else:
                    print(f"  [skip] {image_path.name}: already {natural_w}x{natural_h}px on disk (td fixed)")
            else:
                print(f"  [skip] {image_path.name}: img dims ok ({existing_w}x{existing_h}), td fixed")
            continue

        parent_td = img.find_parent("td")
        if not parent_td:
            # No parent td; just stamp natural size to prevent Outlook guessing
            img["width"] = str(natural_w)
            img["height"] = str(natural_h)
            _clean_img_style(img)
            fixed += 1
            continue

        display_w = _td_content_width_px(parent_td, email_width)
        display_h = round(display_w * natural_h / natural_w)

        img["width"] = str(display_w)
        img["height"] = str(display_h)
        _clean_img_style(img)

        _fix_td_for_outlook(parent_td, display_h)

        if resize and (natural_w, natural_h) != (display_w, display_h):
            _resize_image(image_path, display_w, display_h, img_format)
            print(f"  {image_path.name}: stamped + resized {natural_w}x{natural_h} -> {display_w}x{display_h}px")
        else:
            print(f"  {image_path.name}: {natural_w}x{natural_h} -> stamped {display_w}x{display_h}px")

        fixed += 1

    print(f"  Fixed {fixed} image(s).")
    return str(soup)


def _resize_image(image_path: Path, width: int, height: int, fmt: str) -> None:
    """Resize *image_path* in-place to (*width*, *height*) using high-quality downsampling."""
    with Image.open(image_path) as im:
        # Preserve palette/transparency modes where possible
        if im.mode in ("P", "PA"):
            im = im.convert("RGBA")
        resized = im.resize((width, height), Image.LANCZOS)
        # PNG stays PNG, JPEG stays JPEG, etc.
        save_fmt = fmt if fmt in ("PNG", "JPEG", "JPG", "WEBP") else None
        if save_fmt == "JPG":
            save_fmt = "JPEG"
        resized.save(image_path, format=save_fmt)


def _clean_img_style(img: Tag) -> None:
    """
    Replace width:100% and height:auto in the img style with max-width:100%
    (keeps mobile clients happy) and ensure display:block; border:0.
    """
    style = img.get("style", "")
    style = re.sub(r'\bwidth\s*:\s*100%\s*;?\s*', '', style)
    style = re.sub(r'\bheight\s*:\s*auto\s*;?\s*', '', style)
    style = style.strip().strip(";").strip()

    # Ensure the mandatory baseline properties are present
    base = "display:block; border:0; max-width:100%;"
    parts = {p.strip() for p in style.split(";") if p.strip()}
    for part in base.rstrip(";").split(";"):
        parts.add(part.strip())
    img["style"] = " ".join(p + ";" for p in sorted(parts) if p)


def _fix_td_for_outlook(td: Tag, img_height: int) -> None:
    """
    Apply the full set of Outlook Windows (Word-engine) td fixes for reliable
    image sizing — works on Outlook 2007 through 365.

    Required attributes / styles:
      valign="top"                   – HTML attribute (CSS vertical-align ignored by Word engine)
      height="<px>"                  – HTML attribute matching the image height
      vertical-align:top             – CSS fallback for non-Outlook clients
      mso-line-height-rule:exactly   – prevents Word engine adding extra line-height
      font-size:0                    – removes the font descender gap that shifts images down
      line-height:0                  – belt-and-suspenders companion to font-size:0
    """
    td["valign"] = "top"
    td["height"] = str(img_height)

    style = td.get("style", "")

    # Collect existing declarations (excluding ones we will set explicitly)
    _MANAGED = {"vertical-align", "mso-line-height-rule", "font-size", "line-height"}
    kept = []
    for decl in style.split(";"):
        decl = decl.strip()
        if not decl:
            continue
        prop = decl.split(":")[0].strip().lower()
        if prop not in _MANAGED:
            kept.append(decl)

    kept += [
        "vertical-align:top",
        "mso-line-height-rule:exactly",
        "font-size:0",
        "line-height:0",
    ]
    td["style"] = "; ".join(kept) + ";"


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description=(
            "Pre-process HTML email templates: stamp explicit pixel dimensions on "
            "local <img> tags so Outlook Windows renders gallery photos correctly."
        )
    )
    parser.add_argument("--template", required=True, help="Path to the HTML template file.")
    parser.add_argument(
        "--email-width",
        type=int,
        default=EMAIL_WIDTH_DEFAULT,
        metavar="PX",
        help=f"Email content width in pixels (default: {EMAIL_WIDTH_DEFAULT}).",
    )
    parser.add_argument(
        "--resize",
        action="store_true",
        help=(
            "Physically resize each local image file to its display dimensions. "
            "This is the most reliable fix for Outlook Windows, which sometimes "
            "ignores HTML width/height attributes and renders images at their "
            "natural size. WARNING: overwrites the image files in place."
        ),
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Print the processed HTML to stdout without writing to disk (does not resize images).",
    )
    return parser


def main() -> int:
    parser = build_parser()
    args = parser.parse_args()

    template_path = Path(args.template)
    if not template_path.exists():
        print(f"Error: template not found: {template_path}", file=sys.stderr)
        return 1

    resize = args.resize and not args.dry_run
    if args.resize and args.dry_run:
        print("  Note: --resize is ignored in --dry-run mode.")

    print(f"Processing: {template_path}")
    html = template_path.read_text(encoding="utf-8")
    fixed_html = fix_gallery_images(html, template_path.parent, args.email_width, resize=resize)

    if args.dry_run:
        print(fixed_html)
    else:
        template_path.write_text(fixed_html, encoding="utf-8")
        print(f"Saved: {template_path}")

    return 0


if __name__ == "__main__":
    sys.exit(main())
