# PDF Utilities — Usage Guide

A set of Python scripts for common PDF/image operations. All scripts live in the `pdf/` folder and are run from the command line.

## Requirements

```bash
pip install pymupdf pikepdf Pillow
```

| Script | Library |
|---|---|
| `pdf_to_jpg.py` | `pymupdf` |
| `jpg_to_pdf.py` | `pymupdf` |
| `pdf_merge.py` | `pymupdf` |
| `pdf_compress.py` | `pikepdf`, `Pillow` |

---

## pdf_to_jpg.py — Convert PDF pages to JPG images

Renders each page of a PDF as a separate JPEG file.

**Basic usage:**
```bash
python pdf_to_jpg.py document.pdf
```

**Options:**

| Flag | Default | Description |
|---|---|---|
| `--output-dir`, `-o` | same folder as input | Directory to save images |
| `--dpi` | `200` | Render resolution (`150`=screen, `200`=good, `300`=print) |
| `--quality` | `90` | JPEG quality (1–95) |
| `--pages` | all | Pages to convert, e.g. `1`, `1,3`, `2-5`, `1,4-6` |
| `--password` | _(none)_ | Password for encrypted PDFs |

**Examples:**
```bash
# All pages at default quality
python pdf_to_jpg.py document.pdf

# Pages 1, 3 and 5–8 at 300 DPI into a specific folder
python pdf_to_jpg.py document.pdf --pages 1,3,5-8 --dpi 300 -o ./images

# Encrypted PDF
python pdf_to_jpg.py secret.pdf --password mypassword
```

Output files are named `<stem>_p01.jpg`, `<stem>_p02.jpg`, etc.

---

## jpg_to_pdf.py — Combine images into a PDF

Combines one or more JPG/PNG (or any image) files into a single PDF. The order of files on the command line determines page order.

**Basic usage:**
```bash
python jpg_to_pdf.py image1.jpg image2.jpg image3.jpg
```

**Options:**

| Flag | Default | Description |
|---|---|---|
| `--output`, `-o` | `output.pdf` | Output PDF path |

**Examples:**
```bash
# Merge three images into output.pdf
python jpg_to_pdf.py scan1.jpg scan2.jpg scan3.jpg

# Specify a custom output name
python jpg_to_pdf.py page1.png page2.png -o combined.pdf

# Using a glob (shell expands it)
python jpg_to_pdf.py pages/*.jpg -o book.pdf
```

---

## pdf_merge.py — Merge multiple PDFs into one

Joins any number of PDF files into a single PDF. Page order follows the order of the input files.

**Basic usage:**
```bash
python pdf_merge.py file1.pdf file2.pdf file3.pdf
```

**Options:**

| Flag | Default | Description |
|---|---|---|
| `--output`, `-o` | `merged.pdf` | Output PDF path |
| `--password` | _(none)_ | Password for encrypted PDFs |

**Examples:**
```bash
# Merge two PDFs
python pdf_merge.py chapter1.pdf chapter2.pdf -o book.pdf

# Merge all PDFs in a folder (shell glob)
python pdf_merge.py parts/*.pdf -o complete.pdf

# Merge password-protected files
python pdf_merge.py a.pdf b.pdf --password secret
```

---

## pdf_compress.py — Compress a PDF

Reduces PDF file size using stream compression, object deduplication, and optional image re-compression. By default, metadata (author, title, etc.) is stripped.

**Basic usage:**
```bash
python pdf_compress.py input.pdf
```
Output is saved as `input_compressed.pdf` in the same folder unless you specify otherwise.

**Options:**

| Flag | Default | Description |
|---|---|---|
| `output` _(positional)_ | `<input>_compressed.pdf` | Output PDF path |
| `--images` | off | Re-compress embedded images as JPEG (lossy) |
| `--quality` | `75` | JPEG quality for image compression (1–95, requires `--images`) |
| `--max-dpi` | `150` | Downsample images above this DPI (requires `--images`) |
| `--keep-metadata` | off | Keep author/title/etc. metadata |
| `--password` | _(none)_ | Password for encrypted PDFs |

**Examples:**
```bash
# Basic compression (structure only, no image changes)
python pdf_compress.py report.pdf

# With a custom output name
python pdf_compress.py report.pdf report_small.pdf

# Aggressive: also re-compress images at quality 60, max 96 DPI
python pdf_compress.py report.pdf --images --quality 60 --max-dpi 96

# Preserve metadata
python pdf_compress.py report.pdf --keep-metadata

# Compress a password-protected PDF
python pdf_compress.py secret.pdf --password mypassword
```

After compression the script prints a summary:
```
Input:  report.pdf  (4.2 MB)
Output: report_compressed.pdf
Result: 1.8 MB  (-2.4 MB / 57.1% reduction)
```
