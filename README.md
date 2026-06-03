# extract-to-md

`extract-to-md` is a local-first command-line tool for turning documents into
Markdown that is convenient for LLM and agent workflows.

It uses plain local tools where possible, can call
[Microsoft MarkItDown](https://github.com/microsoft/markitdown) when it is
available, and adds OCR-oriented fallbacks for scanned PDFs, image-heavy slide
decks, and image files.

## Features

- Text-like files: Markdown, TXT, CSV, JSON, XML, HTML, YAML
- Word: extracts body text and tables from `.docx`
- Excel: extracts `.xlsx` worksheets as Markdown tables
- PowerPoint: extracts real `.pptx` slide text and speaker notes
- PowerPoint OCR fallback: renders image-heavy decks through LibreOffice and OCRs them
- PDF: uses `pdftotext -layout`, then OCRs sparse/scanned PDFs
- Images: OCRs PNG, JPG, WebP, BMP, and TIFF files with Tesseract
- MarkItDown fallback: uses `markitdown` for formats not handled directly

## Install

From this repository:

```powershell
uv tool install .
```

With MarkItDown support:

```powershell
uv tool install ".[markitdown]"
```

From GitHub:

```powershell
uv tool install "extract-to-md[markitdown] @ git+https://github.com/byezero/extract-to-md.git"
```

## External Tools

Some formats require external command-line tools on `PATH`:

- `pdftotext` and `pdftoppm` from Poppler for PDF text extraction and OCR rendering
- `tesseract` for image and scanned-PDF OCR
- `soffice` from LibreOffice for PowerPoint OCR fallback
- `markitdown` for best-effort conversion of additional formats

The default OCR language is `chi_sim+eng`.

## Usage

Convert a file and print Markdown to stdout:

```powershell
extract-to-md input.pdf
```

Write to a Markdown file:

```powershell
extract-to-md input.pdf -o output.md
```

Force OCR for a PDF:

```powershell
extract-to-md scanned.pdf --force-ocr -o scanned.md
```

Use a different Tesseract language:

```powershell
extract-to-md image.png --lang eng -o image.md
```

Tune OCR for sparse slide layouts:

```powershell
extract-to-md deck.pptx --psm 11 -o deck.md
```

Limit Excel extraction:

```powershell
extract-to-md workbook.xlsx --max-rows-per-sheet 200 -o workbook.md
```

## Relationship to MarkItDown

MarkItDown is a strong general-purpose converter. This tool wraps it as one
possible conversion path, then adds local extraction and OCR decisions that are
useful for document-heavy agent workflows:

- prefer layout-preserving PDF text extraction before OCR
- automatically OCR sparse/scanned PDFs
- OCR images and image-heavy PowerPoint decks with Tesseract
- keep Excel output bounded for context windows
- provide a single command with consistent Markdown output

## License

MIT
