# extract-to-md

[中文](README.md) | English

`extract-to-md` is a local-first command-line tool for turning documents into
Markdown for LLM and agent workflows.

It can use [Microsoft MarkItDown](https://github.com/microsoft/markitdown) when
available, while adding local extraction and OCR fallbacks for PDFs, scanned
documents, images, PowerPoint decks, and bounded Excel output.

## Supported Formats

| Type | Behavior | External tools |
| --- | --- | --- |
| `.md` / `.txt` / `.csv` / `.json` / `.xml` / `.html` / `.yaml` | Read as text | None |
| `.docx` | Extract body text and tables; can prefer MarkItDown | Optional `markitdown` |
| `.xlsx` | Extract worksheets as Markdown tables | None |
| `.pptx` | Extract real slide text and speaker notes | None |
| Image-heavy `.pptx` | Render to PDF, then OCR | `soffice`, `pdftoppm`, `tesseract` |
| Text-layer `.pdf` | Use `pdftotext -layout` | Poppler `pdftotext` |
| Scanned `.pdf` | Render pages and OCR | Poppler `pdftoppm`, `tesseract` |
| Images | OCR PNG, JPG, WebP, BMP, and TIFF | `tesseract` |
| Other MarkItDown-supported formats | Best-effort fallback | Optional `markitdown` |

The default OCR language is `chi_sim+eng`.

## Install

### Minimal install

The minimal install only installs this CLI. It does not install MarkItDown,
Poppler, Tesseract, or LibreOffice.

```powershell
uv tool install git+https://github.com/byezero/extract-to-md.git
```

Use this for text files, `.docx`, `.xlsx`, and ordinary `.pptx` files that do
not need OCR.

### Full local install

On Windows, install the external tools with `winget`:

```powershell
winget install -e --id Python.Python.3.12
winget install -e --id astral-sh.uv
winget install -e --id oschwartz10612.Poppler
winget install -e --id UB-Mannheim.TesseractOCR
winget install -e --id TheDocumentFoundation.LibreOffice
```

Then open a new PowerShell and install the CLI with MarkItDown support:

```powershell
uv tool install "extract-to-md[markitdown] @ git+https://github.com/byezero/extract-to-md.git"
```

From a local clone:

```powershell
uv tool install ".[markitdown]"
```

## Diagnose Your Environment

Run:

```powershell
extract-to-md --doctor
```

It checks:

- the `extract-to-md` package
- `markitdown`
- `pdftotext` / `pdftoppm`
- `tesseract`
- OCR languages `chi_sim` / `eng`
- `soffice`

`WARN` means the capability is optional until you convert a file type that needs
it.

You can also verify commands manually:

```powershell
python --version
uv --version
pdftotext -v
pdftoppm -v
tesseract --version
tesseract --list-langs
Get-Command soffice
```

If `chi_sim` is missing, install Simplified Chinese Tesseract language data and
place `chi_sim.traineddata` in Tesseract's `tessdata` directory.

## macOS Tools

```bash
brew install python uv poppler tesseract libreoffice
brew install tesseract-lang
uv tool install "extract-to-md[markitdown] @ git+https://github.com/byezero/extract-to-md.git"
extract-to-md --doctor
```

## Linux Tools

Ubuntu / Debian:

```bash
sudo apt update
sudo apt install -y python3 python3-pip poppler-utils tesseract-ocr tesseract-ocr-chi-sim libreoffice
curl -LsSf https://astral.sh/uv/install.sh | sh
uv tool install "extract-to-md[markitdown] @ git+https://github.com/byezero/extract-to-md.git"
extract-to-md --doctor
```

## Usage

```powershell
extract-to-md --version
extract-to-md input.pdf
extract-to-md input.pdf -o output.md
extract-to-md scanned.pdf --force-ocr -o scanned.md
extract-to-md image.png --lang eng -o image.md
extract-to-md deck.pptx --psm 11 -o deck.md
extract-to-md workbook.xlsx --max-rows-per-sheet 200 -o workbook.md
```

## Common Errors

### Required tool not found: pdftotext

Install Poppler and reopen your terminal:

```powershell
winget install -e --id oschwartz10612.Poppler
pdftotext -v
```

### Required tool not found: tesseract

Install Tesseract and reopen your terminal:

```powershell
winget install -e --id UB-Mannheim.TesseractOCR
tesseract --version
tesseract --list-langs
```

### Missing chi_sim.traineddata

Install Simplified Chinese Tesseract language data. `tesseract --list-langs`
should include both `chi_sim` and `eng`.

### Image-heavy PPTX is not OCRed

Install LibreOffice so `soffice` is available:

```powershell
winget install -e --id TheDocumentFoundation.LibreOffice
Get-Command soffice
```

## Development

```powershell
git clone https://github.com/byezero/extract-to-md.git
cd extract-to-md
python src\extract_to_md\cli.py --help
python src\extract_to_md\cli.py --doctor
uv run --extra dev python -m pytest
uv build
```

## License

MIT
