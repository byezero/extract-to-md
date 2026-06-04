from __future__ import annotations

import argparse
import sys
from pathlib import Path

if __package__ in (None, ""):
    sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
    from extract_to_md import __version__
    from extract_to_md.diagnostics import doctor_report
    from extract_to_md.extractors import ExtractOptions, extract_any, format_output
    from extract_to_md.utils import ExtractToMdError
else:
    from . import __version__
    from .diagnostics import doctor_report
    from .extractors import ExtractOptions, extract_any, format_output
    from .utils import ExtractToMdError


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="extract-to-md",
        description="Extract document information into Markdown for local agents.",
    )
    parser.add_argument("input", nargs="?", help="Input file path")
    parser.add_argument("-o", "--output", help="Output Markdown file")
    parser.add_argument("--force-ocr", action="store_true", help="Force OCR for PDF files")
    parser.add_argument("--lang", default="chi_sim+eng", help="Tesseract language, default: chi_sim+eng")
    parser.add_argument("--dpi", type=int, default=300, help="PDF render DPI for OCR, default: 300")
    parser.add_argument("--psm", type=int, default=6, help="Tesseract PSM mode, default: 6")
    parser.add_argument("--min-text-chars", type=int, default=80, help="Minimum text-layer characters before skipping OCR")
    parser.add_argument("--max-rows-per-sheet", type=int, default=500, help="Max rows per Excel sheet")
    parser.add_argument(
        "--pdf-engine",
        choices=["auto", "text", "ocr"],
        default="auto",
        help="PDF extraction engine: auto, text, or ocr. Default: auto",
    )
    parser.add_argument("--doctor", action="store_true", help="Check local tools and OCR language availability")
    parser.add_argument("--version", action="version", version=f"extract-to-md {__version__}")
    return parser


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)

    if args.doctor:
        print(doctor_report())
        return 0

    if not args.input:
        parser.error("input is required unless --doctor or --version is used")

    path = Path(args.input)

    if not path.exists():
        print(f"extract-to-md error: Input file not found: {path}", file=sys.stderr)
        return 1

    options = ExtractOptions(
        force_ocr=args.force_ocr,
        lang=args.lang,
        dpi=args.dpi,
        psm=args.psm,
        min_text_chars=args.min_text_chars,
        max_rows_per_sheet=args.max_rows_per_sheet,
        pdf_engine=args.pdf_engine,
    )

    try:
        text = extract_any(path, options)
        md = format_output(path, text)

        if args.output:
            Path(args.output).write_text(md, encoding="utf-8")
        else:
            print(md, end="")
        return 0
    except ExtractToMdError as exc:
        print(f"extract-to-md error: {exc}", file=sys.stderr)
        return 2
    except Exception as exc:
        print(f"extract-to-md error: {exc}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
