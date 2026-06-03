#!/usr/bin/env python
import argparse
import os
import posixpath
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
S = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
REL = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
PKG_REL = "{http://schemas.openxmlformats.org/package/2006/relationships}"

IMAGE_EXTS = {".png", ".jpg", ".jpeg", ".webp", ".bmp", ".tif", ".tiff"}
TEXT_EXTS = {".md", ".markdown", ".txt", ".csv", ".json", ".xml", ".html", ".htm", ".yaml", ".yml"}


def run(cmd, check=True):
    return subprocess.run(
        cmd,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
        encoding="utf-8",
        errors="replace",
        check=check,
    )


def require(cmd):
    if not shutil.which(cmd):
        raise RuntimeError(f"Required command not found in PATH: {cmd}")


def clean_text(text):
    return text.replace("\r\n", "\n").replace("\r", "\n").strip()


def meaningful_len(text):
    return len(re.sub(r"\s+", "", text or ""))


def md_escape_cell(text):
    text = str(text or "")
    text = text.replace("|", r"\|")
    text = text.replace("\r\n", "<br>").replace("\n", "<br>").replace("\r", "<br>")
    return text


def md_table(rows):
    rows = [[str(c or "") for c in row] for row in rows]
    rows = [row for row in rows if any(c.strip() for c in row)]
    if not rows:
        return ""

    width = max(len(row) for row in rows)
    rows = [row + [""] * (width - len(row)) for row in rows]

    header = rows[0]
    body = rows[1:] if len(rows) > 1 else []

    lines = []
    lines.append("| " + " | ".join(md_escape_cell(c) for c in header) + " |")
    lines.append("| " + " | ".join("---" for _ in header) + " |")
    for row in body:
        lines.append("| " + " | ".join(md_escape_cell(c) for c in row) + " |")
    return "\n".join(lines)


def read_text_file(path):
    return Path(path).read_text(encoding="utf-8", errors="replace")


def try_markitdown(path):
    if not shutil.which("markitdown"):
        return ""

    with tempfile.TemporaryDirectory() as td:
        out = Path(td) / "out.md"
        p = subprocess.run(
            ["markitdown", str(path), "-o", str(out)],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            encoding="utf-8",
            errors="replace",
        )
        if out.exists():
            text = out.read_text(encoding="utf-8", errors="replace")
            return clean_text(text)
        return ""


def extract_docx_native(path):
    out = []
    with zipfile.ZipFile(path, "r") as z:
        xml = z.read("word/document.xml")
    root = ET.fromstring(xml)
    body = root.find(f".//{W}body")
    if body is None:
        return ""

    def para_text(p):
        return "".join(t.text or "" for t in p.findall(f".//{W}t")).strip()

    def table_rows(tbl):
        rows = []
        for tr in tbl.findall(f".//{W}tr"):
            row = []
            for tc in tr.findall(f"{W}tc"):
                parts = []
                for p in tc.findall(f".//{W}p"):
                    s = para_text(p)
                    if s:
                        parts.append(s)
                row.append("\n".join(parts).strip())
            if any(c for c in row):
                rows.append(row)
        return rows

    for child in body:
        if child.tag == f"{W}p":
            text = para_text(child)
            if text:
                out.append(text)
                out.append("")
        elif child.tag == f"{W}tbl":
            table = md_table(table_rows(child))
            if table:
                out.append(table)
                out.append("")

    return clean_text("\n".join(out))


def extract_docx(path):
    text = try_markitdown(path)
    if meaningful_len(text) >= 20:
        return text
    return extract_docx_native(path)


def slide_num(name):
    m = re.search(r"slide(\d+)\.xml$", name)
    return int(m.group(1)) if m else 10**9


def extract_pptx(path):
    out = []

    def paras_from_xml(xml_bytes):
        root = ET.fromstring(xml_bytes)
        paras = []
        for p in root.findall(f".//{A}p"):
            parts = []
            for t in p.findall(f".//{A}t"):
                if t.text:
                    parts.append(t.text)
            line = "".join(parts).strip()
            if line:
                paras.append(line)
        return paras

    with zipfile.ZipFile(path, "r") as z:
        slide_names = [
            n for n in z.namelist()
            if re.match(r"ppt/slides/slide\d+\.xml$", n)
        ]
        slide_names.sort(key=slide_num)

        for idx, name in enumerate(slide_names, start=1):
            out.append(f"## Slide {idx}")
            out.append("")
            paras = paras_from_xml(z.read(name))
            if paras:
                for p in paras:
                    out.append(p)
                    out.append("")
            else:
                out.append("_No text found._")
                out.append("")

        note_names = [
            n for n in z.namelist()
            if re.match(r"ppt/notesSlides/notesSlide\d+\.xml$", n)
        ]
        note_names.sort(key=slide_num)

        if note_names:
            out.append("# Notes")
            out.append("")
            for idx, name in enumerate(note_names, start=1):
                paras = paras_from_xml(z.read(name))
                if paras:
                    out.append(f"## Notes for Slide {idx}")
                    out.append("")
                    for p in paras:
                        out.append(p)
                        out.append("")

    return clean_text("\n".join(out))


def col_to_index(cell_ref):
    m = re.match(r"([A-Z]+)", cell_ref or "")
    if not m:
        return None
    letters = m.group(1)
    n = 0
    for ch in letters:
        n = n * 26 + (ord(ch) - ord("A") + 1)
    return n - 1


def extract_xlsx(path, max_rows_per_sheet=500):
    out = []

    with zipfile.ZipFile(path, "r") as z:
        names = set(z.namelist())

        shared = []
        if "xl/sharedStrings.xml" in names:
            root = ET.fromstring(z.read("xl/sharedStrings.xml"))
            for si in root.findall(f".//{S}si"):
                shared.append("".join(si.itertext()))

        workbook = ET.fromstring(z.read("xl/workbook.xml"))
        rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))

        rid_to_target = {}
        for rel in rels.findall(f"{PKG_REL}Relationship"):
            rid = rel.attrib.get("Id")
            target = rel.attrib.get("Target")
            if not rid or not target:
                continue
            if target.startswith("/"):
                full = target.lstrip("/")
            else:
                full = posixpath.normpath("xl/" + target)
            rid_to_target[rid] = full

        sheets = []
        for s in workbook.findall(f".//{S}sheet"):
            name = s.attrib.get("name", "Sheet")
            rid = s.attrib.get(f"{REL}id")
            target = rid_to_target.get(rid)
            if target:
                sheets.append((name, target))

        for sheet_name, sheet_path in sheets:
            if sheet_path not in names:
                continue

            root = ET.fromstring(z.read(sheet_path))
            rows_out = []

            for row in root.findall(f".//{S}sheetData/{S}row"):
                cells = {}
                max_col = -1

                for c in row.findall(f"{S}c"):
                    ref = c.attrib.get("r", "")
                    col = col_to_index(ref)
                    if col is None:
                        col = max_col + 1
                    max_col = max(max_col, col)

                    cell_type = c.attrib.get("t")
                    value = ""

                    if cell_type == "s":
                        v = c.find(f"{S}v")
                        if v is not None and v.text is not None:
                            idx = int(v.text)
                            value = shared[idx] if 0 <= idx < len(shared) else ""
                    elif cell_type == "inlineStr":
                        value = "".join(c.itertext()).strip()
                    else:
                        v = c.find(f"{S}v")
                        if v is not None and v.text is not None:
                            value = v.text

                    cells[col] = value

                if max_col >= 0:
                    row_values = [cells.get(i, "") for i in range(max_col + 1)]
                    if any(str(x).strip() for x in row_values):
                        rows_out.append(row_values)

            out.append(f"## Sheet: {sheet_name}")
            out.append("")

            if not rows_out:
                out.append("_No data found._")
                out.append("")
                continue

            truncated = False
            if len(rows_out) > max_rows_per_sheet:
                rows_out = rows_out[:max_rows_per_sheet]
                truncated = True

            out.append(md_table(rows_out))
            out.append("")

            if truncated:
                out.append(f"_Truncated to first {max_rows_per_sheet} rows._")
                out.append("")

    return clean_text("\n".join(out))


def extract_pdf_text(path):
    require("pdftotext")
    with tempfile.TemporaryDirectory() as td:
        out = Path(td) / "out.txt"
        p = subprocess.run(
            ["pdftotext", "-layout", str(path), str(out)],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            encoding="utf-8",
            errors="replace",
        )
        if out.exists():
            return clean_text(out.read_text(encoding="utf-8", errors="replace"))
        return ""


def ocr_image(path, lang, psm):
    require("tesseract")
    p = run(["tesseract", str(path), "stdout", "-l", lang, "--psm", str(psm)], check=False)
    return clean_text(p.stdout)


def page_key(path):
    m = re.search(r"-(\d+)\.png$", path.name)
    return int(m.group(1)) if m else 10**9


def ocr_pdf(path, lang, dpi, psm):
    require("pdftoppm")
    require("tesseract")

    with tempfile.TemporaryDirectory() as td:
        prefix = Path(td) / "page"
        run(["pdftoppm", "-r", str(dpi), "-png", str(path), str(prefix)], check=True)

        pages = sorted(Path(td).glob("page-*.png"), key=page_key)
        if not pages:
            return ""

        out = []
        for idx, img in enumerate(pages, start=1):
            out.append(f"## Page {idx}")
            out.append("")
            text = ocr_image(img, lang=lang, psm=psm)
            out.append(text if text else "_No OCR text found._")
            out.append("")

        return clean_text("\n".join(out))


def extract_pdf(path, args):
    if not args.force_ocr:
        text = extract_pdf_text(path)
        if meaningful_len(text) >= args.min_text_chars:
            return text
    return ocr_pdf(path, lang=args.lang, dpi=args.dpi, psm=args.psm)



def extract_pptx_with_ocr_fallback(path, args):
    native = extract_pptx(path)

    # If the PPTX contains real editable text, native extraction is preferred.
    if meaningful_len(native) >= args.min_text_chars and "_No text found._" not in native:
        return native

    # Some PPTX files, especially Marp-generated decks, store slide content
    # as images/SVG-like drawing objects. In that case we render the deck and OCR it.
    if not shutil.which("soffice"):
        return native

    with tempfile.TemporaryDirectory() as td:
        p = subprocess.run(
            ["soffice", "--headless", "--convert-to", "pdf", "--outdir", td, str(path)],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            encoding="utf-8",
            errors="replace",
        )

        pdfs = sorted(Path(td).glob("*.pdf"))
        if not pdfs:
            return native

        ocr_text = ocr_pdf(pdfs[0], lang=args.lang, dpi=args.dpi, psm=args.psm)

        if meaningful_len(ocr_text) > meaningful_len(native):
            return ocr_text

    return native


def extract_any(path, args):
    ext = path.suffix.lower()

    if ext in TEXT_EXTS:
        return read_text_file(path)

    if ext == ".docx":
        return extract_docx(path)

    if ext == ".xlsx":
        return extract_xlsx(path, max_rows_per_sheet=args.max_rows_per_sheet)

    if ext == ".pptx":
        return extract_pptx_with_ocr_fallback(path, args)

    if ext == ".pdf":
        return extract_pdf(path, args)

    if ext in IMAGE_EXTS:
        return ocr_image(path, lang=args.lang, psm=args.psm)

    text = try_markitdown(path)
    if text:
        return text

    raise RuntimeError(f"Unsupported or unreadable file type: {ext}")


def main():
    parser = argparse.ArgumentParser(
        description="Extract document information into Markdown for local agents."
    )
    parser.add_argument("input", help="Input file path")
    parser.add_argument("-o", "--output", help="Output Markdown file")
    parser.add_argument("--force-ocr", action="store_true", help="Force OCR for PDF files")
    parser.add_argument("--lang", default="chi_sim+eng", help="Tesseract language, default: chi_sim+eng")
    parser.add_argument("--dpi", type=int, default=300, help="PDF render DPI for OCR, default: 300")
    parser.add_argument("--psm", type=int, default=6, help="Tesseract PSM mode, default: 6")
    parser.add_argument("--min-text-chars", type=int, default=80, help="Minimum text-layer characters before skipping OCR")
    parser.add_argument("--max-rows-per-sheet", type=int, default=500, help="Max rows per Excel sheet")
    args = parser.parse_args()

    path = Path(args.input)

    if not path.exists():
        print(f"Input file not found: {path}", file=sys.stderr)
        return 1

    try:
        text = extract_any(path, args)
        md = f"<!-- extracted-from: {path.name} -->\n\n{clean_text(text)}\n"

        if args.output:
            Path(args.output).write_text(md, encoding="utf-8")
        else:
            print(md, end="")
        return 0
    except Exception as e:
        print(f"extract-to-md error: {e}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
