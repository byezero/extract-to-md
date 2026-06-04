from __future__ import annotations

from dataclasses import dataclass
import posixpath
import re
import shutil
import tempfile
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

from .utils import (
    ExtractToMdError,
    clean_text,
    command_exists,
    md_table,
    meaningful_len,
    read_text_file,
    require_tool,
    run_command,
)

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
S = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
REL = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
PKG_REL = "{http://schemas.openxmlformats.org/package/2006/relationships}"

IMAGE_EXTS = {".png", ".jpg", ".jpeg", ".webp", ".bmp", ".tif", ".tiff"}
TEXT_EXTS = {".md", ".markdown", ".txt", ".csv", ".json", ".xml", ".html", ".htm", ".yaml", ".yml"}


@dataclass(frozen=True)
class ExtractOptions:
    force_ocr: bool = False
    lang: str = "chi_sim+eng"
    dpi: int = 300
    psm: int = 6
    min_text_chars: int = 80
    max_rows_per_sheet: int = 500
    pdf_engine: str = "auto"


def format_output(path: Path, text: str) -> str:
    return f"<!-- extracted-from: {path.name} -->\n\n{clean_text(text)}\n"


def try_markitdown(path: Path) -> str:
    if not command_exists("markitdown"):
        return ""

    with tempfile.TemporaryDirectory() as td:
        out = Path(td) / "out.md"
        run_command(["markitdown", str(path), "-o", str(out)], check=False)
        if out.exists():
            text = out.read_text(encoding="utf-8", errors="replace")
            return clean_text(text)
        return ""


def extract_docx_native(path: Path) -> str:
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


def extract_docx(path: Path) -> str:
    text = try_markitdown(path)
    if meaningful_len(text) >= 20:
        return text
    return extract_docx_native(path)


def slide_num(name: str) -> int:
    m = re.search(r"slide(\d+)\.xml$", name)
    return int(m.group(1)) if m else 10**9


def extract_pptx(path: Path) -> str:
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


def col_to_index(cell_ref: str | None) -> int | None:
    m = re.match(r"([A-Z]+)", cell_ref or "")
    if not m:
        return None
    letters = m.group(1)
    n = 0
    for ch in letters:
        n = n * 26 + (ord(ch) - ord("A") + 1)
    return n - 1


def extract_xlsx(path: Path, max_rows_per_sheet: int = 500) -> str:
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


def pdf_page_count(path: Path) -> int | None:
    if not command_exists("pdfinfo"):
        return None
    result = run_command(["pdfinfo", str(path)], check=False)
    text = result.stdout or result.stderr
    match = re.search(r"^Pages:\s+(\d+)\s*$", text, flags=re.MULTILINE)
    return int(match.group(1)) if match else None


def split_pdf_text_output(text: str) -> list[str]:
    pages = [clean_text(page) for page in text.split("\f")]
    while pages and not pages[-1]:
        pages.pop()
    return pages


def extract_pdf_text_pages(path: Path, page_count: int | None = None) -> list[str]:
    require_tool("pdftotext")
    result = run_command(["pdftotext", "-layout", str(path), "-"], check=True)
    pages = split_pdf_text_output(result.stdout)
    if page_count is not None:
        if len(pages) < page_count:
            pages.extend([""] * (page_count - len(pages)))
        elif len(pages) > page_count:
            pages = pages[:page_count]
    return pages


def extract_pdf_text(path: Path) -> str:
    return clean_text("\n\n".join(extract_pdf_text_pages(path)))


def usable_pdf_text(text: str, min_text_chars: int) -> bool:
    compact_len = meaningful_len(text)
    threshold = max(12, min(min_text_chars, 24))
    if compact_len < threshold:
        return False

    useful_chars = re.findall(r"[0-9A-Za-z\u3400-\u9fff]", text)
    if len(useful_chars) < 8 and compact_len < 40:
        return False

    bad_chars = text.count("\ufffd") + text.count("□")
    if compact_len and bad_chars / compact_len > 0.1:
        return False

    control_chars = [ch for ch in text if ord(ch) < 32 and ch not in "\n\t"]
    if compact_len and len(control_chars) / compact_len > 0.05:
        return False

    return True


def ocr_image(path: Path, lang: str, psm: int) -> str:
    require_tool("tesseract")
    result = run_command(
        ["tesseract", str(path), "stdout", "-l", lang, "--psm", str(psm)],
        check=True,
    )
    return clean_text(result.stdout)


def page_key(path: Path) -> int:
    m = re.search(r"-(\d+)\.png$", path.name)
    return int(m.group(1)) if m else 10**9


def ocr_pdf(path: Path, lang: str, dpi: int, psm: int) -> str:
    require_tool("pdftoppm")
    require_tool("tesseract")

    with tempfile.TemporaryDirectory() as td:
        prefix = Path(td) / "page"
        run_command(["pdftoppm", "-r", str(dpi), "-png", str(path), str(prefix)], check=True)

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


def ocr_pdf_page(path: Path, page: int, lang: str, dpi: int, psm: int) -> str:
    require_tool("pdftoppm")
    require_tool("tesseract")

    with tempfile.TemporaryDirectory() as td:
        prefix = Path(td) / "page"
        run_command(
            [
                "pdftoppm",
                "-f",
                str(page),
                "-l",
                str(page),
                "-r",
                str(dpi),
                "-png",
                str(path),
                str(prefix),
            ],
            check=True,
        )

        pages = sorted(Path(td).glob("page-*.png"), key=page_key)
        if not pages:
            return ""
        return ocr_image(pages[0], lang=lang, psm=psm)


def pdf_report_comments(engine: str, page_methods: list[str], warnings: list[str]) -> str:
    lines = [
        f"<!-- extractor: pdf-{engine} -->",
        f"<!-- pages: {len(page_methods)} -->",
        f"<!-- page-methods: {', '.join(page_methods)} -->",
    ]
    if warnings:
        lines.append(f"<!-- warnings: {'; '.join(warnings)} -->")
    return "\n".join(lines)


def extract_pdf(path: Path, options: ExtractOptions) -> str:
    engine = "ocr" if options.force_ocr else options.pdf_engine
    if engine not in {"auto", "text", "ocr"}:
        raise ExtractToMdError(f"Unsupported PDF engine: {engine}")

    page_count = pdf_page_count(path)
    text_pages: list[str] = []
    if engine in {"auto", "text"}:
        text_pages = extract_pdf_text_pages(path, page_count=page_count)

    if page_count is None:
        page_count = len(text_pages)

    if page_count == 0 and engine in {"auto", "ocr"}:
        ocr_text = ocr_pdf(path, lang=options.lang, dpi=options.dpi, psm=options.psm)
        methods = ["ocr"] * len(re.findall(r"^## Page ", ocr_text, flags=re.MULTILINE))
        if not methods:
            methods = ["ocr"]
        report = pdf_report_comments(engine, methods, ["page count unavailable; OCR rendered all pages"])
        return clean_text(f"{report}\n\n{ocr_text}")

    out = []
    page_methods = []
    warnings = []

    for page in range(1, page_count + 1):
        text = text_pages[page - 1] if page - 1 < len(text_pages) else ""
        method = "text"

        if engine == "ocr":
            method = "ocr"
            text = ocr_pdf_page(path, page, lang=options.lang, dpi=options.dpi, psm=options.psm)
        elif engine == "auto" and not usable_pdf_text(text, options.min_text_chars):
            method = "ocr"
            warnings.append(f"page {page} text layer too sparse or noisy; used OCR")
            text = ocr_pdf_page(path, page, lang=options.lang, dpi=options.dpi, psm=options.psm)
        elif engine == "text" and not text:
            warnings.append(f"page {page} has no text-layer text")

        if not clean_text(text):
            warnings.append(f"page {page} produced no text")
            text = "_No text found._"

        page_methods.append(f"{page}:{method}")
        out.append(f"## Page {page}")
        out.append("")
        out.append(clean_text(text))
        out.append("")

    body = clean_text("\n".join(out))
    report = pdf_report_comments(engine, page_methods, warnings)
    return clean_text(f"{report}\n\n{body}")


def extract_pptx_with_ocr_fallback(path: Path, options: ExtractOptions) -> str:
    native = extract_pptx(path)

    if meaningful_len(native) >= options.min_text_chars and "_No text found._" not in native:
        return native

    if not command_exists("soffice"):
        return native

    with tempfile.TemporaryDirectory() as td:
        run_command(
            ["soffice", "--headless", "--convert-to", "pdf", "--outdir", td, str(path)],
            check=False,
        )

        pdfs = sorted(Path(td).glob("*.pdf"))
        if not pdfs:
            return native

        ocr_text = ocr_pdf(pdfs[0], lang=options.lang, dpi=options.dpi, psm=options.psm)

        if meaningful_len(ocr_text) > meaningful_len(native):
            return ocr_text

    return native


def extract_any(path: Path, options: ExtractOptions) -> str:
    ext = path.suffix.lower()

    if ext in TEXT_EXTS:
        return read_text_file(path)

    if ext == ".docx":
        return extract_docx(path)

    if ext == ".xlsx":
        return extract_xlsx(path, max_rows_per_sheet=options.max_rows_per_sheet)

    if ext == ".pptx":
        return extract_pptx_with_ocr_fallback(path, options)

    if ext == ".pdf":
        return extract_pdf(path, options)

    if ext in IMAGE_EXTS:
        return ocr_image(path, lang=options.lang, psm=options.psm)

    text = try_markitdown(path)
    if text:
        return text

    raise ExtractToMdError(f"Unsupported or unreadable file type: {ext or '(no extension)'}")
