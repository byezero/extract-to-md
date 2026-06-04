from __future__ import annotations

from pathlib import Path
import zipfile
from xml.sax.saxutils import escape


def write_docx(path: Path, title: str, paragraph: str, table_rows: list[list[str]]) -> None:
    def paragraph_xml(text: str) -> str:
        return f"<w:p><w:r><w:t>{escape(text)}</w:t></w:r></w:p>"

    def table_xml(rows: list[list[str]]) -> str:
        row_xml = []
        for row in rows:
            cells = "".join(f"<w:tc>{paragraph_xml(cell)}</w:tc>" for cell in row)
            row_xml.append(f"<w:tr>{cells}</w:tr>")
        return f"<w:tbl>{''.join(row_xml)}</w:tbl>"

    document = (
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        "<w:body>"
        f"{paragraph_xml(title)}"
        f"{paragraph_xml(paragraph)}"
        f"{table_xml(table_rows)}"
        "</w:body>"
        "</w:document>"
    )
    with zipfile.ZipFile(path, "w") as z:
        z.writestr("word/document.xml", document)


def write_xlsx(path: Path, rows: list[list[str]]) -> None:
    cells = []
    for row_idx, row in enumerate(rows, start=1):
        cell_xml = []
        for col_idx, value in enumerate(row):
            column = chr(ord("A") + col_idx)
            cell_xml.append(
                f'<c r="{column}{row_idx}" t="inlineStr"><is><t>{escape(value)}</t></is></c>'
            )
        cells.append(f'<row r="{row_idx}">{"".join(cell_xml)}</row>')

    with zipfile.ZipFile(path, "w") as z:
        z.writestr(
            "xl/workbook.xml",
            (
                '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
                'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
                '<sheets><sheet name="Quality Sheet" sheetId="1" r:id="rId1"/></sheets>'
                "</workbook>"
            ),
        )
        z.writestr(
            "xl/_rels/workbook.xml.rels",
            (
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                '<Relationship Id="rId1" '
                'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
                'Target="worksheets/sheet1.xml"/>'
                "</Relationships>"
            ),
        )
        z.writestr(
            "xl/worksheets/sheet1.xml",
            (
                '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
                f'<sheetData>{"".join(cells)}</sheetData>'
                "</worksheet>"
            ),
        )


def write_pptx(path: Path, slide_texts: list[list[str]], notes: list[list[str]] | None = None) -> None:
    def slide_xml(lines: list[str]) -> str:
        paragraphs = []
        for line in lines:
            paragraphs.append(f"<a:p><a:r><a:t>{escape(line)}</a:t></a:r></a:p>")
        return (
            '<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" '
            'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
            "<p:cSld><p:spTree>"
            f"{''.join(paragraphs)}"
            "</p:spTree></p:cSld>"
            "</p:sld>"
        )

    notes = notes or []
    with zipfile.ZipFile(path, "w") as z:
        for idx, lines in enumerate(slide_texts, start=1):
            z.writestr(f"ppt/slides/slide{idx}.xml", slide_xml(lines))
        for idx, lines in enumerate(notes, start=1):
            z.writestr(f"ppt/notesSlides/notesSlide{idx}.xml", slide_xml(lines))


def write_text_pdf(path: Path, text: str) -> None:
    write_text_pdf_pages(path, [text])


def write_text_pdf_pages(path: Path, pages: list[str]) -> None:
    objects = [
        b"<< /Type /Catalog /Pages 2 0 R >>",
    ]
    page_object_numbers = []
    content_object_numbers = []

    next_object = 3
    for text in pages:
        page_object_numbers.append(next_object)
        content_object_numbers.append(next_object + 1)
        next_object += 2
    font_object_number = next_object

    kids = " ".join(f"{number} 0 R" for number in page_object_numbers)
    objects.append(f"<< /Type /Pages /Kids [{kids}] /Count {len(pages)} >>".encode("ascii"))

    for page_object_number, content_object_number, text in zip(
        page_object_numbers,
        content_object_numbers,
        pages,
    ):
        objects.append(
            (
                "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] "
                f"/Resources << /Font << /F1 {font_object_number} 0 R >> >> "
                f"/Contents {content_object_number} 0 R >>"
            ).encode("ascii")
        )

        escaped = text.replace("\\", "\\\\").replace("(", "\\(").replace(")", "\\)")
        stream = f"BT /F1 18 Tf 72 720 Td ({escaped}) Tj ET".encode("ascii")
        objects.append(
            b"<< /Length "
            + str(len(stream)).encode("ascii")
            + b" >>\nstream\n"
            + stream
            + b"\nendstream"
        )

    objects.append(b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>")

    chunks = [b"%PDF-1.4\n"]
    offsets = []
    for idx, obj in enumerate(objects, start=1):
        offsets.append(sum(len(chunk) for chunk in chunks))
        chunks.append(f"{idx} 0 obj\n".encode("ascii"))
        chunks.append(obj)
        chunks.append(b"\nendobj\n")

    xref_offset = sum(len(chunk) for chunk in chunks)
    chunks.append(f"xref\n0 {len(objects) + 1}\n".encode("ascii"))
    chunks.append(b"0000000000 65535 f \n")
    for offset in offsets:
        chunks.append(f"{offset:010d} 00000 n \n".encode("ascii"))
    chunks.append(
        (
            f"trailer\n<< /Size {len(objects) + 1} /Root 1 0 R >>\n"
            f"startxref\n{xref_offset}\n%%EOF\n"
        ).encode("ascii")
    )
    path.write_bytes(b"".join(chunks))
