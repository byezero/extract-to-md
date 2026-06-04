from __future__ import annotations

import subprocess
import sys
import zipfile
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from extract_to_md import __version__
from extract_to_md import cli
from extract_to_md.diagnostics import doctor_report
from extract_to_md.extractors import extract_xlsx, ocr_image
from extract_to_md.utils import CommandFailedError, MissingToolError, md_table, run_command


def test_text_file_conversion_includes_source_comment(tmp_path: Path) -> None:
    source = tmp_path / "sample.txt"
    output = tmp_path / "sample.md"
    source.write_text("Hello\n\nWorld", encoding="utf-8")

    assert cli.main([str(source), "-o", str(output)]) == 0

    assert output.read_text(encoding="utf-8") == (
        "<!-- extracted-from: sample.txt -->\n\nHello\n\nWorld\n"
    )


def test_md_table_escapes_pipes_newlines_and_skips_empty_rows() -> None:
    table = md_table(
        [
            ["Name", "Notes"],
            ["alpha|beta", "line 1\nline 2"],
            ["", ""],
        ]
    )

    assert table == (
        "| Name | Notes |\n"
        "| --- | --- |\n"
        r"| alpha\|beta | line 1<br>line 2 |"
    )


def make_xlsx(path: Path, rows: list[list[str]]) -> None:
    cells = []
    for row_idx, row in enumerate(rows, start=1):
        cell_xml = []
        for col_idx, value in enumerate(row):
            column = chr(ord("A") + col_idx)
            cell_xml.append(
                f'<c r="{column}{row_idx}" t="inlineStr"><is><t>{value}</t></is></c>'
            )
        cells.append(f'<row r="{row_idx}">{"".join(cell_xml)}</row>')

    with zipfile.ZipFile(path, "w") as z:
        z.writestr(
            "xl/workbook.xml",
            (
                '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
                'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
                '<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>'
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


def test_xlsx_truncation_message(tmp_path: Path) -> None:
    workbook = tmp_path / "book.xlsx"
    make_xlsx(workbook, [["Name"], ["One"], ["Two"]])

    text = extract_xlsx(workbook, max_rows_per_sheet=2)

    assert "| Name |" in text
    assert "One" in text
    assert "Two" not in text
    assert "_Truncated to first 2 rows._" in text


def test_missing_tool_error_contains_tool_and_install_hint(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    monkeypatch.setattr("extract_to_md.utils.shutil.which", lambda cmd: None)

    with pytest.raises(MissingToolError) as exc:
        ocr_image(tmp_path / "image.png", lang="eng", psm=6)

    message = str(exc.value)
    assert "Required tool not found: tesseract" in message
    assert "Install Tesseract OCR" in message


def test_command_failed_error_contains_command_code_and_stderr(monkeypatch: pytest.MonkeyPatch) -> None:
    def fake_run(*args, **kwargs):
        return subprocess.CompletedProcess(args[0], 7, "", "bad thing happened")

    monkeypatch.setattr("extract_to_md.utils.subprocess.run", fake_run)

    with pytest.raises(CommandFailedError) as exc:
        run_command(["tool", "--flag"], check=True)

    message = str(exc.value)
    assert "exit code 7" in message
    assert "tool --flag" in message
    assert "bad thing happened" in message


def test_version_outputs_current_version(capsys: pytest.CaptureFixture[str]) -> None:
    with pytest.raises(SystemExit) as exc:
        cli.main(["--version"])

    assert exc.value.code == 0
    assert __version__ == "0.4.0"
    assert "extract-to-md 0.4.0" in capsys.readouterr().out


def test_pdf_engine_argument_is_accepted(tmp_path: Path) -> None:
    source = tmp_path / "sample.txt"
    source.write_text("Hello", encoding="utf-8")

    assert cli.main([str(source), "--pdf-engine", "ocr"]) == 0


def test_doctor_reports_ok_and_warn(monkeypatch: pytest.MonkeyPatch) -> None:
    available = {"markitdown", "tesseract"}

    def fake_exists(cmd: str) -> bool:
        return cmd in available

    def fake_run(cmd: list[str], check: bool = True, timeout_seconds=None):
        if cmd == ["tesseract", "--list-langs"]:
            return subprocess.CompletedProcess(cmd, 0, "List of available languages\neng\n", "")
        return subprocess.CompletedProcess(cmd, 0, f"{cmd[0]} version", "")

    monkeypatch.setattr("extract_to_md.diagnostics.command_exists", fake_exists)
    monkeypatch.setattr(
        "extract_to_md.diagnostics.command_path",
        lambda cmd: f"/mock/bin/{cmd}" if cmd in available else None,
    )
    monkeypatch.setattr("extract_to_md.diagnostics.run_command", fake_run)

    report = doctor_report()

    assert "OK   extract-to-md" in report
    assert "OK   markitdown" in report
    assert "WARN pdftotext" in report
    assert "WARN lang:chi_sim" in report
    assert "OK   lang:eng" in report
