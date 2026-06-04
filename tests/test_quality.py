from __future__ import annotations

import subprocess
import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from extract_to_md import __version__
from extract_to_md.diagnostics import tesseract_languages
from extract_to_md.extractors import (
    ExtractOptions,
    extract_any,
    extract_docx,
    extract_pdf_text,
    extract_pptx,
    extract_pptx_with_ocr_fallback,
    extract_xlsx,
)
from extract_to_md.utils import command_exists

from fixtures import write_docx, write_pptx, write_text_pdf, write_xlsx


def test_docx_quality_semantics_without_markitdown(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    docx = tmp_path / "quality.docx"
    write_docx(
        docx,
        title="Quality Title",
        paragraph="Agent readable paragraph",
        table_rows=[["Metric", "Value"], ["Coverage", "Semantic"]],
    )
    monkeypatch.setattr("extract_to_md.extractors.try_markitdown", lambda path: "")

    text = extract_docx(docx)

    assert "Quality Title" in text
    assert "Agent readable paragraph" in text
    assert "| Metric | Value |" in text
    assert "| Coverage | Semantic |" in text


def test_xlsx_quality_semantics_and_truncation(tmp_path: Path) -> None:
    xlsx = tmp_path / "quality.xlsx"
    write_xlsx(xlsx, [["Name", "Score"], ["Alpha", "10"], ["Beta", "20"], ["Gamma", "30"]])

    text = extract_xlsx(xlsx, max_rows_per_sheet=3)

    assert "## Sheet: Quality Sheet" in text
    assert "| Name | Score |" in text
    assert "| Alpha | 10 |" in text
    assert "| Beta | 20 |" in text
    assert "Gamma" not in text
    assert "_Truncated to first 3 rows._" in text


def test_pptx_quality_semantics_for_slides_and_notes(tmp_path: Path) -> None:
    pptx = tmp_path / "quality.pptx"
    write_pptx(
        pptx,
        slide_texts=[
            ["Launch Plan", "First slide body"],
            ["Second Slide", "Follow-up details"],
        ],
        notes=[["Speaker note one"], ["Speaker note two"]],
    )

    text = extract_pptx(pptx)

    assert "## Slide 1" in text
    assert "Launch Plan" in text
    assert "First slide body" in text
    assert "## Slide 2" in text
    assert "Second Slide" in text
    assert "# Notes" in text
    assert "Speaker note one" in text
    assert "Speaker note two" in text


@pytest.mark.integration
@pytest.mark.external_tools
def test_pdf_text_layer_quality_with_poppler(tmp_path: Path) -> None:
    if not command_exists("pdftotext"):
        pytest.skip("pdftotext is not installed")
    pdf = tmp_path / "quality.pdf"
    write_text_pdf(pdf, "Quality PDF text layer")

    text = extract_pdf_text(pdf)

    assert "Quality PDF text layer" in text


def test_doctor_language_parser_accepts_tessdata_prefixed_output(monkeypatch: pytest.MonkeyPatch) -> None:
    def fake_run(cmd: list[str], check: bool = False, timeout_seconds=None):
        return subprocess.CompletedProcess(
            cmd,
            0,
            'List of available languages in "C:\\Program Files\\Tesseract-OCR/" (2):\n'
            "tessdata/chi_sim\n"
            "tessdata/eng\n",
            "",
        )

    monkeypatch.setattr("extract_to_md.diagnostics.command_exists", lambda cmd: True)
    monkeypatch.setattr("extract_to_md.diagnostics.run_command", fake_run)

    assert tesseract_languages() == {"chi_sim", "eng"}


@pytest.mark.integration
@pytest.mark.external_tools
def test_tesseract_language_detection_with_local_tool() -> None:
    if not command_exists("tesseract"):
        pytest.skip("tesseract is not installed")

    languages = tesseract_languages()

    assert "eng" in languages


@pytest.mark.integration
@pytest.mark.external_tools
def test_pptx_ocr_fallback_does_not_raise_when_libreoffice_exists(tmp_path: Path) -> None:
    if not command_exists("soffice"):
        pytest.skip("soffice is not installed")
    pptx = tmp_path / "image-heavy.pptx"
    write_pptx(pptx, slide_texts=[[]])

    text = extract_pptx_with_ocr_fallback(pptx, ExtractOptions())

    assert isinstance(text, str)
    assert "## Slide 1" in text


def test_extract_any_text_sample_still_uses_public_options(tmp_path: Path) -> None:
    sample = tmp_path / "sample.txt"
    sample.write_text("Agent quality baseline", encoding="utf-8")

    assert extract_any(sample, ExtractOptions()) == "Agent quality baseline"
    assert __version__ == "0.3.0"
