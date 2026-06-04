from __future__ import annotations

import re
import shutil
import subprocess
from pathlib import Path


INSTALL_HINTS = {
    "markitdown": "Install with: uv tool install 'markitdown[all]'",
    "pdftotext": "Install Poppler. Windows: winget install -e --id oschwartz10612.Poppler",
    "pdftoppm": "Install Poppler. Windows: winget install -e --id oschwartz10612.Poppler",
    "tesseract": "Install Tesseract OCR. Windows: winget install -e --id UB-Mannheim.TesseractOCR",
    "soffice": "Install LibreOffice. Windows: winget install -e --id TheDocumentFoundation.LibreOffice",
}

TOOL_PURPOSES = {
    "markitdown": "best-effort conversion for extra document formats",
    "pdftotext": "PDF text-layer extraction",
    "pdftoppm": "PDF page rendering for OCR",
    "tesseract": "image and scanned-document OCR",
    "soffice": "PowerPoint rendering for OCR fallback",
}


class ExtractToMdError(Exception):
    """Base class for user-facing errors."""


class MissingToolError(ExtractToMdError):
    def __init__(self, tool: str, purpose: str | None = None):
        self.tool = tool
        self.purpose = purpose or TOOL_PURPOSES.get(tool, "this conversion path")
        hint = INSTALL_HINTS.get(tool, "Install it and make sure it is available on PATH.")
        super().__init__(
            f"Required tool not found: {tool}\n"
            f"Purpose: {self.purpose}\n"
            f"{hint}\n"
            f"After installing, open a new terminal and run: {tool} --version"
        )


class CommandFailedError(ExtractToMdError):
    def __init__(self, cmd: list[str], returncode: int, stderr: str = ""):
        command = " ".join(str(part) for part in cmd)
        stderr = clean_text(stderr)
        if len(stderr) > 1200:
            stderr = stderr[:1200].rstrip() + "\n...stderr truncated..."
        message = f"Command failed with exit code {returncode}: {command}"
        if stderr:
            message += f"\nStderr:\n{stderr}"
        super().__init__(message)


def command_exists(cmd: str) -> bool:
    return shutil.which(cmd) is not None


def command_path(cmd: str) -> str | None:
    return shutil.which(cmd)


def require_tool(cmd: str, purpose: str | None = None) -> None:
    if not command_exists(cmd):
        raise MissingToolError(cmd, purpose)


def run_command(
    cmd: list[str],
    check: bool = True,
    timeout_seconds: float | None = None,
) -> subprocess.CompletedProcess[str]:
    try:
        result = subprocess.run(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=timeout_seconds,
        )
    except subprocess.TimeoutExpired:
        stderr = f"Timed out after {timeout_seconds} seconds."
        if check:
            raise CommandFailedError(cmd, -1, stderr) from None
        return subprocess.CompletedProcess(cmd, 124, "", stderr)
    if check and result.returncode != 0:
        raise CommandFailedError(cmd, result.returncode, result.stderr)
    return result


def clean_text(text: str) -> str:
    return text.replace("\r\n", "\n").replace("\r", "\n").strip()


def meaningful_len(text: str | None) -> int:
    return len(re.sub(r"\s+", "", text or ""))


def md_escape_cell(text: object) -> str:
    value = str(text or "")
    value = value.replace("|", r"\|")
    value = value.replace("\r\n", "<br>").replace("\n", "<br>").replace("\r", "<br>")
    return value


def md_table(rows: list[list[object]]) -> str:
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


def read_text_file(path: Path) -> str:
    return path.read_text(encoding="utf-8", errors="replace")
