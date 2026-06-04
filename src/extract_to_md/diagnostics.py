from __future__ import annotations

from dataclasses import dataclass
import sys

from . import __version__
from .utils import command_exists, command_path, run_command


@dataclass(frozen=True)
class CheckResult:
    status: str
    name: str
    detail: str

    def format(self) -> str:
        return f"{self.status:<4} {self.name:<16} {self.detail}"


def _version_detail(cmd: str, args: list[str]) -> str:
    result = run_command([cmd, *args], check=False, timeout_seconds=5)
    output = (result.stdout or result.stderr or "").strip().splitlines()
    if result.returncode == 0 and output:
        return output[0].strip()
    if result.returncode == 0:
        return "available"
    if result.returncode == 124:
        return "found, but version check timed out"
    return "found, but version check failed"


def check_command(cmd: str, version_args: list[str], purpose: str) -> CheckResult:
    if not command_exists(cmd):
        return CheckResult("WARN", cmd, f"not found; needed for {purpose}")
    detail = _version_detail(cmd, version_args)
    status = "WARN" if "timed out" in detail or "failed" in detail else "OK"
    return CheckResult(status, cmd, detail)


def check_presence_only(cmd: str, purpose: str) -> CheckResult:
    path = command_path(cmd)
    if not path:
        return CheckResult("WARN", cmd, f"not found; needed for {purpose}")
    return CheckResult("OK", cmd, f"available at {path}")


def tesseract_languages() -> set[str]:
    if not command_exists("tesseract"):
        return set()
    result = run_command(["tesseract", "--list-langs"], check=False, timeout_seconds=5)
    if result.returncode != 0:
        return set()
    languages = set()
    for line in result.stdout.splitlines():
        line = line.strip()
        if not line or line.lower().startswith("list of available languages"):
            continue
        languages.add(line.replace("\\", "/").split("/")[-1])
    return languages


def run_checks() -> list[CheckResult]:
    checks = [
        CheckResult(
            "OK",
            "extract-to-md",
            f"version {__version__}; Python {sys.version_info.major}.{sys.version_info.minor}.{sys.version_info.micro}",
        ),
        check_command("markitdown", ["--version"], "best-effort conversion fallback"),
        check_command("pdftotext", ["-v"], "PDF text extraction"),
        check_command("pdftoppm", ["-v"], "PDF rendering for OCR"),
        check_command("tesseract", ["--version"], "image and scanned-document OCR"),
    ]

    languages = tesseract_languages()
    for lang in ("chi_sim", "eng"):
        if not command_exists("tesseract"):
            checks.append(CheckResult("WARN", f"lang:{lang}", "tesseract not found"))
        elif lang in languages:
            checks.append(CheckResult("OK", f"lang:{lang}", "available"))
        else:
            checks.append(CheckResult("WARN", f"lang:{lang}", "not found in tesseract --list-langs"))

    checks.append(check_presence_only("soffice", "PowerPoint OCR fallback"))
    return checks


def doctor_report() -> str:
    lines = ["extract-to-md doctor", ""]
    lines.extend(check.format() for check in run_checks())
    lines.append("")
    lines.append("WARN means the tool is optional until you convert a file type that needs it.")
    return "\n".join(lines)
