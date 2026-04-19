from __future__ import annotations

import subprocess
from pathlib import Path


class Phase1ConversionError(ValueError):
    """Raised when a source document cannot enter the phase-1 DOCX pipeline."""


def prepare_phase1_input(input_path: Path) -> Path:
    suffix = input_path.suffix.lower()
    if suffix == ".docx":
        return input_path
    if suffix == ".doc":
        return _convert_doc_to_docx(input_path)
    if suffix == ".pdf":
        raise Phase1ConversionError(
            f"phase-1 conversion is not configured for '{suffix}' inputs; provide a .docx source"
        )

    normalized_suffix = suffix or "<none>"
    raise Phase1ConversionError(
        "phase-1 normalization supports only .docx inputs today; "
        f"received unsupported suffix '{normalized_suffix}'"
    )


def _convert_doc_to_docx(input_path: Path) -> Path:
    output_path = input_path.with_suffix(".docx")
    command = _word_doc_to_docx_command(input_path, output_path)
    result = subprocess.run(
        command,
        capture_output=True,
        text=True,
        check=False,
    )
    if result.returncode != 0 or not output_path.exists():
        stderr = (result.stderr or "").strip()
        stdout = (result.stdout or "").strip()
        detail = stderr or stdout or "Word automation did not produce a .docx file"
        raise Phase1ConversionError(
            f".doc conversion failed for '{input_path.name}': {detail}"
        )
    return output_path


def _word_doc_to_docx_command(input_path: Path, output_path: Path) -> list[str]:
    input_literal = _powershell_literal(str(input_path.resolve()))
    output_literal = _powershell_literal(str(output_path.resolve()))
    script = (
        "$ErrorActionPreference = 'Stop'\n"
        f"$inputPath = '{input_literal}'\n"
        f"$outputPath = '{output_literal}'\n"
        "$word = $null\n"
        "$document = $null\n"
        "try {\n"
        "  $word = New-Object -ComObject Word.Application\n"
        "  $word.Visible = $false\n"
        "  $word.DisplayAlerts = 0\n"
        "  $document = $word.Documents.Open($inputPath)\n"
        "  $document.SaveAs([ref]$outputPath, [ref]16)\n"
        "} finally {\n"
        "  if ($document -ne $null) { $document.Close([ref]$false) }\n"
        "  if ($word -ne $null) { $word.Quit() }\n"
        "}\n"
    )
    return [
        "powershell",
        "-NoProfile",
        "-NonInteractive",
        "-Command",
        script,
    ]


def _powershell_literal(value: str) -> str:
    return value.replace("'", "''")
