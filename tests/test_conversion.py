from __future__ import annotations

from pathlib import Path

import pytest

from paper_format_normalizer.conversion import (
    Phase1ConversionError,
    _powershell_literal,
    prepare_phase1_input,
)


def test_prepare_phase1_input_returns_docx_unchanged(tmp_path: Path) -> None:
    input_path = tmp_path / "paper.docx"
    input_path.write_text("docx", encoding="utf-8")

    assert prepare_phase1_input(input_path) == input_path


def test_prepare_phase1_input_converts_doc_via_converter(tmp_path: Path, monkeypatch) -> None:
    input_path = tmp_path / "paper.doc"
    input_path.write_bytes(b"doc")
    converted_path = tmp_path / "paper.docx"
    converted_path.write_text("converted", encoding="utf-8")
    calls: list[Path] = []

    def fake_convert(path: Path) -> Path:
        calls.append(path)
        return converted_path

    monkeypatch.setattr("paper_format_normalizer.conversion._convert_doc_to_docx", fake_convert)

    result = prepare_phase1_input(input_path)

    assert result == converted_path
    assert calls == [input_path]


def test_prepare_phase1_input_rejects_pdf(tmp_path: Path) -> None:
    input_path = tmp_path / "paper.pdf"
    input_path.write_bytes(b"%PDF")

    with pytest.raises(Phase1ConversionError, match="phase-1 conversion is not configured for '.pdf' inputs"):
        prepare_phase1_input(input_path)


def test_powershell_literal_escapes_single_quotes() -> None:
    assert _powershell_literal("C:\\Users\\admin\\O'Brien\\paper.doc") == "C:\\Users\\admin\\O''Brien\\paper.doc"
