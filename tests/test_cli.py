from __future__ import annotations

import csv
import importlib.util
from pathlib import Path

from docx import Document
from typer.testing import CliRunner

from paper_format_normalizer.cli import (
    BENCHMARK_SUMMARY_NAME,
    DEFAULT_INPUT_DIR,
    DEFAULT_OUTPUT_DIR,
    DEFAULT_RULES_DIR,
    NORMALIZED_SUFFIX,
    REPORT_SUFFIX,
    app,
    normalize,
    normalize_batch,
)

_BUILDER_PATH = Path(__file__).resolve().parent / "fixtures" / "sample_docx_builder.py"
_BUILDER_SPEC = importlib.util.spec_from_file_location(
    "sample_docx_builder",
    _BUILDER_PATH,
)
if _BUILDER_SPEC is None or _BUILDER_SPEC.loader is None:
    raise RuntimeError(f"Unable to load fixture builder from {_BUILDER_PATH}")
_BUILDER_MODULE = importlib.util.module_from_spec(_BUILDER_SPEC)
_BUILDER_SPEC.loader.exec_module(_BUILDER_MODULE)
build_normalization_sample_docx = _BUILDER_MODULE.build_normalization_sample_docx


def test_help_includes_normalize():
    runner = CliRunner()

    result = runner.invoke(app, ["--help"])

    assert result.exit_code == 0
    assert "normalize" in result.output


def test_normalize_help_exits_zero():
    runner = CliRunner()

    result = runner.invoke(app, ["normalize", "--help"])

    assert result.exit_code == 0
    assert ".docx inputs directly" in result.output
    assert ".doc inputs are automatically converted" in result.output
    assert ".pdf inputs remain intentionally rejected" in result.output
    assert "normalized DOCX and a CSV change report" in result.output


def test_normalize_defaults_use_workspace_rule_and_output_directories() -> None:
    assert normalize.__defaults__ == (DEFAULT_RULES_DIR, DEFAULT_OUTPUT_DIR)


def test_normalize_command_writes_docx_and_csv(tmp_path: Path) -> None:
    runner = CliRunner()
    rules_path = Path(__file__).resolve().parent / "fixtures" / "sample_rules"
    input_path = build_normalization_sample_docx(tmp_path / "sample.docx")
    output_dir = tmp_path / "normalized"

    result = runner.invoke(
        app,
        [
            "normalize",
            "--input",
            str(input_path),
            "--rules",
            str(rules_path),
            "--output-dir",
            str(output_dir),
        ],
    )

    output_path = output_dir / f"sample{NORMALIZED_SUFFIX}.docx"
    report_path = output_dir / f"sample{REPORT_SUFFIX}.csv"

    assert result.exit_code == 0
    assert output_path.exists()
    assert report_path.exists()
    assert len(list(output_dir.glob("*.docx"))) == 2
    assert str(output_path) in result.output
    assert str(report_path) in result.output

    document = Document(output_path)
    body = next(paragraph for paragraph in document.paragraphs if paragraph.text == "Body paragraph text")

    assert body.runs[0].font.name == "Times New Roman"

    with report_path.open("r", encoding="utf-8-sig", newline="") as handle:
        rows = list(csv.DictReader(handle))

    assert rows


def test_normalize_rejects_pdf_input(tmp_path: Path) -> None:
    runner = CliRunner()
    rules_path = Path(__file__).resolve().parent / "fixtures" / "sample_rules"
    input_path = tmp_path / "paper.pdf"
    input_path.write_bytes(b"%PDF-1.4")

    result = runner.invoke(
        app,
        [
            "normalize",
            "--input",
            str(input_path),
            "--rules",
            str(rules_path),
            "--output-dir",
            str(tmp_path / "normalized"),
        ],
    )

    assert result.exit_code != 0
    assert "phase-1 conversion is not configured for '.pdf' inputs" in result.output


def test_normalize_accepts_doc_input_via_auto_conversion(tmp_path: Path, monkeypatch) -> None:
    runner = CliRunner()
    rules_path = Path(__file__).resolve().parent / "fixtures" / "sample_rules"
    input_path = tmp_path / "paper.doc"
    input_path.write_bytes(b"doc")
    converted_path = build_normalization_sample_docx(tmp_path / "paper.docx")

    monkeypatch.setattr("paper_format_normalizer.cli.prepare_phase1_input", lambda _: converted_path)

    result = runner.invoke(
        app,
        [
            "normalize",
            "--input",
            str(input_path),
            "--rules",
            str(rules_path),
            "--output-dir",
            str(tmp_path / "normalized"),
        ],
    )

    assert result.exit_code == 0
    assert str(tmp_path / "normalized" / f"paper{NORMALIZED_SUFFIX}.docx") in result.output


def test_normalize_rejects_unknown_suffix_input(tmp_path: Path) -> None:
    runner = CliRunner()
    rules_path = Path(__file__).resolve().parent / "fixtures" / "sample_rules"
    input_path = tmp_path / "paper.txt"
    input_path.write_text("plain text", encoding="utf-8")

    result = runner.invoke(
        app,
        [
            "normalize",
            "--input",
            str(input_path),
            "--rules",
            str(rules_path),
            "--output-dir",
            str(tmp_path / "normalized"),
        ],
    )

    assert result.exit_code != 0
    assert "phase-1 normalization supports only .docx inputs today" in result.output
    assert "received unsupported suffix '.txt'" in result.output


def test_normalize_batch_help_exits_zero() -> None:
    runner = CliRunner()

    result = runner.invoke(app, ["normalize-batch", "--help"])

    assert result.exit_code == 0
    assert "normalize-batch" in result.output
    assert "--input-dir" in result.output
    assert "--output-dir" in result.output


def test_normalize_batch_defaults_use_workspace_directories() -> None:
    assert normalize_batch.__defaults__ == (
        DEFAULT_INPUT_DIR,
        DEFAULT_RULES_DIR,
        DEFAULT_OUTPUT_DIR,
    )


def test_benchmark_command_writes_summary_csv(tmp_path: Path) -> None:
    runner = CliRunner()
    rules_path = Path(__file__).resolve().parent / "fixtures" / "sample_rules"
    input_path = build_normalization_sample_docx(tmp_path / "sample.docx")
    output_dir = tmp_path / "benchmark"

    result = runner.invoke(
        app,
        [
            "benchmark",
            "--input",
            str(input_path),
            "--rules",
            str(rules_path),
            "--output-dir",
            str(output_dir),
            "--repeat",
            "2",
        ],
    )

    summary_path = output_dir / BENCHMARK_SUMMARY_NAME

    assert result.exit_code == 0
    assert summary_path.exists()
    assert "Benchmark summary:" in result.output
    assert "Repeat count: 2" in result.output

    with summary_path.open("r", encoding="utf-8-sig", newline="") as handle:
        rows = list(csv.DictReader(handle))

    assert len(rows) == 1
    row = rows[0]
    assert row["repeat"] == "2"
    assert float(row["prepare_input_seconds"]) >= 0
    assert float(row["load_rules_seconds"]) >= 0
    assert float(row["total_seconds"]) >= 0
    assert row["normalized_docx"].endswith(f"sample{NORMALIZED_SUFFIX}.docx")


def test_normalize_batch_processes_supported_inputs_and_skips_normalized_outputs(
    tmp_path: Path,
    monkeypatch,
) -> None:
    runner = CliRunner()
    rules_path = Path(__file__).resolve().parent / "fixtures" / "sample_rules"
    input_dir = tmp_path / "inputs"
    input_dir.mkdir()
    output_dir = tmp_path / "normalized"

    build_normalization_sample_docx(input_dir / "first.docx")
    second_input = input_dir / "second.doc"
    second_input.write_bytes(b"doc")
    skipped_input = input_dir / f"already_normalized{NORMALIZED_SUFFIX}.docx"
    skipped_input.write_bytes(b"skip-me")
    converted_second = build_normalization_sample_docx(tmp_path / "converted-second.docx")

    def fake_prepare_phase1_input(path: Path) -> Path:
        if path == second_input:
            return converted_second
        return path

    monkeypatch.setattr("paper_format_normalizer.cli.prepare_phase1_input", fake_prepare_phase1_input)

    result = runner.invoke(
        app,
        [
            "normalize-batch",
            "--input-dir",
            str(input_dir),
            "--rules",
            str(rules_path),
            "--output-dir",
            str(output_dir),
        ],
    )

    assert result.exit_code == 0
    assert any(path.name.startswith("first_") and path.suffix == ".docx" for path in output_dir.iterdir())
    assert any(path.name.startswith("second_") and path.suffix == ".docx" for path in output_dir.iterdir())
    assert not any(
        f"already_normalized{NORMALIZED_SUFFIX}{NORMALIZED_SUFFIX}" in path.name
        for path in output_dir.iterdir()
    )

    rows = None
    for candidate in output_dir.glob("*.csv"):
        with candidate.open("r", encoding="utf-8-sig", newline="") as handle:
            candidate_rows = list(csv.DictReader(handle))
        if candidate_rows and "input_name" in candidate_rows[0]:
            rows = candidate_rows
            break

    assert rows is not None
    assert {row["input_name"] for row in rows} == {
        "first.docx",
        "second.doc",
        f"already_normalized{NORMALIZED_SUFFIX}.docx",
    }

    skipped_row = next(row for row in rows if row["input_name"] == skipped_input.name)
    assert skipped_row["status"] == "skipped"
    assert "already normalized" in skipped_row["detail"]


def test_normalize_batch_continues_when_one_file_fails(tmp_path: Path) -> None:
    runner = CliRunner()
    rules_path = Path(__file__).resolve().parent / "fixtures" / "sample_rules"
    input_dir = tmp_path / "inputs"
    input_dir.mkdir()
    output_dir = tmp_path / "normalized"

    good_input = build_normalization_sample_docx(input_dir / "good.docx")
    bad_input = input_dir / "bad.txt"
    bad_input.write_text("plain text", encoding="utf-8")

    result = runner.invoke(
        app,
        [
            "normalize-batch",
            "--input-dir",
            str(input_dir),
            "--rules",
            str(rules_path),
            "--output-dir",
            str(output_dir),
        ],
    )

    assert result.exit_code == 0
    assert any(path.name.startswith("good_") and path.suffix == ".docx" for path in output_dir.iterdir())

    rows = None
    for candidate in output_dir.glob("*.csv"):
        with candidate.open("r", encoding="utf-8-sig", newline="") as handle:
            candidate_rows = list(csv.DictReader(handle))
        if candidate_rows and "input_name" in candidate_rows[0]:
            rows = candidate_rows
            break

    assert rows is not None

    good_row = next(row for row in rows if row["input_name"] == good_input.name)
    bad_row = next(row for row in rows if row["input_name"] == bad_input.name)

    assert good_row["status"] == "success"
    assert bad_row["status"] == "failed"
    assert "unsupported suffix '.txt'" in bad_row["detail"]
