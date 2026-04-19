from __future__ import annotations

import csv
import shutil
from pathlib import Path
from typing import Annotated

import typer

from paper_format_normalizer.conversion import (
    Phase1ConversionError,
    prepare_phase1_input,
)
from paper_format_normalizer.normalize import normalize_document
from paper_format_normalizer.rules import load_rule_set


app = typer.Typer()

WORKSPACE_ROOT = Path(__file__).resolve().parents[2]
DEFAULT_INPUT_DIR = WORKSPACE_ROOT / "\u8f93\u5165\u6587\u6863"
DEFAULT_RULES_DIR = WORKSPACE_ROOT / "\u89c4\u5219\u5316\u6587\u4ef6\u5939"
DEFAULT_OUTPUT_DIR = WORKSPACE_ROOT / "\u8f93\u51fa\u6587\u6863"
NORMALIZED_SUFFIX = "_\u89c4\u8303\u5316"
ANNOTATED_SUFFIX = "_\u89c4\u8303\u5316_\u7ea2\u5b57\u6807\u6ce8\u7248"
REPORT_SUFFIX = "_\u89c4\u8303\u5316_\u4fee\u6539\u62a5\u544a"
BATCH_SUMMARY_NAME = "\u6279\u91cf\u89c4\u8303\u5316\u6c47\u603b.csv"


@app.callback()
def main_command() -> None:
    """Paper format normalization commands."""


@app.command()
def normalize(
    input_path: Annotated[
        Path,
        typer.Option(
            "--input",
            exists=True,
            dir_okay=False,
            file_okay=True,
            readable=True,
            resolve_path=True,
            help=(
                "Phase 1 accepts .docx inputs directly. "
                ".doc inputs are automatically converted to .docx. "
                ".pdf inputs remain intentionally rejected until converters exist."
            ),
        ),
    ],
    rules_path: Annotated[
        Path,
        typer.Option(
            "--rules",
            exists=True,
            file_okay=False,
            dir_okay=True,
            readable=True,
            resolve_path=True,
            help=f"Defaults to {DEFAULT_RULES_DIR}.",
        ),
    ] = DEFAULT_RULES_DIR,
    output_dir: Annotated[
        Path,
        typer.Option(
            "--output-dir",
            file_okay=False,
            dir_okay=True,
            writable=True,
            resolve_path=True,
            help=(
                "Write the normalized DOCX and CSV report into this directory. "
                f"Defaults to {DEFAULT_OUTPUT_DIR}."
            ),
        ),
    ] = DEFAULT_OUTPUT_DIR,
) -> None:
    """Normalize a paper format.

    Phase 1 accepts .docx inputs directly.
    .doc inputs are automatically converted to .docx.
    .pdf inputs remain intentionally rejected until converters exist.
    The command writes a normalized DOCX and a CSV change report.
    """
    try:
        working_input = prepare_phase1_input(input_path)
    except Phase1ConversionError as exc:
        typer.echo(f"Error: {exc}", err=True)
        raise typer.Exit(code=1) from exc

    try:
        rule_set = load_rule_set(rules_path)
    except ValueError as exc:
        typer.echo(f"Error: {exc}", err=True)
        raise typer.Exit(code=1) from exc

    output_path, report_path, annotated_path = normalize_document(
        working_input,
        rule_set,
        output_dir,
    )
    typer.echo(f"Normalized document: {output_path}")
    typer.echo(f"Change report: {report_path}")
    typer.echo(f"Annotated document: {annotated_path}")


@app.command("normalize-batch")
def normalize_batch(
    input_dir: Annotated[
        Path,
        typer.Option(
            "--input-dir",
            exists=True,
            dir_okay=True,
            file_okay=False,
            readable=True,
            resolve_path=True,
            help=(
                "Normalize every supported document in a directory. "
                "Skip files that are already normalized outputs. "
                f"Defaults to {DEFAULT_INPUT_DIR}."
            ),
        ),
    ] = DEFAULT_INPUT_DIR,
    rules_path: Annotated[
        Path,
        typer.Option(
            "--rules",
            exists=True,
            file_okay=False,
            dir_okay=True,
            readable=True,
            resolve_path=True,
            help=f"Defaults to {DEFAULT_RULES_DIR}.",
        ),
    ] = DEFAULT_RULES_DIR,
    output_dir: Annotated[
        Path,
        typer.Option(
            "--output-dir",
            file_okay=False,
            dir_okay=True,
            writable=True,
            resolve_path=True,
            help=(
                "Write every normalized output plus a batch summary CSV into this directory. "
                f"Defaults to {DEFAULT_OUTPUT_DIR}."
            ),
        ),
    ] = DEFAULT_OUTPUT_DIR,
) -> None:
    """Normalize every supported document in a directory."""
    try:
        rule_set = load_rule_set(rules_path)
    except ValueError as exc:
        typer.echo(f"Error: {exc}", err=True)
        raise typer.Exit(code=1) from exc

    output_dir.mkdir(parents=True, exist_ok=True)
    summary_rows: list[dict[str, str]] = []

    for input_path in sorted(path for path in input_dir.iterdir() if path.is_file()):
        if _is_normalized_output(input_path):
            summary_rows.append(
                {
                    "input_name": input_path.name,
                    "status": "skipped",
                    "detail": "already normalized output",
                    "normalized_docx": "",
                    "report_csv": "",
                    "annotated_docx": "",
                }
            )
            continue

        try:
            working_input = prepare_phase1_input(input_path)
            output_path, report_path, annotated_path = normalize_document(
                working_input,
                rule_set,
                output_dir,
            )
            output_path, report_path, annotated_path = _rename_batch_outputs_to_input_stem(
                input_path=input_path,
                output_path=output_path,
                report_path=report_path,
                annotated_path=annotated_path,
            )
        except Phase1ConversionError as exc:
            summary_rows.append(
                {
                    "input_name": input_path.name,
                    "status": "failed",
                    "detail": str(exc),
                    "normalized_docx": "",
                    "report_csv": "",
                    "annotated_docx": "",
                }
            )
            continue

        summary_rows.append(
            {
                "input_name": input_path.name,
                "status": "success",
                "detail": "",
                "normalized_docx": str(output_path),
                "report_csv": str(report_path),
                "annotated_docx": str(annotated_path),
            }
        )

    summary_path = output_dir / BATCH_SUMMARY_NAME
    _write_batch_summary(summary_path, summary_rows)
    typer.echo(f"Batch summary: {summary_path}")


def _is_normalized_output(path: Path) -> bool:
    stem = path.stem
    suffix = path.suffix.lower()
    if suffix == ".docx" and (
        stem.endswith(NORMALIZED_SUFFIX) or stem.endswith(ANNOTATED_SUFFIX)
    ):
        return True
    if suffix == ".csv" and (
        stem.endswith(REPORT_SUFFIX) or path.name == BATCH_SUMMARY_NAME
    ):
        return True
    return False


def _write_batch_summary(path: Path, rows: list[dict[str, str]]) -> None:
    fieldnames = [
        "input_name",
        "status",
        "detail",
        "normalized_docx",
        "report_csv",
        "annotated_docx",
    ]
    with path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)


def _rename_batch_outputs_to_input_stem(
    *,
    input_path: Path,
    output_path: Path,
    report_path: Path,
    annotated_path: Path,
) -> tuple[Path, Path, Path]:
    if output_path.stem.startswith(input_path.stem):
        return output_path, report_path, annotated_path

    renamed_output = output_path.with_name(
        f"{input_path.stem}{NORMALIZED_SUFFIX}{output_path.suffix}"
    )
    renamed_report = report_path.with_name(
        f"{input_path.stem}{REPORT_SUFFIX}{report_path.suffix}"
    )
    renamed_annotated = annotated_path.with_name(
        f"{input_path.stem}{ANNOTATED_SUFFIX}{annotated_path.suffix}"
    )

    shutil.move(str(output_path), str(renamed_output))
    shutil.move(str(report_path), str(renamed_report))
    shutil.move(str(annotated_path), str(renamed_annotated))
    return renamed_output, renamed_report, renamed_annotated


def main() -> None:
    app()
