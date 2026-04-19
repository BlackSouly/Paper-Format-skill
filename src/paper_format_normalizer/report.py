from __future__ import annotations

import csv
from collections.abc import Mapping, Sequence
from pathlib import Path

from paper_format_normalizer.model import RuleSet


def schema_columns(rule_set: RuleSet) -> list[str]:
    ordered_fields = sorted(
        rule_set.report_schema,
        key=lambda field: (field.order, field.column_name),
    )
    return [field.column_name for field in ordered_fields]


def write_report(
    report_path: Path,
    rows: Sequence[Mapping[str, str]],
    rule_set: RuleSet,
) -> Path:
    fieldnames = schema_columns(rule_set)
    report_path.parent.mkdir(parents=True, exist_ok=True)

    with report_path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=fieldnames)
        writer.writeheader()
        for row_index, row in enumerate(rows, start=1):
            _validate_row_shape(row, fieldnames, row_index)
            writer.writerow({fieldname: row[fieldname] for fieldname in fieldnames})

    return report_path


def _validate_row_shape(
    row: Mapping[str, str],
    fieldnames: Sequence[str],
    row_index: int,
) -> None:
    missing = [fieldname for fieldname in fieldnames if fieldname not in row]
    extra = sorted(set(row) - set(fieldnames))
    if missing or extra:
        details: list[str] = []
        if missing:
            details.append(f"missing columns: {', '.join(missing)}")
        if extra:
            details.append(f"extra columns: {', '.join(extra)}")
        raise ValueError(f"Invalid report row {row_index}: {'; '.join(details)}")
