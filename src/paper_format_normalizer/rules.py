from __future__ import annotations

import csv
from pathlib import Path
from typing import TypeVar

from paper_format_normalizer.model import (
    DocumentRule,
    NumberingRule,
    ParagraphRule,
    ReportSchemaField,
    RuleSet,
    SpecialObjectRule,
    TableRule,
)

T = TypeVar("T")

RULE_FILES: tuple[str, ...] = (
    "document_rules.csv",
    "paragraph_rules.csv",
    "numbering_rules.csv",
    "table_rules.csv",
    "special_object_rules.csv",
    "report_schema.csv",
)

DOCUMENT_RULE_COLUMNS = ("rule_id", "priority", "property_name", "value", "scope")
PARAGRAPH_RULE_COLUMNS = (
    "rule_id",
    "priority",
    "match_type",
    "match_value",
    "target_property",
    "target_value",
)
NUMBERING_RULE_COLUMNS = PARAGRAPH_RULE_COLUMNS
TABLE_RULE_COLUMNS = PARAGRAPH_RULE_COLUMNS
SPECIAL_OBJECT_RULE_COLUMNS = (
    "rule_id",
    "priority",
    "object_type",
    "match_type",
    "match_value",
    "target_object_type",
)
REPORT_SCHEMA_COLUMNS = ("column_name", "order", "description")


def load_rule_set(root: Path) -> RuleSet:
    required_files = [name for name in RULE_FILES if not (root / name).is_file()]
    if required_files:
        raise ValueError(f"Missing rule files: {', '.join(required_files)}")

    document_rules = _load_rule_table(
        root / "document_rules.csv",
        DocumentRule,
        DOCUMENT_RULE_COLUMNS,
        int_fields={"priority"},
    )
    paragraph_rules = _load_rule_table(
        root / "paragraph_rules.csv",
        ParagraphRule,
        PARAGRAPH_RULE_COLUMNS,
        int_fields={"priority"},
    )
    numbering_rules = _load_rule_table(
        root / "numbering_rules.csv",
        NumberingRule,
        NUMBERING_RULE_COLUMNS,
        int_fields={"priority"},
    )
    table_rules = _load_rule_table(
        root / "table_rules.csv",
        TableRule,
        TABLE_RULE_COLUMNS,
        int_fields={"priority"},
    )
    special_object_rules = _load_rule_table(
        root / "special_object_rules.csv",
        SpecialObjectRule,
        SPECIAL_OBJECT_RULE_COLUMNS,
        int_fields={"priority"},
    )
    report_schema = _load_rule_table(
        root / "report_schema.csv",
        ReportSchemaField,
        REPORT_SCHEMA_COLUMNS,
        int_fields={"order"},
    )

    document_rules.sort(key=lambda rule: (rule.priority, rule.rule_id))
    paragraph_rules.sort(key=lambda rule: (rule.priority, rule.rule_id))
    numbering_rules.sort(key=lambda rule: (rule.priority, rule.rule_id))
    table_rules.sort(key=lambda rule: (rule.priority, rule.rule_id))
    special_object_rules.sort(key=lambda rule: (rule.priority, rule.rule_id))
    report_schema.sort(key=lambda field: (field.order, field.column_name))

    return RuleSet(
        document_rules=document_rules,
        paragraph_rules=paragraph_rules,
        numbering_rules=numbering_rules,
        table_rules=table_rules,
        special_object_rules=special_object_rules,
        report_schema=report_schema,
    )


def _load_rule_table(
    path: Path,
    model: type[T],
    required_columns: tuple[str, ...],
    *,
    int_fields: set[str],
) -> list[T]:
    with path.open("r", encoding="utf-8-sig", newline="") as handle:
        reader = csv.DictReader(handle)
        _validate_columns(path, reader.fieldnames, required_columns)
        rows: list[T] = []
        seen_rule_ids: set[str] = set()
        seen_column_names: set[str] = set()
        seen_orders: set[int] = set()
        for row_number, row in enumerate(reader, start=2):
            _validate_row_shape(path, row, row_number)
            data: dict[str, object] = {}
            for column in required_columns:
                raw_value = row.get(column)
                _validate_required_value(path, column, raw_value, row_number)
                if column in int_fields:
                    try:
                        data[column] = int(raw_value)
                    except ValueError as exc:
                        raise ValueError(
                            f"Invalid integer value for column '{column}' in {path.name} row {row_number}"
                        ) from exc
                else:
                    data[column] = raw_value
            row_obj = model(**data)  # type: ignore[arg-type]
            _validate_row_identity(
                path,
                row_obj,
                row_number,
                seen_rule_ids=seen_rule_ids,
                seen_column_names=seen_column_names,
                seen_orders=seen_orders,
            )
            rows.append(row_obj)
    return rows


def _validate_required_value(
    path: Path,
    column: str,
    raw_value: str | None,
    row_number: int,
) -> None:
    if raw_value is None or raw_value == "":
        raise ValueError(
            f"Missing required value for column '{column}' in {path.name} row {row_number}"
        )
    if raw_value != raw_value.strip():
        raise ValueError(
            f"Whitespace-corrupted value for column '{column}' in {path.name} row {row_number}"
        )


def _validate_row_identity(
    path: Path,
    row_obj: object,
    row_number: int,
    *,
    seen_rule_ids: set[str],
    seen_column_names: set[str],
    seen_orders: set[int],
) -> None:
    if hasattr(row_obj, "rule_id"):
        rule_id = getattr(row_obj, "rule_id")
        if rule_id in seen_rule_ids:
            raise ValueError(
                f"Duplicate rule_id '{rule_id}' in {path.name} row {row_number}"
            )
        seen_rule_ids.add(rule_id)

    if hasattr(row_obj, "column_name"):
        column_name = getattr(row_obj, "column_name")
        if column_name in seen_column_names:
            raise ValueError(
                f"Duplicate column_name '{column_name}' in {path.name} row {row_number}"
            )
        seen_column_names.add(column_name)

    if hasattr(row_obj, "order"):
        order = getattr(row_obj, "order")
        if order in seen_orders:
            raise ValueError(
                f"Duplicate order '{order}' in {path.name} row {row_number}"
            )
        seen_orders.add(order)


def _validate_columns(
    path: Path,
    fieldnames: list[str] | None,
    required_columns: tuple[str, ...],
) -> None:
    if fieldnames is None:
        raise ValueError(
            f"Malformed headers in {path.name}: expected {', '.join(required_columns)}"
        )
    if any(name is None or name == "" for name in fieldnames):
        raise ValueError(
            f"Malformed headers in {path.name}: expected {', '.join(required_columns)}"
        )
    if len(fieldnames) != len(set(fieldnames)):
        raise ValueError(f"Duplicate header in {path.name}: {', '.join(fieldnames)}")
    if tuple(fieldnames) != required_columns:
        raise ValueError(
            f"Unexpected header schema in {path.name}: expected {', '.join(required_columns)}"
        )


def _validate_row_shape(path: Path, row: dict[str, str | None], row_number: int) -> None:
    extra_cells = row.get(None)
    if extra_cells:
        raise ValueError(
            f"Extra unmapped cells in {path.name} row {row_number}: {extra_cells}"
        )
