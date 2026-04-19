from dataclasses import dataclass


@dataclass(frozen=True)
class DocumentRule:
    rule_id: str
    priority: int
    property_name: str
    value: str
    scope: str


@dataclass(frozen=True)
class ParagraphRule:
    rule_id: str
    priority: int
    match_type: str
    match_value: str
    target_property: str
    target_value: str


@dataclass(frozen=True)
class NumberingRule:
    rule_id: str
    priority: int
    match_type: str
    match_value: str
    target_property: str
    target_value: str


@dataclass(frozen=True)
class TableRule:
    rule_id: str
    priority: int
    match_type: str
    match_value: str
    target_property: str
    target_value: str


@dataclass(frozen=True)
class SpecialObjectRule:
    rule_id: str
    priority: int
    object_type: str
    match_type: str
    match_value: str
    target_object_type: str


@dataclass(frozen=True)
class ReportSchemaField:
    column_name: str
    order: int
    description: str


@dataclass(frozen=True)
class RuleSet:
    document_rules: list[DocumentRule]
    paragraph_rules: list[ParagraphRule]
    numbering_rules: list[NumberingRule]
    table_rules: list[TableRule]
    special_object_rules: list[SpecialObjectRule]
    report_schema: list[ReportSchemaField]
