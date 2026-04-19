from __future__ import annotations

import re
from collections import defaultdict
from copy import deepcopy
from dataclasses import dataclass
from pathlib import Path
from typing import Literal

from docx import Document
from docx.document import Document as DocumentObject
from docx.oxml.ns import qn
from docx.section import Section
from docx.shared import Cm, Inches, Length, Pt, RGBColor
from docx.table import Table
from docx.text.paragraph import Paragraph

from paper_format_normalizer.classify import (
    _MATCH_KIND_ORDER,
    _numbering_rule_matches,
    _paragraph_rule_candidate,
    _table_rule_candidate,
    classify_document,
)
from paper_format_normalizer.model import (
    DocumentRule,
    NumberingRule,
    ParagraphRule,
    RuleSet,
    SpecialObjectRule,
    TableRule,
)
from paper_format_normalizer.parse import (
    ParsedBodyParagraph,
    ParsedBodyTable,
    _iter_body_elements,
    parse_docx,
)
from paper_format_normalizer.report import write_report

ReportStatus = Literal["modified", "unchanged", "unresolved"]
_RED = RGBColor(0xFF, 0x00, 0x00)
_TEXTUAL_PROPERTIES = frozenset(
    {
        "font_name",
        "font_size",
        "header_row_font_name",
        "header_row_font_size",
        "body_rows_font_name",
        "body_rows_font_size",
        "label_font_name",
        "content_font_name",
        "label_font_size",
        "content_font_size",
        "label_bold",
        "content_bold",
    }
)
_ANNOTATION_PROPERTIES = frozenset(
    {
        "first_line_indent",
        "hanging_indent",
        "space_before",
        "space_after",
        "line_spacing",
        "page_margin_top",
        "page_margin_bottom",
        "page_margin_left",
        "page_margin_right",
    }
)
_PROPERTY_LABELS = {
    "first_line_indent": "首行缩进",
    "hanging_indent": "悬挂缩进",
    "space_before": "段前",
    "space_after": "段后",
    "line_spacing": "行距",
    "page_margin_top": "上边距",
    "page_margin_bottom": "下边距",
    "page_margin_left": "左边距",
    "page_margin_right": "右边距",
    "label_font_name": "标签字体",
    "content_font_name": "内容字体",
    "label_font_size": "标签字号",
    "content_font_size": "内容字号",
    "label_bold": "标签加粗",
    "content_bold": "内容加粗",
    "header_row_font_name": "表头字体",
    "header_row_font_size": "表头字号",
    "body_rows_font_name": "数据行字体",
    "body_rows_font_size": "数据行字号",
}
_ANNOTATION_FONT_NAME = "SimSun"
_ANNOTATION_FONT_SIZE = Pt(10.5)
_WESTERN_FONT_NAME = "Times New Roman"
_FONT_DISPLAY_NAMES = {
    "Arial": "Arial（无衬线体）",
    "Calibri": "Calibri（默认无衬线体）",
    "KaiTi": "楷体（KaiTi）",
    "FangSong": "仿宋（FangSong）",
    "SimHei": "黑体（SimHei）",
    "SimSun": "宋体（SimSun）",
    "Times New Roman": "Times New Roman（新罗马体）",
}
_EAST_ASIAN_FONT_NAMES = frozenset({"SimHei", "SimSun", "FangSong", "KaiTi"})
_INLINE_PREFIXES = ("【摘要】", "【关键词】", "【Abstract】", "【KeyWords】")


@dataclass(frozen=True)
class _MatchedParagraphRule:
    rule_id: str
    source_family: str
    priority: int
    match_kind: str
    match_type: str
    match_value: str
    target_property: str
    target_value: str


@dataclass(frozen=True)
class _RunLengthState:
    values: tuple[Length | None, ...]


@dataclass(frozen=True)
class _MatchedTableRule:
    rule_id: str
    priority: int
    match_kind: str
    match_type: str
    match_value: str
    target_property: str
    target_value: str


def normalize_document(
    input_path: Path,
    rule_set: RuleSet,
    output_dir: Path,
) -> tuple[Path, Path, Path]:
    parsed = parse_docx(input_path)
    classification = classify_document(parsed, rule_set)
    document = Document(input_path)

    output_dir.mkdir(parents=True, exist_ok=True)
    output_path = output_dir / f"{input_path.stem}_规范化.docx"
    report_path = output_dir / f"{input_path.stem}_规范化_修改报告.csv"
    annotated_path = output_dir / f"{input_path.stem}_规范化_红字标注版.docx"

    report_rows: list[dict[str, str]] = []
    report_rows.extend(_apply_document_rules(document, rule_set.document_rules))

    handled_locations: set[str] = set()
    body_paragraphs = _body_paragraph_map(document)
    body_tables = _body_table_map(document)
    header_paragraphs = _header_paragraph_map(document)
    classification_by_location = {
        result.location: result for result in classification.object_results
    }

    for body_index, paragraph in body_paragraphs.items():
        parsed_item = parsed.body_items[body_index]
        if not isinstance(parsed_item, ParsedBodyParagraph):
            continue
        location = f"body_items[{body_index}]"
        handled_locations.add(location)
        classification_result = classification_by_location[location]
        report_rows.extend(
            _apply_paragraph_rules(
                paragraph=paragraph,
                parsed_paragraph=parsed_item,
                classification_result=classification_result,
                rule_set=rule_set,
            )
        )

    for location, paragraph in header_paragraphs.items():
        header_index, item_index = _header_location_indexes(location)
        parsed_header_item = parsed.headers[header_index].items[item_index]
        if not isinstance(parsed_header_item, ParsedBodyParagraph):
            continue
        handled_locations.add(location)
        classification_result = classification_by_location[location]
        report_rows.extend(
            _apply_header_paragraph_rules(
                paragraph=paragraph,
                classification_result=classification_result,
                rule_set=rule_set,
            )
        )

    for body_index, table in body_tables.items():
        parsed_item = parsed.body_items[body_index]
        if not isinstance(parsed_item, ParsedBodyTable):
            continue
        location = f"body_items[{body_index}]"
        handled_locations.add(location)
        classification_result = classification_by_location[location]
        report_rows.extend(
            _apply_table_rules(
                table=table,
                parsed_table=parsed_item,
                classification_result=classification_result,
                rule_set=rule_set,
            )
        )

    for result in classification.object_results:
        if result.location in handled_locations:
            continue
        report_rows.append(
            _report_row(
                object_id=result.object_id,
                object_type_before=result.object_type,
                object_type_after=result.object_type,
                location=result.location,
                text_preview=result.original_text,
                property_name="classification",
                before="",
                after="",
                rule_id=result.matched_rule_id or "",
                status="unresolved",
                reason=(
                    result.reason
                    if result.status == "unresolved"
                    else f"phase-1 normalization does not support {result.object_type} objects"
                ),
            )
        )

    document.save(output_path)
    write_report(report_path, report_rows, rule_set)
    _write_annotated_document(
        normalized_path=output_path,
        annotated_path=annotated_path,
        report_rows=report_rows,
    )
    return output_path, report_path, annotated_path


def _apply_document_rules(document: DocumentObject, rules: list[DocumentRule]) -> list[dict[str, str]]:
    rows: list[dict[str, str]] = []
    grouped_rules: dict[str, list[DocumentRule]] = defaultdict(list)
    for rule in rules:
        grouped_rules[rule.property_name].append(rule)

    for property_name, property_rules in grouped_rules.items():
        sorted_rules = sorted(property_rules, key=lambda rule: (rule.priority, rule.rule_id))
        winner = sorted_rules[0]
        conflict_rules = [
            rule for rule in sorted_rules if rule.priority == winner.priority
        ]
        if len(conflict_rules) > 1:
            conflict_ids = ", ".join(rule.rule_id for rule in conflict_rules)
            for section_index, _ in enumerate(document.sections):
                rows.append(
                    _report_row(
                        object_id=f"section-{section_index}",
                        object_type_before="document",
                        object_type_after="document",
                        location=f"sections[{section_index}]",
                        text_preview="",
                        property_name=property_name,
                        before="",
                        after="",
                        rule_id="",
                        status="unresolved",
                        reason=f"conflicting document rules at same priority: {conflict_ids}",
                    )
                )
            continue

        for section_index, section in enumerate(document.sections):
            rows.append(_apply_document_rule(section, section_index, winner))
    return rows


def _apply_document_rule(
    section: Section,
    section_index: int,
    rule: DocumentRule,
) -> dict[str, str]:
    if rule.scope != "document":
        return _report_row(
            object_id=f"section-{section_index}",
            object_type_before="document",
            object_type_after="document",
            location=f"sections[{section_index}]",
            text_preview="",
            property_name=rule.property_name,
            before="",
            after="",
            rule_id=rule.rule_id,
            status="unresolved",
            reason=f"unsupported document rule scope: {rule.scope}",
        )

    property_accessor = _document_property_accessor(rule.property_name)
    if property_accessor is None:
        return _report_row(
            object_id=f"section-{section_index}",
            object_type_before="document",
            object_type_after="document",
            location=f"sections[{section_index}]",
            text_preview="",
            property_name=rule.property_name,
            before="",
            after="",
            rule_id=rule.rule_id,
            status="unresolved",
            reason=f"unsupported document property: {rule.property_name}",
        )

    before_value = property_accessor.get(section)
    try:
        target_value = _parse_measurement(rule.value)
    except ValueError as exc:
        return _report_row(
            object_id=f"section-{section_index}",
            object_type_before="document",
            object_type_after="document",
            location=f"sections[{section_index}]",
            text_preview="",
            property_name=rule.property_name,
            before=property_accessor.format(before_value, rule.value),
            after=rule.value,
            rule_id=rule.rule_id,
            status="unresolved",
            reason=str(exc),
        )

    property_accessor.set(section, target_value)
    status: ReportStatus = "modified" if before_value != target_value else "unchanged"
    return _report_row(
        object_id=f"section-{section_index}",
        object_type_before="document",
        object_type_after="document",
        location=f"sections[{section_index}]",
        text_preview="",
        property_name=rule.property_name,
        before=property_accessor.format(before_value, rule.value),
        after=property_accessor.format(target_value, rule.value),
        rule_id=rule.rule_id,
        status=status,
        reason="",
    )


def _apply_paragraph_rules(
    *,
    paragraph: Paragraph,
    parsed_paragraph: ParsedBodyParagraph,
    classification_result,
    rule_set: RuleSet,
) -> list[dict[str, str]]:
    if classification_result.status == "unresolved":
        return [
            _report_row(
                object_id=classification_result.object_id,
                object_type_before=classification_result.object_type,
                object_type_after=classification_result.object_type,
                location=classification_result.location,
                text_preview=classification_result.original_text,
                property_name="classification",
                before="",
                after="",
                rule_id="",
                status="unresolved",
                reason=classification_result.reason or "no matching classification rule",
            )
        ]

    rows: list[dict[str, str]] = []
    matched_rules = _matched_paragraph_rules(parsed_paragraph, rule_set)
    compatible_rules = _compatible_paragraph_rules(
        matched_rules=matched_rules,
        classification_result=classification_result,
    )
    rules_by_property: dict[str, list[_MatchedParagraphRule]] = defaultdict(list)
    for rule in compatible_rules:
        rules_by_property[rule.target_property].append(rule)

    if not rules_by_property:
        rows.append(
            _report_row(
                object_id=classification_result.object_id,
                object_type_before=classification_result.object_type,
                object_type_after=classification_result.object_type,
                location=classification_result.location,
                text_preview=classification_result.original_text,
                property_name="classification",
                before="",
                after="",
                rule_id=classification_result.matched_rule_id or "",
                status="unresolved",
                reason=(
                    classification_result.reason
                    or "no normalization rule compatible with resolved classification winner"
                ),
            )
        )
        return rows

    for property_name, property_rules in sorted(rules_by_property.items()):
        sorted_rules = sorted(
            property_rules,
            key=lambda rule: (
                _MATCH_KIND_ORDER[rule.match_kind],
                rule.priority,
                rule.rule_id,
            ),
        )
        winner = sorted_rules[0]
        conflict_rules = [
            rule
            for rule in sorted_rules
            if rule.match_kind == winner.match_kind and rule.priority == winner.priority
        ]
        if len(conflict_rules) > 1:
            conflict_ids = ", ".join(rule.rule_id for rule in conflict_rules)
            rows.append(
                _report_row(
                    object_id=classification_result.object_id,
                    object_type_before=classification_result.object_type,
                    object_type_after=classification_result.object_type,
                    location=classification_result.location,
                    text_preview=classification_result.original_text,
                    property_name=property_name,
                    before="",
                    after="",
                    rule_id="",
                    status="unresolved",
                    reason=(
                        "conflicting rules at same priority: "
                        f"{conflict_ids} "
                        f"(match_kind={winner.match_kind}, priority={winner.priority})"
                    ),
                )
            )
            continue

        rows.append(
            _apply_single_paragraph_rule(
                paragraph=paragraph,
                classification_result=classification_result,
                rule=winner,
            )
        )

    return rows


def _apply_header_paragraph_rules(
    *,
    paragraph: Paragraph,
    classification_result,
    rule_set: RuleSet,
) -> list[dict[str, str]]:
    if classification_result.status == "unresolved":
        return [
            _report_row(
                object_id=classification_result.object_id,
                object_type_before=classification_result.object_type,
                object_type_after=classification_result.object_type,
                location=classification_result.location,
                text_preview=classification_result.original_text,
                property_name="classification",
                before="",
                after="",
                rule_id="",
                status="unresolved",
                reason=classification_result.reason or "no matching classification rule",
            )
        ]

    object_class = _special_object_target_object_type(
        rule_set.special_object_rules,
        classification_result.matched_rule_id,
    )
    if object_class is None:
        return [
            _report_row(
                object_id=classification_result.object_id,
                object_type_before=classification_result.object_type,
                object_type_after=classification_result.object_type,
                location=classification_result.location,
                text_preview=classification_result.original_text,
                property_name="classification",
                before="",
                after="",
                rule_id=classification_result.matched_rule_id or "",
                status="unresolved",
                reason="no target_object_type found for resolved header classification rule",
            )
        ]

    matched_rules = _matched_header_paragraph_rules(object_class, rule_set)
    if not matched_rules:
        return [
            _report_row(
                object_id=classification_result.object_id,
                object_type_before=classification_result.object_type,
                object_type_after=classification_result.object_type,
                location=classification_result.location,
                text_preview=classification_result.original_text,
                property_name="classification",
                before="",
                after="",
                rule_id=classification_result.matched_rule_id or "",
                status="unresolved",
                reason="no normalization rule compatible with resolved header classification winner",
            )
        ]

    rows: list[dict[str, str]] = []
    rules_by_property: dict[str, list[_MatchedParagraphRule]] = defaultdict(list)
    for rule in matched_rules:
        rules_by_property[rule.target_property].append(rule)

    for property_name, property_rules in sorted(rules_by_property.items()):
        sorted_rules = sorted(
            property_rules,
            key=lambda rule: (
                _MATCH_KIND_ORDER[rule.match_kind],
                rule.priority,
                rule.rule_id,
            ),
        )
        winner = sorted_rules[0]
        conflict_rules = [
            rule
            for rule in sorted_rules
            if rule.match_kind == winner.match_kind and rule.priority == winner.priority
        ]
        if len(conflict_rules) > 1:
            conflict_ids = ", ".join(rule.rule_id for rule in conflict_rules)
            rows.append(
                _report_row(
                    object_id=classification_result.object_id,
                    object_type_before=classification_result.object_type,
                    object_type_after=classification_result.object_type,
                    location=classification_result.location,
                    text_preview=classification_result.original_text,
                    property_name=property_name,
                    before="",
                    after="",
                    rule_id="",
                    status="unresolved",
                    reason=(
                        "conflicting rules at same priority: "
                        f"{conflict_ids} "
                        f"(match_kind={winner.match_kind}, priority={winner.priority})"
                    ),
                )
            )
            continue

        rows.append(
            _apply_single_paragraph_rule(
                paragraph=paragraph,
                classification_result=classification_result,
                rule=winner,
            )
        )

    return rows


def _apply_table_rules(
    *,
    table: Table,
    parsed_table: ParsedBodyTable,
    classification_result,
    rule_set: RuleSet,
) -> list[dict[str, str]]:
    if classification_result.status == "unresolved":
        return [
            _report_row(
                object_id=classification_result.object_id,
                object_type_before=classification_result.object_type,
                object_type_after=classification_result.object_type,
                location=classification_result.location,
                text_preview=classification_result.original_text,
                property_name="classification",
                before="",
                after="",
                rule_id="",
                status="unresolved",
                reason=classification_result.reason or "no matching classification rule",
            )
        ]

    matched_rules = _matched_table_rules(parsed_table, rule_set)
    compatible_rules = _compatible_table_rules(
        matched_rules=matched_rules,
        classification_result=classification_result,
    )
    if not compatible_rules:
        return [
            _report_row(
                object_id=classification_result.object_id,
                object_type_before=classification_result.object_type,
                object_type_after=classification_result.object_type,
                location=classification_result.location,
                text_preview=classification_result.original_text,
                property_name="classification",
                before="",
                after="",
                rule_id=classification_result.matched_rule_id or "",
                status="unresolved",
                reason="no normalization rule compatible with resolved table classification winner",
            )
        ]

    rows: list[dict[str, str]] = []
    rules_by_property: dict[str, list[_MatchedTableRule]] = defaultdict(list)
    for rule in compatible_rules:
        rules_by_property[rule.target_property].append(rule)

    for property_name, property_rules in sorted(rules_by_property.items()):
        sorted_rules = sorted(
            property_rules,
            key=lambda rule: (
                _MATCH_KIND_ORDER[rule.match_kind],
                rule.priority,
                rule.rule_id,
            ),
        )
        winner = sorted_rules[0]
        conflict_rules = [
            rule
            for rule in sorted_rules
            if rule.match_kind == winner.match_kind and rule.priority == winner.priority
        ]
        if len(conflict_rules) > 1:
            conflict_ids = ", ".join(rule.rule_id for rule in conflict_rules)
            rows.append(
                _report_row(
                    object_id=classification_result.object_id,
                    object_type_before=classification_result.object_type,
                    object_type_after=classification_result.object_type,
                    location=classification_result.location,
                    text_preview=classification_result.original_text,
                    property_name=property_name,
                    before="",
                    after="",
                    rule_id="",
                    status="unresolved",
                    reason=(
                        "conflicting rules at same priority: "
                        f"{conflict_ids} "
                        f"(match_kind={winner.match_kind}, priority={winner.priority})"
                    ),
                )
            )
            continue

        rows.append(
            _apply_single_table_rule(
                table=table,
                classification_result=classification_result,
                rule=winner,
            )
        )

    return rows


def _apply_single_paragraph_rule(
    *,
    paragraph: Paragraph,
    classification_result,
    rule: _MatchedParagraphRule,
) -> dict[str, str]:
    property_accessor = _paragraph_property_accessor(rule.target_property)
    if property_accessor is None:
        return _report_row(
            object_id=classification_result.object_id,
            object_type_before=classification_result.object_type,
            object_type_after=classification_result.object_type,
            location=classification_result.location,
            text_preview=classification_result.original_text,
            property_name=rule.target_property,
            before="",
            after="",
            rule_id=rule.rule_id,
            status="unresolved",
            reason=f"unsupported paragraph property: {rule.target_property}",
        )

    before_value = property_accessor.get(paragraph)
    try:
        target_value = property_accessor.parse(rule.target_value)
    except ValueError as exc:
        return _report_row(
            object_id=classification_result.object_id,
            object_type_before=classification_result.object_type,
            object_type_after=classification_result.object_type,
            location=classification_result.location,
            text_preview=classification_result.original_text,
            property_name=rule.target_property,
            before=property_accessor.format(before_value, rule.target_value),
            after=rule.target_value,
            rule_id=rule.rule_id,
            status="unresolved",
            reason=str(exc),
        )

    property_accessor.set(paragraph, target_value)
    after_value = property_accessor.get(paragraph)
    status: ReportStatus = "modified" if before_value != after_value else "unchanged"
    return _report_row(
        object_id=classification_result.object_id,
        object_type_before=classification_result.object_type,
        object_type_after=classification_result.object_type,
        location=classification_result.location,
        text_preview=classification_result.original_text,
        property_name=rule.target_property,
        before=property_accessor.format(before_value, rule.target_value),
        after=property_accessor.format(after_value, rule.target_value),
        rule_id=rule.rule_id,
        status=status,
        reason="",
    )


def _apply_single_table_rule(
    *,
    table: Table,
    classification_result,
    rule: _MatchedTableRule,
) -> dict[str, str]:
    property_accessor = _table_property_accessor(rule.target_property)
    if property_accessor is None:
        return _report_row(
            object_id=classification_result.object_id,
            object_type_before=classification_result.object_type,
            object_type_after=classification_result.object_type,
            location=classification_result.location,
            text_preview=classification_result.original_text,
            property_name=rule.target_property,
            before="",
            after="",
            rule_id=rule.rule_id,
            status="unresolved",
            reason=f"unsupported table property: {rule.target_property}",
        )

    before_value = property_accessor.get(table)
    try:
        target_value = property_accessor.parse(rule.target_value)
    except ValueError as exc:
        return _report_row(
            object_id=classification_result.object_id,
            object_type_before=classification_result.object_type,
            object_type_after=classification_result.object_type,
            location=classification_result.location,
            text_preview=classification_result.original_text,
            property_name=rule.target_property,
            before=property_accessor.format(before_value, rule.target_value),
            after=rule.target_value,
            rule_id=rule.rule_id,
            status="unresolved",
            reason=str(exc),
        )

    property_accessor.set(table, target_value)
    after_value = property_accessor.get(table)
    status: ReportStatus = "modified" if before_value != after_value else "unchanged"
    return _report_row(
        object_id=classification_result.object_id,
        object_type_before=classification_result.object_type,
        object_type_after=classification_result.object_type,
        location=classification_result.location,
        text_preview=classification_result.original_text,
        property_name=rule.target_property,
        before=property_accessor.format(before_value, rule.target_value),
        after=property_accessor.format(after_value, rule.target_value),
        rule_id=rule.rule_id,
        status=status,
        reason="",
    )


def _matched_paragraph_rules(
    paragraph: ParsedBodyParagraph,
    rule_set: RuleSet,
) -> list[_MatchedParagraphRule]:
    matched_rules: list[_MatchedParagraphRule] = []
    for rule in rule_set.paragraph_rules:
        candidate = _paragraph_rule_candidate(rule, paragraph)
        if candidate is None:
            continue
        matched_rules.append(
            _MatchedParagraphRule(
                rule_id=rule.rule_id,
                source_family="paragraph",
                priority=rule.priority,
                match_kind=candidate.match_kind,
                match_type=rule.match_type,
                match_value=rule.match_value,
                target_property=rule.target_property,
                target_value=rule.target_value,
            )
        )

    for rule in rule_set.numbering_rules:
        if not _numbering_rule_matches(rule, paragraph):
            continue
        matched_rules.append(
            _MatchedParagraphRule(
                rule_id=rule.rule_id,
                source_family="numbering",
                priority=rule.priority,
                match_kind="numbering",
                match_type=rule.match_type,
                match_value=rule.match_value,
                target_property=rule.target_property,
                target_value=rule.target_value,
            )
        )

    return matched_rules


def _matched_header_paragraph_rules(
    object_class: str,
    rule_set: RuleSet,
) -> list[_MatchedParagraphRule]:
    matched_rules: list[_MatchedParagraphRule] = []
    for rule in rule_set.paragraph_rules:
        if rule.match_type != "class" or rule.match_value != object_class:
            continue
        matched_rules.append(
            _MatchedParagraphRule(
                rule_id=rule.rule_id,
                source_family="header_class",
                priority=rule.priority,
                match_kind="structural",
                match_type=rule.match_type,
                match_value=rule.match_value,
                target_property=rule.target_property,
                target_value=rule.target_value,
            )
        )
    return matched_rules


def _matched_table_rules(
    table: ParsedBodyTable,
    rule_set: RuleSet,
) -> list[_MatchedTableRule]:
    flattened_text = "\n".join("\t".join(cell for cell in row) for row in table.rows)
    matched_rules: list[_MatchedTableRule] = []
    for rule in rule_set.table_rules:
        candidate = _table_rule_candidate(rule, flattened_text)
        if candidate is None:
            continue
        matched_rules.append(
            _MatchedTableRule(
                rule_id=rule.rule_id,
                priority=rule.priority,
                match_kind=candidate.match_kind,
                match_type=rule.match_type,
                match_value=rule.match_value,
                target_property=rule.target_property,
                target_value=rule.target_value,
            )
        )
    return matched_rules


def _compatible_paragraph_rules(
    *,
    matched_rules: list[_MatchedParagraphRule],
    classification_result,
) -> list[_MatchedParagraphRule]:
    if classification_result.status != "matched" or classification_result.matched_rule_id is None:
        return []

    winner_rule = next(
        (rule for rule in matched_rules if rule.rule_id == classification_result.matched_rule_id),
        None,
    )
    if winner_rule is None:
        return []

    return [
        rule
        for rule in matched_rules
        if rule.source_family == winner_rule.source_family
        and rule.match_kind == winner_rule.match_kind
        and rule.match_type == winner_rule.match_type
        and rule.match_value == winner_rule.match_value
    ]


def _compatible_table_rules(
    *,
    matched_rules: list[_MatchedTableRule],
    classification_result,
) -> list[_MatchedTableRule]:
    if classification_result.status != "matched" or classification_result.matched_rule_id is None:
        return []

    winner_rule = next(
        (rule for rule in matched_rules if rule.rule_id == classification_result.matched_rule_id),
        None,
    )
    if winner_rule is None:
        return []

    return [
        rule
        for rule in matched_rules
        if rule.match_kind == winner_rule.match_kind
        and rule.match_type == winner_rule.match_type
        and rule.match_value == winner_rule.match_value
    ]


def _body_paragraph_map(document: DocumentObject) -> dict[int, Paragraph]:
    body_paragraphs: dict[int, Paragraph] = {}
    body_index = 0
    for element in _iter_body_elements(document):
        if isinstance(element, Paragraph):
            if not element.text.strip():
                continue
            body_paragraphs[body_index] = element
            body_index += 1
            continue
        if isinstance(element, Table):
            body_index += 1
    return body_paragraphs


def _body_table_map(document: DocumentObject) -> dict[int, Table]:
    body_tables: dict[int, Table] = {}
    body_index = 0
    for element in _iter_body_elements(document):
        if isinstance(element, Paragraph):
            if not element.text.strip():
                continue
            body_index += 1
            continue
        body_tables[body_index] = element
        body_index += 1
    return body_tables


def _header_paragraph_map(document: DocumentObject) -> dict[str, Paragraph]:
    mapped: dict[str, Paragraph] = {}
    seen_parts: set[int] = set()
    header_index = 0
    for section in document.sections:
        for header in (section.header, section.first_page_header, section.even_page_header):
            part_key = id(header.part)
            if part_key in seen_parts:
                continue
            seen_parts.add(part_key)
            header_items: list[Paragraph] = []
            item_index = 0
            for element in header.iter_inner_content():
                if isinstance(element, Paragraph):
                    if not element.text.strip():
                        continue
                    header_items.append(element)
                    item_index += 1
                    continue
                item_index += 1
            if not header_items:
                continue
            for mapped_item_index, paragraph in enumerate(header_items):
                mapped[f"headers[{header_index}].items[{mapped_item_index}]"] = paragraph
            header_index += 1
    return mapped


def _header_location_indexes(location: str) -> tuple[int, int]:
    match = re.fullmatch(r"headers\[(\d+)\]\.items\[(\d+)\]", location)
    if match is None:
        raise ValueError(f"unsupported header location format: {location}")
    return int(match.group(1)), int(match.group(2))


def _write_annotated_document(
    *,
    normalized_path: Path,
    annotated_path: Path,
    report_rows: list[dict[str, str]],
) -> None:
    document = Document(normalized_path)
    body_paragraphs = {
        f"body_items[{index}]": paragraph
        for index, paragraph in _body_paragraph_map(document).items()
    }
    body_tables = {
        f"body_items[{index}]": table
        for index, table in _body_table_map(document).items()
    }
    header_paragraphs = _header_paragraph_map(document)

    modified_rows = [row for row in report_rows if row["status"] == "modified"]
    rows_by_location: dict[str, list[dict[str, str]]] = defaultdict(list)
    section_rows: list[dict[str, str]] = []
    for row in modified_rows:
        if row["location"].startswith("sections["):
            section_rows.append(row)
            continue
        rows_by_location[row["location"]].append(row)

    if section_rows:
        _insert_document_annotation(
            document,
            _build_section_annotation(section_rows),
        )

    for location, rows in rows_by_location.items():
        textual_rows = [row for row in rows if row["property"] in _TEXTUAL_PROPERTIES]
        annotation_rows = [row for row in rows if row["property"] in _ANNOTATION_PROPERTIES]

        if location in body_paragraphs:
            paragraph = body_paragraphs[location]
            if textual_rows:
                _annotate_paragraph_textual_changes(paragraph, textual_rows)
            if annotation_rows:
                _insert_note_after_paragraph(paragraph, _build_annotation_text(annotation_rows))
            continue

        if location in header_paragraphs:
            paragraph = header_paragraphs[location]
            if textual_rows:
                _annotate_paragraph_textual_changes(paragraph, textual_rows)
            if annotation_rows:
                _insert_note_after_paragraph(paragraph, _build_annotation_text(annotation_rows))
            continue

        if location in body_tables:
            table = body_tables[location]
            if textual_rows:
                _annotate_table_textual_changes(table, textual_rows)
            if annotation_rows:
                _insert_note_after_table(table, _build_annotation_text(annotation_rows))

    document.save(annotated_path)


def _mark_paragraph_red(paragraph: Paragraph) -> None:
    for run in paragraph.runs:
        run.font.color.rgb = _RED


def _mark_table_red(table: Table) -> None:
    for paragraph in _iter_table_paragraphs(table):
        _mark_paragraph_red(paragraph)


def _annotate_paragraph_textual_changes(
    paragraph: Paragraph,
    rows: list[dict[str, str]],
) -> None:
    inline_rows = [row for row in rows if _is_inline_textual_property(row["property"])]
    general_rows = [row for row in rows if not _is_inline_textual_property(row["property"])]
    if inline_rows:
        _mark_inline_textual_segments_red(paragraph, inline_rows)
    if general_rows:
        _mark_paragraph_red(paragraph)
    _append_annotation_run(paragraph, _build_textual_annotation_text(rows))


def _annotate_table_textual_changes(
    table: Table,
    rows: list[dict[str, str]],
) -> None:
    rows_by_paragraph: dict[int, tuple[Paragraph, list[dict[str, str]]]] = {}
    for row in rows:
        for paragraph in _table_paragraphs_for_property(table, row["property"]):
            if not paragraph.text.strip():
                continue
            record = rows_by_paragraph.setdefault(id(paragraph), (paragraph, []))
            record[1].append(row)

    for paragraph, paragraph_rows in rows_by_paragraph.values():
        _annotate_paragraph_textual_changes(paragraph, paragraph_rows)


def _insert_note_after_paragraph(paragraph: Paragraph, text: str) -> Paragraph:
    note = paragraph._parent.add_paragraph()
    _style_annotation_paragraph(note, text)
    paragraph._p.addnext(note._p)
    return note


def _insert_note_after_table(table: Table, text: str) -> Paragraph:
    note = table._parent.add_paragraph()
    _style_annotation_paragraph(note, text)
    table._tbl.addnext(note._p)
    return note


def _insert_document_annotation(document: DocumentObject, text: str) -> Paragraph:
    note = document.add_paragraph()
    _style_annotation_paragraph(note, text)
    first_paragraph = next((paragraph for paragraph in document.paragraphs), None)
    if first_paragraph is not None:
        first_paragraph._p.addprevious(note._p)
    return note


def _style_annotation_paragraph(paragraph: Paragraph, text: str) -> None:
    paragraph.paragraph_format.first_line_indent = Cm(0)
    paragraph.paragraph_format.space_before = Pt(0)
    paragraph.paragraph_format.space_after = Pt(0)
    run = paragraph.add_run(text)
    run.font.color.rgb = _RED
    run.font.name = _ANNOTATION_FONT_NAME
    run.font.size = _ANNOTATION_FONT_SIZE
    run._element.get_or_add_rPr().rFonts.set(qn("w:ascii"), _ANNOTATION_FONT_NAME)
    run._element.get_or_add_rPr().rFonts.set(qn("w:hAnsi"), _ANNOTATION_FONT_NAME)
    run._element.get_or_add_rPr().rFonts.set(qn("w:eastAsia"), _ANNOTATION_FONT_NAME)
    run._element.get_or_add_rPr().rFonts.set(qn("w:cs"), _ANNOTATION_FONT_NAME)


def _append_annotation_run(paragraph: Paragraph, text: str) -> None:
    run = paragraph.add_run(text)
    run.font.color.rgb = _RED
    run.font.name = _ANNOTATION_FONT_NAME
    run.font.size = _ANNOTATION_FONT_SIZE
    run._element.get_or_add_rPr().rFonts.set(qn("w:ascii"), _ANNOTATION_FONT_NAME)
    run._element.get_or_add_rPr().rFonts.set(qn("w:hAnsi"), _ANNOTATION_FONT_NAME)
    run._element.get_or_add_rPr().rFonts.set(qn("w:eastAsia"), _ANNOTATION_FONT_NAME)
    run._element.get_or_add_rPr().rFonts.set(qn("w:cs"), _ANNOTATION_FONT_NAME)


def _build_textual_annotation_text(rows: list[dict[str, str]]) -> str:
    ordered_rows = sorted(
        rows,
        key=lambda row: (
            0 if row["property"] == "font_name" else 1,
            row["rule_id"],
        ),
    )
    fragments = [_describe_textual_change(row) for row in ordered_rows]
    return f"【规范化说明：{'；'.join(fragments)}】"


def _build_textual_annotation_text(rows: list[dict[str, str]]) -> str:
    ordered_rows = sorted(
        rows,
        key=lambda row: (
            _textual_property_order(row["property"]),
            row["rule_id"],
        ),
    )
    fragments = [_describe_textual_change(row) for row in ordered_rows]
    return "【规范化说明：" + "；".join(fragments) + "】"


def _is_inline_textual_property(property_name: str) -> bool:
    return property_name.startswith("label_") or property_name.startswith("content_")


def _mark_inline_textual_segments_red(paragraph: Paragraph, rows: list[dict[str, str]]) -> None:
    label_run, content_run = _ensure_inline_segment_runs(paragraph)
    for row in rows:
        if row["property"].startswith("label_"):
            label_run.font.color.rgb = _RED
            continue
        if row["property"].startswith("content_") and content_run is not None:
            content_run.font.color.rgb = _RED


def _textual_property_order(property_name: str) -> int:
    order = [
        "header_row_font_name",
        "header_row_font_size",
        "body_rows_font_name",
        "body_rows_font_size",
        "label_font_name",
        "label_font_size",
        "label_bold",
        "content_font_name",
        "content_font_size",
        "content_bold",
        "font_name",
        "font_size",
    ]
    if property_name in order:
        return order.index(property_name)
    return len(order)


def _build_annotation_text(rows: list[dict[str, str]]) -> str:
    ordered_rows = sorted(
        rows,
        key=lambda row: (
            _annotation_property_order(row["property"]),
            row["rule_id"],
        ),
    )
    fragments = [
        f"{_PROPERTY_LABELS.get(row['property'], row['property'])}原为 {_prettify_annotation_value(row['before'])}，规范为 {_prettify_annotation_value(row['after'])}"
        for row in ordered_rows
    ]
    return f"[规范化批注] {'；'.join(fragments)}"


def _build_section_annotation(rows: list[dict[str, str]]) -> str:
    deduped: dict[str, dict[str, str]] = {}
    for row in rows:
        deduped.setdefault(row["property"], row)
    ordered_rows = sorted(
        deduped.values(),
        key=lambda row: _annotation_property_order(row["property"]),
    )
    fragments = [
        f"{_PROPERTY_LABELS.get(row['property'], row['property'])}已按规范设为 {_prettify_annotation_value(row['after'])}"
        for row in ordered_rows
    ]
    return f"[规范化批注] 页面设置：{'；'.join(fragments)}"


def _describe_textual_change(row: dict[str, str]) -> str:
    property_name = row["property"]
    before = row["before"]
    after = row["after"]
    if property_name == "font_name":
        return f"字体原为 {_prettify_font_name_value(before)}，调整为 {_prettify_font_name_value(after)}"
    if property_name == "font_size":
        return f"字号原为 {_prettify_font_size_value(before)}，调整为 {_prettify_font_size_value(after)}"
    return f"{_PROPERTY_LABELS.get(property_name, property_name)}原为 {_prettify_annotation_value(before)}，现调整为 {_prettify_annotation_value(after)}"


def _describe_textual_change(row: dict[str, str]) -> str:
    property_name = row["property"]
    before = row["before"]
    after = row["after"]
    if property_name == "font_name":
        return f"字体原为 {_prettify_font_name_value(before)}，调整为 {_prettify_font_name_value(after)}"
    if property_name == "font_size":
        return f"字号原为 {_prettify_font_size_value(before)}，调整为 {_prettify_font_size_value(after)}"
    if property_name == "header_row_font_name":
        return f"表头字体原为 {_prettify_font_name_value(before)}，调整为 {_prettify_font_name_value(after)}"
    if property_name == "header_row_font_size":
        return f"表头字号原为 {_prettify_font_size_value(before)}，调整为 {_prettify_font_size_value(after)}"
    if property_name == "body_rows_font_name":
        return f"数据行字体原为 {_prettify_font_name_value(before)}，调整为 {_prettify_font_name_value(after)}"
    if property_name == "body_rows_font_size":
        return f"数据行字号原为 {_prettify_font_size_value(before)}，调整为 {_prettify_font_size_value(after)}"
    if property_name == "label_font_name":
        return f"标签字体原为 {_prettify_font_name_value(before)}，调整为 {_prettify_font_name_value(after)}"
    if property_name == "label_font_size":
        return f"标签字号原为 {_prettify_font_size_value(before)}，调整为 {_prettify_font_size_value(after)}"
    if property_name == "content_font_name":
        return f"内容字体原为 {_prettify_font_name_value(before)}，调整为 {_prettify_font_name_value(after)}"
    if property_name == "content_font_size":
        return f"内容字号原为 {_prettify_font_size_value(before)}，调整为 {_prettify_font_size_value(after)}"
    if property_name == "label_bold":
        return f"标签加粗原为 {_format_bool_display(before)}，调整为 {_format_bool_display(after)}"
    if property_name == "content_bold":
        return f"内容加粗原为 {_format_bool_display(before)}，调整为 {_format_bool_display(after)}"
    return f"{_PROPERTY_LABELS.get(property_name, property_name)}原为 {_prettify_annotation_value(before)}，调整为 {_prettify_annotation_value(after)}"


def _annotation_property_order(property_name: str) -> int:
    order = [
        "first_line_indent",
        "hanging_indent",
        "space_before",
        "space_after",
        "line_spacing",
        "page_margin_top",
        "page_margin_bottom",
        "page_margin_left",
        "page_margin_right",
    ]
    if property_name in order:
        return order.index(property_name)
    return len(order)


def _prettify_annotation_value(value: str) -> str:
    if value == "":
        return "未显式设置（继承样式）"
    cm_match = re.fullmatch(r"(-?\d+(?:\.\d+)?)cm", value)
    if cm_match is not None:
        amount = float(cm_match.group(1))
        return f"{amount:.2f}".rstrip("0").rstrip(".") + "cm"

    pt_match = re.fullmatch(r"(-?\d+(?:\.\d+)?)pt", value)
    if pt_match is not None:
        amount = float(pt_match.group(1))
        return f"{amount:g}pt"

    return value


def _prettify_font_size_value(value: str) -> str:
    if value.startswith("mixed[") and value.endswith("]"):
        items = value[len("mixed[") : -1].split("|")
        return "混合[" + "|".join(_prettify_font_size_value(item) for item in items) + "]"

    pt_match = re.fullmatch(r"(-?\d+(?:\.\d+)?)pt", value)
    if pt_match is None:
        return value

    amount = float(pt_match.group(1))
    size_name = _font_size_name(amount)
    if size_name is None:
        return f"{amount:g}pt"
    return f"{size_name}（{amount:g}pt）"


def _prettify_font_name_value(value: str) -> str:
    normalized = value.strip()
    if not normalized:
        return "未显式设置（继承样式）"
    if "|" in normalized:
        return "、".join(_prettify_font_name_value(part) for part in normalized.split("|"))
    return _FONT_DISPLAY_NAMES.get(normalized, normalized)


def _parse_bool(value: str) -> bool:
    normalized = value.strip().lower()
    if normalized in {"true", "1", "yes", "是"}:
        return True
    if normalized in {"false", "0", "no", "否"}:
        return False
    raise ValueError(f"unsupported boolean value: {value}")


def _format_bool_display(value) -> str:
    if isinstance(value, str):
        normalized = value.strip().lower()
        if normalized in {"true", "1", "yes", "是"}:
            return "是"
        if normalized in {"false", "0", "no", "否", ""}:
            return "否"
    return "是" if bool(value) else "否"


def _font_size_name(value_pt: float) -> str | None:
    named_sizes = (
        (26.0, "一号"),
        (24.0, "小一"),
        (22.0, "二号"),
        (18.0, "小二"),
        (16.0, "三号"),
        (15.0, "小三"),
        (14.0, "四号"),
        (12.0, "小四"),
        (10.5, "五号"),
        (9.0, "小五"),
        (7.5, "六号"),
        (6.5, "小六"),
        (5.5, "七号"),
        (5.0, "八号"),
    )
    for amount, label in named_sizes:
        if abs(value_pt - amount) < 0.05:
            return label
    return None


@dataclass(frozen=True)
class _SectionPropertyAccessor:
    get: callable
    set: callable
    format: callable


def _document_property_accessor(property_name: str) -> _SectionPropertyAccessor | None:
    accessors: dict[str, _SectionPropertyAccessor] = {
        "page_margin_top": _SectionPropertyAccessor(
            get=lambda section: section.top_margin,
            set=lambda section, value: setattr(section, "top_margin", value),
            format=lambda value, raw: _format_length(value, raw),
        ),
        "page_margin_bottom": _SectionPropertyAccessor(
            get=lambda section: section.bottom_margin,
            set=lambda section, value: setattr(section, "bottom_margin", value),
            format=lambda value, raw: _format_length(value, raw),
        ),
        "page_margin_left": _SectionPropertyAccessor(
            get=lambda section: section.left_margin,
            set=lambda section, value: setattr(section, "left_margin", value),
            format=lambda value, raw: _format_length(value, raw),
        ),
        "page_margin_right": _SectionPropertyAccessor(
            get=lambda section: section.right_margin,
            set=lambda section, value: setattr(section, "right_margin", value),
            format=lambda value, raw: _format_length(value, raw),
        ),
    }
    return accessors.get(property_name)


@dataclass(frozen=True)
class _ParagraphPropertyAccessor:
    get: callable
    set: callable
    parse: callable
    format: callable


@dataclass(frozen=True)
class _TablePropertyAccessor:
    get: callable
    set: callable
    parse: callable
    format: callable


def _paragraph_property_accessor(property_name: str) -> _ParagraphPropertyAccessor | None:
    accessors: dict[str, _ParagraphPropertyAccessor] = {
        "font_name": _ParagraphPropertyAccessor(
            get=_get_paragraph_font_name,
            set=_set_paragraph_font_name,
            parse=lambda value: value,
            format=lambda value, _: value,
        ),
        "font_size": _ParagraphPropertyAccessor(
            get=lambda paragraph: _get_run_length_state(paragraph, "size"),
            set=lambda paragraph, value: _set_run_length(paragraph, "size", value),
            parse=_parse_measurement,
            format=_format_run_length_state,
        ),
        "label_font_name": _ParagraphPropertyAccessor(
            get=_get_label_font_name,
            set=_set_label_font_name,
            parse=lambda value: value,
            format=lambda value, _: value,
        ),
        "content_font_name": _ParagraphPropertyAccessor(
            get=_get_content_font_name,
            set=_set_content_font_name,
            parse=lambda value: value,
            format=lambda value, _: value,
        ),
        "label_font_size": _ParagraphPropertyAccessor(
            get=_get_label_font_size,
            set=_set_label_font_size,
            parse=_parse_measurement,
            format=_format_length,
        ),
        "content_font_size": _ParagraphPropertyAccessor(
            get=_get_content_font_size,
            set=_set_content_font_size,
            parse=_parse_measurement,
            format=_format_length,
        ),
        "label_bold": _ParagraphPropertyAccessor(
            get=_get_label_bold,
            set=_set_label_bold,
            parse=_parse_bool,
            format=lambda value, _: _format_bool_display(value),
        ),
        "content_bold": _ParagraphPropertyAccessor(
            get=_get_content_bold,
            set=_set_content_bold,
            parse=_parse_bool,
            format=lambda value, _: _format_bool_display(value),
        ),
        "first_line_indent": _ParagraphPropertyAccessor(
            get=lambda paragraph: paragraph.paragraph_format.first_line_indent,
            set=lambda paragraph, value: setattr(paragraph.paragraph_format, "first_line_indent", value),
            parse=_parse_measurement,
            format=_format_length,
        ),
        "hanging_indent": _ParagraphPropertyAccessor(
            get=_get_paragraph_hanging_indent,
            set=_set_paragraph_hanging_indent,
            parse=_parse_measurement,
            format=_format_length,
        ),
        "space_before": _ParagraphPropertyAccessor(
            get=lambda paragraph: paragraph.paragraph_format.space_before,
            set=lambda paragraph, value: setattr(paragraph.paragraph_format, "space_before", value),
            parse=_parse_measurement,
            format=_format_length,
        ),
        "space_after": _ParagraphPropertyAccessor(
            get=lambda paragraph: paragraph.paragraph_format.space_after,
            set=lambda paragraph, value: setattr(paragraph.paragraph_format, "space_after", value),
            parse=_parse_measurement,
            format=_format_length,
        ),
        "line_spacing": _ParagraphPropertyAccessor(
            get=lambda paragraph: paragraph.paragraph_format.line_spacing,
            set=lambda paragraph, value: setattr(paragraph.paragraph_format, "line_spacing", value),
            parse=_parse_line_spacing,
            format=_format_line_spacing,
        ),
    }
    return accessors.get(property_name)


def _table_property_accessor(property_name: str) -> _TablePropertyAccessor | None:
    accessors: dict[str, _TablePropertyAccessor] = {
        "font_name": _TablePropertyAccessor(
            get=_get_table_font_name,
            set=_set_table_font_name,
            parse=lambda value: value,
            format=lambda value, _: value,
        ),
        "font_size": _TablePropertyAccessor(
            get=lambda table: _get_table_run_length_state(table, "size"),
            set=lambda table, value: _set_table_run_length(table, "size", value),
            parse=_parse_measurement,
            format=_format_run_length_state,
        ),
        "header_row_font_name": _TablePropertyAccessor(
            get=lambda table: _get_table_font_name_in_rows(table, row_selector="header"),
            set=lambda table, value: _set_table_font_name_in_rows(table, value, row_selector="header"),
            parse=lambda value: value,
            format=lambda value, _: value,
        ),
        "header_row_font_size": _TablePropertyAccessor(
            get=lambda table: _get_table_run_length_state_in_rows(table, "size", row_selector="header"),
            set=lambda table, value: _set_table_run_length_in_rows(table, "size", value, row_selector="header"),
            parse=_parse_measurement,
            format=_format_run_length_state,
        ),
        "body_rows_font_name": _TablePropertyAccessor(
            get=lambda table: _get_table_font_name_in_rows(table, row_selector="body"),
            set=lambda table, value: _set_table_font_name_in_rows(table, value, row_selector="body"),
            parse=lambda value: value,
            format=lambda value, _: value,
        ),
        "body_rows_font_size": _TablePropertyAccessor(
            get=lambda table: _get_table_run_length_state_in_rows(table, "size", row_selector="body"),
            set=lambda table, value: _set_table_run_length_in_rows(table, "size", value, row_selector="body"),
            parse=_parse_measurement,
            format=_format_run_length_state,
        ),
    }
    return accessors.get(property_name)


def _get_paragraph_font_name(paragraph: Paragraph) -> str:
    fonts: list[str] = []
    seen_fonts: set[str] = set()
    for run in paragraph.runs:
        for font_name in _used_run_fonts(run):
            if font_name in seen_fonts:
                continue
            seen_fonts.add(font_name)
            fonts.append(font_name)
    return "|".join(fonts)


def _get_paragraph_hanging_indent(paragraph: Paragraph) -> Length:
    first_line_indent = paragraph.paragraph_format.first_line_indent
    if first_line_indent is None or int(first_line_indent) >= 0:
        return Cm(0)
    return Length(abs(int(first_line_indent)))


def _set_paragraph_hanging_indent(paragraph: Paragraph, value: Length) -> None:
    paragraph.paragraph_format.left_indent = value
    paragraph.paragraph_format.first_line_indent = Length(-int(value))


def _get_label_font_name(paragraph: Paragraph) -> str:
    segmented_runs = _inline_segment_runs_if_present(paragraph)
    if segmented_runs is not None:
        label_run, _ = segmented_runs
    else:
        label_run = next((run for run in paragraph.runs if run.text), None)
        if label_run is None:
            return ""
    return "|".join(_used_run_fonts(label_run))


def _set_label_font_name(paragraph: Paragraph, font_name: str) -> None:
    label_run, _ = _ensure_inline_segment_runs(paragraph)
    _set_run_font_name(label_run, font_name)


def _get_content_font_name(paragraph: Paragraph) -> str:
    segmented_runs = _inline_segment_runs_if_present(paragraph)
    if segmented_runs is not None:
        _, content_run = segmented_runs
    else:
        content_run = next((run for run in paragraph.runs if run.text), None)
    if content_run is None:
        return ""
    return "|".join(_used_run_fonts(content_run))


def _get_label_font_size(paragraph: Paragraph):
    segmented_runs = _inline_segment_runs_if_present(paragraph)
    if segmented_runs is not None:
        label_run, _ = segmented_runs
    else:
        label_run = next((run for run in paragraph.runs if run.text), None)
        if label_run is None:
            return None
    return label_run.font.size


def _set_label_font_size(paragraph: Paragraph, value: Length) -> None:
    label_run, _ = _ensure_inline_segment_runs(paragraph)
    label_run.font.size = value


def _get_content_font_size(paragraph: Paragraph):
    segmented_runs = _inline_segment_runs_if_present(paragraph)
    if segmented_runs is not None:
        _, content_run = segmented_runs
    else:
        content_run = next((run for run in paragraph.runs if run.text), None)
    if content_run is None:
        return None
    return content_run.font.size


def _set_content_font_size(paragraph: Paragraph, value: Length) -> None:
    _, content_run = _ensure_inline_segment_runs(paragraph)
    if content_run is None:
        return
    content_run.font.size = value


def _set_content_font_name(paragraph: Paragraph, font_name: str) -> None:
    _, content_run = _ensure_inline_segment_runs(paragraph)
    if content_run is None:
        return
    _set_run_font_name(content_run, font_name)


def _get_label_bold(paragraph: Paragraph) -> bool:
    segmented_runs = _inline_segment_runs_if_present(paragraph)
    if segmented_runs is not None:
        label_run, _ = segmented_runs
    else:
        label_run = next((run for run in paragraph.runs if run.text), None)
        if label_run is None:
            return False
    return bool(label_run.font.bold)


def _set_label_bold(paragraph: Paragraph, value: bool) -> None:
    label_run, _ = _ensure_inline_segment_runs(paragraph)
    label_run.font.bold = value


def _get_content_bold(paragraph: Paragraph) -> bool:
    segmented_runs = _inline_segment_runs_if_present(paragraph)
    if segmented_runs is not None:
        _, content_run = segmented_runs
    else:
        content_run = next((run for run in paragraph.runs if run.text), None)
    if content_run is None:
        return False
    return bool(content_run.font.bold)


def _set_content_bold(paragraph: Paragraph, value: bool) -> None:
    _, content_run = _ensure_inline_segment_runs(paragraph)
    if content_run is None:
        return
    content_run.font.bold = value


def _set_paragraph_font_name(paragraph: Paragraph, font_name: str) -> None:
    if font_name in _EAST_ASIAN_FONT_NAMES:
        for run in paragraph.runs:
            _set_run_script_aware_font_name(
                run,
                east_asian_font_name=font_name,
                western_font_name=_WESTERN_FONT_NAME,
            )
        return

    for run in paragraph.runs:
        _set_run_uniform_font_name(run, font_name)


def _get_table_font_name(table: Table) -> str:
    fonts: list[str] = []
    seen_fonts: set[str] = set()
    for paragraph in _iter_table_paragraphs(table):
        for run in paragraph.runs:
            for font_name in _used_run_fonts(run):
                if font_name in seen_fonts:
                    continue
                seen_fonts.add(font_name)
                fonts.append(font_name)
    return "|".join(fonts)


def _set_table_font_name(table: Table, font_name: str) -> None:
    for paragraph in _iter_table_paragraphs(table):
        _set_paragraph_font_name(paragraph, font_name)


def _get_table_font_name_in_rows(table: Table, *, row_selector: str) -> str:
    fonts: list[str] = []
    seen_fonts: set[str] = set()
    for paragraph in _iter_table_paragraphs_in_rows(table, row_selector=row_selector):
        for run in paragraph.runs:
            for font_name in _used_run_fonts(run):
                if font_name in seen_fonts:
                    continue
                seen_fonts.add(font_name)
                fonts.append(font_name)
    return "|".join(fonts)


def _set_table_font_name_in_rows(table: Table, font_name: str, *, row_selector: str) -> None:
    for paragraph in _iter_table_paragraphs_in_rows(table, row_selector=row_selector):
        _set_paragraph_font_name(paragraph, font_name)


def _ensure_inline_segment_runs(paragraph: Paragraph) -> tuple:
    inline_parts = _split_inline_paragraph_text(paragraph.text)
    if inline_parts is None:
        raise ValueError(f"paragraph does not match a supported inline-prefix rule: {paragraph.text!r}")

    label_text, content_text = inline_parts
    if len(paragraph.runs) == 2 and paragraph.runs[0].text == label_text and paragraph.runs[1].text == content_text:
        return paragraph.runs[0], paragraph.runs[1]
    if len(paragraph.runs) == 1 and paragraph.runs[0].text == label_text and content_text == "":
        return paragraph.runs[0], None

    _rebuild_inline_segment_runs(paragraph, label_text, content_text)
    if content_text == "":
        return paragraph.runs[0], None
    return paragraph.runs[0], paragraph.runs[1]


def _inline_segment_runs_if_present(paragraph: Paragraph) -> tuple | None:
    inline_parts = _split_inline_paragraph_text(paragraph.text)
    if inline_parts is None:
        return None
    label_text, content_text = inline_parts
    if len(paragraph.runs) == 2 and paragraph.runs[0].text == label_text and paragraph.runs[1].text == content_text:
        return paragraph.runs[0], paragraph.runs[1]
    if len(paragraph.runs) == 1 and paragraph.runs[0].text == label_text and content_text == "":
        return paragraph.runs[0], None
    return None


def _split_inline_paragraph_text(text: str) -> tuple[str, str] | None:
    for prefix in _INLINE_PREFIXES:
        if text.startswith(prefix):
            return prefix, text[len(prefix) :]
    return None


def _rebuild_inline_segment_runs(paragraph: Paragraph, label_text: str, content_text: str) -> None:
    source_run = next((run for run in paragraph.runs if run.text), None)
    source_rpr = None
    if source_run is not None and source_run._element.rPr is not None:
        source_rpr = deepcopy(source_run._element.rPr)

    for child in list(paragraph._p):
        if child.tag == qn("w:pPr"):
            continue
        paragraph._p.remove(child)

    if label_text:
        label_run = paragraph.add_run(label_text)
        if source_rpr is not None:
            label_run._element.insert(0, deepcopy(source_rpr))
    if content_text:
        content_run = paragraph.add_run(content_text)
        if source_rpr is not None:
            content_run._element.insert(0, deepcopy(source_rpr))


def _used_run_fonts(run) -> tuple[str, ...]:
    fonts: list[str] = []
    seen_fonts: set[str] = set()
    for bucket in _ordered_script_buckets(run.text):
        font_name = _run_font_for_bucket(run, bucket)
        if font_name in seen_fonts:
            continue
        seen_fonts.add(font_name)
        fonts.append(font_name)
    return tuple(fonts)


def _ordered_script_buckets(text: str) -> tuple[str, ...]:
    buckets: list[str] = []
    previous_bucket: str | None = None
    for character in text:
        bucket = _script_bucket(character)
        if bucket == previous_bucket:
            continue
        buckets.append(bucket)
        previous_bucket = bucket
    return tuple(buckets)


def _script_bucket(character: str) -> str:
    if _is_east_asian_character(character):
        return "east_asia"
    return "western"


def _is_east_asian_character(character: str) -> bool:
    code_point = ord(character)
    east_asian_ranges = (
        (0x2E80, 0x2EFF),
        (0x2F00, 0x2FDF),
        (0x3000, 0x303F),
        (0x3040, 0x30FF),
        (0x3100, 0x312F),
        (0x31A0, 0x31BF),
        (0x3400, 0x4DBF),
        (0x4E00, 0x9FFF),
        (0xF900, 0xFAFF),
        (0xFE30, 0xFE4F),
        (0xFF00, 0xFFEF),
    )
    return any(start <= code_point <= end for start, end in east_asian_ranges)


def _run_font_for_bucket(run, bucket: str) -> str:
    r_pr = run._element.rPr
    r_fonts = None if r_pr is None else r_pr.rFonts
    if bucket == "east_asia":
        if r_fonts is not None:
            east_asian = r_fonts.get(qn("w:eastAsia"))
            if east_asian:
                return east_asian
        return run.font.name or ""

    if r_fonts is not None:
        for slot in ("w:ascii", "w:hAnsi", "w:cs"):
            font_name = r_fonts.get(qn(slot))
            if font_name:
                return font_name
    return run.font.name or ""


def _set_run_uniform_font_name(run, font_name: str) -> None:
    run.font.name = font_name
    r_fonts = run._element.get_or_add_rPr().get_or_add_rFonts()
    r_fonts.set(qn("w:ascii"), font_name)
    r_fonts.set(qn("w:hAnsi"), font_name)
    r_fonts.set(qn("w:eastAsia"), font_name)
    r_fonts.set(qn("w:cs"), font_name)


def _set_run_script_aware_font_name(
    run,
    *,
    east_asian_font_name: str,
    western_font_name: str,
) -> None:
    r_fonts = run._element.get_or_add_rPr().get_or_add_rFonts()
    r_fonts.set(qn("w:ascii"), western_font_name)
    r_fonts.set(qn("w:hAnsi"), western_font_name)
    r_fonts.set(qn("w:eastAsia"), east_asian_font_name)
    r_fonts.set(qn("w:cs"), western_font_name)


def _set_run_font_name(run, font_name: str) -> None:
    if font_name in _EAST_ASIAN_FONT_NAMES:
        _set_run_script_aware_font_name(
            run,
            east_asian_font_name=font_name,
            western_font_name=_WESTERN_FONT_NAME,
        )
        return
    _set_run_uniform_font_name(run, font_name)


def _get_run_length_state(paragraph: Paragraph, property_name: str) -> _RunLengthState:
    return _RunLengthState(
        values=tuple(getattr(run.font, property_name) for run in paragraph.runs)
    )


def _set_run_length(paragraph: Paragraph, property_name: str, value: Length) -> None:
    for run in paragraph.runs:
        setattr(run.font, property_name, value)


def _get_table_run_length_state(table: Table, property_name: str) -> _RunLengthState:
    values: list[Length | None] = []
    for paragraph in _iter_table_paragraphs(table):
        values.extend(getattr(run.font, property_name) for run in paragraph.runs)
    return _RunLengthState(values=tuple(values))


def _set_table_run_length(table: Table, property_name: str, value: Length) -> None:
    for paragraph in _iter_table_paragraphs(table):
        _set_run_length(paragraph, property_name, value)


def _get_table_run_length_state_in_rows(
    table: Table,
    property_name: str,
    *,
    row_selector: str,
) -> _RunLengthState:
    values: list[Length | None] = []
    for paragraph in _iter_table_paragraphs_in_rows(table, row_selector=row_selector):
        values.extend(getattr(run.font, property_name) for run in paragraph.runs)
    return _RunLengthState(values=tuple(values))


def _set_table_run_length_in_rows(
    table: Table,
    property_name: str,
    value: Length,
    *,
    row_selector: str,
) -> None:
    for paragraph in _iter_table_paragraphs_in_rows(table, row_selector=row_selector):
        _set_run_length(paragraph, property_name, value)


def _parse_line_spacing(value: str) -> float | Length:
    stripped = value.strip()
    if re.fullmatch(r"-?\d+(?:\.\d+)?", stripped):
        return float(stripped)
    return _parse_measurement(stripped)


def _format_line_spacing(value: object, raw_target: str) -> str:
    if value is None:
        return ""
    if isinstance(value, Length):
        return _format_length(value, raw_target)
    if isinstance(value, (int, float)):
        return f"{value:g}"
    return _format_length(value, raw_target)


def _parse_measurement(value: str) -> Length:
    match = re.fullmatch(r"\s*(-?\d+(?:\.\d+)?)\s*(cm|pt|in)\s*", value)
    if match is None:
        raise ValueError(f"unsupported measurement value: {value}")
    amount = float(match.group(1))
    unit = match.group(2)
    if unit == "cm":
        return Cm(amount)
    if unit == "pt":
        return Pt(amount)
    return Inches(amount)


def _format_length(value: object, raw_target: str) -> str:
    if value is None:
        return ""
    if not isinstance(value, Length):
        return str(value)
    unit_match = re.fullmatch(r"\s*-?\d+(?:\.\d+)?\s*(cm|pt|in)\s*", raw_target)
    unit = unit_match.group(1) if unit_match is not None else "pt"
    if unit == "cm":
        return f"{value.cm:g}cm"
    if unit == "in":
        return f"{value.inches:g}in"
    return f"{value.pt:g}pt"


def _format_run_length_state(value: object, raw_target: str) -> str:
    if not isinstance(value, _RunLengthState):
        return _format_length(value, raw_target)
    if not value.values:
        return ""

    formatted_values = [_format_length(item, raw_target) for item in value.values]
    if len(set(formatted_values)) == 1:
        return formatted_values[0]
    return f"mixed[{'|'.join(formatted_values)}]"


def _iter_table_paragraphs(table: Table):
    for row in table.rows:
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                yield paragraph


def _iter_table_paragraphs_in_rows(table: Table, *, row_selector: str):
    if row_selector == "header":
        row_indexes = range(0, min(1, len(table.rows)))
    elif row_selector == "body":
        row_indexes = range(1, len(table.rows))
    else:
        raise ValueError(f"unsupported table row selector: {row_selector}")

    for row_index in row_indexes:
        for cell in table.rows[row_index].cells:
            for paragraph in cell.paragraphs:
                yield paragraph


def _table_paragraphs_for_property(table: Table, property_name: str):
    if property_name.startswith("header_row_"):
        yield from _iter_table_paragraphs_in_rows(table, row_selector="header")
        return
    if property_name.startswith("body_rows_"):
        yield from _iter_table_paragraphs_in_rows(table, row_selector="body")
        return
    yield from _iter_table_paragraphs(table)


def _special_object_target_object_type(
    rules: list[SpecialObjectRule],
    rule_id: str | None,
) -> str | None:
    if rule_id is None:
        return None
    for rule in rules:
        if rule.rule_id == rule_id:
            return rule.target_object_type
    return None


def _report_row(
    *,
    object_id: str,
    object_type_before: str,
    object_type_after: str,
    location: str,
    text_preview: str,
    property_name: str,
    before: str,
    after: str,
    rule_id: str,
    status: ReportStatus,
    reason: str,
) -> dict[str, str]:
    return {
        "object_id": object_id,
        "object_type_before": object_type_before,
        "object_type_after": object_type_after,
        "location": location,
        "text_preview": text_preview,
        "property": property_name,
        "before": before,
        "after": after,
        "rule_id": rule_id,
        "status": status,
        "reason": reason,
    }
