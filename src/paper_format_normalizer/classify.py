from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Literal

from paper_format_normalizer.model import (
    NumberingRule,
    ParagraphRule,
    RuleSet,
    SpecialObjectRule,
    TableRule,
)
from paper_format_normalizer.parse import (
    ParsedBodyParagraph,
    ParsedBodyTable,
    ParsedDocument,
    ParsedHeader,
)

ClassificationStatus = Literal["matched", "unresolved"]
MatchKind = Literal["exact_text", "pattern", "structural", "numbering", "default_body"]

_MATCH_KIND_ORDER: dict[MatchKind, int] = {
    "exact_text": 0,
    "pattern": 1,
    "structural": 2,
    "numbering": 3,
    "default_body": 4,
}


@dataclass(frozen=True)
class ClassificationCandidate:
    rule_id: str
    priority: int
    match_kind: MatchKind


@dataclass(frozen=True)
class ClassifiedObjectResult:
    object_id: str
    location: str
    object_type: str
    original_text: str
    status: ClassificationStatus
    matched_rule_id: str | None
    match_kind: MatchKind | None
    reason: str | None


@dataclass(frozen=True)
class DocumentClassification:
    object_results: tuple[ClassifiedObjectResult, ...]


def classify_document(
    document: ParsedDocument,
    rule_set: RuleSet,
) -> DocumentClassification:
    results: list[ClassifiedObjectResult] = []

    for header_index, header in enumerate(document.headers):
        for item_index, item in enumerate(header.items):
            if isinstance(item, ParsedBodyParagraph):
                results.append(
                    _classify_header_paragraph(
                        header=header,
                        header_index=header_index,
                        item=item,
                        item_index=item_index,
                        rule_set=rule_set,
                    )
                )
                continue
            results.append(
                _classify_header_table(
                    header=header,
                    header_index=header_index,
                    item=item,
                    item_index=item_index,
                    rule_set=rule_set,
                )
            )

    for body_index, item in enumerate(document.body_items):
        if isinstance(item, ParsedBodyParagraph):
            results.append(
                _classify_body_paragraph(
                    item=item,
                    body_index=body_index,
                    rule_set=rule_set,
                )
            )
            continue
        results.append(
            _classify_table(
                item=item,
                body_index=body_index,
                rule_set=rule_set,
            )
        )

    return DocumentClassification(object_results=tuple(results))


def _classify_header_paragraph(
    *,
    header: ParsedHeader,
    header_index: int,
    item: ParsedBodyParagraph,
    item_index: int,
    rule_set: RuleSet,
) -> ClassifiedObjectResult:
    return _resolve_result(
        object_id=f"header-{header_index}-item-{item_index}",
        location=f"headers[{header_index}].items[{item_index}]",
        object_type="header",
        original_text=item.text,
        candidates=_header_special_candidates(rule_set.special_object_rules, item.text),
        no_match_reason=(
            "no matching classification rule for header object"
            f" ({header.variant}, sections={','.join(str(index) for index in header.section_indices)})"
        ),
    )


def _classify_body_paragraph(
    *,
    item: ParsedBodyParagraph,
    body_index: int,
    rule_set: RuleSet,
) -> ClassifiedObjectResult:
    candidates: list[ClassificationCandidate] = []
    for rule in rule_set.paragraph_rules:
        candidate = _paragraph_rule_candidate(rule, item)
        if candidate is not None:
            candidates.append(candidate)

    for rule in rule_set.numbering_rules:
        if _numbering_rule_matches(rule, item):
            candidates.append(
                ClassificationCandidate(
                    rule_id=rule.rule_id,
                    priority=rule.priority,
                    match_kind="numbering",
                )
            )

    return _resolve_result(
        object_id=f"body-{body_index}",
        location=f"body_items[{body_index}]",
        object_type="paragraph",
        original_text=item.text,
        candidates=candidates,
        no_match_reason="no matching classification rule",
    )


def _classify_header_table(
    *,
    header: ParsedHeader,
    header_index: int,
    item: ParsedBodyTable,
    item_index: int,
    rule_set: RuleSet,
) -> ClassifiedObjectResult:
    flattened_text = _flatten_table_text(item)
    return _resolve_result(
        object_id=f"header-{header_index}-item-{item_index}",
        location=f"headers[{header_index}].items[{item_index}]",
        object_type="header_table",
        original_text=flattened_text,
        candidates=_header_special_candidates(
            rule_set.special_object_rules,
            flattened_text,
        ),
        no_match_reason=(
            "no matching classification rule for header table"
            f" ({header.variant}, sections={','.join(str(index) for index in header.section_indices)})"
        ),
    )


def _classify_table(
    *,
    item: ParsedBodyTable,
    body_index: int,
    rule_set: RuleSet,
) -> ClassifiedObjectResult:
    flattened_text = _flatten_table_text(item)
    return _resolve_result(
        object_id=f"body-{body_index}",
        location=f"body_items[{body_index}]",
        object_type="table",
        original_text=flattened_text,
        candidates=_table_candidates(rule_set.table_rules, flattened_text),
        no_match_reason="no matching classification rule",
    )


def _paragraph_rule_candidate(
    rule: ParagraphRule,
    item: ParsedBodyParagraph,
) -> ClassificationCandidate | None:
    if rule.match_type == "text" and item.text == rule.match_value:
        return ClassificationCandidate(
            rule_id=rule.rule_id,
            priority=rule.priority,
            match_kind="exact_text",
        )
    if rule.match_type in {"regex", "pattern"} and re.search(rule.match_value, item.text):
        return ClassificationCandidate(
            rule_id=rule.rule_id,
            priority=rule.priority,
            match_kind="pattern",
        )
    if rule.match_type == "style" and item.style_name == rule.match_value:
        return ClassificationCandidate(
            rule_id=rule.rule_id,
            priority=rule.priority,
            match_kind="structural",
        )
    if rule.match_type == "default" and rule.match_value == "body":
        return ClassificationCandidate(
            rule_id=rule.rule_id,
            priority=rule.priority,
            match_kind="default_body",
        )
    return None


def _table_rule_candidate(
    rule: TableRule,
    flattened_text: str,
) -> ClassificationCandidate | None:
    if rule.match_type == "text" and flattened_text == rule.match_value:
        return ClassificationCandidate(
            rule_id=rule.rule_id,
            priority=rule.priority,
            match_kind="exact_text",
        )
    if rule.match_type in {"regex", "pattern"} and re.search(
        rule.match_value,
        flattened_text,
    ):
        return ClassificationCandidate(
            rule_id=rule.rule_id,
            priority=rule.priority,
            match_kind="pattern",
        )
    return None


def _table_candidates(
    rules: list[TableRule],
    flattened_text: str,
) -> list[ClassificationCandidate]:
    candidates: list[ClassificationCandidate] = []
    for rule in rules:
        candidate = _table_rule_candidate(rule, flattened_text)
        if candidate is not None:
            candidates.append(candidate)
    return candidates


def _header_special_candidates(
    rules: list[SpecialObjectRule],
    text: str,
) -> list[ClassificationCandidate]:
    candidates: list[ClassificationCandidate] = []
    for rule in rules:
        if rule.object_type != "header":
            continue
        if _special_object_rule_matches(rule, text):
            candidates.append(
                ClassificationCandidate(
                    rule_id=rule.rule_id,
                    priority=rule.priority,
                    match_kind="structural",
                )
            )
    return candidates


def _numbering_rule_matches(rule: NumberingRule, item: ParsedBodyParagraph) -> bool:
    if rule.match_type == "style":
        return item.style_name == rule.match_value
    if rule.match_type == "text":
        return item.text == rule.match_value
    if rule.match_type in {"regex", "pattern"}:
        return re.search(rule.match_value, item.text) is not None
    return False


def _special_object_rule_matches(rule: SpecialObjectRule, text: str) -> bool:
    if rule.match_type == "text":
        return text == rule.match_value
    if rule.match_type in {"regex", "pattern"}:
        return re.search(rule.match_value, text) is not None
    return False


def _resolve_result(
    *,
    object_id: str,
    location: str,
    object_type: str,
    original_text: str,
    candidates: list[ClassificationCandidate],
    no_match_reason: str,
) -> ClassifiedObjectResult:
    if not candidates:
        return ClassifiedObjectResult(
            object_id=object_id,
            location=location,
            object_type=object_type,
            original_text=original_text,
            status="unresolved",
            matched_rule_id=None,
            match_kind=None,
            reason=no_match_reason,
        )

    ranked_candidates = sorted(
        candidates,
        key=lambda candidate: (
            _MATCH_KIND_ORDER[candidate.match_kind],
            candidate.priority,
            candidate.rule_id,
        ),
    )
    winner = ranked_candidates[0]
    conflicting_candidates = [
        candidate
        for candidate in ranked_candidates
        if candidate.match_kind == winner.match_kind
        and candidate.priority == winner.priority
    ]
    if len(conflicting_candidates) > 1:
        conflicting_rule_ids = ", ".join(
            candidate.rule_id for candidate in conflicting_candidates
        )
        return ClassifiedObjectResult(
            object_id=object_id,
            location=location,
            object_type=object_type,
            original_text=original_text,
            status="unresolved",
            matched_rule_id=None,
            match_kind=None,
            reason=(
                "conflicting rules at same priority: "
                f"{conflicting_rule_ids} "
                f"(match_kind={winner.match_kind}, priority={winner.priority})"
            ),
        )

    return ClassifiedObjectResult(
        object_id=object_id,
        location=location,
        object_type=object_type,
        original_text=original_text,
        status="matched",
        matched_rule_id=winner.rule_id,
        match_kind=winner.match_kind,
        reason=None,
    )


def _flatten_table_text(item: ParsedBodyTable) -> str:
    return "\n".join("\t".join(cell for cell in row) for row in item.rows)
