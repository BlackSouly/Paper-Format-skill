from paper_format_normalizer.classify import classify_document
from paper_format_normalizer.model import (
    NumberingRule,
    ParagraphRule,
    ReportSchemaField,
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


def test_exact_text_beats_default_body_rule() -> None:
    parsed = ParsedDocument(
        headers=(),
        body_items=(ParsedBodyParagraph(text="Abstract", style_name="Body Text"),),
        paragraphs=(ParsedBodyParagraph(text="Abstract", style_name="Body Text"),),
        tables=(),
    )
    rules = _rule_set(
        paragraph_rules=[
            ParagraphRule(
                rule_id="PAR-DEFAULT",
                priority=100,
                match_type="default",
                match_value="body",
                target_property="object_type",
                target_value="body",
            ),
            ParagraphRule(
                rule_id="PAR-EXACT",
                priority=300,
                match_type="text",
                match_value="Abstract",
                target_property="object_type",
                target_value="abstract",
            ),
        ]
    )

    result = classify_document(parsed, rules)

    assert len(result.object_results) == 1
    assert result.object_results[0].status == "matched"
    assert result.object_results[0].matched_rule_id == "PAR-EXACT"
    assert result.object_results[0].match_kind == "exact_text"


def test_regex_match_works() -> None:
    parsed = ParsedDocument(
        headers=(),
        body_items=(ParsedBodyParagraph(text="Keywords: alpha, beta", style_name=None),),
        paragraphs=(ParsedBodyParagraph(text="Keywords: alpha, beta", style_name=None),),
        tables=(),
    )
    rules = _rule_set(
        paragraph_rules=[
            ParagraphRule(
                rule_id="PAR-REGEX",
                priority=200,
                match_type="regex",
                match_value=r"^Keywords:",
                target_property="object_type",
                target_value="keywords",
            )
        ]
    )

    result = classify_document(parsed, rules)

    assert result.object_results[0].status == "matched"
    assert result.object_results[0].matched_rule_id == "PAR-REGEX"
    assert result.object_results[0].match_kind == "pattern"


def test_word_object_match_classifies_header_item() -> None:
    header_item = ParsedBodyParagraph(text="Section 1", style_name="Header")
    parsed = ParsedDocument(
        headers=(ParsedHeader(variant="default", section_indices=(0,), items=(header_item,)),),
        body_items=(),
        paragraphs=(),
        tables=(),
    )
    rules = _rule_set(
        special_object_rules=[
            SpecialObjectRule(
                rule_id="HDR-STRUCT",
                priority=500,
                object_type="header",
                match_type="text",
                match_value="Section 1",
                target_object_type="running_header",
            )
        ]
    )

    result = classify_document(parsed, rules)

    assert result.object_results[0].status == "matched"
    assert result.object_results[0].matched_rule_id == "HDR-STRUCT"
    assert result.object_results[0].object_type == "header"
    assert result.object_results[0].match_kind == "structural"


def test_default_body_classification_works() -> None:
    paragraph = ParsedBodyParagraph(text="Plain body text", style_name="Body Text")
    parsed = ParsedDocument(
        headers=(),
        body_items=(paragraph,),
        paragraphs=(paragraph,),
        tables=(),
    )
    rules = _rule_set(
        paragraph_rules=[
            ParagraphRule(
                rule_id="PAR-DEFAULT",
                priority=100,
                match_type="default",
                match_value="body",
                target_property="object_type",
                target_value="body",
            )
        ]
    )

    result = classify_document(parsed, rules)

    assert result.object_results[0].status == "matched"
    assert result.object_results[0].matched_rule_id == "PAR-DEFAULT"
    assert result.object_results[0].match_kind == "default_body"


def test_ambiguous_top_ranked_rules_become_unresolved() -> None:
    paragraph = ParsedBodyParagraph(text="Abstract", style_name=None)
    parsed = ParsedDocument(
        headers=(),
        body_items=(paragraph,),
        paragraphs=(paragraph,),
        tables=(),
    )
    rules = _rule_set(
        paragraph_rules=[
            ParagraphRule(
                rule_id="PAR-A",
                priority=300,
                match_type="text",
                match_value="Abstract",
                target_property="object_type",
                target_value="abstract",
            ),
            ParagraphRule(
                rule_id="PAR-B",
                priority=300,
                match_type="text",
                match_value="Abstract",
                target_property="object_type",
                target_value="summary",
            ),
        ]
    )

    result = classify_document(parsed, rules)

    assert result.object_results[0].status == "unresolved"
    assert result.object_results[0].matched_rule_id is None
    assert "PAR-A" in result.object_results[0].reason
    assert "PAR-B" in result.object_results[0].reason
    assert "same priority" in result.object_results[0].reason


def test_lower_numeric_priority_wins_within_same_match_kind() -> None:
    paragraph = ParsedBodyParagraph(text="Abstract", style_name=None)
    parsed = ParsedDocument(
        headers=(),
        body_items=(paragraph,),
        paragraphs=(paragraph,),
        tables=(),
    )
    rules = _rule_set(
        paragraph_rules=[
            ParagraphRule(
                rule_id="PAR-LOW",
                priority=10,
                match_type="text",
                match_value="Abstract",
                target_property="object_type",
                target_value="abstract",
            ),
            ParagraphRule(
                rule_id="PAR-HIGH",
                priority=20,
                match_type="text",
                match_value="Abstract",
                target_property="object_type",
                target_value="summary",
            ),
        ]
    )

    result = classify_document(parsed, rules)

    assert result.object_results[0].status == "matched"
    assert result.object_results[0].matched_rule_id == "PAR-LOW"
    assert result.object_results[0].match_kind == "exact_text"


def test_unmatched_object_becomes_unresolved() -> None:
    paragraph = ParsedBodyParagraph(text="No matching rule here", style_name=None)
    parsed = ParsedDocument(
        headers=(),
        body_items=(paragraph,),
        paragraphs=(paragraph,),
        tables=(),
    )
    rules = _rule_set()

    result = classify_document(parsed, rules)

    assert result.object_results[0].status == "unresolved"
    assert result.object_results[0].matched_rule_id is None
    assert result.object_results[0].reason == "no matching classification rule"


def test_numbering_rule_applies_after_structural_rules() -> None:
    paragraph = ParsedBodyParagraph(text="1. Scope", style_name="Heading 1")
    parsed = ParsedDocument(
        headers=(),
        body_items=(paragraph,),
        paragraphs=(paragraph,),
        tables=(),
    )
    rules = _rule_set(
        numbering_rules=[
            NumberingRule(
                rule_id="NUM-HEADING",
                priority=400,
                match_type="style",
                match_value="Heading 1",
                target_property="level",
                target_value="1",
            )
        ]
    )

    result = classify_document(parsed, rules)

    assert result.object_results[0].status == "matched"
    assert result.object_results[0].matched_rule_id == "NUM-HEADING"
    assert result.object_results[0].match_kind == "numbering"


def test_header_table_is_reported_as_unresolved_when_no_rule_matches() -> None:
    header_table = ParsedBodyTable(rows=(("Header Cell",),))
    parsed = ParsedDocument(
        headers=(ParsedHeader(variant="default", section_indices=(0,), items=(header_table,)),),
        body_items=(),
        paragraphs=(),
        tables=(),
    )
    rules = _rule_set()

    result = classify_document(parsed, rules)

    assert len(result.object_results) == 1
    assert result.object_results[0].object_type == "header_table"
    assert result.object_results[0].status == "unresolved"
    assert result.object_results[0].location == "headers[0].items[0]"
    assert result.object_results[0].original_text == "Header Cell"


def test_header_table_can_match_header_special_object_rule() -> None:
    header_table = ParsedBodyTable(rows=(("Running Header",),))
    parsed = ParsedDocument(
        headers=(ParsedHeader(variant="default", section_indices=(0,), items=(header_table,)),),
        body_items=(),
        paragraphs=(),
        tables=(),
    )
    rules = _rule_set(
        special_object_rules=[
            SpecialObjectRule(
                rule_id="HDR-TABLE",
                priority=10,
                object_type="header",
                match_type="text",
                match_value="Running Header",
                target_object_type="running_header",
            )
        ]
    )

    result = classify_document(parsed, rules)

    assert len(result.object_results) == 1
    assert result.object_results[0].status == "matched"
    assert result.object_results[0].matched_rule_id == "HDR-TABLE"
    assert result.object_results[0].match_kind == "structural"
    assert result.object_results[0].object_type == "header_table"


def test_header_table_prefers_header_semantics_over_generic_table_rules() -> None:
    header_table = ParsedBodyTable(rows=(("Running Header",),))
    parsed = ParsedDocument(
        headers=(ParsedHeader(variant="default", section_indices=(0,), items=(header_table,)),),
        body_items=(),
        paragraphs=(),
        tables=(),
    )
    rules = _rule_set(
        special_object_rules=[
            SpecialObjectRule(
                rule_id="HDR-TABLE",
                priority=20,
                object_type="header",
                match_type="text",
                match_value="Running Header",
                target_object_type="running_header",
            )
        ],
        table_rules=[
            TableRule(
                rule_id="TAB-TEXT",
                priority=10,
                match_type="text",
                match_value="Running Header",
                target_property="object_type",
                target_value="table_caption",
            )
        ],
    )

    result = classify_document(parsed, rules)

    assert len(result.object_results) == 1
    assert result.object_results[0].status == "matched"
    assert result.object_results[0].matched_rule_id == "HDR-TABLE"
    assert result.object_results[0].match_kind == "structural"
    assert result.object_results[0].object_type == "header_table"


def _rule_set(
    *,
    paragraph_rules: list[ParagraphRule] | None = None,
    numbering_rules: list[NumberingRule] | None = None,
    special_object_rules: list[SpecialObjectRule] | None = None,
    table_rules: list[TableRule] | None = None,
) -> RuleSet:
    return RuleSet(
        document_rules=[],
        paragraph_rules=paragraph_rules or [],
        numbering_rules=numbering_rules or [],
        table_rules=table_rules or [],
        special_object_rules=special_object_rules or [],
        report_schema=[
            ReportSchemaField(column_name="object_id", order=1, description="Object ID")
        ],
    )
