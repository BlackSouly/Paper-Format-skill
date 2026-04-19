from pathlib import Path

import pytest

from paper_format_normalizer.model import RuleSet
from paper_format_normalizer.rules import load_rule_set


def test_load_rule_set_reads_all_csv_files() -> None:
    rules = load_rule_set(Path("tests/fixtures/sample_rules"))

    assert isinstance(rules, RuleSet)
    assert [rule.rule_id for rule in rules.document_rules] == ["DOC-002", "DOC-001"]
    assert [rule.rule_id for rule in rules.paragraph_rules] == [
        "PAR-002",
        "PAR-003",
        "PAR-001",
    ]
    assert [rule.rule_id for rule in rules.numbering_rules] == ["NUM-002", "NUM-001"]
    assert [rule.rule_id for rule in rules.table_rules] == ["TAB-002", "TAB-001"]
    assert [rule.rule_id for rule in rules.special_object_rules] == [
        "SPO-002",
        "SPO-001",
    ]
    assert [field.column_name for field in rules.report_schema] == [
        "object_id",
        "object_type_before",
        "object_type_after",
        "location",
        "text_preview",
        "property",
        "before",
        "after",
        "rule_id",
        "status",
        "reason",
    ]


def test_load_rule_set_rejects_missing_required_columns(tmp_path: Path) -> None:
    broken = tmp_path / "rules"
    broken.mkdir()
    (broken / "document_rules.csv").write_text("rule_id\nDOC-001\n", encoding="utf-8")
    (broken / "paragraph_rules.csv").write_text(
        "rule_id,priority,match_type,match_value,target_property,target_value\n"
        "PAR-001,10,text,Title,font_name,SimSun\n",
        encoding="utf-8",
    )
    (broken / "numbering_rules.csv").write_text(
        "rule_id,priority,match_type,match_value,target_property,target_value\n"
        "NUM-001,10,style,List Number,level,1\n",
        encoding="utf-8",
    )
    (broken / "table_rules.csv").write_text(
        "rule_id,priority,match_type,match_value,target_property,target_value\n"
        "TAB-001,10,tag,Table,alignment,center\n",
        encoding="utf-8",
    )
    (broken / "special_object_rules.csv").write_text(
        "rule_id,priority,object_type,match_type,match_value,target_object_type\n"
        "SPO-001,10,header,text,Section,header\n",
        encoding="utf-8",
    )
    (broken / "report_schema.csv").write_text(
        "column_name,order\nobject_id,1\n",
        encoding="utf-8",
    )

    with pytest.raises(ValueError, match="Unexpected header schema in document_rules.csv"):
        load_rule_set(broken)


def test_load_rule_set_rejects_duplicate_rule_id(tmp_path: Path) -> None:
    rules_dir = _write_valid_rules_dir(tmp_path)
    (rules_dir / "paragraph_rules.csv").write_text(
        "rule_id,priority,match_type,match_value,target_property,target_value\n"
        "PAR-001,20,text,Abstract,font_name,SimSun\n"
        "PAR-001,10,style,Body Text,font_name,Times New Roman\n",
        encoding="utf-8",
    )

    with pytest.raises(ValueError, match="Duplicate rule_id 'PAR-001'"):
        load_rule_set(rules_dir)


def test_load_rule_set_rejects_duplicate_report_schema_column_or_order(
    tmp_path: Path,
) -> None:
    rules_dir = _write_valid_rules_dir(tmp_path)
    (rules_dir / "report_schema.csv").write_text(
        "column_name,order,description\n"
        "object_id,1,Stable object identifier\n"
        "object_id,2,Duplicate column name\n",
        encoding="utf-8",
    )

    with pytest.raises(ValueError, match="Duplicate column_name 'object_id'"):
        load_rule_set(rules_dir)

    (rules_dir / "report_schema.csv").write_text(
        "column_name,order,description\n"
        "object_id,1,Stable object identifier\n"
        "status,1,Duplicate order\n",
        encoding="utf-8",
    )

    with pytest.raises(ValueError, match="Duplicate order '1'"):
        load_rule_set(rules_dir)


def test_load_rule_set_rejects_blank_required_cell(tmp_path: Path) -> None:
    rules_dir = _write_valid_rules_dir(tmp_path)
    (rules_dir / "document_rules.csv").write_text(
        "rule_id,priority,property_name,value,scope\n"
        "DOC-001,20,,2.54cm,document\n",
        encoding="utf-8",
    )

    with pytest.raises(ValueError, match="Missing required value for column 'property_name'"):
        load_rule_set(rules_dir)


def test_load_rule_set_rejects_invalid_integer(tmp_path: Path) -> None:
    rules_dir = _write_valid_rules_dir(tmp_path)
    (rules_dir / "document_rules.csv").write_text(
        "rule_id,priority,property_name,value,scope\n"
        "DOC-001,high,page_margin_top,2.54cm,document\n",
        encoding="utf-8",
    )

    with pytest.raises(ValueError, match="Invalid integer value for column 'priority'"):
        load_rule_set(rules_dir)


def test_load_rule_set_rejects_whitespace_corrupted_value(tmp_path: Path) -> None:
    rules_dir = _write_valid_rules_dir(tmp_path)
    (rules_dir / "paragraph_rules.csv").write_text(
        "rule_id,priority,match_type,match_value,target_property,target_value\n"
        "PAR-001,20, text,Abstract,font_name,SimSun\n",
        encoding="utf-8",
    )

    with pytest.raises(ValueError, match="Whitespace-corrupted value for column 'match_type'"):
        load_rule_set(rules_dir)


def test_load_rule_set_rejects_duplicate_header(tmp_path: Path) -> None:
    rules_dir = _write_valid_rules_dir(tmp_path)
    (rules_dir / "document_rules.csv").write_text(
        "rule_id,priority,property_name,property_name,scope\n"
        "DOC-001,20,page_margin_top,2.54cm,document\n",
        encoding="utf-8",
    )

    with pytest.raises(ValueError, match="Duplicate header in document_rules.csv"):
        load_rule_set(rules_dir)


def test_load_rule_set_rejects_unexpected_header(tmp_path: Path) -> None:
    rules_dir = _write_valid_rules_dir(tmp_path)
    (rules_dir / "document_rules.csv").write_text(
        "rule_id,priority,scope,property_name,value\n"
        "DOC-001,20,document,page_margin_top,2.54cm\n",
        encoding="utf-8",
    )

    with pytest.raises(ValueError, match="Unexpected header schema in document_rules.csv"):
        load_rule_set(rules_dir)


def test_load_rule_set_rejects_overlong_row_with_extra_cell(tmp_path: Path) -> None:
    rules_dir = _write_valid_rules_dir(tmp_path)
    (rules_dir / "paragraph_rules.csv").write_text(
        "rule_id,priority,match_type,match_value,target_property,target_value\n"
        "PAR-001,10,text,Abstract,font_name,SimSun,EXTRA\n",
        encoding="utf-8",
    )

    with pytest.raises(ValueError, match="Extra unmapped cells in paragraph_rules.csv row 2"):
        load_rule_set(rules_dir)


def _write_valid_rules_dir(tmp_path: Path) -> Path:
    rules_dir = tmp_path / "rules"
    rules_dir.mkdir()
    (rules_dir / "document_rules.csv").write_text(
        "rule_id,priority,property_name,value,scope\n"
        "DOC-001,20,page_margin_top,2.54cm,document\n",
        encoding="utf-8",
    )
    (rules_dir / "paragraph_rules.csv").write_text(
        "rule_id,priority,match_type,match_value,target_property,target_value\n"
        "PAR-001,10,text,Abstract,font_name,SimSun\n",
        encoding="utf-8",
    )
    (rules_dir / "numbering_rules.csv").write_text(
        "rule_id,priority,match_type,match_value,target_property,target_value\n"
        "NUM-001,10,style,Heading 1,level,1\n",
        encoding="utf-8",
    )
    (rules_dir / "table_rules.csv").write_text(
        "rule_id,priority,match_type,match_value,target_property,target_value\n"
        "TAB-001,10,style,Table Grid,border_style,single\n",
        encoding="utf-8",
    )
    (rules_dir / "special_object_rules.csv").write_text(
        "rule_id,priority,object_type,match_type,match_value,target_object_type\n"
        "SPO-001,10,header,text,Section,header\n",
        encoding="utf-8",
    )
    (rules_dir / "report_schema.csv").write_text(
        "column_name,order,description\n"
        "object_id,1,Stable object identifier\n"
        "status,2,Change status\n",
        encoding="utf-8",
    )
    return rules_dir
