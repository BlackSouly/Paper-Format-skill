# Paper Format Normalizer Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build a rule-driven single-document paper formatter that reads structured CSV rules, normalizes a `.docx` paper to those rules, and emits both a new `.docx` file and a strict CSV change report.

**Architecture:** Use a small Python package with explicit stages: rule loading, document parsing, deterministic classification, normalization, and report emission. Keep classification and formatting separate so unsupported objects can fail or remain unresolved explicitly rather than being silently patched.

**Tech Stack:** Python 3.12, `python-docx`, `lxml`, `pytest`, `typer`

---

## File Structure

- Create: `C:\Users\admin\Desktop\JJBand\pyproject.toml`
- Create: `C:\Users\admin\Desktop\JJBand\src\paper_format_normalizer\__init__.py`
- Create: `C:\Users\admin\Desktop\JJBand\src\paper_format_normalizer\cli.py`
- Create: `C:\Users\admin\Desktop\JJBand\src\paper_format_normalizer\model.py`
- Create: `C:\Users\admin\Desktop\JJBand\src\paper_format_normalizer\rules.py`
- Create: `C:\Users\admin\Desktop\JJBand\src\paper_format_normalizer\parse.py`
- Create: `C:\Users\admin\Desktop\JJBand\src\paper_format_normalizer\classify.py`
- Create: `C:\Users\admin\Desktop\JJBand\src\paper_format_normalizer\normalize.py`
- Create: `C:\Users\admin\Desktop\JJBand\src\paper_format_normalizer\report.py`
- Create: `C:\Users\admin\Desktop\JJBand\src\paper_format_normalizer\conversion.py`
- Create: `C:\Users\admin\Desktop\JJBand\tests\conftest.py`
- Create: `C:\Users\admin\Desktop\JJBand\tests\test_rules.py`
- Create: `C:\Users\admin\Desktop\JJBand\tests\test_classify.py`
- Create: `C:\Users\admin\Desktop\JJBand\tests\test_normalize.py`
- Create: `C:\Users\admin\Desktop\JJBand\tests\test_cli.py`
- Create: `C:\Users\admin\Desktop\JJBand\tests\fixtures\sample_rules\document_rules.csv`
- Create: `C:\Users\admin\Desktop\JJBand\tests\fixtures\sample_rules\paragraph_rules.csv`
- Create: `C:\Users\admin\Desktop\JJBand\tests\fixtures\sample_rules\numbering_rules.csv`
- Create: `C:\Users\admin\Desktop\JJBand\tests\fixtures\sample_rules\table_rules.csv`
- Create: `C:\Users\admin\Desktop\JJBand\tests\fixtures\sample_rules\special_object_rules.csv`
- Create: `C:\Users\admin\Desktop\JJBand\tests\fixtures\sample_rules\report_schema.csv`
- Create: `C:\Users\admin\Desktop\JJBand\tests\fixtures\sample_docx_builder.py`
- Modify: `C:\Users\admin\Desktop\JJBand\skills\paper-format-normalizer\SKILL.md`

### Task 1: Bootstrap The Python Package

**Files:**
- Create: `C:\Users\admin\Desktop\JJBand\pyproject.toml`
- Create: `C:\Users\admin\Desktop\JJBand\src\paper_format_normalizer\__init__.py`
- Create: `C:\Users\admin\Desktop\JJBand\src\paper_format_normalizer\cli.py`
- Test: `C:\Users\admin\Desktop\JJBand\tests\test_cli.py`

- [ ] **Step 1: Write the failing CLI smoke test**

```python
from typer.testing import CliRunner

from paper_format_normalizer.cli import app


def test_cli_shows_help() -> None:
    runner = CliRunner()
    result = runner.invoke(app, ["--help"])
    assert result.exit_code == 0
    assert "normalize" in result.stdout
```

- [ ] **Step 2: Run test to verify it fails**

Run: `pytest C:\Users\admin\Desktop\JJBand\tests\test_cli.py::test_cli_shows_help -v`
Expected: FAIL with `ModuleNotFoundError` or import failure because the package does not exist yet.

- [ ] **Step 3: Write minimal package and CLI implementation**

```toml
[project]
name = "paper-format-normalizer"
version = "0.1.0"
requires-python = ">=3.12"
dependencies = [
  "lxml>=5.2.0",
  "python-docx>=1.1.2",
  "typer>=0.12.3",
]

[project.optional-dependencies]
dev = [
  "pytest>=8.2.0",
]

[project.scripts]
paper-format-normalizer = "paper_format_normalizer.cli:main"

[tool.pytest.ini_options]
testpaths = ["tests"]
pythonpath = ["src"]
```

```python
import typer

app = typer.Typer(help="Normalize one paper document from CSV formatting rules.")


@app.command()
def normalize() -> None:
    """Placeholder command body for initial CLI wiring."""
    raise typer.Exit(code=0)


def main() -> None:
    app()
```

- [ ] **Step 4: Run test to verify it passes**

Run: `pytest C:\Users\admin\Desktop\JJBand\tests\test_cli.py::test_cli_shows_help -v`
Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add C:\Users\admin\Desktop\JJBand\pyproject.toml C:\Users\admin\Desktop\JJBand\src\paper_format_normalizer\__init__.py C:\Users\admin\Desktop\JJBand\src\paper_format_normalizer\cli.py C:\Users\admin\Desktop\JJBand\tests\test_cli.py
git commit -m "feat: bootstrap paper format normalizer package"
```

### Task 2: Define Rule Models And CSV Loading

**Files:**
- Create: `C:\Users\admin\Desktop\JJBand\src\paper_format_normalizer\model.py`
- Create: `C:\Users\admin\Desktop\JJBand\src\paper_format_normalizer\rules.py`
- Create: `C:\Users\admin\Desktop\JJBand\tests\test_rules.py`
- Create: `C:\Users\admin\Desktop\JJBand\tests\fixtures\sample_rules\document_rules.csv`
- Create: `C:\Users\admin\Desktop\JJBand\tests\fixtures\sample_rules\paragraph_rules.csv`
- Create: `C:\Users\admin\Desktop\JJBand\tests\fixtures\sample_rules\numbering_rules.csv`
- Create: `C:\Users\admin\Desktop\JJBand\tests\fixtures\sample_rules\table_rules.csv`
- Create: `C:\Users\admin\Desktop\JJBand\tests\fixtures\sample_rules\special_object_rules.csv`
- Create: `C:\Users\admin\Desktop\JJBand\tests\fixtures\sample_rules\report_schema.csv`

- [ ] **Step 1: Write failing tests for rule loading and validation**

```python
from pathlib import Path

import pytest

from paper_format_normalizer.rules import RuleSet, load_rule_set


def test_load_rule_set_reads_all_csv_files(tmp_path: Path) -> None:
    rules = load_rule_set(Path("tests/fixtures/sample_rules"))
    assert isinstance(rules, RuleSet)
    assert len(rules.paragraph_rules) >= 1
    assert len(rules.special_rules) >= 1


def test_load_rule_set_rejects_missing_required_columns(tmp_path: Path) -> None:
    broken = tmp_path / "rules"
    broken.mkdir()
    (broken / "document_rules.csv").write_text("rule_id\nDOC-001\n", encoding="utf-8")
    for name in [
        "paragraph_rules.csv",
        "numbering_rules.csv",
        "table_rules.csv",
        "special_object_rules.csv",
        "report_schema.csv",
    ]:
        (broken / name).write_text("", encoding="utf-8")

    with pytest.raises(ValueError, match="Missing required columns"):
        load_rule_set(broken)
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `pytest C:\Users\admin\Desktop\JJBand\tests\test_rules.py -v`
Expected: FAIL because `load_rule_set` and `RuleSet` are undefined.

- [ ] **Step 3: Implement the rule models and CSV loader**

```python
from dataclasses import dataclass
from pathlib import Path


@dataclass(frozen=True)
class ParagraphRule:
    rule_id: str
    object_type: str
    match_type: str
    match_value: str
    priority: int


@dataclass(frozen=True)
class RuleSet:
    paragraph_rules: list[ParagraphRule]
    special_rules: list[dict[str, str]]


def load_rule_set(root: Path) -> RuleSet:
    required = {
        "document_rules.csv",
        "paragraph_rules.csv",
        "numbering_rules.csv",
        "table_rules.csv",
        "special_object_rules.csv",
        "report_schema.csv",
    }
    missing = [name for name in required if not (root / name).exists()]
    if missing:
        raise ValueError(f"Missing rule files: {', '.join(sorted(missing))}")
    return RuleSet(paragraph_rules=[], special_rules=[])
```

- [ ] **Step 4: Expand implementation until tests pass with real validation**

Run: `pytest C:\Users\admin\Desktop\JJBand\tests\test_rules.py -v`
Expected: PASS after implementing:
- CSV parsing with `csv.DictReader`
- required-column checks per file type
- priority coercion to `int`
- deterministic sorting by `priority` then `rule_id`

- [ ] **Step 5: Commit**

```bash
git add C:\Users\admin\Desktop\JJBand\src\paper_format_normalizer\model.py C:\Users\admin\Desktop\JJBand\src\paper_format_normalizer\rules.py C:\Users\admin\Desktop\JJBand\tests\test_rules.py C:\Users\admin\Desktop\JJBand\tests\fixtures\sample_rules
git commit -m "feat: load and validate formatting rule tables"
```

### Task 3: Add Parsed Document Models And Fixture Builder

**Files:**
- Create: `C:\Users\admin\Desktop\JJBand\src\paper_format_normalizer\parse.py`
- Create: `C:\Users\admin\Desktop\JJBand\tests\conftest.py`
- Create: `C:\Users\admin\Desktop\JJBand\tests\fixtures\sample_docx_builder.py`

- [ ] **Step 1: Write the failing parse test**

```python
from paper_format_normalizer.parse import parse_docx


def test_parse_docx_extracts_paragraphs_tables_and_headers(sample_docx_path) -> None:
    parsed = parse_docx(sample_docx_path)
    assert parsed.paragraphs
    assert parsed.headers
    assert parsed.tables
```

- [ ] **Step 2: Run test to verify it fails**

Run: `pytest C:\Users\admin\Desktop\JJBand\tests\test_classify.py::test_parse_docx_extracts_paragraphs_tables_and_headers -v`
Expected: FAIL because `parse_docx` and the sample document fixture do not exist.

- [ ] **Step 3: Build the fixture generator and parsed models**

```python
from dataclasses import dataclass
from pathlib import Path


@dataclass
class ParsedParagraph:
    object_id: str
    text: str
    style_name: str | None
    in_table: bool


@dataclass
class ParsedDocument:
    paragraphs: list[ParsedParagraph]
    headers: list[ParsedParagraph]
    tables: list[list[str]]


def parse_docx(path: Path) -> ParsedDocument:
    ...
```

- [ ] **Step 4: Run the parse-focused tests**

Run: `pytest C:\Users\admin\Desktop\JJBand\tests\test_classify.py::test_parse_docx_extracts_paragraphs_tables_and_headers -v`
Expected: PASS after the builder creates a DOCX with a header paragraph, body paragraphs, and a table, and the parser extracts them into deterministic objects.

- [ ] **Step 5: Commit**

```bash
git add C:\Users\admin\Desktop\JJBand\src\paper_format_normalizer\parse.py C:\Users\admin\Desktop\JJBand\tests\conftest.py C:\Users\admin\Desktop\JJBand\tests\fixtures\sample_docx_builder.py
git commit -m "feat: parse docx structure into typed objects"
```

### Task 4: Implement Deterministic Classification

**Files:**
- Create: `C:\Users\admin\Desktop\JJBand\src\paper_format_normalizer\classify.py`
- Create: `C:\Users\admin\Desktop\JJBand\tests\test_classify.py`

- [ ] **Step 1: Write failing classification tests**

```python
from paper_format_normalizer.classify import classify_document


def test_classify_document_prefers_exact_text_over_default(sample_docx_path, sample_rules) -> None:
    classified = classify_document(sample_docx_path, sample_rules)
    abstract_title = next(item for item in classified if item.text == "摘要")
    assert abstract_title.object_type_after == "abstract_title"
    assert abstract_title.rule_id == "TXT-002"


def test_classify_document_marks_ambiguous_objects_unresolved(sample_docx_path, conflicting_rules) -> None:
    classified = classify_document(sample_docx_path, conflicting_rules)
    assert any(item.status == "unresolved" for item in classified)
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `pytest C:\Users\admin\Desktop\JJBand\tests\test_classify.py -v`
Expected: FAIL because the classifier does not exist yet.

- [ ] **Step 3: Implement classification with explicit precedence**

```python
def classify_object(obj, rules):
    candidates = []
    candidates.extend(match_exact_text(obj, rules))
    candidates.extend(match_pattern(obj, rules))
    candidates.extend(match_word_structure(obj, rules))
    candidates.extend(match_numbering(obj, rules))
    candidates.extend(match_default_body(obj, rules))
    ranked = sorted(candidates, key=lambda item: (item.priority, item.rule_id))
    if len(ranked) >= 2 and ranked[0].priority == ranked[1].priority:
        return unresolved_result(obj, "Conflicting rules at the same priority")
    if not ranked:
        return unresolved_result(obj, "No deterministic rule match")
    return matched_result(obj, ranked[0])
```

- [ ] **Step 4: Run tests to verify precedence and ambiguity handling**

Run: `pytest C:\Users\admin\Desktop\JJBand\tests\test_classify.py -v`
Expected: PASS with coverage for:
- exact text match
- regex match
- word-object match for footnotes or headers
- default body classification
- unresolved output for ambiguous or unmatched objects

- [ ] **Step 5: Commit**

```bash
git add C:\Users\admin\Desktop\JJBand\src\paper_format_normalizer\classify.py C:\Users\admin\Desktop\JJBand\tests\test_classify.py
git commit -m "feat: classify document objects from deterministic evidence"
```

### Task 5: Implement Paragraph And Document Normalization

**Files:**
- Create: `C:\Users\admin\Desktop\JJBand\src\paper_format_normalizer\normalize.py`
- Create: `C:\Users\admin\Desktop\JJBand\tests\test_normalize.py`

- [ ] **Step 1: Write failing normalization tests**

```python
from docx import Document

from paper_format_normalizer.normalize import normalize_document


def test_normalize_document_resets_body_paragraph_font_and_spacing(tmp_path, sample_docx_path, sample_rules) -> None:
    output_path, report_rows = normalize_document(sample_docx_path, sample_rules, tmp_path)
    document = Document(output_path)
    paragraph = next(p for p in document.paragraphs if p.text == "这是正文第一段。")
    assert paragraph.paragraph_format.first_line_indent is not None
    assert report_rows
```
 
- [ ] **Step 2: Run tests to verify they fail**

Run: `pytest C:\Users\admin\Desktop\JJBand\tests\test_normalize.py -v`
Expected: FAIL because no normalization pipeline exists yet.

- [ ] **Step 3: Implement minimal normalization pipeline**

```python
from pathlib import Path


def normalize_document(input_path: Path, rules, output_dir: Path):
    classified = classify_document(input_path, rules)
    # open the original docx, apply paragraph/document-level rules, save copy
    # return saved path and report rows
```

- [ ] **Step 4: Expand implementation until formatting and output assertions pass**

Run: `pytest C:\Users\admin\Desktop\JJBand\tests\test_normalize.py -v`
Expected: PASS with coverage for:
- non-destructive output naming
- page-margin normalization
- heading/body font resets
- paragraph indent and spacing resets
- unresolved items preserved in report without silent modification

- [ ] **Step 5: Commit**

```bash
git add C:\Users\admin\Desktop\JJBand\src\paper_format_normalizer\normalize.py C:\Users\admin\Desktop\JJBand\tests\test_normalize.py
git commit -m "feat: normalize docx paragraphs and page settings"
```

### Task 6: Emit Strict CSV Audit Reports

**Files:**
- Create: `C:\Users\admin\Desktop\JJBand\src\paper_format_normalizer\report.py`
- Modify: `C:\Users\admin\Desktop\JJBand\src\paper_format_normalizer\normalize.py`
- Modify: `C:\Users\admin\Desktop\JJBand\tests\test_normalize.py`

- [ ] **Step 1: Write the failing report test**

```python
import csv

from paper_format_normalizer.normalize import normalize_document


def test_normalize_document_writes_strict_change_report(tmp_path, sample_docx_path, sample_rules) -> None:
    output_path, report_path = normalize_document(sample_docx_path, sample_rules, tmp_path)
    rows = list(csv.DictReader(report_path.open("r", encoding="utf-8-sig", newline="")))
    assert rows
    assert {"object_id", "property", "before", "after", "rule_id", "status"} <= set(rows[0])
```

- [ ] **Step 2: Run test to verify it fails**

Run: `pytest C:\Users\admin\Desktop\JJBand\tests\test_normalize.py::test_normalize_document_writes_strict_change_report -v`
Expected: FAIL because the report writer does not exist yet.

- [ ] **Step 3: Implement report rows and CSV emission**

```python
import csv


def write_report(report_path, rows, fieldnames):
    with report_path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)
```

- [ ] **Step 4: Run tests to verify schema compliance**

Run: `pytest C:\Users\admin\Desktop\JJBand\tests\test_normalize.py -v`
Expected: PASS with assertions for:
- UTF-8 BOM compatible CSV output
- exact schema column order from `report_schema.csv`
- at least one `modified` row and one `unchanged` or `unresolved` row in fixture scenarios

- [ ] **Step 5: Commit**

```bash
git add C:\Users\admin\Desktop\JJBand\src\paper_format_normalizer\report.py C:\Users\admin\Desktop\JJBand\src\paper_format_normalizer\normalize.py C:\Users\admin\Desktop\JJBand\tests\test_normalize.py
git commit -m "feat: export strict normalization audit reports"
```

### Task 7: Wire The End-To-End CLI Command

**Files:**
- Modify: `C:\Users\admin\Desktop\JJBand\src\paper_format_normalizer\cli.py`
- Modify: `C:\Users\admin\Desktop\JJBand\tests\test_cli.py`

- [ ] **Step 1: Write the failing end-to-end CLI test**

```python
from pathlib import Path

from typer.testing import CliRunner

from paper_format_normalizer.cli import app


def test_cli_normalize_writes_docx_and_csv(tmp_path: Path, sample_docx_path: Path) -> None:
    runner = CliRunner()
    result = runner.invoke(
        app,
        [
            "normalize",
            "--input",
            str(sample_docx_path),
            "--rules",
            "tests/fixtures/sample_rules",
            "--output-dir",
            str(tmp_path),
        ],
    )
    assert result.exit_code == 0
    assert any(path.suffix == ".docx" for path in tmp_path.iterdir())
    assert any(path.suffix == ".csv" for path in tmp_path.iterdir())
```

- [ ] **Step 2: Run test to verify it fails**

Run: `pytest C:\Users\admin\Desktop\JJBand\tests\test_cli.py::test_cli_normalize_writes_docx_and_csv -v`
Expected: FAIL because the command does not accept arguments or run the pipeline yet.

- [ ] **Step 3: Implement CLI argument handling and pipeline orchestration**

```python
@app.command()
def normalize(
    input: Path = typer.Option(..., exists=True, dir_okay=False),
    rules: Path = typer.Option(..., exists=True, file_okay=False),
    output_dir: Path = typer.Option(..., file_okay=False),
) -> None:
    rule_set = load_rule_set(rules)
    output_path, report_path = normalize_document(input, rule_set, output_dir)
    typer.echo(f"DOCX={output_path}")
    typer.echo(f"REPORT={report_path}")
```

- [ ] **Step 4: Run CLI and full test suite**

Run: `pytest C:\Users\admin\Desktop\JJBand\tests -v`
Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add C:\Users\admin\Desktop\JJBand\src\paper_format_normalizer\cli.py C:\Users\admin\Desktop\JJBand\tests\test_cli.py
git commit -m "feat: run normalization pipeline from the cli"
```

### Task 8: Add Conversion Stubs And Explicit Rejection Paths For `.doc` And PDF

**Files:**
- Create: `C:\Users\admin\Desktop\JJBand\src\paper_format_normalizer\conversion.py`
- Modify: `C:\Users\admin\Desktop\JJBand\src\paper_format_normalizer\cli.py`
- Modify: `C:\Users\admin\Desktop\JJBand\tests\test_cli.py`

- [ ] **Step 1: Write failing rejection tests**

```python
import pytest

from paper_format_normalizer.conversion import ensure_working_docx


def test_ensure_working_docx_returns_native_docx_path(tmp_path) -> None:
    path = tmp_path / "paper.docx"
    path.write_bytes(b"stub")
    assert ensure_working_docx(path, tmp_path) == path


def test_ensure_working_docx_rejects_pdf_without_supported_conversion(tmp_path) -> None:
    path = tmp_path / "paper.pdf"
    path.write_bytes(b"%PDF-1.7")
    with pytest.raises(ValueError, match="PDF conversion is not configured"):
        ensure_working_docx(path, tmp_path)
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `pytest C:\Users\admin\Desktop\JJBand\tests\test_cli.py -v`
Expected: FAIL because conversion handling does not exist.

- [ ] **Step 3: Implement explicit conversion policy**

```python
from pathlib import Path


def ensure_working_docx(input_path: Path, work_dir: Path) -> Path:
    suffix = input_path.suffix.lower()
    if suffix == ".docx":
        return input_path
    if suffix == ".doc":
        raise ValueError("DOC conversion is not configured in phase 1")
    if suffix == ".pdf":
        raise ValueError("PDF conversion is not configured in phase 1")
    raise ValueError(f"Unsupported input type: {suffix}")
```

- [ ] **Step 4: Run tests to verify the phase-1 contract**

Run: `pytest C:\Users\admin\Desktop\JJBand\tests\test_cli.py -v`
Expected: PASS with explicit rejection coverage for `.doc` and `.pdf` until real converters are added.

- [ ] **Step 5: Commit**

```bash
git add C:\Users\admin\Desktop\JJBand\src\paper_format_normalizer\conversion.py C:\Users\admin\Desktop\JJBand\src\paper_format_normalizer\cli.py C:\Users\admin\Desktop\JJBand\tests\test_cli.py
git commit -m "feat: enforce explicit phase-one conversion boundaries"
```

### Task 9: Align The Skill Documentation With The Phase-1 Implementation

**Files:**
- Modify: `C:\Users\admin\Desktop\JJBand\skills\paper-format-normalizer\SKILL.md`

- [ ] **Step 1: Write the failing documentation checklist**

```text
Verify the skill states:
1. Phase 1 implements native `.docx`
2. `.doc` and `PDF` are planned but explicitly rejected until converters are added
3. CSV rule tables are required
4. unresolved objects are reported instead of silently fixed
```

- [ ] **Step 2: Run the checklist manually and confirm it fails**

Run: review `C:\Users\admin\Desktop\JJBand\skills\paper-format-normalizer\SKILL.md`
Expected: FAIL because the current skill text reads as broader support than the first shipped implementation.

- [ ] **Step 3: Update the skill wording to match shipping behavior**

```markdown
Supported in phase 1:
- `.docx`

Planned but not yet implemented:
- `.doc`
- PDF

If `.doc` or PDF is supplied before converters exist, fail explicitly.
```

- [ ] **Step 4: Re-run the checklist**

Run: review `C:\Users\admin\Desktop\JJBand\skills\paper-format-normalizer\SKILL.md`
Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add C:\Users\admin\Desktop\JJBand\skills\paper-format-normalizer\SKILL.md
git commit -m "docs: align skill contract with phase-one implementation"
```

### Task 10: Final Verification

**Files:**
- Modify as needed from prior tasks

- [ ] **Step 1: Run the full test suite**

Run: `pytest C:\Users\admin\Desktop\JJBand\tests -v`
Expected: PASS

- [ ] **Step 2: Run the CLI manually on a fixture**

Run: `python -m paper_format_normalizer.cli normalize --input C:\Users\admin\Desktop\JJBand\tests\fixtures\sample.docx --rules C:\Users\admin\Desktop\JJBand\tests\fixtures\sample_rules --output-dir C:\Users\admin\Desktop\JJBand\tmp\manual-run`
Expected: a new `*_规范化.docx` and `*_规范化_修改报告.csv` are written to `tmp\manual-run`

- [ ] **Step 3: Review outputs manually**

Run:
- inspect the generated CSV for `modified`, `unchanged`, and `unresolved`
- open the generated DOCX and verify title, abstract, heading, and body formatting changed as expected

Expected: output matches the phase-1 fixture rules with no overwrite of the input file.

- [ ] **Step 4: Commit the final green state**

```bash
git add C:\Users\admin\Desktop\JJBand
git commit -m "feat: ship phase-one paper format normalizer"
```

## Self-Review

### Spec Coverage

- Rule-driven classification is covered by Task 4.
- Strict validation and CSV reporting are covered by Tasks 2 and 6.
- Non-destructive DOCX output is covered by Tasks 5 and 7.
- Explicit phase-1 rejection for `.doc` and PDF is covered by Task 8.
- Skill documentation alignment is covered by Task 9.

### Placeholder Scan

- No `TODO`, `TBD`, or deferred pseudo-steps remain.
- Every coding task contains example code, concrete commands, and expected results.

### Type Consistency

- `RuleSet`, `parse_docx`, `classify_document`, `normalize_document`, and `ensure_working_docx` are introduced once and reused consistently.
