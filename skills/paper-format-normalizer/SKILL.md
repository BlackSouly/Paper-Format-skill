---
name: paper-format-normalizer
description: Use when a user wants one paper or thesis document normalized to a school-provided formatting standard using a structured rule table, with strict validation and an exported change report.
---

# Paper Format Normalizer

## Overview

Normalize a single paper document against an explicit formatting rule table and export both a normalized DOCX and a strict change report.

This skill is rule-driven. It must not guess from appearance, use weak heuristics, or silently patch ambiguous cases.

## When To Use

- The user provides one document plus a structured formatting standard
- The task requires exact fonts, sizes, spacing, numbering, and page layout normalization
- The user wants a full audit trail of all changes

Do not use this skill for free-form document beautification or style improvements without an explicit rule table.

## Inputs

Required:

- One source document
- One complete rule-table set

Supported document types:

- `.docx` is supported natively in phase 1
- `.doc` is planned, but phase-1 CLI rejects it explicitly until a converter is added
- `PDF` is planned, but phase-1 CLI rejects it explicitly until a converter is added

Rule-table files:

- `templates/paper-format-rules/document_rules.csv`
- `templates/paper-format-rules/paragraph_rules.csv`
- `templates/paper-format-rules/numbering_rules.csv`
- `templates/paper-format-rules/table_rules.csv`
- `templates/paper-format-rules/special_object_rules.csv`

Report schema:

- `templates/paper-format-rules/report_schema.csv`

## Hard Constraints

1. Only classify objects from allowed evidence:
   - exact text
   - explicit regex or pattern
   - Word structure
   - numbering metadata
   - residual default body classification
2. Never infer object type from visual similarity alone.
3. Never overwrite the original document.
4. If an object cannot be classified deterministically, mark it `unresolved`.
5. If rule tables are missing required fields or contain conflicting rules at the same priority, stop and surface an explicit error.

## Workflow

### 1. Validate inputs

- Verify the document type is supported
- Verify all required CSV templates are present
- Verify required columns are populated

### 2. Convert to working DOCX

- Use native `.docx` directly
- Reject `.doc` with an explicit phase-1 conversion-not-configured error
- Reject `PDF` with an explicit phase-1 conversion-not-configured error
- Reject any unknown suffix explicitly

### 3. Parse document structure

Extract all relevant objects, including:

- sections
- paragraphs
- runs
- numbering
- tables and cells
- headers and footers
- footnotes or endnotes
- captions
- TOC fields
- text boxes
- equations when available

### 4. Classify objects

Apply evidence in priority order:

1. fixed semantic text matches
2. explicit patterns
3. Word object relationships
4. numbering rules
5. residual default body rule

### 5. Normalize

Reset properties from the matched rule:

- fonts
- font size
- bold, italic, underline
- alignment
- indent
- spacing
- page layout
- numbering
- table formatting
- special objects defined in the rule tables

### 6. Emit strict report

For each inspected property, write:

- object id
- object type before and after
- location
- preview
- property name
- before value
- after value
- rule id
- status
- reason

### 7. Save outputs

- `originalName_规范化.docx`
- `originalName_规范化_修改报告.csv`

## Failure Conditions

Stop and report clearly when:

- conversion fails
- `.doc` or `PDF` conversion is requested before converters are configured
- an unknown input suffix is provided
- rule tables are incomplete
- classification is ambiguous
- a required object type has no applicable rule

## Recommended Implementation Notes

- Prefer `python-docx` for supported DOCX operations
- Use deeper OOXML access for numbering, text boxes, and unsupported `python-docx` areas
- Treat `.doc` and `PDF` support as future converter-backed paths, not best-effort repair paths
- Keep strict rule-table validation and unresolved classification reporting enabled; phase 1 does not guess past missing or ambiguous rules

## Templates

Start from the CSV files in `templates/paper-format-rules/`.

These templates are examples and must be replaced with the user's real school requirements before normalization.
