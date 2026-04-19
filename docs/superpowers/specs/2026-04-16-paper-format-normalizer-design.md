# Paper Format Normalizer Design

## Goal

Create a skill-driven document normalizer that accepts one paper-format rule table plus one source document and produces:

- `xxx_规范化.docx`
- `xxx_规范化_修改报告.csv`

The system must avoid degradation handling, fallback hacks, weak heuristics, and post-processing bandages. It should only make decisions based on explicit rules, Word structure, and deterministic text patterns.

## Supported Inputs

- Native `.docx`
- Legacy `.doc`, converted to `.docx` before normalization
- PDF only when text extraction and DOCX conversion preserve usable structure

Unsupported PDFs include scanned-image PDFs and files whose conversion destroys paragraph structure. Those must be rejected as not meeting the preconditions for reliable normalization.

## Primary Interaction Model

Version 1 supports a single-document workflow.

Inputs:

- One structured rule table set
- One document

Outputs:

- One normalized DOCX file
- One strict change report

## Design Principles

1. Rule-driven only. All transformations must be justified by the rule tables.
2. Structural evidence first. Use Word object structure, numbering, fields, and exact text matches before defaulting to body text.
3. No visual guessing. Do not classify content from appearance alone.
4. Strict auditability. Every changed property must be logged as `before -> after`.
5. No silent repair. If an object cannot be classified from allowed evidence, mark it `unresolved` and report it.

## Object Classification Model

Objects are classified using the following evidence sources only:

1. Exact text match
2. Explicit text pattern
3. Word structural relationships
4. Numbering metadata
5. Residual default classification after stronger rules are exhausted

Priority order:

1. Fixed semantic sections such as `摘要`, `关键词`, `参考文献`
2. Heading levels determined by numbering patterns and structure
3. Captions, footnotes, table content, headers, footers, TOC, formulas, text boxes
4. Body paragraphs as the residual class

## Rule Table Layout

The skill should consume five CSV tables:

1. `document_rules.csv`
2. `paragraph_rules.csv`
3. `numbering_rules.csv`
4. `table_rules.csv`
5. `special_object_rules.csv`

An additional report schema file defines the strict report columns.

## Processing Pipeline

### 1. Ingestion

- Detect source file type
- Convert `.doc` and eligible PDF to intermediate `.docx`
- Record conversion metadata

### 2. Structural Parse

Extract:

- Sections
- Paragraphs
- Runs
- Styles
- Numbering
- Tables and cells
- Headers and footers
- Footnotes and endnotes
- Images and captions
- Text boxes and shape text
- TOC fields
- Equations when available

### 3. Classification

For each object:

- Apply fixed semantic matches
- Apply explicit regex or text-pattern rules
- Apply Word-structure rules
- Apply numbering rules
- If no stronger rule matches, classify as default body text where allowed
- Otherwise mark unresolved

### 4. Normalization

Apply the matched rule by fully resetting the target properties instead of attempting minimal patching.

Examples:

- Reset Chinese font, Western font, size, bold, italic, underline
- Reset alignment, indent, line spacing, spacing before and after
- Reset margins, paper size, page numbering
- Rebuild numbering when required by rule
- Normalize tables, captions, headers, footers, footnotes, and text boxes

### 5. Strict Validation and Audit

For each object and property:

- Capture original value
- Capture normalized value
- Record matched rule id
- Record object location and preview
- Record status as `modified`, `unchanged`, or `unresolved`

### 6. Output

- Save normalized document as `originalName_规范化.docx`
- Save report as `originalName_规范化_修改报告.csv`

## Output Contract

### Normalized DOCX

- Never overwrite the original file
- Preserve document text content except where formatting normalization requires object conversion

### Strict Change Report

Each row represents one inspected property on one object, with enough context for manual review.

Required report columns:

- `object_id`
- `object_type_before`
- `object_type_after`
- `location`
- `text_preview`
- `property`
- `before`
- `after`
- `rule_id`
- `status`
- `reason`

## Rejection Conditions

The workflow must stop with an explicit error when:

- The source file cannot be converted into a structurally usable DOCX
- Required rule tables are missing
- Required fields in the rule tables are blank or malformed
- Multiple conflicting rules match the same object at the same priority without a deterministic tie-break

## Open Implementation Notes

- `.doc` conversion likely requires LibreOffice or Word automation
- PDF conversion must be treated as a precondition check, not a best-effort cleanup path
- Some objects such as equations and legacy text boxes may require OOXML-level access beyond `python-docx`

## First Implementation Scope

Phase 1 should prioritize:

1. Native `.docx`
2. Rule loading and validation
3. Paragraph, heading, table, caption, footnote, header, footer handling
4. Strict report generation

Later phases can extend `.doc` conversion, PDF intake, equations, and more complex OOXML objects.

## Review Checklist

- The system remains rule-driven and auditable
- Unsupported inputs fail explicitly
- No weak visual heuristics are introduced
- Report granularity is strict rather than summary-only
- Outputs use non-destructive filenames
