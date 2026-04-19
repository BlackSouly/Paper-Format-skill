from paper_format_normalizer.parse import ParsedBodyParagraph, ParsedBodyTable, parse_docx


def test_parse_docx_extracts_section_aware_headers_and_body_order(sample_docx_path) -> None:
    parsed = parse_docx(sample_docx_path)

    assert [
        (header.variant, header.section_indices, [item.text for item in header.items if isinstance(item, ParsedBodyParagraph)])
        for header in parsed.headers
    ] == [
        ("default", (0, 1), ["Shared header"]),
        ("first_page", (0, 1), ["Section 1 first-page header"]),
        ("even_page", (0, 1), ["Section 1 even-page header"]),
    ]

    assert [type(item) for item in parsed.body_items] == [
        ParsedBodyParagraph,
        ParsedBodyTable,
        ParsedBodyParagraph,
    ]
    assert [
        item.text if isinstance(item, ParsedBodyParagraph) else item.rows
        for item in parsed.body_items
    ] == [
        "First body paragraph",
        (("Cell 1", "Cell 2"), ("Cell 3", "Cell 4")),
        "Second body paragraph",
    ]

    assert [paragraph.text for paragraph in parsed.paragraphs] == [
        "First body paragraph",
        "Second body paragraph",
    ]
    assert [table.rows for table in parsed.tables] == [
        (("Cell 1", "Cell 2"), ("Cell 3", "Cell 4")),
    ]
