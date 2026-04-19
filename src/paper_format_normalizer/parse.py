from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Iterator

from docx import Document
from docx.oxml.ns import qn
from docx.table import Table
from docx.text.paragraph import Paragraph


@dataclass(frozen=True)
class ParsedBodyParagraph:
    text: str
    style_name: str | None


@dataclass(frozen=True)
class ParsedBodyTable:
    rows: tuple[tuple[str, ...], ...]


ParsedBodyItem = ParsedBodyParagraph | ParsedBodyTable


@dataclass(frozen=True)
class ParsedHeader:
    variant: str
    section_indices: tuple[int, ...]
    items: tuple[ParsedBodyItem, ...]


@dataclass(frozen=True)
class ParsedDocument:
    headers: tuple[ParsedHeader, ...]
    body_items: tuple[ParsedBodyItem, ...]
    paragraphs: tuple[ParsedBodyParagraph, ...]
    tables: tuple[ParsedBodyTable, ...]


@dataclass
class _HeaderRecord:
    variant: str
    section_indices: list[int]
    items: tuple[ParsedBodyItem, ...]


def parse_docx(path: Path) -> ParsedDocument:
    document = Document(path)
    headers = _parse_headers(document)
    body_items = _parse_body_items(document)
    paragraphs = tuple(
        item for item in body_items if isinstance(item, ParsedBodyParagraph)
    )
    tables = tuple(item for item in body_items if isinstance(item, ParsedBodyTable))
    return ParsedDocument(
        headers=headers,
        body_items=body_items,
        paragraphs=paragraphs,
        tables=tables,
    )


def _parse_headers(document) -> tuple[ParsedHeader, ...]:
    header_records: dict[int, _HeaderRecord] = {}
    header_order: list[int] = []
    for section_index, section in enumerate(document.sections):
        for variant, header in (
            ("default", section.header),
            ("first_page", section.first_page_header),
            ("even_page", section.even_page_header),
        ):
            items = _parse_content_items(header.iter_inner_content())
            if not items:
                continue
            part_key = id(header.part)
            record = header_records.get(part_key)
            if record is None:
                header_records[part_key] = _HeaderRecord(
                    variant=variant,
                    section_indices=[section_index],
                    items=items,
                )
                header_order.append(part_key)
                continue
            if section_index not in record.section_indices:
                record.section_indices.append(section_index)
    return tuple(
        ParsedHeader(
            variant=header_records[part_key].variant,
            section_indices=tuple(header_records[part_key].section_indices),
            items=header_records[part_key].items,
        )
        for part_key in header_order
    )


def _parse_body_items(document) -> tuple[ParsedBodyItem, ...]:
    return _parse_content_items(_iter_body_elements(document))


def _parse_content_items(
    elements: Iterator[Paragraph | Table],
) -> tuple[ParsedBodyItem, ...]:
    items: list[ParsedBodyItem] = []
    for element in elements:
        if isinstance(element, Paragraph):
            text = element.text.strip()
            if not text:
                continue
            items.append(
                ParsedBodyParagraph(
                    text=text,
                    style_name=element.style.name if element.style is not None else None,
                )
            )
            continue

        rows = tuple(
            tuple(cell.text for cell in row.cells)
            for row in element.rows
        )
        items.append(ParsedBodyTable(rows=rows))
    return tuple(items)


def _iter_body_elements(document) -> Iterator[Paragraph | Table]:
    for child in document.element.body.iterchildren():
        if child.tag == qn("w:p"):
            yield Paragraph(child, document)
        elif child.tag == qn("w:tbl"):
            yield Table(child, document)
