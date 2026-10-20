"""Ingest: .docx -> Document.

Walks the document body element in order so that tables and paragraphs
keep their original positions. This is what makes the preview render
the document in the same order Word shows it, rather than dumping all
paragraphs first and all tables last.
"""
from __future__ import annotations

import os

import docx
from docx.oxml.ns import qn

from .helpers import (
    cell_text,
    heading_level_from_style,
    make_spans_for_text,
    paragraph_is_list_item,
    paragraph_text,
    paragraph_text_boxes,
)
from .model import Block, BlockType, Document, Table


def ingest(path: str) -> Document:
    doc = docx.Document(path)
    document = Document(filename=os.path.basename(path))

    body = doc.element.body
    para_iter = iter(doc.paragraphs)
    table_iter = iter(doc.tables)

    block_index = 0
    table_index = 0

    for child in body.iterchildren():
        if child.tag == qn("w:p"):
            try:
                para = next(para_iter)
            except StopIteration:
                continue
            block_index = _ingest_paragraph(para, document, block_index)
        elif child.tag == qn("w:tbl"):
            try:
                raw_table = next(table_iter)
            except StopIteration:
                continue
            table = _ingest_table(raw_table, table_index, block_index)
            document.tables.append(table)
            table_index += 1

    _count_drawings(doc, document)
    return document


def _ingest_paragraph(para, document: Document, block_index: int) -> int:
    text, math_count = paragraph_text(para)
    if text.strip() or math_count:
        style_name = (para.style.name or "") if para.style else ""
        level = heading_level_from_style(style_name)
        if level:
            block_type = BlockType.HEADING
        elif paragraph_is_list_item(para):
            block_type = BlockType.LIST_ITEM
        else:
            block_type = BlockType.PARAGRAPH

        block = Block(
            block_type=block_type,
            block_index=block_index,
            style_name=style_name,
        )
        block.raw_text = text
        block.spans = make_spans_for_text(text, block_index=block_index)
        block.math_object_count = math_count
        document.blocks.append(block)
        block_index += 1

    for tb_text, tb_math in paragraph_text_boxes(para):
        if not tb_text.strip() and not tb_math:
            continue
        block = Block(
            block_type=BlockType.TEXT_BOX,
            block_index=block_index,
        )
        block.raw_text = tb_text
        block.spans = make_spans_for_text(tb_text, block_index=block_index)
        block.math_object_count = tb_math
        document.blocks.append(block)
        document.text_boxes.append(block)
        block_index += 1

    return block_index


def _ingest_table(raw_table, t_idx: int, position: int) -> Table:
    table = Table(table_index=t_idx, block_index=position)

    for r_idx, raw_row in enumerate(raw_table.rows):
        row_blocks = []
        seen_tc_in_row: set[int] = set()
        for c_idx, raw_cell in enumerate(raw_row.cells):
            tc_id = id(raw_cell._tc)
            synthetic_index = 100_000 + t_idx * 1000 + r_idx * 100 + c_idx

            if tc_id in seen_tc_in_row:
                cell_block = Block(
                    block_type=BlockType.TABLE_CELL,
                    block_index=synthetic_index,
                )
                cell_block.raw_text = ""
                cell_block.spans = []
                cell_block.math_object_count = 0
                row_blocks.append(cell_block)
                continue
            seen_tc_in_row.add(tc_id)

            text, math_count = cell_text(raw_cell)
            cell_block = Block(
                block_type=BlockType.TABLE_CELL,
                block_index=synthetic_index,
            )
            cell_block.raw_text = text
            cell_block.spans = make_spans_for_text(text, block_index=synthetic_index)
            cell_block.math_object_count = math_count
            row_blocks.append(cell_block)
        table.rows.append(row_blocks)

    table.header_row_index = _detect_header_row(raw_table)
    return table


def _detect_header_row(raw_table) -> int:
    try:
        for r_idx, row in enumerate(raw_table.rows):
            trPr = row._tr.find(qn("w:trPr"))
            if trPr is not None:
                if trPr.find(qn("w:tblHeader")) is not None:
                    return r_idx
    except Exception:
        pass
    return 0


def _count_drawings(doc, document: Document) -> None:
    for node in doc.element.body.iter():
        if node.tag == qn("w:drawing"):
            document.drawings += 1
            for inner in node.iter():
                if inner.tag == qn("a:t") and (inner.text or "").strip():
                    document.drawings_with_text += 1
                    break