"""Tests for the Writer patch DSL parser."""

# pyright: reportMissingImports=false, reportAttributeAccessIssue=false

from __future__ import annotations

import pytest


def test_parse_patch_empty_string_returns_empty_operations():
    from writer.patch import parse_patch

    assert parse_patch("") == []


def test_parse_patch_single_insert_text_block():
    from writer.patch import parse_patch

    operations = parse_patch(
        "[operation]\ntype = insert_text\ntext = Hello from patch\n"
    )

    assert len(operations) == 1
    operation = operations[0]
    assert operation.operation_type == "insert_text"
    assert operation.selector is None
    assert operation.payload["text"] == "Hello from patch"


def test_parse_patch_multiple_blocks_preserve_order():
    from writer.patch import parse_patch

    operations = parse_patch(
        "[operation]\n"
        "type = insert_text\n"
        "text = First\n\n"
        "[operation]\n"
        "type = delete_table\n"
        'selector = name:"Quarterly Results"\n'
    )

    assert [operation.operation_type for operation in operations] == [
        "insert_text",
        "delete_table",
    ]


def test_parse_patch_missing_required_key_raises_patch_syntax_error():
    from writer.exceptions import PatchSyntaxError
    from writer.patch import parse_patch

    with pytest.raises(PatchSyntaxError):
        parse_patch('[operation]\ntype = replace_text\nselector = contains:"old"\n')


def test_parse_patch_unknown_operation_type_raises_patch_syntax_error():
    from writer.exceptions import PatchSyntaxError
    from writer.patch import parse_patch

    with pytest.raises(PatchSyntaxError):
        parse_patch("[operation]\ntype = explode_document\ntext = nope\n")


def test_parse_patch_ignores_comments_and_blank_lines():
    from writer.patch import parse_patch

    operations = parse_patch(
        "# comment before block\n\n"
        "[operation]\n"
        "type = insert_text\n"
        "text = First\n\n"
        "# another comment\n"
        "[operation]\n"
        "type = delete_text\n"
        'selector = contains:"First"\n'
    )

    assert len(operations) == 2
    assert operations[0].operation_type == "insert_text"
    assert operations[1].operation_type == "delete_text"


def test_parse_patch_parses_table_data_json():
    from writer.patch import parse_patch

    operations = parse_patch(
        "[operation]\n"
        "type = update_table\n"
        'selector = name:"Budget Table"\n'
        'data = [["A", "B"], ["1", "2"]]\n'
    )

    assert operations[0].payload["data"] == [["A", "B"], ["1", "2"]]


def test_parse_patch_malformed_table_data_json_raises_patch_syntax_error():
    from writer.exceptions import PatchSyntaxError
    from writer.patch import parse_patch

    with pytest.raises(PatchSyntaxError):
        parse_patch(
            "[operation]\n"
            "type = update_table\n"
            'selector = name:"Budget Table"\n'
            'data = [["A", "B"]\n'
        )
