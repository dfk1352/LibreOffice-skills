# pyright: reportMissingImports=false, reportAttributeAccessIssue=false

from __future__ import annotations

import pytest


def test_parse_patch_empty_string_returns_empty_operations():
    from writer.patch import parse_patch

    assert parse_patch("") == []


def test_parse_patch_format_text_block_parses_structured_target_and_formatting():
    from writer.patch import parse_patch

    operations = parse_patch(
        "[operation]\n"
        "type = format_text\n"
        "target.kind = text\n"
        "target.text = quarterly revenue\n"
        "target.after = Financial Summary\n"
        "target.before = Risks\n"
        "target.occurrence = 0\n"
        "format.bold = true\n"
        "format.color = navy\n"
    )

    assert len(operations) == 1
    operation = operations[0]
    assert operation.operation_type == "format_text"
    assert operation.target.kind == "text"
    assert operation.target.text == "quarterly revenue"
    assert operation.target.after == "Financial Summary"
    assert operation.target.before == "Risks"
    assert operation.target.occurrence == 0
    assert operation.payload["formatting"].bold is True
    assert operation.payload["formatting"].color == "navy"


def test_parse_patch_insert_list_block_parses_multiline_json_items():
    from writer.patch import parse_patch

    operations = parse_patch(
        "[operation]\n"
        "type = insert_list\n"
        "target.kind = insertion\n"
        "target.after = Action Items\n"
        "list.ordered = false\n"
        "items <<JSON\n"
        "[\n"
        '  {"text": "Confirm scope", "level": 0},\n'
        '  {"text": "Review output", "level": 0},\n'
        '  {"text": "Update packaging", "level": 1}\n'
        "]\n"
        "JSON\n"
    )

    assert len(operations) == 1
    operation = operations[0]
    assert operation.operation_type == "insert_list"
    assert operation.target.kind == "insertion"
    assert operation.target.after == "Action Items"
    assert operation.payload["ordered"] is False
    assert [(item.text, item.level) for item in operation.payload["items"]] == [
        ("Confirm scope", 0),
        ("Review output", 0),
        ("Update packaging", 1),
    ]


def test_parse_patch_replace_text_supports_multiline_new_text():
    from writer.patch import parse_patch

    operations = parse_patch(
        "[operation]\n"
        "type = replace_text\n"
        "target.kind = text\n"
        "target.text = old paragraph\n"
        "new_text <<TEXT\n"
        "Updated line one.\n"
        "Updated line two.\n"
        "TEXT\n"
    )

    assert operations[0].operation_type == "replace_text"
    assert operations[0].target.text == "old paragraph"
    assert operations[0].payload["new_text"] == "Updated line one.\nUpdated line two."


def test_parse_patch_unknown_operation_type_raises_patch_syntax_error():
    from writer.exceptions import PatchSyntaxError
    from writer.patch import parse_patch

    with pytest.raises(PatchSyntaxError):
        parse_patch("[operation]\ntype = explode_document\ntext = nope\n")


def test_parse_patch_unterminated_heredoc_raises_patch_syntax_error():
    from writer.exceptions import PatchSyntaxError
    from writer.patch import parse_patch

    with pytest.raises(PatchSyntaxError):
        parse_patch(
            "[operation]\n"
            "type = replace_text\n"
            "target.kind = text\n"
            "target.text = old\n"
            "new_text <<TEXT\n"
            "unfinished payload\n"
        )


def test_parse_patch_bad_target_occurrence_integer_raises_patch_syntax_error():
    from writer.exceptions import PatchSyntaxError
    from writer.patch import parse_patch

    with pytest.raises(PatchSyntaxError):
        parse_patch(
            "[operation]\n"
            "type = delete_text\n"
            "target.kind = text\n"
            "target.text = old\n"
            "target.occurrence = first\n"
        )


@pytest.mark.parametrize("key", ["items", "data"])
def test_parse_patch_invalid_json_payload_raises_patch_syntax_error(key):
    from writer.exceptions import PatchSyntaxError
    from writer.patch import parse_patch

    payload_block = (
        "[operation]\n"
        f"type = {'insert_list' if key == 'items' else 'update_table'}\n"
        "target.kind = text\n"
        "target.text = anchor\n"
    )
    if key == "items":
        payload_block += "list.ordered = false\nitems = [invalid\n"
    else:
        payload_block += 'data = [["A", "B"]\n'

    with pytest.raises(PatchSyntaxError):
        parse_patch(payload_block)


# --- Numeric coercion tests (#2) ---


def test_parse_patch_font_size_coerced_to_float():
    """format.font_size = 14 must produce float(14.0), not the string '14' (#2)."""
    from writer.patch import parse_patch

    operations = parse_patch(
        "[operation]\n"
        "type = format_text\n"
        "target.kind = text\n"
        "target.text = hello\n"
        "format.font_size = 14\n"
    )

    formatting = operations[0].payload["formatting"]
    assert formatting.font_size == 14.0
    assert isinstance(formatting.font_size, float)


def test_parse_patch_line_spacing_coerced_to_float():
    """format.line_spacing = 1.5 must produce float(1.5) (#2)."""
    from writer.patch import parse_patch

    operations = parse_patch(
        "[operation]\n"
        "type = format_text\n"
        "target.kind = text\n"
        "target.text = hello\n"
        "format.line_spacing = 1.5\n"
    )

    formatting = operations[0].payload["formatting"]
    assert formatting.line_spacing == 1.5
    assert isinstance(formatting.line_spacing, float)


def test_parse_patch_spacing_before_coerced_to_int():
    """format.spacing_before = 200 must produce int(200) (#2)."""
    from writer.patch import parse_patch

    operations = parse_patch(
        "[operation]\n"
        "type = format_text\n"
        "target.kind = text\n"
        "target.text = hello\n"
        "format.spacing_before = 200\n"
    )

    formatting = operations[0].payload["formatting"]
    assert formatting.spacing_before == 200
    assert isinstance(formatting.spacing_before, int)


def test_parse_patch_spacing_after_coerced_to_int():
    """format.spacing_after = 100 must produce int(100) (#2)."""
    from writer.patch import parse_patch

    operations = parse_patch(
        "[operation]\n"
        "type = format_text\n"
        "target.kind = text\n"
        "target.text = hello\n"
        "format.spacing_after = 100\n"
    )

    formatting = operations[0].payload["formatting"]
    assert formatting.spacing_after == 100
    assert isinstance(formatting.spacing_after, int)


def test_parse_patch_invalid_font_size_raises():
    """format.font_size = not_a_number must raise PatchSyntaxError (#2)."""
    from writer.exceptions import PatchSyntaxError
    from writer.patch import parse_patch

    with pytest.raises(PatchSyntaxError):
        parse_patch(
            "[operation]\n"
            "type = format_text\n"
            "target.kind = text\n"
            "target.text = hello\n"
            "format.font_size = not_a_number\n"
        )
