# pyright: reportMissingImports=false, reportAttributeAccessIssue=false

from __future__ import annotations

import pytest


def test_parse_patch_empty_string_returns_empty_operations():
    from impress.patch import parse_patch

    assert parse_patch("") == []


def test_parse_patch_replace_text_block_parses_multiline_new_text():
    from impress.patch import parse_patch

    operations = parse_patch(
        "[operation]\n"
        "type = replace_text\n"
        "target.kind = text\n"
        "target.slide_index = 2\n"
        "target.placeholder = body\n"
        "new_text <<EOF\n"
        "Quarterly revenue rose 21%.\n"
        "Keep the executive summary concise.\n"
        "EOF\n"
    )

    assert len(operations) == 1
    operation = operations[0]
    assert operation.operation_type == "replace_text"
    assert operation.target.kind == "text"
    assert operation.target.slide_index == 2
    assert operation.target.placeholder == "body"
    assert operation.payload["new_text"] == (
        "Quarterly revenue rose 21%.\nKeep the executive summary concise."
    )


def test_parse_patch_insert_list_block_parses_multiline_json_items():
    from impress.patch import parse_patch

    operations = parse_patch(
        "[operation]\n"
        "type = insert_list\n"
        "target.kind = insertion\n"
        "target.slide_index = 2\n"
        "target.shape_name = Agenda Box\n"
        "target.after = Action Items\n"
        "list.ordered = true\n"
        "items <<JSON\n"
        "[\n"
        '  {"text": "Confirm scope", "level": 0},\n'
        '  {"text": "Review outputs", "level": 0},\n'
        '  {"text": "Update notes", "level": 1}\n'
        "]\n"
        "JSON\n"
    )

    operation = operations[0]
    assert operation.operation_type == "insert_list"
    assert operation.payload["ordered"] is True
    assert [item.text for item in operation.payload["items"]] == [
        "Confirm scope",
        "Review outputs",
        "Update notes",
    ]
    assert [item.level for item in operation.payload["items"]] == [0, 0, 1]


def test_parse_patch_placement_fields_parse_into_shape_placement():
    from impress import ShapePlacement
    from impress.patch import parse_patch

    operations = parse_patch(
        "[operation]\n"
        "type = insert_text_box\n"
        "target.kind = slide\n"
        "target.slide_index = 1\n"
        "text = Summary\n"
        "placement.x_cm = 1.5\n"
        "placement.y_cm = 2.0\n"
        "placement.width_cm = 10.0\n"
        "placement.height_cm = 3.5\n"
        "name = Summary Box\n"
    )

    placement = operations[0].payload["placement"]
    assert isinstance(placement, ShapePlacement)
    assert placement.x_cm == 1.5
    assert placement.y_cm == 2.0
    assert placement.width_cm == 10.0
    assert placement.height_cm == 3.5


def test_parse_patch_master_fields_parse_for_apply_master_page():
    from impress.patch import parse_patch

    operations = parse_patch(
        "[operation]\n"
        "type = apply_master_page\n"
        "target.kind = slide\n"
        "target.slide_index = 1\n"
        "master.kind = master_page\n"
        "master.master_name = Corporate Blue\n"
    )

    operation = operations[0]
    assert operation.operation_type == "apply_master_page"
    assert operation.target.kind == "slide"
    assert operation.payload["master_target"].kind == "master_page"
    assert operation.payload["master_target"].master_name == "Corporate Blue"


def test_parse_patch_unknown_operation_type_raises_patch_syntax_error():
    from impress.exceptions import PatchSyntaxError
    from impress.patch import parse_patch

    with pytest.raises(PatchSyntaxError):
        parse_patch("[operation]\ntype = explode_slide\n")


@pytest.mark.parametrize(
    "line",
    [
        "target.slide_index = second",
        "target.occurrence = later",
    ],
)
def test_parse_patch_bad_integer_coercion_raises_patch_syntax_error(line):
    from impress.exceptions import PatchSyntaxError
    from impress.patch import parse_patch

    patch_text = (
        "[operation]\n"
        "type = replace_text\n"
        "target.kind = text\n"
        f"{line}\n"
        "target.shape_name = Copy Box\n"
        "new_text = Updated\n"
    )

    with pytest.raises(PatchSyntaxError):
        parse_patch(patch_text)


def test_parse_patch_invalid_json_items_raises_patch_syntax_error():
    from impress.exceptions import PatchSyntaxError
    from impress.patch import parse_patch

    with pytest.raises(PatchSyntaxError):
        parse_patch(
            "[operation]\n"
            "type = insert_list\n"
            "target.kind = insertion\n"
            "target.slide_index = 1\n"
            "list.ordered = false\n"
            'items = [{"text": "Broken"}\n'
        )


def test_parse_patch_invalid_json_data_raises_patch_syntax_error():
    from impress.exceptions import PatchSyntaxError
    from impress.patch import parse_patch

    with pytest.raises(PatchSyntaxError):
        parse_patch(
            "[operation]\n"
            "type = insert_chart\n"
            "target.kind = slide\n"
            "target.slide_index = 1\n"
            "chart_type = bar\n"
            "placement.x_cm = 1.0\n"
            "placement.y_cm = 1.0\n"
            "placement.width_cm = 8.0\n"
            "placement.height_cm = 5.0\n"
            'data = [["Category", "Value"]\n'
        )


def test_parse_patch_unterminated_heredoc_raises_patch_syntax_error():
    from impress.exceptions import PatchSyntaxError
    from impress.patch import parse_patch

    with pytest.raises(PatchSyntaxError):
        parse_patch(
            "[operation]\n"
            "type = replace_text\n"
            "target.kind = text\n"
            "target.slide_index = 0\n"
            "new_text <<EOF\n"
            "Never closed\n"
        )


def test_parse_patch_rejects_boolean_list_level():
    from impress.exceptions import PatchSyntaxError
    from impress.patch import parse_patch

    with pytest.raises(PatchSyntaxError, match="List item level must be an integer"):
        parse_patch(
            "[operation]\n"
            "type = insert_list\n"
            "target.kind = insertion\n"
            "target.slide_index = 1\n"
            "list.ordered = false\n"
            "items <<JSON\n"
            '[{"text": "Item", "level": true}]\n'
            "JSON\n"
        )
