"""Tests for the Calc patch DSL parser."""

# pyright: reportMissingImports=false, reportAttributeAccessIssue=false

from __future__ import annotations

import pytest


def test_parse_patch_empty_string_returns_empty_operations():
    from calc.patch import parse_patch

    assert parse_patch("") == []


def test_parse_patch_write_range_block_parses_multiline_json_data():
    from calc.patch import parse_patch

    operations = parse_patch(
        "[operation]\n"
        "type = write_range\n"
        "target.kind = range\n"
        "target.sheet = Data\n"
        "target.row = 0\n"
        "target.col = 0\n"
        "target.end_row = 2\n"
        "target.end_col = 1\n"
        "data <<JSON\n"
        '[["Label", "Value"], ["Revenue", 100], ["Cost", 80]]\n'
        "JSON\n"
    )

    assert len(operations) == 1
    operation = operations[0]
    assert operation.operation_type == "write_range"
    assert operation.target.kind == "range"
    assert operation.target.sheet == "Data"
    assert operation.payload["data"] == [
        ["Label", "Value"],
        ["Revenue", 100],
        ["Cost", 80],
    ]


def test_parse_patch_format_range_block_parses_cell_formatting():
    from calc import CellFormatting
    from calc.patch import parse_patch

    operations = parse_patch(
        "[operation]\n"
        "type = format_range\n"
        "target.kind = range\n"
        "target.sheet = Data\n"
        "target.row = 1\n"
        "target.col = 1\n"
        "target.end_row = 2\n"
        "target.end_col = 1\n"
        "format.bold = true\n"
        "format.number_format = currency\n"
    )

    formatting = operations[0].payload["formatting"]
    assert isinstance(formatting, CellFormatting)
    assert formatting.bold is True
    assert formatting.number_format == "currency"


def test_parse_patch_set_validation_block_parses_validation_rule():
    from calc import ValidationRule
    from calc.patch import parse_patch

    operations = parse_patch(
        "[operation]\n"
        "type = set_validation\n"
        "target.kind = range\n"
        "target.sheet = Data\n"
        "target.row = 1\n"
        "target.col = 1\n"
        "target.end_row = 2\n"
        "target.end_col = 1\n"
        "rule.type = whole\n"
        "rule.condition = between\n"
        "rule.value1 = 1\n"
        "rule.value2 = 1000\n"
        "rule.show_error = true\n"
        "rule.error_message = Value must be positive\n"
    )

    rule = operations[0].payload["rule"]
    assert isinstance(rule, ValidationRule)
    assert rule.type == "whole"
    assert rule.condition == "between"
    assert rule.value1 == 1
    assert rule.value2 == 1000
    assert rule.show_error is True
    assert rule.error_message == "Value must be positive"


def test_parse_patch_create_chart_block_parses_chart_spec():
    from calc import ChartSpec
    from calc.patch import parse_patch

    operations = parse_patch(
        "[operation]\n"
        "type = create_chart\n"
        "target.kind = sheet\n"
        "target.sheet = Dashboard\n"
        "chart.chart_type = line\n"
        "chart.data_range.kind = range\n"
        "chart.data_range.sheet = Dashboard\n"
        "chart.data_range.row = 0\n"
        "chart.data_range.col = 0\n"
        "chart.data_range.end_row = 2\n"
        "chart.data_range.end_col = 1\n"
        "chart.anchor_row = 5\n"
        "chart.anchor_col = 0\n"
        "chart.width = 5000\n"
        "chart.height = 3000\n"
        "chart.title = Revenue Trend\n"
    )

    spec = operations[0].payload["spec"]
    assert isinstance(spec, ChartSpec)
    assert spec.chart_type == "line"
    assert spec.data_range.kind == "range"
    assert spec.data_range.sheet == "Dashboard"
    assert spec.anchor_row == 5
    assert spec.anchor_col == 0
    assert spec.width == 5000
    assert spec.height == 3000
    assert spec.title == "Revenue Trend"


def test_parse_patch_unknown_operation_type_raises_patch_syntax_error():
    from calc.exceptions import PatchSyntaxError
    from calc.patch import parse_patch

    with pytest.raises(PatchSyntaxError):
        parse_patch("[operation]\ntype = explode_sheet\n")


@pytest.mark.parametrize(
    "line",
    [
        "target.row = first",
        "chart.width = wide",
    ],
)
def test_parse_patch_bad_integer_coercion_raises_patch_syntax_error(line):
    from calc.exceptions import PatchSyntaxError
    from calc.patch import parse_patch

    patch_text = (
        "[operation]\n"
        "type = create_chart\n"
        "target.kind = sheet\n"
        "target.sheet = Dashboard\n"
        "chart.chart_type = line\n"
        "chart.data_range.kind = range\n"
        "chart.data_range.sheet = Dashboard\n"
        "chart.data_range.row = 0\n"
        "chart.data_range.col = 0\n"
        "chart.data_range.end_row = 2\n"
        "chart.data_range.end_col = 1\n"
        "chart.anchor_row = 5\n"
        "chart.anchor_col = 0\n"
        f"{line}\n"
        "chart.height = 3000\n"
    )

    with pytest.raises(PatchSyntaxError):
        parse_patch(patch_text)


def test_parse_patch_invalid_json_data_raises_patch_syntax_error():
    from calc.exceptions import PatchSyntaxError
    from calc.patch import parse_patch

    with pytest.raises(PatchSyntaxError):
        parse_patch(
            "[operation]\n"
            "type = write_range\n"
            "target.kind = range\n"
            "target.sheet = Data\n"
            "target.row = 0\n"
            "target.col = 0\n"
            "target.end_row = 0\n"
            "target.end_col = 1\n"
            'data = [["Label", 1]\n'
        )


def test_parse_patch_unterminated_heredoc_raises_patch_syntax_error():
    from calc.exceptions import PatchSyntaxError
    from calc.patch import parse_patch

    with pytest.raises(PatchSyntaxError):
        parse_patch(
            "[operation]\n"
            "type = write_range\n"
            "target.kind = range\n"
            "target.sheet = Data\n"
            "target.row = 0\n"
            "target.col = 0\n"
            "target.end_row = 1\n"
            "target.end_col = 1\n"
            "data <<JSON\n"
            "[[1, 2], [3, 4]]\n"
        )
