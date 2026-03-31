"""Patch parsing and application for Calc spreadsheets."""

from __future__ import annotations

from typing import Any

import patch_base
from patch_base import (
    PatchApplyMode,
    PatchApplyResult,
    PatchOperation,
    PatchOperationResult,
    coerce_bool,
    coerce_float,
    coerce_int,
    coerce_json,
    require_payload_keys,
    require_target,
)
from calc.exceptions import PatchSyntaxError
from calc.targets import (
    CalcTarget,
    CellFormatting,
    ChartSpec,
    ValidationRule,
    parse_target,
)

_OPERATION_TYPES = {
    "write_cell",
    "write_range",
    "format_range",
    "add_sheet",
    "rename_sheet",
    "delete_sheet",
    "define_named_range",
    "delete_named_range",
    "set_validation",
    "clear_validation",
    "create_chart",
    "update_chart",
    "delete_chart",
    "recalculate",
}
_TARGET_INT_KEYS = {
    "sheet_index",
    "row",
    "col",
    "end_row",
    "end_col",
    "index",
}
_FORMAT_BOOL_KEYS = {"format.bold", "format.italic"}
_RULE_BOOL_KEYS = {
    "rule.show_error",
    "rule.show_input",
    "rule.ignore_blank",
}
_RULE_INT_KEYS = {"rule.error_style"}
_CHART_INT_KEYS = {
    "chart.anchor_row",
    "chart.anchor_col",
    "chart.width",
    "chart.height",
}
_CHART_RANGE_INT_KEYS = {
    "row",
    "col",
    "end_row",
    "end_col",
    "sheet_index",
    "index",
}

_E = PatchSyntaxError


def parse_patch(patch_text: str) -> list[PatchOperation]:
    """Parse Calc patch text into ordered operations."""
    return patch_base.parse_patch(
        patch_text,
        operation_types=_OPERATION_TYPES,
        target_int_keys=_TARGET_INT_KEYS,
        parse_target_fn=parse_target,
        parse_payload_fn=_parse_payload,
        validate_fn=_validate_operation_shape,
        error_cls=_E,
    )


def apply_operations(
    session: Any, patch_text: str, mode: PatchApplyMode
) -> PatchApplyResult:
    """Apply patch operations to an already-open Calc session."""
    return patch_base.apply_operations(
        session,
        patch_text,
        mode,
        parse_patch_fn=parse_patch,
        dispatch_fn=_dispatch_operation,
        error_cls=_E,
    )


def patch(
    path: str, patch_text: str, mode: PatchApplyMode = "atomic"
) -> PatchApplyResult:
    """Open a session, apply a patch, and persist if appropriate."""
    from calc.session import CalcSession

    return patch_base.patch(
        path,
        patch_text,
        mode,
        session_cls=CalcSession,
        apply_fn=apply_operations,
    )


# ---- App-specific helpers ------------------------------------------------


def _dispatch_operation(session: Any, operation: PatchOperation) -> None:
    if operation.operation_type == "write_cell":
        session.write_cell(
            operation.target,
            operation.payload["value"],
            value_type=operation.payload.get("value_type", "auto"),
        )
        return
    if operation.operation_type == "write_range":
        session.write_range(operation.target, operation.payload["data"])
        return
    if operation.operation_type == "format_range":
        session.format_range(operation.target, operation.payload["formatting"])
        return
    if operation.operation_type == "add_sheet":
        session.add_sheet(
            operation.payload["name"],
            index=operation.payload.get("index"),
        )
        return
    if operation.operation_type == "rename_sheet":
        session.rename_sheet(operation.target, operation.payload["new_name"])
        return
    if operation.operation_type == "delete_sheet":
        session.delete_sheet(operation.target)
        return
    if operation.operation_type == "define_named_range":
        session.define_named_range(operation.payload["name"], operation.target)
        return
    if operation.operation_type == "delete_named_range":
        session.delete_named_range(operation.target)
        return
    if operation.operation_type == "set_validation":
        session.set_validation(operation.target, operation.payload["rule"])
        return
    if operation.operation_type == "clear_validation":
        session.clear_validation(operation.target)
        return
    if operation.operation_type == "create_chart":
        session.create_chart(operation.target, operation.payload["spec"])
        return
    if operation.operation_type == "update_chart":
        session.update_chart(operation.target, operation.payload["spec"])
        return
    if operation.operation_type == "delete_chart":
        session.delete_chart(operation.target)
        return
    if operation.operation_type == "recalculate":
        session.recalculate()
        return
    raise _E(f"Unsupported operation type: {operation.operation_type}")


def _parse_payload(operation_type: str, block: dict[str, str]) -> dict[str, Any]:
    payload: dict[str, Any] = {}
    for key, raw_value in block.items():
        if key == "type" or key.startswith("target."):
            continue
        if key == "data":
            payload[key] = coerce_json(raw_value, key, _E)
            continue
        if key == "index":
            payload[key] = coerce_int(raw_value, key, _E)
            continue
        if key.startswith("format."):
            payload[key] = _coerce_format_value(key, raw_value)
            continue
        if key.startswith("rule."):
            payload[key] = _coerce_rule_value(key, raw_value)
            continue
        if key.startswith("chart."):
            payload[key] = _coerce_chart_value(key, raw_value)
            continue
        if key == "value":
            payload[key] = _coerce_scalar(raw_value)
            continue
        payload[key] = raw_value

    if operation_type == "format_range":
        payload = {"formatting": _build_formatting(payload)}
    elif operation_type == "set_validation":
        payload = {"rule": _build_validation_rule(payload)}
    elif operation_type in {"create_chart", "update_chart"}:
        payload = {"spec": _build_chart_spec(payload)}
    return payload


def _validate_operation_shape(
    operation_type: str,
    target: CalcTarget | None,
    payload: dict[str, Any],
) -> None:
    has_target = target is not None
    if operation_type == "write_cell":
        require_target(operation_type, has_target, _E)
        require_payload_keys(operation_type, payload, {"value"}, _E)
        return
    if operation_type == "write_range":
        require_target(operation_type, has_target, _E)
        require_payload_keys(operation_type, payload, {"data"}, _E)
        return
    if operation_type == "format_range":
        require_target(operation_type, has_target, _E)
        require_payload_keys(operation_type, payload, {"formatting"}, _E)
        return
    if operation_type == "add_sheet":
        require_payload_keys(operation_type, payload, {"name"}, _E)
        return
    if operation_type == "rename_sheet":
        require_target(operation_type, has_target, _E)
        require_payload_keys(operation_type, payload, {"new_name"}, _E)
        return
    if operation_type in {
        "delete_sheet",
        "delete_named_range",
        "clear_validation",
        "delete_chart",
    }:
        require_target(operation_type, has_target, _E)
        return
    if operation_type == "define_named_range":
        require_target(operation_type, has_target, _E)
        require_payload_keys(operation_type, payload, {"name"}, _E)
        return
    if operation_type == "set_validation":
        require_target(operation_type, has_target, _E)
        require_payload_keys(operation_type, payload, {"rule"}, _E)
        return
    if operation_type in {"create_chart", "update_chart"}:
        require_target(operation_type, has_target, _E)
        require_payload_keys(operation_type, payload, {"spec"}, _E)
        return


def _build_formatting(payload: dict[str, Any]) -> CellFormatting:
    return CellFormatting(
        bold=payload.get("format.bold"),
        italic=payload.get("format.italic"),
        font_name=payload.get("format.font_name"),
        font_size=payload.get("format.font_size"),
        color=payload.get("format.color"),
        number_format=payload.get("format.number_format"),
    )


def _build_validation_rule(payload: dict[str, Any]) -> ValidationRule:
    rule_type = payload.get("rule.type")
    condition = payload.get("rule.condition")
    if rule_type is None or condition is None:
        missing = []
        if rule_type is None:
            missing.append("rule.type")
        if condition is None:
            missing.append("rule.condition")
        raise _E(
            f"Operation set_validation is missing required keys: {', '.join(missing)}"
        )
    return ValidationRule(
        type=str(rule_type),
        condition=str(condition),
        value1=payload.get("rule.value1"),
        value2=payload.get("rule.value2"),
        show_error=payload.get("rule.show_error", False),
        error_message=str(payload.get("rule.error_message", "")),
        show_input=payload.get("rule.show_input", False),
        input_title=str(payload.get("rule.input_title", "")),
        input_message=str(payload.get("rule.input_message", "")),
        ignore_blank=payload.get("rule.ignore_blank", True),
        error_style=payload.get("rule.error_style", 0),
    )


def _build_chart_spec(payload: dict[str, Any]) -> ChartSpec:
    data_range_fields = {
        key.split(".", 2)[2]: value
        for key, value in payload.items()
        if key.startswith("chart.data_range.")
    }
    if not data_range_fields:
        raise _E("Chart operations require chart.data_range.* fields")

    return ChartSpec(
        chart_type=str(payload.get("chart.chart_type", "")),
        data_range=parse_target(data_range_fields),
        anchor_row=_require_int_payload(payload, "chart.anchor_row"),
        anchor_col=_require_int_payload(payload, "chart.anchor_col"),
        width=_require_int_payload(payload, "chart.width"),
        height=_require_int_payload(payload, "chart.height"),
        title=_optional_string(payload.get("chart.title")),
    )


def _coerce_format_value(key: str, value: str) -> Any:
    if key in _FORMAT_BOOL_KEYS:
        return coerce_bool(value, key, _E)
    if key == "format.font_size":
        return coerce_float(value, key, _E)
    return value


def _coerce_rule_value(key: str, value: str) -> Any:
    if key in _RULE_BOOL_KEYS:
        return coerce_bool(value, key, _E)
    if key in _RULE_INT_KEYS:
        return coerce_int(value, key, _E)
    return _coerce_scalar(value)


def _coerce_chart_value(key: str, value: str) -> Any:
    if key in _CHART_INT_KEYS:
        return coerce_int(value, key, _E)
    if key.startswith("chart.data_range."):
        field = key.split(".", 2)[2]
        if field in _CHART_RANGE_INT_KEYS:
            return coerce_int(value, key, _E)
    return value


def _coerce_scalar(value: str) -> Any:
    lowered = value.strip().lower()
    if lowered == "true":
        return True
    if lowered == "false":
        return False
    try:
        return int(value)
    except ValueError:
        pass
    try:
        return float(value)
    except ValueError:
        return value


def _require_int_payload(payload: dict[str, Any], key: str) -> int:
    value = payload.get(key)
    if not isinstance(value, int):
        raise _E(f"{key} must be an integer")
    return value


def _optional_string(value: Any) -> str | None:
    if value is None:
        return None
    cleaned = str(value).strip()
    return cleaned or None
