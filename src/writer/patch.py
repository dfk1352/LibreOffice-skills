"""Patch parsing and application for Writer documents."""

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
from writer.exceptions import PatchSyntaxError
from writer.targets import ListItem, TextFormatting, WriterTarget, parse_target

_OPERATION_TYPES = {
    "insert_text",
    "replace_text",
    "delete_text",
    "format_text",
    "insert_table",
    "update_table",
    "delete_table",
    "insert_image",
    "update_image",
    "delete_image",
    "insert_list",
    "replace_list",
    "delete_list",
}

_INT_KEYS = {"rows", "cols", "width", "height"}
_BOOL_KEYS = {
    "list.ordered",
    "format.bold",
    "format.italic",
    "format.underline",
}
_FLOAT_KEYS = {"format.font_size", "format.line_spacing"}
_FORMAT_INT_KEYS = {"format.spacing_before", "format.spacing_after"}
_TARGET_INT_KEYS = {"occurrence", "index"}

_E = PatchSyntaxError


def parse_patch(patch_text: str) -> list[PatchOperation]:
    """Parse Writer patch text into ordered operations."""
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
    """Apply patch operations to an already-open Writer session."""
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
    from writer.session import WriterSession

    return patch_base.patch(
        path,
        patch_text,
        mode,
        session_cls=WriterSession,
        apply_fn=apply_operations,
    )


# ---- App-specific helpers ------------------------------------------------


def _parse_payload(operation_type: str, block: dict[str, str]) -> dict[str, Any]:
    payload: dict[str, Any] = {}
    for key, raw_value in block.items():
        if key == "type" or key.startswith("target."):
            continue
        if key in _INT_KEYS:
            payload[key] = coerce_int(raw_value, key, _E)
            continue
        if key in _BOOL_KEYS:
            payload[key] = coerce_bool(raw_value, key, _E)
            continue
        if key in _FLOAT_KEYS:
            payload[key] = coerce_float(raw_value, key, _E)
            continue
        if key in _FORMAT_INT_KEYS:
            payload[key] = coerce_int(raw_value, key, _E)
            continue
        if key in {"data", "items"}:
            payload[key] = coerce_json(raw_value, key, _E)
            continue
        payload[key] = raw_value

    if operation_type == "format_text":
        payload = {"formatting": _build_formatting(payload)}
    elif operation_type in {"insert_list", "replace_list"}:
        payload = {
            "ordered": payload.get("list.ordered"),
            "items": _build_list_items(payload.get("items")),
        }
    else:
        payload.pop("list.ordered", None)
    return payload


def _validate_operation_shape(
    operation_type: str,
    target: WriterTarget | None,
    payload: dict[str, Any],
) -> None:
    has_target = target is not None
    if operation_type == "insert_text":
        require_payload_keys(operation_type, payload, {"text"}, _E)
        return
    if operation_type == "replace_text":
        require_target(operation_type, has_target, _E)
        require_payload_keys(operation_type, payload, {"new_text"}, _E)
        return
    if operation_type == "delete_text":
        require_target(operation_type, has_target, _E)
        return
    if operation_type == "format_text":
        require_target(operation_type, has_target, _E)
        require_payload_keys(operation_type, payload, {"formatting"}, _E)
        return
    if operation_type == "insert_table":
        require_payload_keys(operation_type, payload, {"rows", "cols"}, _E)
        return
    if operation_type == "update_table":
        require_target(operation_type, has_target, _E)
        require_payload_keys(operation_type, payload, {"data"}, _E)
        return
    if operation_type == "delete_table":
        require_target(operation_type, has_target, _E)
        return
    if operation_type == "insert_image":
        require_payload_keys(operation_type, payload, {"image_path"}, _E)
        return
    if operation_type == "update_image":
        require_target(operation_type, has_target, _E)
        if not any(key in payload for key in {"image_path", "width", "height"}):
            raise _E(
                "Operation update_image is missing required keys:"
                " image_path, width, or height"
            )
        return
    if operation_type == "delete_image":
        require_target(operation_type, has_target, _E)
        return
    if operation_type == "insert_list":
        require_payload_keys(operation_type, payload, {"ordered", "items"}, _E)
        return
    if operation_type == "replace_list":
        require_target(operation_type, has_target, _E)
        require_payload_keys(operation_type, payload, {"items"}, _E)
        return
    if operation_type == "delete_list":
        require_target(operation_type, has_target, _E)
        return


def _dispatch_operation(session: Any, operation: PatchOperation) -> None:
    if operation.operation_type == "insert_text":
        session.insert_text(operation.payload["text"], operation.target)
        return
    if operation.operation_type == "replace_text":
        session.replace_text(operation.target, operation.payload["new_text"])
        return
    if operation.operation_type == "delete_text":
        session.delete_text(operation.target)
        return
    if operation.operation_type == "format_text":
        session.format_text(operation.target, operation.payload["formatting"])
        return
    if operation.operation_type == "insert_table":
        session.insert_table(
            operation.payload["rows"],
            operation.payload["cols"],
            operation.payload.get("data"),
            operation.payload.get("name"),
            operation.target,
        )
        return
    if operation.operation_type == "update_table":
        session.update_table(operation.target, operation.payload["data"])
        return
    if operation.operation_type == "delete_table":
        session.delete_table(operation.target)
        return
    if operation.operation_type == "insert_image":
        session.insert_image(
            operation.payload["image_path"],
            operation.payload.get("width"),
            operation.payload.get("height"),
            operation.payload.get("name"),
            operation.target,
        )
        return
    if operation.operation_type == "update_image":
        session.update_image(
            operation.target,
            image_path=operation.payload.get("image_path"),
            width=operation.payload.get("width"),
            height=operation.payload.get("height"),
        )
        return
    if operation.operation_type == "delete_image":
        session.delete_image(operation.target)
        return
    if operation.operation_type == "insert_list":
        session.insert_list(
            operation.payload["items"],
            ordered=operation.payload["ordered"],
            target=operation.target,
        )
        return
    if operation.operation_type == "replace_list":
        session.replace_list(
            operation.target,
            operation.payload["items"],
            ordered=operation.payload.get("ordered"),
        )
        return
    if operation.operation_type == "delete_list":
        session.delete_list(operation.target)
        return
    raise _E(f"Unsupported operation type: {operation.operation_type}")


def _build_formatting(payload: dict[str, Any]) -> TextFormatting:
    return TextFormatting(
        bold=payload.get("format.bold"),
        italic=payload.get("format.italic"),
        underline=payload.get("format.underline"),
        font_name=payload.get("format.font_name"),
        font_size=payload.get("format.font_size"),
        color=payload.get("format.color"),
        align=payload.get("format.align"),
        line_spacing=payload.get("format.line_spacing"),
        spacing_before=payload.get("format.spacing_before"),
        spacing_after=payload.get("format.spacing_after"),
    )


def _build_list_items(value: Any) -> list[ListItem]:
    if value is None:
        return []
    if not isinstance(value, list):
        raise _E("items must be a JSON array")
    items: list[ListItem] = []
    for entry in value:
        if not isinstance(entry, dict) or "text" not in entry:
            raise _E("Each list item must be an object with text")
        level = entry.get("level", 0)
        try:
            level_int = int(level)
        except (TypeError, ValueError) as exc:
            raise _E("List item level must be an integer") from exc
        items.append(ListItem(text=str(entry["text"]), level=level_int))
    return items
