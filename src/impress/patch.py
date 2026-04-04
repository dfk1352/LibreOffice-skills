"""Patch parsing and application for Impress presentations."""

from __future__ import annotations

from typing import Any

import patch_base
from patch_base import (
    PatchApplyMode,
    PatchApplyResult,
    PatchOperation,
    coerce_bool,
    coerce_float,
    coerce_int,
    coerce_json,
    require_payload_keys,
    require_target,
)
from impress.exceptions import PatchSyntaxError
from impress.targets import (
    ImpressTarget,
    ListItem,
    ShapePlacement,
    TextFormatting,
    parse_target,
)

_OPERATION_TYPES = {
    "add_slide",
    "delete_slide",
    "move_slide",
    "duplicate_slide",
    "insert_text",
    "replace_text",
    "format_text",
    "insert_list",
    "replace_list",
    "insert_text_box",
    "insert_shape",
    "delete_item",
    "insert_image",
    "replace_image",
    "insert_table",
    "update_table",
    "insert_chart",
    "update_chart",
    "insert_media",
    "replace_media",
    "set_notes",
    "apply_master_page",
    "set_master_background",
}
_TARGET_INT_KEYS = {"slide_index", "shape_index", "occurrence"}
_INT_KEYS = {"index", "to_index", "rows", "cols"}
_BOOL_KEYS = {
    "list.ordered",
    "format.bold",
    "format.italic",
    "format.underline",
}
_FLOAT_KEYS = {
    "placement.x_cm",
    "placement.y_cm",
    "placement.width_cm",
    "placement.height_cm",
    "format.font_size",
}

_E = PatchSyntaxError


def parse_patch(patch_text: str) -> list[PatchOperation]:
    """Parse Impress patch text into ordered operations."""
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
    """Apply patch operations to an already-open Impress session."""
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
    from impress.session import ImpressSession

    return patch_base.patch(
        path,
        patch_text,
        mode,
        session_cls=ImpressSession,
        apply_fn=apply_operations,
    )


# ---- App-specific helpers ------------------------------------------------


def _dispatch_operation(session: Any, operation: PatchOperation) -> None:
    op = operation.operation_type
    if op == "add_slide":
        session.add_slide(
            index=operation.payload.get("index"),
            layout=operation.payload.get("layout", "BLANK"),
        )
        return
    if op == "delete_slide":
        session.delete_slide(operation.target)
        return
    if op == "move_slide":
        session.move_slide(operation.target, operation.payload["to_index"])
        return
    if op == "duplicate_slide":
        session.duplicate_slide(operation.target)
        return
    if op == "insert_text":
        session.insert_text(operation.payload["text"], operation.target)
        return
    if op == "replace_text":
        session.replace_text(operation.target, operation.payload["new_text"])
        return
    if op == "format_text":
        session.format_text(operation.target, operation.payload["formatting"])
        return
    if op == "insert_list":
        session.insert_list(
            operation.payload["items"],
            ordered=operation.payload["ordered"],
            target=operation.target,
        )
        return
    if op == "replace_list":
        session.replace_list(
            operation.target,
            operation.payload["items"],
            ordered=operation.payload.get("ordered"),
        )
        return
    if op == "insert_text_box":
        session.insert_text_box(
            operation.target,
            operation.payload["text"],
            operation.payload["placement"],
            name=operation.payload.get("name"),
        )
        return
    if op == "insert_shape":
        session.insert_shape(
            operation.target,
            operation.payload["shape_type"],
            operation.payload["placement"],
            fill_color=operation.payload.get("fill_color"),
            line_color=operation.payload.get("line_color"),
            name=operation.payload.get("name"),
        )
        return
    if op == "delete_item":
        session.delete_item(operation.target)
        return
    if op == "insert_image":
        session.insert_image(
            operation.target,
            operation.payload["image_path"],
            operation.payload["placement"],
            name=operation.payload.get("name"),
        )
        return
    if op == "replace_image":
        session.replace_image(
            operation.target,
            image_path=operation.payload.get("image_path"),
            placement=operation.payload.get("placement"),
        )
        return
    if op == "insert_table":
        session.insert_table(
            operation.target,
            operation.payload["rows"],
            operation.payload["cols"],
            operation.payload["placement"],
            data=operation.payload.get("data"),
            name=operation.payload.get("name"),
        )
        return
    if op == "update_table":
        session.update_table(operation.target, operation.payload["data"])
        return
    if op == "insert_chart":
        session.insert_chart(
            operation.target,
            operation.payload["chart_type"],
            operation.payload["data"],
            operation.payload["placement"],
            title=operation.payload.get("title"),
            name=operation.payload.get("name"),
        )
        return
    if op == "update_chart":
        session.update_chart(
            operation.target,
            chart_type=operation.payload.get("chart_type"),
            data=operation.payload.get("data"),
            placement=operation.payload.get("placement"),
            title=operation.payload.get("title"),
        )
        return
    if op == "insert_media":
        session.insert_media(
            operation.target,
            operation.payload["media_path"],
            operation.payload["placement"],
            name=operation.payload.get("name"),
        )
        return
    if op == "replace_media":
        session.replace_media(
            operation.target,
            media_path=operation.payload.get("media_path"),
            placement=operation.payload.get("placement"),
        )
        return
    if op == "set_notes":
        session.set_notes(operation.target, operation.payload["text"])
        return
    if op == "apply_master_page":
        session.apply_master_page(operation.target, operation.payload["master_target"])
        return
    if op == "set_master_background":
        session.set_master_background(operation.target, operation.payload["color"])
        return
    raise _E(f"Unsupported operation type: {op}")


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
        if key in {"items", "data"}:
            payload[key] = coerce_json(raw_value, key, _E)
            continue
        payload[key] = raw_value

    if any(key.startswith("format.") for key in payload):
        payload["formatting"] = TextFormatting(
            bold=payload.get("format.bold"),
            italic=payload.get("format.italic"),
            underline=payload.get("format.underline"),
            font_name=payload.get("format.font_name"),
            font_size=payload.get("format.font_size"),
            color=payload.get("format.color"),
            align=payload.get("format.align"),
        )
    if any(key.startswith("placement.") for key in payload):
        payload["placement"] = ShapePlacement(
            x_cm=_require_float_payload(payload, "placement.x_cm"),
            y_cm=_require_float_payload(payload, "placement.y_cm"),
            width_cm=_require_float_payload(payload, "placement.width_cm"),
            height_cm=_require_float_payload(payload, "placement.height_cm"),
        )
    if operation_type in {"insert_list", "replace_list"}:
        payload["items"] = _build_list_items(payload.get("items"))
        payload["ordered"] = payload.get("list.ordered")
    if any(key.startswith("master.") for key in block):
        master_fields = {
            key.split(".", 1)[1]: patch_base.coerce_target_value(
                key.split(".", 1)[1], value, _TARGET_INT_KEYS, _E
            )
            for key, value in block.items()
            if key.startswith("master.")
        }
        payload["master_target"] = parse_target(master_fields)
    return payload


def _validate_operation_shape(
    operation_type: str,
    target: ImpressTarget | None,
    payload: dict[str, Any],
) -> None:
    has_target = target is not None
    if operation_type == "add_slide":
        return
    if operation_type in {
        "delete_slide",
        "duplicate_slide",
        "delete_item",
    }:
        require_target(operation_type, has_target, _E)
        return
    if operation_type == "move_slide":
        require_target(operation_type, has_target, _E)
        require_payload_keys(operation_type, payload, {"to_index"}, _E)
        return
    if operation_type == "insert_text":
        require_payload_keys(operation_type, payload, {"text"}, _E)
        return
    if operation_type == "replace_text":
        require_target(operation_type, has_target, _E)
        require_payload_keys(operation_type, payload, {"new_text"}, _E)
        return
    if operation_type == "format_text":
        require_target(operation_type, has_target, _E)
        require_payload_keys(operation_type, payload, {"formatting"}, _E)
        return
    if operation_type == "insert_list":
        require_payload_keys(operation_type, payload, {"ordered", "items"}, _E)
        return
    if operation_type == "replace_list":
        require_target(operation_type, has_target, _E)
        require_payload_keys(operation_type, payload, {"items"}, _E)
        return
    if operation_type == "insert_text_box":
        require_target(operation_type, has_target, _E)
        require_payload_keys(operation_type, payload, {"text", "placement"}, _E)
        return
    if operation_type == "insert_shape":
        require_target(operation_type, has_target, _E)
        require_payload_keys(
            operation_type,
            payload,
            {"shape_type", "placement"},
            _E,
        )
        return
    if operation_type == "insert_image":
        require_target(operation_type, has_target, _E)
        require_payload_keys(
            operation_type,
            payload,
            {"image_path", "placement"},
            _E,
        )
        return
    if operation_type == "replace_image":
        require_target(operation_type, has_target, _E)
        if "image_path" not in payload and "placement" not in payload:
            raise _E(
                "Operation replace_image is missing required keys:"
                " image_path or placement"
            )
        return
    if operation_type == "insert_table":
        require_target(operation_type, has_target, _E)
        require_payload_keys(
            operation_type,
            payload,
            {"rows", "cols", "placement"},
            _E,
        )
        return
    if operation_type == "update_table":
        require_target(operation_type, has_target, _E)
        require_payload_keys(operation_type, payload, {"data"}, _E)
        return
    if operation_type == "insert_chart":
        require_target(operation_type, has_target, _E)
        require_payload_keys(
            operation_type,
            payload,
            {"chart_type", "data", "placement"},
            _E,
        )
        return
    if operation_type == "update_chart":
        require_target(operation_type, has_target, _E)
        if not any(
            key in payload
            for key in {
                "chart_type",
                "data",
                "placement",
                "title",
            }
        ):
            raise _E(
                "Operation update_chart is missing required keys:"
                " chart_type, data, placement, or title"
            )
        return
    if operation_type == "insert_media":
        require_target(operation_type, has_target, _E)
        require_payload_keys(
            operation_type,
            payload,
            {"media_path", "placement"},
            _E,
        )
        return
    if operation_type == "replace_media":
        require_target(operation_type, has_target, _E)
        if "media_path" not in payload and "placement" not in payload:
            raise _E(
                "Operation replace_media is missing required keys:"
                " media_path or placement"
            )
        return
    if operation_type == "set_notes":
        require_target(operation_type, has_target, _E)
        require_payload_keys(operation_type, payload, {"text"}, _E)
        return
    if operation_type == "apply_master_page":
        require_target(operation_type, has_target, _E)
        require_payload_keys(operation_type, payload, {"master_target"}, _E)
        return
    if operation_type == "set_master_background":
        require_target(operation_type, has_target, _E)
        require_payload_keys(operation_type, payload, {"color"}, _E)
        return


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
        if isinstance(level, bool):
            raise _E("List item level must be an integer")
        try:
            level_value = int(level)
        except (TypeError, ValueError) as exc:
            raise _E("List item level must be an integer") from exc
        items.append(ListItem(text=str(entry["text"]), level=level_value))
    return items


def _require_float_payload(payload: dict[str, Any], key: str) -> float:
    value = payload.get(key)
    if not isinstance(value, (int, float)):
        raise _E(f"{key} must be a number")
    return float(value)
