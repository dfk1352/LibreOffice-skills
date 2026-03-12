"""Patch parsing and application for Writer documents."""

from __future__ import annotations

import json
from dataclasses import dataclass
from typing import Any, Literal

from writer.exceptions import PatchSyntaxError

PatchApplyMode = Literal["atomic", "best_effort"]

_REQUIRED_KEYS: dict[str, set[str]] = {
    "insert_text": {"type", "text"},
    "replace_text": {"type", "selector", "new_text"},
    "delete_text": {"type", "selector"},
    "insert_table": {"type", "rows", "cols"},
    "update_table": {"type", "selector", "data"},
    "delete_table": {"type", "selector"},
    "insert_image": {"type", "image_path"},
    "update_image": {"type", "selector"},
    "delete_image": {"type", "selector"},
}

_INT_KEYS = {"rows", "cols", "width", "height"}


@dataclass(frozen=True)
class PatchOperation:
    """Parsed patch operation."""

    operation_type: str
    selector: str | None
    payload: dict[str, Any]


@dataclass
class PatchOperationResult:
    """Result for one patch operation."""

    operation_type: str
    selector: str | None
    status: str
    error: str | None
    mutated: bool


@dataclass
class PatchApplyResult:
    """Aggregate patch application result."""

    mode: PatchApplyMode
    overall_status: str
    operations: list[PatchOperationResult]
    document_persisted: bool


def parse_patch(patch_text: str) -> list[PatchOperation]:
    """Parse Writer patch text into ordered operations."""
    lines = patch_text.splitlines()
    current: dict[str, str] | None = None
    blocks: list[dict[str, str]] = []

    for line in lines:
        stripped = line.strip()
        if not stripped or stripped.startswith("#"):
            continue
        if stripped == "[operation]":
            if current is not None:
                blocks.append(current)
            current = {}
            continue
        if current is None:
            raise PatchSyntaxError("Patch content must be inside [operation] blocks")
        if "=" not in stripped:
            raise PatchSyntaxError(f"Invalid patch line: {line}")
        key, value = stripped.split("=", 1)
        current[key.strip()] = value.strip()

    if current is not None:
        blocks.append(current)

    operations: list[PatchOperation] = []
    for block in blocks:
        operation_type = block.get("type")
        if operation_type is None:
            raise PatchSyntaxError("Operation block is missing type")
        if operation_type not in _REQUIRED_KEYS:
            raise PatchSyntaxError(f"Unknown operation type: {operation_type}")
        missing = _REQUIRED_KEYS[operation_type] - block.keys()
        if missing:
            missing_list = ", ".join(sorted(missing))
            raise PatchSyntaxError(
                f"Operation {operation_type} is missing required keys: {missing_list}"
            )

        selector = block.get("selector")
        payload = _parse_payload(block)
        payload.pop("type", None)
        payload.pop("selector", None)
        operations.append(
            PatchOperation(
                operation_type=operation_type,
                selector=selector,
                payload=payload,
            )
        )
    return operations


def apply_operations(
    session, patch_text: str, mode: PatchApplyMode
) -> PatchApplyResult:
    """Apply patch operations to an already-open Writer session."""
    operations = parse_patch(patch_text)
    results: list[PatchOperationResult] = []

    if mode not in ("atomic", "best_effort"):
        raise PatchSyntaxError(f"Unsupported patch mode: {mode}")

    for index, operation in enumerate(operations):
        try:
            _dispatch_operation(session, operation)
            results.append(
                PatchOperationResult(
                    operation_type=operation.operation_type,
                    selector=operation.selector,
                    status="ok",
                    error=None,
                    mutated=True,
                )
            )
        except Exception as exc:
            results.append(
                PatchOperationResult(
                    operation_type=operation.operation_type,
                    selector=operation.selector,
                    status="failed",
                    error=str(exc),
                    mutated=False,
                )
            )
            if mode == "atomic":
                for skipped in operations[index + 1 :]:
                    results.append(
                        PatchOperationResult(
                            operation_type=skipped.operation_type,
                            selector=skipped.selector,
                            status="skipped",
                            error="Skipped because an earlier atomic operation failed",
                            mutated=False,
                        )
                    )
                session.reset()
                return PatchApplyResult(
                    mode=mode,
                    overall_status="failed",
                    operations=results,
                    document_persisted=False,
                )

    overall_status = "ok"
    if any(result.status == "failed" for result in results):
        overall_status = "partial"
    return PatchApplyResult(
        mode=mode,
        overall_status=overall_status,
        operations=results,
        document_persisted=False,
    )


def patch(
    path: str, patch_text: str, mode: PatchApplyMode = "atomic"
) -> PatchApplyResult:
    """Open a session, apply a patch, and persist if appropriate."""
    from writer.session import open_writer_session

    session = open_writer_session(path)
    try:
        result = apply_operations(session, patch_text, mode)
        should_save = result.overall_status != "failed" and any(
            operation.mutated for operation in result.operations
        )
        session.close(save=should_save)
        result.document_persisted = should_save
        return result
    finally:
        if not session._closed:
            session.close(save=False)


def _parse_payload(block: dict[str, str]) -> dict[str, Any]:
    payload: dict[str, Any] = dict(block)
    for key in _INT_KEYS:
        if key in payload:
            try:
                payload[key] = int(payload[key])
            except ValueError as exc:
                raise PatchSyntaxError(f"{key} must be an integer") from exc
    if "data" in payload:
        try:
            payload["data"] = json.loads(payload["data"])
        except json.JSONDecodeError as exc:
            raise PatchSyntaxError("data must be valid JSON") from exc
    return payload


def _dispatch_operation(session, operation: PatchOperation) -> None:
    if operation.operation_type == "insert_text":
        session.insert_text(operation.payload["text"], operation.selector)
        return
    if operation.operation_type == "replace_text":
        session.replace_text(operation.selector, operation.payload["new_text"])
        return
    if operation.operation_type == "delete_text":
        session.delete_text(operation.selector)
        return
    if operation.operation_type == "insert_table":
        session.insert_table(
            operation.payload["rows"],
            operation.payload["cols"],
            operation.payload.get("data"),
            operation.payload.get("name"),
            operation.selector,
        )
        return
    if operation.operation_type == "update_table":
        session.update_table(operation.selector, operation.payload["data"])
        return
    if operation.operation_type == "delete_table":
        session.delete_table(operation.selector)
        return
    if operation.operation_type == "insert_image":
        session.insert_image(
            operation.payload["image_path"],
            operation.payload.get("width"),
            operation.payload.get("height"),
            operation.payload.get("name"),
            operation.selector,
        )
        return
    if operation.operation_type == "update_image":
        session.update_image(
            operation.selector,
            image_path=operation.payload.get("image_path"),
            width=operation.payload.get("width"),
            height=operation.payload.get("height"),
        )
        return
    if operation.operation_type == "delete_image":
        session.delete_image(operation.selector)
        return
    raise PatchSyntaxError(f"Unsupported operation type: {operation.operation_type}")
