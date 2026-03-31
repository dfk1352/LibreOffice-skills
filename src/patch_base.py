"""Shared patch infrastructure for all LibreOffice skill apps.

Provides the common DSL parser, coercion helpers, dataclasses, and
apply/patch orchestration used by Writer, Calc, and Impress.
"""

from __future__ import annotations

import json
from dataclasses import dataclass
from typing import Any, Callable, Literal

from exceptions import PatchSyntaxError

PatchApplyMode = Literal["atomic", "best_effort"]


# ---------------------------------------------------------------------------
# Dataclasses
# ---------------------------------------------------------------------------


@dataclass(frozen=True)
class PatchOperation:
    """A single parsed patch operation."""

    operation_type: str
    target: Any
    payload: dict[str, Any]


@dataclass
class PatchOperationResult:
    """Result of a single patch operation."""

    operation_type: str
    target: Any
    status: str  # "ok" | "failed" | "skipped"
    error: str | None
    mutated: bool


@dataclass
class PatchApplyResult:
    """Aggregate result of applying a patch."""

    mode: PatchApplyMode
    overall_status: str  # "ok" | "partial" | "failed"
    operations: list[PatchOperationResult]
    document_persisted: bool


# ---------------------------------------------------------------------------
# Block parser
# ---------------------------------------------------------------------------


def parse_blocks(
    patch_text: str,
    error_cls: type[PatchSyntaxError] = PatchSyntaxError,
) -> list[dict[str, str]]:
    """Split patch text into raw key-value blocks.

    Each ``[operation]`` header starts a new block. Lines within a block
    are ``key = value`` pairs. Heredoc syntax (``key <<MARKER``) is
    supported for multiline values.

    Args:
        patch_text: Raw patch DSL text.
        error_cls: Exception class for syntax errors.

    Raises:
        PatchSyntaxError: On malformed input.
    """
    lines = patch_text.splitlines()
    current: dict[str, str] | None = None
    blocks: list[dict[str, str]] = []
    index = 0
    while index < len(lines):
        line = lines[index]
        stripped = line.strip()
        if not stripped or stripped.startswith("#"):
            index += 1
            continue
        if stripped == "[operation]":
            if current is not None:
                blocks.append(current)
            current = {}
            index += 1
            continue
        if current is None:
            raise error_cls("Patch content must be inside [operation] blocks")
        if "<<" in stripped:
            key, marker = stripped.split("<<", 1)
            end_marker = marker.strip()
            value_lines: list[str] = []
            index += 1
            while index < len(lines) and lines[index].strip() != end_marker:
                value_lines.append(lines[index])
                index += 1
            if index >= len(lines):
                raise error_cls(f"Unterminated heredoc for key: {key.strip()}")
            current[key.strip()] = "\n".join(value_lines)
            index += 1
            continue
        if "=" not in stripped:
            raise error_cls(f"Invalid patch line: {line}")
        key, value = stripped.split("=", 1)
        current[key.strip()] = value.strip()
        index += 1

    if current is not None:
        blocks.append(current)
    return blocks


# ---------------------------------------------------------------------------
# Coercion helpers
# ---------------------------------------------------------------------------


def coerce_int(
    value: str,
    key: str,
    error_cls: type[PatchSyntaxError] = PatchSyntaxError,
) -> int:
    """Coerce a string to int or raise ``error_cls``."""
    try:
        return int(value)
    except ValueError as exc:
        raise error_cls(f"{key} must be an integer") from exc


def coerce_float(
    value: str,
    key: str,
    error_cls: type[PatchSyntaxError] = PatchSyntaxError,
) -> float:
    """Coerce a string to float or raise ``error_cls``."""
    try:
        return float(value)
    except ValueError as exc:
        raise error_cls(f"{key} must be a number") from exc


def coerce_bool(
    value: str,
    key: str,
    error_cls: type[PatchSyntaxError] = PatchSyntaxError,
) -> bool:
    """Coerce a string to bool or raise ``error_cls``."""
    lowered = value.strip().lower()
    if lowered == "true":
        return True
    if lowered == "false":
        return False
    raise error_cls(f"{key} must be true or false")


def coerce_json(
    value: str,
    key: str,
    error_cls: type[PatchSyntaxError] = PatchSyntaxError,
) -> Any:
    """Coerce a string to a JSON value or raise ``error_cls``."""
    try:
        return json.loads(value)
    except json.JSONDecodeError as exc:
        raise error_cls(f"{key} must be valid JSON") from exc


def coerce_target_value(
    field: str,
    value: str,
    int_keys: set[str],
    error_cls: type[PatchSyntaxError] = PatchSyntaxError,
) -> Any:
    """Coerce a target field value, converting int-keys to ``int``."""
    if field in int_keys:
        return coerce_int(value, f"target.{field}", error_cls)
    return value


# ---------------------------------------------------------------------------
# Validation helpers
# ---------------------------------------------------------------------------


def require_target(
    operation_type: str,
    has_target: bool,
    error_cls: type[PatchSyntaxError] = PatchSyntaxError,
) -> None:
    """Raise ``error_cls`` if a target is required but missing."""
    if not has_target:
        raise error_cls(f"Operation {operation_type} requires target.* fields")


def require_payload_keys(
    operation_type: str,
    payload: dict[str, Any],
    keys: set[str],
    error_cls: type[PatchSyntaxError] = PatchSyntaxError,
) -> None:
    """Raise ``error_cls`` if required payload keys are missing."""
    missing = [
        key for key in sorted(keys) if key not in payload or payload[key] is None
    ]
    if missing:
        raise error_cls(
            f"Operation {operation_type} is missing required keys: {', '.join(missing)}"
        )


# ---------------------------------------------------------------------------
# Orchestration
# ---------------------------------------------------------------------------


def parse_patch(
    patch_text: str,
    *,
    operation_types: set[str],
    target_int_keys: set[str],
    parse_target_fn: Callable[[dict[str, Any]], Any],
    parse_payload_fn: Callable[[str, dict[str, str]], dict[str, Any]],
    validate_fn: Callable[[str, Any, dict[str, Any]], None],
    error_cls: type[PatchSyntaxError] = PatchSyntaxError,
) -> list[PatchOperation]:
    """Parse patch text into ``PatchOperation`` objects.

    Args:
        patch_text: Raw patch DSL text.
        operation_types: Set of valid operation type strings.
        target_int_keys: Field names that should be coerced to ``int``.
        parse_target_fn: App-specific target constructor.
        parse_payload_fn: App-specific payload parser.
        validate_fn: App-specific operation shape validator.
        error_cls: Exception class for syntax errors.

    Returns:
        Ordered list of parsed operations.

    Raises:
        PatchSyntaxError: On invalid input.
    """
    blocks = parse_blocks(patch_text, error_cls)
    operations: list[PatchOperation] = []
    for block in blocks:
        operation_type = block.get("type")
        if operation_type is None:
            raise error_cls("Operation block is missing type")
        if operation_type not in operation_types:
            raise error_cls(f"Unknown operation type: {operation_type}")

        target_fields = {
            key.split(".", 1)[1]: coerce_target_value(
                key.split(".", 1)[1], value, target_int_keys, error_cls
            )
            for key, value in block.items()
            if key.startswith("target.")
        }
        target = parse_target_fn(target_fields) if target_fields else None
        payload = parse_payload_fn(operation_type, block)
        validate_fn(operation_type, target, payload)
        operations.append(
            PatchOperation(
                operation_type=operation_type,
                target=target,
                payload=payload,
            )
        )
    return operations


def apply_operations(
    session: Any,
    patch_text: str,
    mode: PatchApplyMode,
    *,
    parse_patch_fn: Callable[[str], list[PatchOperation]],
    dispatch_fn: Callable[[Any, PatchOperation], None],
    error_cls: type[PatchSyntaxError] = PatchSyntaxError,
) -> PatchApplyResult:
    """Apply patch operations to an already-open session.

    Args:
        session: An open app session with ``_path`` and
            ``restore_snapshot`` attributes.
        patch_text: Raw patch DSL text.
        mode: ``"atomic"`` or ``"best_effort"``.
        parse_patch_fn: Bound parse function for this app.
        dispatch_fn: App-specific operation dispatcher.
        error_cls: Exception class for syntax errors.

    Returns:
        Aggregate result of the patch application.

    Raises:
        PatchSyntaxError: On invalid mode or input.
    """
    if mode not in ("atomic", "best_effort"):
        raise error_cls(f"Unsupported patch mode: {mode}")

    operations = parse_patch_fn(patch_text)
    results: list[PatchOperationResult] = []
    if mode == "atomic":
        session._doc.store()
        atomic_snapshot = session._path.read_bytes()
    else:
        atomic_snapshot = None

    for index, operation in enumerate(operations):
        try:
            dispatch_fn(session, operation)
            results.append(
                PatchOperationResult(
                    operation_type=operation.operation_type,
                    target=operation.target,
                    status="ok",
                    error=None,
                    mutated=True,
                )
            )
        except Exception as exc:
            results.append(
                PatchOperationResult(
                    operation_type=operation.operation_type,
                    target=operation.target,
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
                            target=skipped.target,
                            status="skipped",
                            error=(
                                "Skipped because an earlier atomic operation failed"
                            ),
                            mutated=False,
                        )
                    )
                assert atomic_snapshot is not None
                session.restore_snapshot(atomic_snapshot)
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
    path: str,
    patch_text: str,
    mode: PatchApplyMode,
    *,
    session_cls: type,
    apply_fn: Callable[..., PatchApplyResult],
) -> PatchApplyResult:
    """Open a session, apply a patch, and persist if appropriate.

    Args:
        path: Path to the document.
        patch_text: Raw patch DSL text.
        mode: ``"atomic"`` or ``"best_effort"``.
        session_cls: App-specific session class to instantiate.
        apply_fn: Bound ``apply_operations`` for this app.

    Returns:
        Aggregate result with ``document_persisted`` set.
    """
    session = session_cls(path)
    try:
        result = apply_fn(session, patch_text, mode)
        should_save = result.overall_status != "failed" and any(
            operation.mutated for operation in result.operations
        )
        session.close(save=should_save)
        result.document_persisted = should_save
        return result
    finally:
        if not session._closed:
            session.close(save=False)
