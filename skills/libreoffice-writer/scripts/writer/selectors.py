"""Selector parsing and resolution for Writer documents."""

from __future__ import annotations

from dataclasses import dataclass
from typing import Any

from writer.exceptions import (
    InvalidSelectorError,
    SelectorAmbiguousError,
    SelectorNoMatchError,
)


@dataclass(frozen=True)
class Selector:
    """Parsed selector value."""

    kind: str | None
    strategy: str
    value: str | int
    operator: str


def parse_selector(raw: str) -> Selector:
    """Parse a selector string into a structured value."""
    if raw.startswith("after:"):
        return Selector("text", "content", _parse_quoted_value(raw, "after"), "after")
    if raw.startswith("before:"):
        return Selector("text", "content", _parse_quoted_value(raw, "before"), "before")
    if raw.startswith("contains:"):
        return Selector(
            "text",
            "content",
            _parse_quoted_value(raw, "contains"),
            "contains",
        )
    if raw.startswith("name:"):
        return Selector(None, "name", _parse_quoted_value(raw, "name"), "name")
    if raw.startswith("index:"):
        suffix = raw[len("index:") :]
        if not suffix.isdigit():
            raise InvalidSelectorError(f"Invalid selector: {raw}")
        return Selector(None, "index", int(suffix), "index")
    raise InvalidSelectorError(f"Invalid selector: {raw}")


def resolve_text_selector(selector: Selector, doc: Any) -> Any:
    """Resolve a text selector to a UNO text range or insertion cursor."""
    if selector.strategy != "content" or not isinstance(selector.value, str):
        raise InvalidSelectorError("Text selectors must use content matching")

    matches = _find_text_matches(doc, selector.value)
    if not matches:
        raise SelectorNoMatchError(f'No text matched selector "{selector.value}"')
    if len(matches) > 1:
        raise SelectorAmbiguousError(
            f'Selector matched multiple spans: "{selector.value}"'
        )

    match = matches[0]
    if selector.operator == "contains":
        return match

    text_obj = doc.Text
    anchor = match.End if selector.operator == "after" else match.Start
    return text_obj.createTextCursorByRange(anchor)


def resolve_table_selector(selector: Selector, doc: Any) -> Any:
    """Resolve a table selector to a UNO text table."""
    tables = doc.getTextTables()
    if selector.strategy == "name":
        table_name = str(selector.value)
        table_names = list(tables.getElementNames())
        matches = [
            candidate
            for candidate in table_names
            if _normalize_name(candidate) == _normalize_name(table_name)
        ]
        if not matches:
            raise SelectorNoMatchError(f'Table not found: "{table_name}"')
        if len(matches) > 1:
            raise SelectorAmbiguousError(f'Table selector is ambiguous: "{table_name}"')
        return tables.getByName(matches[0])

    if selector.strategy == "index":
        index = int(selector.value)
        if index < 0 or index >= tables.getCount():
            raise SelectorNoMatchError(f"Table index out of range: {index}")
        return tables.getByIndex(index)

    raise InvalidSelectorError("Table selectors must use name or index")


def resolve_image_selector(selector: Selector, doc: Any) -> Any:
    """Resolve an image selector to a UNO graphic object."""
    graphics = doc.getGraphicObjects()
    if selector.strategy == "name":
        graphic_name = str(selector.value)
        graphic_names = list(graphics.getElementNames())
        matches = [
            candidate
            for candidate in graphic_names
            if _normalize_name(candidate) == _normalize_name(graphic_name)
        ]
        if not matches:
            raise SelectorNoMatchError(f'Image not found: "{graphic_name}"')
        if len(matches) > 1:
            raise SelectorAmbiguousError(
                f'Image selector is ambiguous: "{graphic_name}"'
            )
        return graphics.getByName(matches[0])

    if selector.strategy == "index":
        index = int(selector.value)
        graphic_names = list(graphics.getElementNames())
        if index < 0 or index >= len(graphic_names):
            raise SelectorNoMatchError(f"Image index out of range: {index}")
        return graphics.getByName(graphic_names[index])

    raise InvalidSelectorError("Image selectors must use name or index")


def _parse_quoted_value(raw: str, prefix: str) -> str:
    expected_prefix = f'{prefix}:"'
    if not raw.startswith(expected_prefix) or not raw.endswith('"'):
        raise InvalidSelectorError(f"Invalid selector: {raw}")
    return raw[len(expected_prefix) : -1]


def _find_text_matches(doc: Any, needle: str) -> list[Any]:
    search = doc.createSearchDescriptor()
    search.SearchString = needle

    matches: list[Any] = []
    found = doc.findFirst(search)
    while found is not None:
        matches.append(found)
        found = doc.findNext(found.End, search)
    return matches


def _normalize_name(value: str) -> str:
    return "_".join(value.split())
