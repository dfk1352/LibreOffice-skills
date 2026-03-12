"""Tests for Writer selector parsing and resolution."""

# pyright: reportMissingImports=false, reportAttributeAccessIssue=false

from __future__ import annotations

import pytest

from tests.writer._helpers import create_test_image, open_uno_doc


def _create_doc_with_named_table_and_image(doc_path, image_path):
    from writer import open_writer_session
    from writer.core import create_document

    create_document(str(doc_path))
    with open_writer_session(str(doc_path)) as session:
        session.insert_text("Introduction\n")
        session.insert_table(1, 1, [["budget"]], "BudgetTable")
        session.insert_image(str(image_path), name="Logo")


def test_parse_selector_after_content_selector():
    from writer.selectors import parse_selector

    selector = parse_selector('after:"Introduction"')

    assert selector.kind == "text"
    assert selector.strategy == "content"
    assert selector.value == "Introduction"


def test_parse_selector_name_selector():
    from writer.selectors import parse_selector

    selector = parse_selector('name:"My Table"')

    assert selector.kind is None
    assert selector.strategy == "name"
    assert selector.value == "My Table"


def test_parse_selector_index_selector():
    from writer.selectors import parse_selector

    selector = parse_selector("index:0")

    assert selector.kind is None
    assert selector.strategy == "index"
    assert selector.value == 0


def test_parse_selector_invalid_selector_raises():
    from writer.exceptions import InvalidSelectorError
    from writer.selectors import parse_selector

    with pytest.raises(InvalidSelectorError):
        parse_selector("garbage")


def test_resolve_table_selector_returns_named_table(tmp_path):
    from writer.selectors import parse_selector, resolve_table_selector

    doc_path = tmp_path / "selectors_table.odt"
    image_path = create_test_image(tmp_path / "table_helper.png")
    _create_doc_with_named_table_and_image(doc_path, image_path)

    with open_uno_doc(doc_path) as doc:
        table = resolve_table_selector(parse_selector('name:"BudgetTable"'), doc)
        assert table.Name == "BudgetTable"


def test_resolve_table_selector_unknown_name_raises(tmp_path):
    from writer.exceptions import SelectorNoMatchError
    from writer.selectors import parse_selector, resolve_table_selector

    doc_path = tmp_path / "selectors_table_missing.odt"
    image_path = create_test_image(tmp_path / "table_missing.png")
    _create_doc_with_named_table_and_image(doc_path, image_path)

    with open_uno_doc(doc_path) as doc:
        with pytest.raises(SelectorNoMatchError):
            resolve_table_selector(parse_selector('name:"MissingTable"'), doc)


def test_resolve_image_selector_index_zero_returns_graphic(tmp_path):
    from writer.selectors import parse_selector, resolve_image_selector

    doc_path = tmp_path / "selectors_image.odt"
    image_path = create_test_image(tmp_path / "image_helper.png")
    _create_doc_with_named_table_and_image(doc_path, image_path)

    with open_uno_doc(doc_path) as doc:
        graphic = resolve_image_selector(parse_selector("index:0"), doc)
        assert graphic.Name == "Logo"


def test_resolve_image_selector_out_of_range_raises(tmp_path):
    from writer.exceptions import SelectorNoMatchError
    from writer.selectors import parse_selector, resolve_image_selector

    doc_path = tmp_path / "selectors_image_missing.odt"
    image_path = create_test_image(tmp_path / "image_missing.png")
    _create_doc_with_named_table_and_image(doc_path, image_path)

    with open_uno_doc(doc_path) as doc:
        with pytest.raises(SelectorNoMatchError):
            resolve_image_selector(parse_selector("index:2"), doc)
