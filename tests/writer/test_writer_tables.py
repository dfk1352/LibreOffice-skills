"""Test Writer table operations."""

import pytest


def test_add_table_rejects_invalid_dimensions(tmp_path):
    from libreoffice_skills.writer.core import create_document
    from libreoffice_skills.writer.tables import add_table
    from libreoffice_skills.writer.exceptions import InvalidTableError

    doc_path = tmp_path / "sample.odt"
    create_document(str(doc_path))

    with pytest.raises(InvalidTableError):
        add_table(str(doc_path), 0, 2)


def test_add_empty_table(tmp_path):
    from libreoffice_skills.writer.core import create_document
    from libreoffice_skills.writer.tables import add_table
    from libreoffice_skills.uno_bridge import uno_context

    doc_path = tmp_path / "test_table.odt"
    create_document(str(doc_path))

    # Add a 3x2 table
    add_table(str(doc_path), 3, 2)

    # Verify table exists with correct dimensions via UNO roundtrip
    with uno_context() as desktop:
        doc = desktop.loadComponentFromURL(doc_path.resolve().as_uri(), "_blank", 0, ())
        try:
            tables = doc.getTextTables()
            assert tables.getCount() >= 1
            table = tables.getByIndex(tables.getCount() - 1)
            assert table.Rows.Count == 3
            assert table.Columns.Count == 2
        finally:
            doc.close(True)


def test_add_table_with_data(tmp_path):
    from libreoffice_skills.writer.core import create_document
    from libreoffice_skills.writer.tables import add_table, get_cell_name
    from libreoffice_skills.uno_bridge import uno_context

    doc_path = tmp_path / "test_table_data.odt"
    create_document(str(doc_path))

    # Add a table with data
    data = [["Name", "Age"], ["Alice", "30"], ["Bob", "25"]]

    add_table(str(doc_path), 3, 2, data)

    # Verify cell values via UNO roundtrip
    with uno_context() as desktop:
        doc = desktop.loadComponentFromURL(doc_path.resolve().as_uri(), "_blank", 0, ())
        try:
            tables = doc.getTextTables()
            table = tables.getByIndex(tables.getCount() - 1)
            for row_idx, row in enumerate(data):
                for col_idx, expected in enumerate(row):
                    cell_name = get_cell_name(row_idx, col_idx)
                    cell = table.getCellByName(cell_name)
                    assert cell.getString() == expected, (
                        f"Cell {cell_name}: expected {expected!r}, "
                        f"got {cell.getString()!r}"
                    )
        finally:
            doc.close(True)


def test_add_table_at_index(tmp_path):
    from libreoffice_skills.writer.core import create_document
    from libreoffice_skills.writer.text import insert_text
    from libreoffice_skills.writer.tables import add_table
    from libreoffice_skills.uno_bridge import uno_context

    doc_path = tmp_path / "test_table_index.odt"
    create_document(str(doc_path))

    insert_text(str(doc_path), "After", position=None)
    add_table(str(doc_path), 1, 1, [["X"]], position=0)

    with uno_context() as desktop:
        doc = desktop.loadComponentFromURL(doc_path.resolve().as_uri(), "_blank", 0, ())
        try:
            enum = doc.Text.createEnumeration()
            first = enum.nextElement()
            second = enum.nextElement()
            first_text = first.getAnchor().getString()
            second_text = second.getAnchor().getString()
            assert first_text == ""
            assert second_text == "After"
        finally:
            doc.close(True)


def test_add_table_rejects_mismatched_data(tmp_path):
    from libreoffice_skills.writer.core import create_document
    from libreoffice_skills.writer.tables import add_table
    from libreoffice_skills.writer.exceptions import InvalidTableError

    doc_path = tmp_path / "test_mismatch.odt"
    create_document(str(doc_path))

    # Data with wrong number of rows
    with pytest.raises(InvalidTableError):
        add_table(str(doc_path), 2, 2, [["A", "B"]])

    # Data with wrong number of cols
    with pytest.raises(InvalidTableError):
        add_table(str(doc_path), 2, 2, [["A"], ["B"]])


def test_get_cell_name():
    from libreoffice_skills.writer.tables import get_cell_name

    assert get_cell_name(0, 0) == "A1"
    assert get_cell_name(0, 1) == "B1"
    assert get_cell_name(1, 0) == "A2"
    assert get_cell_name(1, 1) == "B2"
    assert get_cell_name(0, 25) == "Z1"
    assert get_cell_name(0, 26) == "AA1"
