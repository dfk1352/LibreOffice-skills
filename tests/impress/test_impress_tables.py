"""Tests for Impress table operations."""


def test_add_table_returns_index(tmp_path):
    from impress.core import create_presentation
    from impress.tables import add_table
    from uno_bridge import uno_context

    path = tmp_path / "table.odp"
    create_presentation(str(path))

    result = add_table(str(path), 0, 3, 2, 2.0, 5.0, 15.0, 8.0)

    assert isinstance(result, int)

    # Verify table shape has correct dimensions via UNO roundtrip
    with uno_context() as desktop:
        doc = desktop.loadComponentFromURL(path.resolve().as_uri(), "_blank", 0, ())
        try:
            slide = doc.DrawPages.getByIndex(0)
            table_shape = slide.getByIndex(result)
            model = table_shape.Model
            assert model.Rows.Count == 3
            assert model.Columns.Count == 2
        finally:
            doc.close(True)


def test_add_table_with_data(tmp_path):
    from impress.core import create_presentation
    from impress.tables import add_table
    from uno_bridge import uno_context

    path = tmp_path / "table_data.odp"
    create_presentation(str(path))

    data = [["Name", "Score"], ["Alice", "95"], ["Bob", "87"]]
    shape_idx = add_table(str(path), 0, 3, 2, 2.0, 5.0, 15.0, 8.0, data=data)

    # Verify via UNO roundtrip
    with uno_context() as desktop:
        doc = desktop.loadComponentFromURL(
            (tmp_path / "table_data.odp").resolve().as_uri(),
            "_blank",
            0,
            (),
        )
        try:
            slide = doc.DrawPages.getByIndex(0)
            table_shape = slide.getByIndex(shape_idx)
            model = table_shape.Model
            cell = model.getCellByPosition(0, 0)
            assert cell.getString() == "Name"
            cell = model.getCellByPosition(1, 1)
            assert cell.getString() == "95"
        finally:
            doc.close(True)


def test_set_table_cell(tmp_path):
    from impress.core import create_presentation
    from impress.tables import add_table, set_table_cell
    from uno_bridge import uno_context

    path = tmp_path / "set_cell.odp"
    create_presentation(str(path))

    shape_idx = add_table(str(path), 0, 2, 2, 2.0, 5.0, 10.0, 5.0)
    set_table_cell(str(path), 0, shape_idx, 0, 0, "Updated")

    with uno_context() as desktop:
        doc = desktop.loadComponentFromURL(
            path.resolve().as_uri(),
            "_blank",
            0,
            (),
        )
        try:
            slide = doc.DrawPages.getByIndex(0)
            table_shape = slide.getByIndex(shape_idx)
            model = table_shape.Model
            cell = model.getCellByPosition(0, 0)
            assert cell.getString() == "Updated"
        finally:
            doc.close(True)


def test_format_table_cell(tmp_path):
    from impress.core import create_presentation
    from impress.tables import (
        add_table,
        format_table_cell,
        set_table_cell,
    )
    from uno_bridge import uno_context

    path = tmp_path / "format_cell.odp"
    create_presentation(str(path))

    shape_idx = add_table(str(path), 0, 2, 2, 2.0, 5.0, 10.0, 5.0)
    set_table_cell(str(path), 0, shape_idx, 0, 0, "Header")
    format_table_cell(
        str(path),
        0,
        shape_idx,
        0,
        0,
        bold=True,
        font_size=14,
        fill_color="lightgray",
    )

    # Verify formatting via UNO roundtrip
    with uno_context() as desktop:
        doc = desktop.loadComponentFromURL(path.resolve().as_uri(), "_blank", 0, ())
        try:
            slide = doc.DrawPages.getByIndex(0)
            table_shape = slide.getByIndex(shape_idx)
            model = table_shape.Model
            cell = model.getCellByPosition(0, 0)
            cursor = cell.getText().createTextCursor()
            cursor.gotoStart(False)
            cursor.gotoEnd(True)
            assert cursor.CharWeight == 150  # Bold
            assert cursor.CharHeight == 14
        finally:
            doc.close(True)
