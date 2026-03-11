"""Tests for Calc chart creation."""

import pytest


def test_create_chart(tmp_path) -> None:
    from calc.charts import create_chart
    from calc.core import create_spreadsheet
    from calc.ranges import set_range
    from uno_bridge import uno_context

    path = tmp_path / "chart.ods"
    create_spreadsheet(str(path))
    set_range(str(path), 0, 0, 0, [["Label", "Value"], ["A", 10], ["B", 20]])

    create_chart(
        str(path),
        0,
        (0, 0, 2, 1),
        "bar",
        anchor=(5, 0),
        size=(5000, 4000),
        title="Sample",
    )

    with uno_context() as desktop:
        doc = desktop.loadComponentFromURL(path.resolve().as_uri(), "_blank", 0, ())
        try:
            sheet = doc.Sheets.getByIndex(0)
            charts = sheet.Charts
            assert charts.getCount() == 1
            name = charts.getByIndex(0).Name
            table_chart = charts.getByName(name)
            ranges = table_chart.getRanges()
            assert len(ranges) == 1
            assert ranges[0].EndRow == 2
            assert ranges[0].EndColumn == 1
            embedded = table_chart.EmbeddedObject
            assert embedded.HasMainTitle
            assert embedded.Title.String == "Sample"
        finally:
            doc.close(True)


def test_create_chart_raises_on_missing_file(tmp_path) -> None:
    from calc.charts import create_chart
    from calc.exceptions import DocumentNotFoundError

    with pytest.raises(DocumentNotFoundError):
        create_chart(
            str(tmp_path / "missing.ods"),
            "Sheet1",
            (0, 0, 1, 1),
            "bar",
            anchor=(0, 0),
            size=(100, 100),
        )
