"""Tests for Calc chart creation."""


def test_create_chart(tmp_path) -> None:
    from libreoffice_skills.calc.charts import create_chart
    from libreoffice_skills.calc.core import create_spreadsheet
    from libreoffice_skills.calc.ranges import set_range
    from libreoffice_skills.uno_bridge import uno_context

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
