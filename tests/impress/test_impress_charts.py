"""Tests for Impress chart operations."""

# pyright: reportMissingImports=false

import pytest


def test_add_chart_missing_doc_raises(tmp_path):
    from impress.charts import add_chart
    from impress.exceptions import DocumentNotFoundError

    data = [["Category", "Value"], ["A", 10]]
    with pytest.raises(DocumentNotFoundError):
        add_chart(str(tmp_path / "no_such.odp"), 0, "bar", data, 1.0, 1.0, 5.0, 3.0)


def test_add_chart_returns_index(tmp_path):
    from impress.charts import add_chart
    from impress.core import create_presentation
    from uno_bridge import uno_context

    path = tmp_path / "chart.odp"
    create_presentation(str(path))

    data = [["Category", "Value"], ["A", 10], ["B", 20], ["C", 30]]
    result = add_chart(
        str(path),
        0,
        "bar",
        data,
        2.0,
        2.0,
        15.0,
        10.0,
        title="Sales",
    )

    assert isinstance(result, int)

    # Verify chart shape exists on the slide via UNO roundtrip
    with uno_context() as desktop:
        doc = desktop.loadComponentFromURL(path.resolve().as_uri(), "_blank", 0, ())
        try:
            slide = doc.DrawPages.getByIndex(0)
            shape = slide.getByIndex(result)
            # The shape should be an OLE2 shape containing a chart
            assert shape.supportsService(
                "com.sun.star.drawing.OLE2Shape"
            ) or shape.supportsService("com.sun.star.presentation.OLE2Shape")
        finally:
            doc.close(True)


def test_add_chart_pie(tmp_path):
    from impress.charts import add_chart
    from impress.core import create_presentation
    from uno_bridge import uno_context

    path = tmp_path / "pie.odp"
    create_presentation(str(path))

    data = [["Slice", "Value"], ["Red", 40], ["Blue", 60]]
    result = add_chart(
        str(path),
        0,
        "pie",
        data,
        2.0,
        2.0,
        12.0,
        10.0,
    )

    assert isinstance(result, int)

    # Verify chart shape exists on the slide via UNO roundtrip
    with uno_context() as desktop:
        doc = desktop.loadComponentFromURL(path.resolve().as_uri(), "_blank", 0, ())
        try:
            slide = doc.DrawPages.getByIndex(0)
            shape = slide.getByIndex(result)
            assert shape.supportsService(
                "com.sun.star.drawing.OLE2Shape"
            ) or shape.supportsService("com.sun.star.presentation.OLE2Shape")
        finally:
            doc.close(True)


def test_add_chart_rejects_invalid_type(tmp_path):
    from impress.charts import add_chart
    from impress.core import create_presentation
    from impress.exceptions import ImpressSkillError

    path = tmp_path / "bad_chart.odp"
    create_presentation(str(path))

    data = [["X", "Y"], [1, 2]]
    with pytest.raises(ImpressSkillError):
        add_chart(
            str(path),
            0,
            "radar",
            data,
            2.0,
            2.0,
            10.0,
            8.0,
        )
