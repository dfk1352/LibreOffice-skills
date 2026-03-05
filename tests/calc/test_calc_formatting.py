"""Tests for Calc formatting helpers."""


def test_apply_currency_format(tmp_path) -> None:
    from calc.core import create_spreadsheet
    from calc.formatting import apply_format
    from uno_bridge import uno_context

    path = tmp_path / "format.ods"
    create_spreadsheet(str(path))

    with uno_context() as desktop:
        from com.sun.star.util import NumberFormat

        doc = desktop.loadComponentFromURL(path.resolve().as_uri(), "_blank", 0, ())
        try:
            sheet = doc.Sheets.getByIndex(0)
            cell = sheet.getCellByPosition(0, 1)
            formats = doc.getNumberFormats()
            locale = getattr(formats, "getLocale", lambda: doc.CharLocale)()
            expected_key = formats.getStandardFormat(
                NumberFormat.CURRENCY,
                locale,
            )
        finally:
            doc.close(True)

    apply_format(str(path), 0, 1, 0, {"number_format": "currency"})

    with uno_context() as desktop:
        doc = desktop.loadComponentFromURL(path.resolve().as_uri(), "_blank", 0, ())
        try:
            sheet = doc.Sheets.getByIndex(0)
            cell = sheet.getCellByPosition(0, 1)
            assert cell.NumberFormat == expected_key
        finally:
            doc.close(True)


def test_apply_color_accepts_name(tmp_path) -> None:
    from calc.core import create_spreadsheet
    from calc.formatting import apply_format
    from uno_bridge import uno_context

    path = tmp_path / "format_color.ods"
    create_spreadsheet(str(path))

    apply_format(str(path), 0, 0, 0, {"color": "navy"})

    with uno_context() as desktop:
        doc = desktop.loadComponentFromURL(path.resolve().as_uri(), "_blank", 0, ())
        try:
            cell = doc.Sheets.getByIndex(0).getCellByPosition(0, 0)
            assert cell.CharColor == 0x000080
        finally:
            doc.close(True)
