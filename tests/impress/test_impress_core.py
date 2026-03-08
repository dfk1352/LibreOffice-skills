"""Tests for Impress core presentation operations."""

# pyright: reportMissingImports=false

import zipfile

import pytest


def test_export_presentation_missing_doc_raises(tmp_path):
    from impress.core import export_presentation
    from impress.exceptions import DocumentNotFoundError

    with pytest.raises(DocumentNotFoundError):
        export_presentation(
            str(tmp_path / "no_such.odp"),
            str(tmp_path / "out.pdf"),
            "pdf",
        )


def test_create_presentation_creates_file(tmp_path):
    from impress.core import create_presentation

    output_path = tmp_path / "sample.odp"
    create_presentation(str(output_path))

    assert output_path.exists()
    assert output_path.stat().st_size > 0


def test_create_presentation_is_valid_odp(tmp_path):
    from impress.core import create_presentation

    output_path = tmp_path / "valid.odp"
    create_presentation(str(output_path))

    assert zipfile.is_zipfile(output_path)

    with zipfile.ZipFile(output_path) as zf:
        assert "content.xml" in zf.namelist()


def test_get_slide_count_returns_one_for_new_presentation(tmp_path):
    from impress.core import (
        create_presentation,
        get_slide_count,
    )

    path = tmp_path / "count.odp"
    create_presentation(str(path))

    assert get_slide_count(str(path)) == 1


def test_export_presentation_pdf(tmp_path):
    from impress.core import (
        create_presentation,
        export_presentation,
    )

    path = tmp_path / "export.odp"
    output = tmp_path / "export.pdf"
    create_presentation(str(path))
    export_presentation(str(path), str(output), "pdf")

    assert output.exists()
    assert output.stat().st_size > 0
    # Verify it is a valid PDF by checking magic bytes
    with open(output, "rb") as f:
        assert f.read(5) == b"%PDF-"


def test_export_presentation_pptx(tmp_path):
    from impress.core import (
        create_presentation,
        export_presentation,
    )

    path = tmp_path / "export.odp"
    output = tmp_path / "export.pptx"
    create_presentation(str(path))
    export_presentation(str(path), str(output), "pptx")

    assert output.exists()
    assert output.stat().st_size > 0
    # PPTX is a ZIP file containing [Content_Types].xml
    import zipfile

    assert zipfile.is_zipfile(output)
    with zipfile.ZipFile(output) as zf:
        assert "[Content_Types].xml" in zf.namelist()


def test_export_presentation_rejects_unknown_format(tmp_path):
    from impress.core import (
        create_presentation,
        export_presentation,
    )
    from impress.exceptions import ImpressSkillError

    path = tmp_path / "export.odp"
    create_presentation(str(path))

    with pytest.raises(ImpressSkillError):
        export_presentation(str(path), str(tmp_path / "out.xyz"), "xyz")
