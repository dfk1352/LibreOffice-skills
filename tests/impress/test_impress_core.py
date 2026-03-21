# pyright: reportMissingImports=false

from __future__ import annotations

import zipfile

import pytest


def test_create_presentation_creates_valid_odp_file(tmp_path):
    from impress.core import create_presentation

    output_path = tmp_path / "sample.odp"
    create_presentation(str(output_path))

    assert output_path.exists()
    assert output_path.stat().st_size > 0
    assert zipfile.is_zipfile(output_path)

    with zipfile.ZipFile(output_path) as archive:
        assert "content.xml" in archive.namelist()


def test_export_presentation_missing_doc_raises(tmp_path):
    from impress.core import export_presentation
    from impress.exceptions import DocumentNotFoundError

    with pytest.raises(DocumentNotFoundError):
        export_presentation(
            str(tmp_path / "missing.odp"),
            str(tmp_path / "out.pdf"),
            "pdf",
        )


def test_export_presentation_pdf(tmp_path):
    from impress import ImpressSession, ImpressTarget, ShapePlacement
    from impress.core import create_presentation, export_presentation

    path = tmp_path / "export_pdf.odp"
    output = tmp_path / "export_pdf.pdf"
    create_presentation(str(path))

    with ImpressSession(str(path)) as session:
        session.insert_text_box(
            ImpressTarget(kind="slide", slide_index=0),
            "Export me",
            ShapePlacement(1.0, 1.0, 10.0, 3.0),
            name="ExportBox",
        )

    export_presentation(str(path), str(output), "pdf")

    assert output.exists()
    assert output.stat().st_size > 0
    with open(output, "rb") as handle:
        assert handle.read(5) == b"%PDF-"


def test_export_presentation_pptx(tmp_path):
    from impress.core import create_presentation, export_presentation

    path = tmp_path / "export_pptx.odp"
    output = tmp_path / "export_pptx.pptx"
    create_presentation(str(path))
    export_presentation(str(path), str(output), "pptx")

    assert output.exists()
    assert output.stat().st_size > 0
    assert zipfile.is_zipfile(output)

    with zipfile.ZipFile(output) as archive:
        assert "[Content_Types].xml" in archive.namelist()


def test_export_presentation_rejects_unknown_format(tmp_path):
    from impress.core import create_presentation, export_presentation
    from impress.exceptions import ImpressSkillError

    path = tmp_path / "export_unknown.odp"
    create_presentation(str(path))

    with pytest.raises(ImpressSkillError):
        export_presentation(str(path), str(tmp_path / "out.xyz"), "xyz")
