# pyright: reportMissingImports=false, reportAttributeAccessIssue=false

from __future__ import annotations

from tests.impress._helpers import (
    create_test_audio,
    create_test_image,
    create_test_video,
    get_chart_details,
    get_media_url,
    get_shape_geometry,
    get_shape_text,
    get_slide_shapes,
    get_table_matrix,
)


def test_image_replacement_persists_target_identity_and_geometry_after_reopen(tmp_path):
    from impress import ImpressSession, ImpressTarget, ShapePlacement
    from impress.core import create_presentation

    doc_path = tmp_path / "image_replace_invariant.odp"
    first_image = create_test_image(tmp_path / "first.png", color="blue")
    replacement_image = create_test_image(tmp_path / "replacement.png", color="green")
    create_presentation(str(doc_path))

    with ImpressSession(str(doc_path)) as session:
        session.insert_image(
            ImpressTarget(kind="slide", slide_index=0),
            str(first_image),
            ShapePlacement(2.0, 2.0, 4.0, 4.0),
            name="Logo",
        )
        session.replace_image(
            ImpressTarget(kind="image", slide_index=0, shape_name="Logo"),
            image_path=str(replacement_image),
            placement=ShapePlacement(3.0, 2.5, 5.0, 5.0),
        )

    assert "replacement" in get_media_url(doc_path, 0, name="Logo")
    assert get_shape_geometry(doc_path, 0, name="Logo") == {
        "x": 3000,
        "y": 2500,
        "width": 5000,
        "height": 5000,
    }


def test_move_and_duplicate_preserve_visible_object_traces_after_reopen(tmp_path):
    from impress import ImpressSession, ImpressTarget, ShapePlacement
    from impress.core import create_presentation

    doc_path = tmp_path / "move_duplicate_invariants.odp"
    image_path = create_test_image(tmp_path / "trace.png", color="purple")
    video_path = create_test_video(tmp_path / "trace.mp4")
    audio_path = create_test_audio(tmp_path / "trace.wav")
    create_presentation(str(doc_path))

    with ImpressSession(str(doc_path)) as session:
        session.add_slide(layout="BLANK")
        slide = ImpressTarget(kind="slide", slide_index=1)
        session.insert_text_box(
            slide,
            "Quarterly checkpoint",
            ShapePlacement(1.0, 1.0, 8.0, 2.5),
            name="Trace Box",
        )
        session.insert_image(
            slide,
            str(image_path),
            ShapePlacement(10.0, 1.0, 3.5, 3.5),
            name="Trace Image",
        )
        session.insert_table(
            slide,
            2,
            2,
            ShapePlacement(1.0, 4.0, 6.0, 3.0),
            data=[["Metric", "Value"], ["Revenue", "120"]],
            name="Trace Table",
        )
        session.insert_chart(
            slide,
            "bar",
            [["Category", "Value"], ["Revenue", 120], ["Cost", 80]],
            ShapePlacement(8.0, 4.0, 6.0, 4.5),
            title="Trace Chart",
            name="Trace Chart",
        )
        session.insert_media(
            slide,
            str(video_path),
            ShapePlacement(1.0, 8.0, 4.0, 3.0),
            name="Trace Media",
        )
        session.replace_media(
            ImpressTarget(kind="media", slide_index=1, shape_name="Trace Media"),
            media_path=str(audio_path),
        )
        session.move_slide(ImpressTarget(kind="slide", slide_index=1), 0)
        session.duplicate_slide(ImpressTarget(kind="slide", slide_index=0))

    names = {shape["name"] for shape in get_slide_shapes(doc_path, 0)}
    assert {
        "Trace_Box",
        "Trace_Image",
        "Trace_Table",
        "Trace_Chart",
    } <= names
    assert get_shape_text(doc_path, 0, name="Trace Box") == "Quarterly checkpoint"
    assert "trace" in get_media_url(doc_path, 0, name="Trace Image")
    assert get_table_matrix(doc_path, 0, name="Trace Table") == [
        ["Metric", "Value"],
        ["Revenue", "120"],
    ]
    assert get_chart_details(doc_path, 0, name="Trace Chart")["title"] == "Trace Chart"
    assert "trace" in get_media_url(doc_path, 0, name="Trace Media")

    duplicate_names = {shape["name"] for shape in get_slide_shapes(doc_path, 1)}
    assert {
        "Trace_Box 1",
        "Trace_Image 1",
        "Trace_Table 1",
        "Trace_Chart 1",
        "Trace_Media 1",
    } <= duplicate_names
    assert get_shape_text(doc_path, 1, name="Trace Box 1") == "Quarterly checkpoint"
    assert "trace" in get_media_url(doc_path, 1, name="Trace Image 1")
    assert get_table_matrix(doc_path, 1, name="Trace Table 1") == [
        ["Metric", "Value"],
        ["Revenue", "120"],
    ]
    assert (
        get_chart_details(doc_path, 1, name="Trace Chart 1")["title"] == "Trace Chart"
    )
    assert "trace" in get_media_url(doc_path, 1, name="Trace Media 1")
