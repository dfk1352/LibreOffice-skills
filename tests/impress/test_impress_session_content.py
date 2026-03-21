# pyright: reportMissingImports=false, reportAttributeAccessIssue=false

from __future__ import annotations

from tests.impress._helpers import (
    TITLE_AND_CONTENT_LAYOUT,
    append_slide,
    create_test_audio,
    create_test_image,
    create_test_video,
    get_chart_details,
    get_list_paragraphs,
    get_master_background,
    get_media_url,
    get_notes_text,
    get_shape_geometry,
    get_shape_text,
    get_slide_master_name,
    get_slide_shapes,
    get_table_matrix,
    get_text_properties,
    open_impress_doc,
)


def test_session_slide_operations_preserve_zero_based_slide_order(tmp_path):
    from impress import ImpressSession, ImpressTarget, ShapePlacement
    from impress.core import create_presentation

    doc_path = tmp_path / "slides.odp"
    create_presentation(str(doc_path))

    with ImpressSession(str(doc_path)) as session:
        session.add_slide(layout="BLANK")
        session.add_slide(layout="BLANK")
        session.insert_text_box(
            ImpressTarget(kind="slide", slide_index=0),
            "Slide zero",
            ShapePlacement(1.0, 1.0, 8.0, 3.0),
            name="Zero Box",
        )
        session.insert_text_box(
            ImpressTarget(kind="slide", slide_index=1),
            "Slide one",
            ShapePlacement(1.0, 1.0, 8.0, 3.0),
            name="One Box",
        )
        session.insert_text_box(
            ImpressTarget(kind="slide", slide_index=2),
            "Slide two",
            ShapePlacement(1.0, 1.0, 8.0, 3.0),
            name="Two Box",
        )
        session.move_slide(ImpressTarget(kind="slide", slide_index=0), 2)
        session.duplicate_slide(ImpressTarget(kind="slide", slide_index=1))
        session.delete_slide(ImpressTarget(kind="slide", slide_index=2))
        slide_count = session.get_slide_count()

    assert slide_count == 3
    assert get_shape_text(doc_path, 0, name="One Box") == "Slide one"
    assert get_shape_text(doc_path, 1, name="Two Box") == "Slide two"
    assert get_shape_text(doc_path, 2, name="Zero Box") == "Slide zero"


def test_session_text_operations_work_for_text_boxes_and_placeholders(tmp_path):
    from impress import ImpressSession, ImpressTarget, ShapePlacement
    from impress.core import create_presentation

    doc_path = tmp_path / "text_ops.odp"
    create_presentation(str(doc_path))

    with open_impress_doc(doc_path) as doc:
        append_slide(doc, TITLE_AND_CONTENT_LAYOUT)
        doc.store()

    with ImpressSession(str(doc_path)) as session:
        session.insert_text(
            "Presentation Title",
            ImpressTarget(kind="insertion", slide_index=1, placeholder="title"),
        )
        session.insert_text(
            "Quarterly revenue rose 18%.",
            ImpressTarget(kind="insertion", slide_index=1, placeholder="body"),
        )
        session.insert_text_box(
            ImpressTarget(kind="slide", slide_index=1),
            "Agenda: confirm scope and share results",
            ShapePlacement(1.0, 7.0, 12.0, 3.0),
            name="Agenda Box",
        )
        placeholder_text = session.read_text(
            ImpressTarget(kind="text", slide_index=1, placeholder="body")
        )
        session.replace_text(
            ImpressTarget(kind="text", slide_index=1, placeholder="body"),
            "Quarterly revenue rose 21%.",
        )
        session.delete_item(
            ImpressTarget(
                kind="text",
                slide_index=1,
                shape_name="Agenda Box",
                text="share ",
            )
        )

    assert placeholder_text == "Quarterly revenue rose 18%."
    assert get_shape_text(doc_path, 1, placeholder="title") == "Presentation Title"
    assert (
        get_shape_text(doc_path, 1, placeholder="body") == "Quarterly revenue rose 21%."
    )
    assert get_shape_text(doc_path, 1, name="Agenda Box") == (
        "Agenda: confirm scope and results"
    )


def test_session_format_text_applies_character_and_paragraph_formatting(tmp_path):
    from impress import (
        ImpressSession,
        ImpressTarget,
        ShapePlacement,
        TextFormatting,
    )
    from impress.core import create_presentation

    doc_path = tmp_path / "format_text.odp"
    create_presentation(str(doc_path))

    with ImpressSession(str(doc_path)) as session:
        session.insert_text_box(
            ImpressTarget(kind="slide", slide_index=0),
            "Formatting target",
            ShapePlacement(1.0, 1.0, 10.0, 3.0),
            name="Format Box",
        )
        session.format_text(
            ImpressTarget(kind="text", slide_index=0, shape_name="Format Box"),
            TextFormatting(
                bold=True,
                italic=True,
                underline=True,
                font_name="Liberation Sans Narrow",
                font_size=18,
                color="red",
                align="center",
            ),
        )

    properties = get_text_properties(doc_path, 0, name="Format Box")
    assert properties["char_weight"] == 150
    assert properties["char_posture"] == 2
    assert properties["char_underline"] == 1
    assert properties["font_name"] == "Liberation Sans Narrow"
    assert properties["font_size"] == 18
    assert properties["color"] == 0xFF0000
    assert properties["align"] in {2, 3}


def test_session_list_operations_mutate_structural_list_content(tmp_path):
    from impress import ImpressSession, ImpressTarget, ListItem, ShapePlacement
    from impress.core import create_presentation

    doc_path = tmp_path / "lists.odp"
    create_presentation(str(doc_path))

    with ImpressSession(str(doc_path)) as session:
        session.insert_text_box(
            ImpressTarget(kind="slide", slide_index=0),
            "Action Items",
            ShapePlacement(1.0, 1.0, 12.0, 5.0),
            name="Agenda Box",
        )
        session.insert_list(
            [
                ListItem(text="Confirm scope", level=0),
                ListItem(text="Review outputs", level=0),
            ],
            ordered=True,
            target=ImpressTarget(
                kind="insertion",
                slide_index=0,
                shape_name="Agenda Box",
                after="Action Items",
            ),
        )

    after_insert = get_list_paragraphs(doc_path, 0, name="Agenda Box")
    insert_texts = [p["text"] for p in after_insert]
    assert "Confirm scope" in insert_texts
    assert "Review outputs" in insert_texts

    with ImpressSession(str(doc_path)) as session:
        session.replace_list(
            ImpressTarget(
                kind="list",
                slide_index=0,
                shape_name="Agenda Box",
                text="Confirm scope",
            ),
            [
                ListItem(text="Finalize scope", level=0),
                ListItem(text="Update notes", level=1),
            ],
        )

    after_replace = get_list_paragraphs(doc_path, 0, name="Agenda Box")
    replace_texts = [p["text"] for p in after_replace]
    assert "Finalize scope" in replace_texts
    assert "Confirm scope" not in replace_texts

    with ImpressSession(str(doc_path)) as session:
        session.delete_item(
            ImpressTarget(
                kind="list",
                slide_index=0,
                shape_name="Agenda Box",
                text="Finalize scope",
            )
        )

    assert get_list_paragraphs(doc_path, 0, name="Agenda Box") == []


def test_session_shape_and_image_operations_mutate_slide_inventory_predictably(
    tmp_path,
):
    from impress import ImpressSession, ImpressTarget, ShapePlacement
    from impress.core import create_presentation

    doc_path = tmp_path / "shapes_images.odp"
    create_presentation(str(doc_path))
    first_image = create_test_image(tmp_path / "first.png", color="blue")
    replacement_image = create_test_image(tmp_path / "replacement.png", color="green")

    with ImpressSession(str(doc_path)) as session:
        session.insert_text_box(
            ImpressTarget(kind="slide", slide_index=0),
            "Summary box",
            ShapePlacement(1.0, 1.0, 10.0, 3.0),
            name="Summary Box",
        )
        session.insert_shape(
            ImpressTarget(kind="slide", slide_index=0),
            "rectangle",
            ShapePlacement(1.0, 5.0, 4.0, 2.0),
            fill_color="navy",
            line_color="black",
            name="Accent Shape",
        )
        session.insert_image(
            ImpressTarget(kind="slide", slide_index=0),
            str(first_image),
            ShapePlacement(7.0, 1.0, 4.0, 4.0),
            name="Logo",
        )
        session.replace_image(
            ImpressTarget(kind="image", slide_index=0, shape_name="Logo"),
            image_path=str(replacement_image),
            placement=ShapePlacement(8.0, 1.5, 5.0, 5.0),
        )
        session.delete_item(
            ImpressTarget(kind="shape", slide_index=0, shape_name="Accent Shape")
        )

    names = [shape["name"] for shape in get_slide_shapes(doc_path, 0)]
    assert any(name.lower() == "summary_box" for name in names)
    assert any(name.lower() == "logo" for name in names)
    assert all(name.lower() != "accent_shape" for name in names)
    assert "replacement" in get_media_url(doc_path, 0, name="Logo")
    geometry = get_shape_geometry(doc_path, 0, name="Logo")
    assert geometry["x"] == 8000
    assert geometry["width"] == 5000


def test_session_insert_image_does_not_leak_filesystem_path(tmp_path):
    """Image Description stores only a filename stem, never the full path."""
    from impress import ImpressSession, ImpressTarget, ShapePlacement
    from impress.core import create_presentation

    doc_path = tmp_path / "path_leak.odp"
    create_presentation(str(doc_path))

    nested_dir = tmp_path / "a" / "b" / "c"
    nested_dir.mkdir(parents=True)
    nested_image = create_test_image(nested_dir / "secret_image.png", color="red")

    with ImpressSession(str(doc_path)) as session:
        session.insert_image(
            ImpressTarget(kind="slide", slide_index=0),
            str(nested_image),
            ShapePlacement(1.0, 1.0, 5.0, 5.0),
            name="Nested Image",
        )

    url = get_media_url(doc_path, 0, name="Nested Image")
    assert "/" not in url, f"Description leaks path separator: {url!r}"
    assert "\\" not in url, f"Description leaks backslash: {url!r}"
    assert "file://" not in url, f"Description leaks file URI: {url!r}"


def test_session_table_operations_mutate_table_content(tmp_path):
    from impress import ImpressSession, ImpressTarget, ShapePlacement
    from impress.core import create_presentation

    doc_path = tmp_path / "tables.odp"
    create_presentation(str(doc_path))

    with ImpressSession(str(doc_path)) as session:
        session.insert_table(
            ImpressTarget(kind="slide", slide_index=0),
            3,
            2,
            ShapePlacement(1.0, 1.0, 10.0, 5.0),
            data=[["Label", "Value"], ["Revenue", "100"], ["Cost", "80"]],
            name="Metrics Table",
        )

    after_insert = get_table_matrix(doc_path, 0, name="Metrics Table")
    assert after_insert[1] == ["Revenue", "100"]
    assert after_insert[2] == ["Cost", "80"]

    with ImpressSession(str(doc_path)) as session:
        session.update_table(
            ImpressTarget(kind="table", slide_index=0, shape_name="Metrics Table"),
            [["Label", "Value"], ["Revenue", "120"], ["Cost", "90"]],
        )

    after_update = get_table_matrix(doc_path, 0, name="Metrics Table")
    assert after_update[1] == ["Revenue", "120"]
    assert after_update[2] == ["Cost", "90"]

    with ImpressSession(str(doc_path)) as session:
        session.delete_item(
            ImpressTarget(kind="table", slide_index=0, shape_name="Metrics Table")
        )

    names = [shape["name"] for shape in get_slide_shapes(doc_path, 0)]
    assert all(name.lower() != "metrics_table" for name in names)


def test_session_chart_operations_mutate_chart_state_or_presence(tmp_path):
    from impress import ImpressSession, ImpressTarget, ShapePlacement
    from impress.core import create_presentation

    doc_path = tmp_path / "charts.odp"
    create_presentation(str(doc_path))

    with ImpressSession(str(doc_path)) as session:
        session.insert_chart(
            ImpressTarget(kind="slide", slide_index=0),
            "bar",
            [["Category", "Value"], ["Revenue", 100], ["Cost", 80]],
            ShapePlacement(1.0, 1.0, 10.0, 6.0),
            title="Revenue Trend",
            name="Revenue Chart",
        )

    created = get_chart_details(doc_path, 0, name="Revenue Chart")
    assert created["title"] == "Revenue Trend"
    assert created["width"] == 10000

    with ImpressSession(str(doc_path)) as session:
        session.update_chart(
            ImpressTarget(kind="chart", slide_index=0, shape_name="Revenue Chart"),
            chart_type="line",
            data=[["Category", "Value"], ["Revenue", 110], ["Cost", 75]],
            placement=ShapePlacement(2.0, 1.5, 12.0, 7.0),
            title="Revenue Trend Updated",
        )

    updated = get_chart_details(doc_path, 0, name="Revenue Chart")
    assert updated["title"] == "Revenue Trend Updated"
    assert updated["x"] == 2000
    assert updated["width"] == 12000

    with ImpressSession(str(doc_path)) as session:
        session.delete_item(
            ImpressTarget(kind="chart", slide_index=0, shape_name="Revenue Chart")
        )

    names = [shape["name"] for shape in get_slide_shapes(doc_path, 0)]
    assert all(name.lower() != "revenue_chart" for name in names)


def test_session_media_operations_mutate_media_objects(tmp_path):
    from impress import ImpressSession, ImpressTarget, ShapePlacement
    from impress.core import create_presentation

    doc_path = tmp_path / "media.odp"
    create_presentation(str(doc_path))
    first_video = create_test_video(tmp_path / "first.mp4")
    replacement_audio = create_test_audio(tmp_path / "replacement.wav")

    with ImpressSession(str(doc_path)) as session:
        session.insert_media(
            ImpressTarget(kind="slide", slide_index=0),
            str(first_video),
            ShapePlacement(1.0, 1.0, 5.0, 5.0),
            name="Demo Media",
        )
        session.replace_media(
            ImpressTarget(kind="media", slide_index=0, shape_name="Demo Media"),
            media_path=str(replacement_audio),
            placement=ShapePlacement(2.0, 2.0, 6.0, 4.0),
        )
        session.delete_item(
            ImpressTarget(kind="media", slide_index=0, shape_name="Demo Media")
        )

    names = [shape["name"] for shape in get_slide_shapes(doc_path, 0)]
    assert all(name.lower() != "demo_media" for name in names)


def test_session_notes_and_master_page_operations_are_persisted(tmp_path):
    from impress import ImpressSession, ImpressTarget
    from impress.core import create_presentation

    doc_path = tmp_path / "notes_master.odp"
    create_presentation(str(doc_path))

    with ImpressSession(str(doc_path)) as session:
        session.add_slide(layout="BLANK")
        masters = session.list_master_pages()
        session.set_notes(ImpressTarget(kind="notes", slide_index=1), "Speaker note")
        session.apply_master_page(
            ImpressTarget(kind="slide", slide_index=1),
            ImpressTarget(kind="master_page", master_name=masters[0]),
        )
        session.set_master_background(
            ImpressTarget(kind="master_page", master_name=masters[0]),
            "lightsteelblue",
        )
        read_back = session.get_notes(ImpressTarget(kind="notes", slide_index=1))

    assert read_back == "Speaker note"
    assert get_notes_text(doc_path, 1) == "Speaker note"
    master_name = get_slide_master_name(doc_path, 1)
    assert master_name
    assert get_master_background(doc_path, master_name) == 0xB0C4DE


def test_session_export_writes_alternate_presentation_formats(tmp_path):
    from impress import ImpressSession, ImpressTarget, ShapePlacement
    from impress.core import create_presentation

    doc_path = tmp_path / "export_session.odp"
    pdf_path = tmp_path / "export_session.pdf"
    pptx_path = tmp_path / "export_session.pptx"
    create_presentation(str(doc_path))

    with ImpressSession(str(doc_path)) as session:
        session.insert_text_box(
            ImpressTarget(kind="slide", slide_index=0),
            "Export me",
            ShapePlacement(1.0, 1.0, 10.0, 3.0),
            name="Export Box",
        )
        session.export(str(pdf_path), "pdf")
        session.export(str(pptx_path), "pptx")

    assert pdf_path.exists()
    assert pdf_path.stat().st_size > 0
    with open(pdf_path, "rb") as handle:
        assert handle.read(5) == b"%PDF-"
    assert pptx_path.exists()
    assert pptx_path.stat().st_size > 0
