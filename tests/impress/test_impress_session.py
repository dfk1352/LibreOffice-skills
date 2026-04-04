# pyright: reportMissingImports=false, reportAttributeAccessIssue=false

from __future__ import annotations

import pytest

from tests.impress._helpers import (
    create_test_audio,
    create_test_image,
    create_test_video,
    get_shape_text,
    get_slide_shapes,
)


@pytest.fixture(scope="module")
def closed_impress_session(tmp_path_factory):
    """An ImpressSession that has been opened and immediately closed.

    Module-scoped so a single LibreOffice round-trip is shared across all
    parametrized ``test_closed_session_methods_raise_impress_session_error``
    variants. Each variant only checks that calling a method on the
    already-closed session raises ``ImpressSessionError`` -- no UNO access
    is needed for that assertion.

    Returns a ``(session, assets, tmp_path)`` tuple so parametrized lambdas
    that reference asset paths still work.
    """
    from impress import ImpressSession
    from impress.core import create_presentation

    base = tmp_path_factory.mktemp("closed_impress")
    doc_path = base / "closed.odp"
    create_presentation(str(doc_path))

    template_path = base / "template.odp"
    create_presentation(str(template_path))
    assets = {
        "image": create_test_image(base / "logo.png"),
        "audio": create_test_audio(base / "sample.wav"),
        "video": create_test_video(base / "sample.mp4"),
        "template": template_path,
    }

    session = ImpressSession(str(doc_path))
    session.close()
    return session, assets, base


def test_impress_session_returns_session(tmp_path):
    from impress import ImpressSession
    from impress.core import create_presentation

    doc_path = tmp_path / "session.odp"
    create_presentation(str(doc_path))

    session = ImpressSession(str(doc_path))
    try:
        assert isinstance(session, ImpressSession)
    finally:
        session.close()


def test_impress_session_missing_path_raises(tmp_path):
    from impress import ImpressSession
    from impress.exceptions import DocumentNotFoundError

    with pytest.raises(DocumentNotFoundError):
        ImpressSession(str(tmp_path / "missing.odp"))


def test_session_close_save_true_persists_changes(tmp_path):
    from impress import ImpressSession, ImpressTarget, ShapePlacement
    from impress.core import create_presentation

    doc_path = tmp_path / "persist.odp"
    create_presentation(str(doc_path))

    session = ImpressSession(str(doc_path))
    session.insert_text_box(
        ImpressTarget(kind="slide", slide_index=0),
        "Saved once",
        ShapePlacement(1.0, 1.0, 8.0, 3.0),
        name="Persisted Box",
    )
    session.close(save=True)

    assert get_shape_text(doc_path, 0, name="Persisted Box") == "Saved once"


def test_session_close_save_false_discards_changes(tmp_path):
    from impress import ImpressSession, ImpressTarget, ShapePlacement
    from impress.core import create_presentation

    doc_path = tmp_path / "discard.odp"
    create_presentation(str(doc_path))

    with ImpressSession(str(doc_path)) as initial_session:
        initial_session.insert_text_box(
            ImpressTarget(kind="slide", slide_index=0),
            "Keep me",
            ShapePlacement(1.0, 1.0, 8.0, 3.0),
            name="Persisted Box",
        )

    session = ImpressSession(str(doc_path))
    session.insert_text_box(
        ImpressTarget(kind="slide", slide_index=0),
        "Discard me",
        ShapePlacement(1.0, 5.0, 8.0, 3.0),
        name="Transient Box",
    )
    session.close(save=False)

    texts = [shape["text"] for shape in get_slide_shapes(doc_path, 0)]
    assert "Keep me" in texts
    assert "Discard me" not in texts


def test_session_reset_discards_in_memory_changes_and_reloads_saved_state(tmp_path):
    from impress import ImpressSession, ImpressTarget, ShapePlacement
    from impress.core import create_presentation

    doc_path = tmp_path / "reset.odp"
    create_presentation(str(doc_path))

    with ImpressSession(str(doc_path)) as initial_session:
        initial_session.insert_text_box(
            ImpressTarget(kind="slide", slide_index=0),
            "Persisted",
            ShapePlacement(1.0, 1.0, 8.0, 3.0),
            name="Persisted Box",
        )

    with ImpressSession(str(doc_path)) as session:
        session.insert_text_box(
            ImpressTarget(kind="slide", slide_index=0),
            "Draft",
            ShapePlacement(1.0, 5.0, 8.0, 3.0),
            name="Draft Box",
        )
        session.reset()

    texts = [shape["text"] for shape in get_slide_shapes(doc_path, 0)]
    assert "Persisted" in texts
    assert "Draft" not in texts


def test_session_restore_snapshot_marks_closed_on_reopen_failure(tmp_path):
    from unittest.mock import patch as mock_patch

    from impress import ImpressSession
    from impress.core import create_presentation
    from impress.exceptions import ImpressSessionError

    doc_path = tmp_path / "restore_fail.odp"
    create_presentation(str(doc_path))

    original_bytes = doc_path.read_bytes()
    session = ImpressSession(str(doc_path))

    with mock_patch.object(
        ImpressSession, "_open_document", side_effect=RuntimeError("simulated failure")
    ):
        with pytest.raises(RuntimeError, match="simulated failure"):
            session.restore_snapshot(original_bytes)

    with pytest.raises(ImpressSessionError):
        session.get_slide_count()


@pytest.mark.parametrize(
    ("label", "call"),
    [
        (
            "get_slide_count",
            lambda session, assets, tmp_path: session.get_slide_count(),
        ),
        (
            "get_slide_inventory",
            lambda session, assets, tmp_path: session.get_slide_inventory(
                _slide_target()
            ),
        ),
        (
            "add_slide",
            lambda session, assets, tmp_path: session.add_slide(layout="BLANK"),
        ),
        (
            "delete_slide",
            lambda session, assets, tmp_path: session.delete_slide(_slide_target()),
        ),
        (
            "move_slide",
            lambda session, assets, tmp_path: session.move_slide(_slide_target(), 0),
        ),
        (
            "duplicate_slide",
            lambda session, assets, tmp_path: session.duplicate_slide(_slide_target()),
        ),
        (
            "read_text",
            lambda session, assets, tmp_path: session.read_text(_text_target()),
        ),
        (
            "insert_text",
            lambda session, assets, tmp_path: session.insert_text(
                "after close", _insertion_target()
            ),
        ),
        (
            "replace_text",
            lambda session, assets, tmp_path: session.replace_text(
                _text_target(),
                "updated",
            ),
        ),
        (
            "delete_item",
            lambda session, assets, tmp_path: session.delete_item(_text_target()),
        ),
        (
            "format_text",
            lambda session, assets, tmp_path: session.format_text(
                _text_target(),
                _formatting(),
            ),
        ),
        (
            "insert_list",
            lambda session, assets, tmp_path: session.insert_list(
                _list_items(),
                ordered=False,
                target=_insertion_target(),
            ),
        ),
        (
            "replace_list",
            lambda session, assets, tmp_path: session.replace_list(
                _list_target(),
                _list_items(),
                ordered=True,
            ),
        ),
        (
            "delete_item",
            lambda session, assets, tmp_path: session.delete_item(_list_target()),
        ),
        (
            "insert_text_box",
            lambda session, assets, tmp_path: session.insert_text_box(
                _slide_target(),
                "box",
                _placement(),
                name="Box",
            ),
        ),
        (
            "insert_shape",
            lambda session, assets, tmp_path: session.insert_shape(
                _slide_target(),
                "rectangle",
                _placement(),
                name="Shape",
            ),
        ),
        (
            "delete_item",
            lambda session, assets, tmp_path: session.delete_item(_shape_target()),
        ),
        (
            "insert_image",
            lambda session, assets, tmp_path: session.insert_image(
                _slide_target(),
                str(assets["image"]),
                _placement(),
                name="Logo",
            ),
        ),
        (
            "replace_image",
            lambda session, assets, tmp_path: session.replace_image(
                _image_target(),
                image_path=str(assets["image"]),
                placement=_placement(),
            ),
        ),
        (
            "delete_item",
            lambda session, assets, tmp_path: session.delete_item(_image_target()),
        ),
        (
            "insert_table",
            lambda session, assets, tmp_path: session.insert_table(
                _slide_target(),
                2,
                2,
                _placement(),
                data=[["A", "B"], ["1", "2"]],
                name="Table",
            ),
        ),
        (
            "update_table",
            lambda session, assets, tmp_path: session.update_table(
                _table_target(),
                [["A", "B"], ["1", "2"]],
            ),
        ),
        (
            "delete_item",
            lambda session, assets, tmp_path: session.delete_item(_table_target()),
        ),
        (
            "insert_chart",
            lambda session, assets, tmp_path: session.insert_chart(
                _slide_target(),
                "bar",
                [["Category", "Value"], ["A", 10]],
                _placement(),
                title="Chart",
                name="Chart",
            ),
        ),
        (
            "update_chart",
            lambda session, assets, tmp_path: session.update_chart(
                _chart_target(),
                chart_type="line",
                data=[["Category", "Value"], ["A", 11]],
                placement=_placement(),
                title="Updated",
            ),
        ),
        (
            "delete_item",
            lambda session, assets, tmp_path: session.delete_item(_chart_target()),
        ),
        (
            "insert_media",
            lambda session, assets, tmp_path: session.insert_media(
                _slide_target(),
                str(assets["video"]),
                _placement(),
                name="Media",
            ),
        ),
        (
            "replace_media",
            lambda session, assets, tmp_path: session.replace_media(
                _media_target(),
                media_path=str(assets["audio"]),
                placement=_placement(),
            ),
        ),
        (
            "delete_item",
            lambda session, assets, tmp_path: session.delete_item(_media_target()),
        ),
        (
            "set_notes",
            lambda session, assets, tmp_path: session.set_notes(
                _notes_target(), "Notes"
            ),
        ),
        (
            "get_notes",
            lambda session, assets, tmp_path: session.get_notes(_notes_target()),
        ),
        (
            "list_master_pages",
            lambda session, assets, tmp_path: session.list_master_pages(),
        ),
        (
            "apply_master_page",
            lambda session, assets, tmp_path: session.apply_master_page(
                _slide_target(),
                _master_target(),
            ),
        ),
        (
            "set_master_background",
            lambda session, assets, tmp_path: session.set_master_background(
                _master_target(),
                "navy",
            ),
        ),
        (
            "import_master_page",
            lambda session, assets, tmp_path: session.import_master_page(
                str(assets["template"])
            ),
        ),
        (
            "export",
            lambda session, assets, tmp_path: session.export(
                str(tmp_path / "closed.pdf"),
                "pdf",
            ),
        ),
        (
            "patch",
            lambda session, assets, tmp_path: session.patch(
                "[operation]\ntype = add_slide\n"
            ),
        ),
        ("reset", lambda session, assets, tmp_path: session.reset()),
    ],
)
def test_closed_session_methods_raise_impress_session_error(
    closed_impress_session,
    label,
    call,
):
    from impress.exceptions import ImpressSessionError

    session, assets, base_tmp = closed_impress_session

    with pytest.raises(ImpressSessionError):
        call(session, assets, base_tmp)


def test_session_close_twice_is_idempotent(tmp_path):
    from impress import ImpressSession
    from impress.core import create_presentation

    doc_path = tmp_path / "double_close.odp"
    create_presentation(str(doc_path))

    session = ImpressSession(str(doc_path))
    session.close()
    session.close()  # second close is a no-op, no exception


def test_impress_session_context_manager_closes_after_block(tmp_path):
    from impress import ImpressSession, ImpressTarget, ShapePlacement
    from impress.core import create_presentation
    from impress.exceptions import ImpressSessionError

    doc_path = tmp_path / "context.odp"
    create_presentation(str(doc_path))

    with ImpressSession(str(doc_path)) as session:
        session.insert_text_box(
            ImpressTarget(kind="slide", slide_index=0),
            "Context managed",
            ShapePlacement(1.0, 1.0, 8.0, 3.0),
            name="Context Box",
        )

    with pytest.raises(ImpressSessionError):
        session.get_slide_count()

    assert get_shape_text(doc_path, 0, name="Context Box") == "Context managed"


def _slide_target():
    from impress import ImpressTarget

    return ImpressTarget(kind="slide", slide_index=0)


def _text_target():
    from impress import ImpressTarget

    return ImpressTarget(kind="text", slide_index=0, shape_name="Box", text="seed")


def _insertion_target():
    from impress import ImpressTarget

    return ImpressTarget(kind="insertion", slide_index=0, shape_name="Box")


def _list_target():
    from impress import ImpressTarget

    return ImpressTarget(kind="list", slide_index=0, shape_name="Box", text="Item")


def _shape_target():
    from impress import ImpressTarget

    return ImpressTarget(kind="shape", slide_index=0, shape_name="Shape")


def _image_target():
    from impress import ImpressTarget

    return ImpressTarget(kind="image", slide_index=0, shape_name="Logo")


def _table_target():
    from impress import ImpressTarget

    return ImpressTarget(kind="table", slide_index=0, shape_name="Table")


def _chart_target():
    from impress import ImpressTarget

    return ImpressTarget(kind="chart", slide_index=0, shape_name="Chart")


def _media_target():
    from impress import ImpressTarget

    return ImpressTarget(kind="media", slide_index=0, shape_name="Media")


def _notes_target():
    from impress import ImpressTarget

    return ImpressTarget(kind="notes", slide_index=0)


def _master_target():
    from impress import ImpressTarget

    return ImpressTarget(kind="master_page", master_name="Default")


def _placement():
    from impress import ShapePlacement

    return ShapePlacement(1.0, 1.0, 8.0, 3.0)


def _formatting():
    from impress import TextFormatting

    return TextFormatting(bold=True)


def _list_items():
    from impress import ListItem

    return [ListItem(text="seed", level=0)]


# --- import_master_page content tests (#5) ---


def test_import_master_page_copies_background_from_template(tmp_path):
    """import_master_page must copy visual content, not just the name (#5)."""
    from impress import ImpressSession
    from impress.core import create_presentation

    # Create a template with a coloured background on its master page
    template_path = tmp_path / "template.odp"
    create_presentation(str(template_path))

    with ImpressSession(str(template_path)) as tmpl_session:
        import uno  # type: ignore[import-not-found]

        master = tmpl_session.doc.MasterPages.getByIndex(0)
        bg = master.Background
        bg.setPropertyValue(
            "FillStyle",
            uno.Enum("com.sun.star.drawing.FillStyle", "SOLID"),
        )
        bg.setPropertyValue("FillColor", 0xFF0000)  # Red

    # Create a target presentation and import the master page
    doc_path = tmp_path / "target.odp"
    create_presentation(str(doc_path))

    with ImpressSession(str(doc_path)) as session:
        result = session.import_master_page(str(template_path))
        imported_name = result.master_name
        # Find the imported master page and verify it has the red background
        masters = session.doc.MasterPages
        imported_master = None
        for i in range(masters.Count):
            if str(masters.getByIndex(i).Name) == imported_name:
                imported_master = masters.getByIndex(i)
                break

        assert imported_master is not None
        bg = imported_master.Background
        fill_color = bg.getPropertyValue("FillColor")
        assert fill_color == 0xFF0000, (
            f"Expected red background (0xFF0000), got {fill_color:#x}"
        )


def test_import_master_page_raises_when_shape_copy_fails(tmp_path, monkeypatch):
    from unittest.mock import patch as mock_patch

    from impress import ImpressSession
    from impress.exceptions import ImpressSkillError

    class _Shape:
        ShapeType = "com.sun.star.drawing.TextShape"
        Size = object()
        Position = object()

    class _Master:
        Name = "Template Master"
        Count = 1
        Background = None

        def getByIndex(self, index):
            return _Shape()

    class _Masters:
        Count = 0

        def insertNewByIndex(self, index):
            return type("TargetMaster", (), {"Name": "", "Background": None})()

        def getByIndex(self, index):
            raise AssertionError("unexpected")

    class _TemplateDoc:
        MasterPages = type(
            "TemplateMasters", (), {"getByIndex": lambda self, index: _Master()}
        )()

        def close(self, value):
            return None

    session = ImpressSession.__new__(ImpressSession)
    session._closed = False
    session._doc = type(
        "Doc",
        (),
        {
            "MasterPages": _Masters(),
            "createInstance": lambda self, service: (_ for _ in ()).throw(
                RuntimeError("synthetic create failure")
            ),
        },
    )()
    session._desktop = type(
        "Desktop", (), {"loadComponentFromURL": lambda self, *args: _TemplateDoc()}
    )()

    with mock_patch("impress.session.Path.exists", return_value=True):
        with pytest.raises(ImpressSkillError, match="shape 0 copy failed"):
            session.import_master_page(str(tmp_path / "template_fail.odp"))


def test_import_master_page_returns_structural_result_on_partial_failure(tmp_path):
    from unittest.mock import patch as mock_patch
    from impress import ImpressSession
    from impress.core import create_presentation

    doc_path = tmp_path / "target_warn.odp"
    create_presentation(str(doc_path))
    template_path = tmp_path / "template_warn.odp"
    create_presentation(str(template_path))

    with ImpressSession(str(doc_path)) as session:
        # Mock getPropertyValue on the source background to raise an exception
        # which triggers the warning code path
        original_getattr = getattr

        def fake_getattr(obj, name, default=None):
            if (
                name == "Background"
                and hasattr(obj, "Name")
                and str(obj.Name) == "Default"
            ):

                class FakeBg:
                    def getPropertyValue(self, prop):
                        if prop == "FillColor":
                            raise RuntimeError("Simulated property error")
                        return 0

                return FakeBg()
            return original_getattr(obj, name, default)

        with mock_patch("impress.session.getattr", side_effect=fake_getattr):
            result = session.import_master_page(str(template_path))

        assert result.master_name == "Default"
        assert result.copied_shape_count >= 0
        assert len(result.warnings) > 0
        assert any("Simulated property error" in w for w in result.warnings)


def test_apply_master_page_can_target_a_single_slide(tmp_path):
    from impress import ImpressSession, ImpressTarget
    from impress.core import create_presentation

    doc_path = tmp_path / "single_slide_master.odp"
    create_presentation(str(doc_path))

    with ImpressSession(str(doc_path)) as session:
        session.add_slide()
        session.add_slide()
        custom = session.doc.MasterPages.insertNewByIndex(session.doc.MasterPages.Count)
        custom.Name = "CustomMaster"
        import uno

        background = custom.Background
        background.setPropertyValue(
            "FillStyle", uno.Enum("com.sun.star.drawing.FillStyle", "SOLID")
        )
        background.setPropertyValue("FillColor", 0x123456)

        session.apply_master_page(
            ImpressTarget(kind="slide", slide_index=1),
            ImpressTarget(kind="master_page", master_name="CustomMaster"),
        )

        masters = [
            str(session.doc.DrawPages.getByIndex(index).MasterPage.Name)
            for index in range(session.doc.DrawPages.Count)
        ]

    assert masters == ["Default", "CustomMaster", "Default"]


# --- move_slide fidelity tests (#6) ---


def test_move_slide_preserves_text_and_formatting(tmp_path):
    """move_slide must preserve text content and formatting of shapes (#6)."""
    from impress import ImpressSession, ImpressTarget, ShapePlacement
    from impress.core import create_presentation

    doc_path = tmp_path / "move_fidelity.odp"
    create_presentation(str(doc_path))

    with ImpressSession(str(doc_path)) as session:
        # Add a second and third slide
        session.add_slide()
        session.add_slide()

        # Insert a text box on slide 0 and apply bold + red via UNO
        session.insert_text_box(
            ImpressTarget(kind="slide", slide_index=0),
            "Bold Red Text",
            ShapePlacement(2.0, 2.0, 10.0, 3.0),
            name="BoldBox",
        )
        # Apply formatting directly via UNO
        slide = session.doc.DrawPages.getByIndex(0)
        for i in range(slide.Count):
            shape = slide.getByIndex(i)
            try:
                if shape.getString() == "Bold Red Text":
                    cursor = shape.Text.createTextCursor()
                    cursor.gotoStart(False)
                    cursor.gotoEnd(True)
                    from com.sun.star.awt.FontWeight import BOLD  # type: ignore

                    cursor.CharWeight = BOLD
                    cursor.CharColor = 0xFF0000
                    break
            except Exception:
                continue

    # Move slide 0 to position 2
    with ImpressSession(str(doc_path)) as session:
        session.move_slide(
            ImpressTarget(kind="slide", slide_index=0),
            to_index=2,
        )

    # Verify the moved slide (now at index 2) has the text with formatting
    with ImpressSession(str(doc_path)) as session:
        slide = session.doc.DrawPages.getByIndex(2)
        found_text = False
        for i in range(slide.Count):
            shape = slide.getByIndex(i)
            try:
                text = shape.getString()
            except Exception:
                continue
            if "Bold Red Text" in text:
                found_text = True
                # Check that character formatting survived the move
                cursor = shape.Text.createTextCursor()
                cursor.gotoStart(False)
                cursor.gotoEnd(True)
                from com.sun.star.awt.FontWeight import BOLD  # type: ignore

                assert cursor.CharWeight == BOLD, (
                    f"Expected bold, got CharWeight={cursor.CharWeight}"
                )
                assert cursor.CharColor == 0xFF0000, (
                    f"Expected red (0xFF0000), got {cursor.CharColor:#x}"
                )
                break

        assert found_text, "Moved slide should contain 'Bold Red Text'"


def test_move_slide_preserves_table_data(tmp_path):
    """move_slide must preserve table shape data (#6)."""
    from impress import ImpressSession, ImpressTarget, ShapePlacement
    from impress.core import create_presentation

    doc_path = tmp_path / "move_table.odp"
    create_presentation(str(doc_path))

    with ImpressSession(str(doc_path)) as session:
        session.add_slide()
        # Insert a table on slide 0
        session.insert_table(
            ImpressTarget(kind="slide", slide_index=0),
            2,
            2,
            ShapePlacement(2.0, 2.0, 12.0, 5.0),
            data=[["A", "B"], ["C", "D"]],
            name="TestTable",
        )

    # Move slide 0 to position 1
    with ImpressSession(str(doc_path)) as session:
        session.move_slide(
            ImpressTarget(kind="slide", slide_index=0),
            to_index=1,
        )

    # Verify the moved slide has a table with correct data
    with ImpressSession(str(doc_path)) as session:
        slide = session.doc.DrawPages.getByIndex(1)
        table_found = False
        for i in range(slide.Count):
            shape = slide.getByIndex(i)
            shape_type = str(getattr(shape, "ShapeType", ""))
            if shape_type == "com.sun.star.drawing.TableShape":
                table_found = True
                model = shape.Model
                assert model.Rows.Count == 2
                assert model.Columns.Count == 2
                break

        assert table_found, "Moved slide should contain a table shape"


def test_move_slide_preserves_chart_data(tmp_path):
    from impress import ImpressSession, ImpressTarget, ShapePlacement
    from impress.core import create_presentation
    from tests.impress._helpers import get_chart_details

    doc_path = tmp_path / "move_chart.odp"
    create_presentation(str(doc_path))

    with ImpressSession(str(doc_path)) as session:
        session.add_slide()
        session.insert_chart(
            ImpressTarget(kind="slide", slide_index=0),
            "bar",
            [["Category", "North", "South"], ["Q1", 10, 20], ["Q2", 30, 40]],
            ShapePlacement(2.0, 2.0, 10.0, 6.0),
            title="Regional Sales",
            name="Sales Chart",
        )

    before = get_chart_details(doc_path, 0, name="Sales Chart")

    with ImpressSession(str(doc_path)) as session:
        session.move_slide(ImpressTarget(kind="slide", slide_index=0), to_index=1)

    after = get_chart_details(doc_path, 1, name="Sales Chart")

    assert after["title"] == before["title"]
    assert after["column_descriptions"] == before["column_descriptions"]
    assert after["row_descriptions"] == before["row_descriptions"]
    assert after["data"] == before["data"]
