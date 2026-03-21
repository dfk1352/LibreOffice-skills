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


def test_session_close_twice_raises_impress_session_error(tmp_path):
    from impress import ImpressSession
    from impress.core import create_presentation
    from impress.exceptions import ImpressSessionError

    doc_path = tmp_path / "double_close.odp"
    create_presentation(str(doc_path))

    session = ImpressSession(str(doc_path))
    session.close()

    with pytest.raises(ImpressSessionError):
        session.close()


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
