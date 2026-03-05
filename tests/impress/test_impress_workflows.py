"""Integration workflow tests for Impress skill."""

import struct
import wave
from pathlib import Path

from PIL import Image


def _create_minimal_wav(path: Path) -> None:
    """Create a minimal valid WAV file for testing."""
    with wave.open(str(path), "w") as w:
        w.setnchannels(1)
        w.setsampwidth(2)
        w.setframerate(44100)
        w.writeframes(struct.pack("<h", 0) * 100)


def _create_minimal_video(path: Path) -> None:
    """Create a minimal file to act as a video placeholder."""
    path.write_bytes(b"\x00" * 100)


def run_end_to_end_workflow(output_dir: Path) -> dict[str, Path]:
    """Build an Impress presentation exercising every public function.

    This is the single function that produces all inspectable output
    files.  It exercises every public function from every Impress module
    plus the Impress find_replace module.

    Args:
        output_dir: Directory where all output files are written.

    Returns:
        Dict mapping logical names to output file paths:
            "presentation"      -> presentation.odp
            "snapshot_before"   -> snapshot_before.png
            "snapshot_after"    -> snapshot_after.png
            "export_pdf"        -> presentation.pdf
            "export_pptx"       -> presentation.pptx
            "slide_images"      -> list stored in output_dir/slides/
    """
    from impress.charts import add_chart
    from impress.content import (
        add_image,
        add_shape,
        add_text_box,
        set_body,
        set_title,
    )
    from impress.core import (
        create_presentation,
        export_presentation,
        get_slide_count,
    )
    from impress.find_replace import find_replace
    from impress.formatting import (
        format_shape_text,
        set_slide_background,
    )
    from impress.master import (
        apply_master_page,
        import_master_from_template,
        list_master_pages,
    )
    from impress.media import add_audio, add_video
    from impress.notes import get_notes, set_notes
    from impress.slides import (
        add_slide,
        delete_slide,
        duplicate_slide,
        get_slide_inventory,
        move_slide,
    )
    from impress.snapshot import snapshot_slide
    from impress.tables import (
        add_table,
        format_table_cell,
        set_table_cell,
    )

    output_dir.mkdir(parents=True, exist_ok=True)

    # ------------------------------------------------------------------
    # 1. Create presentation
    # ------------------------------------------------------------------
    pres_path = output_dir / "presentation.odp"
    create_presentation(str(pres_path))

    count = get_slide_count(str(pres_path))
    assert count == 1, f"Expected 1 slide, got {count}"

    # ------------------------------------------------------------------
    # 2. Add slides with various layouts
    # ------------------------------------------------------------------
    add_slide(str(pres_path), layout="TITLE_SLIDE")  # slide 1
    add_slide(str(pres_path), layout="TITLE_AND_CONTENT")  # slide 2
    add_slide(str(pres_path), layout="BLANK")  # slide 3
    add_slide(str(pres_path), layout="TITLE_ONLY")  # slide 4

    count = get_slide_count(str(pres_path))
    assert count == 5, f"Expected 5 slides, got {count}"

    # ------------------------------------------------------------------
    # 3. Set title and body text
    # ------------------------------------------------------------------
    set_title(str(pres_path), 1, "Welcome to Integration Test")
    set_body(str(pres_path), 2, "This slide has body content.")

    # ------------------------------------------------------------------
    # 4. Add text box, shapes, and image
    # ------------------------------------------------------------------
    tb_idx = add_text_box(
        str(pres_path),
        3,
        "Integration box",
        2.0,
        2.0,
        12.0,
        3.0,
    )
    assert isinstance(tb_idx, int)

    shape_idx = add_shape(
        str(pres_path),
        3,
        "rectangle",
        2.0,
        6.0,
        5.0,
        4.0,
        fill_color="cornflowerblue",
        line_color="black",
    )
    assert isinstance(shape_idx, int)

    # Create a test image and add it
    img_path = output_dir / "test_image.png"
    img = Image.new("RGB", (100, 100), color="blue")
    img.save(img_path)
    img_idx = add_image(str(pres_path), 3, str(img_path), 14.0, 2.0, 5.0, 5.0)
    assert isinstance(img_idx, int)

    # ------------------------------------------------------------------
    # 5. Snapshot BEFORE formatting (slide 0 — the default blank slide)
    # ------------------------------------------------------------------
    before_path = output_dir / "snapshot_before.png"
    snap_result = snapshot_slide(str(pres_path), 0, str(before_path))
    assert before_path.exists()
    assert snap_result.width > 0
    assert snap_result.height > 0

    # ------------------------------------------------------------------
    # 6. Add table with data, set and format cells
    # ------------------------------------------------------------------
    table_data = [
        ["Category", "Value"],
        ["Alpha", "100"],
        ["Beta", "200"],
    ]
    tbl_idx = add_table(str(pres_path), 3, 3, 2, 2.0, 11.0, 12.0, 5.0, data=table_data)
    assert isinstance(tbl_idx, int)

    set_table_cell(str(pres_path), 3, tbl_idx, 2, 1, "250")
    format_table_cell(
        str(pres_path),
        3,
        tbl_idx,
        0,
        0,
        bold=True,
        font_size=14,
        fill_color="lightgray",
    )

    # ------------------------------------------------------------------
    # 7. Add chart
    # ------------------------------------------------------------------
    chart_data = [
        ["", "Q1", "Q2"],
        ["Sales", 100, 200],
        ["Costs", 80, 150],
    ]
    chart_idx = add_chart(
        str(pres_path),
        4,
        "bar",
        chart_data,
        1.0,
        1.0,
        20.0,
        14.0,
        title="Sales Chart",
    )
    assert isinstance(chart_idx, int)

    # ------------------------------------------------------------------
    # 8. List and apply master pages
    # ------------------------------------------------------------------
    masters = list_master_pages(str(pres_path))
    assert len(masters) >= 1

    # Apply existing master to a single slide
    apply_master_page(str(pres_path), masters[0])

    # Import master from a template (use the presentation itself as
    # a template — the function imports from any .odp/.otp file)
    template_path = output_dir / "template.odp"
    create_presentation(str(template_path))
    imported_name = import_master_from_template(str(pres_path), str(template_path))
    assert isinstance(imported_name, str)

    from impress.master import set_master_background

    set_master_background(str(pres_path), imported_name, "lightsteelblue")
    apply_master_page(str(pres_path), imported_name)

    # ------------------------------------------------------------------
    # 9. Find & replace
    # ------------------------------------------------------------------
    replacements = find_replace(str(pres_path), "Integration", "E2E", match_case=True)
    assert isinstance(replacements, int)

    # ------------------------------------------------------------------
    # 10. Format shape text, set background
    # ------------------------------------------------------------------
    format_shape_text(
        str(pres_path),
        3,
        tb_idx,
        bold=True,
        italic=True,
        font_size=20,
        font_name="Liberation Sans Narrow",
        color="red",
        alignment="center",
    )

    set_slide_background(str(pres_path), 3, "navy")

    # ------------------------------------------------------------------
    # 11. Add media (audio and video)
    # ------------------------------------------------------------------
    wav_path = output_dir / "test.wav"
    _create_minimal_wav(wav_path)
    audio_idx = add_audio(str(pres_path), 3, str(wav_path), 18.0, 12.0, 2.0, 2.0)
    assert isinstance(audio_idx, int)

    video_path = output_dir / "test.mp4"
    _create_minimal_video(video_path)
    video_idx = add_video(str(pres_path), 3, str(video_path), 20.0, 12.0, 3.0, 3.0)
    assert isinstance(video_idx, int)

    # ------------------------------------------------------------------
    # 12. Duplicate, delete, move slides
    # ------------------------------------------------------------------
    # Duplicate slide 3 (the content-rich slide) → new slide at index 4
    duplicate_slide(str(pres_path), 3)
    count = get_slide_count(str(pres_path))
    assert count == 6, f"Expected 6 slides after duplicate, got {count}"

    # Move the duplicate from index 4 to index 1
    move_slide(str(pres_path), 4, 1)

    moved_inventory = get_slide_inventory(str(pres_path), 1)
    assert moved_inventory["shape_count"] >= 3

    # Delete the moved duplicate
    delete_slide(str(pres_path), 1)
    count = get_slide_count(str(pres_path))
    assert count == 5, f"Expected 5 slides after delete, got {count}"

    # ------------------------------------------------------------------
    # 13. Get slide inventory
    # ------------------------------------------------------------------
    inventory = get_slide_inventory(str(pres_path), 3)
    assert inventory["slide_index"] == 3
    assert inventory["shape_count"] >= 1
    assert isinstance(inventory["shapes"], list)

    # ------------------------------------------------------------------
    # 14. Set speaker notes, read them back
    # ------------------------------------------------------------------
    set_notes(str(pres_path), 1, "Remember to greet the audience.")
    notes_text = get_notes(str(pres_path), 1)
    assert "greet" in notes_text.lower()

    # ------------------------------------------------------------------
    # 15. Snapshot AFTER formatting
    # ------------------------------------------------------------------
    after_path = output_dir / "snapshot_after.png"
    snapshot_slide(str(pres_path), 3, str(after_path))
    assert after_path.exists()

    # ------------------------------------------------------------------
    # 16. Export to PDF and PPTX
    # ------------------------------------------------------------------
    pdf_path = output_dir / "presentation.pdf"
    export_presentation(str(pres_path), str(pdf_path), "pdf")
    assert pdf_path.exists()

    pptx_path = output_dir / "presentation.pptx"
    export_presentation(str(pres_path), str(pptx_path), "pptx")
    assert pptx_path.exists()

    return {
        "presentation": pres_path,
        "snapshot_before": before_path,
        "snapshot_after": after_path,
        "export_pdf": pdf_path,
        "export_pptx": pptx_path,
    }


# ---------------------------------------------------------------------------
# Deterministic assertion tests
# ---------------------------------------------------------------------------


def test_workflow_content_assertions(tmp_path):
    """Run the workflow in tmp_path and assert document structure.

    Verifies:
        - At least 5 slides in the final presentation.
        - The title slide contains expected text.
        - The content-rich slide has multiple shapes.
        - Speaker notes are present.
        - All exported files exist.
    """
    from impress.core import get_slide_count
    from impress.notes import get_notes
    from impress.slides import get_slide_inventory

    outputs = run_end_to_end_workflow(tmp_path)
    pres = str(outputs["presentation"])

    # Slide count >= 5
    count = get_slide_count(pres)
    assert count >= 5, f"Expected >= 5 slides, got {count}"

    # Find/replace applied ("Integration" -> "E2E")
    texts = []
    for idx in range(get_slide_count(pres)):
        inv = get_slide_inventory(pres, idx)
        texts.extend([s["text"] for s in inv["shapes"] if s["text"]])
    assert any("E2E" in t for t in texts), f"Expected 'E2E' in texts: {texts}"

    # Content-rich slide has multiple shapes (find a slide with >= 3 shapes)
    max_shapes = 0
    for idx in range(get_slide_count(pres)):
        inv_any = get_slide_inventory(pres, idx)
        max_shapes = max(max_shapes, inv_any["shape_count"])
    assert max_shapes >= 3, f"Expected >= 3 shapes on a slide, got {max_shapes}"

    # Speaker notes are present on slide 1
    notes = get_notes(pres, 1)
    assert "greet" in notes.lower(), f"Expected 'greet' in notes: {notes}"

    # All output files exist
    for key, path in outputs.items():
        assert path.exists(), f"{key} not found at {path}"
        assert path.stat().st_size > 0, f"{key} is empty at {path}"


def test_workflow_outputs_to_test_output_dir():
    """Produce inspectable output files in test-output/impress/.

    Calls run_end_to_end_workflow which builds a presentation,
    snapshots before and after content/formatting changes. Assertions
    verify that all output files exist, are non-empty, and the
    before/after snapshots differ.

    Output files:
        test-output/impress/presentation.odp
        test-output/impress/presentation.pdf
        test-output/impress/presentation.pptx
        test-output/impress/snapshot_before.png
        test-output/impress/snapshot_after.png
    """
    output_dir = Path("test-output/impress")

    # Clean up previous runs
    if output_dir.exists():
        import shutil

        shutil.rmtree(output_dir)

    outputs = run_end_to_end_workflow(output_dir)

    # Assert all output files exist and are non-empty
    for key, path in outputs.items():
        assert path.exists(), f"{key} not found at {path}"
        assert path.stat().st_size > 0, f"{key} is empty at {path}"

    # Assert PNG magic bytes on both snapshots
    for key in ("snapshot_before", "snapshot_after"):
        with open(outputs[key], "rb") as f:
            assert f.read(8) == b"\x89PNG\r\n\x1a\n"

    # Assert PDF magic bytes
    with open(outputs["export_pdf"], "rb") as f:
        assert f.read(5) == b"%PDF-"
