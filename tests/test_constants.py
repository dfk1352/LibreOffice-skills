def test_alignment_map_matches_uno_paragraph_adjust_enum() -> None:
    """Verify ALIGNMENT_MAP values match UNO ParagraphAdjust.

    LEFT=0, RIGHT=1, BLOCK(justify)=2, CENTER=3.
    """
    from constants import ALIGNMENT_MAP

    assert ALIGNMENT_MAP["left"] == 0
    assert ALIGNMENT_MAP["right"] == 1
    assert ALIGNMENT_MAP["justify"] == 2
    assert ALIGNMENT_MAP["center"] == 3


def test_alignment_map_contains_start_and_end_values() -> None:
    """start/end map to ParagraphAdjust.START (5) and END (6) for bidi support."""
    from constants import ALIGNMENT_MAP

    assert ALIGNMENT_MAP["start"] == 5
    assert ALIGNMENT_MAP["end"] == 6
