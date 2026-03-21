def test_alignment_map_matches_uno_paragraph_adjust_enum() -> None:
    """Verify ALIGNMENT_MAP values match UNO ParagraphAdjust.

    LEFT=0, RIGHT=1, BLOCK(justify)=2, CENTER=3.
    """
    from constants import ALIGNMENT_MAP

    assert ALIGNMENT_MAP["left"] == 0
    assert ALIGNMENT_MAP["right"] == 1
    assert ALIGNMENT_MAP["justify"] == 2
    assert ALIGNMENT_MAP["center"] == 3
