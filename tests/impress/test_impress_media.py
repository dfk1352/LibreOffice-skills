"""Tests for Impress media operations."""

import struct
import wave

import pytest


def _create_minimal_wav(path):
    """Create a minimal valid WAV file for testing."""
    with wave.open(str(path), "w") as w:
        w.setnchannels(1)
        w.setsampwidth(2)
        w.setframerate(44100)
        w.writeframes(struct.pack("<h", 0) * 100)


def test_add_audio_returns_index(tmp_path):
    from libreoffice_skills.impress.core import create_presentation
    from libreoffice_skills.impress.media import add_audio
    from libreoffice_skills.uno_bridge import uno_context

    path = tmp_path / "audio.odp"
    create_presentation(str(path))

    wav_path = tmp_path / "test.wav"
    _create_minimal_wav(wav_path)

    result = add_audio(str(path), 0, str(wav_path), 1.0, 1.0, 3.0, 3.0)

    assert isinstance(result, int)

    # Verify audio shape exists on the slide via UNO roundtrip
    with uno_context() as desktop:
        doc = desktop.loadComponentFromURL(path.resolve().as_uri(), "_blank", 0, ())
        try:
            slide = doc.DrawPages.getByIndex(0)
            shape = slide.getByIndex(result)
            # The shape should have a MediaURL or be a plugin/media shape
            assert shape is not None
            assert slide.Count > result
        finally:
            doc.close(True)


def test_add_video_returns_index(tmp_path):
    from libreoffice_skills.impress.core import create_presentation
    from libreoffice_skills.impress.media import add_video
    from libreoffice_skills.uno_bridge import uno_context

    path = tmp_path / "video.odp"
    create_presentation(str(path))

    # Create a minimal file to act as video placeholder
    video_path = tmp_path / "test.mp4"
    video_path.write_bytes(b"\x00" * 100)

    result = add_video(str(path), 0, str(video_path), 2.0, 2.0, 10.0, 7.0)

    assert isinstance(result, int)

    # Verify video shape exists on the slide via UNO roundtrip
    with uno_context() as desktop:
        doc = desktop.loadComponentFromURL(path.resolve().as_uri(), "_blank", 0, ())
        try:
            slide = doc.DrawPages.getByIndex(0)
            shape = slide.getByIndex(result)
            assert shape is not None
            assert slide.Count > result
        finally:
            doc.close(True)


def test_add_audio_missing_file_raises(tmp_path):
    from libreoffice_skills.impress.core import create_presentation
    from libreoffice_skills.impress.exceptions import MediaNotFoundError
    from libreoffice_skills.impress.media import add_audio

    path = tmp_path / "audio_missing.odp"
    create_presentation(str(path))

    with pytest.raises(MediaNotFoundError):
        add_audio(str(path), 0, str(tmp_path / "missing.wav"), 1.0, 1.0, 3.0, 3.0)


def test_add_video_missing_file_raises(tmp_path):
    from libreoffice_skills.impress.core import create_presentation
    from libreoffice_skills.impress.exceptions import MediaNotFoundError
    from libreoffice_skills.impress.media import add_video

    path = tmp_path / "video_missing.odp"
    create_presentation(str(path))

    with pytest.raises(MediaNotFoundError):
        add_video(str(path), 0, str(tmp_path / "missing.mp4"), 1.0, 1.0, 5.0, 5.0)
