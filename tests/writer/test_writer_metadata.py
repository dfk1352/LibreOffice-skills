"""Test Writer metadata operations."""

import pytest


def test_set_metadata_rejects_empty_key(tmp_path):
    from writer.core import create_document
    from writer.metadata import set_metadata
    from writer.exceptions import InvalidMetadataError

    doc_path = tmp_path / "sample.odt"
    create_document(str(doc_path))

    with pytest.raises(InvalidMetadataError):
        set_metadata(str(doc_path), {"": "value"})


def test_set_and_get_metadata(tmp_path):
    from writer.core import create_document
    from writer.metadata import (
        set_metadata,
        get_metadata,
    )

    doc_path = tmp_path / "test_metadata.odt"
    create_document(str(doc_path))

    # Set metadata
    metadata = {
        "title": "Test Document",
        "author": "Test Author",
        "subject": "Metadata Subject",
        "keywords": "alpha, beta",
        "description": "A description",
    }
    set_metadata(str(doc_path), metadata)

    # Get metadata
    retrieved = get_metadata(str(doc_path))

    assert retrieved["title"] == "Test Document"
    assert retrieved["author"] == "Test Author"
    assert retrieved["subject"] == "Metadata Subject"
    assert retrieved["keywords"] == "alpha, beta"
    assert retrieved["description"] == "A description"
