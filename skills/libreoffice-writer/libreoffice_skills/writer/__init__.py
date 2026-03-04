"""Writer skill package."""

from libreoffice_skills.writer.core import create_document
from libreoffice_skills.writer.snapshot import snapshot_page

__all__ = ["create_document", "snapshot_page"]
