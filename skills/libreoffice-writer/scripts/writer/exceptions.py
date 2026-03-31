"""Custom exceptions for Writer skill."""

from exceptions import (
    DocumentNotFoundError as _BaseDocumentNotFoundError,
    PatchOperationError as _BasePatchOperationError,
    PatchSyntaxError as _BasePatchSyntaxError,
    SessionClosedError,
    SkillError,
    SnapshotError as _BaseSnapshotError,
)


class WriterSkillError(SkillError):
    """Base error for Writer skill."""


class WriterSessionError(WriterSkillError, SessionClosedError):
    """Error for Writer session lifecycle misuse."""


class DocumentNotFoundError(WriterSkillError, _BaseDocumentNotFoundError):
    """Error when document file is not found."""


class InvalidMetadataError(WriterSkillError):
    """Error for invalid metadata parameters."""


class PatchSyntaxError(WriterSkillError, _BasePatchSyntaxError):
    """Error for malformed Writer patch input."""


class PatchOperationError(WriterSkillError, _BasePatchOperationError):
    """Base error for parsed Writer patch operations."""


class InvalidTargetError(PatchOperationError):
    """Error for malformed or unsupported targets."""


class TargetNoMatchError(PatchOperationError):
    """Error when a target matches no document content."""


class TargetAmbiguousError(PatchOperationError):
    """Error when a target matches more than one element."""


class InvalidFormattingError(PatchOperationError):
    """Error for invalid formatting parameters."""


class InvalidListError(PatchOperationError):
    """Error for invalid list parameters."""


class InvalidTableError(PatchOperationError):
    """Error for invalid table parameters."""


class ImageNotFoundError(PatchOperationError):
    """Error when image file is not found."""


class InvalidPayloadError(PatchOperationError):
    """Error when patch payload data is invalid for the target."""


class SnapshotError(WriterSkillError, _BaseSnapshotError):
    """Base error for snapshot operations."""


class InvalidPageError(SnapshotError):
    """Error when page number is out of bounds."""


class FilterError(SnapshotError):
    """Error when PNG export filter fails."""


__all__ = [
    "WriterSkillError",
    "WriterSessionError",
    "SessionClosedError",
    "DocumentNotFoundError",
    "InvalidMetadataError",
    "PatchSyntaxError",
    "PatchOperationError",
    "InvalidTargetError",
    "TargetNoMatchError",
    "TargetAmbiguousError",
    "InvalidFormattingError",
    "InvalidListError",
    "InvalidTableError",
    "ImageNotFoundError",
    "InvalidPayloadError",
    "SnapshotError",
    "InvalidPageError",
    "FilterError",
]
