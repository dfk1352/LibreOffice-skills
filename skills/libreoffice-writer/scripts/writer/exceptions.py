"""Custom exceptions for Writer skill."""

from exceptions import SessionClosedError, UnoBridgeError


class WriterSkillError(Exception):
    """Base error for Writer skill."""


class WriterSessionError(SessionClosedError, WriterSkillError):
    """Error for Writer session lifecycle misuse."""


class DocumentNotFoundError(WriterSkillError):
    """Error when document file is not found."""


class InvalidFormattingError(WriterSkillError):
    """Error for invalid formatting parameters."""


class InvalidTableError(WriterSkillError):
    """Error for invalid table parameters."""


class ImageNotFoundError(WriterSkillError):
    """Error when image file is not found."""


class InvalidMetadataError(WriterSkillError):
    """Error for invalid metadata parameters."""


class PatchSyntaxError(WriterSkillError):
    """Error for malformed Writer patch input."""


class PatchOperationError(WriterSkillError):
    """Base error for parsed Writer patch operations."""


class InvalidSelectorError(PatchOperationError):
    """Error for malformed or unsupported selectors."""


class SelectorNoMatchError(PatchOperationError):
    """Error when a selector matches no document content."""


class SelectorAmbiguousError(PatchOperationError):
    """Error when a selector matches more than one element."""


class InvalidPayloadError(PatchOperationError):
    """Error when patch payload data is invalid for the target."""


# Re-export for compatibility
__all__ = [
    "WriterSkillError",
    "WriterSessionError",
    "UnoBridgeError",
    "SessionClosedError",
    "DocumentNotFoundError",
    "InvalidFormattingError",
    "InvalidTableError",
    "ImageNotFoundError",
    "InvalidMetadataError",
    "PatchSyntaxError",
    "PatchOperationError",
    "InvalidSelectorError",
    "SelectorNoMatchError",
    "SelectorAmbiguousError",
    "InvalidPayloadError",
]
