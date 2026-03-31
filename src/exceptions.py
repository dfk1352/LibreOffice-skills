"""Top-level exceptions for LibreOffice skills."""


class SkillError(Exception):
    """Base class for all skill errors across all apps."""


class UnoBridgeError(Exception):
    """Error related to UNO bridge operations."""


class SessionClosedError(SkillError):
    """Error raised when a session is used after it has closed."""


class PatchSyntaxError(SkillError):
    """Shared base for patch syntax errors across all apps."""


class PatchOperationError(SkillError):
    """Shared base for patch operation errors across all apps."""


class DocumentNotFoundError(SkillError):
    """Shared base for document-not-found errors across all apps."""


class SnapshotError(SkillError):
    """Shared base for snapshot errors across all apps."""
