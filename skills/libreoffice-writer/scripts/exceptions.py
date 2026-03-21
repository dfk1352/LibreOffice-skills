"""Top-level exceptions for LibreOffice skills."""


class UnoBridgeError(Exception):
    """Error related to UNO bridge operations."""


class SessionClosedError(Exception):
    """Error raised when a session is used after it has closed."""
