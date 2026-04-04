"""Shared session primitives for LibreOffice skills."""

from abc import ABC, abstractmethod
from pathlib import Path
from typing import Any, Literal, Self

from exceptions import SessionClosedError


class BaseSession(ABC):
    """Minimal shared protocol for long-lived skill sessions."""

    def __init__(
        self,
        closed_error_type: type[Exception] = SessionClosedError,
    ) -> None:
        self._closed = False
        self._closed_error_type: type[Exception] = closed_error_type
        self._uno_manager: Any = None
        self._desktop: Any = None
        self._doc: Any = None
        self._path: Path | None = None

    def __enter__(self) -> Self:
        return self

    def __exit__(self, exc_type, exc_val, exc_tb) -> Literal[False]:
        self.close(save=(exc_type is None))
        return False

    def close(self, save: bool = True) -> None:
        """Persist changes if requested, then release session resources."""
        if self._closed:
            return
        try:
            if save:
                self._doc.store()
            self._doc.close(True)
        finally:
            try:
                self._uno_manager.__exit__(None, None, None)
            finally:
                self._closed = True
                self._doc = None
                self._desktop = None
                self._uno_manager = None

    def reset(self) -> None:
        """Discard in-memory changes and reopen the backing document."""
        self._require_open()
        self._doc.close(True)
        self._uno_manager.__exit__(None, None, None)
        self._doc = None
        self._desktop = None
        self._uno_manager = None
        try:
            self._open_document()
            self._closed = False
        except Exception:
            self._closed = True
            raise

    def restore_snapshot(self, snapshot: bytes) -> None:
        """Close the document, overwrite the file, and reopen."""
        self._require_open()
        self._doc.close(True)
        self._uno_manager.__exit__(None, None, None)
        self._doc = None
        self._desktop = None
        self._uno_manager = None
        assert self._path is not None
        self._path.write_bytes(snapshot)
        try:
            self._open_document()
            self._closed = False
        except Exception:
            self._closed = True
            raise

    @abstractmethod
    def _open_document(self) -> None:
        """Open the backing document and populate session state."""

    def _require_open(self) -> None:
        if self._closed:
            raise self._closed_error_type("Session is already closed")
