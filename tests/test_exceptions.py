# pyright: reportMissingImports=false, reportAttributeAccessIssue=false
"""Tests for the unified exception hierarchy (#10)."""

from __future__ import annotations


def test_skill_error_is_base_of_all_app_errors():
    from exceptions import SkillError
    from writer.exceptions import WriterSkillError
    from calc.exceptions import CalcSkillError
    from impress.exceptions import ImpressSkillError

    assert issubclass(WriterSkillError, SkillError)
    assert issubclass(CalcSkillError, SkillError)
    assert issubclass(ImpressSkillError, SkillError)


def test_shared_patch_syntax_error_catches_all_app_variants():
    from exceptions import PatchSyntaxError
    from writer.exceptions import PatchSyntaxError as WriterPSE
    from calc.exceptions import PatchSyntaxError as CalcPSE
    from impress.exceptions import PatchSyntaxError as ImPSE

    assert issubclass(WriterPSE, PatchSyntaxError)
    assert issubclass(CalcPSE, PatchSyntaxError)
    assert issubclass(ImPSE, PatchSyntaxError)


def test_shared_patch_operation_error_catches_all_app_variants():
    from exceptions import PatchOperationError
    from writer.exceptions import PatchOperationError as WriterPOE
    from calc.exceptions import PatchOperationError as CalcPOE
    from impress.exceptions import PatchOperationError as ImPOE

    assert issubclass(WriterPOE, PatchOperationError)
    assert issubclass(CalcPOE, PatchOperationError)
    assert issubclass(ImPOE, PatchOperationError)


def test_shared_document_not_found_error_catches_all_app_variants():
    from exceptions import DocumentNotFoundError
    from writer.exceptions import DocumentNotFoundError as WriterDNF
    from calc.exceptions import DocumentNotFoundError as CalcDNF
    from impress.exceptions import DocumentNotFoundError as ImDNF

    assert issubclass(WriterDNF, DocumentNotFoundError)
    assert issubclass(CalcDNF, DocumentNotFoundError)
    assert issubclass(ImDNF, DocumentNotFoundError)


def test_shared_snapshot_error_catches_all_app_variants():
    from exceptions import SnapshotError
    from writer.exceptions import SnapshotError as WriterSE
    from calc.exceptions import SnapshotError as CalcSE
    from impress.exceptions import SnapshotError as ImSE

    assert issubclass(WriterSE, SnapshotError)
    assert issubclass(CalcSE, SnapshotError)
    assert issubclass(ImSE, SnapshotError)


def test_app_errors_still_catchable_by_app_root():
    """Existing except WriterSkillError still catches all Writer errors."""
    from writer.exceptions import (
        WriterSkillError,
        PatchSyntaxError,
        PatchOperationError,
        DocumentNotFoundError,
        SnapshotError,
    )

    assert issubclass(PatchSyntaxError, WriterSkillError)
    assert issubclass(PatchOperationError, WriterSkillError)
    assert issubclass(DocumentNotFoundError, WriterSkillError)
    assert issubclass(SnapshotError, WriterSkillError)


def test_session_errors_share_common_base():
    """App session errors inherit from a shared SessionClosedError."""
    from exceptions import SessionClosedError
    from writer.exceptions import WriterSessionError
    from calc.exceptions import CalcSessionError
    from impress.exceptions import ImpressSessionError

    assert issubclass(WriterSessionError, SessionClosedError)
    assert issubclass(CalcSessionError, SessionClosedError)
    assert issubclass(ImpressSessionError, SessionClosedError)


def test_catch_skill_error_catches_any_app_error():
    """Raising a concrete app error is caught by except SkillError."""
    from exceptions import SkillError
    from writer.exceptions import PatchSyntaxError

    with __import__("pytest").raises(SkillError):
        raise PatchSyntaxError("test message")


def test_uno_bridge_error_is_skill_error():
    from exceptions import SkillError, UnoBridgeError

    assert issubclass(UnoBridgeError, SkillError)
