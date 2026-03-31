import importlib.util
import sys


def test_find_libreoffice():
    from uno_bridge import find_libreoffice

    path = find_libreoffice()
    assert path is not None, "LibreOffice should be found on this system"
    assert "soffice" in path or "libreoffice" in path.lower()


def test_uno_context_connection(tmp_path):
    """Test that we can establish a UNO connection to LibreOffice."""
    from uno_bridge import uno_context

    with uno_context() as desktop:
        assert desktop is not None
        doc = desktop.loadComponentFromURL("private:factory/swriter", "_blank", 0, ())
        try:
            assert doc.supportsService("com.sun.star.text.TextDocument")
        finally:
            doc.close(True)


def test_find_libreoffice_uses_path_lookup_before_common_paths(monkeypatch):
    """find_libreoffice returns the first executable found on PATH."""
    import shutil

    from uno_bridge import find_libreoffice

    def fake_which(executable):
        if executable == "soffice":
            return "/custom/bin/soffice"
        return None

    monkeypatch.setattr(shutil, "which", fake_which)

    assert find_libreoffice() == "/custom/bin/soffice"


def test_resolve_uno_module_adds_env_path_and_checks_uno(monkeypatch, tmp_path):
    """_resolve_uno_module inserts env override path before retrying import."""
    from uno_bridge import _resolve_uno_module

    env_path = tmp_path / "program"
    env_path.mkdir()
    checks = []

    original_find_spec = importlib.util.find_spec
    original_sys_path = list(sys.path)

    def fake_find_spec(name):
        if name != "uno":
            return original_find_spec(name)
        checks.append(list(sys.path))
        if str(env_path) in sys.path:
            return object()
        return None

    def fake_find_libreoffice():
        return None

    monkeypatch.setattr(importlib.util, "find_spec", fake_find_spec)
    monkeypatch.setattr("uno_bridge.find_libreoffice", fake_find_libreoffice)
    monkeypatch.setattr(sys, "path", original_sys_path)
    monkeypatch.setenv("LIBREOFFICE_PROGRAM_PATH", str(env_path))

    _resolve_uno_module()

    assert sys.path[0] == str(env_path)
    assert len(checks) >= 2
    assert str(env_path) not in checks[0]
    assert str(env_path) in checks[-1]


def test_uno_context_does_not_retry_consumer_noconnectexception():
    """NoConnectException raised inside the with-block must propagate, not retry (#7).

    After the refactor, yield is outside the retry loop, so a
    NoConnectException raised by consumer code inside the with-block
    should propagate as-is rather than being caught by retry logic or
    converted to UnoBridgeError.
    """
    from uno_bridge import uno_context, UnoBridgeError

    with __import__("pytest").raises(Exception) as exc_info:
        with uno_context() as desktop:
            assert desktop is not None
            from com.sun.star.connection import NoConnectException

            raise NoConnectException("simulated consumer error", None)

    # The key assertion: the exception must NOT be an UnoBridgeError.
    # If it were, the retry loop swallowed the consumer's exception.
    assert not isinstance(exc_info.value, UnoBridgeError), (
        f"Consumer NoConnectException was converted to UnoBridgeError: {exc_info.value}"
    )
