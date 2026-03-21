import sys

sys.path.extend(
    [
        "/usr/lib/python3/dist-packages",
        "/usr/local/lib/python3.12/dist-packages",
    ]
)


def _patch_saferepr_for_uno() -> None:
    """Prevent UNO ``DisposedException`` from crashing the pytest process.

    When a test fails and local variables include UNO proxy objects whose
    backing ``soffice`` process has already exited, pytest's ``SafeRepr``
    calls ``repr()`` on the proxy.  The proxy's C++ side calls ``abort()``
    when the UNO bridge is disposed -- this is uncatchable from Python.

    We patch both ``SafeRepr.repr`` and ``SafeRepr.repr_instance`` to
    detect UNO proxy types and return a safe placeholder *before* ever
    calling ``repr()`` on them.
    """
    try:
        from _pytest._io.saferepr import SafeRepr
    except ImportError:
        return

    _original_repr = SafeRepr.repr
    _original_repr_instance = SafeRepr.repr_instance

    # UNO proxy module names that can trigger C++ abort on repr.
    _UNO_PROXY_MODULES = frozenset({"pyuno", "uno"})

    def _is_uno_proxy(obj: object) -> bool:
        module = getattr(type(obj), "__module__", "") or ""
        if module in _UNO_PROXY_MODULES:
            return True
        type_name = type(obj).__name__
        if type_name in ("pyuno", "ByteSequence", "Enum", "Type"):
            return True
        return False

    _UNO_PLACEHOLDER = "<UNO proxy -- repr skipped>"

    def _safe_repr(self, x: object) -> str:
        if _is_uno_proxy(x):
            return _UNO_PLACEHOLDER
        return _original_repr(self, x)

    def _safe_repr_instance(self, x: object, level: int) -> str:
        if _is_uno_proxy(x):
            return _UNO_PLACEHOLDER
        return _original_repr_instance(self, x, level)

    SafeRepr.repr = _safe_repr  # type: ignore[assignment]
    SafeRepr.repr_instance = _safe_repr_instance  # type: ignore[assignment]


_patch_saferepr_for_uno()
