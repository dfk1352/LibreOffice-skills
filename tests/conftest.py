"""Pytest configuration for LibreOffice skills tests."""

import glob
import sys

# Dynamically locate system site-packages containing UNO bindings.
# LibreOffice installs its Python bindings into the system site-packages
# directory (e.g. /usr/lib/python3/dist-packages on Debian/Ubuntu).
# Since the project venv does not include system packages, we add any
# matching system path so that ``import uno`` succeeds.
_UNO_SEARCH_PATTERNS = [
    "/usr/lib/python3/dist-packages",
    "/usr/lib/python3.*/dist-packages",
    "/usr/local/lib/python3.*/dist-packages",
    "/usr/lib64/python3.*/site-packages",
]

for _pattern in _UNO_SEARCH_PATTERNS:
    for _candidate in glob.glob(_pattern):
        if _candidate not in sys.path:
            sys.path.append(_candidate)
