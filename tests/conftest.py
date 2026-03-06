"""Pytest configuration for LibreOffice skills tests."""

import sys

# Add system site-packages to path for UNO access
sys.path.extend(
    [
        "/usr/lib/python3/dist-packages",
        "/usr/local/lib/python3.12/dist-packages",
    ]
)
