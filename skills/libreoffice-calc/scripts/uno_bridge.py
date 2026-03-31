"""UNO bridge for connecting to LibreOffice."""

# pyright: reportMissingImports=false

import glob
import importlib.util
import os
import re
import shutil
import subprocess
import sys
import tempfile
import time
from contextlib import contextmanager
from pathlib import Path
from typing import Any, Generator

from exceptions import UnoBridgeError


def find_libreoffice() -> str | None:
    """Auto-detect LibreOffice installation.

    Searches for the soffice/libreoffice executable using multiple
    strategies to handle both distro-packaged and versioned installs
    (e.g. ``libreoffice26.2``).

    Returns:
        Path to soffice executable, or None if not found.
    """
    # 1. Exact names via PATH lookup.
    for exe in ("soffice", "libreoffice"):
        found = shutil.which(exe)
        if found:
            return found

    # 2. Scan PATH for versioned binaries (libreoffice26.2, soffice7.6, etc.).
    lo_pattern = re.compile(r"^(soffice|libreoffice)\d")
    for directory in os.environ.get("PATH", "").split(os.pathsep):
        if not os.path.isdir(directory):
            continue
        try:
            for entry in os.listdir(directory):
                if lo_pattern.match(entry):
                    full = os.path.join(directory, entry)
                    if os.access(full, os.X_OK):
                        return full
        except PermissionError:
            continue

    # 3. Well-known fixed paths.
    common_paths = [
        "/usr/bin/soffice",
        "/usr/bin/libreoffice",
        "/usr/local/bin/soffice",
        "/opt/libreoffice/program/soffice",
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",
        "C:/Program Files/LibreOffice/program/soffice",
    ]
    for path in common_paths:
        if Path(path).exists():
            return path

    # 4. Glob for versioned /opt installs.
    for path in sorted(glob.glob("/opt/libreoffice*/program/soffice")):
        if os.access(path, os.X_OK):
            return path

    return None


def _resolve_uno_module(soffice_path: str | None = None) -> None:
    """Ensure the LibreOffice-provided ``uno`` module can be imported.

    Also configures the ``URE_BOOTSTRAP`` and ``UNO_PATH`` environment
    variables when they are unset, which is required for
    ``uno.getComponentContext()`` to bootstrap a full UNO environment
    (especially with versioned or ``/opt``-based LO installs).

    Resolution order:
    1. Already importable from the current Python environment.
    2. ``LIBREOFFICE_PROGRAM_PATH`` environment variable override.
    3. Parent directory of the detected ``soffice`` executable.

    Args:
        soffice_path: Pre-resolved path to soffice executable, to avoid
            a redundant ``find_libreoffice()`` call.

    Raises:
        UnoBridgeError: If no valid LibreOffice program directory is found.
    """
    if importlib.util.find_spec("uno") is not None:
        _ensure_uno_env()
        return

    candidates: list[Path] = []

    default_candidates = [
        Path("/usr/lib/python3/dist-packages"),
        Path("/usr/lib/libreoffice/program"),
    ]
    candidates.extend(default_candidates)

    env_path = os.environ.get("LIBREOFFICE_PROGRAM_PATH")
    if env_path:
        candidates.append(Path(env_path))

    if soffice_path is None:
        soffice_path = find_libreoffice()
    if soffice_path:
        candidates.append(Path(soffice_path).resolve().parent)

    seen: set[str] = set()
    for candidate in candidates:
        if candidate.is_dir():
            candidate_str = str(candidate)
            if candidate_str in seen:
                continue
            seen.add(candidate_str)
            if candidate_str not in sys.path:
                sys.path.insert(0, candidate_str)
            if importlib.util.find_spec("uno") is not None:
                _ensure_uno_env()
                return

    raise UnoBridgeError(
        "Unable to import the LibreOffice UNO Python module. "
        "Install LibreOffice with Python UNO support or set "
        "LIBREOFFICE_PROGRAM_PATH to the LibreOffice program directory."
    )


def _ensure_uno_env() -> None:
    """Set UNO bootstrap environment variables if not already present.

    When running under system Python (rather than LO's bundled Python
    wrapper), ``URE_BOOTSTRAP`` and ``UNO_PATH`` are typically unset.
    Without them ``uno.getComponentContext()`` creates a minimal local
    context that cannot bridge to a running LibreOffice instance.

    This function locates the LO program directory by inspecting where
    the ``uno`` module was loaded from and sets the variables to point
    at the corresponding ``fundamentalrc`` file.
    """
    if os.environ.get("URE_BOOTSTRAP") and os.environ.get("UNO_PATH"):
        return

    program_dir = _find_program_dir()
    if program_dir is None:
        return

    if not os.environ.get("UNO_PATH"):
        os.environ["UNO_PATH"] = program_dir

    if not os.environ.get("URE_BOOTSTRAP"):
        fundamentalrc = os.path.join(program_dir, "fundamentalrc")
        if os.path.isfile(fundamentalrc):
            os.environ["URE_BOOTSTRAP"] = f"vnd.sun.star.pathname:{fundamentalrc}"


def _find_program_dir() -> str | None:
    """Locate the LibreOffice program directory.

    Tries, in order:
    1. The directory containing the loaded ``uno`` module.
    2. The ``LIBREOFFICE_PROGRAM_PATH`` environment variable.
    3. The parent of the detected ``soffice`` executable.

    Returns:
        Absolute path to the program directory, or None.
    """
    # 1. Derive from the loaded uno module location.
    spec = importlib.util.find_spec("uno")
    if spec and spec.origin:
        candidate = os.path.dirname(os.path.abspath(spec.origin))
        if os.path.isfile(os.path.join(candidate, "fundamentalrc")):
            return candidate

    # 2. Explicit environment override.
    env_path = os.environ.get("LIBREOFFICE_PROGRAM_PATH")
    if env_path and os.path.isdir(env_path):
        return env_path

    # 3. From the soffice binary's resolved parent.
    soffice = find_libreoffice()
    if soffice:
        candidate = str(Path(soffice).resolve().parent)
        if os.path.isdir(candidate):
            return candidate

    return None


def _connect_with_retry(
    resolver: Any,
    connection_string: str,
    max_retries: int = 50,
    delay: float = 0.2,
) -> Any:
    """Connect to a running LibreOffice instance with retries.

    Args:
        resolver: UNO URL resolver instance.
        connection_string: Pipe or socket connection string.
        max_retries: Maximum connection attempts.
        delay: Seconds between attempts.

    Returns:
        The Desktop service proxy.

    Raises:
        UnoBridgeError: If all retries are exhausted.
    """
    from com.sun.star.connection import NoConnectException

    for attempt in range(max_retries):
        try:
            ctx = resolver.resolve(
                f"uno:{connection_string};urp;StarOffice.ComponentContext"
            )
            smgr = ctx.ServiceManager
            return smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
        except NoConnectException:
            if attempt == max_retries - 1:
                raise UnoBridgeError(
                    f"Failed to connect to LibreOffice after {max_retries} attempts"
                )
            time.sleep(delay)
    # Unreachable, but satisfies type checkers.
    raise UnoBridgeError("Failed to connect to LibreOffice")


@contextmanager
def uno_context() -> Generator[Any, None, None]:
    """Context manager for UNO connection to LibreOffice.

    Yields:
        Desktop object for creating/loading documents.

    Raises:
        UnoBridgeError: If LibreOffice cannot be found or connection fails.

    Example:
        with uno_context() as desktop:
            doc = desktop.loadComponentFromURL(...)
    """
    soffice_path = find_libreoffice()
    if not soffice_path:
        raise UnoBridgeError("LibreOffice not found. Please install LibreOffice.")

    _resolve_uno_module(soffice_path)

    import uno

    pipe_name = f"uno_pipe_{os.getpid()}_{int(time.time() * 1000)}"
    connection_string = f"pipe,name={pipe_name}"
    profile_dir = Path(tempfile.mkdtemp(prefix="libreoffice-skills-"))
    profile_url = profile_dir.resolve().as_uri()

    process = subprocess.Popen(
        [
            soffice_path,
            "--headless",
            "--invisible",
            "--nocrashreport",
            "--nodefault",
            "--nofirststartwizard",
            "--nologo",
            "--norestore",
            f"-env:UserInstallation={profile_url}",
            f"--accept=pipe,name={pipe_name};urp;",
        ],
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
    )

    try:
        local_context = uno.getComponentContext()
        resolver = local_context.ServiceManager.createInstanceWithContext(
            "com.sun.star.bridge.UnoUrlResolver", local_context
        )

        desktop = _connect_with_retry(resolver, connection_string)
        yield desktop
    finally:
        try:
            process.terminate()
            process.wait(timeout=5)
        except subprocess.TimeoutExpired:
            process.kill()
            process.wait()
        shutil.rmtree(profile_dir, ignore_errors=True)
