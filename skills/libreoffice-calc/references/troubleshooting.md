# Troubleshooting

## "Binary URP bridge disposed during call" with LibreOffice 26.x

Symptom examples:

- `RuntimeException: Binary URP bridge disposed during call`
- `loadComponentFromURL` succeeds but returns a broken proxy

Cause:

- LibreOffice 26.x installs to `/opt/libreoffice26.x/` with a versioned
  binary (e.g. `libreoffice26.2`) instead of the usual `soffice` in PATH.
  System Python does not inherit the `URE_BOOTSTRAP` and `UNO_PATH` environment
  variables that LO's bundled Python sets automatically. Without them,
  `uno.getComponentContext()` creates a minimal context that cannot bridge to
  a running LO instance.

What to check:

1. Ensure you are using the latest version of this skill (the UNO bridge now
   detects versioned binaries and sets the required variables automatically).
2. If using a custom `PYTHONPATH` setup, verify that the `uno_bridge.py` from
   this skill's `scripts/` directory is the one being imported.
3. As a manual workaround, set the variables before running your script:

```bash
export URE_BOOTSTRAP="vnd.sun.star.pathname:/opt/libreoffice26.2/program/fundamentalrc"
export UNO_PATH="/opt/libreoffice26.2/program"
```

## UNO import fails after setting `PYTHONPATH`

Symptom examples:

- `ModuleNotFoundError: No module named 'uno'`
- Calc imports work, but opening a spreadsheet fails before LibreOffice starts

Cause:

- `PYTHONPATH=<skill_base_dir>/scripts` exposes the bundled skill modules, but
  some environments still need the system LibreOffice Python packages on the
  Python import path.

What to check:

1. LibreOffice is installed on the machine.
2. The Python process can import `uno`.
3. If `uno` is missing, add the distro LibreOffice Python path before running
   your script. A common Linux path is `/usr/lib/python3/dist-packages`.

Example:

```bash
PYTHONPATH="skills/libreoffice-calc/scripts:/usr/lib/python3/dist-packages" python3 your_script.py
```

If your platform uses a different LibreOffice Python location, use that path
instead.

## Patch result says `document_persisted = false`

Meaning:

- `document_persisted` reports whether the patch mutations currently exist in a
  saved spreadsheet state.
- Standalone `calc.patch(...)` saves the file when successful mutations should
  persist.
- `session.patch(...)` reports `true` when the session now holds successful
  in-memory mutations, even before the session is later closed and stored.

## Chart follow-up targeting feels unclear

Recommendations:

1. When `create_chart()` sets `title`, reuse that same string as the most
   stable later chart target name.
2. Use `CalcTarget(kind="chart", sheet=..., name=...)` when the chart title or
   assigned name is known.
3. Use chart `index` only when chart order is stable and intentional.

## Version Compatibility

Some Calc skill features depend on LibreOffice 26.2 or newer. The core editing
API (create, session, export, patch, snapshots) works on older versions.

### JSON / XML import (26.2+)

`create_spreadsheet(path, source="data.json")` and
`create_spreadsheet(path, source="data.xml")` rely on the orcus-based
auto-detection added in LibreOffice 26.2. When a JSON or XML file is passed as
`source`, `loadComponentFromURL` auto-detects the format through orcus without
needing a `FilterName`. The imported data is then saved as an ODS spreadsheet.

On older versions, this auto-detection is not available and the call will fail
with a UNO import error or silently produce an empty spreadsheet.

Workaround: convert JSON / XML to CSV or ODS externally before importing on
pre-26.2 installations.
