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
- Writer imports work, but opening a document fails before LibreOffice starts

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
PYTHONPATH="skills/libreoffice-writer/scripts:/usr/lib/python3/dist-packages" python3 your_script.py
```

If your platform uses a different LibreOffice Python location, use that path
instead.

## Patch result says `document_persisted = false`

Meaning:

- `document_persisted` reports whether the patch mutations currently exist in a
  saved document state.
- Standalone `writer.patch(...)` saves the file when successful mutations should
  persist.
- `session.patch(...)` reports `true` when the session now holds successful
  in-memory mutations, even before the session is later closed and stored.

## Target matching feels brittle

Recommendations:

1. Anchor text targets with full sentences or paragraph-sized phrases, not
   single words.
2. Add `after` and `before` when the same wording may appear elsewhere.
3. Use `occurrence` only when repeated matches are intentional and stable.

## Version Compatibility

Some Writer skill features depend on LibreOffice 26.2 or newer. The core
editing API (create, session, export, patch, snapshots) works on older versions.

### Markdown import / export (26.2+)

`create_document(path, source="doc.md")` and `export_document(path, out, "md")`
rely on the `Markdown` filter, which was added in LibreOffice 26.2. On older
versions, these calls will fail with a UNO filter error because the filter name
is not registered.

Workaround: convert Markdown to ODT or HTML externally before importing, and
export to HTML instead of Markdown on older versions.

### `"start"` / `"end"` paragraph alignment (26.2+)

`TextFormatting(align="start")` and `TextFormatting(align="end")` set the UNO
`ParaAdjust` property to `ParagraphAdjust.START` (5) and
`ParagraphAdjust.END` (6), enum values added in LibreOffice 26.2. These
alignments are direction-aware: `start` resolves to left in LTR paragraphs and
right in RTL paragraphs (and vice versa for `end`). This is useful for styles
shared across languages with different writing directions.

On older versions, setting these values will either be silently ignored or raise
a UNO property error. Use `"left"` / `"right"` instead when targeting
pre-26.2 installations.
