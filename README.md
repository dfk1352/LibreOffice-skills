# LibreOffice Agent Skills

[![CI](https://github.com/dfk1352/LibreOffice-skills/actions/workflows/ci.yml/badge.svg)](https://github.com/dfk1352/LibreOffice-skills/actions/workflows/ci.yml)

A Python skill suite that lets AI agents create and edit LibreOffice documents,
spreadsheets, and presentations through a clean, stable API — without wrestling
with the UNO API directly.

---

## Table of Contents

**Quick start:**
- [Who This Is For](#who-this-is-for)
- [Prerequisites](#prerequisites)
- [Installation](#installation)

**A bit more on the details:**
- [Why This Exists](#why-this-exists)
- [What's Included](#whats-included)
- [How It Works](#how-it-works)
- [Usage Examples](#usage-examples)

**Development & contributing:**
- [Project Structure](#project-structure)
- [Development](#development)
- [Contributing](#contributing)

---

## Who This Is For

This skill suite targets **AI agents** that need to produce or modify real LibreOffice documents as part of their work. Openclaw, Claude Code, Cowork, OpenCode, Codex, Amp, Cursor, Roo Code, Kilo Code, Antigravity, etc. Any harness that's compatible with skills.

In other words, this skill suite is designed for users who want a well-tested, headless Python library on top of LibreOffice's UNO API, automating LibreOffice operations without building the infrastructure themselves.

Local first, free to use.

---

## Prerequisites

| Requirement | Version | Notes |
|---|---|---|
| Python | 3.12+ | |
| LibreOffice | 7.x or 24.x | System-installed; headless mode used |
| uv | any recent | Recommended package manager for development |

### Installing LibreOffice

Visit and download the installer from the [official LibreOffice download page](https://www.libreoffice.org/download/download-libreoffice/).

The skill scans common spots to locate `soffice.exe`.
For non-standard installs, set the `LIBREOFFICE_PROGRAM_PATH` environment variable to the `soffice.exe` path.

---

## Installation

### Via `npx skills`

```bash
# All three skills at once
npx skills add dfk1352/LibreOffice-skills

# Or pick the ones you need
npx skills add dfk1352/LibreOffice-skills --skill libreoffice-writer
npx skills add dfk1352/LibreOffice-skills --skill libreoffice-calc
npx skills add dfk1352/LibreOffice-skills --skill libreoffice-impress
```

### Via `npx openskills`

```bash
npx openskills install dfk1352/LibreOffice-skills
```

### From Source

For development or when you want to pin a specific commit:

```bash
git clone https://github.com/dfk1352/LibreOffice-skills.git
cd LibreOffice-skills
uv sync

# Rebuild the skills/*/scripts/ bundles from src/ (required after any src/ change)
python scripts/sync_bundles.py
```

Then copy (or symlink) the skill folders into your agent harness's skill directory (typically `~/.agents/skills/`):

```bash
cp -r skills/libreoffice-writer  ~/.agents/skills/
cp -r skills/libreoffice-calc    ~/.agents/skills/
cp -r skills/libreoffice-impress ~/.agents/skills/
```

If the `uno` Python module is not on the default path (common on Linux), add the system UNO package alongside it:

```bash
export PYTHONPATH="$HOME/.agents/skills/libreoffice-writer/scripts:/usr/lib/python3/dist-packages"
```

---

## Why This Exists

LibreOffice's UNO API is powerful but notoriously difficult. Agents that try to drive it directly spend most of their token budget on error recovery rather than the actual task.

This skill suite aims to solve that by packaging the UNO complexity behind a small, predictable interface:

- **Session-based editing** — open a document once, make all your changes through a single live connection, close and save. No per-operation process  spawning overhead, no race conditions.
- **Patch DSL** — express a multi-step edit plan as a single structured string and get back a machine-readable result. Supports `atomic` mode (all-or-nothing) and `best_effort` mode (apply what you can, report what failed).
- **Isolated headless process** — each session launches LibreOffice with a throwaway user profile and a unique named pipe. Nothing leaks between sessions; CI servers stay clean.
- **Visual verification** — every skill exposes a snapshot function that exports a PNG of a page, spreadsheet area, or slide, so an agent can inspect the rendered output before handing the file to the user.
- **Zero install on the agent side** — the `scripts/` bundle is a self-contained Python package. Drop it on `PYTHONPATH` and `import writer` / `import calc` / `import impress` just works.

---

## What's Included

Three skills ship in this repository, all sharing the same session/patch design and the same underlying UNO bridge.

### Writer (`libreoffice-writer`) — `.odt` documents

Full document lifecycle: create, open, read, edit, save, export to PDF or DOCX.

Editing operations cover text insertion and replacement, rich character and paragraph formatting (font, size, color, bold, italic, alignment, spacing), tables (insert, update, delete), images (embed with size control), and unordered/ordered lists.

### Calc (`libreoffice-calc`) — `.ods` spreadsheets

Create and edit spreadsheets with full cell and range operations: read/write values and formulas, bulk range writes, named ranges, number formatting, cell styles (font, color, borders, alignment), sheet management (add, rename, delete), data validation rules, and chart creation (bar, line, pie, scatter).

### Impress (`libreoffice-impress`) — `.odp` presentations

Complete slide deck authoring: add, delete, move, and duplicate slides; insert text boxes, shapes, images, tables, charts, and audio/video; set speaker notes; manage master pages (list, apply, import from template, set background color); and export to PDF or PPTX.

### Shared Infrastructure

All three skills include the same set of shared modules:

- **UNO Bridge** — headless LibreOffice process management and connection.
- **Session base** — `BaseSession` with context-manager support and closed-guard semantics.
- **Color helpers** — `resolve_color()` accepts CSS color names (`"cornflowerblue"`) or `0xRRGGBB` integers interchangeably.
- **Exception hierarchy** — app-specific errors (`WriterSkillError`, `CalcSkillError`, `ImpressSkillError`) with precise subclasses for target resolution failures, ambiguous matches, and formatting errors.
- **Snapshot tool** — app-specific utility to help agents with visual modality to verify editting results.

---

## How It Works

Each editing session maps to exactly one headless LibreOffice process and one open document. The process is spawned with a unique named pipe and a temporary, isolated user profile that is discarded on exit. Nothing persists between sessions; no global state can accumulate.

Within a session, every operation goes through the live UNO connection — text insertions, cell writes, slide manipulations — without reopening the file each time. When the session closes (or its context manager exits), the document is saved and the process is terminated.

The **patch interface** is a higher-level layer on top of sessions. It accepts an INI-style string describing one or more operations, executes them in order against the open document, and returns a `PatchApplyResult` with per-operation status.

---

## Usage Examples

### Session API — building a report in Writer

```python
import writer

# Create a blank document
writer.create_document("report.odt")

with writer.open_writer_session("report.odt") as session:
    session.insert_text("Q1 Sales Report\n")
    session.format_text(
        target=writer.WriterTarget(kind="paragraph", occurrence=1),
        formatting=writer.TextFormatting(bold=True, font_size=18),
    )
    session.insert_table(
        rows=4, cols=3,
        data=[
            ["Region", "Revenue", "Growth"],
            ["North",  "142 000", "+12%"],
            ["South",  "98 000",  "+7%"],
            ["West",   "201 000", "+19%"],
        ],
    )
    session.export("report.pdf", format="pdf")

# Render page 1 as PNG for visual check
result = writer.snapshot_page("report.odt", "report_preview.png", page=1)
```

### Patch DSL — batch-editing a spreadsheet

The patch interface lets an agent express an entire edit plan as a single string. This is useful when an agent wants to compose the full set of changes before committing any of them.

```python
import calc

calc.create_spreadsheet("budget.ods")

patch_text = """
[op1]
type = write_range
target.kind = range
target.sheet = Sheet1
target.range = A1:C4
data <<JSON
[
  ["Department", "Budget",  "Spent"],
  ["Engineering", 120000,   98000],
  ["Marketing",    60000,   55000],
  ["Operations",   40000,   38500]
]
JSON

[op2]
type = format_range
target.kind = range
target.sheet = Sheet1
target.range = A1:C1
format.bold = true
format.background_color = 0x4472C4
format.font_color = 0xFFFFFF
"""

result = calc.patch("budget.ods", patch_text, mode="atomic")
print(result.overall_status)   # "ok"
```

---

## Project Structure

```
src/
  uno_bridge.py          # Headless LibreOffice connection
  session.py             # BaseSession ABC
  colors.py              # Shared color name resolution
  exceptions.py          # Base exception hierarchy
  writer/                # Writer skill modules
  calc/                  # Calc skill modules
  impress/               # Impress skill modules
tests/                   # Unit and integration tests
skills/
  libreoffice-writer/
    SKILL.md             # Skill definition with YAML frontmatter (for agents)
    references/          # Troubleshooting guides
    scripts/             # Bundled modules (writer + shared) — set as PYTHONPATH
  libreoffice-calc/
    SKILL.md
    scripts/
  libreoffice-impress/
    SKILL.md
    scripts/
scripts/
  package_skill.py       # Zip a skill directory into a .skill archive
  sync_bundles.py        # Re-sync bundled packages from src/ into skills/
```

---

## Development

```bash
git clone https://github.com/dfk1352/LibreOffice-skills.git
cd LibreOffice-skills
uv sync

# Re-sync bundled packages after changing src/
python scripts/sync_bundles.py

# Run unit tests
uv run pytest tests/ \
  --ignore=tests/writer/test_writer_workflows.py \
  --ignore=tests/calc/test_calc_workflows.py \
  --ignore=tests/impress/test_impress_workflows.py

# Run integration tests (requires a running LibreOffice instance)
uv run pytest tests/writer/test_writer_workflows.py \
              tests/calc/test_calc_workflows.py \
              tests/impress/test_impress_workflows.py
```

---

## Contributing

Contributions are welcome! Feel free to open issues or PRs to make this skill better.
