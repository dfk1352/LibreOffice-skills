# LibreOffice Agent Skills

[![CI](https://github.com/dfk1352/LibreOffice-skills/actions/workflows/ci.yml/badge.svg)](https://github.com/dfk1352/LibreOffice-skills/actions/workflows/ci.yml)

A set of CLI tools designed for LLM agents to create, modify, and visually inspect LibreOffice artifacts.
Currently supports Writer, Calc, and Impress.
Bundled as [agent skills](https://agentskills.io/home), compatible with all mainstream agent harnesses.

---

## Table of Contents

**Quick start:**
- [Prerequisites](#prerequisites)
- [Installation](#installation)

**A bit more on the details:**
- [Why This Exists](#why-this-exists)
- [What's Included](#Features)
- [Usage Examples](#usage-examples)

**Development & contributing:**
- [Project Structure](#project-structure)
- [Development](#development)
- [Contributing](#contributing)

---

## Prerequisites

| Requirement | Version | Notes |
|---|---|---|
| Python | 3.12+ | |
| LibreOffice | 26.2+ | older versions work with limited functionalities (see below) |

### Installing LibreOffice

Visit and download the installer from the [official LibreOffice download page](https://www.libreoffice.org/download/download-libreoffice/).

The skill automatically scans common spots to locate the LibreOffice binary. For non-standard installs, set the `LIBREOFFICE_PROGRAM_PATH` environment variable to the LibreOffice program directory.

Some features, including Markdown import/export and Calc JSON/XML import, require LibreOffice 26.2 or newer.

---

## Installation

### Quick Installation (Recommended)

```bash
npx skills add dfk1352/LibreOffice-skills           # via `npx skills`

npx openskills install dfk1352/LibreOffice-skills   # via `npx openskills`
```

### Build From Source

```bash
git clone https://github.com/dfk1352/LibreOffice-skills.git
cd LibreOffice-skills
uv sync
```

Then copy (or symlink) the skill folders into your agent harness's skill directory (typically `~/.agents/skills/`):

```bash
cp -r skills/libreoffice-writer  ~/.agents/skills/
cp -r skills/libreoffice-calc    ~/.agents/skills/
cp -r skills/libreoffice-impress ~/.agents/skills/
```

If the `uno` Python module is not on the default path, add the system UNO package alongside it:

```bash
export PYTHONPATH="$HOME/.agents/skills/libreoffice-writer/scripts:/usr/lib/python3/dist-packages"
export PYTHONPATH="$HOME/.agents/skills/libreoffice-calc/scripts:/usr/lib/python3/dist-packages"
export PYTHONPATH="$HOME/.agents/skills/libreoffice-impress/scripts:/usr/lib/python3/dist-packages"
```

---

## Why This Exists

Agents that try to drive the LibreOffice's UNO API directly oftentimes spend most of their token budget on error recovery rather than the actual task. This skill suite aims to solve that by packaging the UNO complexity behind a small, predictable interface:

- **Session-based editing** — open a document once, make all changes through a single live connection, close and save. No per-operation process  spawning overhead, no race conditions.
- **Patch DSL** — express a multi-step edit plan as a single structured string and get back a machine-readable result. Supports `atomic` mode (all-or-nothing) and `best_effort` mode (apply what you can, report what failed).
- **Isolated headless process** — each session launches LibreOffice with a throwaway user profile and a unique named pipe. Nothing leaks between sessions; CI servers stay clean.
- **Visual verification** — every skill exposes a snapshot function that exports a PNG of a page, spreadsheet area, or slide, so an agent can inspect the rendered output before handing the file to the user.
- **Zero install on the agent side** — the `scripts/` bundle is a self-contained Python package.

---

## Features

Three skills ship in this repository, all sharing the same session/patch design and the same underlying UNO bridge.

### Writer (`libreoffice-writer`) — `.odt` documents

Full document lifecycle: create, open, read, edit, save, export to PDF, DOCX, or Markdown.

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
- **Snapshot tool** — app-specific utility to help agents with visual modality to verify editing results.

---

## Usage Examples

### Building a report in Writer

```python
import writer

writer.create_document("report.odt")

with writer.WriterSession("report.odt") as session:
    session.insert_text("Q1 Sales Report\n")
    session.format_text(
        target=writer.WriterTarget(kind="text", text="Q1 Sales Report"),
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
    session.export("report.pdf", export_format="pdf")

# Render page 1 as PNG for visual check
result = writer.snapshot_page("report.odt", "report_preview.png", page=1)
```

### Batch-editing a spreadsheet

The patch interface lets an agent express an entire edit plan as a single string. This is useful when an agent wants to compose the full set of changes before committing any of them.

```python
import calc

calc.create_spreadsheet("budget.ods")

patch_text = """
[op1]
type = write_range
target.kind = range
target.sheet = Sheet1
target.row = 0
target.col = 0
target.end_row = 3
target.end_col = 2
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
target.row = 0
target.col = 0
target.end_row = 0
target.end_col = 2
format.bold = true
format.color = 0x4472C4
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
  writer/
  calc/
  impress/
tests/
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