# LibreOffice Skills

Python skill suite for [LibreOffice](https://www.libreoffice.org/) Writer, Calc, and Impress via the UNO API.
Designed for AI agents to create and edit documents, spreadsheets, and presentations through deterministic, file-system-safe operations.

## Prerequisites

- **Python 3.12+**
- **LibreOffice** (system-installed, headless mode)
- **[uv](https://docs.astral.sh/uv/)** (recommended) or pip

### Installing LibreOffice

**Debian / Ubuntu:**

```bash
sudo apt-get install libreoffice-core libreoffice-writer libreoffice-calc libreoffice-impress
```

**Fedora / RHEL:**

```bash
sudo dnf install libreoffice-core libreoffice-writer libreoffice-calc libreoffice-impress
```

**macOS (Homebrew):**

```bash
brew install --cask libreoffice
```

## Installation

Install all skills at once with [`npx skills`](https://github.com/vercel-labs/skills):

```bash
npx skills add dfk1352/LibreOffice-skills
```

Install a specific skill only:

```bash
npx skills add dfk1352/LibreOffice-skills --skill libreoffice-writer
npx skills add dfk1352/LibreOffice-skills --skill libreoffice-calc
npx skills add dfk1352/LibreOffice-skills --skill libreoffice-impress
```

Or with [`openskills`](https://github.com/numman-ali/openskills):

```bash
npx openskills install dfk1352/LibreOffice-skills
```

Each skill folder bundles the LibreOffice skill modules under `scripts/`
alongside its `SKILL.md`. After installation, add the skill's `scripts/`
directory to `PYTHONPATH` so that `import writer`, `import calc`, or
`import impress` resolves — the agent's skill tool reports the base directory
automatically.

## What's Included

### Writer

Create and edit `.odt` documents: text insertion, formatting (bold, italic,
font, color, alignment), tables, images, metadata, and page snapshots.

### Calc

Create and edit `.ods` spreadsheets: cell/range operations, formulas, number
formatting, charts, named ranges, data validation, sheet management, and area
snapshots.

### Impress

Create and edit `.odp` presentations: slides (add, delete, move, duplicate),
text boxes, shapes, images, tables, charts, audio/video, speaker notes, master
pages, find & replace, and slide snapshots.

### Shared

- **UNO Bridge** -- headless LibreOffice discovery and context management.
- **Color Helpers** -- `0xRRGGBB` integers and CSS color names.
- **Snapshot** -- PNG export for visual verification.

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

# Lint and type-check
uv run ruff check
uv run ruff format --check
uv run mypy src/
```

## Project Structure

```
src/
  uno_bridge.py          # Headless LibreOffice connection
  colors.py              # Shared color name resolution
  exceptions.py          # Base exception hierarchy
  writer/                # Writer skill modules
  calc/                  # Calc skill modules
  impress/               # Impress skill modules
tests/                   # Unit and integration tests
skills/
  libreoffice-writer/
    SKILL.md             # Skill definition with YAML frontmatter
    scripts/               # Bundled modules (writer + shared)
  libreoffice-calc/
    SKILL.md
    scripts/               # Bundled modules (calc + shared)
  libreoffice-impress/
    SKILL.md
    scripts/               # Bundled modules (impress + shared)
scripts/
  package_skill.py       # Zip a skill directory into a .skill archive
  sync_bundles.py        # Re-sync bundled packages from src/ into skills/
```

## Contributing

Contributions are welcome. Please open an issue to discuss proposed changes
before submitting a pull request.
