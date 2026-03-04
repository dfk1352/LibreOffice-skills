#!/usr/bin/env python3
"""Sync the libreoffice_skills bundles from src/ into each skill directory.

Each skill directory under skills/ contains a bundled copy of the
libreoffice_skills Python package so that agents can use it directly
after installing the skill via `npx skills add dfk1352/LibreOffice-skills`.

Run this script whenever src/libreoffice_skills/ is updated:

    python scripts/sync_bundles.py

The bundles are app-specific — only the relevant submodule is copied:
  - skills/libreoffice-writer/  ← shared + writer/
  - skills/libreoffice-calc/    ← shared + calc/
  - skills/libreoffice-impress/ ← shared + impress/
"""

from __future__ import annotations

import shutil
import sys
from pathlib import Path


# Shared files present in every bundle (relative to src/libreoffice_skills/).
SHARED_FILES = [
    "__init__.py",
    "uno_bridge.py",
    "exceptions.py",
    "colors.py",
]

# Map: skill directory name → app subpackage name inside libreoffice_skills/.
SKILL_SUBPACKAGE: dict[str, str] = {
    "libreoffice-writer": "writer",
    "libreoffice-calc": "calc",
    "libreoffice-impress": "impress",
}


def _copy_file(src: Path, dst: Path) -> None:
    dst.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(src, dst)


def sync_bundle(src_root: Path, skill_dir: Path, subpackage: str) -> None:
    dest_root = skill_dir / "libreoffice_skills"

    # Remove stale bundle to avoid accumulating deleted files.
    if dest_root.exists():
        shutil.rmtree(dest_root)

    # Copy shared top-level files.
    for filename in SHARED_FILES:
        src_file = src_root / filename
        if not src_file.exists():
            print(f"  WARNING: shared file not found: {src_file}", file=sys.stderr)
            continue
        _copy_file(src_file, dest_root / filename)

    # Copy the app subpackage.
    src_sub = src_root / subpackage
    if not src_sub.is_dir():
        print(f"  ERROR: subpackage directory not found: {src_sub}", file=sys.stderr)
        sys.exit(1)
    for src_file in src_sub.rglob("*"):
        if src_file.is_dir():
            continue
        rel = src_file.relative_to(src_root)
        _copy_file(src_file, dest_root / rel)

    print(f"  synced {subpackage!r} bundle → {dest_root.relative_to(Path.cwd())}")


def main() -> int:
    repo_root = Path(__file__).parent.parent
    src_root = repo_root / "src" / "libreoffice_skills"
    skills_root = repo_root / "skills"

    if not src_root.is_dir():
        print(f"ERROR: source not found: {src_root}", file=sys.stderr)
        return 1

    print(f"Source: {src_root}")
    for skill_name, subpackage in SKILL_SUBPACKAGE.items():
        skill_dir = skills_root / skill_name
        if not skill_dir.is_dir():
            print(f"  WARNING: skill dir not found, skipping: {skill_dir}")
            continue
        sync_bundle(src_root, skill_dir, subpackage)

    print("Done.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
