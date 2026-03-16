#!/usr/bin/env python3
"""Sync skill bundles from src/ into skills/*/scripts/."""

from __future__ import annotations

import argparse
import shutil
import sys
from pathlib import Path


BUNDLE_DIR = "scripts"

SHARED_FILES = [
    "uno_bridge.py",
    "exceptions.py",
    "colors.py",
    "session.py",
]

SKILL_SUBPACKAGE: dict[str, str] = {
    "libreoffice-writer": "writer",
    "libreoffice-calc": "calc",
    "libreoffice-impress": "impress",
}


def _copy_file(src: Path, dst: Path) -> None:
    dst.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(src, dst)


def sync_bundle(src_root: Path, skill_dir: Path, subpackage: str) -> None:
    dest_root = skill_dir / BUNDLE_DIR
    if dest_root.exists():
        shutil.rmtree(dest_root)
    dest_root.mkdir(parents=True, exist_ok=True)

    for filename in SHARED_FILES:
        src_file = src_root / filename
        if not src_file.exists():
            raise FileNotFoundError(f"Missing shared file: {src_file}")
        _copy_file(src_file, dest_root / filename)

    src_sub = src_root / subpackage
    if not src_sub.is_dir():
        raise FileNotFoundError(f"Missing subpackage directory: {src_sub}")

    for src_file in src_sub.rglob("*"):
        if src_file.is_dir():
            continue
        rel = src_file.relative_to(src_root)
        _copy_file(src_file, dest_root / rel)

    print(f"synced {subpackage!r} bundle -> {dest_root}")


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--repo-root",
        type=Path,
        default=Path.cwd(),
        help="Repo root containing src/ and skills/",
    )
    args = parser.parse_args()

    repo_root = args.repo_root.resolve()
    src_root = repo_root / "src"
    skills_root = repo_root / "skills"

    if not src_root.is_dir():
        print(f"ERROR: source not found: {src_root}", file=sys.stderr)
        return 1
    if not skills_root.is_dir():
        print(f"ERROR: skills root not found: {skills_root}", file=sys.stderr)
        return 1

    for skill_name, subpackage in SKILL_SUBPACKAGE.items():
        skill_dir = skills_root / skill_name
        if not skill_dir.is_dir():
            print(f"WARNING: skill dir not found, skipping: {skill_dir}")
            continue
        sync_bundle(src_root, skill_dir, subpackage)

    print("Done.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
