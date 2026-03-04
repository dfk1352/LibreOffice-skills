#!/usr/bin/env python3
"""Package a skill directory into a .skill zip archive.

The libreoffice_skills Python package is expected to be already bundled
inside each skill directory (e.g. skills/libreoffice-writer/libreoffice_skills/).
This script simply zips the skill folder contents for distribution.
"""

from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile


FRONTMATTER_RE = re.compile(r"^---\n(.*?)\n---\n", re.DOTALL)


def _read_frontmatter(text: str) -> dict[str, str]:
    match = FRONTMATTER_RE.match(text)
    if not match:
        raise ValueError("SKILL.md is missing YAML frontmatter")
    raw = match.group(1)
    data: dict[str, str] = {}
    for line in raw.splitlines():
        if not line.strip():
            continue
        if ":" not in line:
            raise ValueError(f"Invalid frontmatter line: {line}")
        key, value = line.split(":", 1)
        data[key.strip()] = value.strip()
    return data


def _validate_skill(skill_dir: Path) -> str:
    skill_file = skill_dir / "SKILL.md"
    if not skill_file.exists():
        raise FileNotFoundError(f"Missing SKILL.md in {skill_dir}")
    text = skill_file.read_text(encoding="utf-8")
    frontmatter = _read_frontmatter(text)
    if "name" not in frontmatter or "description" not in frontmatter:
        raise ValueError("Frontmatter must include name and description")
    name = frontmatter["name"]
    if not re.fullmatch(r"[A-Za-z0-9-]+", name):
        raise ValueError("Skill name must use letters, numbers, and hyphens")
    if not frontmatter["description"].startswith("Use when"):
        raise ValueError("Description must start with 'Use when'")
    bundle_dir = skill_dir / "libreoffice_skills"
    if not bundle_dir.is_dir():
        raise FileNotFoundError(
            f"Missing libreoffice_skills/ bundle in {skill_dir}. "
            "Run scripts/sync_bundles.py first."
        )
    return name


def _add_dir(zip_file: ZipFile, source: Path, archive_root: Path) -> None:
    for path in source.rglob("*"):
        if path.is_dir():
            continue
        rel = path.relative_to(source)
        zip_file.write(path, (archive_root / rel).as_posix())


def package_skill(skill_dir: Path, output_dir: Path) -> Path:
    name = _validate_skill(skill_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    output_path = output_dir / f"{name}.skill"

    with ZipFile(output_path, "w", ZIP_DEFLATED) as zip_file:
        _add_dir(zip_file, skill_dir, Path(name))

    return output_path


def package_all(output_dir: Path) -> list[Path]:
    skill_root = Path("skills")
    outputs = []
    for skill_dir in sorted(skill_root.iterdir()):
        if not skill_dir.is_dir():
            continue
        if not (skill_dir / "SKILL.md").exists():
            continue
        outputs.append(package_skill(skill_dir, output_dir))
    return outputs


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("skill_dir", type=Path, nargs="?")
    parser.add_argument("output_dir", type=Path)
    parser.add_argument(
        "--all",
        action="store_true",
        help="Package all skills under the skills/ directory",
    )
    args = parser.parse_args()

    try:
        if args.all:
            outputs = package_all(args.output_dir)
            for output in outputs:
                print(output)
            return 0
        if args.skill_dir is None:
            raise ValueError("skill_dir is required unless --all is set")
        output = package_skill(args.skill_dir, args.output_dir)
    except (ValueError, FileNotFoundError) as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1
    print(output)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
