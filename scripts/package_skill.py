#!/usr/bin/env python3
"""Validate portable frontmatter/resources and create an installable skill ZIP."""
import argparse
from pathlib import Path
import re
from zipfile import ZipFile, ZIP_DEFLATED

ROOT = Path(__file__).resolve().parents[1]
SKILL = ROOT / "skills/coe-tos"


def validate_skill():
    text = (SKILL / "SKILL.md").read_text(encoding="utf-8")
    if not text.startswith("---\n"):
        raise ValueError("SKILL.md must start with YAML frontmatter")
    front = text.split("---", 2)[1]
    # This package deliberately uses a simple subset of YAML, requiring no YAML dependency.
    fields = dict(re.findall(r"^([a-z][a-z-]*):\s*(.+)$", front, re.M))
    allowed = {"name", "description", "license", "compatibility", "metadata", "allowed-tools"}
    if set(fields) - allowed:
        raise ValueError("Nonstandard frontmatter fields")
    name = fields.get("name", "")
    if name != SKILL.name or not re.fullmatch(r"[a-z0-9]+(?:-[a-z0-9]+)*", name) or len(name) > 64:
        raise ValueError("Skill name must match directory and be lowercase/hyphenated")
    description = fields.get("description", "")
    if not 1 <= len(description) <= 1024:
        raise ValueError("Skill description must be 1-1024 characters")
    if len(fields.get("compatibility", "")) > 500:
        raise ValueError("compatibility exceeds 500 characters")
    for filename in re.findall(r"\]\(([^)#]+)(?:#[^)]*)?\)", text):
        if not filename.startswith("http") and not (SKILL / filename).exists():
            raise ValueError(f"Missing referenced resource: {filename}")
    for required in ("assets/cjc-template.xlsx", "assets/cjc-profile.json", "scripts/coe_tos.py", "scripts/requirements.txt",
                     "scripts/office.py", "scripts/windows_office.py", "scripts/requirements-windows-office.txt"):
        if not (SKILL / required).is_file():
            raise ValueError(f"Missing resource: {required}")


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--out", type=Path, default=ROOT / "dist/coe-tos.zip")
    args = parser.parse_args()
    validate_skill()
    args.out.parent.mkdir(parents=True, exist_ok=True)
    with ZipFile(args.out, "w", ZIP_DEFLATED) as archive:
        for path in sorted(SKILL.rglob("*")):
            if not path.is_file() or "__pycache__" in path.parts or path.suffix == ".pyc" or path.name == ".DS_Store":
                continue
            archive.write(path, str(Path("coe-tos") / path.relative_to(SKILL)))
        for name in ("LICENSE", "THIRD_PARTY_NOTICES.md"):
            archive.write(ROOT / name, f"coe-tos/{name}")
    print("Skill validated and packaged:", args.out)


if __name__ == "__main__":
    main()
