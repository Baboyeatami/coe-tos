#!/usr/bin/env python3
"""Copy the entire portable skill into a harness discovery directory."""
import argparse
from pathlib import Path
import shutil

ROOT = Path(__file__).resolve().parents[1]


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--harness", choices=["codex", "opencode", "claude", "generic"], required=True)
    parser.add_argument("--scope", choices=["user", "project"], default="user")
    parser.add_argument("--project", type=Path, default=Path.cwd())
    parser.add_argument("--destination", type=Path, help="Exact coe-tos destination folder for generic/custom hosts")
    parser.add_argument("--force", action="store_true", help="Explicitly replace an existing installation")
    args = parser.parse_args()
    if args.destination:
        target = args.destination.expanduser().resolve()
    elif args.harness == "generic":
        parser.error("generic requires --destination /path/to/skills/coe-tos")
    else:
        folders = {"codex": ".agents/skills", "claude": ".claude/skills", "opencode": ".opencode/skills"}
        if args.scope == "project":
            base = args.project.expanduser().resolve()
            if not base.is_dir():
                parser.error("--project must be an existing project directory")
            target = base / folders[args.harness] / "coe-tos"
        elif args.harness == "opencode":
            target = Path.home() / ".config/opencode/skills/coe-tos"
        else:
            target = Path.home() / folders[args.harness] / "coe-tos"
    source = ROOT / "skills/coe-tos"
    if source == target or source in target.parents or target in source.parents:
        parser.error("Destination must be separate from the source skill")
    if target.is_symlink():
        parser.error("Destination is a symlink; remove it manually rather than overwrite it")
    if target.exists():
        if not args.force:
            parser.error(f"{target} already exists; inspect it before using --force")
        if not (target / "SKILL.md").is_file():
            parser.error("Refusing to replace a directory that is not a skill installation")
        shutil.rmtree(target)
    target.parent.mkdir(parents=True, exist_ok=True)
    shutil.copytree(source, target, ignore=shutil.ignore_patterns("__pycache__", "*.pyc", ".DS_Store"))
    # Include licensing in folder installs as well as release archives.
    shutil.copy2(ROOT / "LICENSE", target / "LICENSE")
    shutil.copy2(ROOT / "THIRD_PARTY_NOTICES.md", target / "THIRD_PARTY_NOTICES.md")
    print("Installed:", target)
    print("Dependencies:", f'python -m pip install -r "{target / "scripts/requirements.txt"}"')
    print("Restart OpenCode; reload/restart other hosts if the skill is not listed.")


if __name__ == "__main__":
    main()
