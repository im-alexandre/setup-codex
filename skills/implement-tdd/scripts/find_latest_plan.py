#!/usr/bin/env python3
"""Find the newest likely implementation plan for $implement-tdd."""

from __future__ import annotations

import argparse
from pathlib import Path


PATTERNS = (
    "docs/superpowers/plans/*.md",
    ".codex/plans/*.md",
    "specs/**/plan.md",
    "specs/**/tasks.md",
)


def candidates(root: Path) -> list[Path]:
    found: list[Path] = []
    for pattern in PATTERNS:
        found.extend(path for path in root.glob(pattern) if path.is_file())
    return sorted(found, key=lambda path: path.stat().st_mtime, reverse=True)


def main() -> int:
    parser = argparse.ArgumentParser(description="Find newest local implementation plan.")
    parser.add_argument("--root", default=".", help="Project root to scan.")
    args = parser.parse_args()

    root = Path(args.root).resolve()
    plans = candidates(root)
    if not plans:
        return 1

    print(str(plans[0]))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
