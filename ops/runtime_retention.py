#!/usr/bin/env python3
"""Bounded cleanup for generated AIAtools runtime files.

The allow-list is intentionally fixed so a bad systemd argument cannot turn
this utility into a general-purpose recursive deletion command.
"""

from __future__ import annotations

import argparse
import os
import time
from pathlib import Path


ALLOWED_ROOTS = frozenset(
    {
        Path("/home/ubuntu/AIAtools/uploads"),
        Path("/home/ubuntu/AIAtools/outputs"),
        Path("/home/ubuntu/tax_receipt/uploads"),
        Path("/home/ubuntu/tax_receipt/outputs"),
    }
)


def cleanup_root(root: Path, cutoff: float, dry_run: bool) -> tuple[int, int]:
    resolved = root.resolve(strict=True)
    if resolved not in ALLOWED_ROOTS:
        raise ValueError(f"root is not allow-listed: {resolved}")

    removed_count = 0
    removed_bytes = 0
    for directory, _, filenames in os.walk(resolved, followlinks=False):
        directory_path = Path(directory)
        for filename in filenames:
            if filename == ".gitkeep":
                continue
            path = directory_path / filename
            try:
                stat = path.lstat()
            except FileNotFoundError:
                continue
            if path.is_symlink() or not path.is_file() or stat.st_mtime >= cutoff:
                continue
            removed_count += 1
            removed_bytes += stat.st_size
            if not dry_run:
                path.unlink()
    return removed_count, removed_bytes


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--days", type=int, default=30)
    parser.add_argument("--dry-run", action="store_true")
    parser.add_argument("roots", nargs="+")
    args = parser.parse_args()
    if args.days < 1:
        parser.error("--days must be at least 1")

    cutoff = time.time() - args.days * 24 * 60 * 60
    total_count = 0
    total_bytes = 0
    for raw_root in args.roots:
        count, size = cleanup_root(Path(raw_root), cutoff, args.dry_run)
        total_count += count
        total_bytes += size
        print(f"root={raw_root} files={count} bytes={size} dry_run={args.dry_run}")
    print(f"total_files={total_count} total_bytes={total_bytes} dry_run={args.dry_run}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
