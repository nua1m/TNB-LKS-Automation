"""Combine two completed LKS files, preserving edited cells and in-cell pictures."""

import argparse
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from core.services.lks_merge_service import merge_lks


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("previous", type=Path, help="Previous balance LKS .xlsm")
    parser.add_argument("new", type=Path, help="New LKS .xlsm")
    parser.add_argument("--output-dir", type=Path, default=Path.cwd() / "results" / "merged")
    parser.add_argument("--limit", type=int, default=1000)
    args = parser.parse_args()

    result = merge_lks(args.previous, args.new, args.output_dir, args.limit)
    print(f"Unique SOs: {result['unique_sos']}")
    print(f"Duplicate SOs skipped: {result['duplicates_skipped']}")
    for path in result["paths"]:
        print(f"Output: {path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
