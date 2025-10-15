"""Command-line interface for the item bills synchroniser."""

from __future__ import annotations

import argparse
import sys

from .runner import run_item_bills


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Synchronise item bills between Excel and QuickBooks"
    )
    parser.add_argument(
        "--workbook",
        required=True,
        help="Excel workbook containing the account debit vendor worksheet",
    )
    parser.add_argument("--output", help="Optional JSON output path")

    args = parser.parse_args(argv)

    path = run_item_bills("", args.workbook, output_path=args.output)
    print(f"Report written to {path}")
    return 0


if __name__ == "__main__":  # pragma: no cover - manual invocation
    sys.exit(main())
