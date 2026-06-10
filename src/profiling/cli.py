"""Command-line interface for Phase 1 workbook profiling.

Usage::

    python -m src.profiling.cli path/to/file1.xlsx path/to/file2.csv \\
        --output output/profiling_summary.xlsx --log-level INFO
"""

from __future__ import annotations

import argparse
import logging
import sys
from pathlib import Path


def _build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="profiler",
        description="Profile one or more Excel / CSV workbooks and export a summary.",
    )
    p.add_argument("files", nargs="+", metavar="FILE", help="Excel or CSV files to profile")
    p.add_argument(
        "--output", "-o",
        default="output/profiling_summary.xlsx",
        help="Output .xlsx path (default: output/profiling_summary.xlsx)",
    )
    p.add_argument(
        "--log-level", "-l",
        default="INFO",
        choices=["DEBUG", "INFO", "WARNING", "ERROR"],
        help="Logging verbosity",
    )
    return p


def main(argv: list[str] | None = None) -> int:
    parser = _build_parser()
    args = parser.parse_args(argv)

    logging.basicConfig(
        level=getattr(logging, args.log_level),
        format="%(asctime)s | %(levelname)-8s | %(name)s | %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )

    # Deferred imports so the module stays importable without heavy deps
    from src.ingestion.loader import load_many
    from src.profiling.exporter import export_profiling_results
    from src.profiling.profiler import profile_many

    paths = [Path(f) for f in args.files]
    bundles = load_many(paths)
    if not bundles:
        logging.error("No files could be loaded. Exiting.")
        return 1

    metas = profile_many(bundles)
    if not metas:
        logging.error("Profiling produced no results. Exiting.")
        return 1

    out = export_profiling_results(metas, output_path=args.output)
    print(f"Profiling summary written to: {out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
