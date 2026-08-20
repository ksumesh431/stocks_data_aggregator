"""Command-line entry point for the portfolio workbook generator."""

import argparse
import logging
from pathlib import Path

from portfolio_report.errors import WorkbookOutputError
from portfolio_report.workflow import generate_portfolio_workbook

LOGGER = logging.getLogger(__name__)


def build_argument_parser():
    """Create the small public command-line interface."""
    parser = argparse.ArgumentParser(
        description="Create current-year and all-time portfolio performance workbooks."
    )
    parser.add_argument(
        "--reports-dir",
        default="reports_directory",
        help="Directory containing the INDmoney/Alpaca report folders.",
    )
    parser.add_argument(
        "--output",
        default="output/portfolio_performance.xlsx",
        help="Destination .xlsx file.",
    )
    parser.add_argument(
        "--year",
        help="Calendar year for the current-year sheet. Defaults to the holdings report year.",
    )
    return parser


def main():
    """Parse arguments and run the complete workflow."""
    logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")
    args = build_argument_parser().parse_args()

    year = int(args.year) if args.year else None
    try:
        output_path = generate_portfolio_workbook(
            reports_directory=Path(args.reports_dir),
            output_path=Path(args.output),
            current_year=year,
        )
    except WorkbookOutputError as error:
        LOGGER.error("%s", error)
        return 1

    print(f"Workbook created: {output_path.resolve()}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
