"""Locate source reports without mixing discovery with parsing logic."""

from pathlib import Path

REQUIRED_REPORTS = {
    "trades": ("trade_reports", ("*.xls", "*.xlsx")),
    "holdings": ("holding_reports", ("*.xls", "*.xlsx")),
    "wallet": ("running_balance_reports", ("*.xls", "*.xlsx")),
}

OPTIONAL_SINGLE_REPORTS = {
    "brokerage": ("brokerage_reports", ("*.xls", "*.xlsx")),
}

OPTIONAL_MULTIPLE_REPORTS = {
    "fund_movements": ("deposit_and_withdrawl_reports", ("*.xls", "*.xlsx")),
    "monthly_statements": ("monthly_account_statements", ("*.pdf",)),
}


def _find_files(directory, patterns):
    """Return files matching all configured extensions."""
    files = []
    for pattern in patterns:
        files.extend(directory.glob(pattern))
    return sorted(set(files), key=lambda path: path.name.lower())


def _latest_file(files):
    """Choose the most recently modified file when several exports exist."""
    return max(files, key=lambda path: (path.stat().st_mtime, path.name))


def discover_reports(reports_directory):
    """Find required and optional source reports below the supplied directory."""
    reports_directory = Path(reports_directory)
    if not reports_directory.exists():
        raise FileNotFoundError(f"Reports directory not found: {reports_directory}")

    reports = {}

    for name, specification in REQUIRED_REPORTS.items():
        folder_name, patterns = specification
        files = _find_files(reports_directory / folder_name, patterns)
        if not files:
            raise FileNotFoundError(
                f"Required {name} report is missing from {reports_directory / folder_name}"
            )
        reports[name] = _latest_file(files)

    for name, specification in OPTIONAL_SINGLE_REPORTS.items():
        folder_name, patterns = specification
        files = _find_files(reports_directory / folder_name, patterns)
        reports[name] = _latest_file(files) if files else None

    for name, specification in OPTIONAL_MULTIPLE_REPORTS.items():
        folder_name, patterns = specification
        reports[name] = _find_files(reports_directory / folder_name, patterns)

    reports["root"] = reports_directory
    return reports


def build_source_inventory(reports):
    """Describe every discovered file and how the workflow uses it."""
    roles = {
        "trades": (True, "Executions, quantities, gross trade amounts and brokerage"),
        "holdings": (True, "Authoritative open quantity and remaining cost basis"),
        "wallet": (True, "Dividends, withholding tax and current cash balance"),
        "brokerage": (False, "Cross-check of brokerage shown in the trade report"),
        "fund_movements": (
            False,
            "Funding detail; shown for context, never treated as profit",
        ),
        "monthly_statements": (
            False,
            "Periodic broker snapshots retained as validation references",
        ),
    }
    rows = []
    root = reports["root"]

    for category, role in roles.items():
        paths = reports.get(category)
        if paths is None:
            continue
        if not isinstance(paths, list):
            paths = [paths]
        for path in paths:
            used, purpose = role
            rows.append(
                {
                    "category": category,
                    "file": str(path.relative_to(root)),
                    "primary_calculation_input": "Yes" if used else "No",
                    "purpose": purpose,
                }
            )
    return rows
