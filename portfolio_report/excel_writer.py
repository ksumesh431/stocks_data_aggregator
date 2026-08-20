"""Render the calculations as a styled and auditable Excel workbook."""

from datetime import datetime
from pathlib import Path

import numpy as np
import pandas as pd
import xlsxwriter

from portfolio_report.errors import WorkbookOutputError

COLORS = {
    "navy": "#17324D",
    "blue": "#2F75B5",
    "green": "#1B9E5A",
    "red": "#C73E3A",
    "yellow": "#FFD600",
    "pale_green": "#E5F5EC",
    "pale_red": "#FCE8E6",
    "pale_blue": "#EAF2F8",
    "cream": "#FBF8F0",
    "border": "#D8DEE6",
    "gray": "#6B7280",
    "white": "#FFFFFF",
}


def _output_error_message(output_path, error):
    """Build a practical message for a locked or unwritable workbook path."""
    output_path = Path(output_path).resolve()
    alternative = output_path.with_name(f"{output_path.stem}_new{output_path.suffix}")
    if output_path.exists():
        reason = (
            "The workbook is probably open in Excel or LibreOffice, or another "
            "process such as a cloud-sync client is temporarily locking it."
        )
    else:
        reason = "Windows denied write access to the destination folder or filename."
    return (
        "Cannot write the Excel workbook:\n"
        f"  {output_path}\n\n"
        f"Reason: {reason}\n\n"
        "Close the workbook and run the command again. Alternatively, choose a "
        "different filename:\n"
        f'  uv run python generate_report.py --output "{alternative}"\n\n'
        f"Technical detail: {error}"
    )


def ensure_output_available(output_path):
    """Fail early when an existing destination is locked by another process."""
    output_path = Path(output_path)
    try:
        output_path.parent.mkdir(parents=True, exist_ok=True)
        if output_path.exists():
            # Opening without writing is enough to detect Excel's Windows lock.
            with output_path.open("r+b"):
                pass
    except OSError as error:
        raise WorkbookOutputError(_output_error_message(output_path, error)) from error


def _formats(workbook):
    """Create workbook formats once and reuse them everywhere."""
    return {
        "title": workbook.add_format(
            {
                "bold": True,
                "font_size": 24,
                "font_color": COLORS["navy"],
                "align": "center",
                "valign": "vcenter",
                "bg_color": COLORS["cream"],
            }
        ),
        "subtitle": workbook.add_format(
            {
                "font_size": 10,
                "font_color": COLORS["gray"],
                "align": "center",
                "valign": "vcenter",
            }
        ),
        "section": workbook.add_format(
            {
                "bold": True,
                "font_size": 12,
                "font_color": COLORS["white"],
                "bg_color": COLORS["navy"],
                "align": "left",
                "valign": "vcenter",
            }
        ),
        "card_label": workbook.add_format(
            {
                "bold": True,
                "font_size": 10,
                "font_color": COLORS["gray"],
                "align": "center",
                "valign": "vcenter",
                "border": 1,
                "border_color": COLORS["border"],
                "top": 1,
            }
        ),
        "card_value": workbook.add_format(
            {
                "bold": True,
                "font_size": 16,
                "font_color": COLORS["navy"],
                "align": "center",
                "valign": "vcenter",
                "border": 1,
                "border_color": COLORS["border"],
                "num_format": "$#,##0.00;[Red]-$#,##0.00",
            }
        ),
        "card_value_inr": workbook.add_format(
            {
                "bold": True,
                "font_size": 16,
                "font_color": COLORS["navy"],
                "align": "center",
                "valign": "vcenter",
                "border": 1,
                "border_color": COLORS["border"],
                "num_format": "[$₹-en-IN]#,##0.00;[Red]-[$₹-en-IN]#,##0.00",
            }
        ),
        "card_pct": workbook.add_format(
            {
                "bold": True,
                "font_color": COLORS["green"],
                "bg_color": COLORS["pale_green"],
                "align": "center",
                "valign": "vcenter",
                "border": 1,
                "border_color": COLORS["border"],
                "num_format": "+0.00%;[Red]-0.00%;-",
            }
        ),
        "card_pct_blank": workbook.add_format(
            {
                "align": "center",
                "valign": "vcenter",
                "border": 1,
                "border_color": COLORS["border"],
            }
        ),
        "primary_label": workbook.add_format(
            {
                "bold": True,
                "font_size": 13,
                "font_color": COLORS["navy"],
                "bg_color": COLORS["yellow"],
                "align": "center",
                "valign": "vcenter",
                "top": 2,
                "left": 2,
                "right": 2,
                "border_color": COLORS["navy"],
            }
        ),
        "primary_note": workbook.add_format(
            {
                "bold": True,
                "font_size": 10,
                "font_color": COLORS["navy"],
                "bg_color": COLORS["yellow"],
                "align": "center",
                "valign": "vcenter",
                "bottom": 2,
                "left": 2,
                "right": 2,
                "border_color": COLORS["navy"],
            }
        ),
        "metric_label": workbook.add_format(
            {
                "bold": True,
                "font_color": COLORS["gray"],
                "bg_color": COLORS["pale_blue"],
            }
        ),
        "metric_usd": workbook.add_format(
            {"num_format": "$#,##0.00;[Red]-$#,##0.00", "bg_color": COLORS["pale_blue"]}
        ),
        "metric_inr": workbook.add_format(
            {
                "num_format": "[$₹-en-IN]#,##0.00;[Red]-[$₹-en-IN]#,##0.00",
                "bg_color": COLORS["pale_blue"],
            }
        ),
        "note": workbook.add_format(
            {"font_size": 9, "font_color": COLORS["gray"], "text_wrap": True}
        ),
        "header": workbook.add_format(
            {
                "bold": True,
                "font_color": COLORS["white"],
                "bg_color": COLORS["blue"],
                "border": 1,
                "border_color": COLORS["white"],
            }
        ),
        "datetime": workbook.add_format({"num_format": "yyyy-mm-dd hh:mm"}),
    }


def _sign_color_format(workbook, currency_format, value, highlight=False):
    """Use green for profit and red for loss while preserving currency format."""
    color = COLORS["green"] if value >= 0 else COLORS["red"]
    if highlight:
        background = COLORS["yellow"]
    else:
        background = COLORS["pale_green"] if value >= 0 else COLORS["pale_red"]
    return workbook.add_format(
        {
            "bold": True,
            "font_size": 16,
            "font_color": color,
            "bg_color": background,
            "align": "center",
            "valign": "vcenter",
            "border": 1,
            "border_color": COLORS["border"],
            "num_format": currency_format,
        }
    )


def _primary_percentage_format(workbook, value):
    """Create the strongest visual treatment for the headline XIRR."""
    color = COLORS["green"] if value >= 0 else COLORS["red"]
    return workbook.add_format(
        {
            "bold": True,
            "font_size": 28,
            "font_color": color,
            "bg_color": COLORS["yellow"],
            "align": "center",
            "valign": "vcenter",
            "left": 2,
            "right": 2,
            "border_color": COLORS["navy"],
            "num_format": "+0.00%;[Red]-0.00%;0.00%",
        }
    )


def _write_cards(worksheet, workbook, formats, metrics, start_row, currency):
    """Write the five screenshot-style summary cards."""
    cards = [
        ("INVESTED", "invested_usd", None),
        ("CURRENT", "current_usd", None),
        ("REALIZED P&L", "realized_pnl_usd", "realized_pnl_pct"),
        ("UNREALIZED P&L", "unrealized_pnl_usd", "unrealized_pnl_pct"),
        ("NET P&L", "net_pnl_usd", "net_pnl_pct"),
    ]
    signed_currency_format = (
        "+$#,##0.00;[Red]-$#,##0.00;$0.00"
        if currency == "USD"
        else "+[$₹-en-IN]#,##0.00;[Red]-[$₹-en-IN]#,##0.00;[$₹-en-IN]0.00"
    )
    default_value_format = (
        formats["card_value"] if currency == "USD" else formats["card_value_inr"]
    )

    for index, card in enumerate(cards):
        label, value_key, pct_key = card
        first_col = index * 3
        last_col = first_col + 2
        value = metrics[value_key]
        value_format = default_value_format
        if "pnl" in value_key:
            value_format = _sign_color_format(
                workbook,
                signed_currency_format,
                value,
                highlight=label == "NET P&L",
            )

        worksheet.merge_range(
            start_row, first_col, start_row, last_col, label, formats["card_label"]
        )
        worksheet.merge_range(
            start_row + 1,
            first_col,
            start_row + 2,
            last_col,
            value,
            value_format,
        )
        if pct_key:
            worksheet.merge_range(
                start_row + 3,
                first_col,
                start_row + 3,
                last_col,
                metrics[pct_key],
                formats["card_pct"],
            )
        else:
            worksheet.merge_range(
                start_row + 3,
                first_col,
                start_row + 3,
                last_col,
                "",
                formats["card_pct_blank"],
            )


def _excel_value(value):
    """Convert pandas and NumPy scalars into values accepted by XlsxWriter."""
    if pd.isna(value):
        return ""
    if isinstance(value, pd.Timestamp):
        return value.to_pydatetime().replace(tzinfo=None)
    if isinstance(value, np.generic):
        return value.item()
    return value


def _friendly_header(name):
    """Turn normalized column names into readable Excel headings."""
    replacements = {
        "usd": "USD",
        "inr": "INR",
        "pnl": "P&L",
        "pct": "%",
    }
    words = []
    for word in name.split("_"):
        words.append(replacements.get(word, word.title()))
    return " ".join(words)


def _write_dataframe(worksheet, dataframe, start_row, start_col, table_name, formats):
    """Write a DataFrame as an autofiltering Excel table."""
    dataframe = dataframe.copy()
    headers = [_friendly_header(column) for column in dataframe.columns]
    for column_index, header in enumerate(headers):
        worksheet.write(start_row, start_col + column_index, header, formats["header"])

    for row_index, row in enumerate(
        dataframe.itertuples(index=False), start=start_row + 1
    ):
        for column_index, value in enumerate(row, start=start_col):
            value = _excel_value(value)
            if isinstance(value, datetime):
                worksheet.write_datetime(
                    row_index,
                    column_index,
                    value,
                    formats["datetime"],
                )
            else:
                worksheet.write(row_index, column_index, value)

    if not dataframe.empty:
        worksheet.add_table(
            start_row,
            start_col,
            start_row + len(dataframe),
            start_col + len(headers) - 1,
            {
                "name": table_name,
                "style": "Table Style Medium 2",
                "columns": [{"header": header} for header in headers],
            },
        )

    for index, column in enumerate(dataframe.columns):
        values = dataframe[column].tolist()[:200]
        width = max([len(headers[index])] + [len(str(value)) for value in values]) + 2
        worksheet.set_column(
            start_col + index, start_col + index, min(max(width, 11), 32)
        )
    return start_row + max(len(dataframe), 1) + 2


def _inr_card_metrics(inr_metrics):
    """Adapt INR calculation names to the common card renderer."""
    return {
        "invested_usd": inr_metrics["invested_inr"],
        "current_usd": inr_metrics["current_inr"],
        "realized_pnl_usd": inr_metrics["realized_pnl_inr"],
        "unrealized_pnl_usd": inr_metrics["unrealized_pnl_inr"],
        "net_pnl_usd": inr_metrics["net_pnl_inr"],
        "realized_pnl_pct": inr_metrics["realized_pnl_pct"],
        "unrealized_pnl_pct": inr_metrics["unrealized_pnl_pct"],
        "net_pnl_pct": inr_metrics["net_pnl_pct"],
    }


def _add_inr_holding_columns(dataframe, historical_fx, current_fx):
    """Add FX-aware INR basis, value, P&L and return to each holding."""
    result = dataframe.copy()
    result["cost_basis_inr"] = result["cost_basis_usd"] * historical_fx
    result["current_value_inr"] = result["current_value_usd"] * current_fx
    result["unrealized_pnl_inr"] = (
        result["current_value_inr"] - result["cost_basis_inr"]
    )
    result["return_pct_inr"] = result["unrealized_pnl_inr"] / result["cost_basis_inr"]
    return result


def _comparison_percentage_format(workbook, value, highlight=False):
    """Create an easy-to-scan percentage format for benchmark cards."""
    positive = value is not None and value >= 0
    background = COLORS["yellow"] if highlight else (
        COLORS["pale_green"] if positive else COLORS["pale_red"]
    )
    return workbook.add_format(
        {
            "bold": True,
            "font_size": 22 if highlight else 19,
            "font_color": COLORS["green"] if positive else COLORS["red"],
            "bg_color": background,
            "align": "center",
            "valign": "vcenter",
            "border": 1,
            "border_color": COLORS["border"],
            "num_format": "+0.00%;[Red]-0.00%;0.00%",
        }
    )


def _comparison_formats(workbook):
    """Create formats used only by the human-readable benchmark sheet."""
    return {
        "period_title": workbook.add_format(
            {
                "bold": True,
                "font_size": 14,
                "font_color": COLORS["white"],
                "bg_color": COLORS["navy"],
                "align": "left",
                "valign": "vcenter",
            }
        ),
        "period_dates": workbook.add_format(
            {
                "font_size": 10,
                "font_color": COLORS["gray"],
                "align": "left",
                "valign": "vcenter",
            }
        ),
        "card_label": workbook.add_format(
            {
                "bold": True,
                "font_size": 10,
                "font_color": COLORS["gray"],
                "bg_color": COLORS["white"],
                "align": "center",
                "valign": "vcenter",
                "text_wrap": True,
                "border": 1,
                "border_color": COLORS["border"],
            }
        ),
        "subheading": workbook.add_format(
            {
                "bold": True,
                "font_size": 11,
                "font_color": COLORS["navy"],
                "bg_color": COLORS["pale_blue"],
                "align": "left",
                "valign": "vcenter",
            }
        ),
        "verdict": workbook.add_format(
            {
                "bold": True,
                "font_size": 11,
                "font_color": COLORS["navy"],
                "bg_color": COLORS["cream"],
                "align": "left",
                "valign": "vcenter",
                "text_wrap": True,
                "border": 1,
                "border_color": COLORS["border"],
            }
        ),
        "guide_label": workbook.add_format(
            {
                "bold": True,
                "font_color": COLORS["navy"],
                "bg_color": COLORS["pale_blue"],
                "valign": "top",
            }
        ),
        "guide_text": workbook.add_format(
            {
                "font_color": COLORS["gray"],
                "text_wrap": True,
                "valign": "top",
            }
        ),
    }


def _write_comparison_card(
    worksheet,
    workbook,
    formats,
    first_col,
    last_col,
    start_row,
    label,
    value,
    highlight=False,
):
    """Write one benchmark percentage as a compact two-part card."""
    worksheet.merge_range(
        start_row,
        first_col,
        start_row,
        last_col,
        label,
        formats["card_label"],
    )
    if value is None:
        worksheet.merge_range(
            start_row + 1,
            first_col,
            start_row + 3,
            last_col,
            "NOT AVAILABLE",
            formats["card_label"],
        )
        return
    worksheet.merge_range(
        start_row + 1,
        first_col,
        start_row + 3,
        last_col,
        value,
        _comparison_percentage_format(workbook, value, highlight=highlight),
    )


def _lead_description(excess, strategy=True):
    """Return a plain-language lead label and verdict for one comparison."""
    if excess >= 0:
        label = "YOU LED BY"
        if strategy:
            verdict = (
                f"Your investment strategy beat QQQ by {excess * 100:.2f} "
                "percentage points "
                "after removing deposit and withdrawal timing."
            )
        else:
            verdict = (
                "With your exact contribution timing, your money beat a "
                f"QQQ-only alternative by {excess * 100:.2f} percentage points."
            )
    else:
        label = "QQQ LED BY"
        if strategy:
            verdict = (
                f"QQQ beat your investment strategy by {abs(excess) * 100:.2f} "
                "percentage points "
                "after removing deposit and withdrawal timing."
            )
        else:
            verdict = (
                "With your exact contribution timing, a QQQ-only alternative "
                f"led by {abs(excess) * 100:.2f} percentage points."
            )
    return label, verdict


def _write_benchmark_chart(
    worksheet,
    workbook,
    period,
    start_row,
    helper_row,
):
    """Add a two-question portfolio-versus-QQQ comparison chart."""
    sheet_name = worksheet.get_name()
    worksheet.write(helper_row, 16, "Question")
    worksheet.write(helper_row, 17, "Your Portfolio")
    worksheet.write(helper_row, 18, "QQQ")
    worksheet.write(helper_row + 1, 16, "Strategy return")
    worksheet.write(helper_row + 1, 17, period["portfolio_twr"])
    worksheet.write(helper_row + 1, 18, period["benchmark_twr"])
    worksheet.write(helper_row + 2, 16, "Same cash flows")
    worksheet.write(helper_row + 2, 17, period["portfolio_mwr"])
    worksheet.write(helper_row + 2, 18, period["benchmark_mwr"])

    chart = workbook.add_chart({"type": "column"})
    chart.add_series(
        {
            "name": [sheet_name, helper_row, 17],
            "categories": [sheet_name, helper_row + 1, 16, helper_row + 2, 16],
            "values": [sheet_name, helper_row + 1, 17, helper_row + 2, 17],
            "fill": {"color": COLORS["green"]},
            "border": {"color": COLORS["green"]},
            "data_labels": {"value": True, "num_format": "0.0%"},
        }
    )
    chart.add_series(
        {
            "name": [sheet_name, helper_row, 18],
            "categories": [sheet_name, helper_row + 1, 16, helper_row + 2, 16],
            "values": [sheet_name, helper_row + 1, 18, helper_row + 2, 18],
            "fill": {"color": COLORS["blue"]},
            "border": {"color": COLORS["blue"]},
            "data_labels": {"value": True, "num_format": "0.0%"},
        }
    )
    chart.set_title({"name": "Your portfolio versus QQQ — compare each pair"})
    chart.set_y_axis(
        {
            "name": "Return",
            "num_format": "0%",
            "major_gridlines": {"visible": True, "line": {"color": "#E5E7EB"}},
        }
    )
    chart.set_x_axis({"label_position": "low"})
    chart.set_legend({"position": "bottom"})
    chart.set_chartarea({"border": {"none": True}, "fill": {"color": "#FFFFFF"}})
    chart.set_plotarea({"border": {"none": True}, "fill": {"color": "#FFFFFF"}})
    chart.set_style(10)
    chart.set_size({"width": 720, "height": 310})
    chart.show_hidden_data()
    worksheet.insert_chart(start_row, 7, chart, {"x_offset": 10, "y_offset": 4})


def _write_benchmark_period(
    worksheet,
    workbook,
    formats,
    period,
    start_row,
    helper_row,
):
    """Write one current-year or all-time comparison panel."""
    start_date = pd.Timestamp(period["start_date"]).strftime("%d %b %Y")
    end_date = pd.Timestamp(period["end_date"]).strftime("%d %b %Y")
    worksheet.merge_range(
        start_row,
        0,
        start_row,
        14,
        f"{period['label']} COMPARISON",
        formats["period_title"],
    )
    worksheet.merge_range(
        start_row + 1,
        0,
        start_row + 1,
        6,
        f"Actual measurement period: {start_date} to {end_date}",
        formats["period_dates"],
    )

    worksheet.merge_range(
        start_row + 2,
        0,
        start_row + 2,
        6,
        "DID YOUR INVESTMENT STRATEGY BEAT QQQ?",
        formats["subheading"],
    )
    twr_label, twr_verdict = _lead_description(period["twr_excess"], strategy=True)
    _write_comparison_card(
        worksheet,
        workbook,
        formats,
        0,
        1,
        start_row + 3,
        "YOUR STRATEGY RETURN",
        period["portfolio_twr"],
    )
    _write_comparison_card(
        worksheet,
        workbook,
        formats,
        2,
        3,
        start_row + 3,
        "QQQ SAME-PERIOD RETURN",
        period["benchmark_twr"],
    )
    _write_comparison_card(
        worksheet,
        workbook,
        formats,
        4,
        6,
        start_row + 3,
        twr_label,
        abs(period["twr_excess"]),
        highlight=True,
    )
    worksheet.merge_range(
        start_row + 7,
        0,
        start_row + 8,
        6,
        twr_verdict,
        formats["verdict"],
    )

    money_label = (
        "ANNUALIZED MONEY-WEIGHTED RETURN"
        if period["mwr_is_annualized"]
        else "PERIOD MONEY-WEIGHTED RETURN"
    )
    worksheet.merge_range(
        start_row + 10,
        0,
        start_row + 10,
        6,
        "HOW DID YOUR ACTUAL MONEY PERFORM WITH THE SAME CONTRIBUTION TIMING?",
        formats["subheading"],
    )
    mwr_label, mwr_verdict = _lead_description(period["mwr_excess"], strategy=False)
    _write_comparison_card(
        worksheet,
        workbook,
        formats,
        0,
        1,
        start_row + 11,
        f"YOUR {money_label}",
        period["portfolio_mwr"],
    )
    _write_comparison_card(
        worksheet,
        workbook,
        formats,
        2,
        3,
        start_row + 11,
        "SAME CASH FLOWS IN QQQ",
        period["benchmark_mwr"],
    )
    _write_comparison_card(
        worksheet,
        workbook,
        formats,
        4,
        6,
        start_row + 11,
        mwr_label,
        abs(period["mwr_excess"]),
        highlight=True,
    )
    worksheet.merge_range(
        start_row + 15,
        0,
        start_row + 16,
        6,
        mwr_verdict,
        formats["verdict"],
    )
    _write_benchmark_chart(worksheet, workbook, period, start_row + 2, helper_row)


def _write_benchmark_sheet(workbook, benchmark_comparison):
    """Create a simplified, visual portfolio-versus-QQQ workbook sheet."""
    worksheet = workbook.add_worksheet("QQQ Comparison")
    worksheet.hide_gridlines(2)
    worksheet.set_tab_color(COLORS["yellow"])
    worksheet.set_column("A:G", 16)
    worksheet.set_column("H:O", 12)
    worksheet.set_column("Q:S", 2, None, {"hidden": True})
    worksheet.set_row(0, 34)

    formats = _comparison_formats(workbook)
    title_format = workbook.add_format(
        {
            "bold": True,
            "font_size": 24,
            "font_color": COLORS["navy"],
            "align": "center",
            "valign": "vcenter",
            "bg_color": COLORS["cream"],
        }
    )
    subtitle_format = workbook.add_format(
        {
            "font_size": 10,
            "font_color": COLORS["gray"],
            "align": "center",
            "valign": "vcenter",
        }
    )
    worksheet.merge_range(
        "A1:O2",
        "HOW YOUR PORTFOLIO COMPARES WITH QQQ",
        title_format,
    )
    worksheet.merge_range(
        "A3:O3",
        "USD total-return comparison using the same dates; QQQ dividends are reinvested",
        subtitle_format,
    )

    _write_benchmark_period(
        worksheet,
        workbook,
        formats,
        benchmark_comparison["current_year"],
        3,
        1,
    )
    _write_benchmark_period(
        worksheet,
        workbook,
        formats,
        benchmark_comparison["all_time"],
        21,
        5,
    )

    guide_row = 39
    worksheet.merge_range(
        guide_row,
        0,
        guide_row,
        14,
        "HOW TO READ THIS SHEET",
        formats["period_title"],
    )
    guide = [
        (
            "Strategy return",
            "Time-weighted return removes deposits and withdrawals. This is the professional comparison for deciding whether your investment choices beat QQQ.",
        ),
        (
            "Actual money return",
            "Money-weighted return includes your contribution timing. The QQQ alternative uses the exact same cash-flow dates and amounts.",
        ),
        (
            "Net P&L",
            "Net profit remains useful in the dashboard, but it is not compared with QQQ because it is not a standardized rate of return.",
        ),
        (
            "Data source",
            "Portfolio history comes from the reports; historical portfolio and dividend-adjusted QQQ prices come from Yahoo Finance.",
        ),
    ]
    for index, item in enumerate(guide, start=guide_row + 1):
        label, description = item
        worksheet.merge_range(index, 0, index, 2, label, formats["guide_label"])
        worksheet.merge_range(index, 3, index, 14, description, formats["guide_text"])
        worksheet.set_row(index, 34)

    worksheet.freeze_panes(3, 0)
    worksheet.set_landscape()
    worksheet.fit_to_pages(1, 2)
    worksheet.set_margins(0.25, 0.25, 0.5, 0.5)


def _write_dashboard_sheet(workbook, formats, sheet_name, title, result, fx):
    """Create one current-year or all-time dashboard sheet."""
    worksheet = workbook.add_worksheet(sheet_name)
    worksheet.hide_gridlines(2)
    worksheet.set_tab_color(COLORS["blue"])
    worksheet.set_column("A:O", 12)
    worksheet.set_row(0, 34)
    worksheet.merge_range("A1:O2", title, formats["title"])
    worksheet.merge_range(
        "A3:O3",
        "Yahoo market quotes: {} | USD/INR ({}): {:.4f} at {}".format(
            result["quote_summary"], fx["ticker"], fx["rate"], fx["quote_time"]
        ),
        formats["subtitle"],
    )

    worksheet.merge_range("A4:O4", "USD PERFORMANCE", formats["section"])
    _write_cards(worksheet, workbook, formats, result["metrics"], 4, "USD")

    inr_metrics = result["inr_metrics"]
    worksheet.merge_range(
        "A10:O10",
        "INR PERFORMANCE — HISTORICAL FUNDING FX {:.4f} → CURRENT FX {:.4f}".format(
            inr_metrics["historical_funding_fx"], inr_metrics["current_fx"]
        ),
        formats["section"],
    )
    _write_cards(
        worksheet,
        workbook,
        formats,
        _inr_card_metrics(inr_metrics),
        10,
        "INR",
    )

    xirr = result["metrics"]["xirr"]
    worksheet.merge_range("A16:O16", "PRIMARY RETURN — XIRR", formats["primary_label"])
    if xirr is None:
        worksheet.merge_range("A17:O19", "XIRR NOT AVAILABLE", formats["title"])
    else:
        worksheet.merge_range(
            "A17:O19", xirr, _primary_percentage_format(workbook, xirr)
        )
    worksheet.merge_range(
        "A20:O20",
        "XIRR — annualized, money-weighted return using external cash-flow timing and total ending account value",
        formats["primary_note"],
    )

    worksheet.merge_range("A22:O22", "SUPPORTING COMPONENTS", formats["section"])
    components = [
        ("Gross dividends / interest", "gross_income_usd", "gross_income_inr"),
        ("Withholding tax", "withholding_tax_usd", "withholding_tax_inr"),
        ("Net investment income", "net_income_usd", "net_income_inr"),
        ("Brokerage", "brokerage_usd", "brokerage_inr"),
        (
            "Current cash balance (not in cards)",
            "cash_balance_usd",
            "cash_balance_inr",
        ),
        (
            "FX contribution to unrealized P&L",
            None,
            "fx_unrealized_contribution_inr",
        ),
    ]
    for index, component in enumerate(components):
        label, usd_key, inr_key = component
        row = 22 + index
        worksheet.merge_range(row, 0, row, 4, label, formats["metric_label"])
        if usd_key:
            worksheet.merge_range(
                row, 5, row, 9, result["metrics"][usd_key], formats["metric_usd"]
            )
        else:
            worksheet.merge_range(row, 5, row, 9, "Currency effect", formats["note"])
        worksheet.merge_range(
            row,
            10,
            row,
            14,
            inr_metrics[inr_key],
            formats["metric_inr"],
        )

    worksheet.merge_range("A30:O30", "OPEN HOLDINGS", formats["section"])
    holdings = _add_inr_holding_columns(
        result["holdings"],
        inr_metrics["historical_funding_fx"],
        inr_metrics["current_fx"],
    )
    _write_dataframe(worksheet, holdings, 30, 0, result["table_name"], formats)
    worksheet.freeze_panes(4, 0)
    worksheet.set_landscape()
    worksheet.fit_to_pages(1, 0)
    worksheet.set_margins(0.25, 0.25, 0.5, 0.5)


def _write_methodology(workbook, formats, inventory, as_of, current_year, fx):
    """Explain definitions and list every source file inside the workbook."""
    worksheet = workbook.add_worksheet("Sources & Method")
    worksheet.hide_gridlines(2)
    worksheet.set_column("A:A", 27)
    worksheet.set_column("B:B", 100)
    worksheet.merge_range("A1:B2", "SOURCES & CALCULATION METHOD", formats["title"])

    notes = [
        ("Report date", str(pd.Timestamp(as_of).date())),
        ("Current-year period", f"{current_year}-01-01 through the report run"),
        (
            "Primary percentage",
            "XIRR is highlighted because it is the annualized money-weighted return: it accounts for the exact timing of external deposits and withdrawals and includes securities plus cash at the end.",
        ),
        (
            "QQQ strategy comparison",
            "Portfolio time-weighted return is compared with QQQ total return over the same funded days. Time weighting removes the effect of deposit and withdrawal timing, making this the appropriate test of the investment strategy.",
        ),
        (
            "QQQ cash-flow comparison",
            "A second comparison invests the portfolio's exact external cash flows into dividend-adjusted QQQ on matching dates. It compares money-weighted returns and answers how the investor's actual contribution timing performed.",
        ),
        (
            "Invested",
            "Remaining cost basis of open positions. Current-year opening positions, if any, are marked to the last Yahoo close before January 1.",
        ),
        (
            "Current",
            "Open quantity multiplied by the latest available Yahoo Finance price.",
        ),
        (
            "Realized P&L",
            "Gross sale proceeds minus disposed cost basis, before brokerage. The broker holdings cost basis resolves specific-lot differences.",
        ),
        (
            "Unrealized P&L",
            "Current market value minus the remaining period cost basis.",
        ),
        (
            "Net P&L",
            "Realized P&L + unrealized P&L + dividends/interest - withholding tax - brokerage.",
        ),
        (
            "INR values",
            "INR cost uses the USD-weighted average of actual bank deposit exchange rates; current value uses Yahoo {} at {:.4f}. Therefore INR unrealized P&L includes both stock performance and USD/INR movement.".format(
                fx["ticker"], fx["rate"]
            ),
        ),
        (
            "INR allocation",
            "The reports do not map individual remittance dollars to individual stock lots, so historical funding FX is allocated at portfolio level. This is an investment-performance estimate, not a tax calculation.",
        ),
        (
            "Cash",
            "Shown separately and intentionally excluded from the five securities-performance cards, but included in XIRR's ending account value.",
        ),
        (
            "Funding",
            "Deposits and withdrawals come from the bank movement exports and are XIRR cash flows, not P&L. Wallet settlement entries are excluded to prevent duplicates.",
        ),
        (
            "Monthly PDFs",
            "Periodic Alpaca statements were inspected as corroborating snapshots. They are not parsed because the trade, holdings and wallet exports are newer and structured.",
        ),
        (
            "Price warning",
            "Yahoo Finance quotes can be delayed. The workbook records the quote timestamps used on this run.",
        ),
    ]
    row = 3
    for label, value in notes:
        worksheet.write(row, 0, label, formats["metric_label"])
        worksheet.write(row, 1, value, formats["note"])
        worksheet.set_row(row, 32)
        row += 1

    row += 1
    worksheet.merge_range(row, 0, row, 3, "DISCOVERED SOURCE FILES", formats["section"])
    inventory_frame = pd.DataFrame(inventory)
    _write_dataframe(worksheet, inventory_frame, row + 1, 0, "SourceInventory", formats)


def write_workbook(
    output_path,
    all_time_result,
    current_year_result,
    trades,
    wallet,
    fund_movements,
    validations,
    inventory,
    as_of,
    current_year,
    fx,
    benchmark_comparison,
):
    """Write all dashboard, audit and normalized-data sheets."""
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    workbook = xlsxwriter.Workbook(output_path)
    formats = _formats(workbook)

    _write_dashboard_sheet(
        workbook,
        formats,
        "Current Year",
        f"{current_year} CURRENT-YEAR PORTFOLIO PERFORMANCE",
        current_year_result,
        fx,
    )
    _write_dashboard_sheet(
        workbook,
        formats,
        "All Time",
        "ALL-TIME PORTFOLIO PERFORMANCE",
        all_time_result,
        fx,
    )
    _write_benchmark_sheet(workbook, benchmark_comparison)

    xirr_cash_flows = pd.concat(
        [
            current_year_result["cash_flows"].assign(period="Current Year"),
            all_time_result["cash_flows"].assign(period="All Time"),
        ],
        ignore_index=True,
    )[["period", "date", "description", "cash_flow_usd"]]

    detail_sheets = [
        ("Realized - Current", current_year_result["sales"], "CurrentYearSales"),
        ("Realized - All Time", all_time_result["sales"], "AllTimeSales"),
        ("Transactions", trades, "NormalizedTransactions"),
        ("Cash Activity", wallet, "NormalizedCashActivity"),
        ("Fund Movements", fund_movements, "NormalizedFundMovements"),
        ("XIRR Cash Flows", xirr_cash_flows, "XirrCashFlows"),
        ("Validations", validations, "ReconciliationChecks"),
    ]
    for sheet_name, dataframe, table_name in detail_sheets:
        worksheet = workbook.add_worksheet(sheet_name)
        worksheet.hide_gridlines(2)
        _write_dataframe(worksheet, dataframe, 0, 0, table_name, formats)
        worksheet.freeze_panes(1, 0)

    _write_methodology(workbook, formats, inventory, as_of, current_year, fx)
    try:
        workbook.close()
    except (xlsxwriter.exceptions.FileCreateError, PermissionError) as error:
        # The file can become locked after the early check but before this write.
        raise WorkbookOutputError(_output_error_message(output_path, error)) from error
    return output_path
