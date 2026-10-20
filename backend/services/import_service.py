"""Coordinate one attachment import across Outlook-independent Excel helpers."""

import threading
from datetime import date
from pathlib import Path

import win32com.client

from backend.config import (
    DESTINATION_SHEETS, DOR_COMPARISON_EXCLUDING_TX_SECTION_TITLE,
    IMPORTED_DATA_HIGHLIGHT_COLOR, XL_UP,
)
from backend.exceptions import check_cancelled
from backend.logging_setup import logger
from backend.excel.connection import find_open_workbook, find_sheet
from backend.excel.destination import (
    append_dor_import_record, fill_imported_row_formulas, fill_metadata,
    get_destination_header_map, insert_destination_rows, next_empty_row, write_matrix,
)
from backend.excel.comparison import append_dor_comparison_record, get_dor_comparison_mapping
from backend.excel.highlighting import refresh_destination_workbook_data
from backend.excel.source_reader import (
    fallback_week, find_matching_header, find_source_header_row, has_value,
    normalize, read_filtered_source_rows, to_python_date,
)
from backend.excel.ranges import range_to_rows
from backend.excel.summaries import find_existing_week, get_dor_summary_values


def _open_source_workbook_headless(excel, source_path: Path):
    previous_screen_updating = excel.ScreenUpdating
    excel.ScreenUpdating = False
    try:
        source_workbook = excel.Workbooks.Open(
            str(source_path),
            UpdateLinks=0,
            ReadOnly=True,
            IgnoreReadOnlyRecommended=True,
        )
        try:
            source_workbook.Windows(1).Visible = False
        except Exception:
            source_workbook.Close(SaveChanges=False)
            raise
        return source_workbook
    finally:
        excel.ScreenUpdating = previous_screen_updating


def find_source_data_sheet(
    source_workbook,
    cancel_event: threading.Event | None = None,
):
    """
    Searches every worksheet for a likely header row.

    Returns the first worksheet containing both a Work Type header
    and a Username-like header.
    """

    checked_sheets = []

    for source_sheet in source_workbook.Worksheets:
        check_cancelled(cancel_event)
        sheet_name = str(
            source_sheet.Name
        )

        checked_sheets.append(
            sheet_name
        )

        try:
            source_rows = range_to_rows(
                source_sheet.UsedRange.Value
            )
            check_cancelled(cancel_event)

            if not source_rows:
                continue

            find_source_header_row(
                source_rows,
                maximum_rows_to_check=200,
            )

            return source_sheet

        except RuntimeError:
            continue

    raise RuntimeError(
        "No worksheet containing Work Type and Username "
        "headers was found.\n\n"
        "Worksheets checked:\n"
        + "\n".join(checked_sheets)
    )
def import_attachment(
    source_path: Path,
    selected_queue: str,
    destination_workbook_name: str,
    production_date: date,
    week_override: str,
    row_override: int | None = None,
    cancel_event: threading.Event | None = None,
    highlight_imported_data: bool = False,
) -> dict:
    """
    Opens the attachment as read-only and imports the filtered rows
    into the already-open destination workbook.
    """
    check_cancelled(cancel_event)
    excel = win32com.client.Dispatch(
        "Excel.Application"
    )

    excel.Visible = True

    destination_workbook = find_open_workbook(
        excel,
        destination_workbook_name,
    )

    destination_sheet_name = (
        DESTINATION_SHEETS[
            selected_queue
        ]
    )

    destination_sheet = find_sheet(
        destination_workbook,
        destination_sheet_name,
    )
    dor_sheet = find_sheet(
        destination_workbook,
        "DOR",
    )
    comparison_sheet = None
    if normalize(selected_queue) == "ecrash":
        comparison_sheet = find_sheet(
            destination_workbook,
            "DOR Comparison",
        )
    comparison_mapping = None
    comparison_excluding_tx_mapping = None
    if comparison_sheet is not None:
        comparison_mapping = get_dor_comparison_mapping(
            comparison_sheet,
            dor_sheet,
        )
        comparison_excluding_tx_mapping = get_dor_comparison_mapping(
            comparison_sheet,
            dor_sheet,
            DOR_COMPARISON_EXCLUDING_TX_SECTION_TITLE,
        )
    summary_sheet = find_sheet(
        destination_workbook,
        "Summary by date (All states)",
    )
    summary_sheet_wo_tx = None
    if comparison_sheet is not None:
        summary_sheet_wo_tx = find_sheet(
            destination_workbook,
            "Summary by Date (wo TX)",
        )

    source_workbook = None

    try:
        check_cancelled(cancel_event)
        source_workbook = _open_source_workbook_headless(
            excel,
            source_path,
        )

        source_sheet = find_source_data_sheet(
            source_workbook,
            cancel_event,
        )

        high_elapsed_info = {}

        def record_high_elapsed_export(export_path, row_count):
            high_elapsed_info["path"] = str(export_path)
            high_elapsed_info["count"] = row_count

        result = read_filtered_source_rows(
            source_sheet,
            selected_queue,
            cancel_event,
            on_high_elapsed_export=record_high_elapsed_export,
        )

        logger.info(
            "Read %d source rows and %d source columns",
            len(result[1]),
            len(result[0]),
        )

        source_headers, source_rows = result

        destination_headers = (
            get_destination_header_map(
                destination_sheet
            )
        )

        user_column = destination_headers.get(
            normalize("USER")
        )

        if user_column is None:
            raise RuntimeError(
                'Destination sheet does not contain a "USER" column.'
            )

        required_headers = (
            "Week",
            "Day",
            "Production Date",
        )
        missing_headers = [
            header
            for header in required_headers
            if normalize(header) not in destination_headers
        ]

        if missing_headers:
            raise RuntimeError(
                "Destination sheet is missing: "
                + ", ".join(missing_headers)
            )

        if row_override is None:
            first_destination_row = next_empty_row(
                destination_sheet,
                user_column,
            )
        else:
            first_destination_row = row_override

            maximum_start_row = (
                int(destination_sheet.Rows.Count)
                - len(source_rows)
                + 1
            )

            if first_destination_row > maximum_start_row:
                raise RuntimeError(
                    "Insert at row is too large for the available "
                    "Excel worksheet space."
                )

        existing_week = find_existing_week(
            destination_sheet,
            production_date,
            destination_headers,
        )

        if week_override.strip():
            week_value = week_override.strip()
            week_source = "manual override"

        elif existing_week:
            week_value = existing_week
            week_source = "existing destination data"

        else:
            week_value = fallback_week(
                production_date
            )

            week_source = (
                "fallback calculation"
            )

        imported_columns = [
            user_column,
            user_column + len(source_headers) - 1,
            destination_headers[normalize("Week")],
            destination_headers[normalize("Day")],
            destination_headers[normalize("Production Date")],
            15,
            16,
            18,
        ]

        rows_inserted = False

        try:
            check_cancelled(cancel_event)
            if row_override is not None:
                insert_destination_rows(
                    destination_sheet,
                    first_destination_row,
                    len(source_rows),
                )
                rows_inserted = True
                check_cancelled(cancel_event)

            # Source data starts under the destination USER column.
            write_matrix(
                destination_sheet,
                first_destination_row,
                user_column,
                source_rows,
            )

            # Fill Week, Day, and Production Date.
            fill_metadata(
                destination_sheet,
                first_destination_row,
                len(source_rows),
                destination_headers,
                week_value,
                production_date,
            )

            fill_imported_row_formulas(
                destination_sheet,
                first_destination_row,
                len(source_rows),
            )
            check_cancelled(cancel_event)

            if highlight_imported_data:
                destination_sheet.Range(
                    destination_sheet.Cells(
                        first_destination_row,
                        min(imported_columns),
                    ),
                    destination_sheet.Cells(
                        first_destination_row + len(source_rows) - 1,
                        max(imported_columns),
                    ),
                ).Interior.Color = IMPORTED_DATA_HIGHLIGHT_COLOR
        except Exception:
            if rows_inserted:
                inserted_final_row = (
                    first_destination_row
                    + len(source_rows)
                    - 1
                )
                destination_sheet.Rows(
                    f"{first_destination_row}:{inserted_final_row}"
                ).Delete(
                    Shift=XL_UP,
                )
            raise

        refresh_destination_workbook_data(
            excel,
            destination_workbook,
            cancel_event,
        )

        summary_values, pivot_source_ranges = get_dor_summary_values(
            summary_sheet,
            production_date,
            cancel_event,
        )
        summary_values_wo_tx = None
        wo_tx_pivot_source_ranges = []
        if summary_sheet_wo_tx is not None:
            (
                summary_values_wo_tx,
                wo_tx_pivot_source_ranges,
            ) = get_dor_summary_values(
                summary_sheet_wo_tx,
                production_date,
                cancel_event,
            )
        for pivot_source_range in pivot_source_ranges:
            pivot_source_range["workbook"] = str(
                destination_workbook.Name
            )
            pivot_source_range["is_highlighted"] = (
                highlight_imported_data
            )

        for pivot_source_range in wo_tx_pivot_source_ranges:
            pivot_source_range["workbook"] = str(
                destination_workbook.Name
            )
            pivot_source_range["is_highlighted"] = (
                highlight_imported_data
            )

        if highlight_imported_data:
            for pivot_source_range in pivot_source_ranges:
                summary_sheet.Cells(
                    pivot_source_range["first_row"],
                    pivot_source_range["first_column"],
                ).Interior.Color = pivot_source_range["source_color"]
            for pivot_source_range in wo_tx_pivot_source_ranges:
                summary_sheet_wo_tx.Cells(
                    pivot_source_range["first_row"],
                    pivot_source_range["first_column"],
                ).Interior.Color = pivot_source_range["source_color"]

        dor_row, dor_row_inserted = append_dor_import_record(
            dor_sheet,
            production_date,
            summary_values,
            highlight_imported_data,
        )

        comparison_ranges = []
        if comparison_sheet is not None:
            comparison_row = append_dor_comparison_record(
                comparison_sheet,
                dor_sheet,
                dor_row,
                comparison_mapping,
                highlight_imported_data,
            )
            comparison_range = {
                "workbook": str(destination_workbook.Name),
                "sheet": str(comparison_sheet.Name),
                "kind": "dor_comparison",
                **comparison_row,
            }
            comparison_ranges.append(comparison_range)

            manual_count = summary_values_wo_tx[7]
            auto_count = summary_values_wo_tx[9]
            total_count = (manual_count or 0) + (auto_count or 0)
            manual_share = (
                manual_count / total_count
                if total_count
                else 0
            )
            auto_share = (
                auto_count / total_count
                if total_count
                else 0
            )
            excluding_tx_source_values = {
                6: summary_values_wo_tx[3],
                10: manual_count,
                11: manual_share,
                12: summary_values_wo_tx[8],
                13: auto_count,
                14: auto_share,
                15: summary_values_wo_tx[10],
            }
            excluding_tx_row = append_dor_comparison_record(
                comparison_sheet,
                dor_sheet,
                dor_row,
                comparison_excluding_tx_mapping,
                highlight_imported_data,
                excluding_tx_source_values,
                comparison_row["first_row"],
            )
            comparison_ranges.append(
                {
                    "workbook": str(destination_workbook.Name),
                    "sheet": str(comparison_sheet.Name),
                    "kind": "dor_comparison",
                    **excluding_tx_row,
                }
            )

        destination_workbook.Save()

        highlight_ranges = [
            {
                "workbook": str(destination_workbook.Name),
                "sheet": str(destination_sheet.Name),
                "kind": "raw",
                "first_row": first_destination_row,
                "last_row": first_destination_row + len(source_rows) - 1,
                "first_column": min(imported_columns),
                "last_column": max(imported_columns),
            },
            {
                "workbook": str(destination_workbook.Name),
                "sheet": str(dor_sheet.Name),
                "kind": "dor",
                "first_row": dor_row,
                "last_row": dor_row,
                "first_column": 1,
                "last_column": 15,
            },
        ]

        highlight_ranges.extend(comparison_ranges)

        highlight_ranges.extend(pivot_source_ranges)
        highlight_ranges.extend(wo_tx_pivot_source_ranges)

        return {
            "rows": len(source_rows),
            "columns": len(source_headers),
            "first_row": first_destination_row,
            "last_row": (
                first_destination_row
                + len(source_rows)
                - 1
            ),
            "first_column": min(imported_columns),
            "last_column": max(imported_columns),
            "week": week_value,
            "week_source": week_source,
            "dor_row": dor_row,
            "dor_row_inserted": dor_row_inserted,
            "dor_comparison_row": (
                comparison_ranges[0]["first_row"]
                if comparison_ranges
                else None
            ),
            "dor_comparison_excluding_tx_row": (
                comparison_ranges[1]["first_row"]
                if len(comparison_ranges) > 1
                else None
            ),
            "day": production_date.strftime(
                "%A"
            ),
            "production_date": (
                production_date.strftime(
                    "%m/%d/%Y"
                )
            ),
            "sheet": str(
                destination_sheet.Name
            ),
            "workbook": str(
                destination_workbook.Name
            ),
            "dor_row_inserted": dor_row == 6,
            "raw_rows_inserted": row_override is not None,
            "highlight_imported_data": highlight_imported_data,
            "highlight_ranges": highlight_ranges,
            "high_elapsed_export": high_elapsed_info or None,
        }

    finally:
        if source_workbook is not None:
            source_workbook.Close(
                SaveChanges=False
            )
