"""Apply, remove, and relocate tracked import highlights."""

import threading
from datetime import date

import win32com.client

from backend.config import (
    DOR_COLUMN_SOURCE_COLORS,
    IMPORTED_DATA_HIGHLIGHT_COLOR,
    XL_NONE,
)
from backend.exceptions import ProcessCancelled, check_cancelled
from backend.logging_setup import logger
from backend.excel.connection import find_open_workbook, find_sheet
from backend.excel.source_reader import workbook_stem
from backend.excel.summaries import get_dor_summary_values

def resolve_imported_highlight_range(
    worksheet,
    imported_range: dict,
    cancel_event: threading.Event | None = None,
) -> tuple[int, int, int, int]:
    pivot_date = imported_range.get("pivot_source_date")
    if not pivot_date:
        return (
            imported_range["first_row"],
            imported_range["last_row"],
            imported_range["first_column"],
            imported_range["last_column"],
        )

    _, source_ranges = get_dor_summary_values(
        worksheet,
        date.fromisoformat(pivot_date),
        cancel_event,
    )
    source_field = imported_range["source_field"]
    current_range = next(
        (
            source_range
            for source_range in source_ranges
            if source_range["source_field"] == source_field
        ),
        None,
    )
    if current_range is None:
        raise RuntimeError(
            f'Could not relocate pivot source "{source_field}" for '
            f"{pivot_date}."
        )

    return (
        current_range["first_row"],
        current_range["last_row"],
        current_range["first_column"],
        current_range["last_column"],
    )
def clear_imported_data_highlight(
    imported_ranges: list[dict],
    cancel_event: threading.Event | None = None,
) -> None:
    """Clears fill formatting from imported ranges and saves their workbooks."""
    if not imported_ranges:
        return

    check_cancelled(cancel_event)
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = True
    workbooks = {}
    changed_workbooks = {}

    failures = []
    cancelled_error = None
    for imported_range in imported_ranges:
        try:
            check_cancelled(cancel_event)
            workbook_name = imported_range["workbook"]
            workbook_key = workbook_stem(workbook_name)
            workbook = workbooks.get(workbook_key)
            if workbook is None:
                workbook = find_open_workbook(excel, workbook_name)
                workbooks[workbook_key] = workbook
            changed_workbooks[workbook_key] = workbook

            worksheet = find_sheet(
                workbook,
                imported_range["sheet"],
            )
            (
                first_row,
                last_row,
                first_column,
                last_column,
            ) = resolve_imported_highlight_range(
                worksheet,
                imported_range,
                cancel_event,
            )

            worksheet.Range(
                worksheet.Cells(
                    first_row,
                    first_column,
                ),
                worksheet.Cells(
                    last_row,
                    last_column,
                ),
            ).Interior.Pattern = XL_NONE
        except ProcessCancelled as error:
            cancelled_error = error
            break
        except Exception as error:
            failures.append(
                f'{imported_range.get("workbook", "Unknown workbook")} / '
                f'{imported_range.get("sheet", "Unknown sheet")}: {error}'
            )

    for workbook in changed_workbooks.values():
        try:
            workbook.Save()
        except Exception as error:
            failures.append(f"Could not save {workbook.Name}: {error}")

    if cancelled_error is not None:
        raise cancelled_error
    if failures:
        raise RuntimeError(
            "Some imported ranges could not be unhighlighted:\n"
            + "\n".join(failures)
        )
def highlight_imported_data_ranges(
    imported_ranges: list[dict],
    cancel_event: threading.Event | None = None,
) -> None:
    """Applies import and source-specific colors to tracked ranges."""
    if not imported_ranges:
        return

    check_cancelled(cancel_event)
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = True
    workbooks = {}
    changed_workbooks = {}

    try:
        for imported_range in imported_ranges:
            check_cancelled(cancel_event)
            workbook_name = imported_range["workbook"]
            workbook_key = workbook_stem(workbook_name)
            workbook = workbooks.get(workbook_key)
            if workbook is None:
                workbook = find_open_workbook(excel, workbook_name)
                workbooks[workbook_key] = workbook

            worksheet = find_sheet(
                workbook,
                imported_range["sheet"],
            )
            (
                first_row,
                last_row,
                first_column,
                last_column,
            ) = resolve_imported_highlight_range(
                worksheet,
                imported_range,
                cancel_event,
            )

            worksheet.Range(
                worksheet.Cells(first_row, first_column),
                worksheet.Cells(last_row, last_column),
            ).Interior.Color = IMPORTED_DATA_HIGHLIGHT_COLOR

            range_kind = imported_range.get("kind")
            if range_kind == "dor":
                for column_number, source_color in DOR_COLUMN_SOURCE_COLORS.items():
                    worksheet.Cells(
                        first_row,
                        column_number,
                    ).Interior.Color = source_color
            elif range_kind == "dor_comparison":
                for column_number, source_color in imported_range.get(
                    "source_colors",
                    {},
                ).items():
                    worksheet.Cells(
                        first_row,
                        column_number,
                    ).Interior.Color = source_color
            elif range_kind == "pivot":
                worksheet.Cells(
                    first_row,
                    first_column,
                ).Interior.Color = imported_range["source_color"]

            changed_workbooks[workbook_key] = workbook

        check_cancelled(cancel_event)
    except ProcessCancelled:
        for workbook in changed_workbooks.values():
            workbook.Save()
        raise

    for workbook in changed_workbooks.values():
        workbook.Save()
def refresh_destination_workbook_data(
    excel,
    destination_workbook,
    cancel_event: threading.Event | None = None,
) -> None:
    """Refreshes workbook connections and waits for asynchronous queries."""
    check_cancelled(cancel_event)
    destination_workbook.RefreshAll()
    excel.CalculateUntilAsyncQueriesDone()
    check_cancelled(cancel_event)
