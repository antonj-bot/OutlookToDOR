"""Remove imported rows and summary entries from Excel."""

import threading

import win32com.client

from backend.config import XL_NONE, XL_UP
from backend.exceptions import check_cancelled
from backend.logging_setup import logger
from backend.excel.connection import find_open_workbook, find_sheet
from backend.excel.highlighting import (
    highlight_imported_data_ranges, resolve_imported_highlight_range,
)

def remove_imported_data(
    import_result: dict,
    raw_row_segments: list[tuple[int, int]],
    dor_row: int,
    ranges_to_remove: list[dict],
    remaining_ranges: list[dict],
    rehighlight_remaining: bool,
    cancel_event: threading.Event | None = None,
) -> None:
    """Removes data and highlights associated with a completed import."""

    check_cancelled(cancel_event)

    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = True

    workbook = find_open_workbook(
        excel,
        import_result["workbook"],
    )

    raw_sheet = find_sheet(
        workbook,
        import_result["sheet"],
    )

    dor_sheet = find_sheet(
        workbook,
        "DOR",
    )

    # Remove highlights associated with this import.
    for imported_range in ranges_to_remove:
        if (
            imported_range.get("kind") == "pivot"
            and not imported_range.get("is_highlighted", False)
        ):
            continue

        target_sheet = find_sheet(
            workbook,
            imported_range["sheet"],
        )

        (
            first_row,
            last_row,
            first_column,
            last_column,
        ) = resolve_imported_highlight_range(
            target_sheet,
            imported_range,
            cancel_event,
        )

        target_sheet.Range(
            target_sheet.Cells(
                first_row,
                first_column,
            ),
            target_sheet.Cells(
                last_row,
                last_column,
            ),
        ).Interior.Pattern = XL_NONE

    # Remove imported raw-data rows.
    # Delete from bottom to top so earlier deletions do not shift
    # the positions of rows that still need to be removed.
    for first_row, last_row in sorted(
        raw_row_segments,
        reverse=True,
    ):
        raw_sheet.Rows(
            f"{first_row}:{last_row}"
        ).Delete(
            Shift=XL_UP,
        )

    # Remove the corresponding DOR row.
    dor_sheet.Rows(
        f"{dor_row}:{dor_row}"
    ).Delete(
        Shift=XL_UP,
    )

    # Find comparison data associated with this import.
    comparison_ranges = [
        imported_range
        for imported_range in ranges_to_remove
        if imported_range.get("kind") == "dor_comparison"
    ]

    comparison_rows = {}

    for comparison_range in comparison_ranges:
        row_key = (
            comparison_range["sheet"].casefold(),
            comparison_range["first_row"],
            comparison_range["last_row"],
        )

        comparison_rows.setdefault(
            row_key,
            [],
        ).append(
            comparison_range
        )

    # Remove or clear comparison data.
    for (
        sheet_name,
        first_row,
        last_row,
    ), row_ranges in comparison_rows.items():

        comparison_sheet = find_sheet(
            workbook,
            sheet_name,
        )

        # The import created this row, so remove the whole row.
        if any(
            item.get("row_inserted")
            for item in row_ranges
        ):
            comparison_sheet.Rows(
                f"{first_row}:{last_row}"
            ).Delete(
                Shift=XL_UP,
            )

            continue

        # The row already existed before the import,
        # so only clear the cells populated by this import.
        for comparison_range in row_ranges:
            comparison_data_range = comparison_sheet.Range(
                comparison_sheet.Cells(
                    comparison_range["first_row"],
                    comparison_range["first_column"],
                ),
                comparison_sheet.Cells(
                    comparison_range["last_row"],
                    comparison_range["last_column"],
                ),
            )

            comparison_data_range.ClearContents()
            comparison_data_range.Interior.Pattern = XL_NONE

    # Refresh workbook data after removing the import.
    try:
        workbook.RefreshAll()
        excel.CalculateUntilAsyncQueriesDone()

    except Exception as error:
        logger.warning(
            "Could not refresh workbook data after removing import: %s",
            error,
        )

    # Restore highlights belonging to imports that still exist.
    if rehighlight_remaining and remaining_ranges:
        highlight_imported_data_ranges(
            remaining_ranges,
            cancel_event,
        )


