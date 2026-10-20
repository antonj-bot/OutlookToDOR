"""Writing imported rows and DOR summary records to Excel."""

from datetime import date, datetime

from backend.config import (
    DOR_COLUMN_SOURCE_COLORS, IMPORTED_DATA_HIGHLIGHT_COLOR, XL_DOWN, XL_TO_LEFT, XL_UP,
)
from backend.excel.ranges import range_to_rows
from backend.excel.source_reader import has_value, normalize

def get_destination_header_map(
    destination_sheet,
    header_row: int = 1,
) -> dict[str, int]:
    """
    Returns:

        normalized header -> Excel column number

    Excel column numbers are one-based.
    """
    last_column = int(
        destination_sheet
        .Cells(
            header_row,
            destination_sheet.Columns.Count,
        )
        .End(XL_TO_LEFT)
        .Column
    )

    header_map: dict[str, int] = {}

    for column_number in range(
        1,
        last_column + 1,
    ):
        header_value = (
            destination_sheet
            .Cells(
                header_row,
                column_number,
            )
            .Value
        )

        normalized_header = normalize(
            header_value
        )

        if (
            normalized_header
            and normalized_header not in header_map
        ):
            header_map[
                normalized_header
            ] = column_number

    return header_map
def next_empty_row(
    destination_sheet,
    user_column: int,
    header_row: int = 1,
) -> int:
    """
    Finds the next empty row using the destination USER column.
    """
    last_used_row = int(
        destination_sheet
        .Cells(
            destination_sheet.Rows.Count,
            user_column,
        )
        .End(XL_UP)
        .Row
    )

    if last_used_row <= header_row:
        return header_row + 1

    return last_used_row + 1
def insert_destination_rows(
    destination_sheet,
    first_row: int,
    row_count: int,
) -> None:
    """
    Inserts worksheet rows so existing destination data moves down.
    """
    if row_count <= 0:
        return

    final_row = first_row + row_count - 1

    destination_sheet.Rows(
        f"{first_row}:{final_row}"
    ).Insert(
        Shift=XL_DOWN,
    )
def append_dor_import_record(
    dor_sheet,
    production_date: date,
    summary_values: tuple,
    highlight_imported_data: bool = False,
) -> tuple[int, bool]:
    """
    Writes the report's summary values to the DOR sheet in date order
    (earliest at the top). Returns (dor_row, row_inserted).
    """
    first_data_row = 5

    # Excel serial number avoids pywin32 timezone shifting.
    excel_date_serial = (production_date - date(1899, 12, 30)).days

    last_row = int(
        dor_sheet.Cells(dor_sheet.Rows.Count, 1).End(XL_UP).Row
    )

    dor_row = None

    if last_row >= first_data_row:
        column_values = range_to_rows(
            dor_sheet.Range(
                dor_sheet.Cells(first_data_row, 1),
                dor_sheet.Cells(last_row, 1),
            ).Value2
        )

        for offset, row_values in enumerate(column_values):
            value = row_values[0] if row_values else None

            if isinstance(value, (int, float)):
                # First existing date later than ours: go above it.
                if int(value) > excel_date_serial:
                    dor_row = first_data_row + offset
                    break
            elif has_value(value):
                # Non-date text (e.g. a total row): stay above it.
                dor_row = first_data_row + offset
                break

    if dor_row is None:
        dor_row = max(first_data_row, last_row + 1)
        row_inserted = False
    else:
        insert_destination_rows(dor_sheet, dor_row, 1)
        row_inserted = True

    try:
        write_matrix(
            dor_sheet,
            dor_row,
            1,
            [[
                excel_date_serial,
                None,
                *summary_values[:8],
                None,
                summary_values[8],
                summary_values[9],
                None,
                summary_values[10],
            ]],
        )

        dor_sheet.Cells(dor_row, 1).NumberFormat = "dddd, mmmm d, yyyy"
        dor_sheet.Cells(dor_row, 11).FormulaR1C1 = "=IFERROR(RC[-1]/RC[-7],0)"
        dor_sheet.Cells(dor_row, 14).FormulaR1C1 = "=IFERROR(RC[-1]/RC[-10],0)"

        if highlight_imported_data:
            dor_sheet.Range(
                dor_sheet.Cells(dor_row, 1),
                dor_sheet.Cells(dor_row, 15),
            ).Interior.Color = IMPORTED_DATA_HIGHLIGHT_COLOR

        for column_number, source_color in DOR_COLUMN_SOURCE_COLORS.items():
            dor_sheet.Cells(dor_row, column_number).Interior.Color = source_color

    except Exception:
        if row_inserted:
            dor_sheet.Rows(f"{dor_row}:{dor_row}").Delete(Shift=XL_UP)
        else:
            dor_sheet.Range(
                dor_sheet.Cells(dor_row, 1),
                dor_sheet.Cells(dor_row, 15),
            ).ClearContents()
        raise

    return dor_row, row_inserted

def write_matrix(
    destination_sheet,
    first_row: int,
    first_column: int,
    rows: list[list],
) -> None:
    """
    Writes all source rows to Excel in one operation.
    """
    if not rows:
        return

    row_count = len(rows)
    column_count = len(rows[0])

    excel_values = tuple(
        tuple(row)
        for row in rows
    )

    destination_range = destination_sheet.Range(
        destination_sheet.Cells(
            first_row,
            first_column,
        ),
        destination_sheet.Cells(
            first_row + row_count - 1,
            first_column + column_count - 1,
        ),
    )

    destination_range.Value = excel_values
def fill_metadata(
    destination_sheet,
    first_row: int,
    row_count: int,
    destination_headers: dict[str, int],
    week_value: str,
    production_date: date,
) -> None:
    """
    Populates Week, Day, and Production Date for every imported row.
    """
    required_headers = (
        "Week",
        "Day",
        "Production Date",
    )

    missing_headers = [
        header
        for header in required_headers
        if normalize(header)
        not in destination_headers
    ]

    if missing_headers:
        raise RuntimeError(
            "Destination sheet is missing: "
            + ", ".join(missing_headers)
        )

    final_row = (
        first_row
        + row_count
        - 1
    )

    week_column = destination_headers[
        normalize("Week")
    ]

    day_column = destination_headers[
        normalize("Day")
    ]

    production_date_column = (
        destination_headers[
            normalize("Production Date")
        ]
    )

    day_value = production_date.strftime(
        "%A"
    )

    excel_date_serial = (
        production_date - date(1899, 12, 30)
    ).days

    week_range = destination_sheet.Range(
        destination_sheet.Cells(
            first_row,
            week_column,
        ),
        destination_sheet.Cells(
            final_row,
            week_column,
        ),
    )

    day_range = destination_sheet.Range(
        destination_sheet.Cells(
            first_row,
            day_column,
        ),
        destination_sheet.Cells(
            final_row,
            day_column,
        ),
    )

    production_date_range = (
        destination_sheet.Range(
            destination_sheet.Cells(
                first_row,
                production_date_column,
            ),
            destination_sheet.Cells(
                final_row,
                production_date_column,
            ),
        )
    )

    week_range.Value = week_value
    day_range.Value = day_value

    production_date_range.Value2 = excel_date_serial

    production_date_range.NumberFormat = "dd/mm/yyyy"
    production_date_range.NumberFormatLocal = "dd/mm/yyyy"
def fill_imported_row_formulas(
    destination_sheet,
    first_row: int,
    row_count: int,
) -> None:
    """Populates calculated columns for every imported row."""
    if row_count <= 0:
        return

    final_row = first_row + row_count - 1

    formulas_by_column = {
        15: "=RC[-6]-RC[-8]",
        16: "=IF(RC[-7]-RC[-9]>0.5,1,0)",
        18: "=XLOOKUP(RC[-5],'VC Value per Report'!C1,'VC Value per Report'!C2)",
    }

    for column_number, formula in formulas_by_column.items():
        formula_range = destination_sheet.Range(
            destination_sheet.Cells(first_row, column_number),
            destination_sheet.Cells(final_row, column_number),
        )
        formula_range.FormulaR1C1 = formula
