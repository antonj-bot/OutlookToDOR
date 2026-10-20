"""Mapping and updating the DOR Comparison worksheet sections."""

from backend.config import (
    DOR_COMPARISON_EXCLUDING_TX_SECTION_TITLE, DOR_COMPARISON_SECTION_TITLE,
    DOR_HEADER_ROWS, DOR_COLUMN_SOURCE_COLORS, IMPORTED_DATA_HIGHLIGHT_COLOR, XL_UP,
)
from backend.excel.destination import insert_destination_rows
from backend.excel.ranges import range_to_rows
from backend.excel.source_reader import has_value, normalize

def get_header_path_map(
    worksheet,
    first_header_row: int,
    last_header_row: int,
    first_column: int,
    last_column: int,
) -> dict[str, list[int]]:
    """Maps normalized multi-row headers to their worksheet columns."""
    header_rows = range_to_rows(
        worksheet.Range(
            worksheet.Cells(first_header_row, first_column),
            worksheet.Cells(last_header_row, last_column),
        ).Value
    )
    expanded_parent_rows = []

    for header_row in header_rows[:-1]:
        expanded_row = []
        current_header = ""
        for value in header_row:
            if has_value(value):
                current_header = normalize(value)
            expanded_row.append(current_header)
        expanded_parent_rows.append(expanded_row)

    header_columns: dict[str, list[int]] = {}
    for column_offset in range(last_column - first_column + 1):
        leaf = ""
        leaf_row_index = -1
        for row_index in range(len(header_rows) - 1, -1, -1):
            if column_offset < len(header_rows[row_index]):
                leaf = normalize(header_rows[row_index][column_offset])
            if leaf:
                leaf_row_index = row_index
                break

        if not leaf:
            continue

        if leaf_row_index == len(header_rows) - 1:
            parent_headers = [
                parent_row[column_offset]
                for parent_row in expanded_parent_rows
                if column_offset < len(parent_row)
                and parent_row[column_offset]
            ]
            header = normalize(" ".join((*parent_headers, leaf)))
        else:
            header = leaf

        header_columns.setdefault(header, []).append(
            first_column + column_offset
        )

    return header_columns
def get_dor_comparison_mapping(
    comparison_sheet,
    dor_sheet,
    section_title: str = DOR_COMPARISON_SECTION_TITLE,
) -> dict:
    """Locates the comparison section and matches its headers to DOR."""
    used_range = comparison_sheet.UsedRange
    rows = range_to_rows(used_range.Value)
    first_row = int(used_range.Row)
    first_column = int(used_range.Column)
    title_row = None
    title_column = None

    for row_offset, row_values in enumerate(rows):
        for column_offset, value in enumerate(row_values):
            if normalize(value) == normalize(section_title):
                title_row = first_row + row_offset
                title_column = first_column + column_offset
                break
        if title_row is not None:
            break

    if title_row is None or title_column is None:
        raise RuntimeError(
            f'The section "{section_title}" was not found '
            f'on "{comparison_sheet.Name}".'
        )

    last_used_column = (
        first_column + int(used_range.Columns.Count) - 1
    )
    section_last_column = last_used_column
    title_values = range_to_rows(
        comparison_sheet.Range(
            comparison_sheet.Cells(title_row, title_column),
            comparison_sheet.Cells(title_row, last_used_column),
        ).Value
    )[0]
    for column_offset, value in enumerate(title_values[1:], start=1):
        if has_value(value):
            section_last_column = title_column + column_offset - 1
            break

    comparison_header_first_row = title_row + 1
    comparison_header_last_row = title_row + len(DOR_HEADER_ROWS)
    comparison_headers = get_header_path_map(
        comparison_sheet,
        comparison_header_first_row,
        comparison_header_last_row,
        title_column,
        section_last_column,
    )

    dor_used_range = dor_sheet.UsedRange
    dor_first_column = int(dor_used_range.Column)
    dor_last_column = (
        dor_first_column + int(dor_used_range.Columns.Count) - 1
    )
    dor_headers = get_header_path_map(
        dor_sheet,
        DOR_HEADER_ROWS[0],
        DOR_HEADER_ROWS[-1],
        dor_first_column,
        dor_last_column,
    )

    matching_columns = []
    for header, comparison_columns in comparison_headers.items():
        dor_columns = dor_headers.get(header, [])
        if not dor_columns:
            continue
        if len(dor_columns) > 1:
            raise RuntimeError(
                f'The DOR header "{header}" is ambiguous.'
            )
        matching_columns.extend(
            (comparison_column, dor_columns[0])
            for comparison_column in comparison_columns
        )

    production_date_column = next(
        (
            comparison_column
            for header, comparison_columns in comparison_headers.items()
            if header == normalize("Production Date")
            for comparison_column in comparison_columns
        ),
        None,
    )
    if production_date_column is None:
        raise RuntimeError(
            f'The section "{section_title}" does not have '
            'a "Production Date" header.'
        )
    if not matching_columns:
        raise RuntimeError(
            f'No matching headers were found between "DOR" and the '
            f'"{section_title}" section.'
        )

    return {
        "column_pairs": matching_columns,
        "production_date_column": production_date_column,
        "header_last_row": comparison_header_last_row,
        "first_column": min(column for column, _ in matching_columns),
        "last_column": max(column for column, _ in matching_columns),
    }
def append_dor_comparison_record(
    comparison_sheet,
    dor_sheet,
    dor_row: int,
    comparison_mapping: dict,
    highlight_imported_data: bool = False,
    source_column_values: dict[int, object] | None = None,
    target_row: int | None = None,
) -> dict:
    """Copies matching DOR headers into a new comparison section row."""
    date_column = comparison_mapping["production_date_column"]
    used_range = comparison_sheet.UsedRange
    used_last_row = int(used_range.Row) + int(used_range.Rows.Count) - 1
    if target_row is not None:
        row_number = target_row
        row_inserted = False
    else:
        grand_total_row = None
        for row_number in range(
            comparison_mapping["header_last_row"] + 1,
            used_last_row + 1,
        ):
            if normalize(
                comparison_sheet.Cells(row_number, date_column).Value
            ) == "grand total":
                grand_total_row = row_number
                break

        row_inserted = grand_total_row is not None
        if row_inserted:
            row_number = grand_total_row
            insert_destination_rows(comparison_sheet, row_number, 1)
        else:
            last_data_row = int(
                comparison_sheet
                .Cells(comparison_sheet.Rows.Count, date_column)
                .End(XL_UP)
                .Row
            )
            row_number = max(
                comparison_mapping["header_last_row"] + 1,
                last_data_row + 1,
            )

    source_colors = {
        comparison_column: DOR_COLUMN_SOURCE_COLORS[dor_column]
        for comparison_column, dor_column in comparison_mapping["column_pairs"]
        if dor_column in DOR_COLUMN_SOURCE_COLORS
    }

    try:
        for comparison_column, dor_column in comparison_mapping["column_pairs"]:
            source_value = (
                source_column_values[dor_column]
                if source_column_values is not None
                and dor_column in source_column_values
                else dor_sheet.Cells(dor_row, dor_column).Value2
            )
            comparison_sheet.Cells(
                row_number,
                comparison_column,
            ).Value = source_value

        if highlight_imported_data:
            comparison_sheet.Range(
                comparison_sheet.Cells(
                    row_number,
                    comparison_mapping["first_column"],
                ),
                comparison_sheet.Cells(
                    row_number,
                    comparison_mapping["last_column"],
                ),
            ).Interior.Color = IMPORTED_DATA_HIGHLIGHT_COLOR
        for comparison_column, source_color in source_colors.items():
            comparison_sheet.Cells(
                row_number,
                comparison_column,
            ).Interior.Color = source_color
    except Exception:
        if row_inserted:
            comparison_sheet.Rows(
                f"{row_number}:{row_number}"
            ).Delete(Shift=XL_UP)
        else:
            comparison_sheet.Range(
                comparison_sheet.Cells(
                    row_number,
                    comparison_mapping["first_column"],
                ),
                comparison_sheet.Cells(
                    row_number,
                    comparison_mapping["last_column"],
                ),
            ).ClearContents()
        raise

    return {
        "first_row": row_number,
        "last_row": row_number,
        "first_column": comparison_mapping["first_column"],
        "last_column": comparison_mapping["last_column"],
        "row_inserted": row_inserted,
        "source_colors": source_colors,
    }
