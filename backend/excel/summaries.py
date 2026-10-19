"""Utilities for reading DOR summary metrics, tracking their source cells,
and locating existing production dates and week values.
"""

from datetime import date, datetime
import threading

from backend.config import DOR_SOURCE_COLORS, XL_UP
from backend.exceptions import check_cancelled
from backend.excel.ranges import range_to_rows
from backend.excel.source_reader import has_value, normalize, to_python_date

def get_dor_summary_values(
    summary_sheet,
    production_date: date,
    cancel_event: threading.Event | None = None,
) -> tuple[tuple, list[dict]]:
    """Looks up DOR metrics and their source cells by date."""
    check_cancelled(cancel_event)
    used_range = summary_sheet.UsedRange
    rows = range_to_rows(used_range.Value)
    check_cancelled(cancel_event)
    first_excel_row = int(used_range.Row)
    first_excel_column = int(used_range.Column)

    section_titles = {
        "auto vs manual",
        "total completed & average per hour",
        "past 12-hr tat",
        "manual reports",
        "manual ave/hr",
        "auto extraction",
        "auto extraction ave/hr",
    }
    sections = {}

    for row_index, row_values in enumerate(rows):
        for column_index, value in enumerate(row_values):
            normalized_value = normalize(value)
            if normalized_value in section_titles and normalized_value not in sections:
                following_titles = [
                    other_column
                    for other_column, other_value in enumerate(row_values)
                    if other_column > column_index
                    and normalize(other_value) in section_titles
                ]
                sections[normalized_value] = (
                    row_index,
                    column_index,
                    min(following_titles) - 1
                    if following_titles
                    else len(row_values) - 1,
                )

    def find_date_header(section):
        title_row, first_column, last_column = section
        for row_index in range(title_row + 1, min(title_row + 15, len(rows))):
            for column_index in range(first_column, min(last_column + 1, len(rows[row_index]))):
                header_text = normalize(rows[row_index][column_index])
                if header_text == "date":
                    return row_index, column_index
                if (
                    column_index == first_column
                    and header_text in {"state", "row labels"}
                    and any(
                        to_python_date(rows[data_row][column_index])
                        is not None
                        for data_row in range(
                            row_index + 1,
                            min(row_index + 4, len(rows)),
                        )
                        if column_index < len(rows[data_row])
                    )
                ):
                    return row_index, column_index
        return None

    date_headers = {
        section_name: find_date_header(section)
        for section_name, section in sections.items()
    }

    def find_data_row(section_name, date_header):
        title_row, _first_column, _last_column = sections[section_name]
        header_row, date_column = date_header
        timezone_match = None

        for row_index in range(header_row + 1, len(rows)):
            check_cancelled(cancel_event)
            if date_column >= len(rows[row_index]):
                continue
            value = rows[row_index][date_column]
            if to_python_date(value) == production_date:
                return row_index
            if (
                timezone_match is None
                and isinstance(value, datetime)
                and value.tzinfo is not None
                and value.astimezone().date() == production_date
            ):
                timezone_match = row_index

        if timezone_match is not None:
            return timezone_match

        raise RuntimeError(
            "Production date "
            f"{production_date.strftime('%m/%d/%Y')} was not found in "
            f'the "{section_name}" table on "{summary_sheet.Name}".'
        )

    def find_direct_column(section, target_header, header_row):
        title_row, first_column, last_column = section
        for row_index in range(title_row + 1, header_row + 1):
            for column_index in range(
                first_column,
                min(last_column + 1, len(rows[row_index])),
            ):
                if normalize(rows[row_index][column_index]) == target_header:
                    return column_index
        raise RuntimeError(
            f'Could not find "{target_header}" in the '
            f'"{summary_sheet.Name}" summary table.'
        )

    def find_pass_column(section, pass_name, date_header):
        title_row, first_column, last_column = section
        date_header_row, _date_column = date_header

        for row_index in range(title_row + 1, date_header_row):
            parent_row = rows[row_index]
            for parent_column in range(
                first_column,
                min(last_column + 1, len(parent_row)),
            ):
                if normalize(parent_row[parent_column]) != "completed report":
                    continue

                group_end = next(
                    (
                        other_column
                        for other_column in range(parent_column + 1, len(parent_row))
                        if has_value(parent_row[other_column])
                    ),
                    last_column + 1,
                )
                for header_row_index in range(row_index + 1, date_header_row + 1):
                    header_row = rows[header_row_index]
                    for column_index in range(
                        parent_column,
                        min(group_end, len(header_row)),
                    ):
                        if normalize(header_row[column_index]) == pass_name:
                            return column_index

        raise RuntimeError(
            f'Could not find "Completed Report {pass_name}" in '
            f'"{summary_sheet.Name}".'
        )

    main_section = sections.get("total completed & average per hour")
    tat_section = sections.get("past 12-hr tat")
    auto_manual_section = sections.get("auto vs manual")
    manual_reports_name = next(
        (
            name
            for name in ("manual reports", "manual ave/hr")
            if name in sections and date_headers.get(name) is not None
        ),
        "manual reports",
    )
    auto_extraction_name = next(
        (
            name
            for name in ("auto extraction", "auto extraction ave/hr")
            if name in sections and date_headers.get(name) is not None
        ),
        "auto extraction",
    )
    manual_reports_section = sections.get(manual_reports_name)
    auto_extraction_section = sections.get(auto_extraction_name)
    if any(
        section is None
        for section in (
            main_section,
            tat_section,
            auto_manual_section,
            manual_reports_section,
            auto_extraction_section,
        )
    ):
        raise RuntimeError(
            f'Could not find the DOR summary sections on "{summary_sheet.Name}".'
        )

    main_date_header = date_headers["total completed & average per hour"]
    tat_date_header = date_headers["past 12-hr tat"]
    auto_manual_date_header = date_headers["auto vs manual"]
    manual_reports_date_header = date_headers[manual_reports_name]
    auto_extraction_date_header = date_headers[auto_extraction_name]
    if any(
        date_header is None
        for date_header in (
            main_date_header,
            tat_date_header,
            auto_manual_date_header,
            manual_reports_date_header,
            auto_extraction_date_header,
        )
    ):
        raise RuntimeError(
            f'Could not find date columns in the DOR summary tables on "{summary_sheet.Name}".'
        )

    main_data_row = find_data_row(
        "total completed & average per hour",
        main_date_header,
    )
    tat_data_row = find_data_row(
        "past 12-hr tat",
        tat_date_header,
    )
    auto_manual_data_row = find_data_row(
        "auto vs manual",
        auto_manual_date_header,
    )
    manual_reports_data_row = find_data_row(
        manual_reports_name,
        manual_reports_date_header,
    )
    auto_extraction_data_row = find_data_row(
        auto_extraction_name,
        auto_extraction_date_header,
    )

    metric_columns = [
        (
            "completed report pass 1",
            find_pass_column(main_section, "pass 1", main_date_header),
            main_data_row,
        ),
        (
            "completed report pass 2",
            find_pass_column(main_section, "pass 2", main_date_header),
            main_data_row,
        ),
        (
            "total elapsed time",
            find_direct_column(
                main_section,
                "total elapsed time",
                main_date_header[0],
            ),
            main_data_row,
        ),
        (
            "average report/hr",
            find_direct_column(
                main_section,
                "average report/hr",
                main_date_header[0],
            ),
            main_data_row,
        ),
        (
            "completed reports",
            find_direct_column(
                tat_section,
                "completed reports",
                tat_date_header[0],
            ),
            tat_data_row,
        ),
        (
            "late %",
            find_direct_column(
                tat_section,
                "late %",
                tat_date_header[0],
            ),
            tat_data_row,
        ),
        (
            "ontime%",
            find_direct_column(
                tat_section,
                "ontime%",
                tat_date_header[0],
            ),
            tat_data_row,
        ),
        (
            "auto vs manual manual",
            find_direct_column(
                auto_manual_section,
                "manual",
                auto_manual_date_header[0],
            ),
            auto_manual_data_row,
        ),
        (
            "manual reports average report/hr",
            find_direct_column(
                manual_reports_section,
                "average report/hr",
                manual_reports_date_header[0],
            ),
            manual_reports_data_row,
        ),
        (
            "auto vs manual auto extraction",
            find_direct_column(
                auto_manual_section,
                "auto extraction",
                auto_manual_date_header[0],
            ),
            auto_manual_data_row,
        ),
        (
            "auto extraction average report/hr",
            find_direct_column(
                auto_extraction_section,
                "average report/hr",
                auto_extraction_date_header[0],
            ),
            auto_extraction_data_row,
        ),
    ]

    summary_values = []
    source_ranges = []
    for source_field, column_index, row_index in metric_columns:
        if column_index >= len(rows[row_index]):
            raise RuntimeError(
                f'No value for "{source_field}" was found for '
                f"{production_date.strftime('%m/%d/%Y')}."
            )

        summary_values.append(rows[row_index][column_index])
        excel_row = first_excel_row + row_index
        excel_column = first_excel_column + column_index
        source_ranges.append(
            {
                "sheet": str(summary_sheet.Name),
                "first_row": excel_row,
                "last_row": excel_row,
                "first_column": excel_column,
                "last_column": excel_column,
                "kind": "pivot",
                "source_field": source_field,
                "source_color": DOR_SOURCE_COLORS[source_field],
                "pivot_source_date": production_date.isoformat(),
            }
        )

    return tuple(summary_values), source_ranges
def find_existing_week(
    destination_sheet,
    production_date: date,
    destination_headers: dict[str, int],
) -> str | None:
    """
    Checks whether Production Date already exists in the destination.

    When it exists, the Week value from that row is reused.
    """
    production_date_column = (
        destination_headers.get(
            normalize("Production Date")
        )
    )

    week_column = (
        destination_headers.get(
            normalize("Week")
        )
    )

    if (
        production_date_column is None
        or week_column is None
    ):
        return None

    last_used_row = int(
        destination_sheet
        .Cells(
            destination_sheet.Rows.Count,
            production_date_column,
        )
        .End(XL_UP)
        .Row
    )

    first_column = min(
        production_date_column,
        week_column,
    )
    last_column = max(
        production_date_column,
        week_column,
    )

    values = range_to_rows(
        destination_sheet.Range(
            destination_sheet.Cells(
                2,
                first_column,
            ),
            destination_sheet.Cells(
                last_used_row,
                last_column,
            ),
        ).Value
    )

    production_date_offset = (
        production_date_column - first_column
    )
    week_offset = week_column - first_column

    for row_values in values:
        if production_date_offset >= len(row_values):
            continue

        existing_date = to_python_date(
            row_values[production_date_offset]
        )

        if existing_date != production_date:
            continue

        if week_offset >= len(row_values):
            continue

        existing_week = row_values[week_offset]

        if has_value(existing_week):
            return str(existing_week).strip()

    return None
