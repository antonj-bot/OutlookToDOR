"""Utilities for parsing source workbook data, detecting headers and dates,
filtering rows by work type, and preparing worksheet data for import.
"""

import difflib
import re
import threading
from datetime import date, datetime
from pathlib import Path

from backend.config import ALLOWED_EXCEL_EXTENSIONS, WORK_TYPE_MATCH_VALUES
from backend.exceptions import check_cancelled
from backend.logging_setup import logger
from backend.excel.ranges import range_to_rows

def normalize(value) -> str:
    """
    Normalizes headers and other text for comparison.

    Examples:
        "Work Type" -> "work type"
        " WORK   TYPE " -> "work type"
        "PRODUCTION DATE" -> "production date"
    """
    if value is None:
        return ""

    text = str(value)
    text = text.replace("\r", " ")
    text = text.replace("\n", " ")
    text = re.sub(r"\s+", " ", text)

    return text.strip().casefold()
def has_value(value) -> bool:
    """
    Returns True when a cell contains meaningful data.
    """
    if value is None:
        return False

    if isinstance(value, str):
        return bool(value.strip())

    return True
def workbook_stem(name: str) -> str:
    """
    Removes the file extension from a workbook name.

    This lets the app match:

        Ecrash_Jan 2026_DOR_V2_Anton

    against:

        Ecrash_Jan 2026_DOR_V2_Anton.xlsx
    """
    return Path(str(name)).stem.casefold()
def is_excel_file(file_name: str) -> bool:
    extension = Path(file_name).suffix.casefold()

    return extension in ALLOWED_EXCEL_EXTENSIONS
def extract_date_from_filename(
    file_name: str,
) -> date | None:
    """
    Detects dates in attachment filenames.

    Supported examples:

        Report 2025-01-03.xlsx
        Report_2025_01_03.xlsx
        Report.2025.01.03.xlsx
        Report 20250103.xlsx
    """
    file_stem = Path(file_name).stem

    patterns = (
        r"(?<!\d)"
        r"(20\d{2})[-_. ]"
        r"(\d{1,2})[-_. ]"
        r"(\d{1,2})"
        r"(?!\d)",

        r"(?<!\d)"
        r"(20\d{2})"
        r"(\d{2})"
        r"(\d{2})"
        r"(?!\d)",
    )

    for pattern in patterns:
        match = re.search(
            pattern,
            file_stem,
        )

        if not match:
            continue

        try:
            year = int(match.group(1))
            month = int(match.group(2))
            day_number = int(match.group(3))

            return date(
                year,
                month,
                day_number,
            )

        except ValueError:
            continue

    return None
def parse_date(value: str) -> date:
    """
    Parses the date entered in the UI.
    """
    formats = (
        "%Y-%m-%d",
        "%m/%d/%Y",
        "%m/%d/%y",
    )

    for date_format in formats:
        try:
            return datetime.strptime(
                value.strip(),
                date_format,
            ).date()

        except ValueError:
            continue

    raise ValueError(
        "Production Date must be entered as "
        "YYYY-MM-DD or MM/DD/YYYY."
    )
    
def fallback_week(
    production_date: date,
) -> str:
    """
    Calculates the week within the production date's month.

    Days 1 through 7 = Week 1
    Days 8 through 14 = Week 2
    Days 15 through 21 = Week 3
    Days 22 through the end of the month = Week 4

    The importer first tries to reuse the Week value for an existing
    Production Date in the destination workbook.

    This fallback is used only when the date does not already exist
    and the user has not entered a Week override.
    """
    week_number = min(
        ((production_date.day - 1) // 7) + 1,
        4,
    )

    return f"Week {week_number}"

def to_python_date(value) -> date | None:
    """
    Converts an Excel date value into a Python date.
    """
    if isinstance(value, datetime):
        return value.date()

    if isinstance(value, date):
        return value

    if isinstance(value, str):
        try:
            return parse_date(value)
        except ValueError:
            return None

    return None
def find_source_header_row(
    rows: list[list],
    maximum_rows_to_check: int = 200,
) -> tuple[int, dict[str, int]]:
    """
    Finds the source header row in ordinary worksheet cells.

    The source does NOT need to be formatted as an Excel Table.

    Returns:
        header row index
        detected normalized headers and their column indexes
    """

    # Alternative header names that may exist in the attachment.
    work_type_names = {
        "work type",
        "worktype",
        "work category",
        "worktype name",
    }

    username_names = {
        "username",
        "user name",
        "user",
        "employee username",
        "employee name",
        "agent",
        "agent name",
    }

    rows_to_check = min(
        len(rows),
        maximum_rows_to_check,
    )

    best_candidate = None
    best_score = 0

    for row_index in range(rows_to_check):
        row = rows[row_index]

        header_positions = {}

        for column_index, value in enumerate(row):
            header = normalize(value)

            if not header:
                continue

            if header not in header_positions:
                header_positions[header] = column_index

        detected_work_type = next(
            (
                header
                for header in header_positions
                if header in work_type_names
            ),
            None,
        )

        detected_username = next(
            (
                header
                for header in header_positions
                if header in username_names
            ),
            None,
        )

        score = 0

        if detected_work_type:
            score += 1

        if detected_username:
            score += 1

        if score > best_score:
            best_score = score
            best_candidate = (
                row_index,
                header_positions,
            )

        if detected_work_type and detected_username:
            return (
                row_index,
                header_positions,
            )

    if best_candidate:
        candidate_row_index = best_candidate[0]

        detected_values = [
            str(value).strip()
            for value in rows[candidate_row_index]
            if has_value(value)
        ]

        raise RuntimeError(
            "A possible header row was found, but it did not "
            "contain both the Work Type and Username headers.\n\n"
            f"Possible header row: {candidate_row_index + 1}\n"
            f"Detected values:\n"
            + "\n".join(detected_values)
        )

    raise RuntimeError(
        "No possible header row was found in the first "
        f"{rows_to_check} populated rows."
    )
def find_matching_header(
    header_positions: dict[str, int],
    accepted_names: set[str],
    display_name: str,
) -> tuple[str, int]:

    # Exact match first
    for header_name, column_index in header_positions.items():
        if header_name in accepted_names:
            return (
                header_name,
                column_index,
            )

    # Fuzzy match second
    for header_name, column_index in header_positions.items():

        normalized = (
            header_name
            .replace("-", " ")
            .replace(":", "")
            .replace("_", " ")
            .strip()
        )

        best_match = difflib.get_close_matches(
            normalized,
            accepted_names,
            n=1,
            cutoff=0.80,
        )

        if best_match:
            logger.info(
                "Fuzzy matched '%s' to '%s'",
                header_name,
                best_match[0],
            )

            return (
                header_name,
                column_index,
            )

    raise RuntimeError(
        f'Could not find "{display_name}" column.\n\n'
        f"Headers found:\n"
        + "\n".join(header_positions.keys())
    )
def work_type_matches(
    value,
    selected_queue: str,
) -> bool:
    """
    Checks whether Work Type matches Ecrash or CGF.
    """
    work_type_text = normalize(value)

    if not work_type_text:
        return False

    accepted_values = WORK_TYPE_MATCH_VALUES[
        selected_queue
    ]

    return any(
        accepted_value in work_type_text
        for accepted_value in accepted_values
    )
def read_filtered_source_rows(
    source_sheet,
    selected_queue: str,
    cancel_event: threading.Event | None = None,
) -> tuple[list[str], list[list]]:
    """
    Reads ordinary worksheet cells.

    Rules:
    1. Finds a header row.
    2. Copies every data row.
    3. Excludes only the Username column.
    4. Preserves blank cells and columns.
    5. Returns (headers, rows).
    """

    check_cancelled(cancel_event)
    used_range = source_sheet.UsedRange

    source_rows = range_to_rows(
        used_range.Value
    )
    check_cancelled(cancel_event)

    if not source_rows:
        raise RuntimeError(
            f'Source sheet "{source_sheet.Name}" contains no data.'
        )

    (
        header_row_index,
        header_positions,
    ) = find_source_header_row(
        source_rows,
        maximum_rows_to_check=200,
    )

    source_headers = source_rows[
        header_row_index
    ]

    work_type_names = {
        "work type",
        "worktype",
        "work category",
        "worktype name",
    }

    username_names = {
        "username",
        "user name",
        "user",
        "employee username",
        "employee name",
        "agent",
        "agent name",
    }

    _, work_type_index = find_matching_header(
        header_positions,
        work_type_names,
        "Work Type",
    )

    _, username_index = find_matching_header(
        header_positions,
        username_names,
        "Username",
    )

    # Keep every source column except Username so the copied values stay
    # aligned with the attachment's original column order.
    included_column_indexes = [
        column_index
        for column_index in range(len(source_headers))
        if column_index != username_index
    ]

    included_headers = [
        str(source_headers[column_index]).strip()
        if has_value(source_headers[column_index])
        else ""
        for column_index in included_column_indexes
    ]

    if not included_column_indexes:
        raise RuntimeError(
            "No columns remain after excluding Username."
        )

    filtered_rows = []

    data_start_index = (
        header_row_index + 1
    )

    for source_row in source_rows[
        data_start_index:
    ]:
        check_cancelled(cancel_event)

        work_type_value = (
            source_row[work_type_index]
            if work_type_index < len(source_row)
            else None
        )

        if not work_type_matches(
            work_type_value,
            selected_queue,
        ):
            continue

        output_row = []

        for column_index in included_column_indexes:

            if column_index < len(source_row):
                cell_value = source_row[
                    column_index
                ]
            else:
                cell_value = None

            output_row.append(
                cell_value
            )

        # Skip completely empty rows
        if not any(
            has_value(value)
            for value in output_row
        ):
            continue

        filtered_rows.append(
            output_row
        )

    if not filtered_rows:
        raise RuntimeError(
            f'No {selected_queue} data rows were found after '
            "the source header."
        )

    logger.info(
        "Returning %d headers and %d filtered rows",
        len(included_headers),
        len(filtered_rows),
    )

    return (
        included_headers,
        filtered_rows,
    )
