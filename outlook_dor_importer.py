from __future__ import annotations

import ctypes
import base64
import queue
import re
import subprocess
import tempfile
import threading
import traceback
import time
import logging
import difflib
from ctypes import wintypes

from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path

import tkinter as tk
from tkinter import ttk, messagebox

import pythoncom
import win32com.client


# ============================================================
# CONFIGURATION
# ============================================================

DESTINATION_WORKBOOK_NAMES = {
    "Ecrash": "Ecrash_Jan 2026_DOR_V2_Anton",
    "CGF": "CGF_Jan 2026_DOR_V2_Anton",
}


DESTINATION_SHEETS = {
    "Ecrash": "Ecrash Raw",
    "CGF": "CGF Raw",
}


# The filter looks for these words inside the Work Type value.
# Matching is not case-sensitive.
WORK_TYPE_MATCH_VALUES = {
    "Ecrash": (
        "ecrash",
        "e crash",
    ),
    "CGF": (
        "cgf",
    ),
}


ALLOWED_EXCEL_EXTENSIONS = {
    ".xlsx",
    ".xlsm",
    ".xlsb",
    ".xls",
}


# Number of recent Outlook messages to inspect.
OUTLOOK_MESSAGE_LIMIT = 500


# Only attachments whose filenames contain this phrase are eligible.
REQUIRED_ATTACHMENT_KEYWORD = (
    "individual keyed pass per user"
)


# Outlook Inbox constant.
OL_FOLDER_INBOX = 6


# Excel constants.
XL_UP = -4162
XL_TO_LEFT = -4159
XL_DOWN = -4121

# Windows taskbar notification constants.
FLASHW_TRAY = 0x00000002
FLASHW_TIMERNOFG = 0x0000000C


class FlashWindowInfo(ctypes.Structure):
    _fields_ = [
        ("cbSize", wintypes.UINT),
        ("hwnd", wintypes.HWND),
        ("dwFlags", wintypes.DWORD),
        ("uCount", wintypes.UINT),
        ("dwTimeout", wintypes.DWORD),
    ]

# ============================================================
# PROCESS TIMING
# ============================================================

logger = logging.getLogger("outlook_dor_importer")


@dataclass
class ProcessTiming:
    name: str
    duration_seconds: float
    succeeded: bool
    error: str = ""


class SessionTimer:
    def __init__(self):
        self.started_at = time.perf_counter()
        self.processes: list[ProcessTiming] = []
        self._lock = threading.Lock()

    def start(self) -> float:
        return time.perf_counter()

    def finish(
        self,
        name: str,
        started_at: float,
        succeeded: bool,
        error: Exception | None = None,
    ) -> float:
        duration_seconds = time.perf_counter() - started_at
        error_text = "" if error is None else str(error)
        timing = ProcessTiming(
            name=name,
            duration_seconds=duration_seconds,
            succeeded=succeeded,
            error=error_text,
        )

        with self._lock:
            self.processes.append(timing)

        logger.info(
            "Process '%s' %s in %.2f seconds%s",
            name,
            "completed" if succeeded else "failed",
            duration_seconds,
            f": {error_text}" if error_text else "",
        )

        return duration_seconds

    def summary(self) -> str:
        with self._lock:
            processes = list(self.processes)

        session_duration = time.perf_counter() - self.started_at

        if not processes:
            return (
                "Session summary\n"
                f"Total session time: {session_duration:.2f} seconds\n"
                "No processes were completed."
            )

        lines = [
            "Session summary",
            f"Total session time: {session_duration:.2f} seconds",
        ]

        for process in processes:
            result = "completed" if process.succeeded else "failed"
            line = (
                f"- {process.name}: {process.duration_seconds:.2f} "
                f"seconds ({result})"
            )

            if process.error:
                line += f" - {process.error}"

            lines.append(line)

        return "\n".join(lines)

# ============================================================
# ATTACHMENT INFORMATION
# ============================================================

@dataclass
class AttachmentInfo:
    file_name: str
    received_time: datetime
    sender_name: str
    subject: str
    message_entry_id: str
    store_id: str
    attachment_index: int

    @property
    def display_name(self) -> str:
        received_text = self.received_time.strftime(
            "%Y-%m-%d %H:%M"
        )

        return (
            f"{received_text} | "
            f"{self.file_name} | "
            f"{self.sender_name} | "
            f"{self.subject}"
        )


# ============================================================
# GENERAL HELPERS
# ============================================================

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

# ============================================================
# DATE AND WEEK HELPERS
# ============================================================

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
    Fallback week calculation.

    January 1 through January 7 = Week 1
    January 8 through January 14 = Week 2

    The importer first tries to find the date in the destination
    workbook and reuse the Week value found there.

    This function is only used when the date does not already exist
    and the user has not entered a Week override.
    """
    first_day_of_year = date(
        production_date.year,
        1,
        1,
    )

    days_from_start = (
        production_date - first_day_of_year
    ).days

    week_number = (
        days_from_start // 7
    ) + 1

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

# ============================================================
# OUTLOOK FUNCTIONS
# ============================================================

def get_outlook_namespace():
    """
    Connects to classic Outlook desktop.
    """
    outlook = win32com.client.Dispatch(
        "Outlook.Application"
    )

    return outlook.GetNamespace("MAPI")


def get_outlook_folder(
    namespace,
    folder_path: str = "Inbox",
):
    """
    Resolves Inbox or a folder path such as Inbox\Reports.

    A mailbox-rooted path such as Mailbox Name\Reports is also supported.
    """
    requested_path = folder_path.strip() or "Inbox"
    path_parts = [
        part.strip()
        for part in re.split(r"[\\/]+", requested_path)
        if part.strip()
    ]

    if not path_parts:
        path_parts = ["Inbox"]

    if path_parts[0].casefold() == "inbox":
        current_folder = namespace.GetDefaultFolder(
            OL_FOLDER_INBOX
        )
        path_parts = path_parts[1:]
    else:
        current_folder = None

        for folder in namespace.Folders:
            if str(folder.Name).strip().casefold() == path_parts[0].casefold():
                current_folder = folder
                break

        if current_folder is None:
            raise RuntimeError(
                f'Outlook mailbox or folder root "{path_parts[0]}" '
                "was not found."
            )

        path_parts = path_parts[1:]

    for requested_name in path_parts:
        next_folder = None

        for folder in current_folder.Folders:
            if (
                str(folder.Name).strip().casefold()
                == requested_name.casefold()
            ):
                next_folder = folder
                break

        if next_folder is None:
            raise RuntimeError(
                f'Outlook folder "{requested_name}" was not found '
                f'under "{current_folder.Name}".'
            )

        current_folder = next_folder

    return current_folder



def list_excel_attachments(
    file_name_filter: str = "",
    folder_path: str = "Inbox",
) -> list[AttachmentInfo]:
    """
    Searches recent Outlook messages in the selected folder for matching
    Excel attachments.

    Every result must contain REQUIRED_ATTACHMENT_KEYWORD in its filename.
    The optional file_name_filter narrows those results further.

    Results are returned newest first.
    """
    namespace = get_outlook_namespace()

    outlook_folder = get_outlook_folder(
        namespace,
        folder_path,
    )

    messages = outlook_folder.Items

    # Newest messages first.
    messages.Sort(
        "[ReceivedTime]",
        True,
    )

    requested_text = (
        file_name_filter
        .strip()
        .casefold()
    )

    results: list[AttachmentInfo] = []

    for message_number, message in enumerate(
        messages,
        start=1,
    ):
        if message_number > OUTLOOK_MESSAGE_LIMIT:
            break

        try:
            attachment_count = int(
                message.Attachments.Count
            )
        except Exception:
            continue

        if attachment_count == 0:
            continue

        try:
            received_time = (
                message
                .ReceivedTime
                .replace(tzinfo=None)
            )
        except Exception:
            received_time = datetime.min

        try:
            sender_name = str(
                message.SenderName or ""
            )
        except Exception:
            sender_name = ""

        try:
            subject = str(
                message.Subject or ""
            )
        except Exception:
            subject = ""

        try:
            message_entry_id = str(
                message.EntryID
            )

            store_id = str(
                message.Parent.StoreID
            )
        except Exception:
            continue

        for attachment_index in range(
            1,
            attachment_count + 1,
        ):
            try:
                attachment = (
                    message
                    .Attachments
                    .Item(attachment_index)
                )

                file_name = str(
                    attachment.FileName
                )

            except Exception:
                continue

            if not is_excel_file(file_name):
                continue

            file_name_text = file_name.casefold()

            if (
                REQUIRED_ATTACHMENT_KEYWORD
                not in file_name_text
            ):
                continue

            if (
                requested_text
                and requested_text not in file_name_text
            ):
                continue

            results.append(
                AttachmentInfo(
                    file_name=file_name,
                    received_time=received_time,
                    sender_name=sender_name,
                    subject=subject,
                    message_entry_id=message_entry_id,
                    store_id=store_id,
                    attachment_index=attachment_index,
                )
            )

    results.sort(
        key=lambda item: item.received_time,
        reverse=True,
    )

    return results


def save_attachment(
    attachment_info: AttachmentInfo,
) -> Path:
    """
    Downloads the selected Outlook attachment to a temporary folder.

    An existing file is not overwritten. A number is added instead.
    """
    namespace = get_outlook_namespace()

    message = namespace.GetItemFromID(
        attachment_info.message_entry_id,
        attachment_info.store_id,
    )

    attachment = message.Attachments.Item(
        attachment_info.attachment_index
    )

    download_folder = (
        Path(tempfile.gettempdir())
        / "OutlookDORImporter"
    )

    download_folder.mkdir(
        parents=True,
        exist_ok=True,
    )

    output_path = (
        download_folder
        / attachment_info.file_name
    )

    counter = 1

    while output_path.exists():
        original_path = Path(
            attachment_info.file_name
        )

        output_path = (
            download_folder
            / (
                f"{original_path.stem}_"
                f"{counter}"
                f"{original_path.suffix}"
            )
        )

        counter += 1

    attachment.SaveAsFile(
        str(output_path)
    )

    if not output_path.exists():
        raise RuntimeError(
            "Outlook did not save the selected attachment."
        )

    return output_path

# ============================================================
# EXCEL CONNECTION HELPERS
# ============================================================

def find_open_workbook(
    excel,
    requested_name: str,
):
    """
    Finds an already-open Excel workbook.
    """
    requested_stem = workbook_stem(
        requested_name
    )

    for workbook in excel.Workbooks:
        current_stem = workbook_stem(
            workbook.Name
        )

        if current_stem == requested_stem:
            return workbook

    open_workbook_names = [
        str(workbook.Name)
        for workbook in excel.Workbooks
    ]

    if open_workbook_names:
        workbook_list = "\n".join(
            open_workbook_names
        )
    else:
        workbook_list = "None"

    raise RuntimeError(
        f'Workbook "{requested_name}" is not open in Excel.'
        f"\n\nOpen workbooks:\n{workbook_list}"
    )


def find_sheet(
    workbook,
    sheet_name: str,
):
    """
    Finds a sheet without depending on capitalization.
    """
    requested_name = (
        sheet_name
        .strip()
        .casefold()
    )

    for worksheet in workbook.Worksheets:
        current_name = (
            str(worksheet.Name)
            .strip()
            .casefold()
        )

        if current_name == requested_name:
            return worksheet

    available_sheets = [
        str(worksheet.Name)
        for worksheet in workbook.Worksheets
    ]

    raise RuntimeError(
        f'Sheet "{sheet_name}" was not found in '
        f'"{workbook.Name}".'
        f"\n\nAvailable sheets:\n"
        + "\n".join(available_sheets)
    )


def range_to_rows(value) -> list[list]:
    """
    Converts an Excel COM Range.Value result into normal Python rows.
    """
    if value is None:
        return []

    if not isinstance(value, tuple):
        return [[value]]

    if (
        value
        and not isinstance(
            value[0],
            tuple,
        )
    ):
        return [list(value)]

    return [
        list(row)
        for row in value
    ]


# ============================================================
# SOURCE DATA FILTERING
# ============================================================

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

    used_range = source_sheet.UsedRange

    source_rows = range_to_rows(
        used_range.Value
    )

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

# ============================================================
# DESTINATION WORKBOOK HELPERS
# ============================================================

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

    excel_date_value = datetime.combine(
        production_date,
        datetime.min.time(),
    )

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

    production_date_range.Value = (
        excel_date_value
    )

    production_date_range.NumberFormat = (
        "mm/dd/yyyy"
    )


# ============================================================
# MAIN IMPORT FUNCTION
# ============================================================
def find_source_data_sheet(
    source_workbook,
):
    """
    Searches every worksheet for a likely header row.

    Returns the first worksheet containing both a Work Type header
    and a Username-like header.
    """

    checked_sheets = []

    for source_sheet in source_workbook.Worksheets:
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
) -> dict:
    """
    Opens the attachment as read-only and imports the filtered rows
    into the already-open destination workbook.
    """
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

    source_workbook = None

    try:
        source_workbook = excel.Workbooks.Open(
            str(source_path),
            UpdateLinks=0,
            ReadOnly=True,
            IgnoreReadOnlyRecommended=True,
        )

        source_sheet = find_source_data_sheet(
            source_workbook
        )
        
        result = read_filtered_source_rows(
            source_sheet,
            selected_queue,
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

        rows_inserted = False

        try:
            if row_override is not None:
                insert_destination_rows(
                    destination_sheet,
                    first_destination_row,
                    len(source_rows),
                )
                rows_inserted = True

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

        destination_workbook.Save()

        return {
            "rows": len(source_rows),
            "columns": len(source_headers),
            "first_row": first_destination_row,
            "last_row": (
                first_destination_row
                + len(source_rows)
                - 1
            ),
            "week": week_value,
            "week_source": week_source,
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
        }

    finally:
        if source_workbook is not None:
            source_workbook.Close(
                SaveChanges=False
            )

# ============================================================
# TKINTER APPLICATION
# ============================================================

class OutlookDORImporterApp(tk.Tk):

    def __init__(self):
        super().__init__()

        self.title(
            "Outlook DOR Importer"
        )

        self.geometry(
            "1050x680"
        )

        self.minsize(
            900,
            580,
        )

        self.attachments: list[
            AttachmentInfo
        ] = []

        self.queue_var = tk.StringVar(
            value="Ecrash"
        )

        self.file_filter_var = tk.StringVar()

        self.outlook_folder_var = tk.StringVar(
            value="Inbox"
        )

        self.destination_var = tk.StringVar(
            value=DESTINATION_WORKBOOK_NAMES[
                "Ecrash"
            ]
        )

        self.production_date_var = (
            tk.StringVar()
        )

        self.week_override_var = (
            tk.StringVar()
        )

        self.row_override_var = (
            tk.StringVar()
        )

        self.status_var = tk.StringVar(
            value=(
                "Ready. Select Refresh Attachments."
            )
        )

        self.progress_target = 0.0
        self.progress_animation_id = None
        self.session_timer = SessionTimer()
        self.ui_events = queue.Queue()
        self.closing = False
        self.action_active = False

        self.protocol(
            "WM_DELETE_WINDOW",
            self.close_application,
        )

        self.create_widgets()
        self.after(
            50,
            self.process_ui_events,
        )

    def post_ui_event(
        self,
        callback,
        *args,
    ):
        self.ui_events.put(
            (callback, args)
        )

    def process_ui_events(self):
        if self.closing:
            return

        while True:
            try:
                callback, args = self.ui_events.get_nowait()
            except queue.Empty:
                break

            callback(*args)

        self.after(
            50,
            self.process_ui_events,
        )

    def animate_progress(
        self,
        target_value: float,
    ):
        self.progress_target = max(
            0.0,
            min(100.0, float(target_value)),
        )

        if self.progress_animation_id is None:
            self.progress_animation_id = self.after(
                16,
                self._animate_progress_step,
            )

    def _animate_progress_step(self):
        current = float(
            self.progress["value"]
        )

        distance = self.progress_target - current

        if abs(distance) <= 0.25:
            current = self.progress_target
            self.progress["value"] = current
            self.progress_percent_var.set(
                f"{round(current)}%"
            )
            self.progress_animation_id = None
            return

        step = max(
            0.5,
            abs(distance) * 0.12,
        )

        current += step if distance > 0 else -step
        self.progress["value"] = current
        self.progress_percent_var.set(
            f"{round(current)}%"
        )

        self.progress_animation_id = self.after(
            16,
            self._animate_progress_step,
        )


    def notify_action_complete(
        self,
        title: str,
        message: str,
    ):
        """
        Shows a Windows toast and taskbar attention when the app is inactive.
        """
        try:
            window_handle = self.winfo_id()
            user32 = ctypes.windll.user32
            is_inactive = (
                self.state() == "iconic"
                or user32.GetForegroundWindow()
                != window_handle
            )

            if not is_inactive:
                return

            self.bell()

            flash_info = FlashWindowInfo(
                cbSize=ctypes.sizeof(FlashWindowInfo),
                hwnd=window_handle,
                dwFlags=FLASHW_TRAY | FLASHW_TIMERNOFG,
                uCount=3,
                dwTimeout=0,
            )

            user32.FlashWindowEx(
                ctypes.byref(flash_info)
            )

            escaped_title = title.replace(
                "'",
                "''",
            )
            escaped_message = message.replace(
                "'",
                "''",
            )

            notification_icon = Path(
                __file__
            ).with_name(
                "outlook_dor_importer.png"
            )

            if notification_icon.exists():
                toast_template = "ToastImageAndText02"
                image_uri = notification_icon.as_uri().replace(
                    "'",
                    "''",
                )
                image_script = f"""
$image = $xml.GetElementsByTagName('image').Item(0)
$image.SetAttribute('src', '{image_uri}')
"""
            else:
                toast_template = "ToastText02"
                image_script = ""

            powershell_script = f"""
[Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime] > $null
$xml = [Windows.UI.Notifications.ToastNotificationManager]::GetTemplateContent([Windows.UI.Notifications.ToastTemplateType]::{toast_template})
$text = $xml.GetElementsByTagName('text')
$text.Item(0).AppendChild($xml.CreateTextNode('{escaped_title}')) > $null
$text.Item(1).AppendChild($xml.CreateTextNode('{escaped_message}')) > $null
{image_script}
$toast = [Windows.UI.Notifications.ToastNotification]::new($xml)
[Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier('AutoDOR').Show($toast)
"""

            encoded_script = base64.b64encode(
                powershell_script.encode("utf-16le")
            ).decode("ascii")

            subprocess.Popen(
                [
                    "powershell.exe",
                    "-NoProfile",
                    "-ExecutionPolicy",
                    "Bypass",
                    "-EncodedCommand",
                    encoded_script,
                ],
                creationflags=subprocess.CREATE_NO_WINDOW,
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
            )

        except Exception:
            logger.debug(
                "Could not notify about completed action",
                exc_info=True,
            )

    
    def create_widgets(self):
        main_frame = ttk.Frame(
            self,
            padding=16,
        )

        main_frame.pack(
            fill="both",
            expand=True,
        )

        title_label = ttk.Label(
            main_frame,
            text="Outlook to DOR Excel Importer",
            font=(
                "Segoe UI",
                17,
                "bold",
            ),
        )

        title_label.pack(
            anchor="w"
        )

        subtitle_label = ttk.Label(
            main_frame,
            text=(
                "Choose an Outlook Excel attachment, "
                "filter its Work Type, and append it "
                "to the open DOR workbook."
            ),
        )

        subtitle_label.pack(
            anchor="w",
            pady=(3, 14),
        )

        settings_frame = ttk.LabelFrame(
            main_frame,
            text="Import Settings",
            padding=12,
        )

        settings_frame.pack(
            fill="x"
        )

        ttk.Label(
            settings_frame,
            text="Queue:",
        ).grid(
            row=0,
            column=0,
            sticky="w",
            pady=5,
        )

        queue_dropdown = ttk.Combobox(
            settings_frame,
            textvariable=self.queue_var,
            values=(
                "Ecrash",
                "CGF",
            ),
            state="readonly",
            width=18,
        )

        queue_dropdown.grid(
            row=0,
            column=1,
            sticky="w",
            padx=(8, 0),
            pady=5,
        )

        queue_dropdown.bind(
            "<<ComboboxSelected>>",
            self.queue_changed,
        )

        ttk.Label(
            settings_frame,
            text="Outlook folder:",
        ).grid(
            row=0,
            column=3,
            sticky="e",
            padx=(16, 0),
            pady=5,
        )

        outlook_folder_entry = ttk.Entry(
            settings_frame,
            textvariable=self.outlook_folder_var,
            width=24,
        )

        outlook_folder_entry.grid(
            row=0,
            column=4,
            sticky="w",
            padx=(8, 0),
            pady=5,
        )

        ttk.Label(
            settings_frame,
            text="Additional filename filter:",
        ).grid(
            row=1,
            column=0,
            sticky="w",
            pady=5,
        )

        file_filter_entry = ttk.Entry(
            settings_frame,
            textvariable=(
                self.file_filter_var
            ),
        )

        file_filter_entry.grid(
            row=1,
            column=1,
            columnspan=2,
            sticky="ew",
            padx=(8, 0),
            pady=5,
        )

        refresh_button = ttk.Button(
            settings_frame,
            text="Refresh Attachments",
            command=self.refresh_attachments,
        )

        refresh_button.grid(
            row=1,
            column=3,
            padx=(10, 0),
            pady=5,
        )

        latest_button = ttk.Button(
            settings_frame,
            text="Select Latest",
            command=self.select_latest,
        )

        latest_button.grid(
            row=1,
            column=4,
            padx=(8, 0),
            pady=5,
        )

        ttk.Label(
            settings_frame,
            text="Open destination workbook:",
        ).grid(
            row=2,
            column=0,
            sticky="w",
            pady=5,
        )

        destination_entry = ttk.Entry(
            settings_frame,
            textvariable=(
                self.destination_var
            ),
        )

        destination_entry.grid(
            row=2,
            column=1,
            columnspan=4,
            sticky="ew",
            padx=(8, 0),
            pady=5,
        )

        ttk.Label(
            settings_frame,
            text="Production Date:",
        ).grid(
            row=3,
            column=0,
            sticky="w",
            pady=5,
        )

        production_date_entry = ttk.Entry(
            settings_frame,
            textvariable=(
                self.production_date_var
            ),
            width=18,
        )

        production_date_entry.grid(
            row=3,
            column=1,
            sticky="w",
            padx=(8, 0),
            pady=5,
        )

        ttk.Label(
            settings_frame,
            text=(
                "YYYY-MM-DD or MM/DD/YYYY"
            ),
        ).grid(
            row=3,
            column=2,
            sticky="w",
            padx=8,
        )

        ttk.Label(
            settings_frame,
            text="Week override:",
        ).grid(
            row=3,
            column=3,
            sticky="e",
            padx=(8, 0),
        )

        week_override_entry = ttk.Entry(
            settings_frame,
            textvariable=(
                self.week_override_var
            ),
            width=16,
        )

        week_override_entry.grid(
            row=3,
            column=4,
            sticky="w",
            padx=(8, 0),
        )

        ttk.Label(
            settings_frame,
            text="Insert at row (optional):",
        ).grid(
            row=4,
            column=0,
            sticky="w",
            pady=5,
        )

        row_override_entry = ttk.Entry(
            settings_frame,
            textvariable=self.row_override_var,
            width=18,
        )

        row_override_entry.grid(
            row=4,
            column=1,
            sticky="w",
            padx=(8, 0),
            pady=5,
        )

        ttk.Label(
            settings_frame,
            text="Rows at this position will move down.",
        ).grid(
            row=4,
            column=2,
            columnspan=3,
            sticky="w",
            padx=8,
        )

        settings_frame.columnconfigure(
            2,
            weight=1,
        )

        attachment_frame = ttk.LabelFrame(
            main_frame,
            text="Available Excel Attachments",
            padding=10,
        )

        attachment_frame.pack(
            fill="both",
            expand=True,
            pady=(14, 10),
        )

        self.attachment_listbox = tk.Listbox(
            attachment_frame,
            font=(
                "Consolas",
                9,
            ),
            activestyle="dotbox",
        )

        attachment_scrollbar = ttk.Scrollbar(
            attachment_frame,
            orient="vertical",
            command=(
                self
                .attachment_listbox
                .yview
            ),
        )

        self.attachment_listbox.configure(
            yscrollcommand=(
                attachment_scrollbar.set
            )
        )

        self.attachment_listbox.pack(
            side="left",
            fill="both",
            expand=True,
        )

        attachment_scrollbar.pack(
            side="right",
            fill="y",
        )

        self.attachment_listbox.bind(
            "<<ListboxSelect>>",
            self.attachment_selection_changed,
        )

        self.import_button = ttk.Button(
            main_frame,
            text="Import Selected Attachment",
            command=self.run_import,
        )
        
        progress_frame = ttk.Frame(
            main_frame,
        )

        progress_frame.pack(
            fill="x",
            pady=(0, 10),
        )

        self.progress = ttk.Progressbar(
            progress_frame,
            mode="determinate",
        )

        self.progress.pack(
            side="left",
            fill="x",
            expand=True,
        )

        self.progress_percent_var = tk.StringVar(
            value="0%"
        )

        progress_percent_label = ttk.Label(
            progress_frame,
            textvariable=self.progress_percent_var,
            width=5,
            anchor="e",
        )

        progress_percent_label.pack(
            side="right",
            padx=(10, 0),
        )

        self.import_button.pack(
            anchor="e",
            pady=(0, 10),
        )

        status_frame = ttk.LabelFrame(
            main_frame,
            text="Status",
            padding=10,
        )

        status_frame.pack(
            fill="x"
        )

        status_label = ttk.Label(
            status_frame,
            textvariable=self.status_var,
            justify="left",
            wraplength=980,
        )

        status_label.pack(
            anchor="w"
        )


    def set_busy(
        self,
        busy: bool,
    ):
        self.action_active = busy

        if busy:
            self.progress.configure(
                mode="determinate"
            )
            self.progress["value"] = 0
            self.progress_target = 0.0
            self.progress_percent_var.set("0%")

            self.configure(
                cursor="watch"
            )

            self.import_button.configure(
                state="disabled"
            )

        else:
            self.configure(
                cursor=""
            )

            self.import_button.configure(
                state="normal"
            )

        self.update_idletasks()

    def update_progress(
        self,
        value: int,
        message: str = "",
    ):
        value = max(
            0,
            min(100, value),
        )

        self.animate_progress(value)

        if message:
            self.status_var.set(message)

        self.update_idletasks()

    def queue_changed(
        self,
        _event=None,
    ):
        selected_queue = (
            self.queue_var.get()
        )

        destination_name = (
            DESTINATION_WORKBOOK_NAMES[
                selected_queue
            ]
        )

        self.destination_var.set(
            destination_name
        )


    def refresh_attachments(self):
        self.set_busy(True)
        self.update_progress(
            10,
            f"Searching Outlook "
            f"{self.outlook_folder_var.get().strip() or 'Inbox'}..."
        )

        threading.Thread(
            target=self._refresh_attachments_worker,
            args=(
                self.file_filter_var.get(),
                self.outlook_folder_var.get(),
            ),
            daemon=True,
        ).start()

    def _refresh_attachments_worker(
        self,
        file_name_filter: str,
        folder_path: str,
    ):
        attachments = None
        error = None
        process_started_at = self.session_timer.start()

        com_initialized = False
        try:
            pythoncom.CoInitialize()
            com_initialized = True

            attachments = list_excel_attachments(
                file_name_filter,
                folder_path,
            )

        except Exception as caught_error:
            error = caught_error

        finally:
            if com_initialized:
                pythoncom.CoUninitialize()

        duration_seconds = self.session_timer.finish(
            "Refresh Attachments",
            process_started_at,
            error is None,
            error,
        )

        self.post_ui_event(
            self._finish_refresh_attachments,
            attachments,
            error,
            duration_seconds,
        )

    def _finish_refresh_attachments(
        self,
        attachments,
        error,
        duration_seconds: float,
    ):
        if error is not None:
            self.notify_action_complete(
                "Attachment refresh failed",
                "Outlook attachments could not be loaded.",
            )
            self.status_var.set(
                f"Refresh Attachments failed after "
                f"{duration_seconds:.2f} seconds."
            )
            self.show_error(
                "Could not search Outlook",
                error,
            )
            self.set_busy(False)
            return

        self.notify_action_complete(
            "Attachment refresh finished",
            "Outlook attachments are ready to review.",
        )

        self.attachments = attachments or []

        self.update_progress(
            65,
            "Loading matching attachments..."
        )

        self.attachment_listbox.delete(
            0,
            tk.END,
        )

        for attachment in self.attachments:
            self.attachment_listbox.insert(
                tk.END,
                attachment.display_name,
            )

        self.update_progress(
            85,
            "Preparing attachment selection..."
        )

        if not self.attachments:
            self.update_progress(
                100,
                "No matching Excel attachments were found. "
                f"Finished in {duration_seconds:.2f} seconds."
            )
            self.set_busy(False)
            return

        self.select_latest()

        summary = (
            f"Found {len(self.attachments)} "
            "matching attachment(s) in "
            f"{duration_seconds:.2f} seconds."
        )

        self.update_progress(
            100,
            summary,
        )
        self.set_busy(False)



    def select_latest(self):
        if not self.attachments:
            self.refresh_attachments()
            return

        self.attachment_listbox.selection_clear(
            0,
            tk.END,
        )

        self.attachment_listbox.selection_set(
            0
        )

        self.attachment_listbox.activate(
            0
        )

        self.attachment_listbox.see(
            0
        )

        self.attachment_selection_changed()


    def get_selected_attachment(
        self,
    ) -> AttachmentInfo | None:
        selected_indexes = (
            self
            .attachment_listbox
            .curselection()
        )

        if not selected_indexes:
            return None

        selected_index = int(
            selected_indexes[0]
        )

        return self.attachments[
            selected_index
        ]


    def attachment_selection_changed(
        self,
        _event=None,
    ):
        attachment = (
            self.get_selected_attachment()
        )

        if attachment is None:
            return

        detected_date = (
            extract_date_from_filename(
                attachment.file_name
            )
        )

        if detected_date is None:
            self.production_date_var.set(
                ""
            )

            self.status_var.set(
                "No date was detected in the filename. "
                "Enter Production Date manually."
            )

            return

        self.production_date_var.set(
            detected_date.strftime(
                "%Y-%m-%d"
            )
        )

        detected_day = (
            detected_date.strftime(
                "%A"
            )
        )

        self.status_var.set(
            f"Detected "
            f"{detected_date.strftime('%m/%d/%Y')} "
            f"({detected_day}) from the filename."
        )


    def run_import(self):
        selected_attachment = (
            self.get_selected_attachment()
        )

        if selected_attachment is None:
            messagebox.showwarning(
                "No attachment",
                "Select an attachment first.",
            )

            return

        destination_name = (
            self.destination_var
            .get()
            .strip()
        )

        if not destination_name:
            messagebox.showwarning(
                "Missing destination",
                "Enter the name of the open "
                "destination workbook.",
            )

            return

        try:
            production_date = parse_date(
                self.production_date_var.get()
            )

        except ValueError as error:
            messagebox.showerror(
                "Invalid date",
                str(error),
            )

            return

        row_override_text = self.row_override_var.get().strip()
        row_override = None

        if row_override_text:
            try:
                row_override = int(row_override_text)
            except ValueError:
                messagebox.showerror(
                    "Invalid row",
                    "Insert at row must be a whole number.",
                )

                return

            if row_override < 2:
                messagebox.showerror(
                    "Invalid row",
                    "Insert at row must be 2 or greater.",
                )

                return

        self.set_busy(True)
        self.update_progress(
            10,
            f'Downloading "'
            f'{selected_attachment.file_name}'
            f'" from Outlook...'
        )

        threading.Thread(
            target=self._import_worker,
            args=(
                selected_attachment,
                destination_name,
                production_date,
                self.queue_var.get(),
                self.week_override_var.get(),
                row_override,
            ),
            daemon=True,
        ).start()

    def _import_worker(
        self,
        selected_attachment: AttachmentInfo,
        destination_name: str,
        production_date: date,
        selected_queue: str,
        week_override: str,
        row_override: int | None,
    ):
        result = None
        error = None
        process_started_at = self.session_timer.start()
        stage_durations = {}

        com_initialized = False
        try:
            pythoncom.CoInitialize()
            com_initialized = True

            download_started_at = self.session_timer.start()

            try:
                downloaded_path = save_attachment(
                    selected_attachment
                )
            except Exception as caught_error:
                stage_durations["Download attachment"] = (
                    self.session_timer.finish(
                        "Download attachment",
                        download_started_at,
                        False,
                        caught_error,
                    )
                )
                raise
            else:
                stage_durations["Download attachment"] = (
                    self.session_timer.finish(
                        "Download attachment",
                        download_started_at,
                        True,
                    )
                )

            self.post_ui_event(
                self.update_progress,
                45,
                "Copying attachment contents...",
            )

            import_started_at = self.session_timer.start()

            try:
                result = import_attachment(
                    source_path=downloaded_path,
                    selected_queue=selected_queue,
                    destination_workbook_name=destination_name,
                    production_date=production_date,
                    week_override=week_override,
                    row_override=row_override,
                )
            except Exception as caught_error:
                stage_durations["Copy into Excel"] = (
                    self.session_timer.finish(
                        "Copy into Excel",
                        import_started_at,
                        False,
                        caught_error,
                    )
                )
                raise
            else:
                stage_durations["Copy into Excel"] = (
                    self.session_timer.finish(
                        "Copy into Excel",
                        import_started_at,
                        True,
                    )
                )

        except Exception as caught_error:
            error = caught_error

        finally:
            if com_initialized:
                pythoncom.CoUninitialize()

        duration_seconds = self.session_timer.finish(
            "Import Selected Attachment",
            process_started_at,
            error is None,
            error,
        )

        self.post_ui_event(
            self._finish_import,
            result,
            error,
            duration_seconds,
            stage_durations,
        )

    def _finish_import(
        self,
        result,
        error,
        duration_seconds: float,
        stage_durations: dict[str, float],
    ):
        if error is not None:
            self.notify_action_complete(
                "Excel import failed",
                "The selected report could not be imported.",
            )
            self.status_var.set(
                f"Import failed after {duration_seconds:.2f} seconds."
            )
            self.show_error(
                "Import failed",
                error,
            )
            self.set_busy(False)
            return

        self.notify_action_complete(
            "Excel import finished",
            "The selected report has finished processing.",
        )

        self.update_progress(
            100,
            "Import completed successfully."
        )

        summary = (
            "Import completed successfully.\n\n"
            f"Workbook: {result['workbook']}\n"
            f"Sheet: {result['sheet']}\n"
            f"Rows imported: {result['rows']}\n"
            f"Columns imported: {result['columns']}\n"
            f"Destination rows: "
            f"{result['first_row']} to "
            f"{result['last_row']}\n"
            f"Week: {result['week']} "
            f"({result['week_source']})\n"
            f"Day: {result['day']}\n"
            f"Production Date: "
            f"{result['production_date']}\n"
            f"Download time: "
            f"{stage_durations['Download attachment']:.2f} seconds\n"
            f"Excel copy time: "
            f"{stage_durations['Copy into Excel']:.2f} seconds\n"
            f"Total import time: {duration_seconds:.2f} seconds"
        )

        self.status_var.set(summary)

        self.set_busy(False)

        messagebox.showinfo(
            "Import Complete",
            summary,
        )

    def close_application(self):
        if self.action_active:
            messagebox.showwarning(
                "Action in progress",
                "Please wait for the current action to finish "
                "before closing the importer.",
            )
            return

        self.closing = True
        summary = self.session_timer.summary()
        logger.info(summary)
        messagebox.showinfo(
            "Session Summary",
            summary,
        )
        self.destroy()


    def show_error(
        self,
        title: str,
        error: Exception,
    ):
        traceback.print_exc()

        error_message = (
            f"{type(error).__name__}: "
            f"{error}"
        )

        self.status_var.set(
            error_message
        )

        messagebox.showerror(
            title,
            error_message,
        )

# ============================================================
# START APPLICATION
# ============================================================

if __name__ == "__main__":
    log_folder = (
        Path(tempfile.gettempdir())
        / "OutlookDORImporter"
    )
    log_folder.mkdir(
        parents=True,
        exist_ok=True,
    )
    logging.basicConfig(
        filename=str(log_folder / "process_timings.log"),
        level=logging.INFO,
        format="%(asctime)s %(levelname)s %(message)s",
    )

    application = OutlookDORImporterApp()
    application.mainloop()

