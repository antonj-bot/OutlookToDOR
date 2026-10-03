from __future__ import annotations

import difflib
import logging
import re
import tempfile
import threading
import time

from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path

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
XL_NONE = -4142
IMPORTED_DATA_HIGHLIGHT_COLOR = 65535
DOR_PASS_1_COLOR = 16247773
DOR_PASS_2_COLOR = 14348258
DOR_SOURCE_COLORS = {
    "completed report pass 1": DOR_PASS_1_COLOR,
    "completed report pass 2": DOR_PASS_2_COLOR,
    "total elapsed time": 8696052,
    "average report/hr": 14067636,
    "completed reports": 14539150,
    "late %": 13421812,
    "ontime%": 11065270,
    "auto vs manual manual": 10086143,
    "manual reports average report/hr": 15323865,
    "auto vs manual auto extraction": 14934224,
    "auto extraction average report/hr": 13493756,
}

DOR_COLUMN_SOURCE_COLORS = {
    3: DOR_SOURCE_COLORS["completed report pass 1"],
    4: DOR_SOURCE_COLORS["completed report pass 2"],
    5: DOR_SOURCE_COLORS["total elapsed time"],
    6: DOR_SOURCE_COLORS["average report/hr"],
    7: DOR_SOURCE_COLORS["completed reports"],
    8: DOR_SOURCE_COLORS["late %"],
    9: DOR_SOURCE_COLORS["ontime%"],
    10: DOR_SOURCE_COLORS["auto vs manual manual"],
    12: DOR_SOURCE_COLORS["manual reports average report/hr"],
    13: DOR_SOURCE_COLORS["auto vs manual auto extraction"],
    15: DOR_SOURCE_COLORS["auto extraction average report/hr"],
}

# ============================================================
# PROCESS TIMING
# ============================================================

logger = logging.getLogger("outlook_dor_importer")


class ProcessCancelled(Exception):
    """Raised when a user requests cancellation at a safe checkpoint."""


def check_cancelled(cancel_event: threading.Event | None) -> None:
    if cancel_event is not None and cancel_event.is_set():
        raise ProcessCancelled("Operation cancelled by the user.")


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


def list_outlook_folder_paths(
    namespace,
    cancel_event: threading.Event | None = None,
) -> list[tuple[str, str]]:
    """Returns folder names paired with their mailbox-rooted paths."""
    check_cancelled(cancel_event)
    folders = [("Inbox", "Inbox")]

    def collect_folders(folder, parent_path=""):
        check_cancelled(cancel_event)
        folder_name = str(folder.Name).strip()
        current_path = (
            f"{parent_path}\\{folder_name}"
            if parent_path
            else folder_name
        )
        folders.append((folder_name, current_path))

        for child_folder in folder.Folders:
            check_cancelled(cancel_event)
            collect_folders(child_folder, current_path)

    for root_folder in namespace.Folders:
        check_cancelled(cancel_event)
        collect_folders(root_folder)

    sorted_folders = sorted(
        folders,
        key=lambda item: (item[0].casefold(), item[1].casefold()),
    )
    check_cancelled(cancel_event)
    return sorted_folders


def filter_outlook_folder_choices(
    folders: list[tuple[str, str]],
    query: str,
) -> list[tuple[str, str]]:
    normalized_query = query.strip().casefold()

    if not normalized_query:
        return folders

    return [
        folder
        for folder in folders
        if normalized_query in folder[0].casefold()
        or normalized_query in folder[1].casefold()
    ]


def get_outlook_folder(
    namespace,
    folder_path: str = "Inbox",
    cancel_event: threading.Event | None = None,
):
    """
    Resolves Inbox paths, mailbox-rooted paths, and unique bare folder names.
    """
    requested_path = folder_path.strip() or "Inbox"
    check_cancelled(cancel_event)
    path_parts = [
        part.strip()
        for part in re.split(r"[\\/]+", requested_path)
        if part.strip()
    ]

    if not path_parts:
        path_parts = ["Inbox"]

    logger.info("Resolving Outlook folder path: %s", requested_path)

    if path_parts[0].casefold() == "inbox":
        current_folder = namespace.GetDefaultFolder(
            OL_FOLDER_INBOX
        )
        path_parts = path_parts[1:]
    elif len(path_parts) == 1:
        all_folders = []

        def collect_folders(folder, parent_path=""):
            check_cancelled(cancel_event)
            folder_name = str(folder.Name).strip()
            current_path = (
                f"{parent_path}\\{folder_name}"
                if parent_path
                else folder_name
            )
            all_folders.append(
                (folder_name, current_path, folder)
            )

            for child_folder in folder.Folders:
                check_cancelled(cancel_event)
                collect_folders(child_folder, current_path)

        for root_folder in namespace.Folders:
            check_cancelled(cancel_event)
            collect_folders(root_folder)

        requested_name = path_parts[0].casefold()
        exact_matches = [
            match
            for match in all_folders
            if match[0].casefold() == requested_name
        ]
        matches = exact_matches or [
            match
            for match in all_folders
            if requested_name in match[0].casefold()
        ]

        if len(matches) == 1:
            logger.info(
                'Resolved Outlook folder "%s" to "%s"',
                requested_path,
                matches[0][1],
            )
            return matches[0][2]

        if matches:
            matching_paths = "\n".join(
                match[1]
                for match in matches
            )
            raise RuntimeError(
                f'Outlook folder name "{path_parts[0]}" matches '
                "multiple folders. Enter a full folder path:\n\n"
                f"{matching_paths}"
            )

        available_paths = "\n".join(
            match[1]
            for match in all_folders
        ) or "None"
        raise RuntimeError(
            f'No Outlook folder matched "{path_parts[0]}".\n\n'
            f"Available Outlook folders:\n{available_paths}"
        )
    else:
        current_folder = None
        mailbox_folders = list(namespace.Folders)

        for folder in mailbox_folders:
            check_cancelled(cancel_event)
            if str(folder.Name).strip().casefold() == path_parts[0].casefold():
                current_folder = folder
                break

        if current_folder is None:
            available_mailboxes = "\n".join(
                str(folder.Name)
                for folder in mailbox_folders
            ) or "None"
            raise RuntimeError(
                f'Outlook mailbox or folder root "{path_parts[0]}" '
                "was not found.\n\n"
                f"Available mailbox roots:\n{available_mailboxes}"
            )

        path_parts = path_parts[1:]

    for requested_name in path_parts:
        check_cancelled(cancel_event)
        next_folder = None

        for folder in current_folder.Folders:
            check_cancelled(cancel_event)
            if (
                str(folder.Name).strip().casefold()
                == requested_name.casefold()
            ):
                next_folder = folder
                break

        if next_folder is None:
            available_folders = "\n".join(
                str(folder.Name)
                for folder in current_folder.Folders
            ) or "None"
            raise RuntimeError(
                f'Outlook folder "{requested_name}" was not found '
                f'under "{current_folder.Name}".\n\n'
                f"Available folders under this location:\n"
                f"{available_folders}"
            )

        current_folder = next_folder

    logger.info(
        'Resolved Outlook folder "%s" to "%s"',
        requested_path,
        current_folder.Name,
    )
    return current_folder



def list_excel_attachments(
    file_name_filter: str = "",
    folder_path: str = "Inbox",
    cancel_event: threading.Event | None = None,
) -> list[AttachmentInfo]:
    """
    Searches recent Outlook messages in the selected folder for matching
    Excel attachments.

    Every result must contain REQUIRED_ATTACHMENT_KEYWORD in its filename.
    The optional file_name_filter narrows those results further.

    Results are returned newest first.
    """
    check_cancelled(cancel_event)
    namespace = get_outlook_namespace()

    outlook_folder = get_outlook_folder(
        namespace,
        folder_path,
        cancel_event,
    )

    messages = outlook_folder.Items

    # Newest messages first.
    messages.Sort(
        "[ReceivedTime]",
        True,
    )
    check_cancelled(cancel_event)

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
        check_cancelled(cancel_event)
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
            check_cancelled(cancel_event)
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
    cancel_event: threading.Event | None = None,
) -> Path:
    """
    Downloads the selected Outlook attachment to a temporary folder.

    An existing file is not overwritten. A number is added instead.
    """
    check_cancelled(cancel_event)
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
    try:
        check_cancelled(cancel_event)
    except ProcessCancelled:
        output_path.unlink(missing_ok=True)
        raise

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


def append_dor_import_record(
    dor_sheet,
    production_date: date,
    summary_values: tuple,
    highlight_imported_data: bool = False,
) -> int:
    """Writes the imported report's summary values to the DOR sheet."""
    dor_row = 5
    row_inserted = False

    if has_value(dor_sheet.Cells(5, 1).Value):
        dor_row = 6
        insert_destination_rows(
            dor_sheet,
            dor_row,
            1,
        )
        row_inserted = True

    report_date = datetime.combine(
        production_date,
        datetime.min.time(),
    )

    try:
        write_matrix(
            dor_sheet,
            dor_row,
            1,
            [[
                report_date,
                None,
                *summary_values[:8],
                None,
                summary_values[8],
                summary_values[9],
                None,
                summary_values[10],
            ]],
        )
        dor_sheet.Cells(
            dor_row,
            1,
        ).NumberFormat = "dddd, mmmm d, yyyy"
        dor_sheet.Cells(
            dor_row,
            11,
        ).FormulaR1C1 = "=IFERROR(RC[-1]/RC[-7],0)"
        dor_sheet.Cells(
            dor_row,
            14,
        ).FormulaR1C1 = "=IFERROR(RC[-1]/RC[-10],0)"
        if highlight_imported_data:
            dor_sheet.Range(
                dor_sheet.Cells(dor_row, 1),
                dor_sheet.Cells(dor_row, 15),
            ).Interior.Color = IMPORTED_DATA_HIGHLIGHT_COLOR
        for column_number, source_color in DOR_COLUMN_SOURCE_COLORS.items():
            dor_sheet.Cells(
                dor_row,
                column_number,
            ).Interior.Color = source_color
    except Exception:
        if row_inserted:
            dor_sheet.Rows("6:6").Delete(
                Shift=XL_UP,
            )
        else:
            dor_sheet.Range(
                dor_sheet.Cells(dor_row, 1),
                dor_sheet.Cells(dor_row, 15),
            ).ClearContents()
        raise

    return dor_row


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
                worksheet.Cells(
                    first_row,
                    first_column,
                ),
                worksheet.Cells(
                    last_row,
                    last_column,
                ),
            ).Interior.Pattern = XL_NONE
            changed_workbooks[workbook_key] = workbook

        check_cancelled(cancel_event)
    except ProcessCancelled:
        for workbook in changed_workbooks.values():
            workbook.Save()
        raise

    for workbook in changed_workbooks.values():
        workbook.Save()


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


def undo_imported_data(
    import_result: dict,
    raw_row_segments: list[tuple[int, int]],
    dor_row: int,
    ranges_to_remove: list[dict],
    remaining_ranges: list[dict],
    rehighlight_remaining: bool,
    cancel_event: threading.Event | None = None,
) -> None:
    """Removes one completed import from raw data and DOR, then refreshes pivots."""
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
    dor_sheet = find_sheet(workbook, "DOR")

    for imported_range in ranges_to_remove:
        if imported_range.get("kind") == "pivot" and not imported_range.get(
            "is_highlighted",
            False,
        ):
            continue
        pivot_sheet = find_sheet(
            workbook,
            imported_range["sheet"],
        )
        first_row, last_row, first_column, last_column = resolve_imported_highlight_range(
            pivot_sheet,
            imported_range,
            cancel_event,
        )
        pivot_sheet.Range(
            pivot_sheet.Cells(first_row, first_column),
            pivot_sheet.Cells(last_row, last_column),
        ).Interior.Pattern = XL_NONE

    for first_row, last_row in sorted(raw_row_segments, reverse=True):
        raw_sheet.Rows(
            f"{first_row}:{last_row}"
        ).Delete(
            Shift=XL_UP,
        )

    dor_sheet.Rows(
        f"{dor_row}:{dor_row}"
    ).Delete(
        Shift=XL_UP,
    )

    try:
        workbook.RefreshAll()
        excel.CalculateUntilAsyncQueriesDone()
    except Exception as error:
        logger.warning(
            "Could not refresh workbook data after undo: %s",
            error,
        )

    if rehighlight_remaining and remaining_ranges:
        highlight_imported_data_ranges(
            remaining_ranges,
            cancel_event,
        )

    workbook.Save()


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
        "auto extraction",
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
    manual_reports_section = sections.get("manual reports")
    auto_extraction_section = sections.get("auto extraction")
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
    manual_reports_date_header = date_headers["manual reports"]
    auto_extraction_date_header = date_headers["auto extraction"]
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
        "manual reports",
        manual_reports_date_header,
    )
    auto_extraction_data_row = find_data_row(
        "auto extraction",
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


# ============================================================
# MAIN IMPORT FUNCTION
# ============================================================
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
    summary_sheet = find_sheet(
        destination_workbook,
        "Summary by date (All states)",
    )

    source_workbook = None

    try:
        check_cancelled(cancel_event)
        source_workbook = excel.Workbooks.Open(
            str(source_path),
            UpdateLinks=0,
            ReadOnly=True,
            IgnoreReadOnlyRecommended=True,
        )

        source_sheet = find_source_data_sheet(
            source_workbook,
            cancel_event,
        )
        
        result = read_filtered_source_rows(
            source_sheet,
            selected_queue,
            cancel_event,
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
        for pivot_source_range in pivot_source_ranges:
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

        dor_row = append_dor_import_record(
            dor_sheet,
            production_date,
            summary_values,
            highlight_imported_data,
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

        highlight_ranges.extend(pivot_source_ranges)

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
        }

    finally:
        if source_workbook is not None:
            source_workbook.Close(
                SaveChanges=False
            )

# ============================================================
# START APPLICATION
# ============================================================

if __name__ == "__main__":
    log_file = Path(__file__).with_name(
        "process_timings.log"
    )
    logging.basicConfig(
        filename=str(log_file),
        level=logging.INFO,
        format="%(asctime)s %(levelname)s %(message)s",
    )

    import sys

    sys.modules.setdefault("outlook_dor_importer", sys.modules[__name__])
    from outlook_dor_importer_ui import OutlookDORImporterApp

    application = OutlookDORImporterApp()
    application.mainloop()
