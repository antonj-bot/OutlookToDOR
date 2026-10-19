"""Compatibility exports and application launcher.

Implementation lives in the backend package; these imports preserve the
original module API for existing scripts and integrations.
"""

from pathlib import Path

from backend.config import (
    ALLOWED_EXCEL_EXTENSIONS,
    DESTINATION_SHEETS,
    DESTINATION_WORKBOOK_NAMES,
    DOR_COLUMN_SOURCE_COLORS,
    DOR_COMPARISON_EXCLUDING_TX_SECTION_TITLE,
    DOR_COMPARISON_SECTION_TITLE,
    DOR_HEADER_ROWS,
    DOR_PASS_1_COLOR,
    DOR_PASS_2_COLOR,
    DOR_SOURCE_COLORS,
    IMPORTED_DATA_HIGHLIGHT_COLOR,
    OL_FOLDER_INBOX,
    OUTLOOK_MESSAGE_LIMIT,
    REQUIRED_ATTACHMENT_KEYWORD,
    WORK_TYPE_MATCH_VALUES,
    XL_DOWN,
    XL_NONE,
    XL_TO_LEFT,
    XL_UP,
)
from backend.excel.comparison import (
    append_dor_comparison_record,
    get_dor_comparison_mapping,
    get_header_path_map,
)
from backend.excel.connection import (
    find_open_workbook,
    find_sheet,
)
from backend.excel.destination import (
    append_dor_import_record,
    fill_imported_row_formulas,
    fill_metadata,
    get_destination_header_map,
    insert_destination_rows,
    next_empty_row,
    write_matrix,
)
from backend.excel.highlighting import (
    clear_imported_data_highlight,
    highlight_imported_data_ranges,
    refresh_destination_workbook_data,
    resolve_imported_highlight_range,
)
from backend.excel.ranges import (
    range_to_rows,
)
from backend.excel.source_reader import (
    extract_date_from_filename,
    fallback_week,
    find_matching_header,
    find_source_header_row,
    has_value,
    is_excel_file,
    normalize,
    parse_date,
    read_filtered_source_rows,
    to_python_date,
    work_type_matches,
    workbook_stem,
)
from backend.excel.summaries import (
    find_existing_week,
    get_dor_summary_values,
)
from backend.excel.remove import (
    remove_imported_data,
)
from backend.exceptions import (
    ProcessCancelled,
    check_cancelled,
)
from backend.logging_setup import (
    logger,
)
from backend.models import (
    AttachmentInfo,
    ProcessTiming,
    SessionTimer,
)
from backend.outlook.attachments import (
    list_excel_attachments,
    save_attachment,
)
from backend.outlook.connection import (
    get_outlook_namespace,
)
from backend.outlook.folders import (
    filter_outlook_folder_choices,
    get_outlook_folder,
    list_outlook_folder_paths,
)
from backend.services.import_service import (
    find_source_data_sheet,
    import_attachment,
)


def main() -> None:
    from backend.logging_setup import configure_logging
    from outlook_dor_importer_ui import OutlookDORImporterApp

    configure_logging(Path(__file__).with_name("process_timings.log"))
    OutlookDORImporterApp().mainloop()


if __name__ == "__main__":
    main()
