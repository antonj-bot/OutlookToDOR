"""Helpers for locating open Excel workbooks and worksheets."""

from backend.excel.source_reader import workbook_stem

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
