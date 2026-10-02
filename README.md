
# Outlook DOR Importer

A Windows desktop utility that finds qualifying Excel attachments in Outlook and appends filtered rows to an open DOR workbook in Excel.

## What It Does

- Searches an Outlook folder for recent Excel attachments.
- Only shows attachments whose filename contains `Individual Keyed Pass per User`.
- Supports an additional optional filename filter.
- Filters source rows by queue:
	- `Ecrash`: Work Type contains `ecrash` or `e crash`.
	- `CGF`: Work Type contains `cgf`.
- Excludes the source Username column from copied data.
- Adds Week, Day, and Production Date to imported rows.
- Reuses an existing Week value for the same Production Date when available.
- Reports elapsed time for refresh, download, and Excel import operations.
- Displays a session summary when the application closes.

## Requirements

- Windows
- Python 3.10 or later
- Classic desktop Outlook configured with the target mailbox
- Microsoft Excel desktop with the destination workbook already open
- Python packages:
	- `pywin32`

Install the Python dependency from PowerShell:

```powershell
python -m pip install pywin32
```

## Run

From the project folder:

```powershell
python outlook_dor_importer.py
```

## Typical Workflow

1. Open the destination DOR workbook in Excel.
2. Start the importer.
3. Select `Ecrash` or `CGF`.
4. Select the Outlook folder from the dropdown. Click **Refresh Folders** to
	reload the mailbox folder list; the default selection is `Inbox`.
5. Optionally enter an additional filename filter.
6. Select **Refresh Attachments**.
7. Select an attachment and confirm the Production Date.
8. Enter a Week override only when needed.
9. Leave **Insert at row (optional)** blank to append, or enter a row
	number to insert the imported rows at that position. Existing rows at
	that position and below move down.
10. Select **Import Selected Attachment**.

When the window is minimized or another application is active, the importer
flashes its taskbar button and plays an attention sound when refreshing or
importing finishes.

## Outlook Folder Override

The Outlook folder dropdown shows actual folder names. Each choice retains its
full mailbox path for searching, so folders with the same name appear as
separate choices. The selected folder's full path appears in the status line.
The field is also editable: type a folder name such as `DOR` or a full path
such as `Mailbox Name\Reports` to display matches as you type. If a typed name
matches multiple folders, use the dropdown to select the intended folder.
Click **Refresh Folders** after mailbox folders change. The application
searches up to the most recent 500 messages in the selected folder.

## Expected Source Data

The source workbook can contain multiple worksheets. The importer looks for a worksheet with both:

- A Work Type column, including names such as `Work Type`, `Worktype`, or `Work Category`.
- A Username column, including names such as `Username`, `User Name`, `Agent`, or `Employee Name`.

The destination sheet must contain:

- `USER`
- `Week`
- `Day`
- `Production Date`

After a successful import, the importer also writes a summary record to the
open workbook's `DOR` sheet. It uses the selected Production Date in column A,
leaves column B blank, and copies `H10` and `I10` from `Summary by date (All
states)` into columns C and D. It writes to row 5 when `DOR!A5` is empty;
otherwise, it inserts a row at 6.

Destination workbook and sheet names are configured near the top of `outlook_dor_importer.py` in `DESTINATION_WORKBOOK_NAMES` and `DESTINATION_SHEETS`.

## Timing Logs

Process timings are written to:

```text
process_timings.log
```

The log file is created next to `outlook_dor_importer.py`. It includes folder
lookup details and process timings. The application also shows the timing
summary when the window is closed.

## Troubleshooting

### No attachments are found

- Confirm the message is in the selected Outlook folder.
- Confirm the attachment is an Excel file with an allowed extension: `.xlsx`, `.xlsm`, `.xlsb`, or `.xls`.
- Confirm the filename contains `Individual Keyed Pass per User`.
- Clear the optional additional filename filter.

### The destination workbook is not found

Open the workbook in Excel and ensure its name matches the configured workbook name, with or without the file extension.

### The import reports missing headers

Check the source and destination header names. The source header search allows several common aliases, but the destination metadata headers must be present.

### Outlook or Excel automation fails

Ensure Outlook and Excel are installed as desktop applications, close unexpected dialog boxes, and run the importer with the same Windows user account that owns the Office session.

## Project Files

| File | Purpose |
| --- | --- |
| `outlook_dor_importer.py` | Outlook/Excel import logic and application launcher |
| `outlook_dor_importer_ui.py` | Tkinter user interface |
| `LICENSE` | MIT license |

## License

This project is available under the MIT License. See [LICENSE](LICENSE).
