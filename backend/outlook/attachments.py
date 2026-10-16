"""Find and save qualifying Excel attachments in Outlook."""

import tempfile
import threading
from datetime import datetime
from pathlib import Path

from backend.config import OUTLOOK_MESSAGE_LIMIT, REQUIRED_ATTACHMENT_KEYWORD
from backend.exceptions import ProcessCancelled, check_cancelled
from backend.logging_setup import logger
from backend.models import AttachmentInfo
from backend.outlook.connection import get_outlook_namespace
from backend.outlook.folders import get_outlook_folder
from backend.excel.source_reader import is_excel_file

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
