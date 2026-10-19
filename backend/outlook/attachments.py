"""Find and save qualifying Excel attachments in Outlook."""

import tempfile
import threading
import gc
import time

from datetime import datetime
from pathlib import Path

from backend.exceptions import ProcessCancelled, check_cancelled
from backend.logging_setup import logger
from backend.models import AttachmentInfo
from backend.outlook.connection import get_outlook_namespace
from backend.outlook.folders import get_outlook_folder
from backend.excel.source_reader import is_excel_file

def list_excel_attachments(
    required_keyword: str = "",
    file_name_filter: str = "",
    folder_path: str = "Inbox",
    cancel_event: threading.Event | None = None,
) -> list[AttachmentInfo]:
    """
    Searches Outlook messages for supported attachments whose filenames
    contain the required keyword and optional additional filter.

    An empty required keyword includes all supported attachments.
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

    normalized_required_keyword = (
        required_keyword
        .strip()
        .casefold()
    )

    requested_text = (
        file_name_filter
        .strip()
        .casefold()
    )

    results: list[AttachmentInfo] = []

    for message in messages:
        check_cancelled(cancel_event)
       
        try:
            attachment_count = int(
                message.Attachments.Count
            )
        except Exception:
            logger.exception(
                "Failed to read an Outlook message's attachments.")
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
                    attachment.FileName or ""
                ).strip()

                attachment_type = int(
                    attachment.Type
                )

            except Exception:
                logger.exception(
                    "Could not read attachment %s from message %r.",
                    attachment_index,
                    subject,
                )
                continue

            extension = (
                Path(file_name)
                .suffix
                .casefold()
            )

            file_name_text = file_name.casefold()

            is_regular_excel_file = is_excel_file(
                file_name
            )

            # Outlook olEmbeddeditem
            is_embedded_outlook_item = (
                attachment_type == 5
            )

            logger.info(
                "Attachment candidate: filename=%r, "
                "extension=%r, type=%r, "
                "is_excel=%s, is_embedded=%s",
                file_name,
                extension,
                attachment_type,
                is_regular_excel_file,
                is_embedded_outlook_item,
            )

            # Keep actual Excel files and embedded Outlook messages.
            if (
                not is_regular_excel_file
                and not is_embedded_outlook_item
            ):
                logger.info(
                    "Skipping unsupported attachment: %r",
                    file_name,
                )
                continue

            if (
                normalized_required_keyword
                and normalized_required_keyword not in file_name_text
            ):
                logger.info(
                    "Skipping %r because required keyword %r "
                    "is missing.",
                    file_name,
                    normalized_required_keyword,
                )
                continue

            if (
                requested_text
                and requested_text not in file_name_text
            ):
                logger.info(
                    "Skipping %r because additional filter %r "
                    "is missing.",
                    file_name,
                    requested_text,
                )
                continue

            logger.info(
                "Adding attachment to UI: filename=%r, type=%r",
                file_name,
                attachment_type,
            )

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
    Save a normal Excel attachment or extract an Excel workbook from an
    attached Outlook message.

    Returns the path to the actual Excel workbook.
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

    try:
        attachment_type = int(attachment.Type)
    except Exception:
        attachment_type = None

    # Outlook olEmbeddeditem
    is_embedded_outlook_item = attachment_type == 5

    actual_file_name = str(
        attachment.FileName or attachment_info.file_name
    ).strip()

    if is_embedded_outlook_item:
        return _extract_excel_from_embedded_message(
            attachment=attachment,
            embedded_name=actual_file_name,
            download_folder=download_folder,
            namespace=namespace,
            cancel_event=cancel_event,
        )

    output_path = _unique_output_path(
        download_folder,
        actual_file_name,
    )

    attachment.SaveAsFile(
        str(output_path)
    )

    try:
        check_cancelled(cancel_event)
    except ProcessCancelled:
        output_path.unlink(missing_ok=True)
        raise

    if not output_path.is_file():
        raise RuntimeError(
            "Outlook did not save the selected attachment."
        )

    if not is_excel_file(output_path.name):
        output_path.unlink(missing_ok=True)

        raise ValueError(
            f'"{actual_file_name}" is not a supported Excel attachment.'
        )

    return output_path

def _unique_output_path(
    download_folder: Path,
    file_name: str,
) -> Path:
    """
    Return a non-existing path inside the download folder.
    """
    safe_file_name = Path(file_name).name
    output_path = download_folder / safe_file_name

    counter = 1

    while output_path.exists():
        original_path = Path(safe_file_name)

        output_path = download_folder / (
            f"{original_path.stem}_"
            f"{counter}"
            f"{original_path.suffix}"
        )

        counter += 1

    return output_path

def _delete_temporary_file(
    file_path: Path,
    attempts: int = 5,
    delay_seconds: float = 0.2,
) -> None:
    """
    Delete a temporary file, retrying if Outlook still has it locked.

    Cleanup failure is logged but does not fail a successful import.
    """
    for attempt in range(1, attempts + 1):
        try:
            file_path.unlink(missing_ok=True)
            return

        except PermissionError:
            if attempt == attempts:
                logger.warning(
                    "Could not delete locked temporary file: %s",
                    file_path,
                    exc_info=True,
                )
                return

            time.sleep(delay_seconds)


def _extract_excel_from_embedded_message(
    attachment,
    embedded_name: str,
    download_folder: Path,
    namespace,
    cancel_event: threading.Event | None,
) -> Path:
    """
    Save an Outlook embedded item as MSG, open it, and extract the first
    supported Excel attachment.
    """
    check_cancelled(cancel_event)

    safe_stem = Path(embedded_name).stem.strip()

    if not safe_stem:
        safe_stem = "embedded_message"

    msg_path = _unique_output_path(
        download_folder,
        f"{safe_stem}.msg",
    )

    attachment.SaveAsFile(
        str(msg_path)
    )

    if not msg_path.is_file():
        raise RuntimeError(
            "Outlook did not save the embedded message."
        )

    embedded_message = None

    try:
        check_cancelled(cancel_event)

        embedded_message = namespace.OpenSharedItem(
            str(msg_path)
        )

        inner_attachment_count = int(
            embedded_message.Attachments.Count
        )

        logger.info(
            "Embedded Outlook item %r contains %s attachment(s).",
            embedded_name,
            inner_attachment_count,
        )

        for inner_index in range(
            1,
            inner_attachment_count + 1,
        ):
            check_cancelled(cancel_event)

            inner_attachment = (
                embedded_message
                .Attachments
                .Item(inner_index)
            )

            inner_file_name = str(
                inner_attachment.FileName or ""
            ).strip()
            
            logger.info(
                "Checking inner attachment %s: %r",
                inner_index,
                inner_file_name,
            )

            inner_extension = (
                Path(inner_file_name)
                .suffix
                .casefold()
            )

            # Only accept actual Excel workbooks inside the embedded email.
            if inner_extension not in {
                ".xlsx",
                ".xlsm",
                ".xlsb",
                ".xls",
            }:
                logger.info(
                    "Skipping non-Excel inner attachment: "
                    "filename=%r, extension=%r",
                    inner_file_name,
                    inner_extension,
                )
                continue

            excel_path = _unique_output_path(
                download_folder,
                inner_file_name,
            )

            logger.info(
                "Saving inner Excel attachment %r to %s",
                inner_file_name,
                excel_path,
            )

            inner_attachment.SaveAsFile(
                str(excel_path)
            )

            check_cancelled(cancel_event)

            if not excel_path.is_file():
                raise RuntimeError(
                    "Outlook did not save the inner Excel "
                    f'attachment "{inner_file_name}".'
                )

            logger.info(
                "Successfully extracted Excel attachment %r "
                "from embedded Outlook item %r.",
                inner_file_name,
                embedded_name,
            )

            return excel_path

        raise ValueError(
            "No supported Excel attachment was found inside "
            f'the attached Outlook item "{embedded_name}".'
        )

    except ProcessCancelled:
        logger.info(
            "Embedded-message extraction was cancelled: %r",
            embedded_name,
        )
        raise

    except Exception:
        logger.exception(
            "Failed to extract an Excel attachment from "
            "embedded Outlook item %r.",
            embedded_name,
        )
        raise

    finally:
        try:
            inner_attachment = None
        except UnboundLocalError:
            pass

        if embedded_message is not None:
            try:
                embedded_message.Close(1)
            except Exception:
                logger.debug(
                    "Could not close embedded Outlook item %r.",
                    embedded_name,
                    exc_info=True,
                )
            finally:
                embedded_message = None

        gc.collect()

        _delete_temporary_file(msg_path)          

