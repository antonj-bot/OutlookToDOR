"""Outlook folder discovery, filtering, and path resolution."""

import re
import threading

from backend.config import OL_FOLDER_INBOX
from backend.exceptions import check_cancelled
from backend.logging_setup import logger

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
