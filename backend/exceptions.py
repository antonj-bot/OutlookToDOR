"""Cancellation support shared by background operations."""

import threading

class ProcessCancelled(Exception):
    """Raised when a user requests cancellation at a safe checkpoint."""
def check_cancelled(cancel_event: threading.Event | None) -> None:
    if cancel_event is not None and cancel_event.is_set():
        raise ProcessCancelled("Operation cancelled by the user.")
