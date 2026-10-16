"""Small data models used by the UI and backend services."""

import threading
import time
from dataclasses import dataclass
from datetime import datetime

from backend.logging_setup import logger

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
