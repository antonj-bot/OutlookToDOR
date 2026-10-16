"""Logging configuration shared by the application and backend."""

import logging
from pathlib import Path

logger = logging.getLogger("outlook_dor_importer")

def configure_logging(log_path: Path) -> None:
    logging.basicConfig(
        filename=str(log_path),
        level=logging.INFO,
        format="%(asctime)s %(levelname)s %(message)s",
    )

logger = logging.getLogger("outlook_dor_importer")
