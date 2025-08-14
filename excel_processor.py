"""Excel processing utilities. Do not modify."""

from logger import get_logger

logger = get_logger(__name__)


def process_excel(file_path):
    """Process an Excel file at the given path."""
    logger.info("Processing Excel file: %s", file_path)
    return []
