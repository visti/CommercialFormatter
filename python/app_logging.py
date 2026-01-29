"""Logging configuration for commercial formatter."""

from __future__ import annotations

import logging
import sys
from datetime import datetime
from pathlib import Path
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from stations import Station

from settings import get_settings

# Module-level logger
_logger: logging.Logger | None = None
_log_file_path: Path | None = None


def setup_logging(station: "Station", output_dir: Path | None = None) -> logging.Logger:
    """Set up logging for a processing session.

    Args:
        station: The station being processed
        output_dir: Directory for log files. If None, uses current directory.

    Returns:
        Configured logger instance
    """
    global _logger, _log_file_path

    settings = get_settings()

    # Create logger
    logger = logging.getLogger("komm_fmt")
    logger.setLevel(logging.DEBUG)  # Capture all, filter at handler level

    # Remove existing handlers
    logger.handlers.clear()

    # Console handler (only warnings and errors)
    console_handler = logging.StreamHandler(sys.stderr)
    console_handler.setLevel(logging.WARNING)
    console_format = logging.Formatter("%(levelname)s: %(message)s")
    console_handler.setFormatter(console_format)
    logger.addHandler(console_handler)

    # File handler (if enabled)
    if settings.logging.enabled:
        if output_dir is None:
            output_dir = Path.cwd()

        # Format log filename
        date_str = datetime.now().strftime("%Y-%m-%d")
        log_filename = settings.logging.filename.format(
            date=date_str, station=station.name.lower()
        )

        # Handle relative paths
        log_path = Path(log_filename)
        if not log_path.is_absolute():
            log_path = output_dir / log_path

        # Ensure log directory exists
        log_path.parent.mkdir(parents=True, exist_ok=True)

        _log_file_path = log_path

        # Set up file handler
        file_handler = logging.FileHandler(log_path, encoding="utf-8")
        level = getattr(logging, settings.logging.level, logging.INFO)
        file_handler.setLevel(level)

        file_format = logging.Formatter(
            "%(asctime)s | %(levelname)-8s | %(message)s",
            datefmt="%Y-%m-%d %H:%M:%S",
        )
        file_handler.setFormatter(file_format)
        logger.addHandler(file_handler)

        # Log session start
        logger.info("=" * 60)
        logger.info(f"Processing session started for station: {station.name}")
        logger.info("=" * 60)

    _logger = logger
    return logger


def get_logger() -> logging.Logger:
    """Get the application logger.

    Returns:
        The configured logger, or a default logger if not set up.
    """
    global _logger
    if _logger is None:
        # Return a basic logger if not set up
        _logger = logging.getLogger("komm_fmt")
        if not _logger.handlers:
            _logger.addHandler(logging.NullHandler())
    return _logger


def get_log_file_path() -> Path | None:
    """Get the current log file path.

    Returns:
        Path to the log file, or None if logging to file is disabled.
    """
    return _log_file_path


def log_file_read(file_path: Path, encoding: str, line_count: int) -> None:
    """Log a file read operation."""
    logger = get_logger()
    logger.info(f"Read file: {file_path.name} ({encoding}, {line_count} lines)")


def log_rejection(line_num: int, reason: str, content_preview: str) -> None:
    """Log a line rejection."""
    logger = get_logger()
    preview = content_preview[:50] + "..." if len(content_preview) > 50 else content_preview
    logger.debug(f"Rejected line {line_num}: {reason} - {preview}")


def log_stopword_match(line_num: int, stopword: str) -> None:
    """Log a stopword match."""
    logger = get_logger()
    logger.debug(f"Line {line_num} matched stopword: {stopword}")


def log_overflow_fix(line_num: int, original: str, fixed: str) -> None:
    """Log a playing time overflow fix."""
    logger = get_logger()
    logger.info(f"Line {line_num}: Fixed overflow time {original} -> {fixed}")


def log_user_choice(choice_type: str, description: str, action: str) -> None:
    """Log a user choice."""
    logger = get_logger()
    logger.info(f"User choice ({choice_type}): {description} -> {action}")


def log_duplicate_found(
    line_num: int, title: str, artist: str, date: str, first_line: int
) -> None:
    """Log a duplicate track found."""
    logger = get_logger()
    logger.info(
        f"Duplicate at line {line_num}: '{title}' by '{artist}' on {date} "
        f"(first seen at line {first_line})"
    )


def log_processing_complete(
    files: int, lines_processed: int, lines_rejected: int, duration: float
) -> None:
    """Log processing completion."""
    logger = get_logger()
    logger.info("-" * 60)
    logger.info(f"Processing complete:")
    logger.info(f"  Files processed: {files}")
    logger.info(f"  Lines processed: {lines_processed}")
    logger.info(f"  Lines rejected: {lines_rejected}")
    logger.info(f"  Duration: {duration:.2f}s")
    logger.info("=" * 60)


def log_error(message: str, exc: Exception | None = None) -> None:
    """Log an error."""
    logger = get_logger()
    if exc:
        logger.error(f"{message}: {exc}", exc_info=True)
    else:
        logger.error(message)


def log_backup_created(source: Path, dest: Path) -> None:
    """Log a backup creation."""
    logger = get_logger()
    logger.info(f"Backup created: {source.name} -> {dest}")
