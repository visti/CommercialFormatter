"""Field formatting utilities for date, time, and duration fields."""

from __future__ import annotations

import app_logging as logging
from settings import get_settings

# Constants for time calculations
SECONDS_PER_DAY = 86400
MINUTES_PER_DAY = 1440


class FieldFormatter:
    """Formats and normalizes field values for broadcast metadata."""

    def __init__(self, overflow_threshold: int | None = None):
        """Initialize formatter.

        Args:
            overflow_threshold: Minutes threshold for overflow detection.
                               If None, loads from settings on first use.
        """
        self._overflow_threshold = overflow_threshold

    @property
    def overflow_threshold(self) -> int:
        """Get overflow threshold, loading from settings if needed."""
        if self._overflow_threshold is None:
            settings = get_settings()
            self._overflow_threshold = settings.thresholds.overflow_threshold_minutes
        return self._overflow_threshold

    @staticmethod
    def format_date(date_str: str) -> str:
        """Normalize date to DD-MM-YYYY format.

        Handles:
        - YYMMDD (6 digits) -> DD-MM-YYYY
        - DDMMYYYY (8 digits) -> DD-MM-YYYY
        - YYYY-MM-DD -> DD-MM-YYYY
        - DD.MM.YYYY -> DD-MM-YYYY
        - DD-MM-YYYY -> unchanged

        Args:
            date_str: Date string to normalize.

        Returns:
            Normalized date string in DD-MM-YYYY format.
        """
        date_str = date_str.strip()

        # Handle 6-digit format: YYMMDD
        if len(date_str) == 6 and date_str.isdigit():
            yy = date_str[0:2]
            mm = date_str[2:4]
            dd = date_str[4:6]
            year = f"20{yy}" if int(yy) < 50 else f"19{yy}"
            return f"{dd}-{mm}-{year}"

        # Handle 8-digit format: DDMMYYYY
        if len(date_str) == 8 and date_str.isdigit():
            dd = date_str[0:2]
            mm = date_str[2:4]
            yyyy = date_str[4:8]
            return f"{dd}-{mm}-{yyyy}"

        # Handle YYYY-MM-DD format (convert to DD-MM-YYYY)
        if len(date_str) == 10 and date_str[4] == "-" and date_str[7] == "-":
            parts = date_str.split("-")
            if len(parts) == 3 and len(parts[0]) == 4:
                yyyy, mm, dd = parts
                return f"{dd}-{mm}-{yyyy}"

        # Handle DD.MM.YYYY format (convert to DD-MM-YYYY)
        if len(date_str) == 10 and date_str[2] == "." and date_str[5] == ".":
            return date_str.replace(".", "-")

        # Already DD-MM-YYYY or unknown format, return as-is
        return date_str

    @staticmethod
    def format_time(time_str: str) -> str:
        """Format a 6-digit time string (HHMMSS) to HH:MM:SS.

        Args:
            time_str: Time string to format.

        Returns:
            Formatted time string, or original if not 6 digits.
        """
        if len(time_str) == 6 and time_str.isdigit():
            return f"{time_str[0:2]}:{time_str[2:4]}:{time_str[4:6]}"
        return time_str

    def format_duration(self, duration_str: str) -> str:
        """Fix buffer overflow in playing time for tracks spanning midnight.

        Some stations incorrectly calculate playing time by adding 1440 minutes
        when tracks span midnight.

        Args:
            duration_str: Duration string in MM:SS or HH:MM:SS format.

        Returns:
            Corrected duration string.
        """
        if ":" not in duration_str:
            return duration_str

        parts = duration_str.split(":")
        if len(parts) < 2:
            return duration_str

        try:
            minutes = int(parts[0])
            seconds = int(parts[1])
        except ValueError:
            return duration_str

        if minutes < self.overflow_threshold:
            return duration_str

        # Convert to total seconds and correct for day overflow
        total_seconds = minutes * 60 + seconds
        correct_seconds = SECONDS_PER_DAY - total_seconds

        if correct_seconds < 0:
            return duration_str

        # Convert back to MM:SS
        correct_minutes = correct_seconds // 60
        correct_secs = correct_seconds % 60
        corrected = f"{correct_minutes:02d}:{correct_secs:02d}"

        logging.log_overflow_fix(0, duration_str, corrected)

        return corrected

    @staticmethod
    def get_duration_minutes(duration_str: str) -> int | None:
        """Extract total minutes from a duration string.

        Args:
            duration_str: Duration string in MM:SS format.

        Returns:
            Total minutes, or None if parsing fails.
        """
        if ":" not in duration_str:
            return None

        parts = duration_str.split(":")
        if len(parts) < 2:
            return None

        try:
            return int(parts[0])
        except ValueError:
            return None


# Module-level instance for convenience
_formatter: FieldFormatter | None = None


def get_formatter() -> FieldFormatter:
    """Get the global formatter instance."""
    global _formatter
    if _formatter is None:
        _formatter = FieldFormatter()
    return _formatter


# Convenience functions that use the global formatter
def format_date(date_str: str) -> str:
    """Normalize date to DD-MM-YYYY format."""
    return FieldFormatter.format_date(date_str)


def format_time(time_str: str) -> str:
    """Format a 6-digit time string to HH:MM:SS."""
    return FieldFormatter.format_time(time_str)


def format_duration(duration_str: str) -> str:
    """Fix buffer overflow in playing time."""
    return get_formatter().format_duration(duration_str)


def get_duration_minutes(duration_str: str) -> int | None:
    """Extract total minutes from a duration string."""
    return FieldFormatter.get_duration_minutes(duration_str)
