"""Core file processing logic for commercial formatter."""

from __future__ import annotations

import json
import re
import shutil
import subprocess
import sys
from collections import Counter
from contextlib import ExitStack
from datetime import datetime
from pathlib import Path
from typing import IO, TYPE_CHECKING

import chardet

import app_logging as logging
import output as console
from choices import get_choices_manager
from config import ADDITIONAL_POSTFIX, DELETE_COLS_SCRIPT, REJECTDIR
from settings import get_settings
from stations import Station

if TYPE_CHECKING:
    from output import ProcessingStats


# Checkpoint file for error recovery
CHECKPOINT_FILE = ".komm_fmt_checkpoint.json"


def detect_encoding(file_path: Path) -> str:
    """Detect file encoding using chardet.

    Args:
        file_path: Path to the file to detect encoding for.

    Returns:
        Detected encoding string, defaulting to utf-8 if detection fails.
    """
    with open(file_path, "rb") as f:
        raw_data = f.read()
    result = chardet.detect(raw_data)
    encoding = result.get("encoding", "utf-8")
    # Handle common encoding aliases
    if encoding and encoding.lower() in ("ascii", "iso-8859-1", "latin-1", "latin1"):
        # These are often misdetected; try cp1252 which is a superset
        encoding = "cp1252"
    return encoding or "utf-8"


def backup_files(files: list[Path], backup_dir: str = "backup") -> Path | None:
    """Create backups of input files before processing.

    Args:
        files: List of files to backup.
        backup_dir: Name of backup directory.

    Returns:
        Path to backup directory, or None if backup is disabled.
    """
    settings = get_settings()
    if not settings.backup.enabled:
        return None

    backup_path = Path.cwd() / backup_dir
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_path = backup_path / timestamp

    backup_path.mkdir(parents=True, exist_ok=True)

    for file_path in files:
        dest = backup_path / file_path.name
        shutil.copy2(file_path, dest)
        logging.log_backup_created(file_path, dest)

    console.info(f"Backup created: {console.dim(str(backup_path))}")
    return backup_path


def get_files(station: Station, exclude_filename: str = "") -> list[Path]:
    """Find files in current directory matching station's extensions.

    Args:
        station: Station configuration with file extensions.
        exclude_filename: Filename to exclude from results.

    Returns:
        List of matching file paths.

    Raises:
        SystemExit: If no eligible files are found.
    """
    cwd = Path.cwd()
    files: list[Path] = []

    # Pre-lowercase extensions for efficient comparison
    extensions_lower = [ext.lower() for ext in station.ext]

    for f in cwd.iterdir():
        if f.is_file() and f.name != exclude_filename:
            name_lower = f.name.lower()
            if any(ext in name_lower for ext in extensions_lower):
                files.append(f)

    if not files:
        console.error("No eligible files found.")
        sys.exit(1)

    console.info(f"Found {console.bold(str(len(files)))} file(s):")
    for f in files:
        console.info(f"  {console.dim('-')} {f.name}")
    print()

    return files


def save_checkpoint(
    checkpoint_data: dict,
    output_dir: Path | None = None,
) -> None:
    """Save processing checkpoint for error recovery.

    Args:
        checkpoint_data: Data to save in checkpoint.
        output_dir: Directory for checkpoint file.
    """
    if output_dir is None:
        output_dir = Path.cwd()

    checkpoint_path = output_dir / CHECKPOINT_FILE
    try:
        with open(checkpoint_path, "w", encoding="utf-8") as f:
            json.dump(checkpoint_data, f, indent=2, default=str)
    except Exception as e:
        logging.log_error(f"Failed to save checkpoint: {e}")


def load_checkpoint(output_dir: Path | None = None) -> dict | None:
    """Load processing checkpoint if it exists.

    Args:
        output_dir: Directory containing checkpoint file.

    Returns:
        Checkpoint data dict, or None if no checkpoint exists.
    """
    if output_dir is None:
        output_dir = Path.cwd()

    checkpoint_path = output_dir / CHECKPOINT_FILE
    if not checkpoint_path.exists():
        return None

    try:
        with open(checkpoint_path, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return None


def clear_checkpoint(output_dir: Path | None = None) -> None:
    """Remove checkpoint file after successful processing.

    Args:
        output_dir: Directory containing checkpoint file.
    """
    if output_dir is None:
        output_dir = Path.cwd()

    checkpoint_path = output_dir / CHECKPOINT_FILE
    if checkpoint_path.exists():
        checkpoint_path.unlink()


def read_files(
    files: list[Path],
    station: Station,
    stats: ProcessingStats | None = None,
) -> tuple[str, set[int]]:
    """Read and concatenate all input files, applying station-specific transformations.

    Args:
        files: List of input file paths.
        station: Station configuration.
        stats: Optional stats object to update.

    Returns:
        Tuple of (concatenated content string, set of line indices to reject).
    """
    joined_content: list[str] = []
    total_files = len(files)
    all_dash_reject_indices: set[int] = set()
    choices_manager = get_choices_manager()
    settings = get_settings()

    for idx, file_path in enumerate(files, 1):
        console.progress(idx, total_files, f"Reading {file_path.name}")
        if stats:
            stats.files_processed = idx

        try:
            # Detect encoding for each file
            encoding = detect_encoding(file_path)
            console.info(console.dim(f"    Encoding: {encoding}"))
            content = file_path.read_text(encoding=encoding, errors="replace")
        except Exception as e:
            console.error(f"Could not read file {file_path}: {e}")
            logging.log_error(f"Could not read file {file_path}", e)
            continue

        lines = content.splitlines()
        logging.log_file_read(file_path, encoding, len(lines))

        # Skip header lines based on station config
        if station.skip_lines > 0 and lines:
            lines = lines[station.skip_lines :]

        # Apply configured transformations
        if station.transformations:
            lines = station.apply_transformations(lines, filename=file_path.stem)

        # Fix artist/title split error where title contains " - " (configurable)
        if station.fix_artist_title_split:
            title_idx = station.get_field_index("title", -1)
            artist_idx = station.get_field_index("artist", -1)

            if title_idx >= 0 and artist_idx >= 0:
                # Skip lines that match stopwords (they'll be rejected later anyway)
                issues_by_key: dict[tuple[str, str], list[int]] = {}
                for i, line in enumerate(lines):
                    if line.strip():
                        if station.matches_stopword_lower(line.lower()):
                            continue
                        fields = line.split(station.separator)
                        if len(fields) > max(title_idx, artist_idx) and " - " in fields[title_idx]:
                            key = (fields[title_idx], fields[artist_idx])
                            if key not in issues_by_key:
                                issues_by_key[key] = []
                            issues_by_key[key].append(i)

                # Prompt for each unique issue
                if issues_by_key:
                    console.warning(
                        f"Found {len(issues_by_key)} unique artist/title split issue(s)"
                    )
                    print()

                lines_to_fix: list[int] = []
                dash_reject_indices: set[int] = set()

                for (title, artist), indices in issues_by_key.items():
                    title_parts = title.split(" - ", 1)
                    fixed_title = title_parts[1]
                    fixed_artist = f"{artist}-{title_parts[0]}"

                    count = len(indices)

                    # Check for remembered choice
                    remembered = choices_manager.get_artist_title_choice(title, artist)
                    if remembered:
                        if remembered == "fix":
                            lines_to_fix.extend(indices)
                            console.info(
                                f"  Auto-fix (remembered): \"{title}\" -> \"{fixed_title}\" ({count}x)"
                            )
                            logging.log_user_choice(
                                "artist_title", f"{title} by {artist}", "fix (remembered)"
                            )
                        elif remembered == "reject":
                            dash_reject_indices.update(indices)
                            console.info(
                                f"  Auto-reject (remembered): \"{title}\" ({count}x)"
                            )
                            logging.log_user_choice(
                                "artist_title", f"{title} by {artist}", "reject (remembered)"
                            )
                        else:
                            console.info(f"  Auto-skip (remembered): \"{title}\" ({count}x)")
                            logging.log_user_choice(
                                "artist_title", f"{title} by {artist}", "skip (remembered)"
                            )
                        continue

                    console.info(
                        f"  Current: Title=\"{title}\" Artist=\"{artist}\" ({count}x)"
                    )
                    console.info(
                        f"  Fixed:   Title=\"{fixed_title}\" Artist=\"{fixed_artist}\""
                    )

                    response = input("  [Y]es fix / [N]o skip / [X] reject: ").strip().lower()
                    if response in ("y", "yes", ""):
                        lines_to_fix.extend(indices)
                        console.success(f"  Will fix {count} occurrence(s)")
                        choices_manager.remember_artist_title_choice(title, artist, "fix")
                        logging.log_user_choice("artist_title", f"{title} by {artist}", "fix")
                    elif response in ("x", "reject"):
                        dash_reject_indices.update(indices)
                        console.info(f"  Will reject {count} occurrence(s)")
                        choices_manager.remember_artist_title_choice(title, artist, "reject")
                        logging.log_user_choice(
                            "artist_title", f"{title} by {artist}", "reject"
                        )
                    else:
                        console.info("  Skipped")
                        choices_manager.remember_artist_title_choice(title, artist, "skip")
                        logging.log_user_choice("artist_title", f"{title} by {artist}", "skip")
                    print()

                # Apply fixes
                for i in lines_to_fix:
                    fields = lines[i].split(station.separator)
                    title_parts = fields[title_idx].split(" - ", 1)
                    fields[title_idx] = title_parts[1]
                    fields[artist_idx] = f"{fields[artist_idx]}-{title_parts[0]}"
                    lines[i] = station.separator.join(fields)

                # Add dash reject indices with offset for global position
                offset = len(joined_content)
                all_dash_reject_indices.update(i + offset for i in dash_reject_indices)

        joined_content.extend(lines)

    # Check for long playing times in CSV-format stations
    if not station.positional:
        # Get field indices from station config (auto-derived from headlines)
        title_idx = station.get_field_index("title", -1)
        artist_idx = station.get_field_index("artist", -1)
        time_idx = station.get_field_index("duration", -1)

        if title_idx >= 0 and artist_idx >= 0 and time_idx >= 0:
            joined_content, reject_indices = check_long_playing_times(
                joined_content, station, title_idx, artist_idx, time_idx
            )
        else:
            reject_indices = set()
    else:
        reject_indices = set()

    # Merge dash delimiter rejections with other rejections
    reject_indices.update(all_dash_reject_indices)

    return "\n".join(joined_content), reject_indices


def make_additional_filename(base_path: Path) -> Path:
    """Create additional output filename with postfix.

    Args:
        base_path: Base output file path.

    Returns:
        Path for additional output file.
    """
    stem = base_path.stem
    ext = base_path.suffix
    return base_path.parent / f"{stem}{ADDITIONAL_POSTFIX}{ext}"


def generate_rejection_filename(station: Station) -> Path:
    """Generate the rejection file path with current date.

    Args:
        station: Station configuration.

    Returns:
        Path for rejection file.
    """
    now = datetime.now()
    date_str = f"{now.year}-{now.month}-{now.day}"
    filename = f"{date_str}-reject-{station.name}.csv"
    return REJECTDIR / filename


def format_time_field(time_str: str) -> str:
    """Format a 6-digit time string (HHMMSS) to HH:MM:SS.

    Args:
        time_str: Time string to format.

    Returns:
        Formatted time string.
    """
    if len(time_str) == 6:
        return f"{time_str[0:2]}:{time_str[2:4]}:{time_str[4:6]}"
    return time_str


def fix_overflow_playing_time(playing_time: str, threshold: int | None = None) -> str:
    """Fix buffer overflow in playing time for tracks starting before midnight.

    Some stations incorrectly calculate playing time by adding 1440 minutes
    (minutes in a day) when tracks span midnight.

    Args:
        playing_time: Time string in MM:SS format.
        threshold: Minutes threshold for overflow detection. If None, uses settings.

    Returns:
        Corrected time string in MM:SS format.
    """
    if threshold is None:
        settings = get_settings()
        threshold = settings.thresholds.overflow_threshold_minutes

    if ":" not in playing_time:
        return playing_time

    parts = playing_time.split(":")
    if len(parts) < 2:
        return playing_time

    try:
        minutes = int(parts[0])
        seconds = int(parts[1])
    except ValueError:
        return playing_time

    if minutes < threshold:
        return playing_time

    # Convert to total seconds
    total_seconds = minutes * 60 + seconds
    # 1440 minutes = 86400 seconds (one day)
    correct_seconds = 86400 - total_seconds

    if correct_seconds < 0:
        return playing_time

    # Convert back to MM:SS
    correct_minutes = correct_seconds // 60
    correct_secs = correct_seconds % 60

    logging.log_overflow_fix(0, playing_time, f"{correct_minutes:02d}:{correct_secs:02d}")

    return f"{correct_minutes:02d}:{correct_secs:02d}"


def get_playing_time_minutes(playing_time: str) -> int | None:
    """Extract total minutes from a playing time string.

    Args:
        playing_time: Time string in MM:SS format.

    Returns:
        Total minutes, or None if parsing fails.
    """
    if ":" not in playing_time:
        return None

    parts = playing_time.split(":")
    if len(parts) < 2:
        return None

    try:
        return int(parts[0])
    except ValueError:
        return None


def check_long_playing_times(
    lines: list[str],
    station: Station,
    title_idx: int = 7,
    artist_idx: int = 8,
    time_idx: int = 2,
) -> tuple[list[str], set[int]]:
    """Check for tracks with long playing times and prompt user.

    Args:
        lines: List of input lines.
        station: Station configuration.
        title_idx: Field index for track title.
        artist_idx: Field index for artist.
        time_idx: Field index for playing time.

    Returns:
        Tuple of (modified lines, set of line indices to reject).
    """
    settings = get_settings()
    threshold = settings.thresholds.long_playing_time_minutes
    choices_manager = get_choices_manager()

    # Find lines with long playing times, grouped by unique track
    issues_by_key: dict[tuple[str, str, str], list[int]] = {}

    for i, line in enumerate(lines):
        if not line.strip():
            continue

        fields = line.split(";")
        if len(fields) <= max(title_idx, artist_idx, time_idx):
            continue

        # Get playing time and apply overflow fix first
        playing_time = fix_overflow_playing_time(fields[time_idx])
        minutes = get_playing_time_minutes(playing_time)

        if minutes is not None and minutes >= threshold:
            title = fields[title_idx] if len(fields) > title_idx else ""
            artist = fields[artist_idx] if len(fields) > artist_idx else ""
            key = (title, artist, playing_time)

            if key not in issues_by_key:
                issues_by_key[key] = []
            issues_by_key[key].append(i)

    if not issues_by_key:
        return lines, set()

    console.warning(
        f"Found {len(issues_by_key)} unique track(s) with playing time over {threshold} minutes"
    )
    print()

    lines_to_reject: set[int] = set()
    modified_lines = lines.copy()

    for (title, artist, playing_time), indices in issues_by_key.items():
        count = len(indices)

        # Check for remembered choice
        remembered_action, remembered_time = choices_manager.get_playing_time_choice(
            title, artist, playing_time
        )

        if remembered_action:
            if remembered_action == "accept":
                console.info(
                    f"  Auto-accept (remembered): \"{title}\" - {playing_time} ({count}x)"
                )
                logging.log_user_choice(
                    "long_time", f"{title} ({playing_time})", "accept (remembered)"
                )
            elif remembered_action == "reject":
                lines_to_reject.update(indices)
                console.info(
                    f"  Auto-reject (remembered): \"{title}\" - {playing_time} ({count}x)"
                )
                logging.log_user_choice(
                    "long_time", f"{title} ({playing_time})", "reject (remembered)"
                )
            elif remembered_action == "edit" and remembered_time:
                for idx in indices:
                    fields = modified_lines[idx].split(";")
                    if len(fields) > time_idx:
                        fields[time_idx] = remembered_time
                        modified_lines[idx] = ";".join(fields)
                console.info(
                    f"  Auto-edit (remembered): \"{title}\" -> {remembered_time} ({count}x)"
                )
                logging.log_user_choice(
                    "long_time", f"{title} ({playing_time})", f"edit to {remembered_time} (remembered)"
                )
            continue

        console.info(f'  Title:  "{title}"')
        console.info(f'  Artist: "{artist}"')
        console.info(f"  Time:   {playing_time} ({count}x)")

        while True:
            response = input("  [A]ccept / [R]eject / [E]dit: ").strip().lower()

            if response in ("a", "accept", ""):
                console.success(f"  Accepted {count} occurrence(s)")
                choices_manager.remember_playing_time_choice(
                    title, artist, playing_time, "accept"
                )
                logging.log_user_choice("long_time", f"{title} ({playing_time})", "accept")
                break

            elif response in ("r", "reject"):
                lines_to_reject.update(indices)
                console.info(f"  Will reject {count} occurrence(s)")
                choices_manager.remember_playing_time_choice(
                    title, artist, playing_time, "reject"
                )
                logging.log_user_choice("long_time", f"{title} ({playing_time})", "reject")
                break

            elif response in ("e", "edit"):
                new_time = input("  Enter corrected time (MM:SS): ").strip()
                if ":" in new_time:
                    parts = new_time.split(":")
                    try:
                        int(parts[0])
                        int(parts[1])
                        for idx in indices:
                            fields = modified_lines[idx].split(";")
                            if len(fields) > time_idx:
                                fields[time_idx] = new_time
                                modified_lines[idx] = ";".join(fields)
                        console.success(f"  Updated {count} occurrence(s) to {new_time}")
                        choices_manager.remember_playing_time_choice(
                            title, artist, playing_time, "edit", new_time
                        )
                        logging.log_user_choice(
                            "long_time", f"{title} ({playing_time})", f"edit to {new_time}"
                        )
                        break
                    except ValueError:
                        console.error("  Invalid time format. Use MM:SS (e.g., 03:45)")
                else:
                    console.error("  Invalid time format. Use MM:SS (e.g., 03:45)")

            else:
                console.error("  Invalid choice. Enter A, R, or E")

        print()

    return modified_lines, lines_to_reject


def check_duplicates(
    lines: list[str],
    station: Station,
    title_idx: int = 7,
    artist_idx: int = 8,
    date_idx: int = 0,
) -> tuple[list[str], set[int]]:
    """Check for duplicate tracks and prompt user.

    Args:
        lines: List of input lines.
        station: Station configuration.
        title_idx: Field index for track title.
        artist_idx: Field index for artist.
        date_idx: Field index for date.

    Returns:
        Tuple of (lines, set of line indices to reject as duplicates).
    """
    settings = get_settings()
    if not settings.duplicates.enabled:
        return lines, set()

    # Track seen combinations: (title, artist, date) -> first line index
    seen: dict[tuple[str, str, str], int] = {}
    duplicates: dict[tuple[str, str, str], list[int]] = {}

    for i, line in enumerate(lines):
        if not line.strip():
            continue

        fields = line.split(";")
        if len(fields) <= max(title_idx, artist_idx, date_idx):
            continue

        title = fields[title_idx].strip().lower() if len(fields) > title_idx else ""
        artist = fields[artist_idx].strip().lower() if len(fields) > artist_idx else ""
        date = fields[date_idx].strip() if len(fields) > date_idx else ""

        key = (title, artist, date)

        if key in seen:
            if key not in duplicates:
                duplicates[key] = [seen[key]]
            duplicates[key].append(i)
            logging.log_duplicate_found(i, title, artist, date, seen[key])
        else:
            seen[key] = i

    if not duplicates:
        return lines, set()

    total_dups = sum(len(indices) - 1 for indices in duplicates.values())
    console.warning(f"Found {total_dups} duplicate track(s) across {len(duplicates)} unique tracks")
    print()

    lines_to_reject: set[int] = set()
    action = settings.duplicates.action

    if action == "reject":
        # Auto-reject all duplicates (keep first occurrence)
        for indices in duplicates.values():
            lines_to_reject.update(indices[1:])  # Skip first (original)
        console.info(f"  Auto-rejected {total_dups} duplicate(s)")
    elif action == "keep":
        console.info("  Keeping all duplicates (as configured)")
    else:  # prompt
        for key, indices in duplicates.items():
            title, artist, date = key
            dup_count = len(indices) - 1

            console.info(f'  Title:  "{title}"')
            console.info(f'  Artist: "{artist}"')
            console.info(f"  Date:   {date}")
            console.info(f"  Found {dup_count} duplicate(s) (lines: {indices[1:]})")

            response = input("  [K]eep all / [R]eject duplicates: ").strip().lower()
            if response in ("r", "reject"):
                lines_to_reject.update(indices[1:])
                console.info(f"  Rejected {dup_count} duplicate(s)")
                logging.log_user_choice(
                    "duplicate", f"{title} by {artist} on {date}", "reject"
                )
            else:
                console.info("  Keeping all")
                logging.log_user_choice(
                    "duplicate", f"{title} by {artist} on {date}", "keep"
                )
            print()

    return lines, lines_to_reject


def format_date_field(date_str: str) -> str:
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


def process_positional_line(
    line: str, sorted_positions: list[int], separator: str
) -> str:
    """Extract fields from a fixed-width line using pre-sorted positional indices.

    Args:
        line: Input line to process.
        sorted_positions: List of column boundary positions.
        separator: Output field separator.

    Returns:
        Processed line with fields separated.
    """
    parts: list[str] = []
    start = 0

    for pos in sorted_positions:
        effective_pos = min(pos, len(line))
        if effective_pos >= start:
            part = line[start:effective_pos].strip()
            parts.append(part)
            start = effective_pos

    # Format first field (date) if present
    if parts:
        parts[0] = format_date_field(parts[0])

    return separator.join(parts)


def process_csv_line(line: str, separator: str = ";", input_separator: str = ";") -> str:
    """Process a CSV-delimited line, formatting date and time fields.

    Args:
        line: Input line to process.
        separator: Output field separator.
        input_separator: Input field separator for parsing.

    Returns:
        Processed line with formatted fields.
    """
    fields = line.split(input_separator)

    # Format date field (index 0)
    if fields:
        fields[0] = format_date_field(fields[0])

    # Format time field (index 1) if it's 6 digits
    if len(fields) > 1 and len(fields[1]) == 6:
        fields[1] = format_time_field(fields[1])

    # Fix overflow playing time (index 2)
    if len(fields) > 2:
        fields[2] = fix_overflow_playing_time(fields[2])

    return separator.join(fields)


def write_headlines(file_handle: IO[str], station: Station) -> None:
    """Write headline row to file.

    Args:
        file_handle: File handle to write to.
        station: Station configuration with headlines.
    """
    headlines_line = station.separator.join(station.headlines)
    file_handle.write(headlines_line + "\n")


def clean_empty_file(file_path: Path) -> None:
    """Remove file if it's empty or contains only a header line.

    Args:
        file_path: Path to file to check and potentially remove.
    """
    if file_path.exists():
        content = file_path.read_text(encoding="utf-8", errors="replace")
        line_count = content.count("\n")
        if line_count <= 1:
            file_path.unlink()
            console.info(console.dim(f"Removed empty file: {file_path.name}"))


def run_delete_columns(output_path: Path) -> None:
    """Run the delete_columns.py script on the output file.

    Args:
        output_path: Path to the output file to process.
    """
    try:
        subprocess.run(
            [sys.executable, str(DELETE_COLS_SCRIPT), str(output_path)],
            check=True,
        )
    except subprocess.CalledProcessError as e:
        console.warning(f"delete_columns.py failed: {e}")
        logging.log_error(f"delete_columns.py failed", e)
    except FileNotFoundError:
        console.warning(f"Could not find {DELETE_COLS_SCRIPT}")


def process_line(line: str, station: Station) -> str:
    """Process a single line according to station format.

    Args:
        line: Input line to process.
        station: Station configuration.

    Returns:
        Processed line string.
    """
    if station.positional:
        return process_positional_line(line, station.sorted_positions, station.separator)
    else:
        return process_csv_line(line, station.separator, station.input_separator)


def process_files(
    content: str,
    station: Station,
    output_file: Path,
    additional_filter: str = "",
    use_stopwords: bool = True,
    stats: ProcessingStats | None = None,
    force_reject_indices: set[int] | None = None,
) -> None:
    """Main processing pipeline for station data.

    Args:
        content: Concatenated content to process.
        station: Station configuration.
        output_file: Path for main output file.
        additional_filter: Optional filter for routing to additional file.
        use_stopwords: Whether to apply stopword filtering.
        stats: Optional stats object to update.
        force_reject_indices: Set of line indices to force reject.
    """
    # Track counts
    rejected = 0
    processed = 0
    stopword_counts: Counter[str] = Counter()

    # Prepare file paths
    has_additional = bool(additional_filter)
    additional_path = make_additional_filename(output_file) if has_additional else None
    reject_path = generate_rejection_filename(station)

    # Pre-lowercase additional filter for fast comparison
    additional_filter_lower = additional_filter.lower() if has_additional else ""

    # Ensure rejection directory exists
    reject_path.parent.mkdir(parents=True, exist_ok=True)

    console.info("Processing lines...")

    # Save checkpoint at start
    checkpoint_data = {
        "station": station.name,
        "output_file": str(output_file),
        "start_time": datetime.now().isoformat(),
        "lines_processed": 0,
    }
    save_checkpoint(checkpoint_data)

    try:
        # Open all output files using ExitStack for proper context management
        with ExitStack() as stack:
            out_f = stack.enter_context(open(output_file, "w", encoding="utf-8"))
            reject_f = stack.enter_context(open(reject_path, "w", encoding="utf-8"))

            additional_f: IO[str] | None = None
            if has_additional:
                additional_f = stack.enter_context(
                    open(additional_path, "w", encoding="utf-8")
                )
                console.info(f"Additional filter: {console.cyan(additional_filter)}")
                if stats:
                    stats.additional_file = additional_path.name
                write_headlines(additional_f, station)

            # Write headlines
            write_headlines(out_f, station)
            write_headlines(reject_f, station)

            # Process each line
            line_index = 0
            lines = content.splitlines()
            total_lines = len(lines)

            for line in lines:
                if not line.strip():
                    line_index += 1
                    continue

                # Check if line was marked for forced rejection
                if force_reject_indices and line_index in force_reject_indices:
                    reject_line = process_line(line, station)
                    reject_f.write(reject_line + "\n")
                    rejected += 1
                    logging.log_rejection(line_index, "forced_reject", line)
                    line_index += 1
                    continue

                # Pre-compute lowercase once for efficiency
                line_lower = line.lower()

                # Check stopwords using pre-compiled regex
                if use_stopwords and station.matches_stopword_lower(line_lower):
                    reject_line = process_line(line, station)
                    reject_f.write(reject_line + "\n")
                    rejected += 1

                    # Track which stopword matched for summary
                    matched_stopword = station.get_matched_stopword(line_lower)
                    if matched_stopword:
                        stopword_counts[matched_stopword] += 1
                        logging.log_stopword_match(line_index, matched_stopword)

                    line_index += 1
                    continue

                # Process the line
                output_line = process_line(line, station)
                processed += 1

                # Check for additional routing
                if has_additional and additional_filter_lower in line_lower:
                    additional_f.write(output_line + "\n")
                else:
                    out_f.write(output_line + "\n")

                line_index += 1

                # Update checkpoint periodically
                if line_index % 1000 == 0:
                    checkpoint_data["lines_processed"] = line_index
                    save_checkpoint(checkpoint_data)

        # Update stats
        if stats:
            stats.lines_processed = processed
            stats.lines_rejected = rejected
            stats.stopword_counts = stopword_counts

        # Log completion
        logging.log_processing_complete(
            stats.files_processed if stats else 0,
            processed,
            rejected,
            stats.elapsed() if stats else 0,
        )

        # Clear checkpoint on success
        clear_checkpoint()

    except Exception as e:
        logging.log_error("Processing failed", e)
        console.error(f"Processing failed: {e}")
        console.info("Checkpoint saved. Run again to attempt recovery.")
        raise

    # Cleanup phase
    if has_additional and additional_path:
        clean_empty_file(additional_path)

    clean_empty_file(reject_path)

    # Check if output file is empty (only header)
    if output_file.exists():
        output_content = output_file.read_text(encoding="utf-8", errors="replace")
        if output_content.count("\n") > 1:
            run_delete_columns(output_file)
        else:
            clean_empty_file(output_file)

    # Print stopword summary
    if stopword_counts and stats:
        print_stopword_summary(stopword_counts)


def print_stopword_summary(stopword_counts: Counter[str], top_n: int = 10) -> None:
    """Print a summary of stopwords that caused rejections.

    Args:
        stopword_counts: Counter of stopword matches.
        top_n: Number of top stopwords to show.
    """
    if not stopword_counts:
        return

    print()
    console.info(console.bold("Rejection Summary (top stopwords):"))
    for stopword, count in stopword_counts.most_common(top_n):
        console.info(f"  {count:5,}x  {stopword}")
