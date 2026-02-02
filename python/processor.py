"""Core file processing logic for commercial formatter."""

from __future__ import annotations

import json
import re
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
from decisions import DecisionManager, ARTIST_TITLE_CONFIG, LONG_TIME_CONFIG, DUPLICATE_CONFIG, DecisionConfig, Option
from formatters import format_date, format_time, format_duration, get_duration_minutes
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


def check_artist_title_split(
    lines: list[str],
    station: Station,
) -> tuple[list[str], set[int]]:
    """Check for and fix artist/title split issues where title contains ' - '.

    Args:
        lines: Input lines to check.
        station: Station configuration.

    Returns:
        Tuple of (modified lines, set of line indices to reject).
    """
    title_idx = station.get_field_index("title", -1)
    artist_idx = station.get_field_index("artist", -1)

    if title_idx < 0 or artist_idx < 0:
        return lines, set()

    choices_manager = get_choices_manager()

    # Find issues, skipping lines that match stopwords
    issues: dict[tuple[str, str], list[int]] = {}
    for i, line in enumerate(lines):
        if not line.strip():
            continue
        if station.matches_stopword_lower(line.lower()):
            continue
        fields = line.split(station.separator)
        if len(fields) > max(title_idx, artist_idx) and " - " in fields[title_idx]:
            key = (fields[title_idx], fields[artist_idx])
            if key not in issues:
                issues[key] = []
            issues[key].append(i)

    if not issues:
        return lines, set()

    # Track lines to fix
    lines_to_fix: list[int] = []
    modified_lines = lines.copy()

    def display_issue(key: tuple, indices: list[int], count: int) -> None:
        title, artist = key
        title_parts = title.split(" - ", 1)
        fixed_title = title_parts[1]
        fixed_artist = f"{artist}-{title_parts[0]}"
        console.info(f"  Current: Title=\"{title}\" Artist=\"{artist}\" ({count}x)")
        console.info(f"  Fixed:   Title=\"{fixed_title}\" Artist=\"{fixed_artist}\"")

    def apply_action(action: str, key: tuple, indices: list[int], _extra: any) -> set[int]:
        if action == "fix":
            lines_to_fix.extend(indices)
            console.success(f"  Will fix {len(indices)} occurrence(s)")
            return set()
        elif action == "reject":
            console.info(f"  Will reject {len(indices)} occurrence(s)")
            return set(indices)
        else:
            console.info("  Skipped")
            return set()

    def get_remembered(key: tuple) -> str | None:
        return choices_manager.get_artist_title_choice(key[0], key[1])

    def remember_choice(key: tuple, action: str, _extra: any) -> None:
        choices_manager.remember_artist_title_choice(key[0], key[1], action)

    manager = DecisionManager(
        ARTIST_TITLE_CONFIG,
        get_remembered=get_remembered,
        remember_choice=remember_choice,
    )

    reject_indices = manager.process_issues(
        issues,
        display_issue=display_issue,
        apply_action=apply_action,
        summary_message=f"Found {len(issues)} unique artist/title split issue(s)",
    )

    # Apply fixes to lines
    for i in lines_to_fix:
        fields = modified_lines[i].split(station.separator)
        title_parts = fields[title_idx].split(" - ", 1)
        fields[title_idx] = title_parts[1]
        fields[artist_idx] = f"{fields[artist_idx]}-{title_parts[0]}"
        modified_lines[i] = station.separator.join(fields)

    return modified_lines, reject_indices


def read_single_file(file_path: Path, station: Station) -> list[str]:
    """Read a single file and apply initial transformations.

    Args:
        file_path: Path to file to read.
        station: Station configuration.

    Returns:
        List of processed lines, or empty list on error.
    """
    try:
        encoding = detect_encoding(file_path)
        console.info(console.dim(f"    Encoding: {encoding}"))
        content = file_path.read_text(encoding=encoding, errors="replace")
    except Exception as e:
        console.error(f"Could not read file {file_path}: {e}")
        logging.log_error(f"Could not read file {file_path}", e)
        return []

    lines = content.splitlines()
    logging.log_file_read(file_path, encoding, len(lines))

    # Skip header lines
    if station.skip_lines > 0 and lines:
        lines = lines[station.skip_lines:]

    # Apply configured transformations
    if station.transformations:
        lines = station.apply_transformations(lines, filename=file_path.stem)

    return lines


def read_files(
    files: list[Path],
    station: Station,
    stats: ProcessingStats | None = None,
) -> tuple[str, set[int]]:
    """Read and concatenate all input files, applying transformations and validation.

    Args:
        files: List of input file paths.
        station: Station configuration.
        stats: Optional stats object to update.

    Returns:
        Tuple of (concatenated content string, set of line indices to reject).
    """
    joined_content: list[str] = []
    all_reject_indices: set[int] = set()
    total_files = len(files)

    # Phase 1: Read and transform files
    for idx, file_path in enumerate(files, 1):
        console.progress(idx, total_files, f"Reading {file_path.name}")
        if stats:
            stats.files_processed = idx

        lines = read_single_file(file_path, station)
        if not lines:
            continue

        # Phase 2: Check artist/title split (if configured)
        if station.fix_artist_title_split:
            lines, reject_indices = check_artist_title_split(lines, station)
            # Offset indices to global position
            offset = len(joined_content)
            all_reject_indices.update(i + offset for i in reject_indices)

        joined_content.extend(lines)

    # Phase 3: Check for multiple years
    joined_content = check_multiple_years(joined_content, station)

    # Phase 4: Check long playing times (CSV stations only)
    if not station.positional:
        title_idx = station.get_field_index("title", -1)
        artist_idx = station.get_field_index("artist", -1)
        time_idx = station.get_field_index("duration", -1)

        if title_idx >= 0 and artist_idx >= 0 and time_idx >= 0:
            joined_content, time_reject_indices = check_long_playing_times(
                joined_content, station, title_idx, artist_idx, time_idx
            )
            all_reject_indices.update(time_reject_indices)

    return "\n".join(joined_content), all_reject_indices


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
    issues: dict[tuple[str, str, str], list[int]] = {}
    sep = station.separator

    for i, line in enumerate(lines):
        if not line.strip():
            continue

        fields = line.split(sep)
        if len(fields) <= max(title_idx, artist_idx, time_idx):
            continue

        # Get playing time and apply overflow fix first
        playing_time = format_duration(fields[time_idx])
        minutes = get_duration_minutes(playing_time)

        if minutes is not None and minutes >= threshold:
            title = fields[title_idx] if len(fields) > title_idx else ""
            artist = fields[artist_idx] if len(fields) > artist_idx else ""
            key = (title, artist, playing_time)

            if key not in issues:
                issues[key] = []
            issues[key].append(i)

    if not issues:
        return lines, set()

    # Track edits to apply
    edits_to_apply: dict[int, str] = {}
    modified_lines = lines.copy()

    def display_issue(key: tuple, indices: list[int], count: int) -> None:
        title, artist, playing_time = key
        console.info(f'  Title:  "{title}"')
        console.info(f'  Artist: "{artist}"')
        console.info(f"  Time:   {playing_time} ({count}x)")

    def apply_action(action: str, key: tuple, indices: list[int], extra: any) -> set[int]:
        title, artist, playing_time = key
        count = len(indices)

        if action == "accept":
            console.success(f"  Accepted {count} occurrence(s)")
            return set()
        elif action == "reject":
            console.info(f"  Will reject {count} occurrence(s)")
            return set(indices)
        elif action == "edit" and extra:
            new_time = extra
            for idx in indices:
                edits_to_apply[idx] = new_time
            console.success(f"  Updated {count} occurrence(s) to {new_time}")
            return set()
        return set()

    def get_remembered(key: tuple) -> tuple[str, str | None] | None:
        title, artist, playing_time = key
        action, new_time = choices_manager.get_playing_time_choice(title, artist, playing_time)
        if action:
            return (action, new_time)
        return None

    def remember_choice(key: tuple, action: str, extra: any) -> None:
        title, artist, playing_time = key
        choices_manager.remember_playing_time_choice(title, artist, playing_time, action, extra)

    def handle_edit_input(action: str) -> any:
        if action != "edit":
            return None

        new_time = input("  Enter corrected time (MM:SS): ").strip()
        if ":" not in new_time:
            console.error("  Invalid time format. Use MM:SS (e.g., 03:45)")
            return False  # Signal to re-prompt

        parts = new_time.split(":")
        try:
            int(parts[0])
            int(parts[1])
            return new_time
        except ValueError:
            console.error("  Invalid time format. Use MM:SS (e.g., 03:45)")
            return False  # Signal to re-prompt

    manager = DecisionManager(
        LONG_TIME_CONFIG,
        get_remembered=get_remembered,
        remember_choice=remember_choice,
    )

    reject_indices = manager.process_issues(
        issues,
        display_issue=display_issue,
        apply_action=apply_action,
        summary_message=f"Found {len(issues)} unique track(s) with playing time over {threshold} minutes",
        extra_input_handler=handle_edit_input,
    )

    # Apply edits to lines
    for idx, new_time in edits_to_apply.items():
        fields = modified_lines[idx].split(sep)
        if len(fields) > time_idx:
            fields[time_idx] = new_time
            modified_lines[idx] = sep.join(fields)

    return modified_lines, reject_indices


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
    sep = station.separator

    for i, line in enumerate(lines):
        if not line.strip():
            continue

        fields = line.split(sep)
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
    action = settings.duplicates.action

    # Handle auto-actions from config
    if action == "reject":
        lines_to_reject = set()
        for indices in duplicates.values():
            lines_to_reject.update(indices[1:])  # Skip first (original)
        console.warning(f"Found {total_dups} duplicate track(s) - auto-rejected (as configured)")
        return lines, lines_to_reject
    elif action == "keep":
        console.warning(f"Found {total_dups} duplicate track(s) - keeping all (as configured)")
        return lines, set()

    # Prompt mode - use DecisionManager
    def display_issue(key: tuple, indices: list[int], count: int) -> None:
        title, artist, date = key
        dup_count = len(indices) - 1
        console.info(f'  Title:  "{title}"')
        console.info(f'  Artist: "{artist}"')
        console.info(f"  Date:   {date}")
        console.info(f"  Found {dup_count} duplicate(s) (lines: {indices[1:]})")

    def apply_action(action: str, key: tuple, indices: list[int], _extra: any) -> set[int]:
        dup_count = len(indices) - 1
        if action == "reject":
            console.info(f"  Rejected {dup_count} duplicate(s)")
            return set(indices[1:])  # Keep first, reject rest
        else:
            console.info("  Keeping all")
            return set()

    manager = DecisionManager(DUPLICATE_CONFIG)

    reject_indices = manager.process_issues(
        duplicates,
        display_issue=display_issue,
        apply_action=apply_action,
        summary_message=f"Found {total_dups} duplicate track(s) across {len(duplicates)} unique tracks",
    )

    return lines, reject_indices


def extract_year_from_date(date_str: str) -> str | None:
    """Extract year from a date string by normalizing it first.

    Uses format_date() to normalize to DD-MM-YYYY, then extracts the year.
    This handles all supported date formats automatically.

    Args:
        date_str: Date string in any supported format.

    Returns:
        Year as string (YYYY), or None if cannot be parsed.
    """
    normalized = format_date(date_str.strip())

    # format_date returns DD-MM-YYYY format
    # Check if it looks like a valid normalized date
    if len(normalized) == 10 and normalized[2] == "-" and normalized[5] == "-":
        year = normalized[6:10]
        # Verify it's actually a year (4 digits)
        if year.isdigit():
            return year

    return None


def check_multiple_years(
    lines: list[str],
    station: Station,
    date_idx: int = 0,
) -> list[str]:
    """Check for dates from multiple years and prompt user if found.

    Args:
        lines: List of input lines.
        station: Station configuration.
        date_idx: Field index for date (default 0).

    Returns:
        Lines, possibly filtered to a specific year.
    """
    sep = station.input_separator
    year_counts: Counter[str] = Counter()
    lines_by_year: dict[str, list[int]] = {}

    for i, line in enumerate(lines):
        if not line.strip():
            continue

        fields = line.split(sep)
        if len(fields) <= date_idx:
            continue

        year = extract_year_from_date(fields[date_idx])
        if year:
            year_counts[year] += 1
            if year not in lines_by_year:
                lines_by_year[year] = []
            lines_by_year[year].append(i)

    # If only one year or no years found, return unchanged
    if len(year_counts) <= 1:
        return lines

    # Multiple years found - warn and prompt
    total_lines = sum(year_counts.values())
    sorted_years = sorted(year_counts.keys())

    console.warning(f"Found dates from {len(year_counts)} different years:")
    for year in sorted_years:
        count = year_counts[year]
        pct = (count / total_lines) * 100
        console.info(f"  {year}: {count:,} lines ({pct:.1f}%)")
    print()

    # Build options dynamically based on years found
    options = [
        Option("a", ["all", "keep"], "All years", "all", is_default=True),
    ]
    for i, year in enumerate(sorted_years):
        # Use number keys for year selection (1, 2, 3, ...)
        key = str(i + 1)
        options.append(Option(key, [year], f"{key}={year}", year))

    config = DecisionConfig(
        name="multiple_years",
        options=options,
        prompt_prefix="  Keep: ",
    )

    # Build custom prompt since years don't fit the [X]label pattern
    prompt_parts = ["[A]ll"]
    for i, year in enumerate(sorted_years):
        prompt_parts.append(f"[{i + 1}] {year}")
    prompt_text = "  Keep: " + " / ".join(prompt_parts) + ": "

    # Prompt user
    while True:
        response = input(prompt_text).strip().lower()
        action = config.parse_response(response)

        if action is None:
            valid_keys = ", ".join(o.key.upper() for o in options)
            console.error(f"  Invalid choice. Enter one of: {valid_keys}")
            continue
        break

    if action == "all":
        console.info("  Keeping all years")
        return lines

    # Filter to selected year
    selected_year = action
    keep_indices = set(lines_by_year.get(selected_year, []))
    filtered_lines = [line for i, line in enumerate(lines) if i in keep_indices]

    removed_count = len(lines) - len(filtered_lines)
    console.success(f"  Filtered to {selected_year}: kept {len(filtered_lines):,} lines, removed {removed_count:,}")

    return filtered_lines


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
        parts[0] = format_date(parts[0])

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
        fields[0] = format_date(fields[0])

    # Format time field (index 1) if it's 6 digits
    if len(fields) > 1 and len(fields[1]) == 6:
        fields[1] = format_time(fields[1])

    # Fix overflow playing time (index 2)
    if len(fields) > 2:
        fields[2] = format_duration(fields[2])

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
    save_reject_file: bool = True,
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
        save_reject_file: Whether to save rejected lines to a file.
    """
    # Track counts
    rejected = 0
    processed = 0
    stopword_counts: Counter[str] = Counter()

    # Prepare file paths
    has_additional = bool(additional_filter)
    additional_path = make_additional_filename(output_file) if has_additional else None
    reject_path = generate_rejection_filename(station) if save_reject_file else None

    # Pre-lowercase additional filter for fast comparison
    additional_filter_lower = additional_filter.lower() if has_additional else ""

    # Ensure rejection directory exists
    if reject_path:
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

            reject_f: IO[str] | None = None
            if reject_path:
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
            if reject_f:
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
                    if reject_f:
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
                    if reject_f:
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

    if reject_path:
        clean_empty_file(reject_path)

    # Check if output file is empty (only header) and run delete_columns
    if output_file.exists():
        output_content = output_file.read_text(encoding="utf-8", errors="replace")
        if output_content.count("\n") > 1:
            run_delete_columns(output_file)
        else:
            clean_empty_file(output_file)

    # Also run delete_columns on rejection file
    if reject_path and reject_path.exists():
        reject_content = reject_path.read_text(encoding="utf-8", errors="replace")
        if reject_content.count("\n") > 1:
            run_delete_columns(reject_path)

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
