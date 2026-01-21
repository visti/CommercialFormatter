"""Core file processing logic for commercial formatter."""

import subprocess
import sys
from contextlib import ExitStack
from datetime import datetime
from pathlib import Path
from typing import TYPE_CHECKING

import chardet
import output as console
from config import ADDITIONAL_POSTFIX, DELETE_COLS_SCRIPT, REJECTDIR
from stations import Station


def detect_encoding(file_path: Path) -> str:
    """Detect file encoding using chardet."""
    with open(file_path, "rb") as f:
        raw_data = f.read()
    result = chardet.detect(raw_data)
    encoding = result.get("encoding", "utf-8")
    # Handle common encoding aliases
    if encoding and encoding.lower() in ("ascii", "iso-8859-1", "latin-1", "latin1"):
        # These are often misdetected; try cp1252 which is a superset
        encoding = "cp1252"
    return encoding or "utf-8"

if TYPE_CHECKING:
    from output import ProcessingStats


def get_files(station: Station, exclude_filename: str = "") -> list[Path]:
    """Find files in current directory matching station's extensions."""
    cwd = Path.cwd()
    files = []

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


def read_files(files: list[Path], station: Station, stats: "ProcessingStats" = None) -> str:
    """Read and concatenate all input files, applying station-specific transformations."""
    joined_content = []
    total_files = len(files)

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
            continue

        lines = content.splitlines()

        # Skip header lines based on station config
        if station.skip_lines > 0 and lines:
            lines = lines[station.skip_lines:]

        # Globus-specific: prepend filename and replace " - " with ";"
        if station.name == "Globus":
            base_filename = file_path.stem
            processed_lines = []
            for line in lines:
                if line.strip():
                    line = line.replace(" - ", ";")
                    line = f"{base_filename}{line}"
                    processed_lines.append(line)
            lines = processed_lines

        # Skive-specific: split "DD-MM-YYYY HH-HH" into "DD-MM-YYYY;HH:00:00"
        if station.name == "Skive":
            processed_lines = []
            for line in lines:
                if line.strip():
                    # Split first field (Tidspunkt) which contains "01-12-2025 00-01"
                    parts = line.split(";", 1)
                    if len(parts) >= 1 and " " in parts[0]:
                        date_time = parts[0].split(" ", 1)
                        if len(date_time) == 2:
                            date_part = date_time[0]
                            # Time is "00-01" format (hour range), take first hour
                            time_range = date_time[1]
                            hour = time_range.split("-")[0] if "-" in time_range else time_range
                            time_part = f"{hour}:00:00"
                            # Reconstruct: date;time;rest_of_fields
                            rest = parts[1] if len(parts) > 1 else ""
                            line = f"{date_part};{time_part};{rest}"
                    processed_lines.append(line)
            lines = processed_lines

        joined_content.extend(lines)

    return "\n".join(joined_content)


def make_additional_filename(base_path: Path) -> Path:
    """Create additional output filename with postfix."""
    stem = base_path.stem
    ext = base_path.suffix
    return base_path.parent / f"{stem}{ADDITIONAL_POSTFIX}{ext}"


def generate_rejection_filename(station: Station) -> Path:
    """Generate the rejection file path with current date."""
    now = datetime.now()
    date_str = f"{now.year}-{now.month}-{now.day}"
    filename = f"{date_str}-reject-{station.name}.csv"
    return REJECTDIR / filename


def format_time_field(time_str: str) -> str:
    """Format a 6-digit time string (HHMMSS) to HH:MM:SS."""
    if len(time_str) == 6:
        return f"{time_str[0:2]}:{time_str[2:4]}:{time_str[4:6]}"
    return time_str


def format_date_field(date_str: str) -> str:
    """Normalize date to DD-MM-YYYY format.

    Handles:
    - YYMMDD (6 digits) -> DD-MM-YYYY
    - YYYY-MM-DD -> DD-MM-YYYY
    - DD-MM-YYYY -> unchanged
    """
    date_str = date_str.strip()

    # Handle 6-digit format: YYMMDD
    if len(date_str) == 6 and date_str.isdigit():
        yy = date_str[0:2]
        mm = date_str[2:4]
        dd = date_str[4:6]
        year = f"20{yy}" if int(yy) < 50 else f"19{yy}"
        return f"{dd}-{mm}-{year}"

    # Handle YYYY-MM-DD format (convert to DD-MM-YYYY)
    if len(date_str) == 10 and date_str[4] == "-" and date_str[7] == "-":
        parts = date_str.split("-")
        if len(parts) == 3 and len(parts[0]) == 4:
            yyyy, mm, dd = parts
            return f"{dd}-{mm}-{yyyy}"

    # Already DD-MM-YYYY or unknown format, return as-is
    return date_str


def process_positional_line(line: str, sorted_positions: list[int], separator: str) -> str:
    """Extract fields from a fixed-width line using pre-sorted positional indices."""
    parts = []
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


def process_csv_line(line: str, separator: str = ";") -> str:
    """Process a CSV/semicolon-delimited line, formatting date and time fields."""
    fields = line.split(";")

    # Format date field (index 0)
    if fields:
        fields[0] = format_date_field(fields[0])

    # Format time field (index 1) if it's 6 digits
    if len(fields) > 1 and len(fields[1]) == 6:
        fields[1] = format_time_field(fields[1])

    return separator.join(fields)


def write_headlines(file_handle, station: Station):
    """Write headline row to file."""
    headlines_line = station.separator.join(station.headlines)
    file_handle.write(headlines_line + "\n")


def clean_empty_file(file_path: Path):
    """Remove file if it's empty or contains only a header line."""
    if file_path.exists():
        content = file_path.read_text(encoding="utf-8", errors="replace")
        line_count = content.count("\n")
        if line_count <= 1:
            file_path.unlink()
            console.info(console.dim(f"Removed empty file: {file_path.name}"))


def run_delete_columns(output_path: Path):
    """Run the delete_columns.py script on the output file."""
    try:
        subprocess.run(
            [sys.executable, str(DELETE_COLS_SCRIPT), str(output_path)],
            check=True
        )
    except subprocess.CalledProcessError as e:
        console.warning(f"delete_columns.py failed: {e}")
    except FileNotFoundError:
        console.warning(f"Could not find {DELETE_COLS_SCRIPT}")


def process_line(line: str, station: Station) -> str:
    """Process a single line according to station format."""
    if station.positional:
        return process_positional_line(line, station.sorted_positions, station.separator)
    else:
        return process_csv_line(line, station.separator)


def process_files(
    content: str,
    station: Station,
    output_file: Path,
    additional_filter: str = "",
    use_stopwords: bool = True,
    stats: "ProcessingStats" = None,
):
    """Main processing pipeline for station data."""
    # Track counts
    rejected = 0
    processed = 0

    # Prepare file paths
    has_additional = bool(additional_filter)
    additional_path = make_additional_filename(output_file) if has_additional else None
    reject_path = generate_rejection_filename(station)

    # Pre-lowercase additional filter for fast comparison
    additional_filter_lower = additional_filter.lower() if has_additional else ""

    # Ensure rejection directory exists
    reject_path.parent.mkdir(parents=True, exist_ok=True)

    console.info("Processing lines...")

    # Open all output files using ExitStack for proper context management
    with ExitStack() as stack:
        out_f = stack.enter_context(open(output_file, "w", encoding="utf-8"))
        reject_f = stack.enter_context(open(reject_path, "w", encoding="utf-8"))

        additional_f = None
        if has_additional:
            additional_f = stack.enter_context(open(additional_path, "w", encoding="utf-8"))
            console.info(f"Additional filter: {console.cyan(additional_filter)}")
            if stats:
                stats.additional_file = additional_path.name
            write_headlines(additional_f, station)

        # Write headlines
        write_headlines(out_f, station)
        write_headlines(reject_f, station)

        # Process each line
        for line in content.splitlines():
            if not line.strip():
                continue

            # Pre-compute lowercase once for efficiency
            line_lower = line.lower()

            # Check stopwords using pre-compiled regex
            if use_stopwords and station.matches_stopword_lower(line_lower):
                reject_line = process_line(line, station)
                reject_f.write(reject_line + "\n")
                rejected += 1
                continue

            # Process the line
            output_line = process_line(line, station)
            processed += 1

            # Check for additional routing
            if has_additional and additional_filter_lower in line_lower:
                additional_f.write(output_line + "\n")
            else:
                out_f.write(output_line + "\n")

    # Update stats
    if stats:
        stats.lines_processed = processed
        stats.lines_rejected = rejected

    # Cleanup phase
    if has_additional:
        clean_empty_file(additional_path)

    clean_empty_file(reject_path)

    # Check if output file is empty (only header)
    if output_file.exists():
        content = output_file.read_text(encoding="utf-8", errors="replace")
        if content.count("\n") > 1:
            # Only run delete_columns if there's actual data
            run_delete_columns(output_file)
        else:
            clean_empty_file(output_file)
