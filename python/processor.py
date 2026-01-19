"""Core file processing logic for commercial formatter."""

import subprocess
import sys
from datetime import datetime
from pathlib import Path

from config import ADDITIONAL_POSTFIX, DELETE_COLS_SCRIPT, REJECTDIR
from stations import Station


def get_files(station: Station, exclude_filename: str = "") -> list[Path]:
    """Find files in current directory matching station's extensions."""
    cwd = Path.cwd()
    files = []

    for ext in station.ext:
        for f in cwd.iterdir():
            if f.is_file() and f.name != exclude_filename:
                if ext.lower() in f.name.lower():
                    files.append(f)

    if not files:
        print("No eligible files found.")
        sys.exit(1)

    print("---------------------")
    print(f"Found Files: {';'.join(f.name for f in files)}")
    print("---------------------\n")

    return files


def read_files(files: list[Path], station: Station) -> str:
    """Read and concatenate all input files, applying station-specific transformations."""
    joined_content = []

    for file_path in files:
        try:
            content = file_path.read_text(encoding="utf-8-sig", errors="replace")
        except Exception as e:
            print(f"Could not read file {file_path}: {e}")
            continue

        lines = content.splitlines()

        # Skip header line if station has headlines
        if station.has_headlines and lines:
            lines = lines[1:]

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
    """Format a 6-digit date string (YYMMDD) to DD-MM-YYYY."""
    date_str = date_str.strip()
    if len(date_str) == 6 and date_str.isdigit():
        yy = date_str[0:2]
        mm = date_str[2:4]
        dd = date_str[4:6]
        year = f"20{yy}" if int(yy) < 50 else f"19{yy}"
        return f"{dd}-{mm}-{year}"
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
            print(f"Removed empty file: {file_path.name}")


def run_delete_columns(output_path: Path):
    """Run the delete_columns.py script on the output file."""
    try:
        subprocess.run(
            [sys.executable, str(DELETE_COLS_SCRIPT), str(output_path)],
            check=True
        )
    except subprocess.CalledProcessError as e:
        print(f"Warning: delete_columns.py failed: {e}")
    except FileNotFoundError:
        print(f"Warning: Could not find {DELETE_COLS_SCRIPT}")


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

    # Open all output files
    with (
        open(output_file, "w", encoding="utf-8") as out_f,
        open(reject_path, "w", encoding="utf-8") as reject_f,
    ):
        additional_f = None
        if has_additional:
            additional_f = open(additional_path, "w", encoding="utf-8")
            print(f"Additional file opened for write: {additional_path}")
            write_headlines(additional_f, station)

        try:
            # Write headlines
            write_headlines(out_f, station)
            write_headlines(reject_f, station)

            # Process each line
            for line in content.splitlines():
                if not line.strip():
                    continue

                # Check stopwords using pre-compiled regex
                if use_stopwords and station.matches_stopword(line):
                    reject_line = process_line(line, station)
                    reject_f.write(reject_line + "\n")
                    rejected += 1
                    continue

                # Process the line
                output_line = process_line(line, station)
                processed += 1

                # Check for additional routing
                if has_additional and additional_filter_lower in line.lower():
                    additional_f.write(output_line + "\n")
                else:
                    out_f.write(output_line + "\n")

        finally:
            if additional_f:
                additional_f.close()

    # Print summary
    print(f"Processed Lines: {processed}")
    print(f"Rejected Lines: {rejected}")

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
