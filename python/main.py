#!/usr/bin/env python3
"""Commercial Formatter - Main entry point."""

import argparse
import subprocess
import sys
import tomllib
from pathlib import Path

import output as console
from config import CONVERT_SCRIPT, NOTVALIDSTATION
from processor import get_files, make_additional_filename, process_files, read_files
from stations import get_station, list_aliases, list_stations

# Folder mapping config
FOLDERS_CONFIG = Path(__file__).parent / "config" / "folders.toml"


def load_folder_mapping() -> dict[str, str]:
    """Load folder to station mapping from config."""
    if not FOLDERS_CONFIG.exists():
        return {}

    with open(FOLDERS_CONFIG, "rb") as f:
        data = tomllib.load(f)

    return data.get("folders", {})


def detect_station_from_path() -> str | None:
    """Detect station alias from current working directory path.

    Checks if any configured folder name appears in the current path.
    """
    folder_mapping = load_folder_mapping()
    if not folder_mapping:
        return None

    cwd_str = str(Path.cwd()).lower()

    for folder_name, station_alias in folder_mapping.items():
        # Check if folder name appears in the path (case-insensitive)
        if folder_name.lower() in cwd_str:
            return station_alias

    return None


def run_convert_script():
    """Run the XLSX to CSV conversion script."""
    # Find xlsx files manually since Windows doesn't expand globs
    # Use case-insensitive matching
    xlsx_files = [f for f in Path.cwd().iterdir() if f.is_file() and "xls" in f.name.lower()]
    if not xlsx_files:
        console.warning("No xlsx files found to convert")
        return

    console.info(f"Converting {len(xlsx_files)} xlsx file(s)...")
    try:
        subprocess.run(
            [sys.executable, str(CONVERT_SCRIPT)] + [str(f) for f in xlsx_files],
            check=True,
            cwd=Path.cwd()
        )
    except subprocess.CalledProcessError as e:
        console.warning(f"convert.py failed: {e}")
    except FileNotFoundError:
        console.warning(f"Could not find {CONVERT_SCRIPT}")


def suggest_output_filename() -> str:
    """Suggest output filename based on folder structure.

    Expects structure like: /Silkeborg/2025/q4/ -> 2025_q4_silkeborg.csv
    """
    cwd = Path.cwd()
    quarter = cwd.name.lower()           # e.g., "q4"
    year = cwd.parent.name               # e.g., "2025"
    station_folder = cwd.parent.parent.name.lower()  # e.g., "silkeborg"

    # Validate that we have a reasonable structure
    if quarter.startswith("q") and year.isdigit() and station_folder:
        return f"{year}_{quarter}_{station_folder}.csv"
    return ""


def check_file_accessible(file_path: Path) -> bool:
    """Check if a file can be written to (not locked by another process).

    Args:
        file_path: Path to the file to check

    Returns:
        True if file is accessible or doesn't exist, False if locked
    """
    if not file_path.exists():
        return True

    try:
        # Try to open the file in append mode to check if it's locked
        with open(file_path, "a"):
            pass
        return True
    except (PermissionError, OSError):
        return False


def ensure_file_accessible(file_path: Path) -> bool:
    """Ensure a file is accessible, prompting user to close it if needed.

    Args:
        file_path: Path to the file to check

    Returns:
        True if file is accessible, False if user chose to quit
    """
    while not check_file_accessible(file_path):
        console.error(f"Cannot access '{file_path.name}' - file is in use.")
        response = input("Close the file and press Enter to retry, or X to quit: ").strip().lower()
        if response == "x":
            return False
    return True


def print_stations_and_aliases():
    """Print all stations and their aliases."""
    aliases = list_aliases()
    print("Available stations and aliases:")
    print("-" * 40)
    for station_key in list_stations():
        station_aliases = aliases.get(station_key, [])
        if station_aliases:
            print(f"  {station_key}: {', '.join(station_aliases)}")
        else:
            print(f"  {station_key}")
    print()


def main():
    parser = argparse.ArgumentParser(
        prog="komm_fmt",
        description="Commercial Formatter — Process broadcast metadata files",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
    komm_fmt                            (auto-detect station from folder path)
    komm_fmt Globus
    komm_fmt Globus --no-stopwords
    komm_fmt Radio4 --additional=Boulevard
    komm_fmt 100fm --additional=Hits    (uses Bauer via alias)
    komm_fmt GoFM                       (uses Jyskfynske via alias)
    komm_fmt --list-stations            (show all stations and aliases)
        """,
    )

    parser.add_argument(
        "station",
        nargs="?",
        help="Name of station or alias (use --list-stations to see all)",
    )
    parser.add_argument(
        "--additional",
        metavar="FILTER",
        default="",
        help="Routes matching lines to an additional CSV",
    )
    parser.add_argument(
        "--additional-postfix",
        metavar="POSTFIX",
        default="_additional",
        help="Postfix for additional output file (default: _additional)",
    )
    parser.add_argument(
        "--no-stopwords",
        action="store_true",
        help="Disable stopword rejection entirely",
    )
    parser.add_argument(
        "--list-stations",
        action="store_true",
        help="List all available stations and their aliases",
    )
    parser.add_argument(
        "--edit-choices",
        action="store_true",
        help="Open remembered choices file in nvim for editing",
    )

    args = parser.parse_args()

    # Handle --list-stations
    if args.list_stations:
        print_stations_and_aliases()
        sys.exit(0)

    # Handle --edit-choices
    if args.edit_choices:
        choices_file = Path(__file__).parent / "config" / "remembered_choices.toml"
        if not choices_file.exists():
            console.info(f"Creating {choices_file.name}...")
            choices_file.parent.mkdir(parents=True, exist_ok=True)
            choices_file.write_text("# Remembered user choices\n\n[artist_title_fixes]\n\n[long_playing_times]\n")
        subprocess.run(["nvim", str(choices_file)])
        sys.exit(0)

    # Determine station: use argument, or auto-detect from path
    station_name = args.station
    if not station_name:
        station_name = detect_station_from_path()
        if station_name:
            console.info(f"Auto-detected station: {console.cyan(station_name)}")
        else:
            parser.print_help()
            print()
            print_stations_and_aliases()
            sys.exit(1)

    # Get station configuration
    station = get_station(station_name)
    if station is None:
        console.error(NOTVALIDSTATION)
        console.error(f"'{station_name}' is not a valid station or alias.")
        print()
        print_stations_and_aliases()
        sys.exit(1)

    console.header(f"Station: {station.name}")

    # Run conversion script if station requires it
    if station.convert:
        run_convert_script()

    # Ask for output filename with suggestion
    suggested = suggest_output_filename()
    if suggested:
        output_filename = input(f"Set output filename [{suggested}]: ").strip()
        if not output_filename:
            output_filename = suggested
    else:
        output_filename = input("Set output filename: ").strip()
        if not output_filename:
            console.error("Output filename cannot be empty!")
            sys.exit(1)

    output_path = Path.cwd() / output_filename

    # Check if output file is accessible (not locked by another process)
    if not ensure_file_accessible(output_path):
        console.info("Exiting.")
        sys.exit(0)

    # Also check additional output file if --additional is specified
    if args.additional:
        additional_path = make_additional_filename(output_path)
        if not ensure_file_accessible(additional_path):
            console.info("Exiting.")
            sys.exit(0)

    # Delete existing output file if it exists
    if output_path.exists():
        output_path.unlink()
        console.info(f"Overwriting existing {output_filename}")

    if args.no_stopwords:
        console.warning("Stopword filtering DISABLED (--no-stopwords)")

    # Initialize stats tracking
    stats = console.ProcessingStats(output_file=output_filename)

    # Find and process files
    files = get_files(station, exclude_filename=output_filename)
    stats.files_total = len(files)

    content, reject_indices = read_files(files, station, stats)

    process_files(
        content=content,
        station=station,
        output_file=output_path,
        additional_filter=args.additional,
        use_stopwords=not args.no_stopwords,
        stats=stats,
        force_reject_indices=reject_indices,
    )

    # Print summary
    console.print_summary_box(stats)


if __name__ == "__main__":
    main()
