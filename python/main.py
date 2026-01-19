#!/usr/bin/env python3
"""Commercial Formatter - Main entry point."""

import argparse
import subprocess
import sys
from pathlib import Path

from config import CONVERT_SCRIPT, NOTVALIDSTATION
from processor import get_files, process_files, read_files
from stations import get_station, list_aliases, list_stations


def run_convert_script():
    """Run the XLSX to CSV conversion script."""
    try:
        subprocess.run(
            [sys.executable, str(CONVERT_SCRIPT), "*xls*"],
            check=True,
            cwd=Path.cwd()
        )
    except subprocess.CalledProcessError as e:
        print(f"Warning: convert.py failed: {e}")
    except FileNotFoundError:
        print(f"Warning: Could not find {CONVERT_SCRIPT}")


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

    args = parser.parse_args()

    # Handle --list-stations
    if args.list_stations:
        print_stations_and_aliases()
        sys.exit(0)

    # Require station if not listing
    if not args.station:
        parser.print_help()
        print()
        print_stations_and_aliases()
        sys.exit(1)

    # Get station configuration
    station = get_station(args.station)
    if station is None:
        print(NOTVALIDSTATION)
        print(f"'{args.station}' is not a valid station or alias.")
        print()
        print_stations_and_aliases()
        sys.exit(1)

    print("---------------------")
    print(f"Station: {station.name}")
    print("---------------------\n")

    # Run conversion script if station requires it
    if station.convert:
        run_convert_script()

    # Ask for output filename
    output_filename = input("Set output filename: ").strip()
    if not output_filename:
        print("ERROR: Output filename cannot be empty!")
        sys.exit(1)

    output_path = Path.cwd() / output_filename

    # Delete existing output file if it exists
    if output_path.exists():
        output_path.unlink()
        print(f"Initialized {output_filename}")

    if args.no_stopwords:
        print("Stopword filtering DISABLED (--no-stopwords)")

    # Find and process files
    files = get_files(station, exclude_filename=output_filename)
    content = read_files(files, station)

    process_files(
        content=content,
        station=station,
        output_file=output_path,
        additional_filter=args.additional,
        use_stopwords=not args.no_stopwords,
    )


if __name__ == "__main__":
    main()
