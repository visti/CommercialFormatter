#!/usr/bin/env python3
import sys
import csv
import os
import chardet

from utils import get_file_path_from_args, validate_csv_file

# Simple ANSI colors
USE_COLORS = hasattr(sys.stdout, "isatty") and sys.stdout.isatty() and not os.environ.get("NO_COLOR")
DIM = "\033[2m" if USE_COLORS else ""
CYAN = "\033[96m" if USE_COLORS else ""
YELLOW = "\033[93m" if USE_COLORS else ""
GREEN = "\033[92m" if USE_COLORS else ""
RESET = "\033[0m" if USE_COLORS else ""


def detect_encoding(file_path):
    """Detect the encoding of the file using chardet."""
    with open(file_path, 'rb') as f:
        raw_data = f.read()
    result = chardet.detect(raw_data)
    encoding = result.get('encoding', 'utf-8')
    return encoding

def remove_delete_columns_and_empty_rows(file_path):
    """Remove columns named 'DELETE' and drop rows where 'Main Artist' or 'Track Title' is empty."""
    detected_encoding = detect_encoding(file_path)

    with open(file_path, 'r', newline='', encoding=detected_encoding, errors='replace') as infile:
        reader = csv.reader(infile, delimiter=';')
        rows = list(reader)

    if not rows:
        print(f"{YELLOW}Warning:{RESET} No data found in {os.path.basename(file_path)}")
        sys.stdout.flush()
        return

    original_header = rows[0]
    delete_columns = [i for i, column in enumerate(original_header) if column == "DELETE"]

    cleaned_rows = []
    for row in rows:
        # Handle short rows safely
        padded_row = row + [""] * (len(original_header) - len(row))
        cleaned_row = [value for i, value in enumerate(padded_row) if i not in delete_columns]
        cleaned_rows.append(cleaned_row)

    cleaned_header = cleaned_rows[0]
    try:
        main_artist_idx = cleaned_header.index("Main Artist")
    except ValueError:
        main_artist_idx = None

    try:
        track_title_idx = cleaned_header.index("Track Title")
    except ValueError:
        track_title_idx = None

    final_rows = [cleaned_header]
    deleted_empty = 0
    deleted_malformed = 0

    for row in cleaned_rows[1:]:
        # Skip short rows
        if len(row) <= max(main_artist_idx or 0, track_title_idx or 0):
            deleted_malformed += 1
            continue

        if main_artist_idx is not None and track_title_idx is not None:
            artist_val = row[main_artist_idx].strip()
            title_val = row[track_title_idx].strip()
            if artist_val == "" or title_val == "":
                deleted_empty += 1
                continue
        final_rows.append(row)

    with open(file_path, 'w', newline='', encoding='utf-8') as outfile:
        writer = csv.writer(outfile, delimiter=';')
        writer.writerows(final_rows)

    # Print summary
    total_deleted = deleted_empty + deleted_malformed
    if total_deleted > 0:
        parts = []
        if deleted_empty > 0:
            parts.append(f"{deleted_empty} empty artist/title")
        if deleted_malformed > 0:
            parts.append(f"{deleted_malformed} malformed")
        print(f"{YELLOW}Cleanup:{RESET} Removed {total_deleted} row(s) ({', '.join(parts)})")

    print(f"{GREEN}Done:{RESET} {len(delete_columns)} columns removed → {CYAN}{os.path.basename(file_path)}{RESET}")
    sys.stdout.flush()

def main():
    print(f"{DIM}Cleaning up columns and empty rows...{RESET}")
    sys.stdout.flush()

    file_path = get_file_path_from_args()
    if not file_path:
        print("Usage: python remove_delete_columns.py <filename>")
        sys.exit(1)

    if validate_csv_file(file_path):
        try:
            remove_delete_columns_and_empty_rows(file_path)
        except Exception as e:
            print(f"Error processing '{file_path}': {e}")
    
    sys.stdout.flush()
    sys.exit(0)

if __name__ == "__main__":
    main()
