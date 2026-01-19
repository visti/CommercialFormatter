#!/usr/bin/env python3
import sys
import os
import csv
import chardet

def detect_encoding(file_path):
    """Detect the encoding of the file using chardet."""
    with open(file_path, 'rb') as f:
        raw_data = f.read()
    result = chardet.detect(raw_data)
    encoding = result.get('encoding', 'utf-8')
    print(f"Detected encoding: {encoding}")
    sys.stdout.flush()
    return encoding

def remove_delete_columns_and_empty_rows(file_path):
    """Remove columns named 'DELETE' and drop rows where 'Main Artist' or 'Track Title' is empty,
    printing any deleted rows to the console."""
    detected_encoding = detect_encoding(file_path)
    
    with open(file_path, 'r', newline='', encoding=detected_encoding, errors='replace') as infile:
        reader = csv.reader(infile, delimiter=';')
        rows = list(reader)
    
    if not rows:
        print(f"No data found in file: {file_path}")
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
    for row in cleaned_rows[1:]:
        # Skip short rows
        if len(row) <= max(main_artist_idx or 0, track_title_idx or 0):
            print(f"Skipping malformed row (too few columns): {row}")
            sys.stdout.flush()
            continue

        if main_artist_idx is not None and track_title_idx is not None:
            artist_val = row[main_artist_idx].strip()
            title_val = row[track_title_idx].strip()
            if artist_val == "" or title_val == "":
                print(f"Deleted row: {row}")
                sys.stdout.flush()
                continue  # skip this row
        final_rows.append(row)

    with open(file_path, 'w', newline='', encoding='utf-8') as outfile:
        writer = csv.writer(outfile, delimiter=';')
        writer.writerows(final_rows)
    
    print(f"Finished cleaning. Updated file: {file_path}")
    sys.stdout.flush()

def main():
    print("Deleting extraneous columns and dropping empty rows.")
    sys.stdout.flush()
    
    if len(sys.argv) < 2:
        print("Usage: python remove_delete_columns.py <filename>")
        sys.exit(1)
    
    file_path = " ".join(sys.argv[1:]).strip()
    
    if os.path.isfile(file_path):
        base, ext = os.path.splitext(file_path)
        if ext.lower() == ".csv":
            try:
                remove_delete_columns_and_empty_rows(file_path)
            except Exception as e:
                print(f"Error processing '{file_path}': {e}")
        else:
            print(f"Skipping non-CSV file: '{file_path}'")
    else:
        print(f"File not found: '{file_path}'")
    
    sys.stdout.flush()
    sys.exit(0)

if __name__ == "__main__":
    main()
