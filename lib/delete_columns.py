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

def remove_delete_columns(file_path):
    """Remove columns named 'DELETE' from the CSV file and edit it in place."""
    # Detect the file's encoding
    detected_encoding = detect_encoding(file_path)
    
    # Read the file with the detected encoding (replacing errors if needed)
    with open(file_path, 'r', newline='', encoding=detected_encoding, errors='replace') as infile:
        reader = csv.reader(infile, delimiter=';')
        rows = list(reader)
    
    # Get the header (first row)
    header = rows[0]
    
    # Identify columns with "DELETE" in the header
    delete_columns = [i for i, column in enumerate(header) if column == "DELETE"]
    
    # Remove "DELETE" columns from all rows
    cleaned_rows = []
    for row in rows:
        cleaned_row = [value for i, value in enumerate(row) if i not in delete_columns]
        cleaned_rows.append(cleaned_row)
    
    # Write the cleaned rows back to the same file using UTF-8 encoding
    with open(file_path, 'w', newline='', encoding='utf-8') as outfile:
        writer = csv.writer(outfile, delimiter=';')
        writer.writerows(cleaned_rows)
    
    print(f"Columns named 'DELETE' have been removed. File updated: {file_path}")
    sys.stdout.flush()

def main():
    print("Deleting extraneous columns.")
    sys.stdout.flush()
    
    if len(sys.argv) < 2:
        print("Usage: python remove_delete_columns.py <filename>")
        sys.exit(1)
    
    # Join all arguments into one string in case the filename contains spaces
    file_path = " ".join(sys.argv[1:]).strip()
    
    if os.path.isfile(file_path):
        base, ext = os.path.splitext(file_path)
        if ext.lower() == ".csv":
            try:
                remove_delete_columns(file_path)
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
