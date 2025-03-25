#!/usr/bin/env python3
import sys
import os
import csv

def remove_delete_columns(input_file, output_file):
    """Remove columns named 'DELETE' from the input CSV and write to the output file."""
    with open(input_file, 'r', newline='', encoding='utf-8') as infile:
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
    
    # Write the cleaned rows to the output file
    with open(output_file, 'w', newline='', encoding='utf-8') as outfile:
        writer = csv.writer(outfile, delimiter=';')
        writer.writerows(cleaned_rows)

    print(f"Columns named 'DELETE' have been removed. Cleaned file saved as {output_file}")

def main():
    if len(sys.argv) < 2:
        print("Usage: python remove_delete_columns.py file1.csv file2.csv ...")
        sys.exit(1)
    
    # Process each CSV file provided in the command-line arguments
    for file_path in sys.argv[1:]:
        if os.path.isfile(file_path):
            try:
                # Construct the output file name by adding a "_cleaned" suffix before the file extension
                base, ext = os.path.splitext(file_path)
                if ext.lower() == ".csv":
                    output_file = f"{base}_cleaned.csv"
                    remove_delete_columns(file_path, output_file)
                else:
                    print(f"Skipping non-CSV file: '{file_path}'")
            except Exception as e:
                print(f"Error processing '{file_path}': {e}")
        else:
            print(f"File not found: '{file_path}'")

if __name__ == "__main__":
    main()

