#!/usr/bin/env python3
import sys
import os
import pandas as pd

def delete_podcast_only_rows(file_path):
    """
    Remove rows from a CSV file where the column "Podcast only" equals TRUE.
    The check is done case-insensitively on the string representation of the value.
    The file is edited in place.
    """
    # Read the CSV file using semicolon as the delimiter
    df = pd.read_csv(file_path, delimiter=';')
    
    if "Podcast only" in df.columns:
        # Remove rows where "Podcast only" equals "TRUE" (case-insensitive)
        df = df[~(df["Podcast only"].astype(str).str.upper() == "TRUE")]
    
    # Write the processed data back to the same file, preserving the semicolon delimiter
    df.to_csv(file_path, index=False, sep=';')
    print(f"Rows with 'Podcast only' == TRUE have been removed. File updated: {file_path}")

def main():
    if len(sys.argv) < 2:
        print("Usage: python delete_podcast_rows_in_place.py <filename>")
        sys.exit(1)
    
    # Join all arguments into one string in case the filename contains spaces
    file_path = " ".join(sys.argv[1:]).strip()
    
    if os.path.isfile(file_path):
        base, ext = os.path.splitext(file_path)
        if ext.lower() == '.csv':
            try:
                delete_podcast_only_rows(file_path)
            except Exception as e:
                print(f"Error processing '{file_path}': {e}")
        else:
            print(f"Skipping non-CSV file: '{file_path}'")
    else:
        print(f"File not found: '{file_path}'")

if __name__ == "__main__":
    main()

