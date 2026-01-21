#!/usr/bin/env python3
import sys
import pandas as pd

from utils import get_file_path_from_args, validate_csv_file


def delete_podcast_only_rows(file_path):
    """
    Remove rows from a CSV file where the column "Podcast only" equals TRUE.
    The check is done case-insensitively on the string representation of the value.
    The file is edited in place.
    """
    # Read the CSV file using semicolon as the delimiter
    df = pd.read_csv(file_path, delimiter=';', encoding='utf-8')
    
    if "Podcast only" in df.columns:
        # Remove rows where "Podcast only" equals "TRUE" (case-insensitive)
        df = df[~(df["Podcast only"].astype(str).str.upper() == "TRUE")]
    
    # Write the processed data back to the same file, preserving the semicolon delimiter
    df.to_csv(file_path, index=False, sep=';', encoding='utf-8')
    print(f"Rows with 'Podcast only' == TRUE have been removed. File updated: {file_path}")

def main():
    file_path = get_file_path_from_args()
    if not file_path:
        print("Usage: python delete_podcast_rows_in_place.py <filename>")
        sys.exit(1)

    if validate_csv_file(file_path):
        try:
            delete_podcast_only_rows(file_path)
        except Exception as e:
            print(f"Error processing '{file_path}': {e}")

if __name__ == "__main__":
    main()

