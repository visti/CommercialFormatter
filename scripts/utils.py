#!/usr/bin/env python3
"""Shared utilities for lib scripts."""

import os
import subprocess
import sys
from pathlib import Path


def get_file_path_from_args() -> str:
    """Get file path from command line arguments, handling spaces in filenames."""
    if len(sys.argv) < 2:
        return ""
    return " ".join(sys.argv[1:]).strip()


def validate_csv_file(file_path: str) -> bool:
    """Validate that file exists and is a CSV file."""
    if not os.path.isfile(file_path):
        print(f"File not found: '{file_path}'")
        return False

    _, ext = os.path.splitext(file_path)
    if ext.lower() != ".csv":
        print(f"Skipping non-CSV file: '{file_path}'")
        return False

    return True


def run_python_script(script_path: Path, *args: str) -> bool:
    """Run a Python script with error handling.

    Returns True if successful, False otherwise.
    """
    try:
        subprocess.run(
            [sys.executable, str(script_path)] + list(args),
            check=True
        )
        return True
    except subprocess.CalledProcessError as e:
        print(f"Warning: {script_path.name} failed: {e}")
        return False
    except FileNotFoundError:
        print(f"Warning: Could not find {script_path}")
        return False
