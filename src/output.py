"""Console output utilities with colors and formatting."""

import os
import sys
import time
from dataclasses import dataclass, field

# ANSI color codes
class Colors:
    """ANSI escape codes for terminal colors."""
    RESET = "\033[0m"
    BOLD = "\033[1m"
    DIM = "\033[2m"

    # Regular colors
    RED = "\033[31m"
    GREEN = "\033[32m"
    YELLOW = "\033[33m"
    BLUE = "\033[34m"
    MAGENTA = "\033[35m"
    CYAN = "\033[36m"
    WHITE = "\033[37m"

    # Bright colors
    BRIGHT_RED = "\033[91m"
    BRIGHT_GREEN = "\033[92m"
    BRIGHT_YELLOW = "\033[93m"
    BRIGHT_BLUE = "\033[94m"
    BRIGHT_CYAN = "\033[96m"


def _supports_color() -> bool:
    """Check if the terminal supports color output."""
    # Disable colors if NO_COLOR env var is set
    if os.environ.get("NO_COLOR"):
        return False
    # Check if stdout is a terminal
    if not hasattr(sys.stdout, "isatty") or not sys.stdout.isatty():
        return False
    # Enable virtual terminal processing on Windows
    if sys.platform == "win32":
        try:
            import ctypes
            kernel32 = ctypes.windll.kernel32
            # Enable ANSI escape code processing
            kernel32.SetConsoleMode(kernel32.GetStdHandle(-11), 7)
        except Exception:
            pass
    return True


# Check color support once at module load
USE_COLORS = _supports_color()


def _color(text: str, color: str) -> str:
    """Apply color to text if colors are supported."""
    if USE_COLORS:
        return f"{color}{text}{Colors.RESET}"
    return text


# Convenience functions for colored output
def red(text: str) -> str:
    return _color(text, Colors.BRIGHT_RED)

def green(text: str) -> str:
    return _color(text, Colors.BRIGHT_GREEN)

def yellow(text: str) -> str:
    return _color(text, Colors.BRIGHT_YELLOW)

def blue(text: str) -> str:
    return _color(text, Colors.BRIGHT_BLUE)

def cyan(text: str) -> str:
    return _color(text, Colors.BRIGHT_CYAN)

def bold(text: str) -> str:
    return _color(text, Colors.BOLD)

def dim(text: str) -> str:
    return _color(text, Colors.DIM)


# Output functions
def info(message: str):
    """Print an info message."""
    print(message)

def success(message: str):
    """Print a success message in green."""
    print(green(message))

def warning(message: str):
    """Print a warning message in yellow."""
    print(yellow(f"Warning: {message}"))

def error(message: str):
    """Print an error message in red."""
    print(red(f"Error: {message}"))

def progress(current: int, total: int, message: str):
    """Print a progress message."""
    print(dim(f"[{current}/{total}]") + f" {message}")


def header(title: str):
    """Print a section header."""
    line = "-" * 21
    print(line)
    print(bold(title))
    print(line)


def format_duration(seconds: float) -> str:
    """Format a duration in seconds to a human-readable string."""
    if seconds < 1:
        return f"{seconds * 1000:.0f}ms"
    elif seconds < 60:
        return f"{seconds:.1f}s"
    else:
        minutes = int(seconds // 60)
        secs = seconds % 60
        return f"{minutes}m {secs:.1f}s"


def format_number(n: int) -> str:
    """Format a number with thousand separators."""
    return f"{n:,}"


@dataclass
class ProcessingStats:
    """Track processing statistics."""
    start_time: float = field(default_factory=time.time)
    files_total: int = 0
    files_processed: int = 0
    lines_processed: int = 0
    lines_rejected: int = 0
    output_file: str = ""
    additional_file: str = ""

    def elapsed(self) -> float:
        """Return elapsed time in seconds."""
        return time.time() - self.start_time


def print_summary_box(stats: ProcessingStats):
    """Print a summary box with processing statistics."""
    duration = format_duration(stats.elapsed())

    # Box characters
    if USE_COLORS:
        tl, tr, bl, br = "┌", "┐", "└", "┘"
        h, v, m = "─", "│", "├"
    else:
        tl, tr, bl, br = "+", "+", "+", "+"
        h, v, m = "-", "|", "+"

    width = 35
    inner = width - 2

    def pad(text: str, raw_len: int = None) -> str:
        """Pad text to fill the box width."""
        if raw_len is None:
            raw_len = len(text)
        padding = inner - raw_len
        return text + " " * padding

    print()
    print(f"{tl}{h * inner}{tr}")

    # Title
    title = "Processing Complete"
    title_colored = green(title) if USE_COLORS else title
    print(f"{v} {pad(title_colored, len(title))} {v}")

    print(f"{m}{h * inner}{m}")

    # Stats
    lines = [
        ("Files processed:", format_number(stats.files_processed)),
        ("Lines processed:", format_number(stats.lines_processed)),
        ("Lines rejected:", format_number(stats.lines_rejected)),
        ("Time elapsed:", duration),
    ]

    for label, value in lines:
        line = f"{label:18} {value}"
        print(f"{v} {pad(line)} {v}")

    print(f"{m}{h * inner}{m}")

    # Output file
    output_label = "Output:"
    output_colored = cyan(stats.output_file) if USE_COLORS else stats.output_file
    print(f"{v} {output_label:8} {pad(output_colored, len(stats.output_file) + 9)} {v}")

    if stats.additional_file:
        add_colored = cyan(stats.additional_file) if USE_COLORS else stats.additional_file
        print(f"{v} {'Extra:':8} {pad(add_colored, len(stats.additional_file) + 9)} {v}")

    print(f"{bl}{h * inner}{br}")
