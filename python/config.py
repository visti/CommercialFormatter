"""Global configuration constants for commercial formatter."""

from pathlib import Path

# Paths
BASE_DIR = Path(__file__).parent
LIB_DIR = BASE_DIR / "lib"
CONFIG_DIR = BASE_DIR / "config"
CONVERT_SCRIPT = LIB_DIR / "convert.py"
DELETE_COLS_SCRIPT = LIB_DIR / "delete_columns.py"
REJECTDIR = Path("C:/Users/eva/Gramex/Rapporteringer - Documents/Afviste_linjer_kom_land")

# Defaults
ADDITIONAL_POSTFIX = "_additional"

# Error messages
NOTVALIDSTATION = "ERROR: No valid station selected as argument."
