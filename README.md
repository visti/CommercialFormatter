# Commercial Formatter

A tool for processing broadcast metadata files from Danish radio stations. Extracts track information, filters unwanted content (jingles, promos, etc.), and outputs clean CSV files for reporting.

## Implementations

Two implementations are available:
- **Python** (recommended): Easier to configure, install via pip
- **Odin**: Original implementation, compiled binary

---

## Python Implementation

### Installation

```bash
cd python
pip install -e .
```

This installs the `komm_fmt` command globally.

### Usage

```bash
komm_fmt <station> [options]

# Examples
komm_fmt Bauer
komm_fmt GoFM                        # Uses Jyskfynske via alias
komm_fmt Radio4 --additional=Boulevard
komm_fmt Bauer --no-stopwords
komm_fmt --list-stations             # Show all stations and aliases
```

### Options

| Option | Description |
|--------|-------------|
| `--additional=FILTER` | Route lines containing FILTER to a separate CSV |
| `--additional-postfix=TEXT` | Suffix for additional file (default: `_additional`) |
| `--no-stopwords` | Disable stopword filtering |
| `--list-stations` | List all stations and their aliases |

---

## File Structure

```
commercial_formatter/
├── python/                     # Python implementation
│   ├── main.py                 # CLI entry point
│   ├── processor.py            # Core processing logic
│   ├── stations.py             # Station loading from TOML
│   ├── config.py               # Global paths and constants
│   ├── pyproject.toml          # Package configuration
│   ├── config/
│   │   ├── stations.toml       # Station definitions
│   │   └── stopwords.toml      # Rejection wordlists
│   └── lib/
│       ├── convert.py          # XLSX to CSV converter
│       ├── delete_columns.py   # Remove DELETE columns from output
│       └── delete_podcast.py   # Podcast removal utility
│
└── odin/                       # Odin implementation
    ├── main.odin               # Entry point
    ├── stations.odin           # Station definitions
    ├── process_files.odin      # Processing logic
    ├── io.odin                 # File I/O operations
    ├── globals.odin            # Global constants
    ├── helpers.odin            # Utility functions
    ├── build.sh                # Build script
    └── komm_fmt_odin           # Compiled binary
```

---

## Python File Details

### `python/main.py`
**CLI entry point.** Handles argument parsing, station selection, and orchestrates the processing pipeline.

- Parses command-line arguments using `argparse`
- Validates station name or alias
- Prompts for output filename
- Calls processor functions in sequence
- Provides `--list-stations` to show available stations/aliases

### `python/processor.py`
**Core processing logic.** Reads input files, applies filtering, and writes output.

- `get_files()`: Finds files matching station's extensions in current directory
- `read_files()`: Reads and concatenates files, handles BOM, applies Globus-specific transforms
- `process_line()`: Formats a single line (positional extraction or CSV parsing)
- `process_files()`: Main pipeline - stopword filtering, line processing, routing to output files
- `format_date_field()`: Converts YYMMDD to DD-MM-YYYY
- `format_time_field()`: Converts HHMMSS to HH:MM:SS
- `run_delete_columns()`: Calls helper script to remove DELETE columns

### `python/stations.py`
**Station configuration loader.** Reads TOML files and builds Station objects.

- `Station` dataclass: Holds station config (name, extensions, positions, headlines, etc.)
- `_load_stopwords()`: Loads stopwords from `config/stopwords.toml`
- `_load_stations_config()`: Loads stations from `config/stations.toml`
- `_compile_stopword_pattern()`: Compiles all stopwords into single regex for fast matching
- `_build_stations()`: Creates Station objects with pre-compiled stopword matchers
- `get_station()`: Looks up station by name or alias (direct names take precedence)
- `list_stations()`: Returns list of station names
- `list_aliases()`: Returns dict mapping stations to their aliases
- `reload_config()`: Reloads config from disk without restarting

### `python/config.py`
**Global configuration constants.**

- `BASE_DIR`: Python package directory
- `LIB_DIR`: Path to helper scripts
- `CONFIG_DIR`: Path to TOML config files
- `CONVERT_SCRIPT`: Path to XLSX converter
- `DELETE_COLS_SCRIPT`: Path to column deletion script
- `REJECTDIR`: Directory for rejected lines files
- `ADDITIONAL_POSTFIX`: Default suffix for additional output files

### `python/config/stations.toml`
**Station definitions.** Edit this file to add/modify stations without changing code.

Each station section defines:
- `name`: Display name
- `aliases`: Alternative names that resolve to this station
- `extensions`: File extensions to process (e.g., `["txt", "den"]`)
- `positional`: `true` for fixed-width files, `false` for delimited
- `positions`: Column boundaries for positional extraction
- `has_headlines`: `true` if input files have a header row to skip
- `convert`: `true` to run XLSX converter before processing
- `separator`: Output field separator (`;` or `:`)
- `headlines`: Column headers for output file (use `"DELETE"` to mark columns for removal)

### `python/config/stopwords.toml`
**Rejection wordlists.** Lines containing any stopword are rejected.

- `[default]`: Stopwords applied to ALL stations (DJ mixes, megamixes, etc.)
- `[stationname]`: Additional stopwords for specific stations

Edit this file to add/remove stopwords without changing code.

### `python/lib/convert.py`
**XLSX to CSV converter.** Converts Excel files to CSV format before processing.

- Finds `.xlsx` files matching pattern
- Converts to CSV with semicolon delimiter
- Archives original files

### `python/lib/delete_columns.py`
**Column removal utility.** Post-processes output files to remove unwanted columns.

- Reads the output CSV
- Removes columns with header `"DELETE"`
- Removes empty rows
- Detects file encoding automatically

### `python/lib/delete_podcast.py`
**Podcast removal utility.** Filters out podcast-related entries.

---

## Odin File Details

### `odin/main.odin`
Entry point. Parses CLI arguments, handles memory tracking in debug mode, orchestrates processing.

### `odin/stations.odin`
Station definitions and stopword lists. Hardcoded station configurations.

### `odin/process_files.odin`
Core processing logic. Stopword checking, positional extraction, CSV parsing, file output.

### `odin/io.odin`
File I/O operations. Reading files, writing headlines, generating rejection filenames.

### `odin/globals.odin`
Global constants. Paths, separators, error messages.

### `odin/helpers.odin`
Utility functions. Debug printing, file deletion.

### `odin/build.sh`
Build script. Compiles Odin source to executable.

---

## Configuration Examples

### Adding a New Station

Edit `python/config/stations.toml`:

```toml
[newstation]
name = "New Station"
aliases = ["ns", "newst"]
extensions = ["txt"]
positional = true
positions = [100, 80, 60, 40, 20, 10]
has_headlines = false
separator = ";"
headlines = [
    "Date of Broadcasting",
    "Track starting time",
    "Track Title",
    "Main Artist",
]
```

### Adding Stopwords

Edit `python/config/stopwords.toml`:

```toml
[default]
words = [
    # ... existing words ...
    "new stopword",
]

[newstation]
words = [
    "station-specific stopword",
]
```

### Adding Aliases

Edit `python/config/stations.toml`:

```toml
[existingstation]
aliases = ["alias1", "alias2", "newalias"]
```

---

## Output Files

| File | Description |
|------|-------------|
| `<output>.csv` | Main output with processed tracks |
| `<output>_additional.csv` | Lines matching `--additional` filter (if used) |
| `<date>-reject-<station>.csv` | Rejected lines (stopword matches) |

---

## License

MIT License. See [LICENSE](LICENSE) file.
