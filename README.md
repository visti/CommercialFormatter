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
komm_fmt [station] [options]

# Examples
komm_fmt Bauer
komm_fmt                             # Auto-detect station from folder path
komm_fmt GoFM                        # Uses ABC via alias
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

## Features

### Auto-Detection from Folder Path

When no station is specified, the tool automatically detects the station from the current working directory path using mappings in `config/folders.toml`.

```bash
# Working in: C:\Users\eva\Gramex\...\ABC\2025\q4\
komm_fmt
# Output: Auto-detected station: abc
```

### Suggested Output Filename

The tool suggests an output filename based on folder structure:

```
# Working in: .../Silkeborg/2025/q4/
Set output filename [2025_q4_silkeborg.csv]:
```

Press Enter to accept, or type a custom name.

### File Access Handling

If the output file is open in another application (e.g., Excel), the tool prompts you:

```
Cannot access '2025_q4_abc.csv' - file is in use.
Close the file and press Enter to retry, or X to quit:
```

### Playing Time Overflow Fix

Tracks starting before midnight sometimes have incorrect playing times due to a buffer overflow (adds 1440 minutes). The tool automatically corrects times where minutes >= 1400:

```
Original: 1436:41  →  Corrected: 03:19
Original: 1438:03  →  Corrected: 01:57
```

### Long Playing Time Check

Tracks with playing times over 30 minutes trigger an interactive prompt:

```
Found 2 unique track(s) with playing time over 30 minutes

  Title:  "Extended Club Mix"
  Artist: "DJ Artist"
  Time:   45:30 (3x)
  [A]ccept / [R]eject / [E]dit:
```

Options:
- **A** (or Enter): Accept the track as-is
- **R**: Reject - move to rejection file
- **E**: Edit - enter a corrected time (MM:SS format)

Example of editing:
```
  [A]ccept / [R]eject / [E]dit: e
  Enter corrected time (MM:SS): 04:30
  Updated 3 occurrence(s) to 04:30
```

### ABC Station Enhancements

#### PowerHit Suffix Removal

Automatically removes " - ABC PowerHit" (case-insensitive) from track titles:

```
Before: "Stay (if you wanna dance) - ABC PowerHit"
After:  "Stay (if you wanna dance)"

Before: "Ferrari - ABC Powerhit"
After:  "Ferrari"
```

#### Artist/Title Split Fix

Detects and offers to fix incorrectly split artist/title fields containing " - ":

```
Found 3 unique artist/title split issue(s)

  Current: Title="bi-dua - Krig Og Fred" Artist="Shu" (5x)
  Fixed:   Title="Krig Og Fred" Artist="Shu-bi-dua"
  [Y]es fix / [N]o skip / [X] reject:
```

Options:
- **Y** (or Enter): Apply the fix
- **N**: Skip - keep original
- **X**: Reject - move to rejection file

#### ABC-Specific Stopwords

Lines containing these patterns are automatically rejected:
- "Radio ABC"
- "ABC live"

These are checked before the artist/title split prompt, so rejected lines won't appear in the prompt.

### Date Format Normalization

Automatically normalizes various date formats to DD-MM-YYYY:

| Input Format | Example | Output |
|--------------|---------|--------|
| YYMMDD | 251001 | 01-10-2025 |
| YYYY-MM-DD | 2025-10-01 | 01-10-2025 |
| DD.MM.YYYY | 01.10.2025 | 01-10-2025 |

### Time Format Normalization

Converts 6-digit time strings to HH:MM:SS:

```
Input:  235917
Output: 23:59:17
```

---

## File Structure

```
commercial_formatter/
├── python/                     # Python implementation
│   ├── main.py                 # CLI entry point
│   ├── processor.py            # Core processing logic
│   ├── stations.py             # Station loading from TOML
│   ├── output.py               # Console output and formatting
│   ├── config.py               # Global paths and constants
│   ├── pyproject.toml          # Package configuration
│   ├── config/
│   │   ├── stations.toml       # Station definitions
│   │   ├── stopwords.toml      # Rejection wordlists
│   │   └── folders.toml        # Folder-to-station mappings
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

## Configuration

### Station Definitions (`config/stations.toml`)

Each station section defines parsing rules:

```toml
[abc]
name = "ABC"
aliases = ["solo", "gofm", "radiom", "silkeborg"]
extensions = ["txt", "csv"]
positional = false
has_headlines = false
convert = false
separator = ";"
positions = []
headlines = [
  "Date of Broadcasting",
  "Track Starting Time",
  "Track Playing Time",
  "DELETE",              # Columns marked DELETE are removed
  "DELETE",
  "DELETE",
  "DELETE",
  "Track Title",
  "Main Artist",
  "DELETE",
  "DELETE",
  "DELETE",
  "DELETE",
  "DELETE",
]
```

| Field | Description |
|-------|-------------|
| `name` | Display name |
| `aliases` | Alternative names that resolve to this station |
| `extensions` | File extensions to process |
| `positional` | `true` for fixed-width, `false` for delimited |
| `positions` | Column boundaries for positional extraction |
| `has_headlines` | `true` if input files have a header row |
| `skip_lines` | Number of header lines to skip (default: 1 if has_headlines) |
| `convert` | `true` to run XLSX converter first |
| `separator` | Output field separator (`;` or `:`) |
| `headlines` | Column headers (`"DELETE"` marks columns for removal) |

### Stopwords (`config/stopwords.toml`)

Lines containing any stopword are rejected:

```toml
[default]
# Applied to ALL stations
words = [
    "promo",
    "jingle",
    "sweeper",
    "nyheder",
    "vejr",
    "reklame",
]

[bauer]
# Additional stopwords for Bauer stations
words = [
    "PODCAST",
    "TOP HOUR",
    "NO News",
]

[abc]
# ABC-specific stopwords
words = [
    "Radio ABC",
    "ABC live",
]
```

### Folder Mappings (`config/folders.toml`)

Maps folder names to station aliases for auto-detection:

```toml
[folders]
# Bauer stations
"100FM" = "100fm"
"myRock" = "myrock"
"Nova" = "nova"
"Voice" = "voice"

# ABC stations
"ABC" = "abc"
"Silkeborg" = "silkeborg"
"Solo" = "solo"

# Other stations
"Radio4" = "radio4"
"Globus" = "globus"
```

---

## Processing Pipeline

1. **File Discovery**: Find files matching station's extensions
2. **Encoding Detection**: Auto-detect file encoding (UTF-8, CP1252, etc.)
3. **Header Skipping**: Skip configured number of header lines
4. **Station Transforms**: Apply station-specific preprocessing
   - Globus: Prepend filename, replace " - " with ";"
   - Skive: Split date/time field
   - ABC: Remove PowerHit suffix, fix artist/title splits
5. **Long Playing Time Check**: Prompt for tracks > 30 minutes
6. **Stopword Filtering**: Reject lines matching stopwords
7. **Line Processing**: Format dates, times, extract fields
8. **Playing Time Fix**: Correct midnight overflow times
9. **Output Routing**: Write to main, additional, or reject file
10. **Column Cleanup**: Remove DELETE columns, empty rows

---

## Output Files

| File | Description |
|------|-------------|
| `<output>.csv` | Main output with processed tracks |
| `<output>_additional.csv` | Lines matching `--additional` filter |
| `<date>-reject-<station>.csv` | Rejected lines (stopwords, user rejections) |

---

## Examples

### Basic Usage

```bash
# Process Bauer station files
komm_fmt bauer

# Auto-detect station from current folder
komm_fmt

# Process with additional filter for specific artist
komm_fmt abc --additional="Spotify"
```

### Interactive Session Example

```
$ komm_fmt abc

Station: ABC
Found 2 file(s):
  - Radio ABC okt nov dec 2025.txt

Set output filename [2025_q4_abc.csv]:

Found 1 unique artist/title split issue(s)

  Current: Title="bi-dua - Krig Og Fred" Artist="Shu" (5x)
  Fixed:   Title="Krig Og Fred" Artist="Shu-bi-dua"
  [Y]es fix / [N]o skip / [X] reject: y
  Will fix 5 occurrence(s)

Found 2 unique track(s) with playing time over 30 minutes

  Title:  "Extended Remix"
  Artist: "Various"
  Time:   65:30 (1x)
  [A]ccept / [R]eject / [E]dit: e
  Enter corrected time (MM:SS): 06:30
  Updated 1 occurrence(s) to 06:30

  Title:  "Live Recording"
  Artist: "Band"
  Time:   45:00 (2x)
  [A]ccept / [R]eject / [E]dit: r
  Will reject 2 occurrence(s)

Processing lines...

┌─────────────────────────────────────┐
│            Summary                  │
├─────────────────────────────────────┤
│  Files processed:    2              │
│  Lines processed:    1,234          │
│  Lines rejected:     56             │
│  Output file:        2025_q4_abc.csv│
└─────────────────────────────────────┘
```

### Handling Locked Files

```
$ komm_fmt abc

Set output filename [2025_q4_abc.csv]:

Cannot access '2025_q4_abc.csv' - file is in use.
Close the file and press Enter to retry, or X to quit:

[User closes Excel, presses Enter]

Overwriting existing 2025_q4_abc.csv
Processing lines...
```

---

## Adding a New Station

1. Edit `python/config/stations.toml`:

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
    "Track Starting Time",
    "Track Playing Time",
    "Track Title",
    "Main Artist",
]
```

2. Add stopwords in `python/config/stopwords.toml`:

```toml
[newstation]
words = [
    "station jingle",
    "news break",
]
```

3. Add folder mapping in `python/config/folders.toml`:

```toml
[folders]
"NewStation" = "ns"
```

---

## Supported Stations

| Station | Aliases | Format |
|---------|---------|--------|
| Bauer | nova, radio100, 100fm, thevoice, voice, myrock, pop | Positional |
| Jyskfynske | jfm, vlr, viborg, classicfm, skala | Positional |
| Globus | - | Delimited |
| Radio4 | r4, "radio 4" | CSV (XLSX convert) |
| ANR | nordjyske, radionordjylland, nordjylland | Positional |
| ABC | solo, gofm, radiom, silkeborg | Delimited |
| Skive | "radio skive" | Delimited |

---

## License

MIT License. See [LICENSE](LICENSE) file.
