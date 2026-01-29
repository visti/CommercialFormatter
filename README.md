# Commercial Formatter

A tool for processing broadcast metadata files from Danish radio stations. Extracts track information, filters unwanted content (jingles, promos, etc.), and outputs clean CSV files for reporting.

## Installation

### From Source

```bash
cd python
pip install -e .
```

This installs the `komm_fmt` command globally.

### Docker

```bash
cd python
docker build -t commercial-formatter .
docker run -it -v /path/to/files:/data commercial-formatter [station]
```

### Usage

```bash
komm_fmt [station] [options]

# Examples
komm_fmt Bauer
komm_fmt                             # Auto-detect station from folder path
komm_fmt GoFM                        # Uses ABC via alias
komm_fmt Radio4 --additional=Boulevard
komm_fmt Bauer --no-stopwords
komm_fmt abc -f "Radio ABC.txt"      # Process specific file only
komm_fmt --list-stations             # Show all stations and aliases
komm_fmt --edit-choices              # Edit remembered decisions in nvim
komm_fmt --reject-path               # Open rejection log folder
```

### Options

| Option | Description |
|--------|-------------|
| `-f, --file FILE` | Process specific file(s) instead of all matching (can be repeated) |
| `--additional=FILTER` | Route lines containing FILTER to a separate CSV |
| `--additional-postfix=TEXT` | Suffix for additional file (default: `_additional`) |
| `--no-stopwords` | Disable stopword filtering |
| `--no-reject-file` | Skip saving rejected lines to the rejection log |
| `--list-stations` | List all stations and their aliases |
| `--edit-choices` | Open remembered choices file in nvim for editing |
| `--reject-path` | Open the rejection log folder in file explorer |

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

### Date Format Normalization

Automatically normalizes various date formats to DD-MM-YYYY:

| Input Format | Example | Output |
|--------------|---------|--------|
| YYMMDD | 251001 | 01-10-2025 |
| DDMMYYYY | 01102025 | 01-10-2025 |
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
└── python/
    ├── main.py                 # CLI entry point
    ├── processor.py            # Core processing logic
    ├── stations.py             # Station loading from TOML
    ├── output.py               # Console output and formatting
    ├── config.py               # Global paths and constants
    ├── choices.py              # User choice persistence
    ├── decisions.py            # Decision management (remember/prompt/apply)
    ├── formatters.py           # Field formatting (date, time, duration)
    ├── settings.py             # Settings management
    ├── app_logging.py          # Logging utilities
    ├── pyproject.toml          # Package configuration
    ├── Dockerfile              # Docker build file
    ├── config/
    │   ├── stations.toml       # Station definitions
    │   ├── stopwords.toml      # Rejection wordlists
    │   ├── folders.toml        # Folder-to-station mappings
    │   └── settings.toml       # Application settings
    └── lib/
        ├── convert.py          # XLSX to CSV converter
        ├── delete_columns.py   # Remove DELETE columns from output
        └── delete_podcast.py   # Podcast removal utility
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
transformations = ["remove_title_suffix"]
fix_artist_title_split = true
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

#### Basic Fields

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
| `separator` | Output field separator |
| `input_separator` | Input field separator (defaults to `separator`) |
| `headlines` | Column headers (`"DELETE"` marks columns for removal) |

#### Advanced Fields

| Field | Description |
|-------|-------------|
| `transformations` | List of line transformations to apply (see below) |
| `fix_artist_title_split` | `true` to prompt for fixing "Artist - Title" issues |
| `field_mappings` | Custom field indices (auto-derived from headlines if not set) |

#### Available Transformations

| Transformation | Description | Used By |
|----------------|-------------|---------|
| `prepend_filename` | Prepend filename to each line | Globus |
| `replace_dash_separator` | Replace " - " with separator | Globus |
| `split_datetime` | Split "DD-MM-YYYY HH-HH" into date + time | Skive, Skaga |
| `remove_title_suffix` | Remove " - ABC Powerhit" from titles | ABC |

#### Field Mappings

Field indices are automatically derived from the `headlines` array by looking for these column names:
- `"Track Title"` → `title`
- `"Main Artist"` → `artist`
- `"Track Playing Time"` → `duration`
- `"Date of Broadcasting"` → `date`
- `"Track Starting Time"` → `time`

You can override with explicit mappings if needed:
```toml
field_mappings = { title = 7, artist = 8, duration = 2 }
```

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

```
┌─────────────────────────────────────────────────────────────────────────────┐
│                              main.py                                         │
│  ┌─────────────┐    ┌─────────────┐    ┌─────────────┐    ┌─────────────┐   │
│  │ parse_args  │───▶│ get_station │───▶│  get_files  │───▶│ read_files  │   │
│  └─────────────┘    └─────────────┘    └─────────────┘    └──────┬──────┘   │
│                                                                   │          │
│                                                                   ▼          │
│                                                          ┌─────────────┐     │
│                                                          │process_files│     │
│                                                          └──────┬──────┘     │
│                                                                 │            │
│                                                                 ▼            │
│                                                      ┌──────────────────┐    │
│                                                      │print_summary_box │    │
│                                                      └──────────────────┘    │
└─────────────────────────────────────────────────────────────────────────────┘

┌─────────────────────────────────────────────────────────────────────────────┐
│                           read_files() Flow                                  │
│                                                                              │
│  ┌──────────────────┐                                                        │
│  │ read_single_file │  (for each file)                                       │
│  └────────┬─────────┘                                                        │
│           │                                                                  │
│           ▼                                                                  │
│  ┌──────────────────┐    ┌───────────────────────────┐                       │
│  │ detect_encoding  │───▶│ apply_transformations     │                       │
│  └──────────────────┘    │ (split_datetime,          │                       │
│                          │  prepend_filename, etc.)  │                       │
│                          └─────────────┬─────────────┘                       │
│                                        │                                     │
│           ┌────────────────────────────┼────────────────────────────┐        │
│           ▼                            ▼                            ▼        │
│  ┌─────────────────────┐    ┌─────────────────────┐    ┌─────────────────┐   │
│  │check_artist_title   │    │check_long_playing   │    │ check_duplicates│   │
│  │_split()             │    │_times()             │    │ ()              │   │
│  └─────────┬───────────┘    └─────────┬───────────┘    └────────┬────────┘   │
│            │                          │                         │            │
│            └──────────────────────────┼─────────────────────────┘            │
│                                       ▼                                      │
│                            ┌─────────────────────┐                           │
│                            │  DecisionManager    │                           │
│                            │  (remember/prompt/  │                           │
│                            │   apply pattern)    │                           │
│                            └─────────────────────┘                           │
└─────────────────────────────────────────────────────────────────────────────┘

┌─────────────────────────────────────────────────────────────────────────────┐
│                         process_files() Flow                                 │
│                                                                              │
│  ┌─────────────────┐                                                         │
│  │ For each line:  │                                                         │
│  └────────┬────────┘                                                         │
│           │                                                                  │
│           ▼                                                                  │
│  ┌─────────────────────┐     ┌─────────────────────┐                         │
│  │ Force rejected?     │─Yes─▶│ Write to reject_f   │                         │
│  └─────────┬───────────┘     └─────────────────────┘                         │
│            │ No                                                              │
│            ▼                                                                  │
│  ┌─────────────────────┐     ┌─────────────────────┐                         │
│  │ Matches stopword?   │─Yes─▶│ Write to reject_f   │                         │
│  └─────────┬───────────┘     └─────────────────────┘                         │
│            │ No                                                              │
│            ▼                                                                  │
│  ┌─────────────────────┐                                                     │
│  │   process_line()    │                                                     │
│  │  ┌───────────────┐  │                                                     │
│  │  │ format_date   │  │                                                     │
│  │  │ format_time   │  │                                                     │
│  │  │format_duration│  │                                                     │
│  │  └───────────────┘  │                                                     │
│  └─────────┬───────────┘                                                     │
│            │                                                                  │
│            ▼                                                                  │
│  ┌─────────────────────┐     ┌─────────────────────┐                         │
│  │ Matches additional? │─Yes─▶│ Write to additional │                         │
│  └─────────┬───────────┘     └─────────────────────┘                         │
│            │ No                                                              │
│            ▼                                                                  │
│  ┌─────────────────────┐                                                     │
│  │ Write to output     │                                                     │
│  └─────────────────────┘                                                     │
│                                                                              │
│  ┌─────────────────────┐                                                     │
│  │ Cleanup:            │                                                     │
│  │ run_delete_columns()│──▶ Removes DELETE columns and empty rows            │
│  └─────────────────────┘                                                     │
└─────────────────────────────────────────────────────────────────────────────┘
```

### Pipeline Steps

1. **File Discovery**: Find files matching station's extensions
2. **Encoding Detection**: Auto-detect file encoding (UTF-8, CP1252, etc.)
3. **Header Skipping**: Skip configured number of header lines
4. **Transformations**: Apply configured transformations (e.g., `split_datetime`, `prepend_filename`)
5. **Artist/Title Fix**: Prompt for "Artist - Title" issues (if `fix_artist_title_split = true`)
6. **Long Playing Time Check**: Prompt for tracks > 30 minutes
7. **Stopword Filtering**: Reject lines matching stopwords
8. **Line Processing**: Format dates, times, extract fields
9. **Playing Time Fix**: Correct midnight overflow times
10. **Output Routing**: Write to main, additional, or reject file
11. **Column Cleanup**: Remove DELETE columns, empty rows

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

Most new stations can be added with **config-only changes** - no Python code needed.

### 1. Add station definition (`config/stations.toml`)

```toml
[newstation]
name = "New Station"
aliases = ["ns", "newst"]
extensions = ["txt"]
positional = false           # false = delimited, true = fixed-width
has_headlines = true         # true if first row is headers
skip_lines = 1               # lines to skip (default: 1 if has_headlines)
separator = ";"              # output separator
input_separator = ","        # input separator (if different from output)
transformations = []         # optional: ["split_datetime", "prepend_filename", etc.]
fix_artist_title_split = false  # optional: prompt to fix "Artist - Title" issues
headlines = [
    "Date of Broadcasting",
    "Track Starting Time",
    "Track Playing Time",
    "Track Title",
    "Main Artist",
    "DELETE",                # columns marked DELETE are removed
]
```

### 2. Add stopwords (`config/stopwords.toml`)

```toml
[newstation]
words = [
    "station jingle",
    "news break",
]
```

### 3. Add folder mapping (`config/folders.toml`)

```toml
[folders]
"NewStation" = "ns"
"New Station Folder" = "newstation"
```

### Examples by Format Type

**CSV with comma input, semicolon output (like Limfjord):**
```toml
[mystation]
name = "My Station"
extensions = ["csv"]
positional = false
has_headlines = true
separator = ";"
input_separator = ","
headlines = ["Date of Broadcasting", "Track Starting Time", ...]
```

**Combined date/time field (like Skive/Skaga):**
```toml
[mystation]
name = "My Station"
extensions = ["csv"]
positional = false
has_headlines = true
skip_lines = 2              # skip extra header line
separator = ";"
transformations = ["split_datetime"]
headlines = ["Date of Broadcasting", "Track Starting Time", ...]  # after split
```

**Fixed-width format (like Bauer):**
```toml
[mystation]
name = "My Station"
extensions = ["txt"]
positional = true
positions = [185, 179, 173, 168, 163, 153, 128, 78, 28, 19, 14, 6]
has_headlines = false
separator = ";"
headlines = ["Date of Broadcasting", "Track Starting Time", ...]
```

---

## Supported Stations

| Station | Aliases | Format | Notes |
|---------|---------|--------|-------|
| Bauer | nova, radio100, 100fm, thevoice, voice, myrock, pop | Positional | |
| Jyskfynske | jfm, vlr, viborg, classicfm, skala | Positional | |
| Globus | - | Delimited | Prepends filename |
| Radio4 | r4, "radio 4" | CSV | XLSX conversion |
| ANR | nordjyske, radionordjylland, nordjylland | Positional | |
| ABC | solo, gofm, radiom, silkeborg | Delimited | PowerHit removal, artist/title fix |
| Skive | "radio skive" | CSV | Date/time split |
| Skaga | "radio hirtshals", hirtshals | CSV | Date/time split |
| Limfjord | "limfjord mix", "limfjord slager", "radio limfjord" | CSV | Comma input |

---

## License

MIT License. See [LICENSE](LICENSE) file.
