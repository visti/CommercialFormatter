"""Station definitions loaded from TOML configuration."""

import re
import tomllib
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional

# Config directory path
CONFIG_DIR = Path(__file__).parent / "config"
STATIONS_FILE = CONFIG_DIR / "stations.toml"
STOPWORDS_FILE = CONFIG_DIR / "stopwords.toml"


@dataclass
class Station:
    """Station configuration with pre-compiled stopword matcher."""
    name: str
    ext: list[str]
    positional: bool = False
    positions: list[int] = field(default_factory=list)
    has_headlines: bool = False
    skip_lines: int = 0  # Number of header lines to skip (0 = auto from has_headlines)
    convert: bool = False
    separator: str = ";"  # Output separator
    input_separator: str = ";"  # Input separator (defaults to same as output)
    headlines: list[str] = field(default_factory=list)
    field_mappings: dict[str, int] = field(default_factory=dict)  # Semantic field indices
    transformations: list[str] = field(default_factory=list)  # Line transformations to apply
    fix_artist_title_split: bool = False  # Prompt to fix "Artist - Title" split issues
    # Pre-compiled for efficiency
    _stopword_pattern: Optional[re.Pattern] = field(default=None, repr=False)
    _stopwords: list[str] = field(default_factory=list, repr=False)
    _sorted_positions: list[int] = field(default_factory=list, repr=False)
    _title_suffix_pattern: Optional[re.Pattern] = field(default=None, repr=False)

    def __post_init__(self):
        """Pre-sort positions and derive field mappings from headlines if needed."""
        if self.positions:
            self._sorted_positions = sorted(self.positions)

        # Auto-derive field_mappings from headlines if not explicitly set
        if not self.field_mappings and self.headlines:
            self._derive_field_mappings()

        # Pre-compile title suffix pattern if needed
        if "remove_title_suffix" in self.transformations:
            self._title_suffix_pattern = re.compile(r" - ABC Powerhit", re.IGNORECASE)

    def _derive_field_mappings(self):
        """Derive field indices from headlines array."""
        headline_to_field = {
            "Track Title": "title",
            "Main Artist": "artist",
            "Track Playing Time": "duration",
            "Date of Broadcasting": "date",
            "Track Starting Time": "time",
        }
        for idx, headline in enumerate(self.headlines):
            if headline in headline_to_field:
                self.field_mappings[headline_to_field[headline]] = idx

    def get_field_index(self, field_name: str, default: int = -1) -> int:
        """Get field index by semantic name (title, artist, duration, etc.)."""
        return self.field_mappings.get(field_name, default)

    @property
    def sorted_positions(self) -> list[int]:
        """Return pre-sorted positions."""
        return self._sorted_positions

    def matches_stopword(self, line: str) -> bool:
        """Check if line matches any stopword using compiled regex."""
        if self._stopword_pattern is None:
            return False
        return self._stopword_pattern.search(line.lower()) is not None

    def matches_stopword_lower(self, line_lower: str) -> bool:
        """Check if pre-lowercased line matches any stopword using compiled regex."""
        if self._stopword_pattern is None:
            return False
        return self._stopword_pattern.search(line_lower) is not None

    def get_matched_stopword(self, line_lower: str) -> Optional[str]:
        """Return the stopword that matched the line, or None if no match."""
        if self._stopword_pattern is None:
            return None
        match = self._stopword_pattern.search(line_lower)
        if match:
            return match.group(0)
        return None

    def apply_transformations(self, lines: list[str], filename: str = "") -> list[str]:
        """Apply configured transformations to lines.

        Args:
            lines: Input lines to transform.
            filename: Base filename (without extension) for prepend_filename.

        Returns:
            Transformed lines.
        """
        for transform in self.transformations:
            if transform == "prepend_filename":
                lines = self._transform_prepend_filename(lines, filename)
            elif transform == "replace_dash_separator":
                lines = self._transform_replace_dash(lines)
            elif transform == "split_datetime":
                lines = self._transform_split_datetime(lines)
            elif transform == "remove_title_suffix":
                lines = self._transform_remove_title_suffix(lines)
        return lines

    def _transform_prepend_filename(self, lines: list[str], filename: str) -> list[str]:
        """Prepend filename to each line (for Globus-style files)."""
        return [f"{filename}{line}" for line in lines if line.strip()]

    def _transform_replace_dash(self, lines: list[str]) -> list[str]:
        """Replace ' - ' with separator (for Globus-style files)."""
        return [line.replace(" - ", self.separator) for line in lines]

    def _transform_split_datetime(self, lines: list[str]) -> list[str]:
        """Split 'DD-MM-YYYY HH-HH' into separate date and time fields."""
        result = []
        for line in lines:
            if not line.strip():
                continue
            parts = line.split(self.separator, 1)
            if len(parts) >= 1 and " " in parts[0]:
                date_time = parts[0].split(" ", 1)
                if len(date_time) == 2:
                    date_part = date_time[0]
                    time_range = date_time[1]
                    hour = time_range.split("-")[0] if "-" in time_range else time_range
                    time_part = f"{hour}:00:00"
                    rest = parts[1] if len(parts) > 1 else ""
                    line = f"{date_part}{self.separator}{time_part}{self.separator}{rest}"
            result.append(line)
        return result

    def _transform_remove_title_suffix(self, lines: list[str]) -> list[str]:
        """Remove title suffix pattern (e.g., ' - ABC Powerhit')."""
        if self._title_suffix_pattern is None:
            return lines
        title_idx = self.get_field_index("title", -1)
        if title_idx < 0:
            return lines

        result = []
        for line in lines:
            if line.strip():
                fields = line.split(self.separator)
                if len(fields) > title_idx:
                    fields[title_idx] = self._title_suffix_pattern.sub("", fields[title_idx])
                    line = self.separator.join(fields)
            result.append(line)
        return result


def _load_stopwords() -> dict[str, list[str]]:
    """Load stopwords from TOML config file."""
    if not STOPWORDS_FILE.exists():
        return {"default": []}

    with open(STOPWORDS_FILE, "rb") as f:
        data = tomllib.load(f)

    return {key: value.get("words", []) for key, value in data.items()}


def _load_stations_config() -> dict:
    """Load station definitions from TOML config file."""
    if not STATIONS_FILE.exists():
        raise FileNotFoundError(f"Stations config not found: {STATIONS_FILE}")

    with open(STATIONS_FILE, "rb") as f:
        return tomllib.load(f)


def _compile_stopword_pattern(words: list[str]) -> Optional[re.Pattern]:
    """Compile stopwords into a single regex pattern for fast matching."""
    if not words:
        return None

    # Escape special regex chars and join with | for OR matching
    # Pre-lowercase all patterns
    escaped = [re.escape(word.lower()) for word in words]
    pattern = "|".join(escaped)
    return re.compile(pattern)


def _build_stations() -> tuple[dict[str, Station], dict[str, str]]:
    """Build station objects and alias map from config files."""
    stations_config = _load_stations_config()
    stopwords_config = _load_stopwords()
    default_stopwords = stopwords_config.get("default", [])

    stations = {}
    aliases = {}  # Maps alias -> station key

    for key, cfg in stations_config.items():
        # Combine default + station-specific stopwords
        station_stopwords = stopwords_config.get(key, {})
        if isinstance(station_stopwords, dict):
            station_stopwords = station_stopwords.get("words", [])
        all_stopwords = default_stopwords + station_stopwords

        has_headlines = cfg.get("has_headlines", False)
        # Default skip_lines: 1 if has_headlines, else 0
        default_skip = 1 if has_headlines else 0
        skip_lines = cfg.get("skip_lines", default_skip)

        output_sep = cfg.get("separator", ";")
        input_sep = cfg.get("input_separator", output_sep)

        station = Station(
            name=cfg.get("name", key),
            ext=cfg.get("extensions", []),
            positional=cfg.get("positional", False),
            positions=cfg.get("positions", []),
            has_headlines=has_headlines,
            skip_lines=skip_lines,
            convert=cfg.get("convert", False),
            separator=output_sep,
            input_separator=input_sep,
            headlines=cfg.get("headlines", []),
            field_mappings=cfg.get("field_mappings", {}),
            transformations=cfg.get("transformations", []),
            fix_artist_title_split=cfg.get("fix_artist_title_split", False),
        )

        # Store stopwords and compile pattern
        station._stopwords = all_stopwords
        station._stopword_pattern = _compile_stopword_pattern(all_stopwords)

        stations[key.lower()] = station

        # Build alias map
        for alias in cfg.get("aliases", []):
            aliases[alias.lower()] = key.lower()

    return stations, aliases


# Load stations at module import time
STATIONS, ALIASES = _build_stations()


def get_station(name: str) -> Optional[Station]:
    """Get a station by name or alias (case-insensitive).

    Direct station definitions take precedence over aliases.
    """
    name_lower = name.lower()

    # Direct station lookup first
    if name_lower in STATIONS:
        return STATIONS[name_lower]

    # Fall back to alias resolution
    if name_lower in ALIASES:
        resolved = ALIASES[name_lower]
        return STATIONS.get(resolved)

    return None


def list_stations() -> list[str]:
    """Return list of available station names."""
    return list(STATIONS.keys())


def list_aliases() -> dict[str, list[str]]:
    """Return dict of station -> list of aliases."""
    result = {key: [] for key in STATIONS.keys()}
    for alias, station in ALIASES.items():
        result[station].append(alias)
    return result
