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
    separator: str = ";"
    headlines: list[str] = field(default_factory=list)
    # Pre-compiled for efficiency
    _stopword_pattern: Optional[re.Pattern] = field(default=None, repr=False)
    _sorted_positions: list[int] = field(default_factory=list, repr=False)

    def __post_init__(self):
        """Pre-sort positions for efficient processing."""
        if self.positions:
            self._sorted_positions = sorted(self.positions)

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

        station = Station(
            name=cfg.get("name", key),
            ext=cfg.get("extensions", []),
            positional=cfg.get("positional", False),
            positions=cfg.get("positions", []),
            has_headlines=has_headlines,
            skip_lines=skip_lines,
            convert=cfg.get("convert", False),
            separator=cfg.get("separator", ";"),
            headlines=cfg.get("headlines", []),
        )

        # Compile stopword pattern
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
