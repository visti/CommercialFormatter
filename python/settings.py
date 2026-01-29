"""Settings management for commercial formatter."""

from __future__ import annotations

import tomllib
from dataclasses import dataclass, field
from pathlib import Path
from typing import Literal


@dataclass
class ThresholdSettings:
    """Threshold configuration."""
    long_playing_time_minutes: int = 30
    overflow_threshold_minutes: int = 1400


@dataclass
class BackupSettings:
    """Backup configuration."""
    enabled: bool = True
    directory: str = "backup"


@dataclass
class LoggingSettings:
    """Logging configuration."""
    enabled: bool = True
    filename: str = "logs/komm_fmt_{date}_{station}.log"
    level: Literal["DEBUG", "INFO", "WARNING", "ERROR"] = "INFO"


@dataclass
class DuplicateSettings:
    """Duplicate detection configuration."""
    enabled: bool = True
    action: Literal["prompt", "keep", "reject"] = "prompt"


@dataclass
class ChoicesSettings:
    """User choices configuration."""
    remember_fixes: bool = True
    choices_file: str = "config/remembered_choices.toml"


@dataclass
class Settings:
    """Global application settings."""
    thresholds: ThresholdSettings = field(default_factory=ThresholdSettings)
    backup: BackupSettings = field(default_factory=BackupSettings)
    logging: LoggingSettings = field(default_factory=LoggingSettings)
    duplicates: DuplicateSettings = field(default_factory=DuplicateSettings)
    choices: ChoicesSettings = field(default_factory=ChoicesSettings)


def load_settings(config_dir: Path | None = None) -> Settings:
    """Load settings from TOML file.

    Args:
        config_dir: Directory containing settings.toml. If None, uses default.

    Returns:
        Settings object with loaded or default values.
    """
    if config_dir is None:
        config_dir = Path(__file__).parent / "config"

    settings_file = config_dir / "settings.toml"

    if not settings_file.exists():
        return Settings()

    try:
        with open(settings_file, "rb") as f:
            data = tomllib.load(f)
    except Exception:
        return Settings()

    # Build settings from TOML data
    thresholds = ThresholdSettings(
        long_playing_time_minutes=data.get("thresholds", {}).get(
            "long_playing_time_minutes", 30
        ),
        overflow_threshold_minutes=data.get("thresholds", {}).get(
            "overflow_threshold_minutes", 1400
        ),
    )

    backup = BackupSettings(
        enabled=data.get("backup", {}).get("enabled", True),
        directory=data.get("backup", {}).get("directory", "backup"),
    )

    logging_data = data.get("logging", {})
    logging = LoggingSettings(
        enabled=logging_data.get("enabled", True),
        filename=logging_data.get("filename", "logs/komm_fmt_{date}_{station}.log"),
        level=logging_data.get("level", "INFO"),
    )

    duplicates = DuplicateSettings(
        enabled=data.get("duplicates", {}).get("enabled", True),
        action=data.get("duplicates", {}).get("action", "prompt"),
    )

    choices = ChoicesSettings(
        remember_fixes=data.get("choices", {}).get("remember_fixes", True),
        choices_file=data.get("choices", {}).get(
            "choices_file", "config/remembered_choices.toml"
        ),
    )

    return Settings(
        thresholds=thresholds,
        backup=backup,
        logging=logging,
        duplicates=duplicates,
        choices=choices,
    )


# Global settings instance (loaded once)
_settings: Settings | None = None


def get_settings() -> Settings:
    """Get the global settings instance."""
    global _settings
    if _settings is None:
        _settings = load_settings()
    return _settings


def reload_settings() -> Settings:
    """Reload settings from disk."""
    global _settings
    _settings = load_settings()
    return _settings
