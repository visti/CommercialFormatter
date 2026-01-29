"""User choice persistence for commercial formatter."""

from __future__ import annotations

import tomllib
from datetime import datetime
from pathlib import Path
from typing import Literal

from settings import get_settings

# Choice types
ChoiceAction = Literal["fix", "skip", "reject"]


class ChoicesManager:
    """Manages remembered user choices for artist/title fixes."""

    def __init__(self, config_dir: Path | None = None):
        """Initialize the choices manager.

        Args:
            config_dir: Directory for config files. If None, uses default.
        """
        if config_dir is None:
            config_dir = Path(__file__).parent / "config"

        self.config_dir = config_dir
        settings = get_settings()
        self.choices_file = config_dir / Path(settings.choices.choices_file).name
        self.enabled = settings.choices.remember_fixes
        self._choices: dict[str, dict] = {}
        self._load()

    def _load(self) -> None:
        """Load choices from file."""
        if not self.choices_file.exists():
            self._choices = {"artist_title_fixes": {}, "long_playing_times": {}}
            return

        try:
            with open(self.choices_file, "rb") as f:
                self._choices = tomllib.load(f)
        except Exception:
            self._choices = {"artist_title_fixes": {}, "long_playing_times": {}}

    def _save(self) -> None:
        """Save choices to file."""
        if not self.enabled:
            return

        # Ensure directory exists
        self.choices_file.parent.mkdir(parents=True, exist_ok=True)

        # Write TOML manually (tomllib is read-only)
        lines = [
            "# Remembered user choices",
            f"# Last updated: {datetime.now().isoformat()}",
            "",
            "[artist_title_fixes]",
        ]

        for key, data in self._choices.get("artist_title_fixes", {}).items():
            # Escape the key for TOML
            safe_key = key.replace("\\", "\\\\").replace('"', '\\"')
            lines.append(f'"{safe_key}" = "{data["action"]}"')

        lines.append("")
        lines.append("[long_playing_times]")

        for key, data in self._choices.get("long_playing_times", {}).items():
            safe_key = key.replace("\\", "\\\\").replace('"', '\\"')
            if data["action"] == "edit":
                lines.append(f'"{safe_key}" = {{ action = "edit", time = "{data["time"]}" }}')
            else:
                lines.append(f'"{safe_key}" = "{data["action"]}"')

        try:
            with open(self.choices_file, "w", encoding="utf-8") as f:
                f.write("\n".join(lines) + "\n")
        except Exception:
            pass  # Silently fail if we can't save

    def _make_key(self, title: str, artist: str) -> str:
        """Create a unique key for a title/artist combination."""
        return f"{title}|||{artist}"

    def get_artist_title_choice(
        self, title: str, artist: str
    ) -> ChoiceAction | None:
        """Get remembered choice for an artist/title fix.

        Args:
            title: Track title
            artist: Artist name

        Returns:
            The remembered action, or None if not remembered
        """
        if not self.enabled:
            return None

        key = self._make_key(title, artist)
        data = self._choices.get("artist_title_fixes", {}).get(key)

        if data is None:
            return None

        if isinstance(data, str):
            return data  # type: ignore
        return data.get("action")

    def remember_artist_title_choice(
        self, title: str, artist: str, action: ChoiceAction
    ) -> None:
        """Remember a choice for an artist/title fix.

        Args:
            title: Track title
            artist: Artist name
            action: The action taken (fix, skip, reject)
        """
        if not self.enabled:
            return

        key = self._make_key(title, artist)
        if "artist_title_fixes" not in self._choices:
            self._choices["artist_title_fixes"] = {}

        self._choices["artist_title_fixes"][key] = {"action": action}
        self._save()

    def get_playing_time_choice(
        self, title: str, artist: str, time: str
    ) -> tuple[ChoiceAction | None, str | None]:
        """Get remembered choice for a long playing time.

        Args:
            title: Track title
            artist: Artist name
            time: Playing time string

        Returns:
            Tuple of (action, edited_time) or (None, None) if not remembered
        """
        if not self.enabled:
            return None, None

        key = self._make_key(title, artist)
        data = self._choices.get("long_playing_times", {}).get(key)

        if data is None:
            return None, None

        if isinstance(data, str):
            return data, None  # type: ignore

        action = data.get("action")
        edited_time = data.get("time")
        return action, edited_time

    def remember_playing_time_choice(
        self,
        title: str,
        artist: str,
        time: str,
        action: ChoiceAction,
        edited_time: str | None = None,
    ) -> None:
        """Remember a choice for a long playing time.

        Args:
            title: Track title
            artist: Artist name
            time: Original playing time
            action: The action taken (accept, reject, edit)
            edited_time: The corrected time if action was edit
        """
        if not self.enabled:
            return

        key = self._make_key(title, artist)
        if "long_playing_times" not in self._choices:
            self._choices["long_playing_times"] = {}

        if action == "edit" and edited_time:
            self._choices["long_playing_times"][key] = {
                "action": action,
                "time": edited_time,
            }
        else:
            self._choices["long_playing_times"][key] = {"action": action}

        self._save()

    def clear_all(self) -> None:
        """Clear all remembered choices."""
        self._choices = {"artist_title_fixes": {}, "long_playing_times": {}}
        self._save()


# Global instance
_manager: ChoicesManager | None = None


def get_choices_manager() -> ChoicesManager:
    """Get the global choices manager instance."""
    global _manager
    if _manager is None:
        _manager = ChoicesManager()
    return _manager
