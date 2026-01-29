"""Decision management for user prompts with remember/apply flow."""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any, Callable

import app_logging as logging
import output as console


@dataclass
class Issue:
    """An issue requiring user decision."""

    key: tuple[Any, ...]  # Unique identifier
    indices: list[int]  # Line indices affected
    display_data: dict[str, Any] = field(default_factory=dict)  # Data for display


@dataclass
class Option:
    """A decision option."""

    key: str  # Single char like "y", "r", "e"
    aliases: list[str]  # Other accepted inputs like "yes", "reject"
    label: str  # Display label like "Yes fix"
    action: str  # Action name like "fix", "reject", "skip"
    is_default: bool = False  # If true, empty input selects this


@dataclass
class DecisionConfig:
    """Configuration for a decision type."""

    name: str  # "artist_title", "long_time", "duplicate"
    options: list[Option]
    prompt_prefix: str = "  "

    def get_prompt_text(self) -> str:
        """Generate prompt text from options."""
        parts = []
        for opt in self.options:
            label = f"[{opt.key.upper()}]{opt.label[1:]}" if opt.label else f"[{opt.key.upper()}]"
            parts.append(label)
        return self.prompt_prefix + " / ".join(parts) + ": "

    def parse_response(self, response: str) -> str | None:
        """Parse user response to action name."""
        response = response.strip().lower()

        # Check for default (empty input)
        if not response:
            for opt in self.options:
                if opt.is_default:
                    return opt.action
            return None

        # Check each option
        for opt in self.options:
            if response == opt.key or response in opt.aliases:
                return opt.action

        return None


class DecisionManager:
    """Manages user decision flow with remember/prompt/apply pattern."""

    def __init__(
        self,
        config: DecisionConfig,
        get_remembered: Callable[[tuple], Any] | None = None,
        remember_choice: Callable[[tuple, str, Any], None] | None = None,
    ):
        """Initialize decision manager.

        Args:
            config: Decision configuration.
            get_remembered: Callback to get remembered choice for a key.
            remember_choice: Callback to remember a choice.
        """
        self.config = config
        self.get_remembered = get_remembered
        self.remember_choice = remember_choice

    def process_issues(
        self,
        issues: dict[tuple, list[int]],
        display_issue: Callable[[tuple, list[int], int], None],
        apply_action: Callable[[str, tuple, list[int], Any], set[int]],
        summary_message: str | None = None,
        extra_input_handler: Callable[[str], Any] | None = None,
    ) -> set[int]:
        """Process issues with remember/prompt/apply flow.

        Args:
            issues: Dict mapping issue keys to affected line indices.
            display_issue: Callback to display issue details.
                          Args: (key, indices, count)
            apply_action: Callback to apply an action.
                         Args: (action, key, indices, extra_data)
                         Returns: set of indices to reject
            summary_message: Optional message to show before processing.
            extra_input_handler: Optional handler for actions needing extra input.
                                Args: (action) -> extra_data or None

        Returns:
            Set of line indices to reject.
        """
        if not issues:
            return set()

        if summary_message:
            console.warning(summary_message)
            print()

        reject_indices: set[int] = set()

        for key, indices in issues.items():
            count = len(indices)

            # Check for remembered choice
            if self.get_remembered:
                remembered = self.get_remembered(key)
                if remembered:
                    action, extra_data = self._parse_remembered(remembered)
                    self._log_remembered(action, key, count)
                    reject_indices.update(apply_action(action, key, indices, extra_data))
                    continue

            # Display issue
            display_issue(key, indices, count)

            # Prompt for action
            action, extra_data = self._prompt_user(extra_input_handler)

            # Remember choice
            if self.remember_choice and action:
                self.remember_choice(key, action, extra_data)

            # Log choice
            self._log_choice(action, key)

            # Apply action
            if action:
                reject_indices.update(apply_action(action, key, indices, extra_data))

            print()

        return reject_indices

    def _parse_remembered(self, remembered: Any) -> tuple[str, Any]:
        """Parse remembered choice into (action, extra_data)."""
        if isinstance(remembered, tuple):
            return remembered[0], remembered[1] if len(remembered) > 1 else None
        return remembered, None

    def _prompt_user(
        self, extra_input_handler: Callable[[str], Any] | None = None
    ) -> tuple[str | None, Any]:
        """Prompt user and return (action, extra_data)."""
        while True:
            response = input(self.config.get_prompt_text()).strip().lower()
            action = self.config.parse_response(response)

            if action is None:
                console.error(f"  Invalid choice. Enter one of: {', '.join(o.key.upper() for o in self.config.options)}")
                continue

            # Handle actions that need extra input
            extra_data = None
            if extra_input_handler:
                extra_data = extra_input_handler(action)
                if extra_data is False:  # Signal to re-prompt
                    continue

            return action, extra_data

    def _log_remembered(self, action: str, key: tuple, count: int) -> None:
        """Log auto-applied remembered choice."""
        key_str = self._format_key(key)
        console.info(f"  Auto-{action} (remembered): {key_str} ({count}x)")
        logging.log_user_choice(self.config.name, key_str, f"{action} (remembered)")

    def _log_choice(self, action: str | None, key: tuple) -> None:
        """Log user choice."""
        if action:
            key_str = self._format_key(key)
            logging.log_user_choice(self.config.name, key_str, action)

    def _format_key(self, key: tuple) -> str:
        """Format key tuple for display."""
        if len(key) == 2:
            return f'"{key[0]}" by {key[1]}'
        elif len(key) == 3:
            return f'"{key[0]}" by {key[1]} ({key[2]})'
        return str(key)


# Pre-configured decision types

ARTIST_TITLE_CONFIG = DecisionConfig(
    name="artist_title",
    options=[
        Option("y", ["yes"], "Yes fix", "fix", is_default=True),
        Option("n", ["no"], "No skip", "skip"),
        Option("x", ["reject"], "Reject", "reject"),
    ],
)

LONG_TIME_CONFIG = DecisionConfig(
    name="long_time",
    options=[
        Option("a", ["accept"], "Accept", "accept", is_default=True),
        Option("r", ["reject"], "Reject", "reject"),
        Option("e", ["edit"], "Edit", "edit"),
    ],
)

DUPLICATE_CONFIG = DecisionConfig(
    name="duplicate",
    options=[
        Option("k", ["keep"], "Keep all", "keep", is_default=True),
        Option("r", ["reject"], "Reject duplicates", "reject"),
    ],
)
