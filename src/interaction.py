"""Interaction handler abstraction for user prompts (CLI or GUI)."""
from __future__ import annotations

from abc import ABC, abstractmethod
from typing import Any, Callable, TYPE_CHECKING

if TYPE_CHECKING:
    from .decisions import DecisionConfig


class InteractionHandler(ABC):
    """Abstract interface for user-facing prompts."""

    @abstractmethod
    def prompt_decision(
        self,
        config: DecisionConfig,
        key: Any = None,
        count: int = 0,
        extra_input_fn: Callable[[str], Any] | None = None,
    ) -> tuple[str, Any]:
        """Show a decision prompt. Returns (action, extra_data)."""
        ...

    @abstractmethod
    def prompt_string(self, message: str, default: str = "") -> str:
        """Prompt for a free-text value. Returns the entered string."""
        ...


class CLIInteractionHandler(InteractionHandler):
    """Interaction handler that wraps input() for terminal use."""

    def prompt_decision(
        self,
        config: DecisionConfig,
        key: Any = None,
        count: int = 0,
        extra_input_fn: Callable[[str], Any] | None = None,
    ) -> tuple[str, Any]:
        from . import output as console

        while True:
            response = input(config.get_prompt_text()).strip().lower()
            action = config.parse_response(response)

            if action is None:
                valid = ", ".join(o.key.upper() for o in config.options)
                console.error(f"  Invalid choice. Enter one of: {valid}")
                continue

            extra_data = None
            if extra_input_fn:
                extra_data = extra_input_fn(action)
                if extra_data is False:
                    continue

            return action, extra_data

    def prompt_string(self, message: str, default: str = "") -> str:
        return input(message).strip()
