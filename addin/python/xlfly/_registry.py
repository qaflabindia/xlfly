"""Registries filled by the decorators in user modules."""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any, Callable

EVENT_TYPES = frozenset({"sheet_change", "selection_change", "workbook_open"})


@dataclass
class Registry:
    functions: dict[str, Callable[..., Any]] = field(default_factory=dict)
    commands: dict[str, Callable[..., Any]] = field(default_factory=dict)
    events: dict[str, list[Callable[..., Any]]] = field(default_factory=dict)

    def clear(self) -> None:
        self.functions.clear()
        self.commands.clear()
        self.events.clear()


REGISTRY = Registry()


def _key(name: str) -> str:
    key = name.strip().upper()
    if not key or not all(ch.isalnum() or ch in "._" for ch in key):
        raise ValueError(f"invalid name: {name!r}")
    return key


def _decorator(target: dict[str, Callable[..., Any]], name_or_fn: Any):
    def register(fn: Callable[..., Any], name: str | None = None):
        target[_key(name or fn.__name__)] = fn
        return fn

    if callable(name_or_fn):
        return register(name_or_fn)
    return lambda fn: register(fn, name_or_fn)


def func(name_or_fn: Any = None):
    """Expose a function to worksheets: ``=XLFLY.CALL("NAME", args...)``.

    Worksheet functions must be synchronous and side-effect free.
    """
    return _decorator(REGISTRY.functions, name_or_fn)


def command(name_or_fn: Any = None):
    """Expose a command, run from the task pane. May be ``async`` and may use the workbook API."""
    return _decorator(REGISTRY.commands, name_or_fn)


def on(event: str):
    """Subscribe a handler to a workbook event: sheet_change, selection_change, workbook_open."""
    if event not in EVENT_TYPES:
        raise ValueError(f"unknown event {event!r}; expected one of {sorted(EVENT_TYPES)}")

    def register(fn: Callable[..., Any]):
        REGISTRY.events.setdefault(event, []).append(fn)
        return fn

    return register
