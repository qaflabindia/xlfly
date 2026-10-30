"""Execution contexts.

function  -- running a worksheet function; Excel is waiting on the result.
             Calls back into Excel are refused (they would deadlock and
             Excel forbids side effects from worksheet functions).
command   -- running a command or event handler; workbook access allowed.
"""

from __future__ import annotations

import contextvars
from contextlib import contextmanager
from typing import Iterator

FUNCTION = "function"
COMMAND = "command"

current_context: contextvars.ContextVar[str | None] = contextvars.ContextVar("xlfly_context", default=None)
current_token: contextvars.ContextVar[str | None] = contextvars.ContextVar("xlfly_token", default=None)


class ContextError(RuntimeError):
    """Raised when Excel is accessed from a context that does not allow it."""


@contextmanager
def running(kind: str, token: str | None = None) -> Iterator[None]:
    c = current_context.set(kind)
    t = current_token.set(token)
    try:
        yield
    finally:
        current_context.reset(c)
        current_token.reset(t)


def require_command(operation: str) -> str:
    kind = current_context.get()
    if kind != COMMAND:
        where = "a worksheet function" if kind == FUNCTION else "outside a command"
        raise ContextError(
            f"{operation} is not allowed from {where}; use @xlfly.command or an event handler"
        )
    token = current_token.get()
    if not token:
        raise ContextError("missing command token")
    return token
