"""Bridge from Python to the Excel host.

The kernel installs a bridge: ``async bridge(method, params_json, token) -> result_json``.
In Pyodide it wraps a JavaScript function returning a Promise; in tests it is
a plain coroutine function.
"""

from __future__ import annotations

import json
from typing import Any, Awaitable, Callable

from ._context import require_command

Bridge = Callable[[str, str, str], Awaitable[str]]

_bridge: Bridge | None = None


def install_bridge(bridge: Bridge | None) -> None:
    global _bridge
    _bridge = bridge


class HostError(RuntimeError):
    def __init__(self, code: str, message: str):
        super().__init__(f"{code}: {message}")
        self.code = code


async def request(method: str, params: dict[str, Any]) -> Any:
    token = require_command(method)
    if _bridge is None:
        raise HostError("NO_HOST", "xlfly is not running inside Excel")
    try:
        raw = await _bridge(method, json.dumps(params), token)
    except Exception as exc:  # JavaScript errors surface as JsException in Pyodide
        message = str(exc)
        code = message.split(":", 1)[0] if ":" in message else "HOST_ERROR"
        raise HostError(code.strip(), message) from None
    return json.loads(raw) if raw else None
