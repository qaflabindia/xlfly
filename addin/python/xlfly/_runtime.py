"""Kernel entry points called by the JavaScript worker.

Every entry point takes and returns JSON strings so that no Python or
JavaScript proxies cross the boundary.
"""

from __future__ import annotations

import ast
import inspect
import json
import sys
import traceback
import types
from collections import OrderedDict
from typing import Any

import xlfly

from ._context import COMMAND, FUNCTION, running
from ._registry import REGISTRY
from ._types import XLError, from_wire, to_wire, unwrap_single

USER_PACKAGE = "xlfly_user"
_CODE_CACHE: "OrderedDict[str, tuple[Any, Any]]" = OrderedDict()
_CODE_CACHE_SIZE = 256


def _error(exc: BaseException, include_trace: bool = False) -> dict[str, Any]:
    if isinstance(exc, XLError):
        return {"code": exc.code, "message": exc.message or exc.code}
    message = f"{type(exc).__name__}: {exc}"
    error: dict[str, Any] = {"code": "#VALUE!", "message": message}
    if include_trace:
        error["trace"] = "".join(traceback.format_exception(type(exc), exc, exc.__traceback__, limit=-8))
    return error


def describe() -> dict[str, Any]:
    def info(fn: Any) -> dict[str, Any]:
        doc = inspect.getdoc(fn) or ""
        try:
            signature = str(inspect.signature(fn))
        except (TypeError, ValueError):
            signature = "(...)"
        return {"signature": signature, "doc": doc.splitlines()[0] if doc else ""}

    return {
        "functions": {name: info(fn) for name, fn in sorted(REGISTRY.functions.items())},
        "commands": {name: info(fn) for name, fn in sorted(REGISTRY.commands.items())},
        "events": sorted(REGISTRY.events),
    }


def load_modules(modules_json: str) -> str:
    """Replace all user modules. Input: [{"name": str, "source": str}, ...]."""
    modules = json.loads(modules_json)
    REGISTRY.clear()
    _CODE_CACHE.clear()
    for name in [n for n in sys.modules if n == USER_PACKAGE or n.startswith(USER_PACKAGE + ".")]:
        del sys.modules[name]
    package = types.ModuleType(USER_PACKAGE)
    package.__path__ = []
    sys.modules[USER_PACKAGE] = package

    errors = []
    for module in modules:
        name = module["name"]
        if not name.isidentifier():
            errors.append({"module": name, "code": "#NAME?", "message": "module name must be a Python identifier"})
            continue
        qualified = f"{USER_PACKAGE}.{name}"
        mod = types.ModuleType(qualified)
        mod.__file__ = f"<xlfly:{name}>"
        sys.modules[qualified] = mod
        setattr(package, name, mod)
        try:
            exec(compile(module["source"], mod.__file__, "exec"), mod.__dict__)
        except BaseException as exc:  # a user module must never take the kernel down
            if isinstance(exc, (KeyboardInterrupt, SystemExit)):
                exc = RuntimeError(f"module raised {type(exc).__name__}")
            errors.append({"module": name, **_error(exc, include_trace=True)})
    return json.dumps({**describe(), "errors": errors})


def _compile_cell(code: str) -> tuple[Any, Any]:
    cached = _CODE_CACHE.get(code)
    if cached is not None:
        _CODE_CACHE.move_to_end(code)
        return cached
    tree = ast.parse(code, filename="<cell>", mode="exec")
    last = None
    if tree.body and isinstance(tree.body[-1], ast.Expr):
        last = compile(ast.Expression(tree.body.pop().value), "<cell>", "eval")
    body = compile(tree, "<cell>", "exec")
    _CODE_CACHE[code] = (body, last)
    if len(_CODE_CACHE) > _CODE_CACHE_SIZE:
        _CODE_CACHE.popitem(last=False)
    return body, last


def evaluate(code: str, args: list[Any]) -> Any:
    """Run cell code; the value of a trailing expression is the result."""
    body, last = _compile_cell(code)
    namespace: dict[str, Any] = {
        "__name__": "__xlfly_cell__",
        "xlfly": xlfly,
        "user": sys.modules.get(USER_PACKAGE),
        "args": args,
    }
    for index, value in enumerate(args, start=1):
        namespace[f"arg{index}"] = value
    exec(body, namespace)
    return eval(last, namespace) if last is not None else namespace.get("result")


def _invoke(item: dict[str, Any]) -> Any:
    args = [unwrap_single(from_wire(a)) for a in item.get("args", [])]
    if item["kind"] == "eval":
        value = evaluate(item["code"], args)
    elif item["kind"] == "call":
        fn = REGISTRY.functions.get(str(item["name"]).upper())
        if fn is None:
            raise XLError("#NAME?", f"no xlfly function named {item['name']!r}")
        value = fn(*args)
    else:
        raise ValueError(f"unknown call kind {item['kind']!r}")
    if inspect.isawaitable(value):
        close = getattr(value, "close", None)
        if close:
            close()
        raise TypeError("worksheet functions must be synchronous; use @xlfly.command for async work")
    return value


def call_functions(batch_json: str) -> str:
    """Evaluate a batch of worksheet calls in function context."""
    results = []
    with running(FUNCTION):
        for item in json.loads(batch_json):
            try:
                results.append({"id": item["id"], "ok": True, "value": to_wire(_invoke(item))})
            except BaseException as exc:
                if isinstance(exc, (KeyboardInterrupt, SystemExit)):
                    exc = RuntimeError(f"function raised {type(exc).__name__}")
                results.append({"id": item["id"], "ok": False, "error": _error(exc)})
    return json.dumps(results)


async def run_command(name: str, token: str) -> str:
    fn = REGISTRY.commands.get(name.upper())
    if fn is None:
        return json.dumps({"ok": False, "error": {"code": "NOT_FOUND", "message": f"no command named {name!r}"}})
    with running(COMMAND, token):
        try:
            value = fn()
            if inspect.isawaitable(value):
                value = await value
            return json.dumps({"ok": True, "value": to_wire(value)})
        except BaseException as exc:
            if isinstance(exc, (KeyboardInterrupt, SystemExit)):
                exc = RuntimeError(f"command raised {type(exc).__name__}")
            return json.dumps({"ok": False, "error": _error(exc, include_trace=True)})


async def dispatch_event(event_json: str, token: str) -> str:
    event = json.loads(event_json)
    errors = []
    with running(COMMAND, token):
        for handler in list(REGISTRY.events.get(event.get("type", ""), [])):
            try:
                value = handler(event)
                if inspect.isawaitable(value):
                    await value
            except BaseException as exc:
                if isinstance(exc, (KeyboardInterrupt, SystemExit)):
                    exc = RuntimeError(f"handler raised {type(exc).__name__}")
                errors.append({"handler": getattr(handler, "__name__", "?"), **_error(exc, include_trace=True)})
    return json.dumps({"ok": not errors, "errors": errors})
