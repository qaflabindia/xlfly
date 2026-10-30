"""Workbook API for commands and event handlers. All calls are bulk and async."""

from __future__ import annotations

import base64
import io
from typing import Any

from . import _host
from ._types import from_wire, to_wire

CHART_KINDS = frozenset({"line", "scatter", "column", "bar", "area", "pie", "histogram", "box"})


class Range:
    """A workbook range, e.g. ``Range("Data!A1:D100")`` or ``Range("A1")`` on the active sheet."""

    def __init__(self, address: str):
        if not isinstance(address, str) or not address.strip():
            raise ValueError("address must be a non-empty string")
        self.address = address.strip()

    def __repr__(self) -> str:
        return f"Range({self.address!r})"

    async def read(self) -> list[list[Any]]:
        """Values as a 2-D list. Errors become XLError, empty cells None."""
        result = await _host.request("range_read", {"address": self.address})
        return from_wire(result["values"])

    async def to_frame(self, header: bool = True):
        """Values as a pandas DataFrame."""
        import pandas as pd

        rows = await self.read()
        if header and rows:
            return pd.DataFrame(rows[1:], columns=[str(c) for c in rows[0]])
        return pd.DataFrame(rows)

    async def write(self, values: Any, header: bool = True) -> str:
        """Write values starting at the range's top-left cell. Returns the written address."""
        wire = to_wire(values, header=header)
        if not isinstance(wire, list):
            wire = [[wire]]
        result = await _host.request("range_write", {"address": self.address, "values": wire})
        return result["address"]

    async def from_frame(self, frame: Any, header: bool = True) -> str:
        return await self.write(frame, header=header)


async def call(name: str, *args: Any) -> Any:
    """Call an Excel worksheet function with values, e.g. ``await call("SUM", 1, 2)``."""
    result = await _host.request("excel_call", {"name": name, "args": [to_wire(a) for a in args]})
    return from_wire(result["value"])


async def chart(kind: str, source: str, anchor: str, title: str | None = None) -> str:
    """Insert a native, editable Excel chart over ``source``. Returns the chart name."""
    if kind not in CHART_KINDS:
        raise ValueError(f"unknown chart kind {kind!r}; expected one of {sorted(CHART_KINDS)}")
    result = await _host.request(
        "chart_native", {"kind": kind, "source": source, "anchor": anchor, "title": title}
    )
    return result["name"]


def _figure_png(figure: Any, dpi: int) -> bytes:
    if isinstance(figure, (bytes, bytearray)):
        return bytes(figure)
    savefig = getattr(figure, "savefig", None)
    if savefig is None:
        raise TypeError("image() expects a matplotlib Figure or PNG bytes")
    buffer = io.BytesIO()
    savefig(buffer, format="png", dpi=dpi, bbox_inches="tight")
    return buffer.getvalue()


async def image(figure: Any, anchor: str, name: str | None = None, dpi: int = 144) -> str:
    """Insert a matplotlib figure (or PNG bytes) as a picture at ``anchor``.

    If a picture with the same ``name`` exists it is replaced.
    """
    png = _figure_png(figure, dpi)
    if not png.startswith(b"\x89PNG"):
        raise ValueError("image data is not a PNG")
    result = await _host.request(
        "image_insert",
        {"base64": base64.b64encode(png).decode("ascii"), "anchor": anchor, "name": name},
    )
    return result["name"]


async def log(message: str) -> None:
    """Write a line to the xlfly task pane log."""
    await _host.request("log", {"message": str(message)})
