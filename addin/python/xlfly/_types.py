"""Excel <-> Python value conversion.

Wire format (JSON, shared with the TypeScript host):
  number | string | bool | null       scalar cell value
  {"$err": "#N/A"}                    Excel error
  [[...], [...]]                      2-D array, row-major
"""

from __future__ import annotations

import datetime as _dt
import math
from typing import Any

ERROR_CODES = frozenset(
    {"#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A", "#GETTING_DATA", "#SPILL!", "#CALC!", "#BUSY!"}
)

_EXCEL_EPOCH = _dt.datetime(1899, 12, 30)


class XLError(Exception):
    """An Excel error value such as #N/A. Can be returned or raised."""

    def __init__(self, code: str = "#VALUE!", message: str = ""):
        if code not in ERROR_CODES:
            raise ValueError(f"unknown Excel error code: {code!r}")
        super().__init__(message or code)
        self.code = code
        self.message = message

    def __eq__(self, other: object) -> bool:
        return isinstance(other, XLError) and other.code == self.code

    def __hash__(self) -> int:
        return hash(self.code)

    def __repr__(self) -> str:
        return f"XLError({self.code!r})"


def to_excel_date(value: _dt.date | _dt.datetime) -> float:
    """Convert a date or datetime to an Excel serial number (1900 date system)."""
    if not isinstance(value, _dt.datetime):
        value = _dt.datetime(value.year, value.month, value.day)
    if value.tzinfo is not None:
        value = value.replace(tzinfo=None)
    delta = value - _EXCEL_EPOCH
    return delta.days + delta.seconds / 86400 + delta.microseconds / 86_400_000_000


def from_excel_date(serial: float) -> _dt.datetime:
    """Convert an Excel serial number (1900 date system) to a datetime."""
    return _EXCEL_EPOCH + _dt.timedelta(days=float(serial))


# ---------------------------------------------------------------- wire -> Python

def from_wire(value: Any) -> Any:
    if isinstance(value, dict):
        if "$err" in value:
            return XLError(value["$err"])
        raise ValueError(f"unexpected object on the wire: {value!r}")
    if isinstance(value, list):
        return [from_wire(v) for v in value]
    return value


# ---------------------------------------------------------------- Python -> wire

def _scalar_to_wire(value: Any) -> Any:
    if value is None:
        return None
    if isinstance(value, XLError):
        return {"$err": value.code}
    if isinstance(value, bool):
        return value
    if isinstance(value, (int, float)):
        f = float(value)
        if math.isnan(f) or math.isinf(f):
            return {"$err": "#NUM!"}
        return int(value) if isinstance(value, int) and abs(value) < 2**53 else f
    if isinstance(value, str):
        return value
    if isinstance(value, (_dt.datetime, _dt.date)):
        return to_excel_date(value)
    # numpy / pandas scalars without importing them
    item = getattr(value, "item", None)
    if callable(item) and getattr(value, "shape", None) == ():
        return _scalar_to_wire(item())
    type_name = type(value).__name__
    if type_name in ("NaTType", "NAType"):
        return None
    if type_name == "Timestamp":
        return to_excel_date(value.to_pydatetime())
    return str(value)


def _rectangular(rows: list[list[Any]]) -> list[list[Any]]:
    width = max((len(r) for r in rows), default=0)
    if width == 0:
        return [[None]]
    return [r + [None] * (width - len(r)) for r in rows]


def to_wire(value: Any, header: bool = True) -> Any:
    """Convert a Python result to the wire format.

    Scalars stay scalars. 1-D sequences become a single column.
    DataFrames include a header row when ``header`` is true.
    """
    type_name = type(value).__name__
    if type_name == "DataFrame":
        rows = value.values.tolist()
        body = [[_scalar_to_wire(v) for v in r] for r in rows]
        if header:
            body.insert(0, [_scalar_to_wire(c) for c in value.columns])
        return _rectangular(body) if body else [[None]]
    if type_name == "Series":
        return [[_scalar_to_wire(v)] for v in value.tolist()] or [[None]]
    if type_name == "ndarray":
        if value.ndim == 0:
            return _scalar_to_wire(value.item())
        if value.ndim == 1:
            return [[_scalar_to_wire(v)] for v in value.tolist()] or [[None]]
        if value.ndim == 2:
            return _rectangular([[_scalar_to_wire(v) for v in r] for r in value.tolist()])
        raise ValueError("arrays with more than 2 dimensions cannot be returned to Excel")
    if isinstance(value, (list, tuple)):
        if not value:
            return [[None]]
        if all(isinstance(r, (list, tuple)) for r in value):
            return _rectangular([[_scalar_to_wire(v) for v in r] for r in value])
        return [[_scalar_to_wire(v)] for v in value]
    return _scalar_to_wire(value)


def unwrap_single(value: Any) -> Any:
    """Return the scalar inside a 1x1 array, otherwise the value unchanged."""
    if isinstance(value, list) and len(value) == 1 and isinstance(value[0], list) and len(value[0]) == 1:
        return value[0][0]
    return value
