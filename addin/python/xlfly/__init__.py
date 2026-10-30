"""xlfly: local Python for Excel.

Worksheet functions::

    import xlfly

    @xlfly.func
    def growth(values, rate):
        return [[v * (1 + rate)] for row in values for v in row]

    # In a cell: =XLFLY.CALL("GROWTH", A1:A10, 0.05)

Commands (run from the task pane, may read and write the workbook)::

    @xlfly.command
    async def report():
        df = await xlfly.Range("Data!A1:C100").to_frame()
        await xlfly.Range("Report!A1").write(df.describe().reset_index())
"""

from ._api import Range, call, chart, image, log
from ._context import ContextError
from ._host import HostError
from ._registry import command, func, on
from ._types import XLError, from_excel_date, to_excel_date

__version__ = "0.1.0"

__all__ = [
    "ContextError",
    "HostError",
    "Range",
    "XLError",
    "call",
    "chart",
    "command",
    "from_excel_date",
    "func",
    "image",
    "log",
    "on",
    "to_excel_date",
]
