import xlfly


def numbers(value):
    rows = value if isinstance(value, list) else [[value]]
    for row in rows:
        cells = row if isinstance(row, list) else [row]
        for cell in cells:
            if isinstance(cell, (int, float)) and not isinstance(cell, bool):
                yield cell


@xlfly.func
def scale(values, factor):
    """Scale the numeric values and return one column."""
    return [[value * factor] for value in numbers(values)]


@xlfly.command
async def summary():
    """Read A1:A4 and write a small summary on the active sheet."""
    rows = await xlfly.Range("A1:A4").read()
    values = list(numbers(rows))
    if not values:
        raise ValueError("A1:A4 contains no numbers")
    report = [
        ["Metric", "Value"],
        ["Count", len(values)],
        ["Total", sum(values)],
        ["Average", sum(values) / len(values)],
    ]
    await xlfly.Range("D1").write(report)
    await xlfly.chart(
        "column", "D1:E4", "G2", title="Input summary"
    )
    await xlfly.log("Summary written to D1:E4")
