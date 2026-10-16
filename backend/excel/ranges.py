"""Conversions for values returned by Excel COM ranges."""

def range_to_rows(value) -> list[list]:
    """
    Converts an Excel COM Range.Value result into normal Python rows.
    """
    if value is None:
        return []

    if not isinstance(value, tuple):
        return [[value]]

    if (
        value
        and not isinstance(
            value[0],
            tuple,
        )
    ):
        return [list(value)]

    return [
        list(row)
        for row in value
    ]
