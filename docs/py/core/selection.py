"""Percentage maths for First (no vendor) and Second selection."""

FIELDS = ("total", "sensible", "flow", "power")


def factor(pct, direction):
    p = float(pct or 0) / 100.0
    return 1.0 - p if str(direction).lower().startswith("dec") else 1.0 + p


def scale(value, pct, direction, apply=True):
    if value is None:
        return None
    if not apply:
        return value
    return round(value * factor(pct, direction), 4)


def apply_all(values, pct, direction, apply_to=None):
    """values: dict total/sensible/flow/power -> value. apply_to: dict of bools (default all)."""
    apply_to = apply_to or {}
    return {k: scale(values.get(k), pct, direction, apply_to.get(k, True)) for k in FIELDS}
