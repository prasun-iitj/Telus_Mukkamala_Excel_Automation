"""Chart display helpers — normalize zero values for chart data."""

from __future__ import annotations

from typing import Iterable, Optional


def normalize_chart_value(value: Optional[float]) -> Optional[float]:
    """Use integers for whole-number chart values (0.0 → 0)."""
    if value is None:
        return None
    if isinstance(value, float) and value == int(value):
        return int(value)
    if isinstance(value, int):
        return value
    return float(value)


def normalize_chart_series(values: Iterable[Optional[float]]) -> list[Optional[float]]:
    """Normalize a full series for chart replacement."""
    return [normalize_chart_value(v) for v in values]


def chart_series_number_format(values: Iterable[Optional[float]]) -> Optional[str]:
    """Return '0' for whole-number series; leave decimals/percentages unchanged."""
    nums = [v for v in values if isinstance(v, (int, float))]
    if not nums:
        return None
    if all(float(v) == int(float(v)) for v in nums):
        return "0"
    return None


def count_chart_value(value: Optional[float]) -> Optional[float]:
    """Use None for zero counts so overlapping '0' data labels are not drawn."""
    if value is None:
        return None
    if isinstance(value, (int, float)) and float(value) == 0.0:
        return None
    return normalize_chart_value(value)


def availability_chart_value(value: Optional[float]) -> Optional[float]:
    """Availability charts use a 95–100% axis; 0% is off-scale — treat as no data."""
    if value is None:
        return None
    if isinstance(value, (int, float)) and float(value) == 0.0:
        return None
    return normalize_chart_value(value)


def add_chart_series(chart_data, name: str, values: Iterable[Optional[float]]) -> None:
    """Add a series with safe value normalization and optional integer format."""
    normalized = normalize_chart_series(values)
    fmt = chart_series_number_format(normalized)
    if fmt:
        chart_data.add_series(name, normalized, number_format=fmt)
    else:
        chart_data.add_series(name, normalized)
