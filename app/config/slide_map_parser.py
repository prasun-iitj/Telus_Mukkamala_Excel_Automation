"""Parser and validator for slide_map.yaml configuration."""

from __future__ import annotations

import os
from dataclasses import dataclass, field
from typing import Any, Optional

import yaml

from app.exceptions import SlideMapError

# ---------------------------------------------------------------------------
# Supported slide types
# ---------------------------------------------------------------------------
VALID_SLIDE_TYPES = frozenset(
    ["title_slide", "section_divider", "summary_table", "bar_chart", "table", "text_block"]
)

# Valid field values per update
VALID_FIELDS = frozenset(["text", "chart", "table"])

# Valid static source types
VALID_SOURCE_TYPES = frozenset(["static", "month_year"])


# ---------------------------------------------------------------------------
# Config dataclasses
# ---------------------------------------------------------------------------
@dataclass
class SeriesConfig:
    """A single chart data series reference."""

    name: str
    column: str


@dataclass
class SourceConfig:
    """Data source reference for a shape update."""

    # For static / month_year text sources
    type: Optional[str] = None
    value: Optional[str] = None

    # For data-driven sources (chart, table, text_block)
    sheet: Optional[str] = None
    section: Optional[str] = None
    month: Optional[str] = None
    columns: Optional[list[str]] = None
    categories: Optional[list[str]] = None
    series: Optional[list[SeriesConfig]] = None
    format: Optional[str] = None
    include_teams: Optional[bool] = None


@dataclass
class UpdateConfig:
    """A single shape update within a slide."""

    shape_name: str
    field: str  # "text", "chart", or "table"
    source: SourceConfig
    shape_index: Optional[int] = None


@dataclass
class SlideConfig:
    """Configuration for one slide."""

    slide_index: int
    type: str
    updates: list[UpdateConfig]


@dataclass
class SlideMapConfig:
    """Top-level configuration containing all slides."""

    slides: list[SlideConfig] = field(default_factory=list)


# ---------------------------------------------------------------------------
# Parsing helpers
# ---------------------------------------------------------------------------

def _parse_source(raw: dict[str, Any], context: str) -> SourceConfig:
    """Parse a source dict into a SourceConfig."""
    if not isinstance(raw, dict):
        raise SlideMapError(f"{context}: 'source' must be a mapping")

    series_raw = raw.get("series")
    series_list: Optional[list[SeriesConfig]] = None
    if series_raw is not None:
        if not isinstance(series_raw, list):
            raise SlideMapError(f"{context}: 'series' must be a list")
        series_list = []
        for i, s in enumerate(series_raw):
            if not isinstance(s, dict):
                raise SlideMapError(f"{context}: series[{i}] must be a mapping")
            if "name" not in s or "column" not in s:
                raise SlideMapError(
                    f"{context}: series[{i}] requires 'name' and 'column'"
                )
            series_list.append(SeriesConfig(name=s["name"], column=str(s["column"])))

    columns = raw.get("columns")
    if columns is not None:
        columns = [str(c) for c in columns]

    categories = raw.get("categories")
    if categories is not None:
        categories = [str(c) for c in categories]

    return SourceConfig(
        type=raw.get("type"),
        value=raw.get("value"),
        sheet=raw.get("sheet"),
        section=raw.get("section"),
        month=raw.get("month"),
        columns=columns,
        categories=categories,
        series=series_list,
        format=raw.get("format"),
        include_teams=raw.get("include_teams"),
    )


def _parse_update(raw: dict[str, Any], slide_ctx: str, idx: int) -> UpdateConfig:
    """Parse a single update entry."""
    ctx = f"{slide_ctx}, update[{idx}]"

    if not isinstance(raw, dict):
        raise SlideMapError(f"{ctx}: update entry must be a mapping")

    shape_name = raw.get("shape_name")
    if not shape_name:
        raise SlideMapError(f"{ctx}: 'shape_name' is required")

    fld = raw.get("field")
    if not fld:
        raise SlideMapError(f"{ctx}: 'field' is required")
    if fld not in VALID_FIELDS:
        raise SlideMapError(
            f"{ctx}: invalid field '{fld}', must be one of {sorted(VALID_FIELDS)}"
        )

    source_raw = raw.get("source")
    if source_raw is None:
        raise SlideMapError(f"{ctx}: 'source' is required")

    source = _parse_source(source_raw, ctx)
    shape_index = raw.get("shape_index")

    return UpdateConfig(
        shape_name=str(shape_name),
        field=str(fld),
        source=source,
        shape_index=shape_index,
    )


def _validate_update_for_type(update: UpdateConfig, slide_type: str, ctx: str) -> None:
    """Validate that an update has the required source fields for its slide type."""
    src = update.source

    if slide_type in ("title_slide", "section_divider"):
        if update.field != "text":
            raise SlideMapError(
                f"{ctx}: '{slide_type}' updates must have field='text'"
            )
        if src.type is None:
            raise SlideMapError(
                f"{ctx}: '{slide_type}' text source requires 'type' "
                f"(one of {sorted(VALID_SOURCE_TYPES)})"
            )
        if src.type not in VALID_SOURCE_TYPES:
            raise SlideMapError(
                f"{ctx}: invalid source type '{src.type}', "
                f"must be one of {sorted(VALID_SOURCE_TYPES)}"
            )
        if src.type == "static" and src.value is None:
            raise SlideMapError(
                f"{ctx}: static source requires 'value'"
            )

    elif update.field == "chart":
        if src.sheet is None:
            raise SlideMapError(f"{ctx}: chart source requires 'sheet'")
        if src.section is None:
            raise SlideMapError(f"{ctx}: chart source requires 'section'")
        if src.categories is None:
            raise SlideMapError(f"{ctx}: chart source requires 'categories'")
        if src.series is None or len(src.series) == 0:
            raise SlideMapError(f"{ctx}: chart source requires at least one 'series'")

    elif update.field == "table":
        if src.sheet is None:
            raise SlideMapError(f"{ctx}: table source requires 'sheet'")
        if src.section is None:
            raise SlideMapError(f"{ctx}: table source requires 'section'")
        if src.month is None:
            raise SlideMapError(f"{ctx}: table source requires 'month'")
        if src.columns is None:
            raise SlideMapError(f"{ctx}: table source requires 'columns'")

    elif update.field == "text" and slide_type in ("text_block", "bar_chart", "summary_table", "table"):
        # Data-driven text blocks
        if src.type in VALID_SOURCE_TYPES:
            # Static/month_year text is allowed anywhere
            return
        if src.sheet is None:
            raise SlideMapError(f"{ctx}: text source requires 'sheet'")
        if src.section is None:
            raise SlideMapError(f"{ctx}: text source requires 'section'")
        if src.month is None:
            raise SlideMapError(f"{ctx}: text source requires 'month'")
        if src.columns is None:
            raise SlideMapError(f"{ctx}: text source requires 'columns'")


def _parse_slide(raw: dict[str, Any], idx: int) -> SlideConfig:
    """Parse and validate a single slide entry."""
    ctx = f"slides[{idx}]"

    if not isinstance(raw, dict):
        raise SlideMapError(f"{ctx}: slide entry must be a mapping")

    slide_index = raw.get("slide_index")
    if slide_index is None:
        raise SlideMapError(f"{ctx}: 'slide_index' is required")
    if not isinstance(slide_index, int):
        raise SlideMapError(f"{ctx}: 'slide_index' must be an integer")

    slide_type = raw.get("type")
    if not slide_type:
        raise SlideMapError(f"{ctx}: 'type' is required")
    if slide_type not in VALID_SLIDE_TYPES:
        raise SlideMapError(
            f"{ctx}: unknown slide type '{slide_type}', "
            f"must be one of {sorted(VALID_SLIDE_TYPES)}"
        )

    updates_raw = raw.get("updates")
    if updates_raw is None:
        raise SlideMapError(f"{ctx}: 'updates' is required")
    if not isinstance(updates_raw, list):
        raise SlideMapError(f"{ctx}: 'updates' must be a list")

    updates: list[UpdateConfig] = []
    for i, u in enumerate(updates_raw):
        update = _parse_update(u, ctx, i)
        _validate_update_for_type(update, slide_type, f"{ctx}, update[{i}]")
        updates.append(update)

    return SlideConfig(
        slide_index=slide_index,
        type=slide_type,
        updates=updates,
    )


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def load_slide_map(path: str) -> SlideMapConfig:
    """Load, parse, and validate a slide_map.yaml file.

    Args:
        path: Path to the YAML configuration file.

    Returns:
        A fully validated SlideMapConfig.

    Raises:
        SlideMapError: If the file cannot be read, parsed, or fails validation.
    """
    if not os.path.isfile(path):
        raise SlideMapError(f"Slide map file not found: {path}")

    try:
        with open(path, "r", encoding="utf-8") as fh:
            data = yaml.safe_load(fh)
    except yaml.YAMLError as exc:
        raise SlideMapError(f"YAML parse error: {exc}") from exc

    if not isinstance(data, dict):
        raise SlideMapError("Slide map must be a YAML mapping at the top level")

    slides_raw = data.get("slides")
    if slides_raw is None:
        raise SlideMapError("Top-level 'slides' key is required")
    if not isinstance(slides_raw, list):
        raise SlideMapError("'slides' must be a list")

    slides: list[SlideConfig] = []
    for i, s in enumerate(slides_raw):
        slides.append(_parse_slide(s, i))

    return SlideMapConfig(slides=slides)
