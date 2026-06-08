"""Core data models for the Excel-to-PPT automation pipeline."""

from dataclasses import dataclass, field
from typing import Optional, Union

CellValue = Union[str, int, float, None]


@dataclass
class TeamMetrics:
    """Metrics for a single team within a month block."""
    team_name: str
    values: dict[str, CellValue]  # metric_name -> value


@dataclass
class MonthData:
    """Data for a single month across all teams."""
    month_name: str  # e.g., "February"
    teams: list[TeamMetrics]


@dataclass
class SectionData:
    """Data for a complete section (e.g., On-Time Delivery) across all months."""
    section_name: str  # e.g., "On-Time and Planned Delivery"
    sheet_name: str  # e.g., "Master-Apps (DoNotChange)"
    months: list[MonthData]
    column_headers: list[str]  # metric column names


@dataclass
class DataModel:
    """Complete data model containing all sections from both Master sheets."""
    apps_delivery: SectionData
    apps_quality: SectionData
    apps_rollout: SectionData
    apps_product_health: SectionData
    platforms_delivery: SectionData
    platforms_quality: SectionData
    platforms_rollout: SectionData
    platforms_product_health: SectionData
    reporting_month: str  # auto-detected or from --month arg


@dataclass
class SeriesSpec:
    """A single data series for a chart."""
    name: str  # series label
    values: list[Optional[float]]  # one value per category, None for missing


@dataclass
class ChartDataSpec:
    """Specification for chart data updates."""
    categories: list[str]  # x-axis labels (month names)
    series: list[SeriesSpec]


@dataclass
class TableDataSpec:
    """Specification for table data updates."""
    headers: list[str]
    rows: list[list[CellValue]]


@dataclass
class ValidationIssue:
    """A single validation issue (error or warning)."""
    level: str  # "error" or "warning"
    cell_ref: str  # e.g., "Master-Apps!D15"
    message: str


@dataclass
class ValidationReport:
    """Accumulated validation results from the pipeline."""
    errors: list[ValidationIssue] = field(default_factory=list)
    warnings: list[ValidationIssue] = field(default_factory=list)

    @property
    def has_fatal_errors(self) -> bool:
        return len(self.errors) > 0
