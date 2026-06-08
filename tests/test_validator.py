"""Unit tests for the Validator class."""

from unittest.mock import MagicMock

import pytest

from app.exceptions import SheetNotFoundError
from app.models import (
    MonthData,
    SectionData,
    TeamMetrics,
    ValidationReport,
)
from app.validators.validator import Validator


# ------------------------------------------------------------------
# Helpers
# ------------------------------------------------------------------

def _make_reader(available_sheets: list[str]) -> MagicMock:
    """Create a mock ExcelReader whose get_sheet raises for missing sheets."""
    reader = MagicMock()

    def _get_sheet(name: str):
        if name not in available_sheets:
            raise SheetNotFoundError(name)
        return MagicMock()

    reader.get_sheet = MagicMock(side_effect=_get_sheet)
    return reader


def _section(
    section_name: str,
    months: list[MonthData],
    sheet_name: str = "Master-Apps (DoNotChange)",
    column_headers: list[str] | None = None,
) -> SectionData:
    return SectionData(
        section_name=section_name,
        sheet_name=sheet_name,
        months=months,
        column_headers=column_headers or [],
    )


# ------------------------------------------------------------------
# validate_workbook
# ------------------------------------------------------------------

class TestValidateWorkbook:
    def test_both_sheets_present(self):
        reader = _make_reader([
            "Master-Apps (DoNotChange)",
            "Master-Platforms (DoNotChange)",
        ])
        report = Validator().validate_workbook(reader)
        assert not report.has_fatal_errors
        assert len(report.errors) == 0

    def test_apps_sheet_missing(self):
        reader = _make_reader(["Master-Platforms (DoNotChange)"])
        report = Validator().validate_workbook(reader)
        assert report.has_fatal_errors
        assert len(report.errors) == 1
        assert "Master-Apps (DoNotChange)" in report.errors[0].message

    def test_platforms_sheet_missing(self):
        reader = _make_reader(["Master-Apps (DoNotChange)"])
        report = Validator().validate_workbook(reader)
        assert report.has_fatal_errors
        assert len(report.errors) == 1
        assert "Master-Platforms (DoNotChange)" in report.errors[0].message

    def test_both_sheets_missing(self):
        reader = _make_reader([])
        report = Validator().validate_workbook(reader)
        assert report.has_fatal_errors
        assert len(report.errors) == 2


# ------------------------------------------------------------------
# validate_section_data — non-numeric warnings
# ------------------------------------------------------------------

class TestValidateNonNumeric:
    def test_string_in_numeric_column_warns(self):
        data = _section(
            "On-Time and Planned Delivery",
            months=[
                MonthData(
                    month_name="January",
                    teams=[TeamMetrics(team_name="TeamA", values={"Track 1 delivered": "N/A"})],
                )
            ],
        )
        report = Validator().validate_section_data(data, "On-Time and Planned Delivery")
        assert len(report.warnings) == 1
        assert "Non-numeric" in report.warnings[0].message

    def test_numeric_value_no_warning(self):
        data = _section(
            "On-Time and Planned Delivery",
            months=[
                MonthData(
                    month_name="January",
                    teams=[TeamMetrics(team_name="TeamA", values={"Track 1 delivered": 42})],
                )
            ],
        )
        report = Validator().validate_section_data(data, "On-Time and Planned Delivery")
        # Only non-empty month, no warnings expected
        assert len(report.warnings) == 0


# ------------------------------------------------------------------
# validate_section_data — percentage range warnings
# ------------------------------------------------------------------

class TestValidatePercentage:
    def test_valid_decimal_percentage(self):
        data = _section(
            "On-Time and Planned Delivery",
            months=[
                MonthData(
                    month_name="Feb",
                    teams=[TeamMetrics(team_name="T", values={"% on-time Track 1": 0.85})],
                )
            ],
        )
        report = Validator().validate_section_data(data, "On-Time and Planned Delivery")
        assert len(report.warnings) == 0

    def test_valid_whole_number_percentage(self):
        data = _section(
            "On-Time and Planned Delivery",
            months=[
                MonthData(
                    month_name="Feb",
                    teams=[TeamMetrics(team_name="T", values={"% on-time Track 1": 85})],
                )
            ],
        )
        report = Validator().validate_section_data(data, "On-Time and Planned Delivery")
        assert len(report.warnings) == 0

    def test_negative_percentage_warns(self):
        data = _section(
            "On-Time and Planned Delivery",
            months=[
                MonthData(
                    month_name="Feb",
                    teams=[TeamMetrics(team_name="T", values={"% on-time Track 1": -0.5})],
                )
            ],
        )
        report = Validator().validate_section_data(data, "On-Time and Planned Delivery")
        pct_warnings = [w for w in report.warnings if "Percentage" in w.message]
        assert len(pct_warnings) == 1

    def test_percentage_over_100_warns(self):
        data = _section(
            "On-Time and Planned Delivery",
            months=[
                MonthData(
                    month_name="Feb",
                    teams=[TeamMetrics(team_name="T", values={"% on-time Track 1": 150})],
                )
            ],
        )
        report = Validator().validate_section_data(data, "On-Time and Planned Delivery")
        pct_warnings = [w for w in report.warnings if "Percentage" in w.message]
        assert len(pct_warnings) == 1

    def test_rollout_percentage_column_validated(self):
        data = _section(
            "Rollout Success",
            months=[
                MonthData(
                    month_name="Mar",
                    teams=[TeamMetrics(team_name="T", values={"rollout success percentage": -1})],
                )
            ],
        )
        report = Validator().validate_section_data(data, "Rollout Success")
        pct_warnings = [w for w in report.warnings if "Percentage" in w.message]
        assert len(pct_warnings) == 1

    def test_quality_section_no_percentage_columns(self):
        """Quality section has no designated percentage columns, so out-of-range values don't trigger pct warnings."""
        data = _section(
            "Quality and Dev Efficiency",
            months=[
                MonthData(
                    month_name="Mar",
                    teams=[TeamMetrics(team_name="T", values={"Total Defects": 200})],
                )
            ],
        )
        report = Validator().validate_section_data(data, "Quality and Dev Efficiency")
        assert len(report.warnings) == 0


# ------------------------------------------------------------------
# validate_section_data — all-null / all-zero month blocks
# ------------------------------------------------------------------

class TestValidateEmptyMonth:
    def test_all_none_month_warns(self):
        data = _section(
            "On-Time and Planned Delivery",
            months=[
                MonthData(
                    month_name="January",
                    teams=[
                        TeamMetrics(team_name="A", values={"col1": None, "col2": None}),
                        TeamMetrics(team_name="B", values={"col1": None, "col2": None}),
                    ],
                )
            ],
        )
        report = Validator().validate_section_data(data, "On-Time and Planned Delivery")
        empty_warnings = [w for w in report.warnings if "all-null or all-zero" in w.message]
        assert len(empty_warnings) == 1

    def test_all_zero_month_warns(self):
        data = _section(
            "On-Time and Planned Delivery",
            months=[
                MonthData(
                    month_name="January",
                    teams=[
                        TeamMetrics(team_name="A", values={"col1": 0, "col2": 0}),
                    ],
                )
            ],
        )
        report = Validator().validate_section_data(data, "On-Time and Planned Delivery")
        empty_warnings = [w for w in report.warnings if "all-null or all-zero" in w.message]
        assert len(empty_warnings) == 1

    def test_mixed_values_no_empty_warning(self):
        data = _section(
            "On-Time and Planned Delivery",
            months=[
                MonthData(
                    month_name="January",
                    teams=[
                        TeamMetrics(team_name="A", values={"col1": 5, "col2": None}),
                    ],
                )
            ],
        )
        report = Validator().validate_section_data(data, "On-Time and Planned Delivery")
        empty_warnings = [w for w in report.warnings if "all-null or all-zero" in w.message]
        assert len(empty_warnings) == 0
