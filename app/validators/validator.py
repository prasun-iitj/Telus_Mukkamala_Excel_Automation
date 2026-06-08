"""Validator for Excel workbook structure and section data quality."""

import logging
from typing import Union

from app.exceptions import SheetNotFoundError
from app.models import (
    CellValue,
    MonthData,
    SectionData,
    ValidationIssue,
    ValidationReport,
)
from app.readers.excel_reader import ExcelReader

logger = logging.getLogger(__name__)

REQUIRED_SHEETS = [
    "Master-Apps (DoNotChange)",
    "Master-Platforms (DoNotChange)",
]

# Percentage column names by section (case-insensitive matching)
DELIVERY_PERCENTAGE_COLUMNS = {
    "% on-time track 1",
    "% on-time track 2",
    "% on-time sustainability",
    "% on-time subcontracted",
}

ROLLOUT_PERCENTAGE_COLUMNS = {
    "rollout success percentage",
    "ytd success percentage",
}


class Validator:
    """Validates workbook structure and extracted section data."""

    def validate_workbook(self, reader: ExcelReader) -> ValidationReport:
        """Check that all required sheets exist in the workbook.

        Args:
            reader: An initialised ExcelReader instance.

        Returns:
            A ValidationReport with errors for any missing sheets.
        """
        report = ValidationReport()
        for sheet_name in REQUIRED_SHEETS:
            try:
                reader.get_sheet(sheet_name)
            except SheetNotFoundError:
                report.errors.append(
                    ValidationIssue(
                        level="error",
                        cell_ref="",
                        message=f"Required sheet '{sheet_name}' not found",
                    )
                )
        return report

    def validate_section_data(
        self, data: SectionData, section_name: str
    ) -> ValidationReport:
        """Validate extracted section data for quality issues.

        Checks performed (all non-fatal warnings):
        - Non-numeric values in numeric columns
        - Percentage values outside the valid range (0.0-1.0 or 0-100)
        - Entire month blocks with all-null or all-zero values

        Args:
            data: The extracted SectionData to validate.
            section_name: Human-readable section name for messages.

        Returns:
            A ValidationReport with accumulated warnings.
        """
        report = ValidationReport()
        pct_columns = self._percentage_columns_for(section_name)

        for month in data.months:
            month_all_empty = True

            for team in month.teams:
                for col_name, value in team.values.items():
                    col_lower = col_name.lower().strip()
                    is_pct = col_lower in pct_columns
                    cell_ref = f"{data.sheet_name}!{section_name}/{month.month_name}/{team.team_name}/{col_name}"

                    # Skip None values for numeric / percentage checks
                    if value is None:
                        continue

                    # Check if this value makes the month non-empty
                    if self._is_non_empty(value):
                        month_all_empty = False

                    # Non-numeric check for numeric columns
                    if isinstance(value, str):
                        report.warnings.append(
                            ValidationIssue(
                                level="warning",
                                cell_ref=cell_ref,
                                message=f"Non-numeric value '{value}' in numeric column '{col_name}'",
                            )
                        )
                        continue  # skip percentage range check for strings

                    # Percentage range check
                    if is_pct and isinstance(value, (int, float)):
                        if not self._is_valid_percentage(value):
                            report.warnings.append(
                                ValidationIssue(
                                    level="warning",
                                    cell_ref=cell_ref,
                                    message=(
                                        f"Percentage value {value} outside valid range "
                                        f"(0.0-1.0 or 0-100) in column '{col_name}'"
                                    ),
                                )
                            )

            # All-null / all-zero month block check
            if month_all_empty:
                report.warnings.append(
                    ValidationIssue(
                        level="warning",
                        cell_ref=f"{data.sheet_name}!{section_name}/{month.month_name}",
                        message=f"Month '{month.month_name}' has all-null or all-zero values",
                    )
                )

        return report

    # ------------------------------------------------------------------
    # Internal helpers
    # ------------------------------------------------------------------

    @staticmethod
    def _percentage_columns_for(section_name: str) -> set[str]:
        """Return the set of percentage column names for a given section."""
        name_lower = section_name.lower()
        if "delivery" in name_lower:
            return DELIVERY_PERCENTAGE_COLUMNS
        if "rollout" in name_lower:
            return ROLLOUT_PERCENTAGE_COLUMNS
        # Quality section has no specifically marked percentage columns
        return set()

    @staticmethod
    def _is_valid_percentage(value: Union[int, float]) -> bool:
        """Return True if the value is within a valid percentage range.

        Accepts either the 0.0-1.0 decimal form or the 0-100 whole-number form.
        """
        return 0.0 <= value <= 1.0 or 0 <= value <= 100

    @staticmethod
    def _is_non_empty(value: CellValue) -> bool:
        """Return True if the value is neither None nor zero."""
        if value is None:
            return False
        if isinstance(value, (int, float)) and value == 0:
            return False
        if isinstance(value, str) and value.strip() == "":
            return False
        return True
