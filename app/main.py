"""CLI entry point for the Excel-to-PPT automation pipeline."""

from __future__ import annotations

import argparse
import logging
import os
import sys
from datetime import datetime
from typing import Optional

from app.config.slide_map_parser import (
    SourceConfig,
    SlideMapConfig,
    UpdateConfig,
    load_slide_map,
)
from app.exceptions import PipelineError
from app.generators.output_utils import resolve_output_path
from app.generators.ppt_updater import PPTUpdater
from app.models import (
    ChartDataSpec,
    DataModel,
    SectionData,
    SeriesSpec,
    TableDataSpec,
)
from app.readers.excel_reader import ExcelReader
from app.transformers.transformer import Transformer
from app.validators.validator import Validator
from app.generators.slide_handlers import (
    update_slide2_apps_summary,
    update_slide3_apps_delivery,
    update_slide4_apps_quality,
    update_slide5_apps_rollout,
    update_slide6_to_8_apps_product_health,
    update_slide6_threatmetrix,
    update_slide7_apps_outages,
    update_slide9_platforms_divider,
    update_slide10_platforms_summary,
    update_slides_11_to_17,
    update_slide16_data_quality,
    update_slide18_dpos_divider,
    update_slides_19_to_24,
    update_slide21_dpos_delivery_text,
    update_slides_22_to_24,
)

logger = logging.getLogger(__name__)

# Column letter to 0-based index offset (D=0, E=1, F=2, …)
_COL_OFFSET = {chr(c): c - ord("D") for c in range(ord("D"), ord("Z") + 1)}


# ---------------------------------------------------------------------------
# Data resolver – bridge between slide_map config and the DataModel
# ---------------------------------------------------------------------------


def _get_section(data_model: DataModel, sheet: str, section: str) -> Optional[SectionData]:
    """Look up a SectionData by sheet + section header name."""
    sections = {
        ("Master-Apps (DoNotChange)", "On-Time and Planned Delivery"): data_model.apps_delivery,
        ("Master-Apps (DoNotChange)", "Quality and Dev Efficiency"): data_model.apps_quality,
        ("Master-Apps (DoNotChange)", "Rollout Success"): data_model.apps_rollout,
        ("Master-Apps (DoNotChange)", "Product Health Measure"): data_model.apps_product_health,
        ("Master-Platforms (DoNotChange)", "On-Time and Planned Delivery"): data_model.platforms_delivery,
        ("Master-Platforms (DoNotChange)", "Quality and Dev Efficiency"): data_model.platforms_quality,
        ("Master-Platforms (DoNotChange)", "Rollout Success"): data_model.platforms_rollout,
        ("Master-Platforms (DoNotChange)", "Product Health Measure"): data_model.platforms_product_health,
    }
    return sections.get((sheet, section))


def _col_index(col_letter: str) -> int:
    """Convert a column letter (D, E, …) to a 0-based index into column_headers."""
    return _COL_OFFSET.get(col_letter.upper(), -1)


def resolve_chart_data(
    source: SourceConfig,
    data_model: DataModel,
) -> ChartDataSpec:
    """Build a ChartDataSpec from the slide-map source config and data model.

    For chart sources the categories are month abbreviations and each series
    aggregates the *total* across all teams for the specified column in each
    month.
    """
    section = _get_section(data_model, source.sheet, source.section)
    if section is None:
        logger.warning(
            "Section '%s' not found in sheet '%s' — returning empty chart data",
            source.section,
            source.sheet,
        )
        categories = source.categories or []
        series_list = [
            SeriesSpec(name=s.name, values=[None] * len(categories))
            for s in (source.series or [])
        ]
        return ChartDataSpec(categories=categories, series=series_list)

    categories = source.categories or []

    # Build a lookup: month_name -> MonthData
    month_lookup = {m.month_name: m for m in section.months}

    # Map 3-letter abbreviations to full month names
    abbrev_to_full = {
        "Jan": "January", "Feb": "February", "Mar": "March",
        "Apr": "April", "May": "May", "Jun": "June",
        "Jul": "July", "Aug": "August", "Sep": "September",
        "Oct": "October", "Nov": "November", "Dec": "December",
    }

    series_list: list[SeriesSpec] = []
    for s_cfg in (source.series or []):
        col_idx = _col_index(s_cfg.column)
        values: list[Optional[float]] = []
        for cat in categories:
            full_month = abbrev_to_full.get(cat, cat)
            month_data = month_lookup.get(full_month)
            if month_data is None or col_idx < 0 or col_idx >= len(section.column_headers):
                values.append(None)
                continue
            header_name = section.column_headers[col_idx]
            # Sum across all teams for this month
            total = None
            for team in month_data.teams:
                val = team.values.get(header_name)
                if val is not None and isinstance(val, (int, float)):
                    total = (total or 0) + val
            values.append(total)
        series_list.append(SeriesSpec(name=s_cfg.name, values=values))

    return ChartDataSpec(categories=categories, series=series_list)


def resolve_table_data(
    source: SourceConfig,
    data_model: DataModel,
    reporting_month: str,
) -> TableDataSpec:
    """Build a TableDataSpec for the current month's team data."""
    section = _get_section(data_model, source.sheet, source.section)
    if section is None:
        logger.warning(
            "Section '%s' not found in sheet '%s' — returning empty table data",
            source.section,
            source.sheet,
        )
        return TableDataSpec(headers=[], rows=[])

    # Find the current month's data
    month_data = None
    for m in section.months:
        if m.month_name == reporting_month:
            month_data = m
            break

    if month_data is None:
        logger.warning(
            "No data for month '%s' in section '%s'",
            reporting_month,
            source.section,
        )
        return TableDataSpec(headers=[], rows=[])

    # Determine which columns to include
    col_indices = [_col_index(c) for c in (source.columns or [])]
    headers = []
    for idx in col_indices:
        if 0 <= idx < len(section.column_headers):
            headers.append(section.column_headers[idx])
        else:
            headers.append("")

    rows: list[list] = []
    for team in month_data.teams:
        row = [team.team_name]
        for idx in col_indices:
            if 0 <= idx < len(section.column_headers):
                header_name = section.column_headers[idx]
                row.append(team.values.get(header_name))
            else:
                row.append(None)
        rows.append(row)

    # Prepend "Team" to headers
    headers = ["Team"] + headers

    return TableDataSpec(headers=headers, rows=rows)


def resolve_text_data(
    source: SourceConfig,
    data_model: DataModel,
    reporting_month: str,
) -> str:
    """Build text content from the current month's text fields."""
    section = _get_section(data_model, source.sheet, source.section)
    if section is None:
        return ""

    month_data = None
    for m in section.months:
        if m.month_name == reporting_month:
            month_data = m
            break

    if month_data is None:
        return ""

    col_indices = [_col_index(c) for c in (source.columns or [])]
    lines: list[str] = []

    for team in month_data.teams:
        for idx in col_indices:
            if 0 <= idx < len(section.column_headers):
                header_name = section.column_headers[idx]
                val = team.values.get(header_name)
                if val is not None and str(val).strip():
                    lines.append(str(val).strip())

    return "\n".join(lines)


def resolve_data(
    source: SourceConfig,
    data_model: DataModel,
    reporting_month: str,
    year: int,
):
    """Resolve data from a source config – returns ChartDataSpec, TableDataSpec, or str."""
    # Static text
    if source.type == "static":
        return source.value or ""

    # Month/year text
    if source.type == "month_year":
        return f"{reporting_month} {year}"

    # Chart data
    if source.series is not None and source.categories is not None:
        return resolve_chart_data(source, data_model)

    # Table data
    if source.month is not None and source.columns is not None and source.series is None and source.categories is None:
        # Check if it's a text block (has format field)
        if source.format is not None:
            return resolve_text_data(source, data_model, reporting_month)
        return resolve_table_data(source, data_model, reporting_month)

    # Fallback: text data
    if source.sheet and source.section:
        return resolve_text_data(source, data_model, reporting_month)

    return ""


# ---------------------------------------------------------------------------
# CLI argument parsing
# ---------------------------------------------------------------------------


def build_parser() -> argparse.ArgumentParser:
    """Build and return the CLI argument parser."""
    script_dir = os.path.dirname(os.path.abspath(__file__))
    default_config = os.path.join(script_dir, "config", "slide_map.yaml")

    parser = argparse.ArgumentParser(
        description="Excel-to-PPT Automation: generate monthly ops report decks.",
    )
    parser.add_argument(
        "--excel",
        required=True,
        help="Path to input .xlsx workbook",
    )
    parser.add_argument(
        "--template",
        required=True,
        help="Path to template .pptx file",
    )
    parser.add_argument(
        "--config",
        default=default_config,
        help="Path to slide_map.yaml (default: app/config/slide_map.yaml)",
    )
    parser.add_argument(
        "--output",
        required=True,
        help="Path to output .pptx file or directory",
    )
    parser.add_argument(
        "--month",
        default=None,
        help="Reporting month name (e.g. February). Auto-detected if omitted.",
    )
    parser.add_argument(
        "--project",
        default="DSD_Mukkamala",
        help="Project name for output filename (default: DSD_Mukkamala)",
    )
    parser.add_argument(
        "--year",
        type=int,
        default=datetime.now().year,
        help="Reporting year for output filename (default: current year)",
    )
    return parser


# ---------------------------------------------------------------------------
# Pipeline
# ---------------------------------------------------------------------------


def run_pipeline(args: argparse.Namespace) -> str:
    """Execute the full pipeline and return the output file path.

    Raises PipelineError on any step failure.
    """
    # 1. Load Excel workbook
    print(f"Loading Excel workbook: {args.excel}")
    reader = ExcelReader(args.excel)

    # 2. Validate workbook structure
    print("Validating workbook structure...")
    validator = Validator()
    wb_report = validator.validate_workbook(reader)
    if wb_report.has_fatal_errors:
        for err in wb_report.errors:
            print(f"  ERROR: {err.message}", file=sys.stderr)
        raise PipelineError("validation", "Workbook validation failed with fatal errors")

    # 3. Transform data
    print("Transforming data...")
    transformer = Transformer()
    data_model = transformer.build_data_model(reader)

    # 4. Override reporting month if provided
    if args.month:
        data_model.reporting_month = args.month

    reporting_month = data_model.reporting_month
    if not reporting_month:
        raise PipelineError(
            "transform",
            "Could not auto-detect reporting month (no months with data found)",
        )
    print(f"Reporting month: {reporting_month}")

    # 5. Validate section data
    print("Validating section data...")
    all_sections = [
        (data_model.apps_delivery, "Apps Delivery"),
        (data_model.apps_quality, "Apps Quality"),
        (data_model.apps_rollout, "Apps Rollout"),
        (data_model.apps_product_health, "Apps Product Health"),
        (data_model.platforms_delivery, "Platforms Delivery"),
        (data_model.platforms_quality, "Platforms Quality"),
        (data_model.platforms_rollout, "Platforms Rollout"),
        (data_model.platforms_product_health, "Platforms Product Health"),
    ]
    all_warnings = []
    for section_data, section_label in all_sections:
        sec_report = validator.validate_section_data(section_data, section_label)
        all_warnings.extend(sec_report.warnings)

    # 6. Print validation report
    if all_warnings:
        print(f"\nValidation warnings ({len(all_warnings)}):")
        for w in all_warnings:
            print(f"  WARNING [{w.cell_ref}]: {w.message}")
        print()
    else:
        print("No validation warnings.\n")

    # 7. Resolve output path
    output_path = resolve_output_path(
        args.output, args.project, reporting_month, args.year,
    )

    # 8. Load slide map
    print(f"Loading slide map: {args.config}")
    slide_map = load_slide_map(args.config)

    # 9. Clone template and create updater
    print(f"Cloning template: {args.template}")
    updater = PPTUpdater(args.template, output_path)

    # 10. Process YAML-driven slide updates
    print("Updating slides (YAML-driven)...")
    for slide_cfg in slide_map.slides:
        for update in slide_cfg.updates:
            _apply_update(updater, update, slide_cfg.slide_index, data_model, reporting_month, args.year)

    # 11. Process custom slide handlers
    print("Updating slides (custom handlers)...")
    update_slide2_apps_summary(updater, data_model, reporting_month)
    update_slide3_apps_delivery(updater, data_model, reporting_month)
    update_slide4_apps_quality(updater, data_model, reporting_month)
    update_slide5_apps_rollout(updater, data_model, reporting_month)
    update_slide6_to_8_apps_product_health(updater, data_model, reporting_month)
    update_slide6_threatmetrix(updater, args.excel, reporting_month)
    update_slide7_apps_outages(updater, args.excel, reporting_month)
    update_slide9_platforms_divider(updater, data_model, reporting_month, args.year)
    update_slide10_platforms_summary(updater, data_model, reporting_month)
    update_slides_11_to_17(updater, data_model, reporting_month)
    update_slide16_data_quality(updater, args.excel, reporting_month)
    update_slide18_dpos_divider(updater, data_model, reporting_month, args.year)
    update_slides_19_to_24(updater, data_model, reporting_month)
    update_slide21_dpos_delivery_text(updater, args.excel, reporting_month)
    update_slides_22_to_24(updater, data_model, reporting_month)

    # 12. Save output
    saved_path = updater.save()
    return saved_path


def _apply_update(
    updater: PPTUpdater,
    update: UpdateConfig,
    slide_index: int,
    data_model: DataModel,
    reporting_month: str,
    year: int,
) -> None:
    """Apply a single shape update from the slide map."""
    source = update.source
    data = resolve_data(source, data_model, reporting_month, year)

    if update.field == "chart" and isinstance(data, ChartDataSpec):
        # Skip chart update if all series have only None values (empty section)
        has_data = any(
            v is not None
            for s in data.series
            for v in s.values
        )
        if not has_data:
            logger.warning(
                "Skipping chart update on slide %d, shape '%s': no data available",
                slide_index,
                update.shape_name,
            )
            return
        updater.update_chart(
            slide_index,
            update.shape_name,
            data,
            shape_index=update.shape_index,
        )
    elif update.field == "table" and isinstance(data, TableDataSpec):
        # Skip table update if there are no rows (empty section)
        if not data.rows and not data.headers:
            logger.warning(
                "Skipping table update on slide %d, shape '%s': no data available",
                slide_index,
                update.shape_name,
            )
            return
        updater.update_table(
            slide_index,
            update.shape_name,
            data,
            shape_index=update.shape_index,
        )
    elif update.field == "text":
        text_val = data if isinstance(data, str) else str(data)
        updater.update_text_frame(
            slide_index,
            update.shape_name,
            text_val,
            shape_index=update.shape_index,
        )
    else:
        logger.warning(
            "Skipping update on slide %d, shape '%s': "
            "unhandled field '%s' / data type '%s'",
            slide_index,
            update.shape_name,
            update.field,
            type(data).__name__,
        )


# ---------------------------------------------------------------------------
# Main entry point
# ---------------------------------------------------------------------------


def main() -> None:
    """Parse arguments and run the pipeline."""
    # Ensure UTF-8 output on Windows to handle special characters from Excel data
    import io
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    if hasattr(sys.stderr, "reconfigure"):
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")

    parser = build_parser()
    args = parser.parse_args()

    try:
        output_path = run_pipeline(args)
        print(f"Success! Output saved to: {output_path}")
    except PipelineError as exc:
        print(f"Error: {exc}", file=sys.stderr)
        sys.exit(1)
    except Exception as exc:
        print(f"Error: [unexpected] {exc}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
