"""Tests for reporting-month auto-detection."""

import logging

from app.readers.excel_reader import ExcelReader
from app.transformers.transformer import Transformer

logging.disable(logging.CRITICAL)


def test_march_workbook_detects_march_not_april():
    reader = ExcelReader("Ops Report Data Collection_March.xlsx")
    dm = Transformer().build_data_model(reader)
    assert dm.reporting_month == "March"


def test_february_workbook_detects_february():
    reader = ExcelReader("Ops Report Data Collection_February.xlsx")
    dm = Transformer().build_data_model(reader)
    assert dm.reporting_month == "February"
