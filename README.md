# Excel to PowerPoint Automation Tool

Automates generation of monthly DSD Mukkamala ops report PowerPoint decks from Excel workbooks.

Upload Excel → Get PPT. No manual copy-paste.

---

## How It Works

The tool copies the existing PPT template and surgically updates only the data cells (summary tables, section divider dates) while preserving all formatting, charts, colors, fonts, and layout exactly as-is.

**What gets updated:**
- Slide 0 (Title) — date text
- Slide 1 (Applications divider) — date text
- Slide 2 (Applications summary) — 3 data tables (Quality, Rollout, On-Time Delivery)
- Slide 9 (Platforms divider) — date text
- Slide 10 (Platforms summary) — 3 data tables (Quality, Rollout, On-Time Delivery)
- Slide 18 (DPOS divider) — date text

**What stays untouched (23 of 25 slides):**
- All charts (cumulative historical data preserved)
- All Key Highlights text (manually written narrative)
- All grouped visual elements (KPI boxes)
- All mitigation/outage tables (manually written)
- All formatting, fonts, colors, positioning

---

## Quick Start

### Option 1: Web UI (recommended for non-technical users)

```bash
pip install -r requirements.txt
python web_app.py
```

Open `http://localhost:5000` in your browser:
1. Drag & drop the monthly Excel file
2. Click "Generate PowerPoint"
3. Download the updated PPT

Month is auto-detected from the Excel data. Template is built-in.

### Option 2: Command Line

```bash
pip install -r requirements.txt

python -m app.main \
  --excel "Ops Report Data Collection_February.xlsx" \
  --template "DSD_Mukkamala_February 2026_orig.pptx" \
  --output ./output/ \
  --month February \
  --year 2026
```

---

## Project Structure

```
Telus_Excel_Automation/
├── web_app.py                  # Flask web UI (one-click tool)
├── templates/
│   └── index.html              # Web UI page
├── app/
│   ├── main.py                 # CLI entry point
│   ├── models.py               # Data models (dataclasses)
│   ├── exceptions.py           # Custom exceptions
│   ├── readers/
│   │   └── excel_reader.py     # Excel workbook reader (openpyxl)
│   ├── validators/
│   │   └── validator.py        # Data validation
│   ├── transformers/
│   │   └── transformer.py      # Excel → structured data model
│   ├── generators/
│   │   ├── ppt_updater.py      # PPT template clone & fill (python-pptx)
│   │   ├── slide_handlers.py   # Per-slide custom update logic
│   │   └── output_utils.py     # Output filename/directory helpers
│   └── config/
│       ├── slide_map.yaml      # YAML-driven slide update config
│       └── slide_map_parser.py # YAML parser & validator
├── tests/
│   ├── test_validator.py       # 15 unit tests
│   └── test_ppt_updater.py     # 18 unit tests
├── output/                     # Generated PPT files go here
├── requirements.txt            # Python dependencies
└── DSD_Mukkamala_February 2026_orig.pptx  # Template PPT
```

---

## Architecture

```
Excel (.xlsx)
  → ExcelReader (openpyxl, data_only=True)
  → Validator (check required sheets, data types)
  → Transformer (extract sections, months, teams → DataModel)
  → PPTUpdater (shutil.copy2 template → update shapes in-place)
  → Output .pptx
```

**Key design decision:** Template-clone-and-fill, not build-from-scratch. The tool copies the template PPT and only modifies specific table cells. This preserves all formatting, charts, and layout perfectly.

---

## Excel Requirements

The Excel workbook must contain these sheets:
- `Master-Apps (DoNotChange)` — Application team metrics
- `Master-Platforms (DoNotChange)` — Platform team metrics

Each sheet has 3 sections:
- **On-Time and Planned Delivery** — Track 1/2, Subcontracted, Sustainability delivery counts
- **Quality and Dev Efficiency** — Defect counts, leakage, test cases
- **Rollout Success** — Deployments, rollbacks, hotfixes

Data is organized in month blocks (4 teams per block for Apps, 8 for Platforms).

---

## Dependencies

```
openpyxl>=3.1.0       # Excel reading
python-pptx>=0.6.21   # PowerPoint manipulation
PyYAML>=6.0           # Slide map configuration
flask>=3.0            # Web UI
hypothesis>=6.0       # Property-based testing
pytest>=7.0           # Test runner
```

Install: `pip install -r requirements.txt`

---

## Running Tests

```bash
cd Telus_Excel_Automation
python -m pytest tests/ -v
```

33 tests covering:
- Excel reader (merged cells, #REF! errors, sheet validation)
- Validator (required sheets, data types, percentages, empty months)
- PPT updater (shape lookup, chart/table/text updates, template preservation)

---

## How to Use for Next Month

1. Get the new month's Excel workbook (same format, new data)
2. Either:
   - **Web UI:** Run `python web_app.py`, upload the Excel, download the PPT
   - **CLI:** Run the command with `--month March --year 2026`
3. Open the generated PPT
4. Manually update: Key Highlights text, charts (add new month's bar), mitigation tables
5. Save and distribute

The tool handles the tedious numeric data entry. You still write the narrative parts.

---

## Current Status — 13 of 25 slides auto-updated

**Auto-updated from Excel (13 slides):**
- ✅ S00, S01, S09, S18: Section divider dates
- ✅ S02: Apps summary — all 4 tables (Product Health + Quality + Rollout + Delivery)
- ✅ S03: Apps delivery charts — March column overlaid on historical data
- ✅ S08: Apps Product Health charts — degradations + outages
- ✅ S10: Platforms summary — all 4 tables
- ✅ S11: Platforms delivery charts
- ✅ S14: Platforms Product Health availability chart
- ✅ S15: Platforms Product Health charts
- ✅ S19: DPOS summary — all 4 tables
- ✅ S20: DPOS delivery chart

**Preserved from template — need manual editing (12 slides):**
- S04-05: Apps Quality/Rollout (grouped KPI visuals + narrative)
- S06: ThreatMetrix (separate data source, not in Master sheets)
- S07: Apps outage incident tables (manually written per incident)
- S12-13: Platforms Quality/Rollout (grouped KPI visuals + narrative)
- S16: Data Quality charts (separate data source)
- S17: Data Quality incident table (manually written)
- S21: Embedded OLE objects (Excel screenshots)
- S22-23: DPOS Quality/Rollout (groups + narrative)
- S24: DPOS Product Health (narrative + outage table)

---

## Future Enhancements

- Auto-update chart data (append new month's bar to existing charts)
- Auto-generate Key Highlights text from data patterns
- Support for Product Health section data
- Batch processing (multiple months at once)
- Email/Slack notification on generation
