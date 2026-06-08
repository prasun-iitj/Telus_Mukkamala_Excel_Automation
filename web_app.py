"""Simple one-click Excel-to-PPT web tool."""

import os
import logging
import tempfile
from datetime import datetime

from flask import Flask, render_template, request, send_file, flash, redirect, url_for

# Suppress noisy #REF! warnings during generation
logging.getLogger("app.readers.excel_reader").setLevel(logging.ERROR)
logging.getLogger("app.generators.ppt_updater").setLevel(logging.ERROR)

app = Flask(__name__, template_folder="templates")
app.secret_key = os.environ.get("SECRET_KEY", "excel-ppt-2026-dev-only")

# Bundled template — sits next to this file
TEMPLATE_PPT = os.path.join(os.path.dirname(__file__), "DSD_Mukkamala_February 2026_orig.pptx")
UPLOAD_FOLDER = tempfile.mkdtemp()


@app.route("/", methods=["GET"])
def index():
    return render_template("index.html")


@app.route("/generate", methods=["POST"])
def generate():
    excel_file = request.files.get("excel_file")
    if not excel_file or not excel_file.filename:
        flash("Please upload an Excel file.", "error")
        return redirect(url_for("index"))

    excel_path = os.path.join(UPLOAD_FOLDER, "input.xlsx")
    excel_file.save(excel_path)

    try:
        from app.readers.excel_reader import ExcelReader
        from app.transformers.transformer import Transformer
        from app.validators.validator import Validator
        from app.generators.ppt_updater import PPTUpdater
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
        from app.config.slide_map_parser import load_slide_map
        from app.main import _apply_update

        # Load and validate Excel
        reader = ExcelReader(excel_path)
        validator = Validator()
        wb_report = validator.validate_workbook(reader)
        if wb_report.has_fatal_errors:
            msgs = [e.message for e in wb_report.errors]
            flash(f"Excel validation failed: {'; '.join(msgs)}", "error")
            return redirect(url_for("index"))

        # Transform data
        transformer = Transformer()
        data_model = transformer.build_data_model(reader)
        year = datetime.now().year

        # Detect month from filename first (e.g. "Ops Report Data Collection_March.xlsx")
        month = None
        MONTHS = ["January","February","March","April","May","June","July","August","September","October","November","December"]
        fname = excel_file.filename.lower()
        for m in MONTHS:
            if m.lower() in fname:
                month = m
                break
        # Fallback to data-based detection
        if not month:
            month = data_model.reporting_month
        if not month:
            flash("Could not detect reporting month. Include the month name in the filename (e.g. 'Report_March.xlsx').", "error")
            return redirect(url_for("index"))
        data_model.reporting_month = month

        # Generate PPT
        output_path = os.path.join(UPLOAD_FOLDER, f"DSD_Mukkamala_{month}_{year}.pptx")
        updater = PPTUpdater(TEMPLATE_PPT, output_path)

        # YAML-driven updates (date text on slides 0, 1)
        config_path = os.path.join(os.path.dirname(__file__), "app", "config", "slide_map.yaml")
        slide_map = load_slide_map(config_path)
        for slide_cfg in slide_map.slides:
            for update in slide_cfg.updates:
                _apply_update(updater, update, slide_cfg.slide_index, data_model, month, year)

        # Custom handlers — all slides
        update_slide2_apps_summary(updater, data_model, month)
        update_slide3_apps_delivery(updater, data_model, month)
        update_slide4_apps_quality(updater, data_model, month)
        update_slide5_apps_rollout(updater, data_model, month)
        update_slide6_to_8_apps_product_health(updater, data_model, month)
        update_slide6_threatmetrix(updater, excel_path, month)
        update_slide7_apps_outages(updater, excel_path, month)
        update_slide9_platforms_divider(updater, data_model, month, int(year))
        update_slide10_platforms_summary(updater, data_model, month)
        update_slides_11_to_17(updater, data_model, month)
        update_slide16_data_quality(updater, excel_path, month)
        update_slide18_dpos_divider(updater, data_model, month, int(year))
        update_slides_19_to_24(updater, data_model, month)
        update_slide21_dpos_delivery_text(updater, excel_path, month)
        update_slides_22_to_24(updater, data_model, month)

        updater.save()

        response = send_file(
            output_path,
            as_attachment=True,
            download_name=f"DSD_Mukkamala_{month}_{year}.pptx",
            mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        )
        response.set_cookie("ppt_downloaded", "1", max_age=30)
        return response

    except Exception as e:
        flash(f"Error: {str(e)}", "error")
        return redirect(url_for("index"))


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    debug = os.environ.get("FLASK_DEBUG", "").lower() in ("1", "true", "yes")
    app.run(host="0.0.0.0", port=port, debug=debug)
