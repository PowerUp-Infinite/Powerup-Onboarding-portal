"""
m3_worker.py — M3 portfolio transition deck, desktop flow.

Reuses portal/m3_engine.py:
  - read_excel(path)  → parses the uploaded M3 workbook into a dict of sections
  - load_reference_data() → fetches monthly ref data from Google Sheets
  - generate_deck(excel_data, client_name, ref_data, rr_category)
    → returns (BytesIO, filename)

Then uploads the .pptx to M3_OUTPUT_FOLDER_ID on Drive.
"""
from __future__ import annotations

import app_config                        # noqa: F401

from workers.common import upload_pptx_to_drive, PROGRESS


def generate(xlsx_path: str, client_name: str) -> dict:
    """Run the full M3 pipeline. Returns {'url', 'name', 'filename'}."""
    import m3_engine  # type: ignore  # via portal/ on sys.path

    PROGRESS(f"[1/4] Parsing M3 workbook {xlsx_path}...")
    excel_data = m3_engine.read_excel(xlsx_path)

    PROGRESS("[2/4] Loading monthly reference data from Google Sheets...")
    ref_data, rr_category = m3_engine.load_reference_data()

    PROGRESS(f"[3/4] Generating M3 deck for {client_name}...")
    buf, filename = m3_engine.generate_deck(
        excel_data, client_name, ref_data=ref_data, rr_category=rr_category,
    )

    PROGRESS("[4/4] Uploading to Google Drive...")
    result = upload_pptx_to_drive(
        buf, filename, app_config.M3_OUTPUT_FOLDER_ID, convert_to_slides=True,
    )
    PROGRESS(f"Done. Uploaded as '{result.get('name', filename)}'.")
    return {
        'url': result['url'],
        'name': result.get('name', filename),
        'filename': filename,
    }
