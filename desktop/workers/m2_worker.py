"""
m2_worker.py — M2 deck generation, desktop flow.

Takes a local Excel path + a PF_ID + an optional questionnaire name.
1. Reads client data directly from the Excel (NOT synced to Google Sheets).
2. Fetches the questionnaire row from Google Sheets (the form submits there).
3. Calls portal/m2_engine.generate_deck() with the assembled data dict.
4. Uploads the resulting .pptx to M2_OUTPUT_FOLDER_ID on Drive.

Returns a dict with {'url', 'name', 'filename'} or raises.
"""
from __future__ import annotations

import app_config                        # noqa: F401  # bootstraps env + sys.path
import pandas as pd

from workers.common import (
    parse_uploaded_excel, filter_data_to_pf_id,
    fetch_questionnaire, upload_pptx_to_drive, PROGRESS,
)


def generate(xlsx_path: str, pf_id: str, customer_name: str,
             questionnaire_name: str | None = None,
             override_risk_profile: str | None = None) -> dict:
    """Run the full M2 pipeline for one client. Returns
    {'url': drive_url, 'name': drive_title, 'filename': pptx_filename}.

    override_risk_profile (optional): one of the RISK_SCALE values
        ('Very Conservative' / 'Conservative' / 'Balanced' / 'Aggressive'
         / 'Very Aggressive'). Bypasses the calculated questionnaire
        risk profile — affects slide 3 display, slide 4 match indicator,
        slide 6 body text, slide 13 Infinite line, AND which Risk
        Reward slides are inserted at slides 15-18.
    """
    # Late imports — portal/m2_engine.py pulls matplotlib, python-pptx, etc.,
    # which are slow. Deferring keeps the GUI snappy on startup.
    import m2_engine   # type: ignore  # via portal/ on sys.path
    import pandas as pd

    PROGRESS(f"[1/5] Parsing {xlsx_path}...")
    raw = parse_uploaded_excel(xlsx_path)

    PROGRESS(f"[2/5] Filtering to PF_ID {pf_id}...")
    data = filter_data_to_pf_id(raw, pf_id)

    # pf_level must have exactly one row for the chosen pf_id
    if data['pf_level'].empty:
        raise ValueError(
            f"PF_ID {pf_id!r} not found in the PF_level tab of the uploaded file."
        )

    # Sync this client's is_demat rows to the Google Sheet so the cloud/
    # Streamlit flow renders the same SOA/Demat for future runs. Desktop
    # still reads is_demat locally from the Excel for THIS run.
    if not data['is_demat'].empty:
        try:
            import sheets  # type: ignore  # portal/sheets.py via app_config
            PROGRESS("  Syncing Is_demat to Google Sheets...")
            sheets.upsert_is_demat(data['is_demat'])
        except Exception as e:
            PROGRESS(f"  WARN: Is_demat sync skipped ({e}).")

    PROGRESS("[3/5] Fetching questionnaire from Google Sheets...")
    data['questionnaire'] = fetch_questionnaire()

    PROGRESS("[4/5] Downloading categorization + template from Drive...")
    # m2_engine pulls categorization + base deck + risk reward + rating PNGs
    # from Drive on demand via cached helpers. No extra plumbing needed here.
    data['categorization'] = pd.read_excel(m2_engine._get_categorization_path())

    if override_risk_profile:
        PROGRESS(f"[5/5] Generating deck for {customer_name} "
                 f"(override risk profile: {override_risk_profile})...")
    else:
        PROGRESS(f"[5/5] Generating deck for {customer_name}...")
    buf, filename = m2_engine.generate_deck(
        pf_id, customer_name, data=data,
        questionnaire_name=questionnaire_name,
        override_risk_profile=override_risk_profile,
    )

    PROGRESS("Uploading to Google Drive...")
    result = upload_pptx_to_drive(
        buf, filename, app_config.M2_OUTPUT_FOLDER_ID, convert_to_slides=True,
    )
    PROGRESS(f"Done. Uploaded as '{result.get('name', filename)}'.")
    return {
        'url': result['url'],
        'name': result.get('name', filename),
        'filename': filename,
    }
