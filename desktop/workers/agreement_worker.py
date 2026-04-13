"""
agreement_worker.py — Agreement (DOCX) generation, desktop flow.

Reuses portal/agreement_engine.py:
  - generate_elite(client_name, agreement_date, custom_slabs)
      → (BytesIO, filename)
  - generate_nonelite(client_name, email, phone, custom_slabs)
      → (BytesIO, filename)

Then uploads the .docx to AGREEMENT_OUTPUT_FOLDER_ID on Drive (converted
to Google Doc), and returns {url, name} for the GUI to display.
"""
from __future__ import annotations

from datetime import date as _date
from io import BytesIO
from typing import Optional

import app_config                                  # noqa: F401

from workers.common import PROGRESS


def generate(
    is_elite: bool,
    client_name: str,
    *,
    email: str = "",
    phone: str = "",
    agreement_date: Optional[_date] = None,
    custom_slabs: Optional[list[dict]] = None,
) -> dict:
    """Run the full agreement pipeline. Returns {'url', 'name', 'filename'}.
    Raises if generation or upload fails."""
    import sheets               # type: ignore  (via app_config bootstrap)
    import agreement_engine     # type: ignore

    PROGRESS(f"[1/2] Generating "
             f"{'Elite' if is_elite else 'Non-Elite'} agreement for {client_name}...")
    if is_elite:
        buf, filename = agreement_engine.generate_elite(
            client_name.strip(),
            agreement_date=agreement_date,
            custom_slabs=custom_slabs,
        )
    else:
        buf, filename = agreement_engine.generate_nonelite(
            client_name.strip(),
            email=email.strip(),
            phone=phone.strip(),
            custom_slabs=custom_slabs,
        )

    PROGRESS("[2/2] Uploading to Google Drive...")
    result = sheets.upload_docx_to_drive(
        buf, filename,
        app_config.AGREEMENT_OUTPUT_FOLDER_ID,
        convert_to_gdoc=True,
    )

    PROGRESS(f"Done. Uploaded as '{result.get('name', filename)}'.")
    return {
        'url': result['url'],
        'name': result.get('name', filename),
        'filename': filename,
    }
