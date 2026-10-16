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


def _upload_pdf_to_drive(pdf_buf, filename: str, folder_id: str) -> dict:
    """Upload a PDF BytesIO to Drive. Returns {id, url, name}."""
    import sheets  # type: ignore
    created = sheets.create_drive_file_with_retry(
        {"name": filename, "parents": [folder_id]}, pdf_buf, "application/pdf",
    )
    return {
        "id":   created["id"],
        "url":  created.get("webViewLink",
                            f"https://drive.google.com/file/d/{created['id']}/view"),
        "name": created["name"],
    }


def generate(
    is_elite: bool,
    client_name: str,
    *,
    email: str = "",
    phone: str = "",
    agreement_date: Optional[_date] = None,
    custom_slabs: Optional[list[dict]] = None,
    signatory_capacity: str = "Self",
) -> dict:
    """Run the full agreement pipeline.

    Pipeline:
      1. agreement_engine generates the .docx in memory.
      2. .docx is uploaded to Drive AND converted to a Google Doc — this is
         the only reliable way to get a clean PDF (Drive's converter handles
         page breaks / fonts / table styling correctly).
      3. The Google Doc is exported as PDF in memory.
      4. The PDF is uploaded to Drive as the user-facing artifact.
      5. The intermediate Google Doc is deleted to keep the folder clean.
         (If you'd rather keep the editable Doc, comment out step 5.)

    Returns {'url', 'name', 'filename'} — pointing at the PDF.
    """
    import sheets               # type: ignore  (via app_config bootstrap)
    import agreement_engine     # type: ignore
    from google_auth import get_drive_service  # type: ignore

    PROGRESS(f"[1/4] Generating "
             f"{'Elite' if is_elite else 'Non-Elite'} agreement for {client_name}...")
    if is_elite:
        buf, filename = agreement_engine.generate_elite(
            client_name.strip(),
            agreement_date=agreement_date,
            custom_slabs=custom_slabs,
            signatory_capacity=signatory_capacity,
        )
    else:
        buf, filename = agreement_engine.generate_nonelite(
            client_name.strip(),
            email=email.strip(),
            phone=phone.strip(),
            custom_slabs=custom_slabs,
        )

    PROGRESS("[2/4] Uploading and converting to Google Doc...")
    gdoc = sheets.upload_docx_to_drive(
        buf, filename,
        app_config.AGREEMENT_OUTPUT_FOLDER_ID,
        convert_to_gdoc=True,
    )

    PROGRESS("[3/4] Exporting as PDF...")
    pdf_buf = sheets.export_drive_file_as_pdf(gdoc["id"])

    pdf_filename = filename.rsplit(".", 1)[0] + ".pdf"
    PROGRESS(f"[4/4] Uploading {pdf_filename} to Drive...")
    pdf = _upload_pdf_to_drive(
        pdf_buf, pdf_filename, app_config.AGREEMENT_OUTPUT_FOLDER_ID,
    )

    # Clean up the intermediate Google Doc — the PDF is the deliverable.
    # On a shared drive the service account may not be allowed to delete
    # (Drive answers 404), so fall back to moving it to the trash.
    drive = get_drive_service()
    try:
        drive.files().delete(fileId=gdoc["id"], supportsAllDrives=True).execute()
    except Exception:
        try:
            drive.files().update(fileId=gdoc["id"], body={"trashed": True},
                                 supportsAllDrives=True).execute()
        except Exception as e:
            PROGRESS(f"  WARN: couldn't remove intermediate Google Doc: {e}")

    PROGRESS(f"Done. Uploaded as '{pdf['name']}'.")
    return {
        'url':      pdf['url'],
        'name':     pdf['name'],
        'filename': pdf_filename,
    }
