"""
tabs/m3_deck.py — M3 Portfolio Transition Deck tab.

Flow:
  1. User uploads client Excel workbook (PF_Curation or PF_MasterPlan sheet)
  2. User enters (or confirms) client name
  3. "Generate Deck" → m3_engine reads Excel, fetches reference data from Sheets,
     downloads template from Drive, builds PPTX in memory
  4. PPTX uploaded to Google Drive (M3_OUTPUT_FOLDER_ID), converted to Slides
  5. Portal shows clickable link to the generated presentation
"""

import streamlit as st

import sheets
from config import M3_OUTPUT_FOLDER_ID
from m3_engine import read_excel, generate_deck, load_reference_data


def render():
    st.header("M3 — Portfolio Transition Deck")
    st.caption(
        "Upload the client's Excel workbook (PF_Curation or PF_MasterPlan sheet) "
        "and generate a transition deck as a Google Slides presentation."
    )
    st.divider()

    # ── Config check ──────────────────────────────────────────
    if not M3_OUTPUT_FOLDER_ID:
        st.error(
            "**M3_OUTPUT_FOLDER_ID is not configured.**\n\n"
            "Add it to `portal/.env`:\n"
            "```\nM3_OUTPUT_FOLDER_ID=your_folder_id_here\n```"
        )
        return

    # ── File upload ───────────────────────────────────────────
    uploaded = st.file_uploader(
        "Upload client Excel workbook",
        type=["xlsx", "xls"],
        key="m3_upload",
        help="Must contain a PF_Curation_* or PF_MasterPlan_* sheet.",
    )

    if not uploaded:
        st.info("Upload a client Excel file to get started.")
        return

    # ── Parse Excel ───────────────────────────────────────────
    try:
        import tempfile, os
        # Save uploaded file to temp location for openpyxl
        tmp_path = os.path.join(tempfile.gettempdir(), f"m3_upload_{uploaded.name}")
        with open(tmp_path, "wb") as tmp:
            tmp.write(uploaded.read())

        excel_data = read_excel(tmp_path)

        try:
            os.unlink(tmp_path)
        except OSError:
            pass  # file locked — will be cleaned up by OS
    except Exception as e:
        st.error(f"Could not parse Excel file: {e}")
        return

    # Show summary
    sections_found = []
    for key in ('section1', 'section2', 'section3', 'section4'):
        rows = [r for r in excel_data.get(key, []) if not r.get('__grand_total__')]
        if rows:
            sections_found.append(f"{key}: {len(rows)} rows")
    st.success(f"Excel parsed successfully — {', '.join(sections_found)}")

    # ── Client name ───────────────────────────────────────────
    client_name = st.text_input(
        "Client name",
        value="",
        key="m3_client_name",
        help="Full name as it should appear on the deck cover.",
    )

    if not client_name.strip():
        st.warning("Enter the client's name to proceed.")
        return

    # ── Output folder info ────────────────────────────────────
    folder_url = f"https://drive.google.com/drive/folders/{M3_OUTPUT_FOLDER_ID}"
    st.caption(f"Decks saved to: [M3 Output folder]({folder_url})")

    st.divider()

    # ── Generate button ───────────────────────────────────────
    if st.button("Generate Deck", type="primary", use_container_width=True, key="m3_generate"):
        with st.spinner("Loading reference data from Google Sheets..."):
            try:
                ref_data, rr_category = load_reference_data()
            except Exception as e:
                st.error(f"Failed to load reference data: {e}")
                return

        with st.spinner(f"Generating transition deck for {client_name}... (this can take up to 2 minutes)"):
            try:
                buf, filename = generate_deck(
                    excel_data, client_name.strip(),
                    ref_data=ref_data, rr_category=rr_category,
                )
            except Exception as e:
                st.error(f"Deck generation failed: {e}")
                return

        with st.spinner("Uploading to Google Drive..."):
            try:
                result = sheets.upload_pptx_to_drive(
                    buf, filename, M3_OUTPUT_FOLDER_ID, convert_to_slides=True,
                )
            except Exception as e:
                st.error(f"Upload to Drive failed: {e}")
                buf.seek(0)
                st.download_button(
                    "Download PPTX locally",
                    data=buf,
                    file_name=filename,
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                )
                return

        slide_url = result["url"]
        slide_name = result["name"]

        st.success("Transition deck generated and uploaded successfully.")
        st.markdown(
            f"### [{slide_name}]({slide_url})",
            help="Click to open the generated Google Slides presentation.",
        )
        st.link_button("Open deck", slide_url, type="primary")
