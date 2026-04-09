"""
tabs/agreement.py — Agreement Automation tab.

Two agreement types:
  - Elite (PowerUp Infinite Agreement): name + date + fee slabs
  - Non-Elite (UW IA Client Agreement): name + email + phone + fee slabs

Fee modes:
  - Default slabs (already in template — just fill name/date)
  - Custom slabs (user enters fee % and AUA range per row)

Generated DOCX is uploaded to Drive, then user can download as PDF.
"""

from datetime import date

import streamlit as st

import sheets
import config
from agreement_engine import generate_elite, generate_nonelite


def render():
    st.header("Agreement Automation")
    st.caption("Generate client agreements with default or custom fee slabs.")
    st.divider()

    # ── Config check ──────────────────────────────────────────
    if not config.AGREEMENT_OUTPUT_FOLDER_ID:
        st.error(
            "**AGREEMENT_OUTPUT_FOLDER_ID is not configured.**\n\n"
            "Add it to `portal/.env` or Streamlit Cloud secrets."
        )
        return

    # ── Agreement type ────────────────────────────────────────
    agreement_type = st.radio(
        "Agreement type",
        ["Elite (PowerUp Infinite)", "Non-Elite (UW IA Client Agreement)"],
        horizontal=True,
        key="ag_type",
    )
    is_elite = "Elite" in agreement_type

    st.divider()

    # ── Client details ────────────────────────────────────────
    client_name = st.text_input(
        "Client name",
        key="ag_name",
        placeholder="Full name as it should appear on the agreement",
    )

    email = ""
    phone = ""
    if not is_elite:
        col1, col2 = st.columns(2)
        with col1:
            email = st.text_input(
                "Email",
                key="ag_email",
                placeholder="Client email (or NA to skip)",
            )
        with col2:
            phone = st.text_input(
                "Phone",
                key="ag_phone",
                placeholder="Phone number (or NA to skip)",
            )

    agreement_date = None
    if is_elite:
        agreement_date = st.date_input(
            "Agreement date",
            value=date.today(),
            key="ag_date",
        )

    # ── Fee slabs ─────────────────────────────────────────────
    st.divider()
    slab_mode = st.radio(
        "Fee slabs",
        ["Default (3 standard slabs)", "Custom"],
        horizontal=True,
        key="ag_slab_mode",
    )

    custom_slabs = None
    if slab_mode == "Custom":
        st.caption("Enter fee slabs. Use the number input to set how many rows you need.")

        slab_count = st.number_input(
            "Number of fee slabs",
            min_value=1,
            max_value=10,
            value=1,
            step=1,
            key="ag_slab_count",
        )

        slabs_input = []
        for i in range(slab_count):
            col_fee, col_aua = st.columns(2)
            with col_fee:
                fee = st.text_input(
                    f"Fee % (slab {i + 1})",
                    key=f"ag_fee_{i}",
                    placeholder="e.g. 0.75% p.a.",
                )
            with col_aua:
                aua = st.text_input(
                    f"AUA range (slab {i + 1})",
                    key=f"ag_aua_{i}",
                    placeholder="e.g. less than 50 lakhs",
                )
            if fee.strip():
                fee_val = fee.strip()
                if "p.a" not in fee_val.lower():
                    fee_val = fee_val + " p.a."
                slabs_input.append({"fee": fee_val, "aua": aua.strip()})

        if slabs_input:
            custom_slabs = slabs_input

    # ── Generate ──────────────────────────────────────────────
    st.divider()

    folder_url = f"https://drive.google.com/drive/folders/{config.AGREEMENT_OUTPUT_FOLDER_ID}"
    st.caption(f"Agreements saved to: [Agreement Output folder]({folder_url})")

    if not client_name.strip():
        st.warning("Enter the client's name to proceed.")
        return

    if st.button("Generate Agreement", type="primary", use_container_width=True, key="ag_generate"):
        with st.spinner("Generating agreement..."):
            try:
                if is_elite:
                    buf, filename = generate_elite(
                        client_name.strip(),
                        agreement_date=agreement_date,
                        custom_slabs=custom_slabs,
                    )
                else:
                    buf, filename = generate_nonelite(
                        client_name.strip(),
                        email=email.strip(),
                        phone=phone.strip(),
                        custom_slabs=custom_slabs,
                    )
            except Exception as e:
                st.error(f"Agreement generation failed: {e}")
                return

        with st.spinner("Uploading to Google Drive..."):
            try:
                result = sheets.upload_docx_to_drive(
                    buf, filename,
                    config.AGREEMENT_OUTPUT_FOLDER_ID,
                    convert_to_gdoc=True,
                )
            except Exception as e:
                st.error(f"Upload to Drive failed: {e}")
                buf.seek(0)
                st.download_button(
                    "Download DOCX locally",
                    data=buf,
                    file_name=filename,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                )
                return

        doc_url = result["url"]
        doc_name = result["name"]
        doc_id = result["id"]

        st.success("Agreement generated and uploaded successfully.")
        st.markdown(f"### [{doc_name}]({doc_url})")
        st.link_button("Open in Google Docs", doc_url, type="primary")

        # PDF download
        with st.spinner("Preparing PDF download..."):
            try:
                pdf_buf = sheets.export_drive_file_as_pdf(doc_id)
                pdf_name = filename.replace(".docx", ".pdf")
                st.download_button(
                    "Download as PDF",
                    data=pdf_buf,
                    file_name=pdf_name,
                    mime="application/pdf",
                )
            except Exception as e:
                st.caption(f"PDF export not available: {e}")
