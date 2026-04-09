"""
app.py
======
Streamlit webapp for extracting owner financial data.

Upload an .xlsx file, enter a property name, click Extract, and download
the output workbook — all in the browser.

Uses the same extraction logic as the command-line script via extractor_core.py.

TO RUN:
  streamlit run app.py
"""

import os
import tempfile
import streamlit as st

from extractor_core import (
    guess_property_name,
    process_workbook,
    MANAGED_BY_OPTIONS,
    DEFAULT_MANAGED_BY,
)


# ---------------------------------------------------------------------------
# PAGE CONFIG
# ---------------------------------------------------------------------------

st.set_page_config(
    page_title="Owner Financials Extractor",
    page_icon="📊",
    layout="centered",
)


# ---------------------------------------------------------------------------
# HEADER
# ---------------------------------------------------------------------------

st.title("📊 Owner Financials Extractor")
st.markdown(
    "Upload an owner financial workbook, enter a property name, "
    "and download the extracted datapack."
)

st.divider()


# ---------------------------------------------------------------------------
# FILE UPLOAD
# ---------------------------------------------------------------------------

uploaded_file = st.file_uploader(
    "Upload owner financial workbook (.xlsx)",
    type=["xlsx"],
    help="Drop an owner financial Excel file here (Extra Space, Public Storage, CubeSmart, or Other).",
)


# ---------------------------------------------------------------------------
# PROPERTY NAME INPUT
# ---------------------------------------------------------------------------

# Guess a default name from the uploaded filename
if uploaded_file is not None:
    default_name = guess_property_name(uploaded_file.name)
else:
    default_name = ""

property_name = st.text_input(
    "Property name (used in output filename and Rolling IS tab)",
    value=default_name,
    placeholder="e.g. Chattanooga",
)

managed_by = st.selectbox(
    "Management company",
    options=MANAGED_BY_OPTIONS,
    index=MANAGED_BY_OPTIONS.index(DEFAULT_MANAGED_BY),
)


# ---------------------------------------------------------------------------
# EXTRACT BUTTON
# ---------------------------------------------------------------------------

# Only show the button if a file is uploaded and a name is entered
if uploaded_file is not None and property_name.strip():

    if st.button("🚀 Extract Data", type="primary", use_container_width=True):

        with st.spinner("Extracting data..."):

            # Save the uploaded file to a temporary location so openpyxl can read it.
            # (openpyxl needs a file path, not a file-like object in read_only mode)
            with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
                tmp.write(uploaded_file.getvalue())
                tmp_path = tmp.name

            try:
                # Call the shared core extraction function
                result = process_workbook(tmp_path, property_name.strip(),
                                         managed_by=managed_by)
            finally:
                # Clean up the temp file
                os.unlink(tmp_path)

        # -- Display results --

        if result["output_bytes"] is None:
            st.error("Could not process this file. Check the log below.")
        else:
            # Success banner
            st.success(f"Extraction complete: **{result['output_filename']}**")

            # Summary metrics in columns
            summary = result.get("summary", {})
            if summary:
                cols = st.columns(len(summary))
                labels = {
                    "rolling_is": "Rolling IS",
                    "unit_rate": "Unit Rate",
                    "ops_sum": "Ops Sum",
                    "rent_roll": "Rent Roll",
                }
                for col, (key, msg) in zip(cols, summary.items()):
                    col.metric(labels.get(key, key), msg)

            # Download button
            st.download_button(
                label="📥 Download Datapack",
                data=result["output_bytes"],
                file_name=result["output_filename"],
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

        # Processing log (always shown)
        st.divider()
        st.subheader("Processing Log")
        for entry in result["log"]:
            icon = "✅" if entry["status"] == "OK" else "⚠️" if entry["status"] == "WARNING" else "❌"
            st.markdown(f"{icon} **{entry['sheet']}** — {entry['message']}")

elif uploaded_file is not None and not property_name.strip():
    st.info("Enter a property name above to continue.")


# ---------------------------------------------------------------------------
# FOOTER
# ---------------------------------------------------------------------------

st.divider()
st.caption("Owner Financials Extractor v3.0 — Streamlit Edition")
