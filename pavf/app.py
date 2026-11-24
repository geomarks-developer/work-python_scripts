"""pyPAVF Streamlit application."""
from __future__ import annotations

from io import BytesIO
from pathlib import Path
import tempfile

import pandas as pd
import streamlit as st

from modules.pavf_reader import read_pavf_records
from modules.email_evidence_reader import read_email_evidence
from modules.reconciliation_matcher import ReconciliationMatcher

APP_DIR = Path(__file__).resolve().parent
DATA_DIR = APP_DIR / "data"
OUTPUT_DIR = DATA_DIR / "output"
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

st.set_page_config(
    page_title="pyPAVF",
    page_icon="📋",
    layout="wide",
)


def save_upload(uploaded_file, destination: Path) -> Path:
    destination.parent.mkdir(parents=True, exist_ok=True)
    destination.write_bytes(uploaded_file.getbuffer())
    return destination


def save_uploads(uploaded_files, destination: Path) -> list[Path]:
    paths = []
    destination.mkdir(parents=True, exist_ok=True)

    for uploaded in uploaded_files or []:
        paths.append(
            save_upload(
                uploaded,
                destination / uploaded.name
            )
        )

    return paths


st.title("pyPAVF")

st.caption(
    "Build an approval ledger from PAVFs and approval emails"
)

with st.sidebar:

    st.header("Approval Ledger Inputs")

    pavf_uploads = st.file_uploader(
        "PAVF Files",
        type=["xlsx", "xlsm", "csv"],
        accept_multiple_files=True,
    )

    email_uploads = st.file_uploader(
        "Approval Emails",
        type=["eml", "txt"],
        accept_multiple_files=True,
    )

    run_button = st.button(
        "Build Approval Ledger",
        type="primary",
        width="stretch",
    )

col1, col2 = st.columns(2)

col1.metric(
    "PAVFs",
    len(pavf_uploads or [])
)

col2.metric(
    "Emails",
    len(email_uploads or [])
)

st.divider()

if (
    not run_button
    and "ledger_df" not in st.session_state
):
    st.info(
        "Load PAVF files and approval emails, then select "
        "**Build Approval Ledger**."
    )
    st.stop()

if not pavf_uploads:
    st.error("Please load at least one PAVF file.")
    st.stop()

if not email_uploads:
    st.warning(
        "No email evidence was loaded. The approval ledger "
        "can still be built, but approvals will not be available."
    )

try:

    if run_button:

        with st.status(
            "Building approval ledger...",
            expanded=True,
        ) as status:

            with tempfile.TemporaryDirectory(
                prefix="pypavf_"
            ) as temp_name:

                work_dir = Path(temp_name)

                pavf_dir = work_dir / "pavf"
                email_dir = work_dir / "emails"

                st.write("Reading PAVF files...")

                save_uploads(
                    pavf_uploads,
                    pavf_dir,
                )

                pavf_records = read_pavf_records(
                    pavf_dir
                )

                st.write("Reading email evidence...")

                save_uploads(
                    email_uploads,
                    email_dir,
                )

                email_evidence = (
                    read_email_evidence(email_dir)
                    if email_uploads
                    else []
                )

                st.write(
                    "Matching PAVFs to approval emails..."
                )

                matcher = ReconciliationMatcher()

                results = []

                for pavf in pavf_records:

                    approval = matcher._find_thread_approval(
                        breakdown=pavf,
                        pavf=pavf,
                        evidence_list=email_evidence,
                    )

                    results.append(
                        {
                            "Date": pavf.problem_date,
                            "Part Number": pavf.part_number,
                            "Cardboard Boxes": pavf.box_qty,
                            "PO Number": pavf.po_number,
                            "Approval": getattr(
                                approval.approval_status,
                                "value",
                                "NOT_FOUND",
                            ),
                            "Approval Date": approval.approval_date,
                            "Approval Email": approval.evidence_file,
                        }
                    )

                ledger_df = pd.DataFrame(results)

                if not ledger_df.empty:

                    ledger_df = ledger_df.sort_values(
                        by=["Date", "Part Number"]
                    )

                st.session_state["ledger_df"] = ledger_df

                status.update(
                    label="Approval ledger complete",
                    state="complete",
                    expanded=False,
                )

    ledger_df = st.session_state.get(
        "ledger_df",
        pd.DataFrame()
    )

    if ledger_df.empty:
        st.stop()

    approved_count = (
        ledger_df["Approval"]
        .eq("APPROVED")
        .sum()
    )

    missing_count = (
        len(ledger_df)
        - approved_count
    )

    st.subheader("Approval Ledger")

    summary_cols = st.columns(3)

    summary_cols[0].metric(
        "PAVFs",
        len(ledger_df)
    )

    summary_cols[1].metric(
        "Approved",
        approved_count
    )

    summary_cols[2].metric(
        "Missing Approval",
        missing_count
    )

    st.dataframe(
        ledger_df,
        width="stretch",
        hide_index=True,
    )

    #
    # Build filename
    #
    min_date = pd.to_datetime(
        ledger_df["Date"].min()
    )

    max_date = pd.to_datetime(
        ledger_df["Date"].max()
    )

    if min_date.date() == max_date.date():

        file_name = (
            f"Approval_Ledger_"
            f"{min_date:%Y_%m_%d}.xlsx"
        )

    else:

        file_name = (
            f"Approval_Ledger_"
            f"{min_date:%Y_%m_%d}"
            f"_to_"
            f"{max_date:%Y_%m_%d}.xlsx"
        )

    #
    # Create Excel in memory
    #
    export_df = ledger_df.copy()

    if "Approval Date" in export_df.columns:

        export_df["Approval Date"] = (
            pd.to_datetime(
                export_df["Approval Date"],
                errors="coerce",
                utc=True,
            )
            .dt.tz_localize(None)
        )

    buffer = BytesIO()

    export_df.to_excel(
        buffer,
        index=False,
        engine="openpyxl",
    )

    buffer.seek(0)

    st.download_button(
        "Download Approval Ledger",
        data=buffer.getvalue(),
        file_name=file_name,
        mime=(
            "application/vnd.openxmlformats-"
            "officedocument.spreadsheetml.sheet"
        ),
        type="primary",
    )

except Exception as exc:

    st.error(
        "The approval ledger could not be created."
    )

    st.exception(exc)