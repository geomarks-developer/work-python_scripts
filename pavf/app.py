"""pyPAVF Streamlit application."""
from __future__ import annotations

from pathlib import Path
import shutil
import tempfile

import streamlit as st

from modules.breakdown_reader import BreakdownReader
from modules.pavf_reader import read_pavf_records
from modules.email_evidence_reader import read_email_evidence
from modules.reconciliation_matcher import ReconciliationMatcher
from modules.reconciliation_report import create_reconciliation_report


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
        paths.append(save_upload(uploaded, destination / uploaded.name))
    return paths


def build_breakdown_records(path: Path):
    """Convert BreakdownReader dictionaries to project BreakdownRecord objects."""
    from modules.models import BreakdownRecord

    records = []
    for row in BreakdownReader(path).read():
        allowed = BreakdownRecord.__dataclass_fields__.keys()
        values = {key: value for key, value in row.items() if key in allowed}
        records.append(BreakdownRecord(**values))
    return records


st.title("pyPAVF")
st.caption("PAVF reimbursement reconciliation")

with st.sidebar:
    st.header("Reconciliation Inputs")

    breakdown_upload = st.file_uploader(
        "Reimbursement Breakdown",
        type=["xlsx", "xlsm"],
        accept_multiple_files=False,
    )

    pavf_uploads = st.file_uploader(
        "PAVF Files",
        type=["xlsx", "xlsm", "csv"],
        accept_multiple_files=True,
    )

    email_uploads = st.file_uploader(
        "Email Evidence",
        type=["eml", "txt"],
        accept_multiple_files=True,
    )

    run_button = st.button(
        "Run Reconciliation",
        type="primary",
        use_container_width=True,
    )

col1, col2, col3 = st.columns(3)
col1.metric("Breakdown", "Ready" if breakdown_upload else "Not loaded")
col2.metric("PAVFs", len(pavf_uploads or []))
col3.metric("Emails", len(email_uploads or []))

st.divider()

if not run_button:
    st.info(
        "Load a reimbursement breakdown, the applicable PAVF files, and "
        "email evidence. Then select **Run Reconciliation**."
    )
    st.stop()

if breakdown_upload is None:
    st.error("Please load a reimbursement breakdown workbook.")
    st.stop()

if not pavf_uploads:
    st.error("Please load at least one PAVF file.")
    st.stop()

if not email_uploads:
    st.warning(
        "No email evidence was loaded. The reconciliation can still run, "
        "but approvals will not be available."
    )

try:
    with st.status("Running PAVF reconciliation...", expanded=True) as status:
        with tempfile.TemporaryDirectory(prefix="pypavf_") as temp_name:
            work_dir = Path(temp_name)
            breakdown_dir = work_dir / "breakdown"
            pavf_dir = work_dir / "pavf"
            email_dir = work_dir / "emails"

            st.write("Reading reimbursement breakdown...")
            breakdown_path = save_upload(
                breakdown_upload,
                breakdown_dir / breakdown_upload.name,
            )
            breakdown_records = build_breakdown_records(breakdown_path)

            st.write("Reading PAVF files...")
            save_uploads(pavf_uploads, pavf_dir)
            pavf_records = read_pavf_records(pavf_dir)

            st.write("Reading email evidence...")
            save_uploads(email_uploads, email_dir)
            email_evidence = (
                read_email_evidence(email_dir)
                if email_uploads
                else []
            )

            st.write("Matching records...")
            results = ReconciliationMatcher().reconcile(
                breakdown_records=breakdown_records,
                pavf_records=pavf_records,
                email_evidence=email_evidence,
            )

            st.write("Building Excel report...")
            report_path = OUTPUT_DIR / "PAVF_Reconciliation_Report.xlsx"
            create_reconciliation_report(results, report_path)

            status.update(
                label="Reconciliation complete",
                state="complete",
                expanded=False,
            )

    counts = {}
    for result in results:
        status_value = getattr(result.status, "value", str(result.status))
        counts[status_value] = counts.get(status_value, 0) + 1

    st.subheader("Results")
    summary_cols = st.columns(4)
    summary_cols[0].metric("Records", len(results))
    summary_cols[1].metric("Approved", counts.get("APPROVED", 0))
    summary_cols[2].metric("PAVF Not Found", counts.get("PAVF_NOT_FOUND", 0))
    summary_cols[3].metric(
        "Needs Review",
        sum(
            count
            for name, count in counts.items()
            if name not in {"APPROVED", "PAVF_NOT_FOUND"}
        ),
    )

    table_rows = []
    for result in results:
        breakdown = result.breakdown
        pavf = result.pavf
        approval = result.approval
        table_rows.append({
            "Date": getattr(breakdown, "date", None),
            "Part Number": getattr(breakdown, "part_number", ""),
            "Breakdown Boxes": getattr(breakdown, "box_qty", None),
            "PAVF Boxes": getattr(pavf, "box_qty", None) if pavf else None,
            "Approval": (
                getattr(getattr(approval, "approval_status", None), "value", "")
                if approval else ""
            ),
            "Status": getattr(result.status, "value", str(result.status)),
            "Notes": "; ".join(getattr(result, "notes", []) or []),
        })

    st.dataframe(
        table_rows,
        use_container_width=True,
        hide_index=True,
    )

    report_bytes = report_path.read_bytes()
    st.download_button(
        "Download Reconciliation Report",
        data=report_bytes,
        file_name=report_path.name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )

except Exception as exc:
    st.error("The reconciliation could not be completed.")
    st.exception(exc)
