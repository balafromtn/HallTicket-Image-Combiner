# app.py
from __future__ import annotations

import io
import shutil
import tempfile
import zipfile
from datetime import date
from pathlib import Path

import streamlit as st
from openpyxl import Workbook

from single_hallticket import generate_single
from bulk_hallticket import generate_bulk

APP_DIR = Path(__file__).resolve().parent
OUTPUT_ROOT = APP_DIR / "generated"
SUPPORTED_PHOTOS = {".png", ".jpg", ".jpeg", ".webp", ".bmp"}
MAX_FILE_SIZE_BYTES = 1 * 1024 * 1024  # 1MB limit

st.set_page_config(
    page_title="GACS HallTicket Image Combiner",
    page_icon="🎓",
    layout="wide",
    initial_sidebar_state="collapsed",
)

st.markdown("""
<style>
.stApp { background: #f5f8fc; }
[data-testid="stHeader"] { background: rgba(245,248,252,.92); }
.block-container { max-width: 1120px; padding-top: 2rem; padding-bottom: 3rem; }
.hero { padding: 1.6rem 1.8rem; border: 1px solid #dce5f0; border-radius: 20px;
background: linear-gradient(135deg,#fff 0%,#edf4ff 100%); margin-bottom: 1.25rem; }
.hero h1 { color: #183153; font-size: 2rem; margin: 0 0 .4rem 0; }
.hero p { color: #5b6b82; margin: 0; font-size: 1rem; }
div.stButton > button, div.stDownloadButton > button { border-radius: 10px; font-weight: 600; min-height: 2.7rem; }
[data-testid="stFileUploader"] { background: #fff; border: 1px dashed #c8d6e8; border-radius: 12px; padding: .5rem; }
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="hero">
<h1>🎓 GACS HallTicket Image Combiner</h1>
<p>Generate student hall tickets with embedded photographs, one at a time or in bulk.</p>
</div>
""", unsafe_allow_html=True)

if "mode" not in st.session_state:
    st.session_state.mode = "Single Hall Ticket"

def set_mode(mode_name):
    st.session_state.mode = mode_name

left, right = st.columns(2)
with left:
    st.button("📄  Single Hall Ticket", use_container_width=True,
              type="primary" if st.session_state.mode == "Single Hall Ticket" else "secondary",
              on_click=set_mode, args=("Single Hall Ticket",))
with right:
    st.button("📂  Multiple Hall Tickets", use_container_width=True,
              type="primary" if st.session_state.mode == "Multiple Hall Tickets" else "secondary",
              on_click=set_mode, args=("Multiple Hall Tickets",))

st.write("")

def make_zip(files: list[Path], archive_path: Path, base_path: Path) -> None:
    with zipfile.ZipFile(archive_path, "w", zipfile.ZIP_DEFLATED) as archive:
        for file_path in files:
            if file_path.exists() and file_path.is_file():
                archive.write(file_path, arcname=str(file_path.relative_to(base_path)))

def create_sample_excel() -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "Students"
    ws.append(["RegNo", "DOB"])
    ws.append(["24UCS250904", "2005-08-15"])
    ws.append(["24UCS250905", "2004-12-10"])
    
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()

if st.session_state.mode == "Single Hall Ticket":
    st.subheader("Generate one hall ticket")
    st.caption("Enter the student's portal details and upload their photograph.")
    with st.form("single_form"):
        c1, c2 = st.columns(2)
        with c1:
            register_no = st.text_input("Register number", placeholder="e.g. 24UCS250904")
        with c2:
            dob = st.date_input("Date of birth", format="MM/DD/YYYY", min_value=date(2000, 1, 1), max_value=date(2040, 12, 31))
        photo_file = st.file_uploader("Student photo (Max 1MB)", type=["png", "jpg", "jpeg", "webp", "bmp"])
        submit_single = st.form_submit_button("Generate hall ticket", type="primary", use_container_width=True)

    if submit_single:
        if not register_no.strip():
            st.error("Enter the register number.")
        elif photo_file is None:
            st.error("Upload the student's photo.")
        elif photo_file.size > MAX_FILE_SIZE_BYTES:
            st.error("The uploaded photo exceeds the 1MB size limit.")
        else:
            temp_dir = Path(tempfile.mkdtemp(prefix="gacs-single-"))
            try:
                photo_path = temp_dir / Path(photo_file.name).name
                photo_path.write_bytes(photo_file.getvalue())
                with st.spinner("Connecting to the college portal and generating files..."):
                    result = generate_single(register_no.strip(), dob.strftime("%Y-%m-%d"), photo_path, OUTPUT_ROOT)
                st.success("Hall ticket generated successfully.")
                st.write(f"**Register number:** {result['register_no']}")
                c1, c2 = st.columns(2)
                with c1:
                    st.download_button("Download PDF", result["pdf_path"].read_bytes(),
                        file_name=result["pdf_path"].name, mime="application/pdf", use_container_width=True)
                with c2:
                    st.download_button("Download HTML", result["html_path"].read_bytes(),
                        file_name=result["html_path"].name, mime="text/html", use_container_width=True)
                st.caption(f"Local files saved in `{OUTPUT_ROOT}`.")
            except Exception as exc:
                st.error(f"Generation failed: {type(exc).__name__}: {exc}")
            finally:
                shutil.rmtree(temp_dir, ignore_errors=True)

else:
    st.subheader("Generate multiple hall tickets")
    st.caption("Upload an Excel/CSV roster and the corresponding student photo files.")
    st.info("The roster must contain register-number and DOB columns. Each photo filename should match its student's register number, e.g. 24UCS250904.jpg.")
    
    # Download sample button
    st.download_button(
        label="⬇️ Download Sample Input Excel File",
        data=create_sample_excel(),
        file_name="sample_roster.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    
    with st.form("bulk_form"):
        spreadsheet = st.file_uploader("Excel or CSV file (Max 1MB)", type=["xlsx", "csv"], key="bulk_sheet")
        photos = st.file_uploader("Student photos (Max 1MB per file)", type=["png", "jpg", "jpeg", "webp", "bmp"],
                                  accept_multiple_files=True, key="bulk_photos")
        submit_bulk = st.form_submit_button("Start bulk generation", type="primary", use_container_width=True)

    if submit_bulk:
        if spreadsheet is None:
            st.error("Upload an Excel or CSV file.")
        elif spreadsheet.size > MAX_FILE_SIZE_BYTES:
            st.error("The spreadsheet exceeds the 1MB size limit.")
        elif not photos:
            st.error("Upload the student photo files.")
        else:
            oversized_photos = [p.name for p in photos if p.size > MAX_FILE_SIZE_BYTES]
            if oversized_photos:
                st.error(f"The following photos exceed the 1MB size limit: {', '.join(oversized_photos)}")
            else:
                temp_dir = Path(tempfile.mkdtemp(prefix="gacs-bulk-"))
                try:
                    input_path = temp_dir / Path(spreadsheet.name).name
                    input_path.write_bytes(spreadsheet.getvalue())
                    photo_dir = temp_dir / "photos"
                    photo_dir.mkdir()
                    for photo in photos:
                        safe_name = Path(photo.name).name
                        if Path(safe_name).suffix.lower() in SUPPORTED_PHOTOS:
                            (photo_dir / safe_name).write_bytes(photo.getvalue())

                    progress = st.progress(0, text="Preparing records...")
                    latest_status = st.empty()
                    rows_area = st.empty()
                    processed_rows = []

                    def on_progress(position, total, row):
                        progress.progress(position / total, text=f"Processed {position} of {total}")
                        processed_rows.append(row)
                        latest_status.write(f"**Latest:** `{row['regno'] or '(blank register number)'}` · {row['status']}")
                        rows_area.dataframe(processed_rows, use_container_width=True, hide_index=True)

                    with st.spinner("Generating hall tickets. Keep this page open until processing finishes."):
                        result = generate_bulk(input_path, photo_dir, OUTPUT_ROOT, on_progress)

                    c1, c2, c3 = st.columns(3)
                    c1.metric("Total records", result["total"])
                    c2.metric("Successful", result["succeeded"])
                    c3.metric("Failed", result["failed"])
                    st.success("Bulk processing finished.")
                    st.download_button("Download bulk report (CSV)", result["report_path"].read_bytes(),
                        file_name="bulk_report.csv", mime="text/csv", use_container_width=True)

                    generated_files = list(result["html_dir"].glob("*.html")) + list(result["pdf_dir"].glob("*.pdf")) + [result["report_path"]]
                    zip_path = OUTPUT_ROOT / "halltickets_bundle.zip"
                    make_zip(generated_files, zip_path, OUTPUT_ROOT)
                    st.download_button("Download all generated files (ZIP)", zip_path.read_bytes(),
                        file_name="gacs_halltickets.zip", mime="application/zip", use_container_width=True)
                    st.caption(f"Generated files are also saved locally in `{OUTPUT_ROOT}`.")
                except Exception as exc:
                    st.error(f"Bulk processing could not finish: {type(exc).__name__}: {exc}")
                finally:
                    shutil.rmtree(temp_dir, ignore_errors=True)