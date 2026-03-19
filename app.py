from __future__ import annotations
import logging
import os
import shutil
import tempfile
import time
import zipfile
from concurrent.futures import ThreadPoolExecutor, as_completed
from typing import Optional
import pandas as pd
import streamlit as st
MAX_FILE_SIZE_MB: int = 50
MAX_FILE_SIZE_BYTES: int = MAX_FILE_SIZE_MB * 1024 * 1024
LOG_FILE: str = "conversion.log"
logging.basicConfig(
    filename=LOG_FILE,
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
logger = logging.getLogger(__name__)
if "results" not in st.session_state:
    st.session_state["results"]: dict = {}
if "temp_dirs" not in st.session_state:
    st.session_state["temp_dirs"]: list[str] = []
def safe_filename(name: str) -> str:
    return (
        "".join(c if c.isalnum() or c in (" ", "_", "-") else "_" for c in name)
        .strip()
        .replace(" ", "_")
    )
def cleanup_temp_dirs(paths: list[str]) -> None:
    for path in paths:
        shutil.rmtree(path, ignore_errors=True)
def process_sheet(
    xls: pd.ExcelFile,
    base_name: str,
    sheet_name: str,
    output_dir: str,
) -> tuple[Optional[pd.DataFrame], str]:
    start = time.perf_counter()
    try:
        df: pd.DataFrame = pd.read_excel(xls, sheet_name=sheet_name)
        if df.empty:
            msg = f"SKIPPED  | {base_name} | {sheet_name} | empty sheet"
            logger.info(msg)
            return None, msg
        file_name = f"{base_name}_{safe_filename(sheet_name)}.csv"
        path = os.path.join(output_dir, file_name)
        df.to_csv(path, index=False, encoding="utf-8-sig")
        elapsed = time.perf_counter() - start
        msg = (
            f"SAVED    | {base_name} | {sheet_name} | "
            f"rows={len(df)} | cols={len(df.columns)} | {elapsed:.2f}s → {file_name}"
        )
        logger.info(msg)
        return df, msg
    except Exception as exc:  # noqa: BLE001
        msg = f"ERROR    | {base_name} | {sheet_name} | {exc}"
        logger.error(msg)
        return None, msg
def stack_sheets(
    sheet_frames: list[tuple[str, pd.DataFrame]],
    base_name: str,
    temp_dir: str,
) -> Optional[str]:
    if not sheet_frames:
        return None
    pieces: list[pd.DataFrame] = []
    for idx, (sheet_name, df) in enumerate(sheet_frames):
        tagged = df.copy()
        tagged.insert(0, "_sheet", sheet_name)
        pieces.append(tagged)
        if idx < len(sheet_frames) - 1:
            separator = pd.DataFrame(
                [[""] * len(tagged.columns)], columns=tagged.columns
            )
            pieces.append(separator)
    if not pieces:
        return None
    stacked = pd.concat(pieces, ignore_index=True)
    stacked_path = os.path.join(temp_dir, f"{base_name}_stacked.csv")
    stacked.to_csv(stacked_path, index=False, encoding="utf-8-sig")
    logger.info(
        f"STACKED  | {base_name} | sheets={len(sheet_frames)} | total_rows={len(stacked)}"
    )
    return stacked_path
def convert_file(
    excel_path: str,
    file_name: str,
    progress_bar: st.delta_generator.DeltaGenerator,
    status_text: st.delta_generator.DeltaGenerator,
    do_stack: bool,
) -> tuple[str, list[str]]:
    base_name = safe_filename(os.path.splitext(file_name)[0])
    temp_dir = tempfile.mkdtemp()
    st.session_state["temp_dirs"].append(temp_dir)
    output_dir = os.path.join(temp_dir, "csvs")
    os.makedirs(output_dir, exist_ok=True)
    xls = pd.ExcelFile(excel_path)
    sheets: list[str] = xls.sheet_names
    total = len(sheets)
    completed = 0
    logs: list[str] = []
    sheet_results: dict[str, Optional[pd.DataFrame]] = {s: None for s in sheets}
    with ThreadPoolExecutor() as executor:
        future_to_sheet: dict = {
            executor.submit(process_sheet, xls, base_name, sheet, output_dir): sheet
            for sheet in sheets
        }
        for future in as_completed(future_to_sheet):
            sheet_name = future_to_sheet[future]
            df, log = future.result()
            logs.append(log)
            sheet_results[sheet_name] = df
            completed += 1
            progress_bar.progress(completed / total)
            status_text.text(f"Processed {completed}/{total} sheets")
    stacked_path: Optional[str] = None
    if do_stack:
        ordered_frames = [
            (s, sheet_results[s])
            for s in sheets
            if sheet_results[s] is not None and not sheet_results[s].empty
        ]
        stacked_path = stack_sheets(ordered_frames, base_name, temp_dir)
        if stacked_path:
            logs.append(f"STACKED  | Written → {os.path.basename(stacked_path)}")
        else:
            logs.append("STACKED  | Skipped – all sheets were empty")
    zip_path = os.path.join(temp_dir, f"{base_name}.zip")
    with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for csv_file in sorted(os.listdir(output_dir)):
            zf.write(os.path.join(output_dir, csv_file), arcname=csv_file)
        if stacked_path:
            zf.write(stacked_path, arcname=os.path.basename(stacked_path))
    return zip_path, logs
st.set_page_config(
    page_title="Excel → CSV Converter",
    page_icon="📊",
    layout="centered",
    initial_sidebar_state="collapsed",
)
st.markdown(
    unsafe_allow_html=True,
)
st.title("📊 Excel → CSV Converter")
st.caption("Upload one or more `.xlsx` / `.xls` files — each sheet becomes a CSV.")
with st.sidebar:
    st.header("⚙️ Options")
    do_stack = st.checkbox(
        "Stack all sheets into one CSV",
        value=False,
        help=(
            "Appends every sheet one after another into a single `_stacked.csv`.  \n"
            "A `_sheet` column is added so you always know which sheet each row came from.  \n"
            "A blank separator row is inserted between sheets for readability."
        ),
    )
    show_preview = st.checkbox("Preview first 5 rows before converting", value=True)
    st.markdown("---")
    st.markdown(f"**Max file size:** `{MAX_FILE_SIZE_MB} MB`")
uploaded_files = st.file_uploader(
    "Drop Excel files here",
    type=["xlsx", "xls"],
    accept_multiple_files=True,
    label_visibility="collapsed",
)
if uploaded_files:
    st.success(f"**{len(uploaded_files)}** file(s) ready.")
    if show_preview:
        st.markdown("### 🔍 Data Preview")
        for uf in uploaded_files:
            if uf.size > MAX_FILE_SIZE_BYTES:
                st.warning(f"⚠️ `{uf.name}` exceeds {MAX_FILE_SIZE_MB} MB — will be skipped.")
                continue
            with st.expander(f"`{uf.name}`"):
                try:
                    xls_preview = pd.ExcelFile(uf)
                    sheet_choice = st.selectbox(
                        "Sheet",
                        xls_preview.sheet_names,
                        key=f"preview_sheet_{uf.name}",
                    )
                    preview_df = pd.read_excel(xls_preview, sheet_name=sheet_choice, nrows=5)
                    st.dataframe(preview_df, use_container_width=True)
                    st.caption(
                        f"Sheets in this file: **{len(xls_preview.sheet_names)}** — "
                        f"showing first 5 rows of `{sheet_choice}`"
                    )
                except Exception as exc:
                    st.error(f"Could not preview `{uf.name}`: {exc}")
                finally:
                    uf.seek(0)  # rewind buffer for conversion
    if st.button("⚡ Convert All", type="primary"):
        for uf in uploaded_files:
            if uf.size > MAX_FILE_SIZE_BYTES:
                st.warning(f"⏭️ Skipping `{uf.name}` — exceeds {MAX_FILE_SIZE_MB} MB.")
                continue
            st.divider()
            st.subheader(f"📄 {uf.name}")
            with tempfile.NamedTemporaryFile(
                delete=False, suffix=os.path.splitext(uf.name)[1]
            ) as tmp:
                tmp.write(uf.getbuffer())
                tmp_path = tmp.name
            col_prog, col_status = st.columns([3, 1])
            with col_prog:
                progress_bar = st.progress(0)
            with col_status:
                status_text = st.empty()
            try:
                zip_path, logs = convert_file(
                    excel_path=tmp_path,
                    file_name=uf.name,
                    progress_bar=progress_bar,
                    status_text=status_text,
                    do_stack=do_stack,
                )
                st.session_state["results"][uf.name] = {
                    "zip_path": zip_path,
                    "logs": logs,
                }
            except Exception as exc:
                st.error(f"Failed to convert `{uf.name}`: {exc}")
                logger.exception(f"TOP-LEVEL ERROR | {uf.name} | {exc}")
            finally:
                try:
                    os.unlink(tmp_path)
                except OSError:
                    pass
        st.success("✅ All files processed.")
if st.session_state["results"]:
    st.divider()
    st.header("📦 Download Results")
    for file_name, data in st.session_state["results"].items():
        st.subheader(f"`{file_name}`")
        with st.expander("📋 Processing logs"):
            for log in data["logs"]:
                if "ERROR" in log:
                    st.markdown(f":red[{log}]")
                elif "SKIPPED" in log:
                    st.markdown(f":orange[{log}]")
                elif "STACKED" in log:
                    st.markdown(f":blue[{log}]")
                else:
                    st.text(log)
        try:
            with open(data["zip_path"], "rb") as fh:
                st.download_button(
                    label=f"⬇️  Download  {os.path.basename(data['zip_path'])}",
                    data=fh,
                    file_name=os.path.basename(data["zip_path"]),
                    mime="application/zip",
                    key=f"dl_{file_name}",
                )
        except FileNotFoundError:
            st.error(
                f"ZIP for `{file_name}` not found — it may have been cleaned up. "
                "Please re-convert."
            )
    st.divider()
    if st.button("🗑️  Clear Results & Free Disk Space"):
        cleanup_temp_dirs(st.session_state["temp_dirs"])
        st.session_state["results"] = {}
        st.session_state["temp_dirs"] = []
        st.success("Cleared — all temporary files deleted.")
        st.rerun()
