
import io
import pandas as pd
import streamlit as st

# Import library functions
from Malaria_Data_Cleaning import clean_malaria_data
from Malaria_Indicator import compute_indicators

# ---------- Utilities for UI-only presentation ----------
def put_comment_first(df: pd.DataFrame) -> pd.DataFrame:
    return df if "COMMENT" not in df.columns else df[["COMMENT"] + [c for c in df.columns if c != "COMMENT"]]

def put_comment_last(df: pd.DataFrame) -> pd.DataFrame:
    return df if "COMMENT" not in df.columns else df[[c for c in df.columns if c != "COMMENT"] + ["COMMENT"]]

def sanitize_sheet_name(name: str) -> str:
    # Excel sheet names: <=31 chars, no []:*?/\
    bad = '[]:*?/\\'
    table = {ord(ch): ' ' for ch in bad}
    name = str(name).translate(table).strip() or "Sheet"
    return name[:31]

# ---------- Chip-free multi-select dropdown ----------
def multiselect_dropdown(label: str, options: list[str], key: str = "ms_dropdown"):
    sel_key  = f"{key}_selected"
    snap_key = f"{key}_options_snapshot"

    if sel_key not in st.session_state:
        st.session_state[sel_key] = set()
    if snap_key not in st.session_state:
        st.session_state[snap_key] = tuple()

    snapshot = tuple(options)
    if st.session_state[snap_key] != snapshot:
        st.session_state[snap_key] = snapshot
        st.session_state[sel_key] = set()

    def render_checkbox_list():
        cols = st.columns([1, 1])
        with cols[0]:
            if st.button("Select all", key=f"{key}_all"):
                st.session_state[sel_key] = set(options)
        with cols[1]:
            if st.button("Clear", key=f"{key}_clear"):
                st.session_state[sel_key].clear()

        for opt in options:
            checked = opt in st.session_state[sel_key]
            if st.checkbox(opt, value=checked, key=f"{key}_{opt}"):
                st.session_state[sel_key].add(opt)
            else:
                st.session_state[sel_key].discard(opt)

    if hasattr(st, "popover"):
        with st.popover(label):
            render_checkbox_list()
    else:
        with st.expander(label, expanded=False):
            render_checkbox_list()

    selected = [o for o in options if o in st.session_state[sel_key]]
    st.caption(f"Selected: {len(selected)}")
    return selected

# Optional helper: ensure SCREENING_DATE strings look like 'DD-MMM-YY'
import re as _re
from datetime import datetime as _dt
_ISO_DATE = _re.compile(r'^\d{4}[-/]\d{2}[-/]\d{2}(?: \d{2}:\d{2}:\d{2})?$')
def force_screening_date_strings(df: pd.DataFrame) -> pd.DataFrame:
    date_col = next((c for c in df.columns if c.strip().lower() == "screening_date"), None)
    if not date_col:
        return df
    def _fmt(v):
        if isinstance(v, (pd.Timestamp, _dt)):
            return pd.to_datetime(v).strftime('%d-%b-%y')
        if isinstance(v, str) and _ISO_DATE.match(v.strip()):
            dt = pd.to_datetime(v, errors='coerce')
            if pd.notna(dt):
                return dt.strftime('%d-%b-%y')
        return v
    df[date_col] = df[date_col].map(_fmt)
    return df

# ----------------- App -----------------
st.set_page_config(page_title="🦟 Malaria Apps", layout="wide")
st.title("🦟 Malaria Apps")
mode = st.radio("Choose a tool", ["Data Cleaning", "Indicator Calculation"], horizontal=True)

uploaded_file = st.file_uploader("Upload your malaria Excel file", type=["xlsx", "xls"])

if uploaded_file is None:
    st.info("Upload an Excel file to begin.")
    st.stop()

# Reset dropdown state when a different file is uploaded
file_key = f"{uploaded_file.name}:{uploaded_file.size}"
if st.session_state.get("uploaded_file_key") != file_key:
    st.session_state["uploaded_file_key"] = file_key
    for k in list(st.session_state.keys()):
        if k.endswith("_selected") or k.endswith("_options_snapshot"):
            del st.session_state[k]

# Read Excel and list sheets
try:
    xls = pd.ExcelFile(uploaded_file)
    sheet_names = xls.sheet_names
except Exception as e:
    st.error("❌ Could not read the Excel file. Please check the format.")
    st.exception(e)
    st.stop()

selected_sheets = multiselect_dropdown(
    "Select sheet(s) to process",
    sheet_names,
    key="sheets"
)
if not selected_sheets:
    st.warning("Select at least one sheet to continue.")
    st.stop()

if mode == "Data Cleaning":
    cleaned_by_sheet = {}
    for sheet in selected_sheets:
        raw_df  = xls.parse(sheet_name=sheet)
        cleaned = clean_malaria_data(raw_df)
        cleaned = force_screening_date_strings(cleaned)
        cleaned_by_sheet[sheet] = {
            "display": put_comment_first(cleaned.copy()),
            "file":    put_comment_last(cleaned.copy()),
        }

    # Error previews & summary
    st.subheader("⚠️ Error Rows Preview (by sheet)")
    tabs = st.tabs(selected_sheets)
    def _build_error_summary_by_column(error_series: pd.Series) -> pd.DataFrame:
        counts = {}
        for comment in error_series.astype(str):
            parts = [p.strip() for p in comment.split(";") if p.strip()]
            for p in parts:
                if "[" in p and "]" in p:
                    col = p.split("[", 1)[1].split("]", 1)[0]
                    counts[col] = counts.get(col, 0) + 1
        if not counts:
            return pd.DataFrame(columns=["Column", "Error Count"])
        return (
            pd.DataFrame(list(counts.items()), columns=["Column", "Error Count"])
            .sort_values("Error Count", ascending=False)
            .reset_index(drop=True)
        )

    for tab, sheet in zip(tabs, selected_sheets):
        with tab:
            display_df = cleaned_by_sheet[sheet]["display"]
            error_df = display_df[display_df.get("COMMENT", "").astype(str).str.strip() != ""]
            if error_df.empty:
                st.success(f"✅ {sheet}: No data quality issues found.")
            else:
                st.dataframe(error_df.head(50), use_container_width=True)
                st.info(f"{sheet} — Total error rows: {len(error_df)}")
                summary_df = _build_error_summary_by_column(error_df["COMMENT"])
                if not summary_df.empty:
                    st.markdown("**📊 Error Summary by Column**")
                    st.dataframe(summary_df, use_container_width=True)

    # Download cleaned workbook
    st.subheader("⬇️ Download")
    try:
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
            for sheet in selected_sheets:
                cleaned_by_sheet[sheet]["file"].to_excel(
                    writer, index=False, sheet_name=sanitize_sheet_name(sheet)
                )
        buffer.seek(0)
        st.download_button(
            label=f"📥 Download Cleaned Workbook ({len(selected_sheets)} sheets)",
            data=buffer,
            file_name="malaria_cleaned_selected.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    except Exception as e:
        st.error("❌ Failed to generate the cleaned workbook.")
        st.exception(e)

else:  # Indicator Calculation
    outputs = {}
    errors = []
    for sheet in selected_sheets:
        try:
            raw_df = xls.parse(sheet_name=sheet)
            out_df = compute_indicators(raw_df.copy())
            outputs[sheet] = out_df
        except Exception as e:
            errors.append((sheet, e))

    if outputs:
        st.subheader("👀 Preview (first 50 rows)")
        tabs = st.tabs(list(outputs.keys()))
        for tab, sheet in zip(tabs, outputs.keys()):
            with tab:
                st.caption(f"Sheet: {sheet}")
                st.dataframe(outputs[sheet].head(50), use_container_width=True)

    if errors:
        st.warning("Some sheets failed to process:")
        for sheet, e in errors:
            with st.expander(f"Details: {sheet}"):
                st.exception(e)

    # Download indicators workbook
    if outputs:
        st.subheader("⬇️ Download")
        try:
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
                for sheet, df_out in outputs.items():
                    df_out.to_excel(writer, index=False, sheet_name=sanitize_sheet_name(sheet))
            buffer.seek(0)
            st.download_button(
                label=f"📥 Download Indicators Workbook ({len(outputs)} sheet{'s' if len(outputs)>1 else ''})",
                data=buffer,
                file_name="malaria_indicators_selected.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        except Exception as e:
            st.error("❌ Failed to generate the indicators workbook.")
            st.exception(e)
