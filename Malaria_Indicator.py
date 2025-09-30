import pandas as pd
import numpy as np
import io
import streamlit as st



# --- helpers ---
def first_col(df, candidates):
    for c in candidates:
        if c in df.columns:
            return df[c]
    raise KeyError(f"Column not found among: {candidates}")

def move_after(df, col, after):
    cols = list(df.columns)
    cols.insert(cols.index(after) + 1, cols.pop(cols.index(col)))
    return df[cols]

def norm_str(s):
    return s.astype(str).str.strip().str.casefold()

def compute_indicators(df: pd.DataFrame) -> pd.DataFrame:
    """
    Apply Malaria indicator logic to a single sheet's DataFrame.
    Returns the DataFrame with indicator columns added/reordered.
    """
    
    # Normalize blanks (for emptiness checks) and compute a single base mask
    df_norm = df.replace(r'^\s*$', pd.NA, regex=True)
    base_check_cols = [c for c in df.columns if c not in {"SR_NO", "COMMENT"}]
    is_blank_or_sr_only = ~df_norm[base_check_cols].notna().any(axis=1)

    # --- CM Other - 2 (D) ---
    retesting = first_col(df, ["RETESTING", "Retesting"])
    is_repeat = norm_str(retesting).eq("repeat").fillna(False)
    df["CM Other - 2 (D)"] = np.where(is_repeat | is_blank_or_sr_only, np.nan, 1)
    df = move_after(df, "CM Other - 2 (D)", "COMMENT")

    # --- CM Other - 2 (N) ---
    txg = first_col(df, ["TX_GUIDELINE"])
    is_tx_no = norm_str(txg).isin({"no", "refer & no"})
    blank_due_to_D = df["CM Other - 2 (D)"].isna()

    df["CM Other - 2 (N)"] = pd.Series(
        np.where(is_tx_no | is_blank_or_sr_only | blank_due_to_D, pd.NA, 1),
        index=df.index,
        dtype="Int64",
    )
    df = move_after(df, "CM Other - 2 (N)", "CM Other - 2 (D)")

    # --- CM Other - 1 (duplicate of (D)) ---
    df["CM Other - 1"] = df["CM Other - 2 (D)"]
    df = move_after(df, "CM Other - 1", "CM Other - 2 (N)")

    # --- AGE_GP from AGE_YEAR ---
    age = pd.to_numeric(df["AGE_YEAR"], errors="coerce")
    df["AGE_GP"] = np.select(
        [age.lt(5), age.between(5, 14, inclusive="both"), age.ge(15)],
        ["<5", "5-14", ">=15"],
        default=""
    )
    df = move_after(df, "AGE_GP", "CM Other - 1")
    return df

# --- Sheet-name safety for Excel ---
def _sanitize_sheet_name(name: str) -> str:
    bad = '[]:*?/\\'
    table = {ord(ch): ' ' for ch in bad}
    name = str(name).translate(table).strip() or "Sheet"
    return name[:31]

# --- Chip-free multi-select like Cleaning App ---
def multiselect_dropdown(label: str, options: list[str], key: str = "ms_dropdown"):
    sel_key  = f"{key}_selected"
    snap_key = f"{key}_options_snapshot"

    if sel_key not in st.session_state:
        st.session_state[sel_key] = set()
    if snap_key not in st.session_state:
        st.session_state[snap_key] = tuple()

    # reset selection if options change (e.g., different file)
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

    # popover in newer Streamlit; fallback to expander for older versions
    if hasattr(st, "popover"):
        with st.popover(label):
            render_checkbox_list()
    else:
        with st.expander(label, expanded=False):
            render_checkbox_list()

    selected = [o for o in options if o in st.session_state[sel_key]]
    st.caption(f"Selected: {len(selected)}")
    return selected

# ==========================
# 🦟 Malaria Indicator Calculation App (Streamlit)
# Flow: upload Excel → select sheet(s) → compute indicators → download workbook
# ==========================
st.set_page_config(page_title="🦟 Malaria Indicator Calculation App", layout="wide")
st.title("🦟 Malaria Indicator Calculation App")

uploaded_file = st.file_uploader("Upload your malaria Excel file", type=["xlsx", "xls"])

# reset dropdown state when a different file is uploaded
if uploaded_file is not None:
    file_key = f"{uploaded_file.name}:{uploaded_file.size}"
    if st.session_state.get("uploaded_file_key") != file_key:
        st.session_state["uploaded_file_key"] = file_key
        for k in list(st.session_state.keys()):
            if k.endswith("_selected") or k.endswith("_options_snapshot"):
                del st.session_state[k]

if uploaded_file is None:
    st.info("Upload an Excel file to begin.")
    st.stop()

# Read workbook and list sheets
try:
    xls = pd.ExcelFile(uploaded_file)
    sheet_names = xls.sheet_names
except Exception as e:
    st.error("❌ Could not read the Excel file. Please check the format.")
    st.exception(e)
    st.stop()

# Let user pick one or more sheets (dropdown with checkboxes – no red chips)
selected_sheets = multiselect_dropdown(
    "Select sheet(s) to calculate indicators",
    sheet_names,
    key="sheets"
)
if not selected_sheets:
    st.warning("Select at least one sheet to continue.")
    st.stop()

# Process selected sheets with your existing logic
outputs = {}
errors = []
for sheet in selected_sheets:
    try:
        raw_df = xls.parse(sheet_name=sheet)
        out_df = compute_indicators(raw_df.copy())
        outputs[sheet] = out_df
    except Exception as e:
        errors.append((sheet, e))

# Preview results (first 50 rows)
if outputs:
    st.subheader("👀 Preview (first 50 rows)")
    tabs = st.tabs(list(outputs.keys()))
    for tab, sheet in zip(tabs, outputs.keys()):
        with tab:
            st.caption(f"Sheet: {sheet}")
            st.dataframe(outputs[sheet].head(50), use_container_width=True)

# Show any sheet-level errors
if errors:
    st.warning("Some sheets failed to process:")
    for sheet, e in errors:
        with st.expander(f"Details: {sheet}"):
            st.exception(e)

# Download: single workbook containing all selected sheets
if outputs:
    st.subheader("⬇️ Download")
    try:
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
            for sheet, df_out in outputs.items():
                safe_name = _sanitize_sheet_name(sheet)
                df_out.to_excel(writer, index=False, sheet_name=safe_name)
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

