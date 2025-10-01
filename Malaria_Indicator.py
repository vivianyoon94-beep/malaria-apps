import pandas as pd
import numpy as np



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

