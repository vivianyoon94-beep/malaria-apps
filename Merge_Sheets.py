
import argparse
import io
from typing import Iterable, List, Union, Optional
import pandas as pd
import sys
import os

# Optional import: only used when --include-originals is set
try:
    from openpyxl import load_workbook
    from openpyxl.utils.dataframe import dataframe_to_rows
    _HAVE_OPENPYXL = True
except Exception:  # pragma: no cover
    _HAVE_OPENPYXL = False

def _excel_file_type(path: str) -> str:
    ext = os.path.splitext(path)[1].lower()
    if ext in {".xlsx"}:
        return "xlsx"
    if ext in {".xls"}:
        return "xls"
    raise ValueError("Unsupported file type. Please provide .xlsx or .xls")

def _read_excel_file(src: Union[str, bytes, io.BytesIO]) -> pd.ExcelFile:
    """Return a pandas.ExcelFile, choosing engine by extension when possible."""
    if isinstance(src, (bytes, io.BytesIO)):
        # Pandas can infer .xlsx engine from bytes; .xls needs xlrd<2.0
        try:
            return pd.ExcelFile(src)
        except Exception as e:
            raise RuntimeError("Failed to open Excel bytes. For .xls, ensure xlrd<2.0 is installed.") from e
    else:
        kind = _excel_file_type(str(src))
        if kind == "xlsx":
            return pd.ExcelFile(src, engine="openpyxl")
        else:  # xls
            # xlrd 2.0+ dropped .xls support; user must have xlrd==1.2.0 installed
            try:
                return pd.ExcelFile(src, engine="xlrd")
            except Exception as e:
                raise RuntimeError("Opening .xls requires xlrd<2.0 (e.g., xlrd==1.2.0).") from e

def _normalize_columns(df: pd.DataFrame) -> pd.Index:
    """Return a normalized column index for comparison (strip + lower)."""
    return pd.Index([str(c).strip().lower() for c in df.columns])

def _validate_same_headers(dfs: List[pd.DataFrame]) -> List[str]:
    """Ensure all DataFrames have the same headers (order-insensitive). Returns canonical list order
    based on the first DF's columns."""
    if not dfs:
        raise ValueError("No DataFrames provided to validate.")
    base_cols = list(dfs[0].columns)
    base_norm = set(_normalize_columns(dfs[0]))
    for idx, df in enumerate(dfs[1:], start=2):
        if set(_normalize_columns(df)) != base_norm:
            raise ValueError(f"Selected sheets do not share the same headers. First set: {list(dfs[0].columns)}; "
                             f"mismatch at sheet #{idx} with columns {list(df.columns)}")
    return base_cols

def merge_sheets_from_excel(src: Union[str, bytes, io.BytesIO], sheets: Iterable[str]) -> pd.DataFrame:
    """Merge rows from selected sheets in the same Excel file (headers must match).
    Adds DATA_SOURCE column with the sheet name.
    """
    xls = _read_excel_file(src)
    sheets = list(sheets)
    missing = [s for s in sheets if s not in xls.sheet_names]
    if missing:
        raise ValueError(f"Sheet(s) not found: {missing}. Available: {xls.sheet_names}")
    # Read selected sheets
    dfs = []
    for s in sheets:
        df = xls.parse(sheet_name=s)
        df.columns = [str(c) for c in df.columns]  # keep original names as strings
        dfs.append((s, df))
    # Validate same headers (order-insensitive); reorder to match first sheet
    canonical = _validate_same_headers([d for _, d in dfs])
    for i in range(len(dfs)):
        s, df = dfs[i]
        # Reorder columns to canonical order (case-insensitive match)
        mapping = {str(c).strip().lower(): c for c in df.columns}
        ordered = [mapping[str(c).strip().lower()] for c in canonical]
        df = df[ordered].copy()
        df["DATA_SOURCE"] = s
        dfs[i] = (s, df)
    merged = pd.concat([d for _, d in dfs], ignore_index=True)
    return merged

def _write_output_only_merged(dest: str, merged_df: pd.DataFrame, merged_sheet_name: str = "Merged") -> None:
    kind = _excel_file_type(dest)
    if kind == "xlsx":
        with pd.ExcelWriter(dest, engine="xlsxwriter") as writer:
            merged_df.to_excel(writer, index=False, sheet_name=merged_sheet_name)
    else:  # xls
        # xls writer uses xlwt (no formatting anyway). Pandas handles engine choice.
        with pd.ExcelWriter(dest) as writer:
            merged_df.to_excel(writer, index=False, sheet_name=merged_sheet_name)

def _write_output_with_originals(src_path: str, dest: str, merged_df: pd.DataFrame, merged_sheet_name: str = "Merged") -> None:
    # Only supports .xlsx for preserving formats using openpyxl
    if _excel_file_type(src_path) != "xlsx":
        raise ValueError("--include-originals requires an .xlsx input (openpyxl).")
    if not _HAVE_OPENPYXL:
        raise RuntimeError("openpyxl not available. Add 'openpyxl' to requirements.")

    import io
    with open(src_path, "rb") as f:
        original_bytes = f.read()
    from openpyxl import load_workbook
    from openpyxl.utils.dataframe import dataframe_to_rows

    wb = load_workbook(io.BytesIO(original_bytes))
    # Place merged sheet at the end; ensure unique name
    name = merged_sheet_name
    counter = 1
    while name in wb.sheetnames:
        counter += 1
        name = f"{merged_sheet_name}_{counter}"
    ws = wb.create_sheet(title=name)
    for r in dataframe_to_rows(merged_df, index=False, header=True):
        ws.append(r)
    wb.save(dest)

def main(argv: Optional[List[str]] = None) -> int:
    ap = argparse.ArgumentParser(description="Merge selected sheets from an Excel file.")
    ap.add_argument("input", help="Path to input .xlsx or .xls file")
    ap.add_argument("--sheets", required=True, help="Comma-separated list of sheet names to merge")
    ap.add_argument("--out", default="merged.xlsx", help="Output Excel path (default: merged.xlsx)")
    ap.add_argument("--merged-sheet-name", default="Merged", help="Name of the merged sheet")
    ap.add_argument("--include-originals", action="store_true",
                    help="(xlsx only) Keep original sheets and append the merged sheet (preserve formatting)")
    args = ap.parse_args(argv)

    # Prepare selection
    selected = [s.strip() for s in args.sheets.split(",") if s.strip()]
    if not selected:
        print("No sheets provided. Use --sheets \"Sheet1,Sheet2\"", file=sys.stderr)
        return 2

    # Merge
    merged_df = merge_sheets_from_excel(args.input, selected)

    # Write
    if args.include_originals:
        _write_output_with_originals(args.input, args.out, merged_df, args.merged_sheet_name)
    else:
        _write_output_only_merged(args.out, merged_df, args.merged_sheet_name)

    print(f"✅ Wrote {args.out} with sheet: {args.merged_sheet_name}")
    return 0

if __name__ == "__main__":
    raise SystemExit(main())
