import argparse
import io
import os
from pathlib import Path
from typing import Dict, Iterable, List, Optional

import pandas as pd

# Optional import for building/formatting a single combined workbook
try:
    from openpyxl import Workbook
    from openpyxl import load_workbook
    from openpyxl.utils.dataframe import dataframe_to_rows
    _HAVE_OPENPYXL = True
except Exception:  # pragma: no cover
    _HAVE_OPENPYXL = False


# ---------- helpers ----------
def _strip_time_from_datetime_columns(df: pd.DataFrame) -> pd.DataFrame:
    for col, dtype in df.dtypes.items():
        if str(dtype).startswith("datetime64"):
            df[col] = pd.to_datetime(df[col], errors="coerce").dt.date
    return df

def _normalize_headers(df: pd.DataFrame) -> List[str]:
    return [str(c).strip().lower() for c in df.columns]

def _validate_headers_match(dfs: List[pd.DataFrame]) -> None:
    if not dfs:
        raise ValueError("No dataframes selected")
    base = set(_normalize_headers(dfs[0]))
    for i, d in enumerate(dfs[1:], start=2):
        if set(_normalize_headers(d)) != base:
            raise ValueError(
                f"Headers do not match across all selected sheets (mismatch near input #{i})."
            )

def _sanitize_sheet_name(name: str) -> str:
    # Excel sheet name: <=31 chars, no []:*?/\
    bad = '[]:*?/\\'
    trans = {ord(ch): ' ' for ch in bad}
    name = (str(name) or "Sheet").translate(trans).strip()
    return name[:31] if name else "Sheet"


# ---------- core ----------
def merge_across_files(file_sheet_map: Dict[str, Iterable[str]]) -> pd.DataFrame:
    """Merge rows across multiple files/sheets. Adds DATA_SOURCE and FILE_SOURCE."""
    parts: List[pd.DataFrame] = []
    for fname, sheets in file_sheet_map.items():
        xls = pd.ExcelFile(fname)
        for s in sheets:
            df = xls.parse(sheet_name=s)
            df.columns = [str(c) for c in df.columns]
            df["DATA_SOURCE"] = s
            df["FILE_SOURCE"] = Path(fname).name
            parts.append(df)

    if not parts:
        raise ValueError("No sheets selected.")

    _validate_headers_match(parts)

    meta_cols = {"DATA_SOURCE", "FILE_SOURCE"}
    canonical = [c for c in parts[0].columns if c not in meta_cols]

    normalized = []
    for df in parts:
        mapping = {str(c).strip().lower(): c for c in df.columns}
        ordered = [mapping[str(c).strip().lower()] for c in canonical]
        tmp = df[ordered + ["DATA_SOURCE", "FILE_SOURCE"]].copy()
        normalized.append(tmp)

    merged = pd.concat(normalized, ignore_index=True)
    merged = _strip_time_from_datetime_columns(merged)
    return merged


def write_merged_only(path: str, merged_df: pd.DataFrame, sheet_name: str = "Merged") -> None:
    # Use XlsxWriter so Excel displays dates as DD-MMM-YY
    with pd.ExcelWriter(
        path,
        engine="xlsxwriter",
        datetime_format="dd-mmm-yy",
        date_format="dd-mmm-yy",
    ) as writer:
        merged_df.to_excel(writer, index=False, sheet_name=sheet_name)


def write_combined_workbook(
    file_sheet_map: Dict[str, Iterable[str]],
    merged_df: pd.DataFrame,
    out_path: str,
) -> None:
    """
    Create a single Excel workbook containing:
      - 'Merged' (first sheet, date-formatted as DD-MMM-YY)
      - all selected original sheets appended after, sheet names prefixed with '<file>-<sheet>'
    Note: original styles/filters can’t be preserved when consolidating into a new workbook.
    """
    if not _HAVE_OPENPYXL:
        raise RuntimeError("openpyxl not available; add 'openpyxl' to requirements.")

    from datetime import date as _dt_date

    wb = Workbook()
    # Remove the default sheet
    default = wb.active
    wb.remove(default)

    # 1) Add merged as the first sheet
    ws_m = wb.create_sheet(title="Merged", index=0)
    for r in dataframe_to_rows(merged_df, index=False, header=True):
        ws_m.append(r)

    # format date columns in merged
    date_cols_idx = []
    for j, col in enumerate(merged_df.columns, start=1):
        col_s = merged_df[col]
        if pd.api.types.is_datetime64_any_dtype(col_s) or col_s.map(lambda v: isinstance(v, (_dt_date, pd.Timestamp))).any():
            date_cols_idx.append(j)
    for j in date_cols_idx:
        for col_cells in ws_m.iter_cols(min_col=j, max_col=j, min_row=2):
            for cell in col_cells:
                cell.number_format = "DD-MMM-YY"

    # 2) Append all selected original sheets (values only)
    for fpath, sheets in file_sheet_map.items():
        file_base = Path(fpath).stem
        xls = pd.ExcelFile(fpath)
        for s in sheets:
            df = xls.parse(sheet_name=s)
            # Keep column order and values
            title = _sanitize_sheet_name(f"{file_base} - {s}")
            # Ensure uniqueness if collision
            orig_title = title
            k = 2
            while title in wb.sheetnames:
                title = _sanitize_sheet_name(f"{orig_title[:28]}_{k}")
                k += 1
            ws = wb.create_sheet(title=title)
            for r in dataframe_to_rows(df, index=False, header=True):
                ws.append(r)

    wb.save(out_path)


# ---------- CLI ----------
def parse_select_args(select_args: List[str], files: List[str]) -> Dict[str, List[str]]:
    """Parse selections of the form:
       --select "file1.xlsx:SheetA,SheetB"  --select "file2.xlsx:Data"
       If no --select provided, default to all sheets for each file.
    """
    mapping: Dict[str, List[str]] = {}
    if not select_args:
        for f in files:
            mapping[f] = pd.ExcelFile(f).sheet_names
        return mapping

    wanted: Dict[str, List[str]] = {}
    for item in select_args:
        item = item.strip()
        if not item or ":" not in item:
            raise ValueError(f"Bad --select value: {item!r}. Use 'file.xlsx:Sheet1,Sheet2'")
        fname, sheets_str = item.split(":", 1)
        fname = fname.strip()
        if fname not in files:
            raise ValueError(f"--select references unknown file: {fname}")
        sheet_list = [s.strip() for s in sheets_str.split(",") if s.strip()]
        if not sheet_list:
            raise ValueError(f"No sheets provided in --select for {fname}")
        wanted[fname] = sheet_list

    for f in files:
        mapping[f] = wanted.get(f, pd.ExcelFile(f).sheet_names)
    return mapping


def main(argv: Optional[List[str]] = None) -> int:
    ap = argparse.ArgumentParser(description="Merge sheets across one or more Excel files.")
    ap.add_argument("files", nargs="+", help="Paths to input Excel files (.xlsx or .xls)")
    ap.add_argument("-o", "--output", default="merged.xlsx", help="Path to merged-only workbook")
    ap.add_argument("--combined", help="Write one Excel that contains 'Merged' first + all selected original sheets")
    ap.add_argument("--select", action="append", help="Select sheets per file: 'file.xlsx:Sheet1,Sheet2' (repeatable)")
    args = ap.parse_args(argv)

    file_sheet_map = parse_select_args(args.select or [], args.files)
    merged_df = merge_across_files(file_sheet_map)

    # 1) merged-only workbook
    write_merged_only(args.output, merged_df)
    print(f"✅ Wrote merged-only workbook: {args.output}")

    # 2) single combined workbook (optional)
    if args.combined:
        write_combined_workbook(file_sheet_map, merged_df, args.combined)
        print(f"📘 Wrote combined workbook (Merged + originals): {args.combined}")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
