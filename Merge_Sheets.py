#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import argparse
import io
import os
import zipfile
from pathlib import Path
from typing import Dict, Iterable, List, Optional

import pandas as pd

# Optional import for preserving formatting when injecting into .xlsx
try:
    from openpyxl import load_workbook
    from openpyxl.utils.dataframe import dataframe_to_rows
    _HAVE_OPENPYXL = True
except Exception:  # pragma: no cover
    _HAVE_OPENPYXL = False


# ---------- helpers ----------
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


# ---------- core ----------
def merge_across_files(file_sheet_map: Dict[str, Iterable[str]]) -> pd.DataFrame:
    """
    Merge rows across multiple files/sheets. Adds DATA_SOURCE and FILE_SOURCE.
    APPEND-ONLY: do not touch values; keep strings as-is and keep real Excel dates as Timestamps.
    """
    parts: List[pd.DataFrame] = []
    for fname, sheets in file_sheet_map.items():
        xls = pd.ExcelFile(fname)
        for s in sheets:
            df = xls.parse(sheet_name=s)
            # strip headers to avoid 'SCREENING_DATE ' / invisible-space issues
            df.columns = [str(c).strip() for c in df.columns]

            # Do NOT coerce/format dates here; leave exactly as read.

            df["DATA_SOURCE"] = s
            df["FILE_SOURCE"] = Path(fname).name
            parts.append(df)

    if not parts:
        raise ValueError("No sheets selected.")

    _validate_headers_match(parts)

    # Keep first part's column order; ensure meta columns at the end
    canonical = list(parts[0].columns)
    for m in ("DATA_SOURCE", "FILE_SOURCE"):
        if m in canonical:
            canonical.remove(m)
    canonical += ["DATA_SOURCE", "FILE_SOURCE"]

    normalized = []
    for df in parts:
        mapping = {str(c).strip().lower(): c for c in df.columns}
        ordered = [mapping[str(c).strip().lower()] for c in canonical if str(c).strip().lower() in mapping]
        tmp = df[ordered].copy()
        normalized.append(tmp)

    merged = pd.concat(normalized, ignore_index=True)
    return merged


def write_merged_only(path: str, merged_df: pd.DataFrame, sheet_name: str = "Merged") -> None:
    """
    Write merged-only workbook. Keep values as-is, but make real date cells
    display as DD-MMM-YY (e.g., 01-Jan-26) instead of 2026-01-01 00:00:00.
    """
    with pd.ExcelWriter(path, engine="xlsxwriter") as writer:
        merged_df.to_excel(writer, index=False, sheet_name=sheet_name)

        # Apply date column format for display (text values are unaffected)
        ws = writer.sheets[sheet_name]
        try:
            date_col_idx = next(
                i for i, c in enumerate(merged_df.columns)
                if str(c).strip().lower() == "screening_date"
            )
            date_fmt = writer.book.add_format({"num_format": "DD-MMM-YY"})
            ws.set_column(date_col_idx, date_col_idx, 12, date_fmt)
        except StopIteration:
            pass


def write_zip_injected(file_sheet_map: Dict[str, Iterable[str]], merged_df: pd.DataFrame, zip_path: str) -> None:
    """
    Create a ZIP where each original .xlsx gets a merged sheet appended (as first tab),
    preserving original formatting. For .xls inputs, include a new .xlsx containing
    only the merged sheet (values as-is, with date display format applied).
    """
    if not _HAVE_OPENPYXL:
        raise RuntimeError(
            "openpyxl not available; cannot preserve formatting for .xlsx. Add 'openpyxl' to requirements."
        )

    with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for src_path in file_sheet_map.keys():
            fname = os.path.basename(src_path)

            if fname.lower().endswith(".xlsx"):
                with open(src_path, "rb") as f:
                    orig_bytes = f.read()

                wb = load_workbook(io.BytesIO(orig_bytes))

                # Unique sheet name and put the merged sheet FIRST
                base_name = "Merged"
                name = base_name
                counter = 1
                while name in wb.sheetnames:
                    counter += 1
                    name = f"{base_name}_{counter}"
                ws = wb.create_sheet(title=name, index=0)

                # Write header + rows
                from openpyxl.utils.dataframe import dataframe_to_rows
                for r in dataframe_to_rows(merged_df, index=False, header=True):
                    ws.append(r)

                # Apply display format to the real-date cells only
                try:
                    date_col_idx = next(
                        i for i, c in enumerate(merged_df.columns, start=1)
                        if str(c).strip().lower() == "screening_date"
                    )
                    for row_idx in range(2, len(merged_df) + 2):  # skip header
                        cell = ws.cell(row=row_idx, column=date_col_idx)
                        cell.number_format = "DD-MMM-YY"
                except StopIteration:
                    pass

                out = io.BytesIO()
                wb.save(out)
                out.seek(0)
                zf.writestr(fname, out.read())

            else:
                # .xls fallback: cannot preserve original formatting; deliver an xlsx with merged-only
                alt_name = os.path.splitext(fname)[0] + "_merged_only.xlsx"
                bout = io.BytesIO()
                with pd.ExcelWriter(bout, engine="xlsxwriter") as writer:
                    merged_df.to_excel(writer, index=False, sheet_name="Merged")
                    ws = writer.sheets["Merged"]
                    try:
                        idx = next(
                            i for i, c in enumerate(merged_df.columns)
                            if str(c).strip().lower() == "screening_date"
                        )
                        date_fmt = writer.book.add_format({"num_format": "DD-MMM-YY"})
                        ws.set_column(idx, idx, 12, date_fmt)
                    except StopIteration:
                        pass
                bout.seek(0)
                zf.writestr(alt_name, bout.read())


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
    ap.add_argument("-o", "--output", default="merged.xlsx", help="Path to merged-only workbook to write")
    ap.add_argument("--zip", dest="zip_path", help="Create a ZIP that injects the merged sheet back into each original file")
    ap.add_argument("--select", action="append", help="Select sheets per file: 'file.xlsx:Sheet1,Sheet2' (repeatable)")
    args = ap.parse_args(argv)

    file_sheet_map = parse_select_args(args.select or [], args.files)
    merged_df = merge_across_files(file_sheet_map)

    write_merged_only(args.output, merged_df)
    print(f"✅ Wrote merged-only workbook: {args.output}")

    if args.zip_path:
        write_zip_injected(file_sheet_map, merged_df, args.zip_path)
        print(f"📦 Wrote ZIP bundle with injected merged sheets: {args.zip_path}")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
