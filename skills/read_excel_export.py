# -*- coding: utf-8 -*-
"""
Skill 3: read_excel_export.py
PBI AutoGov — Power BI Data Governance Automation Pipeline

Reads CSV or Excel files exported from Power BI visuals (right-click -> Export data)
and produces the standard 8-column metadata Excel that feeds into dax_query_builder.py.

How it works:
  1. Parse filename -> visual name
  2. Read column headers -> field names
  3. Classify each field by analyzing data values (first 20 rows):
     - Values with $, %, or pure numeric -> measure -> Usage = "Visual Value"
     - Text/categorical values -> grouping column -> Usage = "Visual Column"
  4. If --model-root provided, match fields against semantic model for exact
     table names + measure formulas
  5. Without model -> placeholder '<Table>'[FieldName]
  6. Output: standard 8-column metadata Excel

Input:  One or more CSV/Excel files exported from PBI visuals
        Optional: semantic model root for field matching
Output: pbi_report_metadata.xlsx (same format as Skill 1)

Usage:
    python skills/read_excel_export.py <files...> [--model-root PATH] [--output PATH] [--page-name TEXT]
"""

import argparse
import re
from pathlib import Path

import pandas as pd

from tmdl_parser import parse_semantic_model, match_field_to_model


# ============================================================
# Field classification by value analysis
# ============================================================

def classify_field_by_values(series: pd.Series) -> str:
    """Classify a column as 'measure' or 'grouping' by analyzing its values.

    Heuristic (applied to first 20 non-null values):
      - If most values match currency ($), percentage (%), or pure numeric -> measure
      - Otherwise -> grouping column

    Returns: 'measure' or 'grouping'
    """
    sample = series.dropna().head(20).astype(str)
    if len(sample) == 0:
        return "grouping"

    numeric_count = 0
    for val in sample:
        val = val.strip()
        # Currency: $1,234.56 or ($1,234.56) or -$1,234.56
        if re.match(r"^[\-\(]?\$[\d,]+\.?\d*\)?$", val):
            numeric_count += 1
            continue
        # Percentage: 12.34% or -12.34%
        if re.match(r"^[\-]?\d+\.?\d*%$", val):
            numeric_count += 1
            continue
        # Pure numeric (int or float, with optional commas)
        if re.match(r"^[\-]?[\d,]+\.?\d*$", val):
            numeric_count += 1
            continue

    # If >50% of values are numeric-like, classify as measure
    if numeric_count / len(sample) > 0.5:
        return "measure"
    return "grouping"


# ============================================================
# Visual name from filename
# ============================================================

def visual_name_from_filename(filepath: Path) -> str:
    """Derive a visual name from the export filename.

    PBI exports typically name files after the visual title:
      "Total Sales Variance by FiscalMonth and District Manager.csv"
    """
    return filepath.stem


# ============================================================
# Core: process a single exported file
# ============================================================

def process_export_file(filepath: Path, model=None, page_name: str = "Page 1",
                        visual_id: str = "") -> list[dict]:
    """Read a CSV/Excel export and return metadata rows.

    Args:
        filepath: Path to CSV or Excel file
        model: Optional SemanticModel from tmdl_parser for field matching
        page_name: Page name to use in output
        visual_id: Unique visual identifier for this file

    Returns:
        List of dicts with the 9 standard metadata columns (including Visual ID).
    """
    # Read file
    suffix = filepath.suffix.lower()
    if suffix == ".csv":
        df = pd.read_csv(filepath, encoding="utf-8-sig")
    elif suffix in (".xlsx", ".xls"):
        df = pd.read_excel(filepath)
    else:
        print(f"WARNING: Unsupported file type: {suffix} — skipping {filepath.name}")
        return []

    visual_name = visual_name_from_filename(filepath)
    print(f"\n  Processing: {filepath.name}")
    print(f"    Visual name: {visual_name}")
    print(f"    Columns: {list(df.columns)}")
    print(f"    Rows: {len(df)}")

    rows = []
    for col_name in df.columns:
        role = classify_field_by_values(df[col_name])
        usage = "Visual Value" if role == "measure" else "Visual Column"

        # Default: placeholder table/column
        table_sm = "<Table>"
        col_sm = col_name
        formula = ""

        # Try matching against semantic model
        if model is not None:
            match = match_field_to_model(col_name, model)
            if match:
                table_sm = match["table"]
                col_sm = match["field_name"]
                formula = match["formula"]
                # Override role based on match type
                if match["match_type"] in ("measure", "measure_fuzzy"):
                    usage = "Visual Value"
                print(f"    Matched: {col_name!r} -> '{table_sm}'[{col_sm}] ({match['match_type']})")
            else:
                print(f"    No match: {col_name!r} -> '<Table>'[{col_name}]")

        rows.append({
            "Page Name": page_name,
            "Visual/Table Name in PBI": visual_name,
            "Visual ID": visual_id,
            "Visual Type": "unknown",
            "UI Field Name": col_name,
            "Usage (Visual/Filter/Slicer)": usage,
            "Measure Formula": formula,
            "Table in the Semantic Model": table_sm,
            "Column in the Semantic Model": col_sm,
        })

    return rows


# ============================================================
# Export to Excel (same format as Skill 1)
# ============================================================

def export_to_excel(df: pd.DataFrame, output_path: str):
    """Save metadata DataFrame to Excel with auto-sized columns."""
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Report Metadata", index=False)
        ws = writer.sheets["Report Metadata"]
        for col_idx, col_name in enumerate(df.columns, 1):
            max_len = max(
                len(str(col_name)),
                df[col_name].astype(str).str.len().max() if len(df) > 0 else 0,
            )
            ws.column_dimensions[ws.cell(row=1, column=col_idx).column_letter].width = min(max_len + 2, 60)
    print(f"\nExcel file saved to: {output_path}")


# ============================================================
# Main entry point
# ============================================================

def read_excel_exports(files: list[str], model_root: str = None,
                       page_name: str = "Page 1") -> pd.DataFrame:
    """Process multiple CSV/Excel export files into metadata DataFrame.

    Args:
        files: List of file paths to process
        model_root: Optional path to semantic model definition root
        page_name: Page name to use in output

    Returns:
        DataFrame with all extracted metadata rows.
    """
    print("=" * 60)
    print("PBI AutoGov — CSV/Excel Export Reader")
    print("=" * 60)

    # Optionally load semantic model
    model = None
    if model_root:
        print(f"\nLoading semantic model: {model_root}")
        model = parse_semantic_model(model_root)
        print(f"  Measures: {len(model.measures)}")
        print(f"  Columns: {len(model.columns)}")

    all_rows = []
    for file_idx, fpath in enumerate(files, 1):
        filepath = Path(fpath)
        if not filepath.is_file():
            print(f"WARNING: File not found: {fpath}")
            continue
        visual_id = f"csv_{file_idx:03d}"
        rows = process_export_file(filepath, model=model, page_name=page_name,
                                   visual_id=visual_id)
        all_rows.extend(rows)

    df = pd.DataFrame(all_rows, columns=[
        "Page Name",
        "Visual/Table Name in PBI",
        "Visual ID",
        "Visual Type",
        "UI Field Name",
        "Usage (Visual/Filter/Slicer)",
        "Measure Formula",
        "Table in the Semantic Model",
        "Column in the Semantic Model",
    ])

    print(f"\n{'=' * 60}")
    print(f"Total rows extracted: {len(df)}")
    print(f"{'=' * 60}")

    return df


# ============================================================
# Standalone execution
# ============================================================

if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="PBI AutoGov — CSV/Excel Export Reader (Skill 3)"
    )
    parser.add_argument("files", nargs="+", help="CSV or Excel files exported from PBI visuals")
    parser.add_argument("--model-root", required=True,
                        help="Path to semantic model definition root (required for accurate field matching)")
    parser.add_argument("--output", default="pbi_report_metadata.xlsx",
                        help="Output Excel file path")
    parser.add_argument("--page-name", default="Page 1",
                        help="Page name to use in output (default: 'Page 1')")
    args = parser.parse_args()

    df = read_excel_exports(args.files, model_root=args.model_root, page_name=args.page_name)
    if not df.empty:
        export_to_excel(df, args.output)
    else:
        print("No data extracted. Check input files.")
