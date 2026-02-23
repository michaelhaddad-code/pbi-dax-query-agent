# -*- coding: utf-8 -*-
"""
Skill 4: analyze_screenshot.py
PBI AutoGov -- Power BI Visual Metadata Builder

Builds the standard 8-column metadata Excel from a manual visual description.
The Claude Code agent reads the screenshot, identifies the visual type, name,
and fields, then passes them as CLI arguments to this script for semantic model
matching and Excel output.

How it works:
  1. Agent (Claude Code) views the screenshot and identifies visual type, name, fields
  2. Agent calls this script with --visual-type, --visual-name, --field args
  3. Script matches fields against the semantic model (if --model-root provided)
  4. Output: standard 8-column metadata Excel

Input:  Visual description via CLI args (from agent's screenshot analysis)
        Optional: semantic model root for field matching
Output: pbi_report_metadata.xlsx (same format as Skill 1)

Usage:
    python skills/analyze_screenshot.py \
      --visual-type pieChart \
      --visual-name "This Year Sales by Chain" \
      --field "Chain:grouping" \
      --field "This Year Sales:measure" \
      --model-root "data/Store Sales.SemanticModel/definition" \
      --output "output/screenshot_metadata.xlsx"
"""

import argparse
from pathlib import Path

import pandas as pd

from tmdl_parser import parse_semantic_model, match_field_to_model


# ============================================================
# Visual type normalization (human description -> PBI identifier)
# ============================================================

VISUAL_TYPE_NORMALIZE = {
    "bar chart": "barChart",
    "clustered bar chart": "clusteredBarChart",
    "clustered column chart": "clusteredColumnChart",
    "stacked bar chart": "stackedBarChart",
    "stacked column chart": "stackedColumnChart",
    "column chart": "clusteredColumnChart",
    "line chart": "lineChart",
    "area chart": "areaChart",
    "line and column chart": "lineClusteredColumnComboChart",
    "combo chart": "lineClusteredColumnComboChart",
    "ribbon chart": "ribbonChart",
    "waterfall chart": "waterfallChart",
    "funnel chart": "funnelChart",
    "pie chart": "pieChart",
    "donut chart": "donutChart",
    "treemap": "treemap",
    "map": "map",
    "filled map": "filledMap",
    "table": "tableEx",
    "matrix": "pivotTable",
    "card": "card",
    "multi-row card": "multiRowCard",
    "kpi": "kpi",
    "gauge": "gauge",
    "slicer": "slicer",
    "scatter chart": "scatterChart",
    "scatter plot": "scatterChart",
}


def normalize_visual_type(vis_type: str) -> str:
    """Normalize a visual type description to a PBI identifier.

    Accepts both PBI identifiers (pieChart) and human names (pie chart).
    """
    key = vis_type.strip().lower()
    return VISUAL_TYPE_NORMALIZE.get(key, vis_type)


# ============================================================
# Core: build metadata rows from visual description
# ============================================================

def build_metadata_from_description(visual_type: str, visual_name: str,
                                    fields: list, model=None,
                                    page_name: str = "Page 1",
                                    visual_id: str = "") -> list[dict]:
    """Build metadata rows from a visual description.

    Args:
        visual_type: PBI visual type identifier or human-readable name
        visual_name: Visual title/name
        fields: List of (field_name, role) tuples where role is "measure" or "grouping"
        model: Optional SemanticModel from tmdl_parser for field matching
        page_name: Page name to use in output
        visual_id: Unique visual identifier

    Returns:
        List of dicts with the 9 standard metadata columns (including Visual ID).
    """
    vis_type = normalize_visual_type(visual_type)

    print(f"  Visual type: {vis_type}")
    print(f"  Visual name: {visual_name}")
    print(f"  Fields: {len(fields)}")

    rows = []
    for fname, role in fields:
        if role == "measure":
            usage = "Visual Value"
        elif vis_type in ("slicer", "advancedSlicerVisual"):
            usage = "Slicer"
        else:
            usage = "Visual Column"

        # Default: placeholder table/column
        table_sm = "<Table>"
        col_sm = fname
        formula = ""

        # Try matching against semantic model
        if model is not None:
            match = match_field_to_model(fname, model)
            if match:
                table_sm = match["table"]
                col_sm = match["field_name"]
                formula = match["formula"]
                if match["match_type"] in ("measure", "measure_fuzzy"):
                    usage = "Visual Value"
                print(f"    Matched: {fname!r} -> '{table_sm}'[{col_sm}] ({match['match_type']})")
            else:
                print(f"    No match: {fname!r} -> '<Table>'[{fname}]")

        rows.append({
            "Page Name": page_name,
            "Visual/Table Name in PBI": visual_name,
            "Visual ID": visual_id,
            "Visual Type": vis_type,
            "UI Field Name": fname,
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
# Standalone execution
# ============================================================

if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="PBI AutoGov -- Visual Metadata Builder (Skill 4)"
    )
    parser.add_argument("--visual-type", required=True,
                        help="PBI visual type (e.g., pieChart, barChart, card, slicer)")
    parser.add_argument("--visual-name", required=True,
                        help="Visual title/name as shown in the report")
    parser.add_argument("--field", action="append", required=True, dest="fields",
                        help="Field in 'name:role' format where role is 'measure' or 'grouping'. "
                             "Can be repeated. E.g.: --field 'Chain:grouping' --field 'Sales:measure'")
    parser.add_argument("--model-root", required=True,
                        help="Path to semantic model definition root (required for accurate field matching)")
    parser.add_argument("--output", default="pbi_report_metadata.xlsx",
                        help="Output Excel file path")
    parser.add_argument("--page-name", default="Page 1",
                        help="Page name to use in output (default: 'Page 1')")
    args = parser.parse_args()

    print("=" * 60)
    print("PBI AutoGov -- Visual Metadata Builder")
    print("=" * 60)

    # Parse field args ("name:role" format)
    parsed_fields = []
    for f in args.fields:
        if ":" in f:
            name, role = f.rsplit(":", 1)
            role = role.strip().lower()
            if role not in ("measure", "grouping"):
                print(f"WARNING: Unknown role '{role}' for field '{name}', defaulting to 'grouping'")
                role = "grouping"
            parsed_fields.append((name.strip(), role))
        else:
            print(f"WARNING: Field '{f}' missing ':role' suffix, defaulting to 'grouping'")
            parsed_fields.append((f.strip(), "grouping"))

    # Optionally load semantic model
    model = None
    if args.model_root:
        print(f"\nLoading semantic model: {args.model_root}")
        model = parse_semantic_model(args.model_root)
        print(f"  Measures: {len(model.measures)}")
        print(f"  Columns: {len(model.columns)}")

    rows = build_metadata_from_description(
        args.visual_type, args.visual_name, parsed_fields,
        model=model, page_name=args.page_name,
        visual_id="screenshot_001",
    )

    df = pd.DataFrame(rows, columns=[
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

    if not df.empty:
        export_to_excel(df, args.output)
    else:
        print("No data extracted. Check --field arguments.")
