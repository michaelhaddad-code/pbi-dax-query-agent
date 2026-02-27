# -*- coding: utf-8 -*-
"""
Unified CLI for the PBI DAX Query Generation Pipeline.

Chains all skills in sequence: .pbix extraction (optional) → metadata extraction → DAX query generation.

Usage:
    python pbi_pipeline.py "report.pbix"
    python pbi_pipeline.py --report-root "data/X.Report/definition" --model-root "data/X.SemanticModel/definition"
    python pbi_pipeline.py "Revenue Opportunities"
"""

import argparse
import os
import re
import sys

# Windows console encoding fix
os.environ.setdefault("PYTHONIOENCODING", "utf-8")

# Add skills/ to path for imports
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "skills"))

from extract_metadata import extract_metadata, export_to_excel
from dax_query_builder import read_extractor_output, build_bookmark_queries, write_output
from tmdl_parser import parse_semantic_model


def sanitize_filename(name: str) -> str:
    """Sanitize a report name for use as a filename."""
    return re.sub(r'[^\w\-]', '_', name)


def resolve_sample_report(name: str, data_dir: str) -> tuple:
    """Resolve a sample report name to report_root and model_root paths.

    Tries exact match first, then case-insensitive prefix match.

    Returns:
        (report_root, model_root) paths or raises FileNotFoundError
    """
    # Map of known shortcut names → actual directory names
    shortcuts = {
        "revenue opportunities": "Revenue Opportunities",
        "store sales": "Store Sales",
        "ai sample": "Artificial Intelligence Sample (2)",
        "artificial intelligence sample": "Artificial Intelligence Sample (2)",
        "regional sales": "Regional Sales Sample",
        "regional sales sample": "Regional Sales Sample",
        "nations": "nationsSample",
        "nationssample": "nationsSample",
        "data dump": "Data Dump 05152025 - Use As Sample",
    }

    # Try shortcut lookup first
    resolved = shortcuts.get(name.lower())
    if resolved:
        report_root = os.path.join(data_dir, f"{resolved}.Report", "definition")
        model_root = os.path.join(data_dir, f"{resolved}.SemanticModel", "definition")
        if os.path.isdir(report_root):
            return report_root, model_root, resolved

    # Try exact name
    report_root = os.path.join(data_dir, f"{name}.Report", "definition")
    model_root = os.path.join(data_dir, f"{name}.SemanticModel", "definition")
    if os.path.isdir(report_root):
        return report_root, model_root, name

    # Try case-insensitive scan of data_dir
    name_lower = name.lower()
    for entry in os.listdir(data_dir):
        if entry.lower().startswith(name_lower) and entry.endswith(".Report"):
            base = entry[:-len(".Report")]
            report_root = os.path.join(data_dir, entry, "definition")
            model_root = os.path.join(data_dir, f"{base}.SemanticModel", "definition")
            if os.path.isdir(report_root):
                return report_root, model_root, base

    raise FileNotFoundError(
        f"Could not find sample report '{name}' in {data_dir}. "
        f"Available reports: {[d for d in os.listdir(data_dir) if d.endswith('.Report')]}"
    )


def main():
    parser = argparse.ArgumentParser(
        description="PBI DAX Query Generation Pipeline — unified CLI",
        epilog="Examples:\n"
               '  python pbi_pipeline.py "Revenue Opportunities"\n'
               '  python pbi_pipeline.py "report.pbix"\n'
               '  python pbi_pipeline.py --report-root "data/X.Report/definition" '
               '--model-root "data/X.SemanticModel/definition"',
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument(
        "input", nargs="?", default=None,
        help="Path to .pbix file OR sample report name (e.g., 'Revenue Opportunities')",
    )
    parser.add_argument("--report-root", help="Path to PBIP report definition root")
    parser.add_argument("--model-root", help="Path to PBIP semantic model definition root")
    parser.add_argument("--output-dir", default="output", help="Output directory (default: output/)")
    parser.add_argument("--no-bookmarks", action="store_true", help="Skip bookmark extraction")

    args = parser.parse_args()

    # Validate inputs
    if not args.input and not (args.report_root and args.model_root):
        parser.error("Provide either a .pbix file / sample name, or both --report-root and --model-root")

    project_root = os.path.dirname(os.path.abspath(__file__))
    data_dir = os.path.join(project_root, "data")
    output_dir = os.path.join(project_root, args.output_dir)
    os.makedirs(output_dir, exist_ok=True)

    report_root = None
    model_root = None
    report_name = None
    semantic_model_source = ""  # Track model provenance through the pipeline

    # --- Mode detection ---
    if args.report_root and args.model_root:
        # Explicit PBIP paths — source auto-detected from .source marker by parse_semantic_model()
        report_root = args.report_root
        model_root = args.model_root
        # Derive name from report root path
        report_name = os.path.basename(os.path.dirname(os.path.dirname(report_root)))
        if report_name.endswith(".Report"):
            report_name = report_name[:-len(".Report")]
        print(f"Mode: Explicit PBIP paths")

    elif args.input and args.input.lower().endswith(".pbix"):
        # .pbix file — run Skill 0 first
        pbix_path = args.input
        if not os.path.isfile(pbix_path):
            print(f"ERROR: PBIX file not found: {pbix_path}")
            sys.exit(1)

        print("=" * 60)
        print("[0] Extracting .pbix file...")
        print("=" * 60)

        from pbix_extractor import extract_pbix
        result = extract_pbix(pbix_path, output_dir=data_dir, model_root=args.model_root)
        report_root = result.report_root
        model_root = result.model_root
        report_name = result.report_name
        semantic_model_source = result.semantic_model_source

        print(f"\n    Extracted: {report_name}")
        print(f"    Pages: {result.page_count}, Data visuals: {result.data_visual_count}, "
              f"Bookmarks: {result.bookmark_count}")
        print(f"    Semantic model: {result.semantic_model_source}")

        if not model_root:
            print("\n    WARNING: No semantic model available. Measure formulas will be missing.")
            print("    Install pbixray (`pip install pbixray`) or provide --model-root.")

    else:
        # Sample report name
        name = args.input
        try:
            report_root, model_root, report_name = resolve_sample_report(name, data_dir)
        except FileNotFoundError as e:
            print(f"ERROR: {e}")
            sys.exit(1)
        print(f"Mode: Sample report — {report_name}")

    # Check paths exist
    if not os.path.isdir(report_root):
        print(f"ERROR: Report root not found: {report_root}")
        sys.exit(1)
    if model_root and not os.path.isdir(model_root):
        print(f"WARNING: Model root not found: {model_root} — proceeding without semantic model")
        model_root = report_root  # fallback (extract_metadata handles missing tables/)

    safe_name = sanitize_filename(report_name)
    metadata_path = os.path.join(output_dir, f"{safe_name}_metadata.xlsx")
    dax_path = os.path.join(output_dir, f"{safe_name}_dax_queries.xlsx")

    # --- Step 1: Extract metadata ---
    print("\n" + "=" * 60)
    print("[1] Extracting metadata...")
    print("=" * 60)

    include_bookmarks = not args.no_bookmarks
    df, bookmarks_list, filter_expressions = extract_metadata(
        report_root, model_root, include_bookmarks=include_bookmarks,
        semantic_model_source=semantic_model_source,
    )
    export_to_excel(df, metadata_path, bookmarks_list=bookmarks_list,
                    filter_expressions=filter_expressions)

    metadata_rows = len(df)
    bookmark_count = len(bookmarks_list) if bookmarks_list else 0
    print(f"\n    Metadata: {metadata_rows} rows → {metadata_path}")
    if bookmark_count:
        print(f"    Bookmarks: {bookmark_count}")

    # --- Step 2: Generate DAX queries ---
    print("\n" + "=" * 60)
    print("[2] Generating DAX queries...")
    print("=" * 60)

    visuals, page_filters, bookmarks, filter_expr_data = read_extractor_output(metadata_path)

    # Load semantic model for filter redundancy checks
    model = None
    if model_root:
        try:
            model = parse_semantic_model(model_root)
        except Exception:
            pass  # Non-critical — queries still generated without formula checks

    # Build bookmark queries if bookmarks present
    bookmark_queries = None
    if bookmarks:
        bookmark_queries = build_bookmark_queries(bookmarks, visuals, page_filters, model=model)

    visual_count = write_output(visuals, page_filters, dax_path,
                                bookmark_queries=bookmark_queries,
                                filter_expr_data=filter_expr_data,
                                model=model)

    bm_query_count = len(bookmark_queries) if bookmark_queries else 0
    print(f"\n    DAX queries: {visual_count} visuals → {dax_path}")
    if bm_query_count:
        print(f"    Bookmark DAX queries: {bm_query_count}")

    # --- Summary ---
    print("\n" + "=" * 60)
    print("Pipeline complete!")
    print("=" * 60)
    print(f"  Report:     {report_name}")
    print(f"  Visuals:    {visual_count}")
    print(f"  Metadata:   {metadata_path}")
    print(f"  DAX:        {dax_path}")
    if bm_query_count:
        print(f"  Bookmarks:  {bm_query_count} bookmark×visual queries")
    print()


if __name__ == "__main__":
    main()
