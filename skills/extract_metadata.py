# -*- coding: utf-8 -*-
"""
Skill 1: extract_metadata.py
PBI AutoGov — Power BI Data Governance Automation Pipeline

Parses PBIP report files (JSON + TMDL) to extract every visual, field, filter,
and measure used in a report. Recursively resolves nested measure dependencies
to trace all underlying column references.

Input:  PBIP report definition root (pages/, report.json)
        Semantic model tables directory (for measure DAX lookup)
Output: pbi_report_metadata.xlsx
"""

import argparse
import json
import re
from pathlib import Path
from collections import Counter

import pandas as pd

from tmdl_parser import parse_tmdl_files
from bookmark_parser import parse_bookmarks, extract_single_filter


# ============================================================
# Visual type display names
# ============================================================

VISUAL_TYPE_DISPLAY = {
    "barChart": "Bar Chart",
    "clusteredBarChart": "Clustered Bar Chart",
    "clusteredColumnChart": "Clustered Column Chart",
    "stackedBarChart": "Stacked Bar Chart",
    "stackedColumnChart": "Stacked Column Chart",
    "hundredPercentStackedBarChart": "100% Stacked Bar Chart",
    "hundredPercentStackedColumnChart": "100% Stacked Column Chart",
    "lineChart": "Line Chart",
    "areaChart": "Area Chart",
    "stackedAreaChart": "Stacked Area Chart",
    "lineStackedColumnComboChart": "Line & Stacked Column Chart",
    "lineClusteredColumnComboChart": "Line & Clustered Column Chart",
    "ribbonChart": "Ribbon Chart",
    "waterfallChart": "Waterfall Chart",
    "funnelChart": "Funnel Chart",
    "pieChart": "Pie Chart",
    "donutChart": "Donut Chart",
    "treemap": "Treemap",
    "map": "Map",
    "filledMap": "Filled Map",
    "shapeMap": "Shape Map",
    "azureMap": "Azure Map",
    "tableEx": "Table",
    "pivotTable": "Matrix",
    "card": "Card",
    "multiRowCard": "Multi-Row Card",
    "kpi": "KPI",
    "gauge": "Gauge",
    "slicer": "Slicer",
    "scatterChart": "Scatter Chart",
    "decompositionTreeVisual": "Decomposition Tree",
    "keyDriversVisual": "Key Influencers",
    "qnaVisual": "Q&A",
    "scriptVisual": "R Script Visual",
    "pythonVisual": "Python Visual",
    "aiNarratives": "Smart Narrative",
    "paginator": "Paginated Report Visual",
    "cardVisual": "New Card",
    "advancedSlicerVisual": "New Slicer",
    "referenceLabel": "Reference Label",
}

# Visual types to skip (no data fields)
SKIP_VISUAL_TYPES = {
    "actionButton", "image", "textbox", "shape", "bookmarkNavigator",
    "pageNavigator", "groupShape",
}


# ============================================================
# Role → usage label mapping
# ============================================================

ROLE_USAGE_MAP = {
    # Slicer
    ("slicer", "Values"): "Slicer",
    ("advancedSlicerVisual", "Values"): "Slicer",
    # Table / Matrix
    ("tableEx", "Values"): "Visual Column",
    ("tableEx", "Rows"): "Visual Column",
    ("pivotTable", "Values"): "Visual Value",
    ("pivotTable", "Rows"): "Visual Row",
    ("pivotTable", "Columns"): "Visual Column",
    # Cards
    ("card", "Values"): "Visual Value",
    ("cardVisual", "Values"): "Visual Value",
    ("multiRowCard", "Values"): "Visual Value",
    ("kpi", "Value"): "Visual Value",
    ("kpi", "Goal"): "Visual Goal",
    ("kpi", "Trend"): "Visual Trend",
    # Gauge
    ("gauge", "Value"): "Visual Value",
    ("gauge", "MinValue"): "Visual Min",
    ("gauge", "MaxValue"): "Visual Max",
    ("gauge", "TargetValue"): "Visual Target",
}

DEFAULT_ROLE_MAP = {
    "Category": "Visual Column",
    "X": "Visual Value",
    "Y": "Visual Value",
    "Series": "Visual Column",
    "Values": "Visual Value",
    "Rows": "Visual Row",
    "Columns": "Visual Column",
    "Fields": "Visual Column",
    "Analyze": "Visual Value",
    "ExplainBy": "Visual Column",
    "Target": "Visual Column",
    "Location": "Visual Column",
    "Latitude": "Visual Column",
    "Longitude": "Visual Column",
    "Size": "Visual Value",
    "Color": "Visual Column",
    "Tooltips": "Visual Tooltip",
    "Value": "Visual Value",
    "Goal": "Visual Goal",
    "Trend": "Visual Trend",
}


def get_usage_label(vis_type: str, role: str, is_measure: bool) -> str:
    """Determine the usage label for a field based on visual type, role, and measure status."""
    base = ROLE_USAGE_MAP.get((vis_type, role))
    if not base:
        base = DEFAULT_ROLE_MAP.get(role, f"Visual {role}")
    if is_measure:
        return f"{base}, Filter (Measure)"
    return base


# ============================================================
# Field extraction from visual/filter JSON
# ============================================================

def extract_field_info(field: dict) -> list[dict]:
    """Extract table, column/measure, and type from a field definition.
    Used at visual, page, and report level.
    """
    results = []

    if "Column" in field:
        col = field["Column"]
        entity = _get_entity(col)
        prop = col.get("Property", "")
        results.append({"entity": entity, "property": prop, "field_type": "Column"})

    elif "Measure" in field:
        meas = field["Measure"]
        entity = _get_entity(meas)
        prop = meas.get("Property", "")
        results.append({"entity": entity, "property": prop, "field_type": "Measure"})

    elif "Aggregation" in field:
        agg = field["Aggregation"]
        expr = agg.get("Expression", {})
        if "Column" in expr:
            col = expr["Column"]
            entity = _get_entity(col)
            prop = col.get("Property", "")
            agg_func = _get_agg_name(agg.get("Function", 0))
            results.append({
                "entity": entity, "property": prop,
                "field_type": f"Aggregation ({agg_func})",
            })

    elif "HierarchyLevel" in field:
        hl = field["HierarchyLevel"]
        expr = hl.get("Expression", {})
        if "Hierarchy" in expr:
            hier = expr["Hierarchy"]
            entity = _get_entity(hier)
            hierarchy_name = hier.get("Hierarchy", "")
            level_name = hl.get("Level", "")
            prop = level_name or hierarchy_name

            # Fallback: resolve from PropertyVariationSource (auto-generated date hierarchies)
            if not entity:
                inner_expr = hier.get("Expression", {})
                pvs = inner_expr.get("PropertyVariationSource", {})
                if pvs:
                    entity = _get_entity(pvs)
                    prop = pvs.get("Property", prop)

            results.append({"entity": entity, "property": prop, "field_type": "HierarchyLevel"})

    return results


def _get_entity(node: dict) -> str:
    """Extract table name from a field node. Tries Expression > SourceRef > Entity."""
    try:
        return node["Expression"]["SourceRef"]["Entity"]
    except (KeyError, TypeError):
        return ""


def _get_agg_name(func_id: int) -> str:
    """Map Power BI aggregation function ID to readable name."""
    agg_map = {0: "Sum", 1: "Avg", 2: "Count", 3: "Min", 4: "Max",
               5: "CountNonNull", 6: "Median"}
    return agg_map.get(func_id, f"Func{func_id}")


# ============================================================
# Measure dependency resolution (recursive)
# ============================================================

def resolve_measure_dependencies(formula: str, measures_lookup: dict,
                                 visited: set = None) -> list[dict]:
    """Parse a DAX formula and identify all tables/columns it uses,
    including those from nested measures. Uses a visited set to prevent
    infinite loops from circular dependencies.
    """
    if visited is None:
        visited = set()

    dependencies = []

    # Find direct Table[Column] references
    # Pattern: 'TableName'[ColumnName] or TableName[ColumnName]
    direct_refs = re.findall(
        r"(?:'([^']+)'|([A-Za-z_][\w\s]*?))\[([^\]]+)\]",
        formula,
    )
    for quoted_table, unquoted_table, column in direct_refs:
        table = (quoted_table or unquoted_table).strip()
        col = column.strip()
        if table and col:
            dep = {"table": table, "column": col}
            if dep not in dependencies:
                dependencies.append(dep)

    # Find standalone [MeasureName] references (nested measures)
    nested_refs = re.findall(r"(?<!['\w\]])\[([^\]]+)\]", formula)

    for ref_name in nested_refs:
        ref_name = ref_name.strip()
        # Skip if already captured as a direct column reference
        if any(d["column"] == ref_name for d in dependencies):
            continue
        # Skip if already visited (prevents circular dependency loops)
        if ref_name in visited:
            continue
        visited.add(ref_name)

        # Look up this measure in the measures_lookup to find its DAX
        for (tbl, mname), sub_formula in measures_lookup.items():
            if mname == ref_name:
                # Include the nested measure itself as a dependency
                nested_dep = {"table": tbl, "column": mname}
                if nested_dep not in dependencies:
                    dependencies.append(nested_dep)

                # Recursively resolve the nested measure's dependencies
                sub_deps = resolve_measure_dependencies(sub_formula, measures_lookup, visited)
                for dep in sub_deps:
                    if dep not in dependencies:
                        dependencies.append(dep)
                break

    return dependencies


def get_measure_source_tables(entity: str, prop: str, measures_lookup: dict) -> list[dict]:
    """Get all source tables/columns for a measure, including nested dependencies.
    Groups columns by table for cleaner output.
    """
    formula = measures_lookup.get((entity, prop), "")
    if not formula:
        return [{"table": entity, "column": prop}]
    deps = resolve_measure_dependencies(formula, measures_lookup)
    if not deps:
        return [{"table": entity, "column": prop}]

    measure_dep = {"table": entity, "column": prop}
    if measure_dep not in deps:
        deps.insert(0, measure_dep)

    table_cols = {}
    for dep in deps:
        t, c = dep["table"], dep["column"]
        if t not in table_cols:
            table_cols[t] = []
        if c not in table_cols[t]:
            table_cols[t].append(c)
    return [{"table": t, "column": ", ".join(cols)} for t, cols in table_cols.items()]


# ============================================================
# Visual parser
# ============================================================

def get_visual_display_name(vis_type: str) -> str:
    """Convert camelCase visual type to human-readable name."""
    if vis_type in VISUAL_TYPE_DISPLAY:
        return VISUAL_TYPE_DISPLAY[vis_type]
    name = re.sub(r"([A-Z])", r" \1", vis_type).strip()
    return name.title()


def _get_visual_title(vis: dict) -> str:
    """Extract explicit title from visual JSON, if set."""
    try:
        titles = vis.get("visualContainerObjects", {}).get("title", [])
        for t in titles:
            text_expr = t.get("properties", {}).get("text", {})
            if "expr" in text_expr:
                val = text_expr["expr"].get("Literal", {}).get("Value", "")
                cleaned = val.strip("'")
                if cleaned:
                    return cleaned
    except (KeyError, TypeError, AttributeError):
        pass
    return ""


def _process_measure_field(page_name, vis_label, vis_type, display_name, usage, formula,
                           entity, prop, measures_lookup, visual_id=""):
    """Helper: resolve a measure field into output rows (handles nested dependencies)."""
    rows = []
    if formula:
        source_tables = get_measure_source_tables(entity, prop, measures_lookup)
        for st in source_tables:
            rows.append({
                "Page Name": page_name,
                "Visual/Table Name in PBI": vis_label,
                "Visual ID": visual_id,
                "Visual Type": vis_type,
                "UI Field Name": display_name,
                "Usage (Visual/Filter/Slicer)": usage,
                "Measure Formula": formula,
                "Table in the Semantic Model": st["table"],
                "Column in the Semantic Model": st["column"],
            })
    return rows


def parse_visual(visual_json: dict, page_name: str, measures_lookup: dict,
                 vis_type_counter: Counter, visual_id: str = "") -> list[dict]:
    """Parse a single visual.json and return rows for the output."""
    rows = []
    vis = visual_json.get("visual", {})
    vis_type = vis.get("visualType", "unknown")

    if vis_type in SKIP_VISUAL_TYPES:
        return rows

    # Determine visual label (title or auto-generated name)
    vis_title = _get_visual_title(vis)
    if vis_title:
        vis_label = vis_title
    else:
        vis_type_counter[vis_type] += 1
        count = vis_type_counter[vis_type]
        display_type = get_visual_display_name(vis_type)
        vis_label = display_type if count == 1 else f"{display_type} ({count})"

    # --- Query state fields (visual data roles) ---
    query_state = vis.get("query", {}).get("queryState", {})
    for role, role_data in query_state.items():
        projections = role_data.get("projections", [])
        for proj in projections:
            field = proj.get("field", {})
            display_name = proj.get("displayName", "")
            field_infos = extract_field_info(field)

            for fi in field_infos:
                if not display_name:
                    display_name = fi["property"]

                is_measure = fi["field_type"] == "Measure"
                formula = ""
                if is_measure:
                    formula = measures_lookup.get((fi["entity"], fi["property"]), "")

                usage = get_usage_label(vis_type, role, is_measure)

                if is_measure and formula:
                    rows.extend(_process_measure_field(
                        page_name, vis_label, vis_type, display_name, usage, formula,
                        fi["entity"], fi["property"], measures_lookup,
                        visual_id=visual_id,
                    ))
                else:
                    rows.append({
                        "Page Name": page_name,
                        "Visual/Table Name in PBI": vis_label,
                        "Visual ID": visual_id,
                        "Visual Type": vis_type,
                        "UI Field Name": display_name,
                        "Usage (Visual/Filter/Slicer)": usage,
                        "Measure Formula": formula,
                        "Table in the Semantic Model": fi["entity"],
                        "Column in the Semantic Model": fi["property"],
                    })

    # --- Collect fields already captured (to skip duplicate auto-generated filters) ---
    query_fields = set()
    for row in rows:
        query_fields.add((row["Table in the Semantic Model"], row["Column in the Semantic Model"]))

    # --- Visual-level filters ---
    vis_filters = visual_json.get("filterConfig", {}).get("filters", [])
    for flt in vis_filters:
        flt_field = flt.get("field", {})
        field_infos = extract_field_info(flt_field)
        for fi in field_infos:
            # Skip auto-generated filters that duplicate query state fields
            if (fi["entity"], fi["property"]) in query_fields:
                continue

            is_measure = fi["field_type"] == "Measure"
            formula = ""
            if is_measure:
                formula = measures_lookup.get((fi["entity"], fi["property"]), "")
                usage_str = "Filter (Measure)"
                if formula:
                    rows.extend(_process_measure_field(
                        page_name, vis_label, vis_type, fi["property"], usage_str, formula,
                        fi["entity"], fi["property"], measures_lookup,
                        visual_id=visual_id,
                    ))
                    continue
            else:
                usage_str = "Filter"

            rows.append({
                "Page Name": page_name,
                "Visual/Table Name in PBI": vis_label,
                "Visual ID": visual_id,
                "Visual Type": vis_type,
                "UI Field Name": fi["property"],
                "Usage (Visual/Filter/Slicer)": usage_str,
                "Measure Formula": formula,
                "Table in the Semantic Model": fi["entity"],
                "Column in the Semantic Model": fi["property"],
            })

    return rows


# ============================================================
# Page filter parser
# ============================================================

def parse_page_filters(page_json: dict, page_name: str, measures_lookup: dict) -> list[dict]:
    """Extract page-level filters."""
    rows = []
    filters = page_json.get("filterConfig", {}).get("filters", [])
    for flt in filters:
        flt_field = flt.get("field", {})
        field_infos = extract_field_info(flt_field)
        for fi in field_infos:
            is_measure = fi["field_type"] == "Measure"
            formula = ""
            if is_measure:
                formula = measures_lookup.get((fi["entity"], fi["property"]), "")
                usage_str = "Page Filter (Measure)"
                if formula:
                    rows.extend(_process_measure_field(
                        page_name, "Page Filters", "pageFilter", fi["property"],
                        usage_str, formula, fi["entity"], fi["property"], measures_lookup,
                        visual_id="",
                    ))
                    continue
            else:
                usage_str = "Page Filter"

            rows.append({
                "Page Name": page_name,
                "Visual/Table Name in PBI": "Page Filters",
                "Visual ID": "",
                "Visual Type": "pageFilter",
                "UI Field Name": fi["property"],
                "Usage (Visual/Filter/Slicer)": usage_str,
                "Measure Formula": formula,
                "Table in the Semantic Model": fi["entity"],
                "Column in the Semantic Model": fi["property"],
            })
    return rows


# ============================================================
# Filter expression extraction
# ============================================================

def extract_filter_expressions_from_list(filters: list, page_name: str,
                                         visual_name: str, visual_id: str,
                                         level: str) -> list[dict]:
    """Extract DAX filter expressions from a list of filterConfig.filters[] entries.

    Args:
        filters: List of filter objects from filterConfig.filters[]
        page_name: Page display name
        visual_name: Visual display name
        visual_id: Visual container ID
        level: Filter level ("Report", "Page", or "Visual")

    Returns:
        List of dicts for the Filter Expressions sheet.
    """
    results = []
    for flt in filters:
        # Determine the filter field for display
        flt_field = flt.get("field", {})
        field_infos = extract_field_info(flt_field)
        filter_field_str = ""
        if field_infos:
            fi = field_infos[0]
            entity = fi["entity"]
            prop = fi["property"]
            filter_field_str = f"'{entity}'[{prop}]" if entity else f"[{prop}]"

        # Check for TopN filter type
        filter_type = flt.get("type", "")
        if filter_type == "TopN":
            results.append({
                "Page Name": page_name,
                "Visual Name": visual_name,
                "Visual ID": visual_id,
                "Filter Level": level,
                "Filter Field": filter_field_str,
                "Filter DAX Expression": "-- TopN filter (not supported)",
            })
            continue

        # Extract DAX expressions using bookmark_parser's extract_single_filter
        dax_exprs = extract_single_filter(flt)
        if dax_exprs:
            for dax_expr in dax_exprs:
                results.append({
                    "Page Name": page_name,
                    "Visual Name": visual_name,
                    "Visual ID": visual_id,
                    "Filter Level": level,
                    "Filter Field": filter_field_str,
                    "Filter DAX Expression": dax_expr,
                })

    return results


# ============================================================
# Main extraction function
# ============================================================

def extract_metadata(report_root: str, model_root: str,
                     include_bookmarks: bool = True) -> tuple:
    """Main entry point: extract all metadata from a PBIP report.

    Args:
        report_root: Path to PBIP report definition root (contains pages/, report.json)
        model_root: Path to semantic model definition root (contains tables/)
        include_bookmarks: Whether to parse and include bookmark data (default True)

    Returns:
        Tuple of (metadata_df, bookmarks_list, filter_expressions) where
        bookmarks_list and filter_expressions may be empty.
    """
    tables_dir = Path(model_root) / "tables"
    pages_dir = Path(report_root) / "pages"

    print("=" * 60)
    print("PBI AutoGov — Metadata Extractor")
    print("=" * 60)

    # [1] Parse measures from semantic model
    print(f"\n[1] Parsing semantic model: {tables_dir}")
    measures_lookup = parse_tmdl_files(tables_dir)
    print(f"    Found {len(measures_lookup)} measures")
    if measures_lookup:
        print("    Sample measures:")
        for i, ((tbl, mname), _) in enumerate(list(measures_lookup.items())[:3]):
            print(f"      - {tbl}.{mname}")

    # [2] Parse report-level filters
    print(f"\n[2] Checking for report-level filters")
    report_json_path = Path(report_root) / "report.json"
    all_rows = []
    filter_expressions = []  # Accumulate filter DAX expressions for all levels

    if report_json_path.is_file():
        report_json = json.loads(report_json_path.read_text(encoding="utf-8-sig"))
        report_filters = report_json.get("filterConfig", {}).get("filters", [])
        if report_filters:
            # Extract filter DAX expressions at report level
            filter_expressions.extend(extract_filter_expressions_from_list(
                report_filters, "(All Pages)", "Report Filters", "", "Report",
            ))
            for flt in report_filters:
                flt_field = flt.get("field", {})
                field_infos = extract_field_info(flt_field)
                for fi in field_infos:
                    is_measure = fi["field_type"] == "Measure"
                    formula = ""
                    if is_measure:
                        formula = measures_lookup.get((fi["entity"], fi["property"]), "")
                        usage_str = "Report Filter (Measure)"
                        if formula:
                            all_rows.extend(_process_measure_field(
                                "(All Pages)", "Report Filters", "reportFilter",
                                fi["property"], usage_str, formula,
                                fi["entity"], fi["property"], measures_lookup,
                                visual_id="",
                            ))
                            continue
                    else:
                        usage_str = "Report Filter"

                    all_rows.append({
                        "Page Name": "(All Pages)",
                        "Visual/Table Name in PBI": "Report Filters",
                        "Visual ID": "",
                        "Visual Type": "reportFilter",
                        "UI Field Name": fi["property"],
                        "Usage (Visual/Filter/Slicer)": usage_str,
                        "Measure Formula": formula,
                        "Table in the Semantic Model": fi["entity"],
                        "Column in the Semantic Model": fi["property"],
                    })
            print(f"    Found {len(all_rows)} report-level filters")
        else:
            print("    No report-level filters found")
    else:
        print("    report.json not found, skipping")

    # [3] Parse report pages — also build ID→name mappings for bookmarks
    print(f"\n[3] Parsing report pages: {pages_dir}")
    if not pages_dir.is_dir():
        print(f"ERROR: Pages directory not found: {pages_dir}")
        return pd.DataFrame(), [], []

    # Mappings for bookmark resolution
    visual_id_to_name = {}      # visual folder name → display label
    page_id_to_name = {}        # page folder name (section ID) → display name
    page_id_to_visual_ids = {}  # page folder name → set of visual container IDs

    for page_folder in sorted(pages_dir.iterdir()):
        if not page_folder.is_dir():
            continue
        page_json_path = page_folder / "page.json"
        if not page_json_path.is_file():
            continue

        page_json = json.loads(page_json_path.read_text(encoding="utf-8-sig"))
        page_name = page_json.get("displayName", page_folder.name)
        print(f"\n    Page: {page_name}")

        # Track page ID → name mapping
        page_id_to_name[page_folder.name] = page_name
        page_id_to_visual_ids[page_folder.name] = set()

        # Page filters
        pf_rows = parse_page_filters(page_json, page_name, measures_lookup)
        all_rows.extend(pf_rows)
        print(f"      Page filters: {len(pf_rows)}")

        # Extract page-level filter DAX expressions
        page_filter_list = page_json.get("filterConfig", {}).get("filters", [])
        if page_filter_list:
            filter_expressions.extend(extract_filter_expressions_from_list(
                page_filter_list, page_name, "Page Filters", "", "Page",
            ))

        # Visuals
        visuals_dir = page_folder / "visuals"
        if not visuals_dir.is_dir():
            print("      No visuals directory found")
            continue

        vis_count = 0
        vis_type_counter = Counter()

        for vis_folder in sorted(visuals_dir.iterdir()):
            if not vis_folder.is_dir():
                continue
            vis_json_path = vis_folder / "visual.json"
            if not vis_json_path.is_file():
                continue

            vis_json = json.loads(vis_json_path.read_text(encoding="utf-8-sig"))

            # Track visual container ID for bookmark resolution
            page_id_to_visual_ids[page_folder.name].add(vis_folder.name)

            vis_rows = parse_visual(vis_json, page_name, measures_lookup, vis_type_counter,
                                    visual_id=vis_folder.name)
            all_rows.extend(vis_rows)
            if vis_rows:
                vis_count += 1
                vis_label = vis_rows[0]["Visual/Table Name in PBI"]
                # Map container ID → the visual label used in the first row
                visual_id_to_name[vis_folder.name] = vis_label

                # Extract visual-level filter DAX expressions
                vis_filters = vis_json.get("filterConfig", {}).get("filters", [])
                if vis_filters:
                    filter_expressions.extend(extract_filter_expressions_from_list(
                        vis_filters, page_name, vis_label, vis_folder.name, "Visual",
                    ))
            else:
                # Even data-less visuals (buttons, images) get their type as name
                vis = vis_json.get("visual", {})
                vis_type = vis.get("visualType", "unknown")
                visual_id_to_name[vis_folder.name] = _get_visual_title(vis) or vis_type

        print(f"      Visuals with data: {vis_count}")

    # Build output DataFrame
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
    print(f"Pages: {df['Page Name'].nunique()}")
    print(f"Visuals: {df.groupby(['Page Name', 'Visual/Table Name in PBI']).ngroups}")
    print(f"{'=' * 60}")

    # [4] Parse bookmarks (if present and enabled)
    bookmarks_list = []
    if include_bookmarks:
        bookmarks_dir = Path(report_root) / "bookmarks"
        if bookmarks_dir.is_dir():
            print(f"\n[4] Parsing bookmarks: {bookmarks_dir}")
            bookmarks_list = parse_bookmarks(
                report_root, visual_id_to_name,
                page_id_to_name, page_id_to_visual_ids,
            )
            if bookmarks_list:
                print(f"    Found {len(bookmarks_list)} bookmarks")
                for bm in bookmarks_list:
                    vis_count = sum(1 for v in bm.visuals if v.visible)
                    print(f"      - {bm.name}: {len(bm.filters)} filters, "
                          f"{vis_count}/{len(bm.visuals)} visuals visible")
            else:
                print("    No bookmarks found")
        else:
            print(f"\n[4] No bookmarks folder found")

    if filter_expressions:
        print(f"\n[5] Filter expressions extracted: {len(filter_expressions)}")
    else:
        print(f"\n[5] No filter expressions extracted")

    return df, bookmarks_list, filter_expressions


def export_to_excel(df: pd.DataFrame, output_path: str, bookmarks_list: list = None,
                    filter_expressions: list = None):
    """Save metadata DataFrame to Excel with auto-sized columns.

    If bookmarks_list is provided, adds a 'Bookmarks' sheet.
    If filter_expressions is provided, adds a 'Filter Expressions' sheet.
    """
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        # --- Report Metadata sheet ---
        df.to_excel(writer, sheet_name="Report Metadata", index=False)
        ws = writer.sheets["Report Metadata"]
        for col_idx, col_name in enumerate(df.columns, 1):
            max_len = max(
                len(str(col_name)),
                df[col_name].astype(str).str.len().max() if len(df) > 0 else 0,
            )
            ws.column_dimensions[ws.cell(row=1, column=col_idx).column_letter].width = min(max_len + 2, 60)

        # --- Bookmarks sheet ---
        if bookmarks_list:
            bm_rows = []
            for bm in bookmarks_list:
                filter_dax = "; ".join(bm.filters) if bm.filters else ""
                for vis in bm.visuals:
                    bm_rows.append({
                        "Bookmark Name": bm.name,
                        "Page Name": bm.page_name,
                        "Visual Container ID": vis.container_id,
                        "Visual Name": vis.visual_name,
                        "Visible": "Y" if vis.visible else "N",
                        "Filter DAX": filter_dax,
                    })

            if bm_rows:
                bm_df = pd.DataFrame(bm_rows, columns=[
                    "Bookmark Name", "Page Name", "Visual Container ID",
                    "Visual Name", "Visible", "Filter DAX",
                ])
                bm_df.to_excel(writer, sheet_name="Bookmarks", index=False)
                ws_bm = writer.sheets["Bookmarks"]
                for col_idx, col_name in enumerate(bm_df.columns, 1):
                    max_len = max(
                        len(str(col_name)),
                        bm_df[col_name].astype(str).str.len().max() if len(bm_df) > 0 else 0,
                    )
                    ws_bm.column_dimensions[ws_bm.cell(row=1, column=col_idx).column_letter].width = min(max_len + 2, 60)
                print(f"  Bookmarks sheet: {len(bm_rows)} rows")

        # --- Filter Expressions sheet ---
        if filter_expressions:
            fe_df = pd.DataFrame(filter_expressions, columns=[
                "Page Name", "Visual Name", "Visual ID",
                "Filter Level", "Filter Field", "Filter DAX Expression",
            ])
            fe_df.to_excel(writer, sheet_name="Filter Expressions", index=False)
            ws_fe = writer.sheets["Filter Expressions"]
            for col_idx, col_name in enumerate(fe_df.columns, 1):
                max_len = max(
                    len(str(col_name)),
                    fe_df[col_name].astype(str).str.len().max() if len(fe_df) > 0 else 0,
                )
                ws_fe.column_dimensions[ws_fe.cell(row=1, column=col_idx).column_letter].width = min(max_len + 2, 60)
            print(f"  Filter Expressions sheet: {len(filter_expressions)} rows")

    print(f"\nExcel file saved to: {output_path}")


# ============================================================
# Standalone execution
# ============================================================

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="PBI AutoGov — Metadata Extractor")
    parser.add_argument("--report-root", required=True, help="Path to PBIP report definition root")
    parser.add_argument("--model-root", required=True, help="Path to semantic model definition root")
    parser.add_argument("--output", default="pbi_report_metadata.xlsx", help="Output Excel file path")
    parser.add_argument("--no-bookmarks", action="store_true",
                        help="Disable bookmark extraction")
    args = parser.parse_args()

    df, bookmarks_list, filter_expressions = extract_metadata(
        args.report_root, args.model_root,
        include_bookmarks=not args.no_bookmarks,
    )
    if not df.empty:
        export_to_excel(df, args.output, bookmarks_list, filter_expressions)
    else:
        print("No data extracted. Check paths.")
