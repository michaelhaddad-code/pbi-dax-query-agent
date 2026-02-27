"""
PBIX to PBIP Converter

Extracts a .pbix ZIP archive into the PBIP folder structure consumed by
extract_metadata.py. Report structure (pages, visuals, filters, bookmarks)
is extracted with pure Python. Semantic model (measures, columns) requires
the optional `pbixray` package.

Usage:
    python skills/pbix_extractor.py "path/to/report.pbix" --output "data/"
    python skills/pbix_extractor.py "path/to/report.pbix" --output "data/" --model-root "path/to/SemanticModel/definition"
"""

import argparse
import json
import logging
import os
import re
import sys
import zipfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Optional

logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")
logger = logging.getLogger(__name__)

# Try importing pbixray at module level — optional dependency
HAS_PBIXRAY = False
PBIXRAY_ERROR = ""
try:
    from pbixray import PBIXRay
    HAS_PBIXRAY = True
except ImportError as e:
    PBIXRAY_ERROR = str(e)
except OSError as e:
    # OSError typically means a missing C compiler or shared library
    PBIXRAY_ERROR = (
        f"pbixray failed to load (likely missing C compiler): {e}. "
        "Install a C compiler (e.g. Visual Studio Build Tools on Windows, "
        "gcc on Linux) and reinstall pbixray, or use PBIP format instead."
    )
    logger.warning(PBIXRAY_ERROR)


@dataclass
class PbixExtractResult:
    """Result of a .pbix extraction."""

    report_root: str
    model_root: Optional[str]
    report_name: str
    page_count: int = 0
    visual_container_count: int = 0
    data_visual_count: int = 0  # visuals with queryState (actual data visuals)
    bookmark_count: int = 0
    semantic_model_source: str = "none"  # "pbixray" | "user-provided" | "none"


def read_layout_json(pbix_path: str) -> dict:
    """Read and parse the Report/Layout JSON from a .pbix ZIP.

    The Layout file is UTF-16LE encoded (with BOM).
    """
    with zipfile.ZipFile(pbix_path, "r") as zf:
        names = zf.namelist()
        # Find the layout file — typically "Report/Layout" but handle casing
        layout_name = None
        for name in names:
            if name.lower() == "report/layout":
                layout_name = name
                break
        if layout_name is None:
            raise FileNotFoundError(
                f"Could not find Report/Layout in {pbix_path}. "
                f"Available entries: {names[:20]}"
            )
        raw = zf.read(layout_name)

    # Decode UTF-16LE (handles BOM automatically)
    try:
        text = raw.decode("utf-16-le")
    except UnicodeDecodeError:
        # Fallback: try utf-8
        text = raw.decode("utf-8-sig")

    # Strip BOM if present
    if text and text[0] == "\ufeff":
        text = text[1:]

    return json.loads(text)


def safe_json_loads(s: Any) -> Any:
    """Parse a stringified JSON field. Returns None on failure."""
    if not s or not isinstance(s, str):
        return None
    try:
        return json.loads(s)
    except (json.JSONDecodeError, TypeError):
        return None


def sanitize_filename(name: str) -> str:
    """Remove characters that are invalid in Windows/Linux filenames."""
    # Replace invalid chars with underscore
    return re.sub(r'[<>:"/\\|?*]', "_", name).strip()


def normalize_filters(filters: list) -> list:
    """Normalize .pbix filter objects to match PBIP format.

    Three transformations:
    1. Rename `expression` → `field` (PBIP key name)
    2. Resolve SourceRef.Source (alias) → SourceRef.Entity (table name) in the field
    3. If no `field`/`expression` exists but `filter.Where` does, synthesize `field`
       from the first Where condition's column reference
    4. Default `type` to "Categorical" if missing
    """
    for f in filters:
        # Build alias→entity map from this filter's From array
        alias_map = {}
        filt = f.get("filter", {})
        for entry in filt.get("From", []):
            alias = entry.get("Name", "")
            entity = entry.get("Entity", "")
            if alias and entity:
                alias_map[alias] = entity

        # Step 1: rename expression → field
        if "expression" in f and "field" not in f:
            f["field"] = f.pop("expression")

        # Step 2: if field exists, resolve aliases in it
        if "field" in f and alias_map:
            _resolve_source_refs(f["field"], alias_map)

        # Step 3: if no field at all, synthesize from filter.Where
        if "field" not in f and filt:
            synthesized = _synthesize_field_from_where(filt, alias_map)
            if synthesized:
                f["field"] = synthesized

        # Step 4: default type
        if "type" not in f:
            f["type"] = "Categorical"

    return filters


def _synthesize_field_from_where(filt: dict, alias_map: dict) -> Optional[dict]:
    """Extract a field reference from a filter's Where clause.

    Walks Where conditions to find the first Column or Measure reference,
    resolves aliases, and returns it in PBIP field format. Checks all
    conditions (not just the first) in case the first has no extractable field.
    """
    where_list = filt.get("Where", [])
    if not where_list:
        return None

    for where_entry in where_list:
        condition = where_entry.get("Condition", {})
        field_ref = _find_field_in_condition(condition)
        if field_ref:
            field_ref = json.loads(json.dumps(field_ref))  # deep copy
            _resolve_source_refs(field_ref, alias_map)
            return field_ref

    return None


def _find_field_in_condition(condition: dict) -> Optional[dict]:
    """Recursively find the first Column or Measure field in a filter condition."""
    if not condition:
        return None

    # Comparison: check Left side
    if "Comparison" in condition:
        comp = condition["Comparison"]
        left = comp.get("Left", {})
        for ftype in ("Column", "Measure"):
            if ftype in left:
                return {ftype: left[ftype]}

    # In: check Expressions[0]
    if "In" in condition:
        exprs = condition["In"].get("Expressions", [])
        if exprs:
            for ftype in ("Column", "Measure"):
                if ftype in exprs[0]:
                    return {ftype: exprs[0][ftype]}

    # Between: check Left (same shape as Comparison)
    if "Between" in condition:
        left = condition["Between"].get("Left", {})
        for ftype in ("Column", "Measure"):
            if ftype in left:
                return {ftype: left[ftype]}

    # Contains / StartsWith / EndsWith: check Left.Column or Left.Measure
    for op in ("Contains", "StartsWith", "EndsWith"):
        if op in condition:
            left = condition[op].get("Left", {})
            for ftype in ("Column", "Measure"):
                if ftype in left:
                    return {ftype: left[ftype]}

    # Not: unwrap and recurse
    if "Not" in condition:
        inner = condition["Not"]
        return _find_field_in_condition(inner)

    # And: check Left branch
    if "And" in condition:
        return _find_field_in_condition(condition["And"].get("Left", {}))

    # Or: check Left branch
    if "Or" in condition:
        return _find_field_in_condition(condition["Or"].get("Left", {}))

    return None


# ---------------------------------------------------------------------------
# Report-level extraction
# ---------------------------------------------------------------------------


def build_report_json(layout: dict) -> dict:
    """Build the PBIP report.json from the layout's top-level config."""
    report = {
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/report/3.1.0/schema.json"
    }

    # Parse top-level config (stringified JSON)
    config = safe_json_loads(layout.get("config", ""))
    if config:
        if "themeCollection" in config:
            report["themeCollection"] = config["themeCollection"]
        if "objects" in config:
            report["objects"] = config["objects"]
        if "settings" in config:
            report["settings"] = config["settings"]
        if "slowDataSourceSettings" in config:
            report["slowDataSourceSettings"] = config["slowDataSourceSettings"]

    # Resource packages (already parsed, not stringified)
    if "resourcePackages" in layout:
        report["resourcePackages"] = layout["resourcePackages"]

    return report


def build_report_filters(layout: dict) -> list:
    """Extract report-level filters from the layout's top-level filters field."""
    filters_str = layout.get("filters")
    if not filters_str:
        return []
    filters = safe_json_loads(filters_str)
    if isinstance(filters, list):
        return normalize_filters(filters)
    return []


# ---------------------------------------------------------------------------
# Page extraction
# ---------------------------------------------------------------------------


def build_pages_json(sections: list) -> dict:
    """Build pages/pages.json from the sections array."""
    # Sort by ordinal to get display order
    sorted_sections = sorted(sections, key=lambda s: s.get("ordinal", 0))
    page_order = [s["name"] for s in sorted_sections if "name" in s]

    return {
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/pagesMetadata/1.0.0/schema.json",
        "pageOrder": page_order,
        "activePageName": page_order[0] if page_order else "",
    }


def build_page_json(section: dict) -> dict:
    """Build a single page.json from a .pbix section entry."""
    page = {
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/page/2.0.0/schema.json",
        "name": section.get("name", ""),
        "displayName": section.get("displayName", section.get("name", "")),
    }

    # Dimensions
    if "width" in section:
        page["width"] = section["width"]
    if "height" in section:
        page["height"] = section["height"]
    if "displayOption" in section:
        page["displayOption"] = section["displayOption"]

    # Parse page config (stringified)
    config = safe_json_loads(section.get("config", ""))
    if config:
        if "objects" in config:
            page["objects"] = config["objects"]
        if "displayOption" in config and "displayOption" not in page:
            page["displayOption"] = config["displayOption"]

    # Parse page-level filters (stringified)
    filters = safe_json_loads(section.get("filters", ""))
    if isinstance(filters, list) and filters:
        page["filterConfig"] = {"filters": normalize_filters(filters)}

    # Page binding (drillthrough config) — already in parsed config
    if config and "pageBinding" in config:
        page["pageBinding"] = config["pageBinding"]

    return page


# ---------------------------------------------------------------------------
# Visual extraction
# ---------------------------------------------------------------------------


def build_visual_json(vc: dict) -> dict:
    """Build a single visual.json from a .pbix visual container.

    The .pbix visual container has stringified `config`, `filters`, and `query`
    fields that must be parsed and restructured into PBIP format.
    """
    config = safe_json_loads(vc.get("config", ""))
    if not config:
        logger.warning("Visual container has no parseable config, skipping")
        return {}

    visual_json = {
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/visualContainer/2.5.0/schema.json"
    }

    # Visual name (container ID)
    name = config.get("name", "")
    if name:
        visual_json["name"] = name

    # Position
    position = {}
    for key in ("x", "y", "z", "width", "height", "tabOrder"):
        if key in vc:
            position[key] = vc[key]
    if position:
        visual_json["position"] = position

    # Parent group reference
    if "parentGroupName" in config:
        visual_json["parentGroupName"] = config["parentGroupName"]

    # Check for singleVisual vs singleVisualGroup
    sv = config.get("singleVisual")
    svg = config.get("singleVisualGroup")

    if sv:
        visual_block = _build_visual_block(sv)
        if visual_block:
            visual_json["visual"] = visual_block
    elif svg:
        # Group container — create a minimal visual block with type "group"
        visual_block = {"visualType": "group"}
        if "displayName" in svg:
            visual_block["displayName"] = svg["displayName"]
        visual_json["visual"] = visual_block
        # Groups can contain children — these are handled as separate visual containers
        # by the caller (extract_visuals_from_section)

    # Filters (stringified in .pbix) — visual-level filters
    filters = safe_json_loads(vc.get("filters", ""))
    if isinstance(filters, list) and filters:
        visual_json["filterConfig"] = {"filters": normalize_filters(filters)}

    # Fallback: if visual still has no query, try the top-level `query` field
    # on the visual container (Commands format). Convert it to queryState.
    if "visual" in visual_json and "query" not in visual_json.get("visual", {}):
        vc_query = safe_json_loads(vc.get("query", ""))
        if vc_query:
            visual_type = visual_json.get("visual", {}).get("visualType", "")
            sv = config.get("singleVisual", {}) if config else {}
            col_props = sv.get("columnProperties", {}) if sv else {}
            commands_qs = _convert_commands_to_query_state(
                vc_query, visual_type, col_props
            )
            if commands_qs:
                visual_json["visual"]["query"] = {"queryState": commands_qs}

    return visual_json


def _build_visual_block(sv: dict) -> dict:
    """Build the `visual` object from a singleVisual config block.

    The .pbix singleVisual stores query data in a different format than PBIP:
      - .pbix: `prototypeQuery` (From/Select) + `projections` (role→queryRef) + `columnProperties` (displayNames)
      - PBIP:  `query.queryState` (role→projections with resolved field + queryRef + displayName)
      - .pbix: `vcObjects` → PBIP: `visualContainerObjects`

    This function converts from .pbix format to PBIP format.
    """
    visual = {}

    if "visualType" in sv:
        visual["visualType"] = sv["visualType"]

    # Build queryState from prototypeQuery + projections (the .pbix → PBIP conversion)
    query_state = _build_query_state(sv)
    if query_state:
        visual["query"] = {"queryState": query_state}
    elif "query" in sv:
        # Fallback: try converting Commands format → queryState
        commands_qs = _convert_commands_to_query_state(
            sv["query"],
            sv.get("visualType", ""),
            sv.get("columnProperties", {}),
        )
        if commands_qs:
            visual["query"] = {"queryState": commands_qs}
        # else: visual has no usable query — extract_metadata will skip it

    if "objects" in sv:
        visual["objects"] = sv["objects"]

    # .pbix uses "vcObjects", PBIP uses "visualContainerObjects"
    vc_objects = sv.get("vcObjects") or sv.get("visualContainerObjects")
    if vc_objects:
        visual["visualContainerObjects"] = vc_objects

    if "drillFilterOtherVisuals" in sv:
        visual["drillFilterOtherVisuals"] = sv["drillFilterOtherVisuals"]

    if "orderBy" in sv:
        visual["orderBy"] = sv["orderBy"]

    if "activeProjections" in sv:
        visual["activeProjections"] = sv["activeProjections"]

    return visual


def _build_query_state(sv: dict) -> Optional[dict]:
    """Convert .pbix prototypeQuery + projections into PBIP queryState format.

    .pbix structure:
        singleVisual.projections = {
            "Category": [{"queryRef": "SalesStage.Sales Stage", "active": true}],
            "Y": [{"queryRef": "Fact.Opportunity Count"}]
        }
        singleVisual.prototypeQuery = {
            "From": [{"Name": "s", "Entity": "SalesStage", "Type": 0}, ...],
            "Select": [
                {"Column": {"Expression": {"SourceRef": {"Source": "s"}}, "Property": "Sales Stage"}, "Name": "SalesStage.Sales Stage"},
                {"Measure": {"Expression": {"SourceRef": {"Source": "f"}}, "Property": "Opportunity Count"}, "Name": "Fact.Opportunity Count"}
            ]
        }
        singleVisual.columnProperties = {"Fact.Revenue": {"displayName": "Opportunity Revenue"}}

    PBIP target:
        query.queryState = {
            "Category": {"projections": [{"field": {"Column": {..., "Entity": "SalesStage"}}, "queryRef": "...", "active": true}]},
            "Y": {"projections": [{"field": {"Measure": {..., "Entity": "Fact"}}, "queryRef": "..."}]}
        }
    """
    projections_map = sv.get("projections")
    prototype_query = sv.get("prototypeQuery")

    if not projections_map or not prototype_query:
        return None

    # Build alias → entity map from From array
    from_entries = prototype_query.get("From", [])
    alias_to_entity = {}
    for entry in from_entries:
        alias = entry.get("Name", "")
        entity = entry.get("Entity", "")
        if alias and entity:
            alias_to_entity[alias] = entity

    # Build queryRef → Select entry map
    select_entries = prototype_query.get("Select", [])
    queryref_to_select = {}
    for sel in select_entries:
        name = sel.get("Name", "")
        if name:
            queryref_to_select[name] = sel

    # Display name overrides
    column_properties = sv.get("columnProperties", {})

    # Build queryState
    query_state = {}
    for role_name, role_projections in projections_map.items():
        if not isinstance(role_projections, list):
            continue
        pbip_projections = []
        for proj in role_projections:
            query_ref = proj.get("queryRef", "")
            sel = queryref_to_select.get(query_ref, {})

            # Build the field object with resolved entity references
            field_obj = _resolve_field_from_select(sel, alias_to_entity)

            pbip_proj: dict[str, Any] = {}
            if field_obj:
                pbip_proj["field"] = field_obj
            if query_ref:
                pbip_proj["queryRef"] = query_ref

            # Display name from columnProperties or nativeQueryRef
            display_name = (
                column_properties.get(query_ref, {}).get("displayName")
                if isinstance(column_properties.get(query_ref), dict)
                else None
            )
            if not display_name:
                display_name = sel.get("nativeQueryRef")
            if display_name:
                pbip_proj["displayName"] = display_name

            # Active flag
            if "active" in proj:
                pbip_proj["active"] = proj["active"]

            if pbip_proj:
                pbip_projections.append(pbip_proj)

        if pbip_projections:
            query_state[role_name] = {"projections": pbip_projections}

    return query_state if query_state else None


# Role inference: map visual type + field type → queryState role name
# Columns are grouping fields, Measures/Aggregations are value fields.
# Types not listed default to Category (grouping) / Y (values), which is
# correct for standard charts (bar, column, line, pie, area, scatter, combo,
# waterfall, ribbon, funnel, donut).
_GROUPING_ROLE = {
    "card": "Values",
    "cardVisual": "Values",
    "multiRowCard": "Values",
    "kpi": "Values",
    "gauge": "Values",
    "slicer": "Values",
    "tableEx": "Values",
    "pivotTable": "Rows",
    "matrix": "Rows",
    "treemap": "Group",
    "decompositionTreeVisual": "Category",
    "qnaVisual": "Category",
}
_VALUE_ROLE = {
    "pivotTable": "Values",
    "matrix": "Values",
    "treemap": "Values",
    "decompositionTreeVisual": "Y",
    "qnaVisual": "Y",
}


def _convert_commands_to_query_state(
    query: dict,
    visual_type: str,
    column_properties: Optional[dict] = None,
) -> Optional[dict]:
    """Convert a SemanticQueryDataShapeCommand query into PBIP queryState.

    This is the fallback path for visuals that don't have prototypeQuery+projections.
    Role names are inferred from visual type and field type (Column vs Measure).
    column_properties (from singleVisual.columnProperties) provides display name
    overrides keyed by queryRef.
    """
    # Navigate to the inner Query object
    commands = query.get("Commands", [])
    if not commands:
        return None
    sqds = commands[0].get("SemanticQueryDataShapeCommand", {})
    inner_query = sqds.get("Query", {})
    if not inner_query:
        return None

    # Build alias → entity map
    alias_to_entity = {}
    for entry in inner_query.get("From", []):
        alias = entry.get("Name", "")
        entity = entry.get("Entity", "")
        if alias and entity:
            alias_to_entity[alias] = entity

    select_entries = inner_query.get("Select", [])
    if not select_entries:
        return None

    if column_properties is None:
        column_properties = {}

    # Determine the role for grouping (Column) vs value (Measure) fields
    grouping_role = _GROUPING_ROLE.get(visual_type, "Category")
    value_role = _VALUE_ROLE.get(visual_type, "Y")

    # Build projections grouped by role
    role_projections: dict[str, list] = {}
    for sel in select_entries:
        field_obj = _resolve_field_from_select(sel, alias_to_entity)
        if not field_obj:
            continue

        query_ref = sel.get("Name", "")

        # Determine role based on field type
        field_type = next(iter(field_obj))  # "Column", "Measure", etc.
        if field_type in ("Measure", "Aggregation"):
            role = value_role
        else:
            role = grouping_role

        proj: dict[str, Any] = {"field": field_obj}
        if query_ref:
            proj["queryRef"] = query_ref

        # Display name: prefer columnProperties override, then nativeQueryRef
        display_name = (
            column_properties.get(query_ref, {}).get("displayName")
            if isinstance(column_properties.get(query_ref), dict)
            else None
        )
        if not display_name:
            display_name = sel.get("nativeQueryRef")
        if display_name:
            proj["displayName"] = display_name

        role_projections.setdefault(role, []).append(proj)

    # For cards/kpis/gauges where both columns and measures go to "Values",
    # the grouping_role == value_role == "Values" so they merge naturally.

    query_state = {}
    for role_name, projs in role_projections.items():
        query_state[role_name] = {"projections": projs}

    return query_state if query_state else None


def _resolve_field_from_select(sel: dict, alias_to_entity: dict) -> Optional[dict]:
    """Extract and resolve a field object from a .pbix Select entry.

    Converts SourceRef.Source (alias) → SourceRef.Entity (table name).
    Handles Column, Measure, Aggregation, and HierarchyLevel field types.
    """
    # Determine which field type is present
    for field_type in ("Column", "Measure", "Aggregation", "HierarchyLevel"):
        if field_type in sel:
            field_data = json.loads(json.dumps(sel[field_type]))  # deep copy
            _resolve_source_refs(field_data, alias_to_entity)
            return {field_type: field_data}
    return None


def _resolve_source_refs(obj: Any, alias_to_entity: dict) -> None:
    """Recursively resolve SourceRef.Source (alias) → SourceRef.Entity (table name)."""
    if isinstance(obj, dict):
        if "SourceRef" in obj and "Source" in obj["SourceRef"]:
            alias = obj["SourceRef"]["Source"]
            entity = alias_to_entity.get(alias, alias)
            # Replace Source with Entity (PBIP format)
            obj["SourceRef"] = {"Entity": entity}
        for value in obj.values():
            _resolve_source_refs(value, alias_to_entity)
    elif isinstance(obj, list):
        for item in obj:
            _resolve_source_refs(item, alias_to_entity)


def extract_visuals_from_section(section: dict) -> list[tuple[str, dict]]:
    """Extract all visual containers from a section, returning (visual_id, visual_json) pairs.

    Handles grouped visuals: group containers and their children are all returned
    as top-level entries (children get parentGroupName set).
    """
    results = []
    vcs = section.get("visualContainers", [])

    for vc in vcs:
        config = safe_json_loads(vc.get("config", ""))
        if not config:
            continue

        visual_id = config.get("name", "")
        if not visual_id:
            logger.warning("Visual container has no name in config, skipping")
            continue

        visual_json = build_visual_json(vc)
        if visual_json:
            results.append((visual_id, visual_json))

        # Handle grouped visuals — if this is a group, extract children
        svg = config.get("singleVisualGroup")
        if svg and "children" in svg:
            for child_vc in svg["children"]:
                child_config = safe_json_loads(child_vc.get("config", ""))
                if not child_config:
                    continue
                child_id = child_config.get("name", "")
                if not child_id:
                    continue
                # Ensure child references the parent group
                if "parentGroupName" not in child_config:
                    child_config["parentGroupName"] = visual_id
                    child_vc["config"] = json.dumps(child_config)
                child_json = build_visual_json(child_vc)
                if child_json:
                    results.append((child_id, child_json))

    return results


# ---------------------------------------------------------------------------
# Bookmark extraction
# ---------------------------------------------------------------------------


def extract_bookmarks(layout: dict) -> list[dict]:
    """Extract bookmarks from the layout JSON.

    Bookmarks can be in:
    1. Top-level config (stringified) → config.bookmarks
    2. Top-level 'bookmarks' key (PBI version variance)

    Returns a list of bookmark objects.
    """
    bookmarks = []

    # Try config.bookmarks first
    config = safe_json_loads(layout.get("config", ""))
    if config and "bookmarks" in config:
        bm_list = config["bookmarks"]
        if isinstance(bm_list, list):
            bookmarks = bm_list
            logger.info(f"Found {len(bookmarks)} bookmarks in config.bookmarks")
        else:
            logger.warning(
                f"config.bookmarks exists but is not a list: {type(bm_list)}"
            )

    # Fallback: top-level bookmarks key
    if not bookmarks:
        top_bm = layout.get("bookmarks")
        if top_bm:
            if isinstance(top_bm, list):
                bookmarks = top_bm
                logger.info(
                    f"Found {len(bookmarks)} bookmarks in top-level bookmarks key"
                )
            elif isinstance(top_bm, str):
                parsed = safe_json_loads(top_bm)
                if isinstance(parsed, list):
                    bookmarks = parsed
                    logger.info(
                        f"Found {len(bookmarks)} bookmarks in top-level bookmarks key (stringified)"
                    )

    if not bookmarks:
        logger.info(
            "No bookmarks found in layout JSON (checked config.bookmarks and top-level bookmarks)"
        )

    return bookmarks


def build_bookmarks_index(bookmarks: list[dict]) -> dict:
    """Build bookmarks/bookmarks.json index file."""
    items = []
    for bm in bookmarks:
        name = bm.get("name", "")
        if name:
            items.append({"name": name})
    return {
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/bookmarksMetadata/1.0.0/schema.json",
        "items": items,
    }


def build_bookmark_file(bm: dict) -> dict:
    """Build an individual bookmark .bookmark.json file from layout bookmark data."""
    bookmark = {
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/bookmark/2.0.0/schema.json"
    }

    # Copy standard fields
    for key in ("displayName", "name", "options", "explorationState"):
        if key in bm:
            bookmark[key] = bm[key]

    return bookmark


# ---------------------------------------------------------------------------
# Semantic model extraction (requires pbixray)
# ---------------------------------------------------------------------------

# Mapping from pandas/pbixray data types to TMDL dataType values
DTYPE_MAP = {
    "int64": "int64",
    "Int64": "int64",
    "float64": "double",
    "Float64": "double",
    "object": "string",
    "string": "string",
    "bool": "boolean",
    "boolean": "boolean",
    "datetime64[ns]": "dateTime",
    "datetime64": "dateTime",
}


def map_data_type(dtype_str: str) -> str:
    """Map a pandas/pbixray dtype string to a TMDL dataType value."""
    dtype_str = str(dtype_str).strip()
    if dtype_str in DTYPE_MAP:
        return DTYPE_MAP[dtype_str]
    if "int" in dtype_str.lower():
        return "int64"
    if "float" in dtype_str.lower() or "double" in dtype_str.lower():
        return "double"
    if "date" in dtype_str.lower() or "time" in dtype_str.lower():
        return "dateTime"
    if "bool" in dtype_str.lower():
        return "boolean"
    if "decimal" in dtype_str.lower():
        return "decimal"
    return "string"


def extract_semantic_model_pbixray(pbix_path: str, model_dir: Path) -> bool:
    """Extract semantic model from .pbix using pbixray and write synthetic TMDL files.

    Returns True if extraction succeeded, False otherwise.
    """
    if not HAS_PBIXRAY:
        return False

    try:
        pbix = PBIXRay(pbix_path)
    except Exception as e:
        logger.warning(f"pbixray failed to open {pbix_path}: {e}")
        return False

    tables_dir = model_dir / "tables"
    tables_dir.mkdir(parents=True, exist_ok=True)

    # Collect measures by table
    measures_by_table: dict[str, list[tuple[str, str]]] = {}
    try:
        dax_measures = pbix.dax_measures
        if dax_measures is not None and not dax_measures.empty:
            for _, row in dax_measures.iterrows():
                table_name = str(row.get("TableName", row.get("tableName", "")))
                measure_name = str(row.get("Name", row.get("name", "")))
                expression = str(row.get("Expression", row.get("expression", "")))
                if table_name and measure_name and expression:
                    measures_by_table.setdefault(table_name, []).append(
                        (measure_name, expression)
                    )
            logger.info(
                f"pbixray: extracted {len(dax_measures)} measures from {len(measures_by_table)} tables"
            )
    except Exception as e:
        logger.warning(f"pbixray: could not extract DAX measures: {e}")

    # Collect columns by table
    columns_by_table: dict[str, list[tuple[str, str]]] = {}
    try:
        schema = pbix.schema
        if schema is not None and not schema.empty:
            for _, row in schema.iterrows():
                table_name = str(row.get("TableName", row.get("tableName", "")))
                col_name = str(row.get("ColumnName", row.get("columnName", "")))
                dtype = str(
                    row.get(
                        "DataType", row.get("dataType", row.get("dtype", "string"))
                    )
                )
                if table_name and col_name:
                    columns_by_table.setdefault(table_name, []).append(
                        (col_name, map_data_type(dtype))
                    )
            logger.info(
                f"pbixray: extracted columns from {len(columns_by_table)} tables"
            )
    except Exception as e:
        logger.warning(f"pbixray: could not extract schema: {e}")

    # Extract relationships
    relationships = []
    try:
        rel_data = pbix.relationships
        if rel_data is not None and not rel_data.empty:
            for _, row in rel_data.iterrows():
                rel = {
                    "fromTable": str(row.get("FromTableName", row.get("fromTableName", ""))),
                    "fromColumn": str(row.get("FromColumnName", row.get("fromColumnName", ""))),
                    "toTable": str(row.get("ToTableName", row.get("toTableName", ""))),
                    "toColumn": str(row.get("ToColumnName", row.get("toColumnName", ""))),
                    "isActive": bool(row.get("IsActive", row.get("isActive", True))),
                    "cardinality": str(row.get("Cardinality", row.get("cardinality", ""))),
                    "crossFiltering": str(row.get("CrossFilteringBehavior", row.get("crossFilteringBehavior", ""))),
                }
                if rel["fromTable"] and rel["toTable"]:
                    relationships.append(rel)
            logger.info(f"pbixray: extracted {len(relationships)} relationships")
    except Exception as e:
        logger.warning(f"pbixray: could not extract relationships: {e}")

    if not measures_by_table and not columns_by_table:
        logger.warning("pbixray: no measures or columns extracted")
        return False

    # Merge all table names
    all_tables = set(measures_by_table.keys()) | set(columns_by_table.keys())

    for table_name in sorted(all_tables):
        tmdl_content = _build_tmdl_file(
            table_name,
            measures_by_table.get(table_name, []),
            columns_by_table.get(table_name, []),
        )
        # Sanitize filename — TMDL files use the table name as filename
        safe_name = sanitize_filename(table_name)
        tmdl_path = tables_dir / f"{safe_name}.tmdl"
        tmdl_path.write_text(tmdl_content, encoding="utf-8")

    # Write minimal model.tmdl stub
    (model_dir / "model.tmdl").write_text(
        "model Model\n\tculture: en-US\n", encoding="utf-8"
    )

    # Write minimal database.tmdl stub
    (model_dir / "database.tmdl").write_text(
        "database Database\n", encoding="utf-8"
    )

    # Write .source marker so downstream code knows types are unreliable
    (model_dir / ".source").write_text("pbixray", encoding="utf-8")

    # Write relationships.json if any were extracted
    if relationships:
        import json as _json
        (model_dir / "relationships.json").write_text(
            _json.dumps(relationships, indent=2, ensure_ascii=False),
            encoding="utf-8",
        )

    logger.info(
        f"pbixray: wrote {len(all_tables)} TMDL files to {tables_dir}"
    )

    # Log known pbixray limitations
    logger.warning(
        "pbixray limitations: column data types are ALL reported as 'string' "
        "(unreliable). Calculated columns/tables may be missing. "
        "The pipeline will use CONVERT() for numeric aggregations to work around this."
    )

    return True


def _build_tmdl_file(
    table_name: str,
    measures: list[tuple[str, str]],
    columns: list[tuple[str, str]],
) -> str:
    """Build a synthetic TMDL file content matching tmdl_parser.py regex patterns.

    Format:
        table <TableName>
        \tmeasure '<MeasureName>' = <expression>
        \tcolumn <ColumnName>
        \t\tdataType: <type>
    """
    # Quote table name if it contains spaces
    quoted_table = f"'{table_name}'" if " " in table_name else table_name
    lines = [f"table {quoted_table}"]

    for measure_name, expression in measures:
        # Quote measure name (tmdl_parser.py expects optional quotes)
        quoted_measure = (
            f"'{measure_name}'" if " " in measure_name else measure_name
        )
        # Handle multi-line expressions: indent continuation lines
        expr_lines = expression.strip().split("\n")
        if len(expr_lines) == 1:
            lines.append(f"\tmeasure {quoted_measure} = {expr_lines[0].strip()}")
        else:
            lines.append(f"\tmeasure {quoted_measure} =")
            lines.append("\t\t```")
            for el in expr_lines:
                lines.append(f"\t\t{el}")
            lines.append("\t\t```")
        # Add a lineageTag placeholder so tmdl_parser stops correctly
        lines.append(f"\t\tlineageTag: auto-generated")
        lines.append("")

    for col_name, data_type in columns:
        quoted_col = f"'{col_name}'" if " " in col_name else col_name
        lines.append(f"\tcolumn {quoted_col}")
        lines.append(f"\t\tdataType: {data_type}")
        lines.append(f"\t\tlineageTag: auto-generated")
        lines.append(f"\t\tsummarizeBy: none")
        lines.append(f"\t\tsourceColumn: {col_name}")
        lines.append("")

    return "\n".join(lines) + "\n"


# ---------------------------------------------------------------------------
# Main extraction orchestrator
# ---------------------------------------------------------------------------


def extract_pbix(
    pbix_path: str,
    output_dir: str = "data/",
    model_root: Optional[str] = None,
) -> PbixExtractResult:
    """Extract a .pbix file into PBIP folder structure.

    Args:
        pbix_path: Path to the .pbix file
        output_dir: Output directory (default: data/)
        model_root: Optional path to existing PBIP semantic model (skips pbixray)

    Returns:
        PbixExtractResult with paths and counts
    """
    pbix_path = str(Path(pbix_path).resolve())
    if not os.path.isfile(pbix_path):
        raise FileNotFoundError(f"PBIX file not found: {pbix_path}")
    if not zipfile.is_zipfile(pbix_path):
        raise ValueError(f"Not a valid ZIP/PBIX file: {pbix_path}")

    # Derive report name from filename (without extension)
    report_name = Path(pbix_path).stem
    output_base = Path(output_dir).resolve()

    # PBIP-style output paths
    report_dir = output_base / f"{report_name}.Report" / "definition"
    report_dir.mkdir(parents=True, exist_ok=True)

    logger.info(f"Extracting {pbix_path} → {report_dir}")

    # Read the monolithic Layout JSON
    layout = read_layout_json(pbix_path)

    sections = layout.get("sections", [])
    if not sections:
        logger.warning("No sections (pages) found in Layout JSON")

    # ---- Step 1: report.json ----
    report_json = build_report_json(layout)

    # Add report-level filters if present
    report_filters = build_report_filters(layout)
    if report_filters:
        report_json["filterConfig"] = {"filters": report_filters}

    _write_json(report_dir / "report.json", report_json)

    # ---- Step 2: pages/pages.json ----
    pages_dir = report_dir / "pages"
    pages_dir.mkdir(exist_ok=True)
    pages_json = build_pages_json(sections)
    _write_json(pages_dir / "pages.json", pages_json)

    # ---- Step 3: Per-page extraction ----
    total_containers = 0
    data_visuals = 0
    for section in sections:
        section_name = section.get("name", "")
        if not section_name:
            continue

        page_dir = pages_dir / section_name
        page_dir.mkdir(exist_ok=True)

        # page.json
        page_json = build_page_json(section)
        _write_json(page_dir / "page.json", page_json)

        # Visuals
        visuals = extract_visuals_from_section(section)
        if visuals:
            visuals_dir = page_dir / "visuals"
            visuals_dir.mkdir(exist_ok=True)

            for visual_id, visual_json in visuals:
                visual_dir = visuals_dir / visual_id
                visual_dir.mkdir(exist_ok=True)
                _write_json(visual_dir / "visual.json", visual_json)
                total_containers += 1
                # Count data visuals (queryState + not a decorative type)
                visual_block = visual_json.get("visual", {})
                vtype = visual_block.get("visualType", "")
                has_query = bool(visual_block.get("query", {}).get("queryState"))
                is_decorative = vtype in (
                    "textbox", "image", "shape", "actionButton", "group",
                )
                if has_query and not is_decorative:
                    data_visuals += 1

    logger.info(
        f"Extracted {len(sections)} pages, {total_containers} visual containers "
        f"({data_visuals} data visuals)"
    )

    # ---- Step 4: Bookmarks ----
    bookmarks = extract_bookmarks(layout)
    bookmark_count = 0
    if bookmarks:
        bookmarks_dir = report_dir / "bookmarks"
        bookmarks_dir.mkdir(exist_ok=True)

        # bookmarks.json index
        bm_index = build_bookmarks_index(bookmarks)
        _write_json(bookmarks_dir / "bookmarks.json", bm_index)

        # Individual bookmark files
        for bm in bookmarks:
            bm_name = bm.get("name", "")
            if not bm_name:
                continue
            bm_file = build_bookmark_file(bm)
            safe_bm_name = sanitize_filename(bm_name)
            _write_json(bookmarks_dir / f"{safe_bm_name}.bookmark.json", bm_file)
            bookmark_count += 1

        logger.info(f"Extracted {bookmark_count} bookmarks")

    # ---- Step 5: Semantic model ----
    result = PbixExtractResult(
        report_root=str(report_dir),
        model_root=None,
        report_name=report_name,
        page_count=len(sections),
        visual_container_count=total_containers,
        data_visual_count=data_visuals,
        bookmark_count=bookmark_count,
        semantic_model_source="none",
    )

    if model_root:
        # User provided an existing semantic model path
        result.model_root = str(Path(model_root).resolve())
        result.semantic_model_source = "user-provided"
        logger.info(f"Using user-provided semantic model: {result.model_root}")
    else:
        # Try pbixray extraction
        model_dir = output_base / f"{report_name}.SemanticModel" / "definition"
        if extract_semantic_model_pbixray(pbix_path, model_dir):
            result.model_root = str(model_dir)
            result.semantic_model_source = "pbixray"
        else:
            if HAS_PBIXRAY:
                logger.warning(
                    "pbixray could not extract the semantic model from this .pbix file. "
                    "Report structure was extracted but measure formulas will be missing."
                )
            elif PBIXRAY_ERROR and "C compiler" in PBIXRAY_ERROR:
                logger.warning(
                    f"pbixray could not load: {PBIXRAY_ERROR}\n"
                    "Report structure was extracted but measure formulas will be missing."
                )
            else:
                logger.warning(
                    "pbixray is not installed. Report structure was extracted but "
                    "measure formulas will be missing. Install with: pip install pbixray"
                )

    return result


def _write_json(path: Path, data: dict) -> None:
    """Write a JSON file with consistent formatting."""
    path.write_text(
        json.dumps(data, indent=2, ensure_ascii=False), encoding="utf-8-sig"
    )


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------


def main():
    parser = argparse.ArgumentParser(
        description="Convert a .pbix file to PBIP folder structure"
    )
    parser.add_argument("pbix_path", help="Path to the .pbix file")
    parser.add_argument(
        "--output",
        default="data/",
        help="Output directory (default: data/)",
    )
    parser.add_argument(
        "--model-root",
        default=None,
        help="Path to existing PBIP semantic model definition folder (skips pbixray)",
    )
    args = parser.parse_args()

    result = extract_pbix(args.pbix_path, args.output, args.model_root)

    print(f"\nExtraction complete: {result.report_name}")
    print(f"  Report root:  {result.report_root}")
    print(f"  Model root:   {result.model_root or '(not extracted)'}")
    print(f"  Pages:        {result.page_count}")
    print(f"  Data visuals: {result.data_visual_count}")
    print(f"  Visual containers (total): {result.visual_container_count}")
    print(f"  Bookmarks:    {result.bookmark_count}")
    print(f"  Model source: {result.semantic_model_source}")


if __name__ == "__main__":
    main()
