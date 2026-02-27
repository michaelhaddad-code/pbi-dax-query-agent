"""
Skill 0: PBIX to PBIP Converter

Extracts a .pbix ZIP archive into the PBIP folder structure consumed by
extract_metadata.py. Report structure (pages, visuals, filters, bookmarks)
is extracted with pure Python. Semantic model extraction (tables, columns,
measures, relationships, hierarchies, variations, partitions, RLS roles)
uses pbixray's PbixUnpacker + SQLiteHandler for full metadata access.

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
import uuid
import zipfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Optional

import pandas as pd

logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")
logger = logging.getLogger(__name__)

# Try importing pbixray internals at module level — optional dependency
HAS_PBIXRAY = False
PBIXRAY_ERROR = ""
try:
    from pbixray.pbix_unpacker import PbixUnpacker
    from pbixray.utils import get_data_slice
    from pbixray.meta.sqlite_handler import SQLiteHandler

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
    semantic_model_source: str = "none"  # "pbixray-sqlite" | "user-provided" | "none"


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

# TOM ExplicitDataType codes → TMDL dataType strings
DATATYPE_MAP = {
    2: "string", 6: "int64", 8: "double", 9: "dateTime",
    10: "decimal", 11: "boolean", 17: "binary",
}

# TOM SummarizeBy codes → TMDL summarizeBy strings (None = omit)
SUMMARIZE_BY_MAP = {
    1: None, 2: "none", 3: "sum", 4: "min", 5: "max",
    6: "count", 7: "average", 8: "distinctCount",
}

# Column.Type: 1=imported, 2=calculated, 3=RowNumber(skip), 4=calcTableCol
# CrossFilteringBehavior: 1=oneDirection(default/omit), 2=bothDirections
# JoinOnDateBehavior: 1=default(omit), 2=datePartOnly
# Partition.Type: 2=calculated, 4=m

# TOM Annotation ObjectType codes
ANNOT_TABLE = 3
ANNOT_COLUMN = 4
ANNOT_MEASURE = 9
ANNOT_HIERARCHY = 10


# ---------------------------------------------------------------------------
# SQLite query helpers
# ---------------------------------------------------------------------------


def _query(handler, sql: str) -> tuple:
    """Execute SQL via SQLiteHandler. Returns (DataFrame, success_bool).

    success=True means the query executed (even if zero rows).
    success=False means the query errored (missing table/column).
    We distinguish by checking if the returned DataFrame has columns.
    """
    df = handler.execute_query(sql)
    return df, len(df.columns) > 0


def _safe_int(val, default=None):
    """Safely convert a value to int, returning default on failure."""
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return default
    try:
        return int(val)
    except (ValueError, TypeError):
        return default


def _safe_str(val, default=""):
    """Return string if non-null, else default."""
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return default
    return str(val)


def _safe_bool(val) -> bool:
    """Return True if val is truthy (handles 0/1, True/False, NaN)."""
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return False
    return bool(int(val))


def query_tables(handler) -> pd.DataFrame:
    """Query user tables from the metadata SQLite.

    Filters by SystemFlags: 0=regular tables, 2=date/calculated tables.
    Excludes SystemFlags=1 (internal RowNumber/hierarchy storage tables).
    """
    sql = """
        SELECT ID, Name, IsHidden, ShowAsVariationsOnly, LineageTag, SystemFlags
        FROM [Table]
        WHERE SystemFlags IN (0, 2)
    """
    df, ok = _query(handler, sql)
    if not ok:
        # Fallback: only essential columns (still filter SystemFlags)
        sql = "SELECT ID, Name FROM [Table] WHERE SystemFlags IN (0, 2)"
        df, ok = _query(handler, sql)
    if not ok:
        # Last resort: no SystemFlags filter (column might not exist)
        sql = "SELECT ID, Name FROM [Table]"
        df, _ = _query(handler, sql)
    return df


def query_columns(handler, table_ids: list) -> pd.DataFrame:
    """Query all columns (excluding RowNumber type 3) with SortByColumn resolution.

    Uses progressive fallback: full query → without IsNameInferred/SortByColumn →
    without SourceColumn/FormatString → minimal (ID, TableID, Name, Type, DataType).
    """
    if not table_ids:
        return pd.DataFrame()
    ids_str = ",".join(str(i) for i in table_ids)

    # Tier 1: Full query with SortByColumn self-join + IsNameInferred
    sql = f"""
        SELECT c.ID, c.TableID, c.ExplicitName, c.InferredName, c.Type,
               c.ExplicitDataType, c.SourceColumn, c.Expression, c.FormatString,
               c.IsHidden, c.SummarizeBy, c.DataCategory, c.LineageTag,
               c.IsNameInferred,
               sc.ExplicitName AS SortByColumnName
        FROM [Column] c
        LEFT JOIN [Column] sc ON c.SortByColumnID = sc.ID
        WHERE c.Type != 3 AND c.TableID IN ({ids_str})
    """
    df, ok = _query(handler, sql)
    if ok:
        return df

    # Tier 2: Without IsNameInferred/InferredName and SortByColumn self-join
    # (these columns don't exist in all PBIX SQLite schemas)
    sql = f"""
        SELECT c.ID, c.TableID, c.ExplicitName, c.Type,
               c.ExplicitDataType, c.SourceColumn, c.Expression, c.FormatString,
               c.IsHidden, c.SummarizeBy, c.DataCategory, c.LineageTag
        FROM [Column] c
        WHERE c.Type != 3 AND c.TableID IN ({ids_str})
    """
    df, ok = _query(handler, sql)
    if ok:
        logger.info("Column query: using Tier 2 (no SortByColumn/IsNameInferred)")
        return df

    # Tier 3: Core columns only (SourceColumn/FormatString may also be missing)
    sql = f"""
        SELECT c.ID, c.TableID, c.ExplicitName, c.Type,
               c.ExplicitDataType, c.Expression, c.IsHidden, c.LineageTag
        FROM [Column] c
        WHERE c.Type != 3 AND c.TableID IN ({ids_str})
    """
    df, ok = _query(handler, sql)
    if ok:
        logger.info("Column query: using Tier 3 (no SourceColumn/FormatString)")
        return df

    # Tier 4: Minimal fallback
    sql = f"""
        SELECT c.ID, c.TableID, c.ExplicitName, c.Type,
               c.ExplicitDataType, c.Expression
        FROM [Column] c
        WHERE c.Type != 3 AND c.TableID IN ({ids_str})
    """
    df, _ = _query(handler, sql)
    logger.info("Column query: using Tier 4 (minimal)")
    return df


def query_measures(handler, table_ids: list) -> pd.DataFrame:
    """Query all DAX measures."""
    if not table_ids:
        return pd.DataFrame()
    ids_str = ",".join(str(i) for i in table_ids)
    sql = f"""
        SELECT m.ID, m.TableID, m.Name, m.Expression, m.FormatString,
               m.IsHidden, m.DisplayFolder, m.Description, m.LineageTag
        FROM [Measure] m
        WHERE m.TableID IN ({ids_str})
    """
    df, ok = _query(handler, sql)
    if not ok:
        sql = f"""
            SELECT m.ID, m.TableID, m.Name, m.Expression
            FROM [Measure] m
            WHERE m.TableID IN ({ids_str})
        """
        df, _ = _query(handler, sql)
    return df


def query_hierarchies_and_levels(handler, table_ids: list) -> pd.DataFrame:
    """Query Hierarchy + Level + Column join. May fail if tables don't exist."""
    if not table_ids:
        return pd.DataFrame()
    ids_str = ",".join(str(i) for i in table_ids)
    sql = f"""
        SELECT h.ID AS HierarchyID, h.Name AS HierarchyName, h.TableID,
               h.IsHidden AS HierarchyIsHidden, h.LineageTag AS HierarchyLineageTag,
               l.ID AS LevelID, l.Ordinal AS LevelOrdinal, l.Name AS LevelName,
               l.LineageTag AS LevelLineageTag,
               c.ExplicitName AS LevelColumnName
        FROM [Hierarchy] h
        LEFT JOIN [Level] l ON l.HierarchyID = h.ID
        LEFT JOIN [Column] c ON l.ColumnID = c.ID
        WHERE h.TableID IN ({ids_str})
        ORDER BY h.ID, l.Ordinal
    """
    df, ok = _query(handler, sql)
    return df if ok else pd.DataFrame()


def query_variations(handler, table_ids: list) -> pd.DataFrame:
    """Query Variation + Column + Relationship + Hierarchy join.

    May fail if Variation/Hierarchy tables don't exist in this SQLite DB.
    """
    if not table_ids:
        return pd.DataFrame()
    ids_str = ",".join(str(i) for i in table_ids)
    sql = f"""
        SELECT v.ID, v.ColumnID, v.Name AS VariationName, v.IsDefault,
               c.ExplicitName AS OwnerColumnName, c.TableID,
               r.Name AS RelationshipName,
               h.Name AS DefaultHierarchyName,
               ht.Name AS HierarchyTableName
        FROM [Variation] v
        JOIN [Column] c ON v.ColumnID = c.ID
        LEFT JOIN [Relationship] r ON v.RelationshipID = r.ID
        LEFT JOIN [Hierarchy] h ON v.DefaultHierarchyID = h.ID
        LEFT JOIN [Table] ht ON h.TableID = ht.ID
        WHERE c.TableID IN ({ids_str})
    """
    df, ok = _query(handler, sql)
    return df if ok else pd.DataFrame()


def query_relationships(handler) -> pd.DataFrame:
    """Query all relationships with table/column name resolution."""
    sql = """
        SELECT r.ID, r.Name AS RelName,
               ft.Name AS FromTableName, fc.ExplicitName AS FromColumnName,
               tt.Name AS ToTableName, tc.ExplicitName AS ToColumnName,
               r.IsActive, r.CrossFilteringBehavior, r.JoinOnDateBehavior
        FROM [Relationship] r
        LEFT JOIN [Table] ft ON r.FromTableID = ft.ID
        LEFT JOIN [Column] fc ON r.FromColumnID = fc.ID
        LEFT JOIN [Table] tt ON r.ToTableID = tt.ID
        LEFT JOIN [Column] tc ON r.ToColumnID = tc.ID
    """
    df, ok = _query(handler, sql)
    if not ok:
        # Fallback: without JoinOnDateBehavior and Relationship.Name
        sql = """
            SELECT r.ID,
                   ft.Name AS FromTableName, fc.ExplicitName AS FromColumnName,
                   tt.Name AS ToTableName, tc.ExplicitName AS ToColumnName,
                   r.IsActive, r.CrossFilteringBehavior
            FROM [Relationship] r
            LEFT JOIN [Table] ft ON r.FromTableID = ft.ID
            LEFT JOIN [Column] fc ON r.FromColumnID = fc.ID
            LEFT JOIN [Table] tt ON r.ToTableID = tt.ID
            LEFT JOIN [Column] tc ON r.ToColumnID = tc.ID
        """
        df, _ = _query(handler, sql)
    return df


def query_partitions(handler, table_ids: list) -> pd.DataFrame:
    """Query M (type 4) and calculated (type 2) partitions."""
    if not table_ids:
        return pd.DataFrame()
    ids_str = ",".join(str(i) for i in table_ids)
    sql = f"""
        SELECT p.ID, p.TableID, p.Name AS PartitionName, p.Type,
               p.QueryDefinition, p.Mode
        FROM [partition] p
        WHERE p.TableID IN ({ids_str}) AND p.Type IN (2, 4)
    """
    df, ok = _query(handler, sql)
    if not ok:
        # Fallback: without Mode and Name
        sql = f"""
            SELECT p.ID, p.TableID, p.Type, p.QueryDefinition
            FROM [partition] p
            WHERE p.TableID IN ({ids_str}) AND p.Type IN (2, 4)
        """
        df, _ = _query(handler, sql)
    return df


def query_rls_roles(handler) -> pd.DataFrame:
    """Query RLS role definitions (TablePermission + Role + Table)."""
    sql = """
        SELECT r.Name AS RoleName, r.Description AS RoleDescription,
               t.Name AS TableName, tp.FilterExpression
        FROM [TablePermission] tp
        JOIN [Role] r ON tp.RoleID = r.ID
        JOIN [Table] t ON tp.TableID = t.ID
    """
    df, _ = _query(handler, sql)
    return df


def query_annotations(handler) -> pd.DataFrame:
    """Query all annotations with ObjectType and ObjectID for mapping."""
    sql = "SELECT ObjectType, ObjectID, Name, Value FROM [Annotation]"
    df, ok = _query(handler, sql)
    if not ok:
        # Fallback: model-level only (ObjectType=1)
        sql = "SELECT Name, Value FROM [Annotation] WHERE ObjectType = 1"
        df, _ = _query(handler, sql)
    return df


# ---------------------------------------------------------------------------
# TMDL generation helpers
# ---------------------------------------------------------------------------


def tmdl_quote(name: str) -> str:
    """Single-quote names containing spaces; escape embedded single quotes."""
    if " " in name or "'" in name:
        return "'" + name.replace("'", "''") + "'"
    return name


def _tmdl_col_ref(table_name: str, col_name: str) -> str:
    """Format a Table.Column reference for TMDL (fromColumn/toColumn)."""
    return f"{tmdl_quote(table_name)}.{tmdl_quote(col_name)}"


def _emit_measure(m: pd.Series, annotations: list) -> list:
    """Generate TMDL lines for a single measure block."""
    lines = []
    name = _safe_str(m.get("Name", ""))
    expression = _safe_str(m.get("Expression", ""))

    # Declaration line with expression
    expr_lines = expression.strip().split("\n") if expression.strip() else [""]
    if len(expr_lines) == 1:
        lines.append(f"\tmeasure {tmdl_quote(name)} = {expr_lines[0].strip()}")
    else:
        lines.append(f"\tmeasure {tmdl_quote(name)} =")
        lines.append("\t\t```")
        for el in expr_lines:
            lines.append(f"\t\t{el}")
        lines.append("\t\t```")

    # Properties
    fmt = _safe_str(m.get("FormatString"))
    if fmt:
        lines.append(f"\t\tformatString: {fmt}")

    tag = _safe_str(m.get("LineageTag"))
    if tag:
        lines.append(f"\t\tlineageTag: {tag}")

    if _safe_bool(m.get("IsHidden")):
        lines.append("\t\tisHidden")

    display_folder = _safe_str(m.get("DisplayFolder"))
    if display_folder:
        lines.append(f"\t\tdisplayFolder: {display_folder}")

    # Annotations
    for ann_name, ann_val in annotations:
        lines.append("")
        lines.append(f"\t\tannotation {ann_name} = {ann_val}")

    lines.append("")  # blank line after measure
    return lines


def _emit_column(c: pd.Series, col_variations: pd.DataFrame,
                 annotations: list) -> list:
    """Generate TMDL lines for a single column block."""
    lines = []
    col_type = _safe_int(c.get("Type"), 1)

    # Determine display name
    is_name_inferred = _safe_bool(c.get("IsNameInferred"))
    if is_name_inferred and _safe_str(c.get("InferredName")):
        name = _safe_str(c["InferredName"])
    else:
        name = _safe_str(c.get("ExplicitName", ""))

    # Declaration line
    if col_type == 2:
        # Calculated column: column Name = expression
        expression = _safe_str(c.get("Expression", ""))
        expr_lines = expression.strip().split("\n") if expression.strip() else [""]
        if len(expr_lines) == 1:
            lines.append(f"\tcolumn {tmdl_quote(name)} = {expr_lines[0].strip()}")
        else:
            lines.append(f"\tcolumn {tmdl_quote(name)} =")
            lines.append("\t\t```")
            for el in expr_lines:
                lines.append(f"\t\t{el}")
            lines.append("\t\t```")
    else:
        # Imported (1) or calc table col (4)
        lines.append(f"\tcolumn {tmdl_quote(name)}")

    # Properties — order matches PBI Desktop output
    if col_type in (1, 4):
        dtype_code = _safe_int(c.get("ExplicitDataType"))
        if dtype_code is not None:
            dtype_str = DATATYPE_MAP.get(dtype_code, "string")
            lines.append(f"\t\tdataType: {dtype_str}")

    if _safe_bool(c.get("IsHidden")):
        lines.append("\t\tisHidden")

    fmt = _safe_str(c.get("FormatString"))
    if fmt:
        lines.append(f"\t\tformatString: {fmt}")

    tag = _safe_str(c.get("LineageTag"))
    if tag:
        lines.append(f"\t\tlineageTag: {tag}")

    data_cat = _safe_str(c.get("DataCategory"))
    if data_cat:
        lines.append(f"\t\tdataCategory: {data_cat}")

    # summarizeBy
    sb_code = _safe_int(c.get("SummarizeBy"))
    if sb_code is not None:
        sb_str = SUMMARIZE_BY_MAP.get(sb_code)
        if sb_str:
            lines.append(f"\t\tsummarizeBy: {sb_str}")

    if is_name_inferred:
        lines.append("\t\tisNameInferred")

    # sourceColumn (imported=real name, calcTableCol=[Name], calculated=none)
    if col_type == 1:
        src = _safe_str(c.get("SourceColumn"))
        if src:
            lines.append(f"\t\tsourceColumn: {src}")
        else:
            lines.append(f"\t\tsourceColumn: {name}")
    elif col_type == 4:
        lines.append(f"\t\tsourceColumn: [{name}]")

    # sortByColumn
    sort_by = _safe_str(c.get("SortByColumnName"))
    if sort_by:
        lines.append(f"\t\tsortByColumn: {sort_by}")

    # Variation sub-blocks (nested under column at 2-tab level)
    if not col_variations.empty:
        for _, v in col_variations.iterrows():
            lines.append("")
            var_name = _safe_str(v.get("VariationName", "Variation"))
            lines.append(f"\t\tvariation {var_name}")
            if _safe_bool(v.get("IsDefault")):
                lines.append("\t\t\tisDefault")
            rel_name = _safe_str(v.get("RelationshipName"))
            if rel_name:
                lines.append(f"\t\t\trelationship: {rel_name}")
            hier_name = _safe_str(v.get("DefaultHierarchyName"))
            hier_table = _safe_str(v.get("HierarchyTableName"))
            if hier_name and hier_table:
                lines.append(
                    f"\t\t\tdefaultHierarchy: "
                    f"{tmdl_quote(hier_table)}.{tmdl_quote(hier_name)}"
                )

    # Annotations
    for ann_name, ann_val in annotations:
        lines.append("")
        lines.append(f"\t\tannotation {ann_name} = {ann_val}")

    lines.append("")  # blank line after column
    return lines


def _emit_hierarchy(hier_rows: pd.DataFrame, annotations: list) -> list:
    """Generate TMDL lines for a single hierarchy block.

    hier_rows: all rows for one hierarchy (one per level), sorted by Ordinal.
    """
    lines = []
    first = hier_rows.iloc[0]
    hier_name = _safe_str(first.get("HierarchyName", ""))
    lines.append(f"\thierarchy {tmdl_quote(hier_name)}")

    tag = _safe_str(first.get("HierarchyLineageTag"))
    if tag:
        lines.append(f"\t\tlineageTag: {tag}")

    # Levels ordered by ordinal
    for _, lvl in hier_rows.iterrows():
        level_name = _safe_str(lvl.get("LevelName", ""))
        if not level_name:
            continue
        lines.append("")
        lines.append(f"\t\tlevel {tmdl_quote(level_name)}")
        lvl_tag = _safe_str(lvl.get("LevelLineageTag"))
        if lvl_tag:
            lines.append(f"\t\t\tlineageTag: {lvl_tag}")
        col_name = _safe_str(lvl.get("LevelColumnName"))
        if col_name:
            lines.append(f"\t\t\tcolumn: {col_name}")

    # Annotations (e.g., TemplateId = DateHierarchy)
    for ann_name, ann_val in annotations:
        lines.append("")
        lines.append(f"\t\tannotation {ann_name} = {ann_val}")

    lines.append("")  # blank line after hierarchy
    return lines


def _emit_partition(p: pd.Series) -> list:
    """Generate TMDL lines for a partition block."""
    lines = []
    part_type = _safe_int(p.get("Type"), 4)
    part_name = _safe_str(p.get("PartitionName"))
    query_def = _safe_str(p.get("QueryDefinition", ""))

    # Partition type keyword
    type_kw = "calculated" if part_type == 2 else "m"

    # Generate a name if missing
    if not part_name:
        part_name = str(uuid.uuid4())

    lines.append(f"\tpartition {part_name} = {type_kw}")

    # Mode (default to import)
    mode_val = _safe_int(p.get("Mode"), 0)
    mode_str = "directQuery" if mode_val == 1 else "import"
    lines.append(f"\t\tmode: {mode_str}")

    # Source expression
    if query_def:
        qd_lines = query_def.split("\n")
        if part_type == 2 or len(qd_lines) == 1:
            # Single-line (calculated DAX or short M)
            lines.append(f"\t\tsource = {qd_lines[0].strip()}")
            # Remaining lines (rare for calculated, but handle gracefully)
            for extra in qd_lines[1:]:
                lines.append(f"\t\t\t{extra}")
        else:
            # Multi-line M expression — indent at 3 tabs
            lines.append("\t\tsource =")
            for ml in qd_lines:
                lines.append(f"\t\t\t{ml}")

    lines.append("")  # blank line after partition
    return lines


def generate_table_tmdl(
    table_name: str,
    table_props: dict,
    measures_df: pd.DataFrame,
    columns_df: pd.DataFrame,
    hier_levels_df: pd.DataFrame,
    variations_df: pd.DataFrame,
    partitions_df: pd.DataFrame,
    annotations_map: dict,
) -> str:
    """Assemble complete TMDL content for one table.

    annotations_map has keys: 'table', 'columns', 'measures', 'hierarchies'
    Each maps object ID → list of (name, value) tuples.
    """
    lines = []

    # Table header
    lines.append(f"table {tmdl_quote(table_name)}")

    # Table-level flags
    if table_props.get("IsHidden"):
        lines.append("\tisHidden")
    if table_props.get("IsPrivate"):
        lines.append("\tisPrivate")
    if table_props.get("ShowAsVariationsOnly"):
        lines.append("\tshowAsVariationsOnly")

    tag = table_props.get("LineageTag", "")
    if tag:
        lines.append(f"\tlineageTag: {tag}")

    lines.append("")  # blank line after table header

    # Measures
    measure_annots = annotations_map.get("measures", {})
    for _, m in measures_df.iterrows():
        m_id = _safe_int(m.get("ID"))
        annots = measure_annots.get(m_id, [])
        lines.extend(_emit_measure(m, annots))

    # Columns
    col_annots = annotations_map.get("columns", {})
    for _, c in columns_df.iterrows():
        c_id = _safe_int(c.get("ID"))
        # Get variations for this column
        col_vars = pd.DataFrame()
        if not variations_df.empty and "ColumnID" in variations_df.columns:
            col_vars = variations_df[variations_df["ColumnID"] == c_id]
        annots = col_annots.get(c_id, [])
        lines.extend(_emit_column(c, col_vars, annots))

    # Hierarchies
    hier_annots = annotations_map.get("hierarchies", {})
    if not hier_levels_df.empty and "HierarchyID" in hier_levels_df.columns:
        for hier_id in hier_levels_df["HierarchyID"].unique():
            hier_data = hier_levels_df[hier_levels_df["HierarchyID"] == hier_id]
            annots = hier_annots.get(_safe_int(hier_id), [])
            lines.extend(_emit_hierarchy(hier_data, annots))

    # Partitions
    for _, p in partitions_df.iterrows():
        lines.extend(_emit_partition(p))

    # Table-level annotations
    table_annots = annotations_map.get("table", [])
    for ann_name, ann_val in table_annots:
        lines.append(f"\tannotation {ann_name} = {ann_val}")
        lines.append("")

    return "\n".join(lines) + "\n"


def generate_relationships_tmdl(rel_df: pd.DataFrame) -> str:
    """Generate relationships.tmdl from a DataFrame of relationships."""
    lines = []
    for _, r in rel_df.iterrows():
        # Relationship name (UUID) — use Name if available, else generate one
        rel_name = _safe_str(r.get("RelName"))
        if not rel_name:
            rel_name = str(r.get("ID", uuid.uuid4()))
        lines.append(f"relationship {rel_name}")

        # Optional: joinOnDateBehavior (only emit datePartOnly)
        jodb = _safe_int(r.get("JoinOnDateBehavior"))
        if jodb == 2:
            lines.append("\tjoinOnDateBehavior: datePartOnly")

        # Optional: crossFilteringBehavior (only emit bothDirections)
        cfb = _safe_int(r.get("CrossFilteringBehavior"))
        if cfb == 2:
            lines.append("\tcrossFilteringBehavior: bothDirections")

        # Optional: isActive (only emit when false)
        is_active = r.get("IsActive")
        if is_active is not None and not _safe_bool(is_active):
            lines.append("\tisActive: false")

        # fromColumn and toColumn
        from_table = _safe_str(r.get("FromTableName"))
        from_col = _safe_str(r.get("FromColumnName"))
        to_table = _safe_str(r.get("ToTableName"))
        to_col = _safe_str(r.get("ToColumnName"))
        if from_table and from_col:
            lines.append(f"\tfromColumn: {_tmdl_col_ref(from_table, from_col)}")
        if to_table and to_col:
            lines.append(f"\ttoColumn: {_tmdl_col_ref(to_table, to_col)}")

        lines.append("")  # blank line between relationships

    return "\n".join(lines) + "\n" if lines else ""


def generate_role_tmdl(role_name: str, permissions_df: pd.DataFrame) -> str:
    """Generate a role TMDL file from role name + table permissions."""
    lines = []
    lines.append(f"role {tmdl_quote(role_name)}")

    # Optional: role description (from first row)
    if "RoleDescription" in permissions_df.columns:
        desc = _safe_str(permissions_df.iloc[0].get("RoleDescription"))
        if desc:
            lines.append(f"\tdescription: {desc}")

    lines.append("")

    for _, perm in permissions_df.iterrows():
        table_name = _safe_str(perm.get("TableName", ""))
        filter_expr = _safe_str(perm.get("FilterExpression", ""))
        if table_name:
            lines.append(f"\ttablePermission {tmdl_quote(table_name)}")
            if filter_expr:
                filter_lines = filter_expr.strip().split("\n")
                if len(filter_lines) == 1:
                    lines.append(
                        f"\t\tfilterExpression = {filter_lines[0].strip()}"
                    )
                else:
                    lines.append("\t\tfilterExpression =")
                    lines.append("\t\t\t```")
                    for fl in filter_lines:
                        lines.append(f"\t\t\t{fl}")
                    lines.append("\t\t\t```")
            lines.append("")

    return "\n".join(lines) + "\n"


# ---------------------------------------------------------------------------
# Main semantic model extraction from SQLite
# ---------------------------------------------------------------------------


def extract_semantic_model_from_sqlite(pbix_path: str, model_dir: Path) -> bool:
    """Extract full semantic model from .pbix via pbixray's internal SQLite.

    Uses PbixUnpacker + SQLiteHandler for single-decompression-pass extraction
    with full access to every TOM table (columns, measures, relationships,
    hierarchies, variations, partitions, RLS roles, annotations).

    Returns True if extraction succeeded, False otherwise.
    """
    if not HAS_PBIXRAY:
        return False

    try:
        # Step 1: Decompress and open SQLite
        unpacker = PbixUnpacker(pbix_path)
        data_model = unpacker.data_model
        sqlite_bytes = get_data_slice(data_model, "metadata.sqlitedb")
        handler = SQLiteHandler(sqlite_bytes)
    except Exception as e:
        logger.warning(f"pbixray: failed to open SQLite from {pbix_path}: {e}")
        return False

    try:
        # Step 2: Query all metadata
        tables_df = query_tables(handler)
        if tables_df.empty:
            logger.warning("pbixray: no tables found in metadata")
            return False

        table_ids = tables_df["ID"].tolist()

        # Build table name lookup
        table_id_to_name = dict(zip(tables_df["ID"], tables_df["Name"]))

        columns_df = query_columns(handler, table_ids)
        measures_df = query_measures(handler, table_ids)
        hier_levels_df = query_hierarchies_and_levels(handler, table_ids)
        variations_df = query_variations(handler, table_ids)
        relationships_df = query_relationships(handler)
        partitions_df = query_partitions(handler, table_ids)
        rls_df = query_rls_roles(handler)
        annotations_df = query_annotations(handler)

        logger.info(
            f"pbixray: queried {len(tables_df)} tables, "
            f"{len(columns_df)} columns, {len(measures_df)} measures, "
            f"{len(relationships_df)} relationships"
        )

        # Step 3: Build annotation lookup {ObjectType: {ObjectID: [(name, val)]}}
        annot_lookup: dict[int, dict[int, list]] = {}
        if not annotations_df.empty and "ObjectType" in annotations_df.columns:
            for _, ann in annotations_df.iterrows():
                otype = _safe_int(ann.get("ObjectType"))
                oid = _safe_int(ann.get("ObjectID"))
                if otype is not None and oid is not None:
                    annot_lookup.setdefault(otype, {}).setdefault(oid, []).append(
                        (_safe_str(ann["Name"]), _safe_str(ann["Value"]))
                    )

        # Step 4: Generate TMDL files per table
        tables_dir = model_dir / "tables"
        tables_dir.mkdir(parents=True, exist_ok=True)

        for _, tbl in tables_df.iterrows():
            tbl_id = int(tbl["ID"])
            tbl_name = str(tbl["Name"])

            # Table properties
            tbl_props = {
                "IsHidden": _safe_bool(tbl.get("IsHidden")),
                "IsPrivate": False,  # Not directly queryable; inferred from table name
                "ShowAsVariationsOnly": _safe_bool(tbl.get("ShowAsVariationsOnly")),
                "LineageTag": _safe_str(tbl.get("LineageTag")),
            }
            # Heuristic: DateTableTemplate tables are private
            if "DateTableTemplate" in tbl_name:
                tbl_props["IsPrivate"] = True

            # Filter data for this table
            t_measures = measures_df[measures_df["TableID"] == tbl_id] if not measures_df.empty else pd.DataFrame()
            t_columns = columns_df[columns_df["TableID"] == tbl_id] if not columns_df.empty else pd.DataFrame()
            t_hier = pd.DataFrame()
            if not hier_levels_df.empty and "TableID" in hier_levels_df.columns:
                t_hier = hier_levels_df[hier_levels_df["TableID"] == tbl_id]
            t_vars = pd.DataFrame()
            if not variations_df.empty and "TableID" in variations_df.columns:
                t_vars = variations_df[variations_df["TableID"] == tbl_id]
            t_parts = partitions_df[partitions_df["TableID"] == tbl_id] if not partitions_df.empty else pd.DataFrame()

            # Annotations map for this table
            annot_map = {
                "table": annot_lookup.get(ANNOT_TABLE, {}).get(tbl_id, []),
                "columns": annot_lookup.get(ANNOT_COLUMN, {}),
                "measures": annot_lookup.get(ANNOT_MEASURE, {}),
                "hierarchies": annot_lookup.get(ANNOT_HIERARCHY, {}),
            }

            tmdl_content = generate_table_tmdl(
                tbl_name, tbl_props,
                t_measures, t_columns, t_hier, t_vars, t_parts,
                annot_map,
            )

            safe_name = sanitize_filename(tbl_name)
            tmdl_path = tables_dir / f"{safe_name}.tmdl"
            tmdl_path.write_text(tmdl_content, encoding="utf-8")

        logger.info(f"pbixray: wrote {len(tables_df)} table TMDL files")

        # Step 5: Generate relationships.tmdl and relationships.json
        if not relationships_df.empty:
            rel_content = generate_relationships_tmdl(relationships_df)
            (model_dir / "relationships.tmdl").write_text(rel_content, encoding="utf-8")

            # Also write relationships.json for tmdl_parser.py compatibility
            rel_json = []
            for _, r in relationships_df.iterrows():
                cfb = _safe_int(r.get("CrossFilteringBehavior"))
                cf_str = "bothDirections" if cfb == 2 else ""
                rel_json.append({
                    "fromTable": _safe_str(r.get("FromTableName", "")),
                    "fromColumn": _safe_str(r.get("FromColumnName", "")),
                    "toTable": _safe_str(r.get("ToTableName", "")),
                    "toColumn": _safe_str(r.get("ToColumnName", "")),
                    "isActive": bool(_safe_bool(r.get("IsActive"))),
                    "cardinality": "",
                    "crossFiltering": cf_str,
                })
            (model_dir / "relationships.json").write_text(
                json.dumps(rel_json, indent=2, ensure_ascii=False),
                encoding="utf-8",
            )

            logger.info(f"pbixray: wrote {len(relationships_df)} relationships")

        # Step 6: Generate role TMDL files
        if not rls_df.empty:
            roles_dir = model_dir / "roles"
            roles_dir.mkdir(parents=True, exist_ok=True)
            for role_name in rls_df["RoleName"].unique():
                role_perms = rls_df[rls_df["RoleName"] == role_name]
                role_content = generate_role_tmdl(str(role_name), role_perms)
                safe_role = sanitize_filename(str(role_name))
                (roles_dir / f"{safe_role}.tmdl").write_text(
                    role_content, encoding="utf-8"
                )
            logger.info(
                f"pbixray: wrote {len(rls_df['RoleName'].unique())} role files"
            )

        # Step 7: Generate model.tmdl and database.tmdl stubs
        model_lines = ["model Model", "\tculture: en-US",
                        "\tdefaultPowerBIDataSourceVersion: powerBI_V3",
                        "\tsourceQueryCulture: en-US", ""]
        # Add ref table entries
        for _, tbl in tables_df.iterrows():
            model_lines.append(f"ref table {tmdl_quote(str(tbl['Name']))}")
        model_lines.append("")
        (model_dir / "model.tmdl").write_text(
            "\n".join(model_lines) + "\n", encoding="utf-8"
        )

        (model_dir / "database.tmdl").write_text(
            "database\n\tcompatibilityLevel: 1600\n", encoding="utf-8"
        )

        return True

    except Exception as e:
        logger.warning(f"pbixray: semantic model extraction failed: {e}")
        import traceback
        traceback.print_exc()
        return False
    finally:
        try:
            handler.close_connection()
        except Exception:
            pass


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
        if extract_semantic_model_from_sqlite(pbix_path, model_dir):
            result.model_root = str(model_dir)
            result.semantic_model_source = "pbixray-sqlite"
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
