# -*- coding: utf-8 -*-
"""
Bookmark Parser Module
======================
Parses Power BI bookmark JSON files from PBIP report definitions.
Converts bookmark filter conditions into DAX filter expressions and
tracks visual visibility state per bookmark.

Used by extract_metadata.py (Skill 1) to add a Bookmarks sheet,
and by dax_query_builder.py (Skill 2) to generate CALCULATETABLE-wrapped queries.
"""

import json
import re
from dataclasses import dataclass, field
from pathlib import Path


# ============================================================
# Data classes
# ============================================================

@dataclass
class BookmarkVisual:
    """Visibility state of a single visual within a bookmark."""
    container_id: str
    visual_name: str  # resolved display name (from visual_id_to_name map)
    visible: bool


@dataclass
class BookmarkInfo:
    """Parsed bookmark with resolved page/visual names and DAX filters."""
    name: str               # display name (e.g., "TSV Ribbon ON")
    bookmark_id: str        # internal ID (e.g., "Bookmark2f6f395ae3e64d361f93")
    page_name: str          # resolved page display name
    page_id: str            # section ID (e.g., "ReportSection8ef90f03d68d67c9d9ee")
    filters: list           # list of DAX filter expression strings
    visuals: list           # list of BookmarkVisual


# ============================================================
# Literal parsing
# ============================================================

def parse_literal(value_str: str) -> str:
    """Convert a PBI literal value string to a DAX-compatible representation.

    Formats:
      - String:   'New Store'  → "New Store"
      - DateTime: datetime'2020-06-01T00:00:00' → DATE(2020, 6, 1)
      - Integer:  -6L  → -6  (with comment about relative date)
      - Decimal:  0D   → 0
      - Boolean:  true/false → TRUE/FALSE
      - Null:     null → BLANK()
    """
    if value_str is None:
        return "BLANK()"

    s = str(value_str).strip()

    # Null
    if s.lower() == "null":
        return "BLANK()"

    # Boolean
    if s.lower() in ("true", "false"):
        return s.upper()

    # DateTime: datetime'2020-06-01T00:00:00'
    dt_match = re.match(r"datetime'(\d{4})-(\d{2})-(\d{2})T", s)
    if dt_match:
        y, m, d = dt_match.group(1), dt_match.group(2), dt_match.group(3)
        return f"DATE({int(y)}, {int(m)}, {int(d)})"

    # String literal: 'Some Value' → "Some Value"
    if s.startswith("'") and s.endswith("'") and len(s) >= 2:
        inner = s[1:-1]
        # Escape any double quotes inside the string
        inner = inner.replace('"', '""')
        return f'"{inner}"'

    # Integer with L suffix (PBI long integer): 2025L → 2025, 0L → 0
    # Negative values in date filter context (e.g., -6L) are relative offsets
    # that can't be resolved statically — flag those only
    if re.match(r"^-?\d+L$", s):
        num = s[:-1]
        if num.startswith("-"):
            return f"{num} /* relative offset — cannot resolve statically */"
        return num

    # Decimal with D suffix: 0D → 0
    if re.match(r"^-?\d+(\.\d+)?D$", s):
        return s[:-1]

    # Plain number
    if re.match(r"^-?\d+(\.\d+)?$", s):
        return s

    return s


# ============================================================
# Entity resolution from From[] aliases
# ============================================================

def _build_alias_map(from_entities: list) -> dict:
    """Build a mapping from alias Name → Entity table name.

    Example: [{"Name": "s", "Entity": "Store", "Type": 0}]
    Returns: {"s": "Store"}
    """
    alias_map = {}
    for entry in (from_entities or []):
        alias = entry.get("Name", "")
        entity = entry.get("Entity", "")
        if alias and entity:
            alias_map[alias] = entity
    return alias_map


def _resolve_column_ref(col_expr: dict, alias_map: dict) -> tuple:
    """Resolve a Column expression to (table_name, column_name).

    The Column expression can reference either:
      - SourceRef.Source (alias) → look up in alias_map
      - SourceRef.Entity (direct) → use directly
    """
    col = col_expr.get("Column", {})
    prop = col.get("Property", "")
    source_ref = col.get("Expression", {}).get("SourceRef", {})

    # Try alias first
    source_alias = source_ref.get("Source", "")
    if source_alias and source_alias in alias_map:
        table = alias_map[source_alias]
    else:
        # Direct entity reference
        table = source_ref.get("Entity", "")

    return table, prop


# ============================================================
# Condition → DAX conversion
# ============================================================

# ComparisonKind → DAX operator
_COMPARISON_OPS = {
    0: "=",
    1: ">",
    2: ">=",
    3: "<",
    4: "<=",
    5: "<>",
}


def condition_to_dax(condition: dict, from_entities: list) -> str:
    """Convert a bookmark filter Where.Condition to a DAX filter expression.

    Handles:
      - Comparison (=, >, >=, <, <=, <>)
      - In (column IN {values})
      - Not > In, Not > Contains, Not > StartsWith (negation wrapper)
      - And (left && right, recursive)
      - Or (left || right, recursive)
      - Between (col >= lower && col <= upper)
      - Contains / DoesNotContain (CONTAINSSTRING)
      - StartsWith / DoesNotStartWith (LEFT + LEN)
      - IsBlank / IsNotBlank (ISBLANK)

    Args:
        condition: The Condition dict from Where[].Condition
        from_entities: The From[] array for alias resolution

    Returns:
        DAX filter expression string, e.g. 'Store'[Store Type] = "New Store"
    """
    alias_map = _build_alias_map(from_entities)
    return _condition_to_dax_inner(condition, alias_map)


def _condition_to_dax_inner(condition: dict, alias_map: dict) -> str:
    """Recursive inner function for condition_to_dax.

    Handles: Comparison, In, Not (generic negation wrapper), And, Or,
    Between, Contains, DoesNotContain, StartsWith, DoesNotStartWith,
    IsBlank, IsNotBlank.
    """

    # --- Comparison ---
    if "Comparison" in condition:
        comp = condition["Comparison"]
        kind = comp.get("ComparisonKind", 0)
        op = _COMPARISON_OPS.get(kind, "=")

        left = comp.get("Left", {})
        right = comp.get("Right", {})

        # Left side: column reference
        table, col = _resolve_column_ref(left, alias_map)
        col_ref = f"'{table}'[{col}]" if table else f"[{col}]"

        # Right side: literal value
        lit_val = right.get("Literal", {}).get("Value", "")
        dax_val = parse_literal(lit_val)

        return f"{col_ref} {op} {dax_val}"

    # --- In ---
    if "In" in condition:
        return _in_to_dax(condition["In"], alias_map, negated=False)

    # --- Not (generic negation wrapper) ---
    if "Not" in condition:
        inner_expr = condition["Not"].get("Expression", {})
        if "In" in inner_expr:
            return _in_to_dax(inner_expr["In"], alias_map, negated=True)
        # Fallback: generic Not wrapping (handles Not>Contains, Not>StartsWith, etc.)
        inner_dax = _condition_to_dax_inner(inner_expr, alias_map)
        return f"NOT ({inner_dax})"

    # --- And ---
    if "And" in condition:
        and_node = condition["And"]
        left_dax = _condition_to_dax_inner(and_node.get("Left", {}), alias_map)
        right_dax = _condition_to_dax_inner(and_node.get("Right", {}), alias_map)
        return f"{left_dax} && {right_dax}"

    # --- Or ---
    if "Or" in condition:
        or_node = condition["Or"]
        left_dax = _condition_to_dax_inner(or_node.get("Left", {}), alias_map)
        right_dax = _condition_to_dax_inner(or_node.get("Right", {}), alias_map)
        return f"({left_dax}) || ({right_dax})"

    # --- Between (col >= lower && col <= upper) ---
    if "Between" in condition:
        between = condition["Between"]
        left_col = between.get("Left", {})
        table, col = _resolve_column_ref(left_col, alias_map)
        col_ref = f"'{table}'[{col}]" if table else f"[{col}]"
        lower_val = parse_literal(between.get("Lower", {}).get("Literal", {}).get("Value", ""))
        upper_val = parse_literal(between.get("Upper", {}).get("Literal", {}).get("Value", ""))
        return f"{col_ref} >= {lower_val} && {col_ref} <= {upper_val}"

    # --- Contains → CONTAINSSTRING(col, val) ---
    if "Contains" in condition:
        node = condition["Contains"]
        table, col = _resolve_column_ref(node.get("Left", {}), alias_map)
        col_ref = f"'{table}'[{col}]" if table else f"[{col}]"
        lit_val = node.get("Right", {}).get("Literal", {}).get("Value", "")
        dax_val = parse_literal(lit_val)
        return f"CONTAINSSTRING({col_ref}, {dax_val})"

    # --- DoesNotContain → NOT CONTAINSSTRING(col, val) ---
    if "DoesNotContain" in condition:
        node = condition["DoesNotContain"]
        table, col = _resolve_column_ref(node.get("Left", {}), alias_map)
        col_ref = f"'{table}'[{col}]" if table else f"[{col}]"
        lit_val = node.get("Right", {}).get("Literal", {}).get("Value", "")
        dax_val = parse_literal(lit_val)
        return f"NOT CONTAINSSTRING({col_ref}, {dax_val})"

    # --- StartsWith → LEFT(col, LEN(val)) = val ---
    if "StartsWith" in condition:
        node = condition["StartsWith"]
        table, col = _resolve_column_ref(node.get("Left", {}), alias_map)
        col_ref = f"'{table}'[{col}]" if table else f"[{col}]"
        lit_val = node.get("Right", {}).get("Literal", {}).get("Value", "")
        dax_val = parse_literal(lit_val)
        return f"LEFT({col_ref}, LEN({dax_val})) = {dax_val}"

    # --- DoesNotStartWith → NOT (LEFT(col, LEN(val)) = val) ---
    if "DoesNotStartWith" in condition:
        node = condition["DoesNotStartWith"]
        table, col = _resolve_column_ref(node.get("Left", {}), alias_map)
        col_ref = f"'{table}'[{col}]" if table else f"[{col}]"
        lit_val = node.get("Right", {}).get("Literal", {}).get("Value", "")
        dax_val = parse_literal(lit_val)
        return f"NOT (LEFT({col_ref}, LEN({dax_val})) = {dax_val})"

    # --- IsBlank → ISBLANK(col) ---
    if "IsBlank" in condition:
        node = condition["IsBlank"]
        col_expr = node.get("Expression", {})
        table, col = _resolve_column_ref(col_expr, alias_map)
        col_ref = f"'{table}'[{col}]" if table else f"[{col}]"
        return f"ISBLANK({col_ref})"

    # --- IsNotBlank → NOT ISBLANK(col) ---
    if "IsNotBlank" in condition:
        node = condition["IsNotBlank"]
        col_expr = node.get("Expression", {})
        table, col = _resolve_column_ref(col_expr, alias_map)
        col_ref = f"'{table}'[{col}]" if table else f"[{col}]"
        return f"NOT ISBLANK({col_ref})"

    return "-- unsupported condition type"


def _in_to_dax(in_node: dict, alias_map: dict, negated: bool) -> str:
    """Convert an In condition to DAX IN expression."""
    expressions = in_node.get("Expressions", [])
    values = in_node.get("Values", [])

    if not expressions:
        return "-- empty IN expression"

    # Single column IN
    col_expr = expressions[0]
    table, col = _resolve_column_ref(col_expr, alias_map)
    col_ref = f"'{table}'[{col}]" if table else f"[{col}]"

    # Flatten values: each entry in Values is a list of one literal (for single-column IN)
    dax_values = []
    for val_row in values:
        if val_row and isinstance(val_row, list):
            lit = val_row[0].get("Literal", {}).get("Value", "")
            dax_values.append(parse_literal(lit))

    if len(dax_values) == 1:
        values_str = dax_values[0]
        if negated:
            return f"{col_ref} <> {values_str}"
        else:
            return f"{col_ref} = {values_str}"

    values_str = "{" + ", ".join(dax_values) + "}"
    prefix = "NOT " if negated else ""
    return f"{prefix}{col_ref} IN {values_str}"


# ============================================================
# Filter extraction from bookmark sections
# ============================================================

def _extract_filters_from_section(section: dict) -> list:
    """Extract all filter DAX expressions from a bookmark section's filters.

    Processes both byName and byExpr filter collections.
    Only filters with a "filter" key containing Where clauses have actual values.
    """
    dax_filters = []
    filters_block = section.get("filters", {})

    # Process byName filters (keyed dict)
    by_name = filters_block.get("byName", {})
    for _filter_name, filter_obj in by_name.items():
        dax = _extract_single_filter(filter_obj)
        if dax:
            dax_filters.extend(dax)

    # Process byExpr filters (array)
    by_expr = filters_block.get("byExpr", [])
    for filter_obj in by_expr:
        dax = _extract_single_filter(filter_obj)
        if dax:
            dax_filters.extend(dax)

    return dax_filters


def _extract_single_filter(filter_obj: dict) -> list:
    """Extract DAX expressions from a single filter object.

    Only processes filters that have a "filter" key with "Where" clauses.
    Filters with only an "expression" key are field references without values — skip.
    """
    query_filter = filter_obj.get("filter")
    if not query_filter:
        return []

    from_entities = query_filter.get("From", [])
    where_clauses = query_filter.get("Where", [])

    results = []
    for where in where_clauses:
        condition = where.get("Condition", {})
        if condition:
            dax = condition_to_dax(condition, from_entities)
            if dax and not dax.startswith("--"):
                results.append(dax)

    return results


# Public alias for use by extract_metadata.py (filter value extraction)
extract_single_filter = _extract_single_filter


def _extract_slicer_filters(section: dict) -> list:
    """Extract filter conditions embedded in slicer visual objects.

    Some bookmarks store date range filters inside:
      visualContainers.<id>.singleVisual.objects.merge.general[].properties.filter.filter
    """
    dax_filters = []
    visual_containers = section.get("visualContainers", {})

    for _vis_id, vis_data in visual_containers.items():
        single_visual = vis_data.get("singleVisual", {})
        vis_type = single_visual.get("visualType", "")
        if vis_type != "slicer":
            continue

        # Check objects.merge.general for embedded filter
        objects = single_visual.get("objects", {})
        merge = objects.get("merge", {})
        general_list = merge.get("general", [])
        for general in general_list:
            props = general.get("properties", {})
            filter_wrapper = props.get("filter", {})
            query_filter = filter_wrapper.get("filter")
            if not query_filter:
                continue

            from_entities = query_filter.get("From", [])
            where_clauses = query_filter.get("Where", [])
            for where in where_clauses:
                condition = where.get("Condition", {})
                if condition:
                    dax = condition_to_dax(condition, from_entities)
                    if dax and not dax.startswith("--"):
                        dax_filters.append(dax)

    return dax_filters


# ============================================================
# Visual visibility extraction
# ============================================================

def _extract_visual_visibility(section: dict, page_visual_ids: set,
                               visual_id_to_name: dict) -> dict:
    """Determine visibility of each visual on the page from bookmark state.

    Returns: dict of container_id → bool (visible)

    Logic:
      - If a visual appears in visualContainers with display.mode = "hidden" → hidden
      - If a visual appears in visualContainers without display.mode → visible
      - If a visual appears in visualContainerGroups with isHidden = True → hidden
      - Visuals not mentioned in the bookmark default to visible
    """
    visibility = {}

    # Start with all page visuals as visible (default)
    for vid in page_visual_ids:
        visibility[vid] = True

    # Process visualContainers
    visual_containers = section.get("visualContainers", {})
    for vis_id, vis_data in visual_containers.items():
        single_visual = vis_data.get("singleVisual", {})
        display = single_visual.get("display", {})
        mode = display.get("mode", "")
        if mode == "hidden":
            visibility[vis_id] = False
        elif vis_id in visibility:
            visibility[vis_id] = True

    # Process visualContainerGroups (AI Sample pattern)
    visual_groups = section.get("visualContainerGroups", {})
    for group_id, group_data in visual_groups.items():
        group_hidden = group_data.get("isHidden", False)

        # The group ID itself may be a visual container
        if group_id in page_visual_ids:
            visibility[group_id] = not group_hidden

        # Process children within the group
        children = group_data.get("children", {})
        for child_id, child_data in children.items():
            child_hidden = child_data.get("isHidden", False)
            if child_id in page_visual_ids:
                visibility[child_id] = not child_hidden

    return visibility


# ============================================================
# Main bookmark parsing
# ============================================================

def parse_bookmarks(report_root: str, visual_id_to_name: dict,
                    page_id_to_name: dict,
                    page_id_to_visual_ids: dict = None) -> list:
    """Parse all bookmarks from a PBIP report definition.

    Args:
        report_root: Path to the report definition root (contains bookmarks/ folder)
        visual_id_to_name: Mapping of visual container folder name → display name
        page_id_to_name: Mapping of page section folder name → display name
        page_id_to_visual_ids: Mapping of page section folder name → set of visual container IDs

    Returns:
        List of BookmarkInfo objects
    """
    bookmarks_dir = Path(report_root) / "bookmarks"

    # Check if bookmarks folder and index exist
    index_path = bookmarks_dir / "bookmarks.json"
    if not index_path.is_file():
        return []

    # Read the bookmark index
    index_data = json.loads(index_path.read_text(encoding="utf-8-sig"))
    items = index_data.get("items", [])
    if not items:
        return []

    page_id_to_visual_ids = page_id_to_visual_ids or {}
    bookmarks = []

    for item in items:
        bm_name = item.get("name", "")
        if not bm_name:
            continue

        bm_path = bookmarks_dir / f"{bm_name}.bookmark.json"
        if not bm_path.is_file():
            print(f"    WARNING: Bookmark file not found: {bm_path.name}")
            continue

        bm_data = json.loads(bm_path.read_text(encoding="utf-8-sig"))
        bm_info = _parse_single_bookmark(bm_data, visual_id_to_name,
                                         page_id_to_name, page_id_to_visual_ids)
        if bm_info:
            bookmarks.append(bm_info)

    return bookmarks


def _parse_single_bookmark(bm_data: dict, visual_id_to_name: dict,
                           page_id_to_name: dict,
                           page_id_to_visual_ids: dict) -> BookmarkInfo:
    """Parse a single bookmark JSON into a BookmarkInfo."""
    display_name = bm_data.get("displayName", "")
    bookmark_id = bm_data.get("name", "")

    exploration = bm_data.get("explorationState", {})
    active_section = exploration.get("activeSection", "")
    sections = exploration.get("sections", {})

    # Resolve page name
    page_name = page_id_to_name.get(active_section, active_section)
    page_id = active_section

    # Get the section data for the active page
    section = sections.get(active_section, {})
    if not section:
        return None

    # Extract page-level filters (DAX expressions)
    filters = _extract_filters_from_section(section)

    # Also check for slicer-embedded filters
    slicer_filters = _extract_slicer_filters(section)
    filters.extend(slicer_filters)

    # Extract visual visibility
    page_visuals = page_id_to_visual_ids.get(active_section, set())
    visibility = _extract_visual_visibility(section, page_visuals, visual_id_to_name)

    # Build BookmarkVisual list
    visuals = []
    for vid, is_visible in sorted(visibility.items()):
        vname = visual_id_to_name.get(vid, vid)
        visuals.append(BookmarkVisual(
            container_id=vid,
            visual_name=vname,
            visible=is_visible,
        ))

    return BookmarkInfo(
        name=display_name,
        bookmark_id=bookmark_id,
        page_name=page_name,
        page_id=page_id,
        filters=filters,
        visuals=visuals,
    )
