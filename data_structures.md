# Pipeline Data Structures Reference

Quick reference for the agent on return types, dict keys, and function signatures across all modules.

---

## tmdl_parser.py

### Dataclasses

**TmdlColumn**
```
table: str
name: str
data_type: str          # e.g. "string", "int64", "double", "dateTime"
is_hidden: bool
```

**TmdlRelationship**
```
from_table: str
from_column: str
to_table: str
to_column: str
is_active: bool         # default True
cardinality: str        # e.g. "oneToMany", "manyToOne"
cross_filtering: str    # e.g. "oneDirection", "bothDirections"
```

**SemanticModel**
```
measures: dict          # (table_name, measure_name) → DAX formula string
columns: dict           # (table_name, column_name) → TmdlColumn
_measure_index: dict    # lowercase_name → list of (table, name) tuples
_column_index: dict     # lowercase_name → list of (table, name) tuples
source: str             # "pbixray" | "pbip" | ""
relationships: list     # list of TmdlRelationship
```

Properties:
- `model.types_reliable` → bool (False when source is "pbixray" or "")
- `model.measure_names` → the `_measure_index` dict
- `model.column_names` → the `_column_index` dict

### Key Functions

**parse_semantic_model(model_root) → SemanticModel**
- `model_root`: path to semantic model definition root (contains `tables/` dir)
- Reads `.source` marker, `relationships.json` or `relationships.tmdl`, then all `.tmdl` files
- Calls `build_indexes()` automatically

**match_field_to_model(field_name, model) → dict | None**
- Returns: `{"table": str, "field_name": str, "formula": str, "match_type": str}` or None
- `match_type` is one of: `"measure"`, `"column"`, `"measure_fuzzy"`, `"column_fuzzy"`

---

## dax_query_builder.py

### read_extractor_output(filepath) → (visuals, page_filters, bookmarks, filter_expr_data)

Returns a 4-tuple:

**visuals: OrderedDict**
- Key: `(page_name: str, visual_id_or_name: str)` — a **tuple**, not a string
- Value: dict with these keys:
  ```
  {
      "visual_type": str,        # e.g. "barChart", "donutChart", "slicer"
      "visual_name": str,        # display name e.g. "Donut Chart"
      "visual_id": str,          # container folder ID (may be "")
      "z_index": int,            # z-order for bookmark-toggled copies
      "sort_order": str,         # DAX ORDER BY expression (may be "")
      "fields": [field_dict, ...]
  }
  ```

**field_dict** (each entry in the `"fields"` list):
```
{
    "ui_name": str,              # display name in the visual
    "usage": str,                # e.g. "Visual Column", "Visual Value", "Slicer"
    "well": str,                 # e.g. "X-axis", "Y-axis", "Legend", "Values", "Matrix Columns"
    "table_sm": str,             # table name in semantic model
    "col_sm": str,               # column/measure name in semantic model
    "measure_formula": str,      # DAX formula (may be "")
    "agg_func": str,             # implicit agg e.g. "Sum", "Avg" (may be "")
    "data_type": str,            # e.g. "string", "int64" (may be "")
    "model_source": str,         # "pbixray" or "pbip" (may be "")
}
```

> **IMPORTANT**: field dicts do NOT have a `"role"` key. Use `"usage"` for the usage label and `"well"` for the well assignment. To classify a field, call `classify_field(f["usage"], f.get("well", ""))`.

**page_filters: dict**
- Key: `page_name: str`
- Value: `list[field_dict]` — same field_dict shape as above

**bookmarks: list[dict]**
- Each dict:
  ```
  {
      "bookmark_name": str,
      "page_name": str,
      "container_id": str,
      "visual_name": str,
      "visible": str,            # "Y" or "N"
      "filter_dax": str,         # semicolon-separated DAX filter expressions
  }
  ```

**filter_expr_data: list[dict]**
- Each dict:
  ```
  {
      "page_name": str,
      "visual_name": str,
      "visual_id": str,
      "filter_level": str,       # "Report", "Page", "Visual", or "Slicer"
      "filter_field": str,       # e.g. "'Store'[Region]"
      "filter_dax_expr": str,    # full DAX expression
  }
  ```

### Iterating Over Visuals

```python
# CORRECT — unpack the tuple key
for (page, vid), data in visuals.items():
    visual_name = data["visual_name"]
    fields = data["fields"]

# WRONG — these will all fail:
# visuals[0]               → KeyError (keys are tuples, not ints)
# v.page_name              → AttributeError (values are dicts, not objects)
# v["page_name"]           → KeyError (no such key in value dict)
```

### classify_field(usage, well="") → str

Returns one of: `"grouping"`, `"measure"`, `"filter"`, `"slicer"`, `"page_filter"`, `"matrix_column"`, `"other"`

- Prefers `well`-based classification when well is non-empty
- Falls back to parsing the `usage` string

### classify_visual_fields(fields) → (grouping, measures, filters, slicer_fields, matrix_columns)

- Input: list of field_dicts
- Returns 5 lists of field_dicts, split by role
- Implicit measures (field has `agg_func` set) get reclassified from grouping → measure

### build_dax_query(grouping, measures, filters, slicer_fields, visual_type, model=None, matrix_columns=None, sort_order=None) → (pattern, dax)

- Returns a **2-tuple**: `(pattern_name: str, dax_query: str)`
- Pattern names: `"Pattern 1: Single Measure"`, `"Pattern 1: Multiple Measures"`, `"Pattern 2: Columns Only"`, `"Pattern 3: Columns + Measures"`, `"Pattern 3M: Matrix Summary"`, `"Unknown"`

### find_visual(visuals, search_term) → list[keys]

- Returns list of `(page_name, visual_id_or_name)` tuple keys
- Case-insensitive partial match against visual_name and "page / visual_name"
- Sorted by z_index descending (highest z = default visible copy first)

### get_single_visual_query(visuals, page_filters, visual_search, filter_exprs=None, having_exprs=None, model=None, filter_expr_data=None, column_values=None, flat_measures=None) → dict | None

- `visual_search`: str — partial name match (not a page_name keyword arg)
- Returns dict:
  ```
  {
      "page": str,
      "visual_name": str,
      "visual_type": str,
      "pattern": str,
      "dax_query": str,              # filtered DAX
      "base_dax_query": str,         # unfiltered DAX
      "filters_applied": str,        # semicolon-separated
      "having_applied": str,
      "preset_filters_applied": str,
      "matrix_columns": list[dict],  # [{"table", "column", "ui_name"}, ...]
      "values_query": str,           # preflight VALUES() DAX for matrix
      "pivot_dax_query": str,        # pivoted CALCULATE DAX for matrix
      "auto_flat_measures": list,    # measure ui_names excluded from pivot
  }
  ```

### collect_filters_for_visual(page_name, visual_name, visual_id, filter_expr_data) → list[str]

- Returns list of DAX filter expression strings
- Follows hierarchy: Report → Page → Slicer → Visual
- Skips TopN and entries starting with `--`

### wrap_dax_with_filters(base_dax, filter_exprs, pattern) → str

- Wraps DAX with CALCULATETABLE (column filters) and FILTER (measure filters)
- Pattern 1 Single Measure uses CALCULATE instead
- `pattern` must be the pattern name string (e.g. `"Pattern 3: Columns + Measures"`)

### wrap_dax_with_having(dax, having_exprs) → str

- Wraps with FILTER for post-aggregation conditions

---

## bookmark_parser.py

### Dataclasses

**BookmarkVisual**
```
container_id: str
visual_name: str        # resolved display name
visible: bool
```

**BookmarkInfo**
```
name: str               # display name
bookmark_id: str        # internal ID
page_name: str          # resolved page display name
page_id: str            # section folder ID
filters: list[str]      # DAX filter expression strings
visuals: list[BookmarkVisual]
```

### Key Functions

**parse_bookmarks(report_root, visual_id_to_name, page_id_to_name, page_id_to_visual_ids=None) → list[BookmarkInfo]**

**extract_single_filter(filter_obj) → list[str]**
- Extracts DAX expressions from a filter JSON object with `"filter"` key containing `"Where"` clauses

---

## extract_metadata.py

### extract_metadata(report_root, model_root, include_bookmarks=True) → (df, bookmarks_list, filter_expressions)

- `df`: pandas DataFrame with columns:
  ```
  Page Name, Visual/Table Name in PBI, Visual ID, Visual Type,
  UI Field Name, Usage (Visual/Filter/Slicer), Well, Measure Formula,
  Table in the Semantic Model, Column in the Semantic Model,
  Aggregation Function, Data Type, Semantic Model Source, Z Index, Sort Order
  ```
- `bookmarks_list`: list of BookmarkInfo objects
- `filter_expressions`: list of dicts with keys:
  ```
  Page Name, Visual Name, Visual ID, Filter Level, Filter Field, Filter DAX Expression
  ```

### export_to_excel(df, output_path, bookmarks_list=None, filter_expressions=None)

- Writes DataFrame to Excel with sheets: "Report Metadata", optionally "Bookmarks" and "Filter Expressions"

---

## pbix_extractor.py

### PbixExtractResult (dataclass)
```
report_root: str
model_root: Optional[str]
report_name: str
page_count: int
visual_container_count: int
data_visual_count: int
bookmark_count: int
semantic_model_source: str     # "pbixray-sqlite" | "user-provided" | "none"
```

---

## chart_generator.py

### VisualSpec (dataclass)
```
page_name: str
visual_name: str
visual_type: str                  # PBI camelCase e.g. "barChart"
grouping_columns: list[str]       # X-axis / row field names
measure_columns: list[str]        # Y-axis / values field names
y2_columns: list[str]             # secondary axis measures
series_columns: list[str]         # legend / series field names
facet_column: str                 # small multiples field (may be "")
dax_pattern: str                  # e.g. "Pattern 3"
```

### Constants

- `NATIVE_CHART_MAP`: maps PBI visual type → XL_CHART_TYPE enum
- `PNG_FALLBACK_TYPES`: visual types rendered as plotly PNG (`{"waterfallChart", "funnelChart", "treemap", "gauge"}`)
- `PBI_COLORS`: 10-color hex palette
- `_SKIP_TYPES` (from dax_executor patterns): visual types to skip during batch execution

---

## Common Gotchas

1. **visuals is an OrderedDict with tuple keys** — `(page_name, visual_id)`, not integer indices
2. **field dicts use `"usage"` not `"role"`** — there is no `"role"` key
3. **field dicts use `"well"` for well assignment** — e.g. "X-axis", "Y-axis", "Legend"
4. **`get_single_visual_query()` does NOT accept a `page_name` keyword arg** — pass the visual search string as the 3rd positional arg; it does partial matching across all pages
5. **`build_dax_query()` returns a 2-tuple** `(pattern, dax)`, not a dict
6. **`build_matrix_pivot_query()` may return a 2-tuple or 3-tuple** — check `len()` before unpacking
7. **`classify_field()` takes `(usage, well)` not a field dict** — call it as `classify_field(f["usage"], f.get("well", ""))`
8. **CSV output uses UTF-8-sig encoding** with no index column — column names match `ui_name` from metadata
