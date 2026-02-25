---
name: build-dax-queries
description: Generate DAX queries from PBI metadata extractor output. Use when the user asks to generate DAX queries, build DAX, reverse-engineer visuals into queries, or get the data behind a visual. Also use when the user mentions DAX patterns, SUMMARIZECOLUMNS, or EVALUATE.
---

# Build DAX Queries from Report Metadata

## What This Skill Does
Reads the metadata extractor output (Skill 1's Excel) and generates a DAX query for each visual in the report. Classifies each field by role (grouping column, measure, filter, slicer), then applies the appropriate DAX pattern to construct an `EVALUATE` query that reproduces the visual's underlying data.

## When to Use
- User asks to "generate DAX queries" or "build DAX"
- User wants to reverse-engineer a visual into a DAX query
- User has already run extract-metadata and wants the next step
- This is Step 2 of the DAX query generation pipeline

## How to Run

### As an import:
```python
import sys
sys.path.insert(0, "skills")
from dax_query_builder import read_extractor_output, classify_visual_fields, build_dax_query, add_filter_comments, write_output

visuals, page_filters = read_extractor_output("output/pbi_report_metadata.xlsx")
count = write_output(visuals, page_filters, "output/dax_queries.xlsx")
print(f"Generated DAX queries for {count} visuals")
```

### From command line:
```bash
python skills/dax_query_builder.py "output/pbi_report_metadata.xlsx" "output/dax_queries.xlsx"
```

### Revenue Opportunities test run:
```bash
python skills/dax_query_builder.py "output/pbi_report_metadata.xlsx" "output/dax_queries_revenue_opportunities.xlsx"
```

## Required Inputs
1. **Metadata extractor Excel** (positional arg) — output from extract-metadata (Skill 1)
2. **Output path** (optional positional arg) — defaults to `dax_queries_<input_name>.xlsx`

## Output
- Excel file with columns: Page Name, Visual Name, Visual Type, DAX Pattern, DAX Query, Filter Fields, Validated?
- One row per visual with a complete DAX query

## DAX Patterns
| Pattern | When Used | DAX Template |
|---------|-----------|--------------|
| Pattern 1: Single Measure | Card with one measure | `EVALUATE { [Measure] }` |
| Pattern 1: Multiple Measures | Card/KPI with multiple measures | `EVALUATE ROW("Name", [Measure], ...)` |
| Pattern 2: Columns Only | Slicers, column-only visuals | `EVALUATE VALUES(...)` or `EVALUATE DISTINCT(SELECTCOLUMNS(...))` |
| Pattern 3: Columns + Measures | Most chart/table visuals | `EVALUATE SUMMARIZECOLUMNS(columns, "Name", [Measure], ...)` |
| Unknown | Fallback | Comment explaining the issue |

## Validation
- Compare generated queries against `~/Desktop/dax_queries_by_visual.xlsx` (manual reference for Revenue Opportunities)
- For non-table visuals: convert to table view in PBI, then verify DAX output matches the displayed data
- Check that all visuals with data fields have a generated query (not "Unknown")

## Known Limitations
- Filter values are NOT extracted — only filter field references are captured (appear as comments in DAX)
- Complex calculated columns or unusual aggregations may produce imperfect queries
- Page-level filters are appended to all visuals on that page but cannot be embedded as CALCULATETABLE filters (values unknown)

## Critical Rules
- NEVER modify the input Excel file — it is read-only
- ALWAYS save output to the `output/` folder
- Filter comments must clearly state that values are not extracted
- Measure references in DAX use `[MeasureName]` (no table prefix) — this matches how Power BI resolves them
- **Measure filters in CALCULATETABLE/CALCULATE must be wrapped in FILTER().** Bare measure boolean expressions like `[Rev Goal] > 0` are INVALID as CALCULATETABLE filter arguments — they cause "A function 'PLACEHOLDER' has been used in a True/False expression" errors. Wrap them: `FILTER(ALL('Table'), [Measure] > 0)`. Column filters like `'Table'[Column] >= value` work directly without wrapping. Rule of thumb: `[MeasureName]` (no table prefix) → needs FILTER(); `'Table'[Column]` → use directly.
