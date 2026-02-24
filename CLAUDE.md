# PBI DAX Query Generation Agent

## Welcome Message — Display on Every Launch
When a new conversation starts, immediately greet the user with the following message (before they ask anything):

---

**PBI DAX Query Generation Agent**

I reverse-engineer Power BI visuals into executable DAX queries — and now I can also **generate chart images** from the query results that resemble the original PBI visuals.

**Three ways to feed me input:**

1. **PBIP files** — Point me at a `.Report/` and `.SemanticModel/` folder pair. I'll parse the JSON + TMDL files directly and extract every visual, field, filter, and measure. *(Most complete — includes bookmarks if present.)*
2. **CSV / Excel exports** — Right-click a visual in Power BI Desktop, choose "Export data", and give me the resulting file. I'll infer field roles from the data and match against the semantic model. *(Quick per-visual approach.)*
3. **Screenshots** — Give me a PNG/JPG of a Power BI visual. I'll read the chart, identify the visual type, fields, and measures, then build the metadata. *(Works when you only have an image.)*

All three paths produce the same standardized metadata, which I then convert into DAX queries. Once you execute those queries and get tabular data back, I can generate **chart images (PNG)** that visually match the original PBI visuals using Skill 5.

**Sample reports you can try me on right now (already in `data/`):**

| Report | What it tests | Run command |
|---|---|---|
| **Revenue Opportunities** | 11 visuals, no bookmarks. Reference report with manual validation files. | `run-dax-pipeline` with report=Revenue Opportunities |
| **Store Sales** | 17 visuals, 2 bookmarks (Store Type filter + visual show/hide). | `run-dax-pipeline` with report=Store Sales |
| **AI Sample** | 10 visuals, 17 bookmarks (IN, NOT IN, date ranges, relative dates). | `run-dax-pipeline` with report=Artificial Intelligence Sample (2) |
| **CSV export** | Single visual export for Skill 3 testing. | Feed `data/Total Sales Variance by FiscalMonth and District Manager.csv` through the CSV reader |

Just tell me which report or file to process, or drop a screenshot — and I'll get started!

---

## What This Project Does
Reverse-engineers Power BI visuals into executable DAX queries. Given a PBIP report's JSON and TMDL files, the pipeline extracts every visual's fields, filters, and measures, then deterministically constructs DAX `EVALUATE` queries that reproduce each visual's underlying data.

This is the **data retrieval layer** for Lara's AI-Powered Slide Generation project at XP3R. The generated DAX queries feed into an agent that queries Microsoft Fabric Semantic Models, retrieves tabular data, and produces PowerPoint slides with AI-generated insights.

### Bigger Picture: How This Fits
```
  PBIP files -----> [Skill 1] extract_metadata.py -----+
  CSV/Excel  -----> [Skill 3] read_excel_export.py ----+--> 8-col metadata Excel --+--> [Skill 2] dax_query_builder.py --> DAX queries
  Screenshots ----> [Skill 4] analyze_screenshot.py ---+                           |
                                                       ^                           +--> Bookmarks sheet (optional)
                            [Shared] tmdl_parser.py (semantic model matching)      |
                            [Shared] bookmark_parser.py (bookmark filter parsing)--+

  DAX queries --> Execute against Fabric --> tabular CSV data --+--> [Skill 5] chart_generator.py --> .pptx (native chart or PNG fallback)
                                                                |                                    or .png (legacy mode)
  Metadata Excel (visual type + field roles) -------------------+

  .pptx chart slides --> Lara's Agent --> AI insights --> PowerPoint slides
  Bookmark DAX queries--^  (filtered per-bookmark view)
```

The key insight: if you can see data in a PBI visual, you can reconstruct the DAX query that produces that data. Three input methods are supported:
- **PBIP files** (Skill 1): parse JSON + TMDL files directly
- **CSV/Excel exports** (Skill 3): right-click visual in PBI -> Export data
- **Screenshots** (Skill 4): analyze visual images via Claude Vision API

All three produce the same 8-column metadata Excel that feeds into Skill 2 unchanged.

### Current Approach: Deterministic Extraction (Method 1)
The pipeline uses fully automated, code-based DAX construction. No AI in the loop. Field metadata is extracted from the visual JSON, classified by role (grouping column, measure, filter, slicer), then assembled into a DAX `EVALUATE` query using the appropriate pattern (SUMMARIZECOLUMNS, VALUES, ROW, etc.).

AI-generated DAX queries (Method 2) are a future possibility for complex edge cases but are **out of scope for now**.

## Stack
- Python 3.x
- pandas, openpyxl (Excel I/O)
- regex (TMDL file parsing, DAX formula analysis)
- urllib.request, base64 (stdlib -- Claude Vision API for Skill 4)
- plotly, kaleido (chart rendering + static image export for Skill 5 PNG mode/fallback)
- python-pptx (native PowerPoint chart generation for Skill 5 PPTX mode)
- Power BI Desktop PBIP format (JSON + TMDL files)

## Project Structure
```
powerpointTask/
├── CLAUDE.md                       # This file
├── skills/
│   ├── tmdl_parser.py              # Shared: TMDL semantic model parser
│   ├── bookmark_parser.py          # Shared: Bookmark filter parsing + DAX conversion
│   ├── extract_metadata.py         # Skill 1: PBIP metadata extraction (+Bookmarks sheet)
│   ├── dax_query_builder.py        # Skill 2: DAX query generation (+Bookmark DAX Queries sheet)
│   ├── read_excel_export.py        # Skill 3: CSV/Excel export reader
│   ├── analyze_screenshot.py       # Skill 4: Screenshot analyzer (Claude Vision)
│   └── chart_generator.py          # Skill 5: Chart image generator (plotly)
├── data/                           # Input PBIP folders + CSV/screenshots go here
│   ├── <ReportName>.Report/        # PBIP report definition (may include bookmarks/)
│   ├── <ReportName>.SemanticModel/ # PBIP semantic model
│   └── *.csv / *.xlsx / *.png      # Exported data files or screenshots
└── output/                         # All generated outputs go here
```

## Skill Details

### Skill 1: extract_metadata.py
Parses PBIP report files (JSON + TMDL) to extract every visual, field, filter, and measure used in a report. Recursively resolves nested measure dependencies to trace all underlying column references. Optionally parses bookmarks to extract filter values and visual visibility state.

- **Input:**
  - `--report-root` — Path to PBIP report definition root (contains `pages/`, `report.json`)
  - `--model-root` — Path to semantic model definition root (contains `tables/` with `.tmdl` files)
  - `--output` — Output Excel file path (default: `pbi_report_metadata.xlsx`)
  - `--no-bookmarks` — Disable bookmark extraction (default: bookmarks are extracted if present)
- **Output:** Excel with two sheets:
  - **Report Metadata** — 8 columns: Page Name, Visual/Table Name in PBI, Visual Type, UI Field Name, Usage (Visual/Filter/Slicer), Measure Formula, Table in the Semantic Model, Column in the Semantic Model
  - **Bookmarks** (if bookmarks/ folder exists) — 6 columns: Bookmark Name, Page Name, Visual Container ID, Visual Name, Visible (Y/N), Filter DAX
- **Key logic:**
  - `resolve_measure_dependencies()` — recursive DAX formula parsing with visited-set cycle prevention
  - `extract_field_info()` — handles Column, Measure, Aggregation, and HierarchyLevel field types
  - Extracts report-level, page-level, and visual-level filters separately
  - Builds visual_id_to_name and page_id_to_name mappings for bookmark resolution
  - `extract_metadata()` returns `(df, bookmarks_list)` tuple

```bash
python skills/extract_metadata.py \
  --report-root "data/<ReportName>.Report/definition" \
  --model-root "data/<ReportName>.SemanticModel/definition" \
  --output "output/pbi_report_metadata.xlsx"
```

### Skill 2: dax_query_builder.py
Reads the metadata extractor output (Skill 1's Excel) and generates a DAX query for each visual in the report. Classifies each field by role, then applies one of four DAX patterns. If bookmark data is present, generates an additional sheet with filter-aware queries.

- **Input:** Positional arg — path to the metadata extractor Excel file
- **Output:** Excel file with one or two sheets:
  - **DAX Queries by Visual** — columns: Page Name, Visual Name, Visual Type, DAX Pattern, DAX Query, Filter Fields, Validated?
  - **Bookmark DAX Queries** (if Bookmarks sheet present in input) — columns: Bookmark Name, Page Name, Visual Name, Visual Type, DAX Pattern, DAX Query, Filters Applied, Validated?
- **DAX Patterns:**
  - **Pattern 1 (Measures Only):** Cards, KPIs — `EVALUATE { [Measure] }` or `EVALUATE ROW(...)`
  - **Pattern 2 (Columns Only):** Slicers, column-only visuals — `EVALUATE VALUES(...)` or `EVALUATE DISTINCT(SELECTCOLUMNS(...))`
  - **Pattern 3 (Columns + Measures):** Most visuals — `EVALUATE SUMMARIZECOLUMNS(...)`
  - **Unknown:** Fallback comment when pattern can't be determined
- **Bookmark DAX wrapping:**
  - Pattern 1 Single Measure → `EVALUATE { CALCULATE([Measure], filter1, ...) }`
  - All others → `EVALUATE CALCULATETABLE(<inner>, filter1, ...)`
  - Only visible visuals (Visible = Y) get bookmark queries
- **Key logic:**
  - `classify_field()` — maps Usage labels to roles: grouping, measure, filter, slicer, page_filter
  - `build_dax_query()` — selects the DAX pattern based on which field roles are present
  - `wrap_dax_with_filters()` — wraps base DAX with CALCULATETABLE/CALCULATE for bookmark filters
  - `build_bookmark_queries()` — orchestrates bookmark DAX generation for all visible visuals
  - Page-level filters are appended to each visual's filter list
  - Unextracted filter values appear as comments in the main DAX Queries sheet
  - `parse_filter_column_refs()` — extracts `(table, column)` pairs from DAX filter expressions
  - `check_filter_redundancy()` — detects when external filter expressions target columns already referenced in measure formulas, preventing CALCULATETABLE conflicts with internal measure logic
- **Filter redundancy check:**
  - Before wrapping with CALCULATETABLE, the builder checks if measure formulas already reference the filtered column
  - Example: `[Revenue Won]` internally filters `Status = "Won"` — adding an external `'Opportunities'[Status] IN {"Open", "Won"}` would conflict
  - Conflicting filters are automatically skipped with a WARNING printed to stdout
  - Non-conflicting filters are still applied normally
  - Works in both `--visual` single-query mode and bookmark DAX generation
  - Formula source priority: metadata Excel "Measure Formula" column → `--model-root` semantic model fallback
  - `--model-root` — optional CLI arg to load semantic model for formula lookup when metadata Excel lacks formulas (e.g., CSV/screenshot inputs)

```bash
python skills/dax_query_builder.py "output/pbi_report_metadata.xlsx" "output/dax_queries.xlsx"
```

### Shared Module: tmdl_parser.py
Reusable TMDL semantic model parser. Extracts both measures AND columns from TMDL files into a `SemanticModel` dataclass with case-insensitive lookup indexes. Used by Skills 1, 3, and 4.

- **Key functions:**
  - `parse_semantic_model(model_root)` -- returns `SemanticModel` with measures, columns, and indexes
  - `parse_tmdl_files(tables_dir)` -- legacy wrapper for Skill 1 (returns measures dict only)
  - `match_field_to_model(field_name, model)` -- matches a bare field name to the model
    - Priority: exact measure -> exact column -> fuzzy match (normalized) -> None
- **Data classes:**
  - `SemanticModel` -- measures dict, columns dict, case-insensitive name indexes
  - `TmdlColumn` -- table, name, data_type, is_hidden

### Shared Module: bookmark_parser.py
Parses bookmark JSON files from PBIP report definitions. Converts bookmark filter conditions into DAX filter expressions and tracks visual visibility per bookmark.

- **Key functions:**
  - `parse_bookmarks(report_root, visual_id_to_name, page_id_to_name, page_id_to_visual_ids)` -- main entry point, returns list of `BookmarkInfo`
  - `condition_to_dax(condition, from_entities)` -- converts a JSON filter `Where.Condition` to a DAX expression
  - `parse_literal(value_str)` -- converts PBI literal values to DAX format (strings, dates, relative offsets)
- **Supported filter condition types:**
  - `Comparison` (=, >, >=, <, <=, <>) -- e.g., `'Store'[Store Type] = "New Store"`
  - `In` -- e.g., `'Opportunities'[Status] IN {"Open", "Won"}`
  - `Not > In` -- e.g., `NOT 'Opportunities'[Status] IN {"Lost"}`
  - `And` (recursive) -- e.g., `'Calendar'[Date] >= DATE(2020, 6, 1) && 'Calendar'[Date] < DATE(2021, 6, 1)`
- **Literal value formats:**
  - String: `'New Store'` → `"New Store"` (PBI single quotes → DAX double quotes)
  - DateTime: `datetime'2020-06-01T00:00:00'` → `DATE(2020, 6, 1)`
  - Relative: `-6L` → `-6 /* relative offset */` (cannot resolve statically)
- **Data classes:**
  - `BookmarkInfo` -- name, bookmark_id, page_name, page_id, filters (DAX strings), visuals (BookmarkVisual list)
  - `BookmarkVisual` -- container_id, visual_name, visible (bool)
- **Entity resolution:** The `From` array in filter JSON maps aliases to entity names (e.g., `{"Name": "s", "Entity": "Store"}` means `SourceRef.Source: "s"` → table `Store`)

### Skill 3: read_excel_export.py
Reads CSV or Excel files exported from Power BI visuals (right-click -> Export data) and produces the standard 8-column metadata Excel.

- **Input:**
  - Positional args -- one or more CSV/Excel files
  - `--model-root` -- optional path to semantic model (for field matching)
  - `--output` -- output Excel file path (default: `pbi_report_metadata.xlsx`)
  - `--page-name` -- page name to use in output (default: `Page 1`)
- **Output:** Same 8-column Excel as Skill 1
- **How it works:**
  1. Parse filename -> visual name
  2. Read column headers -> field names
  3. Classify each field by analyzing first 20 data values:
     - `$`, `%`, or pure numeric -> measure -> Usage = "Visual Value"
     - Text/categorical -> grouping -> Usage = "Visual Column"
  4. If `--model-root` provided, match fields against semantic model for exact table/formula
  5. Without model -> placeholder `'<Table>'[FieldName]`

```bash
# With semantic model matching
python skills/read_excel_export.py "data/export.csv" \
  --model-root "data/<ReportName>.SemanticModel/definition" \
  --output "output/pbi_report_metadata.xlsx"

# Without model (placeholder tables)
python skills/read_excel_export.py "data/export.csv" \
  --output "output/pbi_report_metadata.xlsx"
```

### Skill 4: analyze_screenshot.py
Builds the standard 8-column metadata Excel from a visual description. The Claude Code agent reads the screenshot directly, identifies the visual type/name/fields, then passes them as CLI arguments. No external API key needed.

- **Input:**
  - `--visual-type` -- PBI visual type (e.g., `pieChart`, `barChart`, `card`)
  - `--visual-name` -- Visual title as shown in the report
  - `--field` -- Field in `name:role` format, repeatable (`role` = `measure` or `grouping`)
  - `--model-root` -- optional path to semantic model (for field matching)
  - `--output` -- output Excel file path (default: `pbi_report_metadata.xlsx`)
  - `--page-name` -- page name to use in output (default: `Page 1`)
- **Output:** Same 8-column Excel as Skill 1
- **How it works:**
  1. Agent (Claude Code) views the screenshot and identifies visual type, name, fields
  2. Agent calls this script with the identified info as CLI args
  3. Script matches fields against semantic model if provided
  4. Normalizes visual type to PBI identifiers
  5. Outputs standard metadata Excel
- **Dependencies:** No external API keys. Only pandas, openpyxl.

```bash
# Agent identifies: pie chart, "This Year Sales by Chain", fields Chain (grouping) + This Year Sales (measure)
python skills/analyze_screenshot.py \
  --visual-type pieChart \
  --visual-name "This Year Sales by Chain" \
  --field "Chain:grouping" \
  --field "This Year Sales:measure" \
  --model-root "data/<ReportName>.SemanticModel/definition" \
  --output "output/screenshot_metadata.xlsx"
```

### Skill 5: chart_generator.py
Generates PBI-style chart visuals from DAX query tabular data. Supports two output formats: **PPTX** (default, single-slide PowerPoint with native editable chart or PNG fallback) and **PNG** (legacy plotly image). Uses python-pptx for native charts and plotly for PNG rendering. Supports 16 chart renderers covering all common PBI visual types.

- **Input (two modes):**
  - **Mode 1 (CSV + Metadata):** `--csv` data file + `--metadata` Excel from Skill 1 + `--visual` name to match
  - **Mode 2 (CSV + Screenshot):** `--csv` data file + `--visual-type` + `--field` args (agent passes visual info from screenshot)
  - `--format` — Output format: `pptx` (default) or `png` (legacy)
  - `--output` — Output directory for chart files (default: `output/charts/`)
  - `--width` / `--height` / `--scale` — Image dimensions for PNG mode (default: 1100x500, scale=2 for 144 DPI)
- **Output:**
  - **PPTX mode (default):** Single-slide `.pptx` file per visual. Contains either a native editable chart or an embedded plotly PNG image.
  - **PNG mode (legacy):** plotly-rendered static PNG image.
- **Dual rendering engine (PPTX mode):**
  - **Native python-pptx charts** (editable in PowerPoint): `barChart`, `clusteredBarChart`, `stackedBarChart`, `hundredPercentStackedBarChart`, `columnChart`, `clusteredColumnChart`, `stackedColumnChart`, `hundredPercentStackedColumnChart`, `lineChart`, `areaChart`, `stackedAreaChart`, `pieChart`, `donutChart`, `scatterChart`
  - **PNG fallback on slide** (plotly image inserted as picture): `lineClusteredColumnComboChart`, `lineStackedColumnComboChart`, `waterfallChart`, `funnelChart`, `treemap`, `gauge`, `card`, `multiRowCard`, `kpi`, `tableEx`, `pivotTable`, `ribbonChart`
  - **Skip:** slicers, maps, AI visuals (not meaningful as static charts)
- **Plotly chart type routing (PNG mode / fallback):** `CHART_TYPE_ROUTER` dict maps PBI visual type identifiers to plotly renderer functions:
  - **Bar/Column:** `barChart`, `clusteredBarChart`, `columnChart`, `clusteredColumnChart` → `go.Bar` (horizontal or vertical)
  - **Stacked:** `stackedBarChart`, `stackedColumnChart`, `hundredPercentStacked*` → `go.Bar` with `barmode="stack"`
  - **Line/Area:** `lineChart`, `areaChart`, `stackedAreaChart` → `go.Scatter` with lines/fill
  - **Pie/Donut:** `pieChart`, `donutChart` → `go.Pie` (hole=0 or 0.4)
  - **Waterfall:** `waterfallChart` → `go.Waterfall` (native plotly)
  - **Combo:** `lineStackedColumnComboChart`, `lineClusteredColumnComboChart` → `make_subplots` with dual Y-axis
  - **Funnel/Treemap:** `funnelChart`, `treemap` → `go.Funnel`, `go.Treemap`
  - **Gauge/Card/KPI:** `gauge`, `card`, `multiRowCard`, `kpi` → `go.Indicator`
  - **Table:** `tableEx`, `pivotTable` → `go.Table` (also fallback for unknown types)
- **Key data classes:**
  - `VisualSpec` — page_name, visual_name, visual_type, grouping_columns, measure_columns, y2_columns, dax_pattern
- **Key functions:**
  - `generate_chart(df, spec)` — plotly API; always returns a plotly Figure (or None). Use for PNG output.
  - `generate_chart_pptx(df, spec)` — PPTX API; always returns a `pptx.presentation.Presentation` (or None). Uses native chart when possible, PNG fallback on slide otherwise.
  - `save_chart(fig, output_path, width, height, scale)` — exports plotly Figure to PNG via kaleido
  - `save_chart_pptx(prs, output_path)` — saves Presentation as .pptx
  - `_add_native_chart(slide, df, spec)` — routes to correct native python-pptx chart type, returns True/False
  - `_build_category_chart_data(df, categories, values)` — builds CategoryChartData for bar/column/line/area/pie
  - `_build_xy_chart_data(df, spec)` — builds XyChartData for scatter charts
  - `_style_native_chart(chart, spec, num_series)` — applies PBI colors, fonts, legend to native charts
  - `_add_png_to_slide(slide, png_path)` — inserts plotly PNG as picture shape (for fallback types)
  - `classify_columns(df, spec)` — matches VisualSpec field names to DataFrame columns (case-insensitive), falls back to dtype inference
  - `_prepare_series_data(df, categories, values)` — shared pivot helper for plotly bar/column/stacked renderers
  - `parse_visual_from_metadata(metadata_excel, visual_name)` — reads Skill 1 output, builds VisualSpec using `classify_field()` from dax_query_builder
- **PBI styling:** All charts use PBI's default color palette (`#118DFF`, `#12239E`, `#E66C37`, ...), Segoe UI font, white background, light gray gridlines. Native PPTX charts use 16:9 widescreen slides (13.333" x 7.5").
- **Dependencies:** plotly, kaleido, pandas, openpyxl, python-pptx

```bash
# Mode 1: CSV + Metadata Excel (default: PPTX output with native chart)
python skills/chart_generator.py \
  --csv "output/revenue_data.csv" \
  --metadata "output/pbi_report_metadata.xlsx" \
  --visual "Pipeline by Stage" \
  --output "output/charts/"

# Mode 2: CSV + Screenshot (agent-driven, PPTX output)
python skills/chart_generator.py \
  --csv "output/revenue_data.csv" \
  --visual-type barChart \
  --visual-name "Pipeline by Stage" \
  --field "Sales Stage:grouping" \
  --field "Opportunity Count:measure" \
  --output "output/charts/"

# Combo chart with secondary axis (auto PNG fallback on slide)
python skills/chart_generator.py \
  --csv "output/combo_data.csv" \
  --visual-type lineClusteredColumnComboChart \
  --visual-name "Revenue and Growth" \
  --field "Month:grouping" \
  --field "Revenue:measure" \
  --field "Growth Rate:y2" \
  --output "output/charts/"

# Legacy PNG mode (backward compatible)
python skills/chart_generator.py \
  --csv "output/revenue_data.csv" \
  --visual-type columnChart \
  --visual-name "Revenue by Region" \
  --field "Region:grouping" \
  --field "Revenue:measure" \
  --format png \
  --output "output/charts/"
```

### Orchestrator: Run Skills in Sequence
All three input skills (1, 3, 4) produce the same metadata Excel. Feed any of them into Skill 2. After executing the DAX queries, feed the tabular results into Skill 5 for chart generation.

```bash
# Path A: PBIP files (Skill 1 -> Skill 2 -> execute DAX -> Skill 5)
python skills/extract_metadata.py \
  --report-root "data/<ReportName>.Report/definition" \
  --model-root "data/<ReportName>.SemanticModel/definition" \
  --output "output/pbi_report_metadata.xlsx"
python skills/dax_query_builder.py "output/pbi_report_metadata.xlsx" "output/dax_queries.xlsx"
# User executes DAX queries against Fabric, saves results as CSV
python skills/chart_generator.py \
  --csv "output/visual_data.csv" \
  --metadata "output/pbi_report_metadata.xlsx" \
  --visual "Pipeline by Stage" \
  --output "output/charts/"

# Path B: CSV/Excel export (Skill 3 -> Skill 2 -> execute DAX -> Skill 5)
python skills/read_excel_export.py "data/export.csv" \
  --model-root "data/<ReportName>.SemanticModel/definition" \
  --output "output/pbi_report_metadata.xlsx"
python skills/dax_query_builder.py "output/pbi_report_metadata.xlsx" "output/dax_queries.xlsx"
# User executes DAX queries against Fabric, saves results as CSV
python skills/chart_generator.py \
  --csv "output/visual_data.csv" \
  --metadata "output/pbi_report_metadata.xlsx" \
  --visual "Sales by Region" \
  --output "output/charts/"

# Path C: Screenshot (Agent reads image -> Skill 4 -> Skill 2, or directly to Skill 5)
# Agent views screenshot, identifies: pieChart, "Sales by Region", fields Region (grouping) + Sales (measure)
python skills/analyze_screenshot.py \
  --visual-type pieChart \
  --visual-name "Sales by Region" \
  --field "Region:grouping" \
  --field "Sales:measure" \
  --model-root "data/<ReportName>.SemanticModel/definition" \
  --output "output/pbi_report_metadata.xlsx"
python skills/dax_query_builder.py "output/pbi_report_metadata.xlsx" "output/dax_queries.xlsx"
# Or skip straight to chart generation with CSV + screenshot info:
python skills/chart_generator.py \
  --csv "output/visual_data.csv" \
  --visual-type pieChart \
  --visual-name "Sales by Region" \
  --field "Region:grouping" \
  --field "Sales:measure" \
  --output "output/charts/"
```

## Test Data
- **Revenue Opportunities** report -- the reference PBIP report for Skill 1 validation (no bookmarks)
  - Report files: `data/Revenue Opportunities.Report/definition/`
  - Semantic model: `data/Revenue Opportunities.SemanticModel/definition/`
- **Store Sales** report -- PBIP report with 2 bookmarks (filter: Store Type = "New Store", visual show/hide toggle)
  - Report files: `data/Store Sales.Report/definition/`
  - Semantic model: `data/Store Sales.SemanticModel/definition/`
  - Bookmarks: `TSV Ribbon ON`, `LYS Combo ON` — both filter by Store Type, swap waterfall/column chart visibility
- **AI Sample** report -- PBIP report with 17 bookmarks (varied condition types: IN, NOT IN, date ranges, relative dates)
  - Report files: `data/Artificial Intelligence Sample (2).Report/definition/`
  - Semantic model: `data/Artificial Intelligence Sample (2).SemanticModel/definition/`
  - Most bookmarks reference deleted/hidden pages; `Last 12` and `Last 90` reference existing page
- **CSV test file:** `data/Total Sales Variance by FiscalMonth and District Manager.csv` -- sample PBI visual export for Skill 3 testing
- **Manual reference files:** `data/manual/` -- manually created reference files for the Revenue Opportunities report:
  - `data/manual/pbi_report_metadata_revopp.xlsx` -- manually verified metadata extraction (30 rows, 11 visuals)
  - `data/manual/dax_queries_by_visual.xlsx` -- manually written DAX queries per visual (11 queries, all validated)

### Full Test Run (Revenue Opportunities)
```bash
python skills/extract_metadata.py \
  --report-root "data/Revenue Opportunities.Report/definition" \
  --model-root "data/Revenue Opportunities.SemanticModel/definition" \
  --output "output/pbi_report_metadata.xlsx"

python skills/dax_query_builder.py "output/pbi_report_metadata.xlsx" "output/dax_queries_revenue_opportunities.xlsx"
```
Then compare against manual references:
- `output/pbi_report_metadata.xlsx` vs `data/manual/pbi_report_metadata_revopp.xlsx`
- `output/dax_queries_revenue_opportunities.xlsx` vs `data/manual/dax_queries_by_visual.xlsx`

## Critical Rules — NEVER BREAK THESE
1. **NEVER modify input PBIP files** — the `.Report/` and `.SemanticModel/` folders are read-only inputs. Never write to, rename, or delete anything inside them.
2. **ALWAYS save outputs to the `output/` folder** — never write output files to `data/` or the project root.
3. **Input data goes in `data/`** — PBIP report and semantic model folders belong under `data/`.
4. **NEVER modify original measure names** — measure names must match exactly as they appear in TMDL files.
5. **ALWAYS resolve nested measure dependencies recursively** — if Measure A references Measure B which references Column C, all three must appear in the metadata output.
6. **Circular measure references must not cause infinite loops** — the visited set in `resolve_measure_dependencies()` prevents this.
7. **Auto-generated visual-level filters that duplicate query state fields must be skipped** — prevents double-counting in metadata.
8. **Base DAX queries do not include filter values** — the main DAX Queries sheet only knows which fields are filtered, not the values (appear as comments). Bookmark-derived filter values are in the separate Bookmark DAX Queries sheet.
9. **NEVER guess filter values blindly** — when the user asks for a filtered DAX query (e.g., "give me the waterfall chart for only new stores"):
   - **Map to the correct field using context.** A person's name (e.g., "Espinoza Brynn") is a Buyer (`'Item'[Buyer]`), not a Store Name (`'Store'[Name]`). Read the visual's fields from the metadata to understand what each column represents before choosing the filter target.
   - **The pipeline has NO access to actual data values.** Metadata only contains field names and tables, not row-level data. The exact spelling, casing, and formatting of values is unknown (e.g., `"Espinoza, Brynn"` vs `"Espinoza Brynn"` — a wrong value silently returns empty results).
   - **Always caveat uncertain values.** Tell the user: "I'm using `"Espinoza Brynn"` but the actual value in the data might differ (e.g., comma-separated, different casing). Can you confirm the exact value?"
   - **Check available data exports first.** Look for CSV/Excel files in `data/` that might contain sample values for the relevant table/column before constructing the filter.
   - **When ambiguous, ask.** If the user says "filter by espinoza" and the visual has both Store Name and Buyer, ask which field they mean rather than guessing.

## Validation Status
The pipeline has been manually cross-checked against three reports:
- **Revenue Opportunities** — 11/11 visuals, 30/30 metadata rows. No bookmarks. Validated against manual reference files in `data/manual/`.
- **Store Sales** — 17/17 visuals across 5 pages. 2 bookmarks with `'Store'[Store Type] = "New Store"` filter. 8 bookmark DAX queries generated (4 per bookmark). Validated by running DAX queries against the Semantic Model.
- **AI Sample** — 10 visuals across 3 pages. 17 bookmarks (IN, NOT IN, date range, relative dates). Bookmarks referencing deleted pages produce filters but 0 matched visuals (expected).

## Validation Approach
1. Run the pipeline on the target report
2. For Revenue Opportunities: compare metadata output against `data/manual/pbi_report_metadata_revopp.xlsx` (expect 30/30 row match) and DAX queries against `data/manual/dax_queries_by_visual.xlsx` (expect 11/11 exact match)
3. For other reports: run the generated DAX queries against the Semantic Model and compare to what the PBI visual displays
4. For non-table visuals, convert to table view in PBI to see the raw data behind it
5. If the DAX query output matches what the visual displays, the query is correct

## Known Limitations
- **Base DAX queries do not include filter values** -- the main DAX Queries sheet captures which fields are filtered but not the actual values (shown as comments). When bookmarks are present, the Bookmark DAX Queries sheet provides `CALCULATETABLE`-wrapped queries with actual filter values applied.
- **Bookmark filters only** -- filter values are extracted from bookmarks, not from the visual's own persisted filter state. If a report has no bookmarks, no filter values are available.
- **Relative date offsets** (e.g., `-6L` months back) cannot be resolved statically and appear as comments in DAX.
- Complex visuals with calculated columns, nested measures, or unusual aggregations may produce queries that don't perfectly match Power BI's internal rendering
- Implicit measures (auto-generated Sum/Count from drag-and-drop) are not tracked in TMDL files
- Cross-table measure dependencies may not be fully caught
- HierarchyLevel fields (date hierarchies) use fallback resolution via PropertyVariationSource
- **Skill 3 (CSV reader):** Field classification relies on value heuristics (currency/percentage/numeric patterns). Ambiguous columns (e.g., IDs that look numeric) may be misclassified. Use `--model-root` for accurate matching.
- **Skill 4 (screenshots):** The agent identifies fields visually — accuracy depends on image clarity and whether field names are visible in the chart (axis labels, headers, legend). Always verify detected fields.
- **Skill 3/4 without model:** Unmatched fields use placeholder `'<Table>'[FieldName]` which must be manually corrected before DAX queries can execute.
- **Skill 5 (chart generator):** Charts are PBI-styled approximations, not pixel-perfect replicas. PPTX mode produces native editable charts for bar, column, line, area, pie, donut, and scatter types; all other types fall back to a plotly PNG image embedded on the slide. Combo charts, waterfall, funnel, treemap, gauge, card, KPI, and tables use PNG fallback because python-pptx lacks native support for these chart types. Combo chart bar/line split defaults to "last measure = line" when no Visual Y2 metadata is available. Map visuals, slicers, and AI visuals are skipped (not meaningful as static images). Ribbon charts are rendered as stacked area (closest plotly equivalent).

## Chat Presentation Rules
- **Always detect and apply ALL filters** — report-level, page-level, and visual-level — when presenting a DAX query in chat. Check the Filter Expressions sheet for any filter that applies to the visual (by scope: report filters apply to all visuals, page filters apply to all visuals on that page, visual filters apply to that specific visual).
- **Respect filter hierarchy when building the filtered expression.** Power BI applies filters in this order: Report → Page → Visual. Inner filters override outer filters when they target the same column. When wrapping with CALCULATETABLE/CALCULATE, list filters from outermost to innermost scope so the DAX engine resolves conflicts the same way PBI does (visual-level wins over page-level, page-level wins over report-level). If a visual-level filter targets the same column as a report-level filter, only include the visual-level filter (it overrides).
- **Always show both versions:** first the **filtered query** (with all applicable filters wrapped via CALCULATETABLE/CALCULATE), then the **unfiltered (base) query** below it. The filtered version is the primary output since it reflects what the user actually sees in the report.

## Coding Conventions
- Use clear variable names (no single letters except loop counters)
- Add inline comments explaining regex patterns
- All file I/O uses UTF-8 with BOM handling (`encoding="utf-8-sig"`)
- Each skill must work both standalone (`if __name__ == "__main__"`) and as an importable module
- Log warnings for unresolved items (don't silently drop data)
