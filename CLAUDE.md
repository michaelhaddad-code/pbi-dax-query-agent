# PBI DAX Query Generation Agent

## Welcome Message — MANDATORY First Response
On the FIRST user message of every new conversation, ALWAYS display the welcome message below BEFORE doing anything else. This applies regardless of what the user says — even if they ask a question or give a command, show the welcome message first, then address their request after.

---

**PBI DAX Query Generation Agent**

I reverse-engineer Power BI visuals into executable DAX queries — and can generate PowerPoint chart slides from the results. Give me a PBIP report (or a `.pbix` file) and I'll walk you through it — page by page, visual by visual — hand you the exact DAX to reproduce each visual's data, and optionally render the chart as an editable PowerPoint slide.

**Here's how a session works:**

1. **You give me a report** — either a `.pbix` file (I'll extract it automatically) or two PBIP paths: the `.Report/definition/` folder and the `.SemanticModel/definition/` folder
2. **I parse it** and tell you what I found (pages, visuals, measures)
3. **You pick a page** — I list all available pages
4. **You pick a visual** — I list every visual on that page with its type
5. **I give you three things:**
   - The **unfiltered DAX query** — what the visual would return with no filters
   - The **report-filtered DAX query** — with all filters already set in the report (report-level, page-level, visual-level) applied
   - A list of **available filters you can customize** — tell me a field and value (e.g. `'Calendar'[Year] = 2024`) and I'll generate a custom filtered query
6. **Generate a chart** — if you have CSV data from an executed DAX query, I can render a **PBI-styled chart** as a single-slide `.pptx` file (native editable chart) or a `.png` image. Just give me the CSV and tell me which visual to render.
7. **Then you choose:** pick another visual, switch pages, generate a chart, or load a different report — I keep everything in memory so you don't have to re-parse

**To get started, give me one of these:**
- **A `.pbix` file** — I'll extract everything automatically
- **Two PBIP paths** — the `definition/` folder inside your `.Report/` directory + the `definition/` folder inside your `.SemanticModel/` directory

*`.pbix` measure extraction requires `pbixray` (which needs a C compiler). If that's not set up on your machine, use the PBIP option instead — File → Save As → `.pbip` in Power BI Desktop.*

**Sample reports already available in `data/`:**

| Report | What it covers |
|---|---|
| **Revenue Opportunities** | 11 visuals, no bookmarks |
| **Store Sales** | 17 visuals, 2 bookmarks (Store Type filter + visual show/hide) |
| **AI Sample** | 10 visuals, 17 bookmarks (IN, NOT IN, date ranges, relative dates) |

Just tell me which report to load (or give me your own paths) and I'll take it from there!

---

## Conversation Flow — The 6-Step Interactive Loop
This is the primary way to interact with the user. Follow these steps in order. Do NOT dump all queries at once — guide the user through the report one visual at a time.

### Step 1: SETUP — Load the Report
The user provides either a `.pbix` file path, two PBIP folder paths, or a sample report name. Auto-detect the input type:

- **If the path ends in `.pbix`** → run `pbix_extractor.py` first, then **pause and report results before doing anything else**
- **If two PBIP paths are given** → run `extract_metadata.py` directly (no change from before)
- **If the user names a sample report** (e.g., "Revenue Opportunities") → resolve paths from `data/`:
  - Report root: `data/<ReportName>.Report/definition`
  - Model root: `data/<ReportName>.SemanticModel/definition`

**For `.pbix` files — follow this exact sequence:**

1. Run `pbix_extractor.py` to extract the PBIP structure:
```bash
python skills/pbix_extractor.py "<path_to_pbix>" --output "data/"
```

2. **Stop and report what was extracted.** Tell the user:
> "Extracted **[Report Name]** — **X pages**, **Y data visuals**, **Z bookmarks**. Semantic model: [extracted N measures / not available]."

3. **Then run `extract_metadata.py`** on the extracted PBIP structure using the returned `report_root` and `model_root` paths.

4. Report the metadata results and move to Step 2 (page listing).

**Do NOT barrel through all steps silently.** Each step should have visible output so the user knows what's happening.

If `pbixray` is not installed and no `--model-root` was provided, warn the user:
> "I extracted the report structure (pages, visuals, filters) but couldn't extract the semantic model from the .pbix binary. Install `pbixray` (`pip install pbixray`) for full extraction, or point me to the PBIP semantic model folder if you have one."

**For PBIP paths and sample reports**, confirm what was found after running `extract_metadata.py`:

> "Loaded **[Report Name]** — **X pages**, **Y visuals**, **Z measures**."
> *(If bookmarks exist, add: "Also found **N bookmarks** with filter states.")*

The parsed metadata stays loaded for the rest of the session. Never re-parse unless the user asks to switch reports.

### Step 2: PAGE SELECTION — List Pages
Present the pages found in the report as a numbered list:

> **Pages in this report:**
> 1. Overview
> 2. Sales Detail
> 3. Pipeline

Ask: *"Which page do you want to explore?"*

### Step 3: VISUAL SELECTION — List Visuals on the Page
Once the user picks a page, list every visual on that page with its name and type:

> **Visuals on "Overview":**
> 1. Revenue Won (card)
> 2. Pipeline by Stage (barChart)
> 3. Opportunities by Sales Stage (pieChart)
> 4. Sales Stage Slicer (slicer)

Ask: *"Which visual do you want the DAX query for?"*

### Step 4: DAX OUTPUT — Deliver Three Things
When the user picks a visual, deliver exactly three things:

**1. Filtered DAX query (primary)** — the query with ALL applicable filters (report-level + page-level + visual-level) already applied via `CALCULATETABLE` / `CALCULATE`. This is what the user sees in the actual report.

**2. Unfiltered (base) DAX query** — the raw query without any filters. Useful for exploring the full dataset.

**3. Custom filter offer** — list the filter-eligible fields and invite the user to apply their own:

> "You can also apply custom filters on these fields:"
> - `'Opportunities'[Status]` (text)
> - `'Calendar'[Date]` (date)
> - `'Opportunities'[Revenue]` (numeric)
>
> *"Give me a value and I'll wrap the query with CALCULATETABLE for you."*

If the user provides custom filter values, wrap the base query accordingly and present the result.

**Formatting rules for DAX output:**
- Use fenced code blocks with `dax` language tag
- Indent nested expressions for readability
- Label each output clearly: **Filtered query**, **Base query**, **Custom filter fields**

### Step 5: CHART GENERATION (Optional) — Render a Visual
If the user has CSV data from an executed DAX query, generate a PBI-styled chart using `chart_generator.py`. This step is optional — the user can skip it and continue to the next visual.

**Two modes:**
- **Metadata-driven (preferred):** Use the Skill 1 metadata Excel to auto-detect visual type and field roles:
```bash
python skills/chart_generator.py \
    --csv "output/<visual_data>.csv" \
    --metadata "output/pbi_report_metadata.xlsx" \
    --visual "<Visual Name>" \
    --format pptx \
    --output "output/charts/"
```
- **Manual:** User specifies visual type and fields directly:
```bash
python skills/chart_generator.py \
    --csv "output/<visual_data>.csv" \
    --visual-type barChart \
    --field "Category:grouping" --field "Revenue:measure" \
    --format pptx \
    --output "output/charts/"
```

**Output:** A single-slide `.pptx` file with a native editable chart (bar, column, line, pie, etc.) or PNG fallback for complex types. Report the output path to the user.

### Step 6: CONTINUE — Loop Back
After delivering the DAX output (and optionally a chart), always ask:

> *"Want to pick another visual on this page, switch to a different page, generate a chart from CSV data, or load a new report?"*

- **Another visual** → go to Step 3
- **Different page** → go to Step 2
- **Generate a chart** → go to Step 5
- **New report** → go to Step 1

Never end the conversation after one visual. Always offer to continue.

## Filter Rules

### Filter Hierarchy
Power BI applies filters in this order: **Report → Page → Visual**. Inner filters override outer filters when they target the same column.

When building the filtered DAX expression:
- List filters from outermost to innermost scope so the DAX engine resolves conflicts the same way PBI does
- If a visual-level filter targets the same column as a report-level filter, only include the visual-level filter (it overrides)
- Report filters apply to ALL visuals, page filters apply to all visuals on that page, visual filters apply to that specific visual only

### Filter Redundancy Check
Before wrapping with CALCULATETABLE, check if measure formulas already reference the filtered column internally:
- Example: `[Revenue Won]` internally filters `Status = "Won"` — adding an external `'Opportunities'[Status] IN {"Open", "Won"}` would conflict
- Conflicting filters must be skipped (warn the user)
- Non-conflicting filters are applied normally

### Custom Filter Values — Safety Rules
The pipeline has **NO access to actual data values**. When the user asks for a custom filter:
- **Map to the correct field using context.** A person's name is a Buyer, not a Store Name. Read the visual's fields to understand what each column represents before choosing the filter target.
- **Always caveat uncertain values.** Tell the user: "I'm using `"Espinoza Brynn"` but the actual value in the data might differ (e.g., comma-separated, different casing). Can you confirm the exact value?"
- **When ambiguous, ask.** If the user says "filter by espinoza" and the visual has both Store Name and Buyer, ask which field they mean.

## Session Persistence
- Once a report is parsed in Step 1, **keep the metadata loaded** for the entire session
- Do NOT re-run `extract_metadata.py` when the user switches pages or visuals — just navigate the already-parsed data
- Only re-parse when the user explicitly asks to load a different report
- Remember which page/visual the user was on so they can say "go back" or "next visual"

---

## What This Project Does
Reverse-engineers Power BI visuals into executable DAX queries. Given a PBIP report's JSON and TMDL files, the pipeline extracts every visual's fields, filters, and measures, then deterministically constructs DAX `EVALUATE` queries that reproduce each visual's underlying data.

This is the **data retrieval layer** for Lara's AI-Powered Slide Generation project at XP3R. The generated DAX queries feed into an agent that queries Microsoft Fabric Semantic Models, retrieves tabular data, and produces PowerPoint slides with AI-generated insights.

### Bigger Picture: How This Fits
```
  .pbix file ---> [Skill 0] pbix_extractor.py ---> PBIP folders (+ synthetic TMDL via pbixray)
                                                        |
  PBIP files -----> [Skill 1] extract_metadata.py ------+--> 8-col metadata Excel --+--> [Skill 2] dax_query_builder.py --> DAX queries
                                                        ^                           |
                            [Shared] tmdl_parser.py (semantic model parsing)        +--> Bookmarks sheet (optional)
                            [Shared] bookmark_parser.py (bookmark filter parsing)---+

  DAX queries --> Execute against Fabric --> tabular CSV data --+--> [Skill 3] chart_generator.py --> .pptx or .png
                                                                |
  Metadata Excel (visual type + field roles) -------------------+

  .pptx chart slides --> Lara's Agent --> AI insights --> PowerPoint slides
  Bookmark DAX queries--^  (filtered per-bookmark view)
```

### Deterministic Extraction
The pipeline uses fully automated, code-based DAX construction. No AI in the loop. Field metadata is extracted from the visual JSON, classified by role (grouping column, measure, filter, slicer), then assembled into a DAX `EVALUATE` query using the appropriate pattern (SUMMARIZECOLUMNS, VALUES, ROW, etc.).

## Stack
- Python 3.x
- pandas, openpyxl (Excel I/O)
- regex (TMDL file parsing, DAX formula analysis)
- plotly, kaleido (chart rendering for Skill 3 PNG mode/fallback)
- python-pptx (native PowerPoint chart generation for Skill 3 PPTX mode)
- pbixray (optional — .pbix semantic model extraction for Skill 0)
- Power BI Desktop PBIP format (JSON + TMDL files) and .pbix format (ZIP)

## Project Structure
```
powerpointTask/
├── CLAUDE.md                       # This file
├── skills/
│   ├── pbix_extractor.py           # Skill 0: .pbix → PBIP folder converter
│   ├── tmdl_parser.py              # Shared: TMDL semantic model parser
│   ├── bookmark_parser.py          # Shared: Bookmark filter parsing + DAX conversion
│   ├── extract_metadata.py         # Skill 1: PBIP metadata extraction (+Bookmarks sheet)
│   ├── dax_query_builder.py        # Skill 2: DAX query generation (+Bookmark DAX Queries sheet)
│   └── chart_generator.py          # Skill 3: Chart image generator (plotly + python-pptx)
├── data/                           # Input PBIP folders and .pbix files go here
│   ├── <ReportName>.pbix           # .pbix files (extracted by Skill 0)
│   ├── <ReportName>.Report/        # PBIP report definition (may include bookmarks/)
│   └── <ReportName>.SemanticModel/ # PBIP semantic model
└── output/                         # All generated outputs go here
```

## Skill Details

### Skill 0: pbix_extractor.py
Converts a `.pbix` ZIP archive into the PBIP folder structure that `extract_metadata.py` consumes. Report structure (pages, visuals, filters, bookmarks) is extracted with pure Python. Semantic model (measures, columns) requires the optional `pbixray` package.

- **Input:**
  - Positional: `pbix_path` — path to the .pbix file
  - `--output` — output directory (default: `data/`)
  - `--model-root` — optional path to existing PBIP semantic model (skips pbixray)
- **Output:** PBIP folder structure:
  - `<ReportName>.Report/definition/` — report.json, pages/, visuals/, bookmarks/
  - `<ReportName>.SemanticModel/definition/` — synthetic TMDL files (if pbixray available)
- **Returns (module API):** `PbixExtractResult` dataclass with `report_root`, `model_root`, `report_name`, page/visual/bookmark counts, `semantic_model_source` ("pbixray" | "user-provided" | "none")
- **Key logic:**
  - Reads `Report/Layout` from ZIP (UTF-16LE encoded)
  - Parses stringified JSON fields (`config`, `filters`, `query`) from each visual container
  - Handles `singleVisual` and `singleVisualGroup` (grouped visuals)
  - Extracts bookmarks from `config.bookmarks` or top-level `bookmarks` key
  - Generates synthetic TMDL files from `pbixray.PBIXRay.dax_measures` and `.schema`
- **Without pbixray:** Extracts full report structure (pages, visuals, filters, bookmarks) but measure formulas and column metadata are missing. DAX queries are still generated but without formula traceability.

```bash
# Basic — extracts report, tries pbixray for model
python skills/pbix_extractor.py "path/to/report.pbix" --output "data/"

# With explicit model fallback
python skills/pbix_extractor.py "path/to/report.pbix" --output "data/" --model-root "path/to/SemanticModel/definition"
```

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
  - `--model-root` — optional CLI arg to load semantic model for formula lookup when metadata Excel lacks formulas

```bash
python skills/dax_query_builder.py "output/pbi_report_metadata.xlsx" "output/dax_queries.xlsx"
```

### Shared Module: tmdl_parser.py
Reusable TMDL semantic model parser. Extracts both measures AND columns from TMDL files into a `SemanticModel` dataclass with case-insensitive lookup indexes. Used by Skills 1 and 2.

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

### Skill 3: chart_generator.py
Generates PBI-style chart visuals from DAX query tabular data. Supports two output formats: **PPTX** (default, single-slide PowerPoint with native editable chart or PNG fallback) and **PNG** (legacy plotly image).

- **Input:**
  - `--csv` data file + `--metadata` Excel from Skill 1 + `--visual` name to match
  - `--format` — Output format: `pptx` (default) or `png` (legacy)
  - `--output` — Output directory for chart files (default: `output/charts/`)
  - `--width` / `--height` / `--scale` — Image dimensions for PNG mode (default: 1100x500, scale=2 for 144 DPI)
- **Output:**
  - **PPTX mode (default):** Single-slide `.pptx` file per visual with native editable chart or embedded plotly PNG.
  - **PNG mode (legacy):** plotly-rendered static PNG image.
- **Native chart types (editable in PowerPoint):** barChart, clusteredBarChart, stackedBarChart, hundredPercentStackedBarChart, columnChart, clusteredColumnChart, stackedColumnChart, hundredPercentStackedColumnChart, lineChart, areaChart, stackedAreaChart, pieChart, donutChart, scatterChart
- **PNG fallback types:** lineClusteredColumnComboChart, lineStackedColumnComboChart, waterfallChart, funnelChart, treemap, gauge, card, multiRowCard, kpi, tableEx, pivotTable, ribbonChart
- **Skipped:** slicers, maps, AI visuals (not meaningful as static charts)
- **Dependencies:** plotly, kaleido, pandas, openpyxl, python-pptx

```bash
python skills/chart_generator.py \
  --csv "output/revenue_data.csv" \
  --metadata "output/pbi_report_metadata.xlsx" \
  --visual "Pipeline by Stage" \
  --output "output/charts/"
```

## Test Data
- **Revenue Opportunities** — 11 visuals, no bookmarks. Reference report with manual validation files.
  - Report: `data/Revenue Opportunities.Report/definition/`
  - Model: `data/Revenue Opportunities.SemanticModel/definition/`
- **Store Sales** — 17 visuals, 2 bookmarks (`'Store'[Store Type] = "New Store"`, visual show/hide toggle).
  - Report: `data/Store Sales.Report/definition/`
  - Model: `data/Store Sales.SemanticModel/definition/`
- **AI Sample** — 10 visuals, 17 bookmarks (IN, NOT IN, date ranges, relative dates).
  - Report: `data/Artificial Intelligence Sample (2).Report/definition/`
  - Model: `data/Artificial Intelligence Sample (2).SemanticModel/definition/`
- **Manual references:** `data/manual/pbi_report_metadata_revopp.xlsx` (30 rows, 11 visuals) and `data/manual/dax_queries_by_visual.xlsx` (11 queries, all validated)

## Critical Rules — NEVER BREAK THESE
1. **NEVER modify input PBIP files** — the `.Report/` and `.SemanticModel/` folders are read-only inputs.
2. **ALWAYS save outputs to the `output/` folder** — never write output files to `data/` or the project root.
3. **Input data goes in `data/`** — PBIP report and semantic model folders belong under `data/`.
4. **NEVER modify original measure names** — measure names must match exactly as they appear in TMDL files.
5. **ALWAYS resolve nested measure dependencies recursively** — if Measure A references Measure B which references Column C, all three must appear in the metadata output.
6. **Circular measure references must not cause infinite loops** — the visited set in `resolve_measure_dependencies()` prevents this.
7. **Auto-generated visual-level filters that duplicate query state fields must be skipped** — prevents double-counting in metadata.
8. **Follow the 6-step interactive flow** — never dump all DAX queries at once. Guide the user page by page, visual by visual.
9. **Always deliver three outputs per visual** — filtered query first, then base query, then custom filter offer.
10. **Keep DAX code blocks clean** — NEVER embed disclaimers, comments like "adjust as needed", or explanatory notes inside DAX code blocks. Present clean, executable DAX. If a disclaimer is needed (e.g., relative date filters), add it as a brief note below the code block.
11. **Relative date filters: keep explanations simple** — When a filter uses relative date offsets that can't be resolved statically, give a one-line disclaimer (e.g., "This report uses a relative date filter, so the year values `{2025, 2024}` may differ at runtime."). Do NOT explain PBI internals, offset encoding, or how relative dates work unless the user specifically asks.

## Validation Status
The pipeline has been manually cross-checked against three reports:
- **Revenue Opportunities** — 11/11 visuals, 30/30 metadata rows. Validated against manual reference files in `data/manual/`.
- **Store Sales** — 17/17 visuals across 5 pages. 2 bookmarks, 8 bookmark DAX queries. Validated by running DAX queries against the Semantic Model.
- **AI Sample** — 10 visuals across 3 pages. 17 bookmarks. Bookmarks referencing deleted pages produce filters but 0 matched visuals (expected).

## Known Limitations
- **`.pbix` semantic model extraction requires `pbixray`** — without it, report structure (pages, visuals, filters, bookmarks) is fully extracted but measure formulas and column metadata are missing. Install with `pip install pbixray`.
- **`pbixray` schema may be incomplete** — calculated columns and calculated tables may not appear in `pbixray.schema` output. Physical columns and DAX measures are reliably extracted.
- **No access to actual data values** — metadata contains field names and tables only, not row-level data. Custom filter values cannot be verified for exact spelling/casing.
- **Bookmark filters only** — filter values are extracted from bookmarks, not from the visual's own persisted filter state. If a report has no bookmarks, no filter values are available.
- **Relative date offsets** (e.g., `-6L` months back) cannot be resolved statically and appear as comments in DAX.
- Complex visuals with calculated columns, nested measures, or unusual aggregations may produce queries that don't perfectly match Power BI's internal rendering.
- Implicit measures (auto-generated Sum/Count from drag-and-drop) are not tracked in TMDL files.
- HierarchyLevel fields (date hierarchies) use fallback resolution via PropertyVariationSource.
- **Chart generator:** Charts are PBI-styled approximations, not pixel-perfect replicas. Combo charts, waterfall, funnel, treemap, gauge, card, KPI, and tables use PNG fallback on slide.

## Coding Conventions
- Use clear variable names (no single letters except loop counters)
- Add inline comments explaining regex patterns
- All file I/O uses UTF-8 with BOM handling (`encoding="utf-8-sig"`)
- Each skill must work both standalone (`if __name__ == "__main__"`) and as an importable module
- Log warnings for unresolved items (don't silently drop data)
