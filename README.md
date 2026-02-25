# PBI DAX Query Generation Agent

An agent that reverse-engineers Power BI visuals into executable DAX queries and generates PBI-styled PowerPoint chart slides from the results. Built as the data retrieval layer for Lara's AI-Powered Slide Generation project at XP3R.

## What It Does

Give it a Power BI report and it walks you through it interactively — page by page, visual by visual — delivering the exact DAX to reproduce each visual's data, plus optional chart generation as editable PowerPoint slides.

```
  .pbix file ──► [Skill 0] pbix_extractor.py ──► PBIP folders
                                                      │
  PBIP files ───► [Skill 1] extract_metadata.py ──────┴──► Metadata Excel ──► [Skill 2] dax_query_builder.py ──► DAX queries
                                                                │
                                                                └──► [Skill 3] chart_generator.py ──► .pptx / .png
```

## Pipeline Steps

| Step | Script | Input | Output |
|------|--------|-------|--------|
| **Skill 0** | `pbix_extractor.py` | `.pbix` file | PBIP folder structure |
| **Skill 1** | `extract_metadata.py` | PBIP report + semantic model folders | Metadata Excel (visuals, fields, filters, measures, bookmarks) |
| **Skill 2** | `dax_query_builder.py` | Metadata Excel | DAX queries per visual + bookmark-filtered variants |
| **Skill 3** | `chart_generator.py` | CSV data + metadata or manual spec | Single-slide `.pptx` (native editable chart) or `.png` |

## Input Options

- **`.pbix` file** — Skill 0 extracts the PBIP structure automatically. Semantic model extraction requires the optional `pbixray` package (needs a C compiler).
- **PBIP folders** — Point directly to the `.Report/definition/` and `.SemanticModel/definition/` folders (most complete, no extra dependencies).
- **Sample reports** — Three validated reports included in `data/`:

| Report | Visuals | Bookmarks | Notes |
|--------|---------|-----------|-------|
| Revenue Opportunities | 11 | 0 | Reference report with manual validation files |
| Store Sales | 17 | 2 | Store Type filter + visual show/hide |
| AI Sample | 10 | 17 | IN, NOT IN, date ranges, relative dates |

## Interactive Session Flow

The agent follows a 6-step interactive loop:

1. **Load the report** — provide a `.pbix`, PBIP paths, or a sample report name
2. **See pages** — numbered list of all pages in the report
3. **Pick a visual** — list of visuals on the selected page with type and key fields
4. **Get DAX queries** — three outputs per visual:
   - **Filtered query** — with all report/page/visual-level filters applied
   - **Base query** — unfiltered, for exploring the full dataset
   - **Custom filter offer** — list of filterable fields; provide a value and get a wrapped query
5. **Generate a chart** (optional) — provide CSV data from an executed DAX query to render a PBI-styled chart as a `.pptx` slide or `.png` image
6. **Continue** — pick another visual, switch pages, generate a chart, or load a new report

## Chart Generation

Skill 3 renders PBI-styled chart visuals in two output formats:

- **PPTX (default)** — single-slide PowerPoint with a native editable chart (bar, column, line, area, pie, donut, scatter) or embedded PNG fallback for complex types
- **PNG** — static plotly-rendered image

**Supported chart types:**

| Category | Types | Output |
|----------|-------|--------|
| Native PPTX | bar, column, stacked bar/column, 100% stacked, line, area, stacked area, pie, donut, scatter | Editable chart |
| PNG fallback | combo (dual-axis), waterfall, funnel, treemap, gauge, card, KPI, table, ribbon | Embedded image |
| Skipped | slicer, map, AI visuals | Not rendered |

**Two input modes:**
```bash
# Mode 1: Metadata-driven (auto-detects visual type + fields from Skill 1 Excel)
python skills/chart_generator.py \
    --csv "output/data.csv" \
    --metadata "output/pbi_report_metadata.xlsx" \
    --visual "Pipeline by Stage" \
    --format pptx --output "output/charts/"

# Mode 2: Manual (specify visual type + fields directly)
python skills/chart_generator.py \
    --csv "output/data.csv" \
    --visual-type barChart \
    --field "Sales Stage:grouping" --field "Revenue:measure" \
    --format pptx --output "output/charts/"
```

## DAX Patterns

| Pattern | When Used | DAX Template |
|---------|-----------|--------------|
| Pattern 1 (Single Measure) | Cards with one measure | `EVALUATE { [Measure] }` |
| Pattern 1 (Multiple Measures) | Cards/KPIs with multiple measures | `EVALUATE ROW("Name", [Measure], ...)` |
| Pattern 2 (Columns Only) | Slicers, column-only visuals | `EVALUATE VALUES(...)` or `EVALUATE DISTINCT(SELECTCOLUMNS(...))` |
| Pattern 3 (Columns + Measures) | Most chart/table visuals | `EVALUATE SUMMARIZECOLUMNS(...)` |

Bookmark-filtered queries wrap the base DAX with `CALCULATETABLE` / `CALCULATE` and apply all bookmark filter conditions (IN, NOT IN, comparisons, date ranges).

## Project Structure

```
powerpointTask/
├── CLAUDE.md                       # Agent instructions and skill documentation
├── README.md                       # This file
├── requirements.txt                # Python dependencies
├── skills/
│   ├── pbix_extractor.py           # Skill 0: .pbix → PBIP folder converter
│   ├── tmdl_parser.py              # Shared: TMDL semantic model parser
│   ├── bookmark_parser.py          # Shared: Bookmark filter parsing + DAX conversion
│   ├── extract_metadata.py         # Skill 1: PBIP metadata extraction
│   ├── dax_query_builder.py        # Skill 2: DAX query generation
│   └── chart_generator.py          # Skill 3: Chart generator (plotly + python-pptx)
├── .claude/skills/                 # Claude Code skill definitions
│   ├── extract-metadata/
│   ├── build-dax-queries/
│   ├── generate-charts/
│   └── run-dax-pipeline/
├── data/                           # Input reports (.pbix, .Report/, .SemanticModel/)
└── output/                         # All generated outputs (metadata, DAX, charts)
```

## Stack

- **Python 3.x** with pandas, openpyxl, regex
- **plotly + kaleido** — chart rendering (PNG mode and PPTX fallback)
- **python-pptx** — native PowerPoint chart generation
- **pbixray** (optional) — `.pbix` semantic model extraction

## Quick Start

```bash
pip install -r requirements.txt

# From .pbix file
python skills/pbix_extractor.py "report.pbix" --output "data/"

# Extract metadata
python skills/extract_metadata.py \
    --report-root "data/Report.Report/definition" \
    --model-root "data/Report.SemanticModel/definition" \
    --output "output/metadata.xlsx"

# Generate DAX queries
python skills/dax_query_builder.py "output/metadata.xlsx" "output/dax_queries.xlsx"

# Generate a chart (after executing a DAX query to CSV)
python skills/chart_generator.py \
    --csv "output/visual_data.csv" \
    --metadata "output/metadata.xlsx" \
    --visual "Revenue by Region" \
    --output "output/charts/"
```

## Validation

The pipeline has been manually cross-checked against three reports:
- **Revenue Opportunities** — 11/11 visuals, 30/30 metadata rows, validated against manual reference files
- **Store Sales** — 17/17 visuals, 2 bookmarks, 8 bookmark DAX queries, validated by running DAX against the Semantic Model
- **AI Sample** — 10 visuals, 17 bookmarks, bookmarks referencing deleted pages produce filters but 0 matched visuals (expected)

## Known Limitations

- `.pbix` semantic model extraction requires `pbixray` (which needs a C compiler) — without it, report structure is fully extracted but measure formulas are missing
- Relative date offsets (e.g., "last 6 months") cannot be resolved statically
- Charts are PBI-styled approximations, not pixel-perfect replicas
