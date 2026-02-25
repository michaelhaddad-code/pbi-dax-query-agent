---
name: generate-charts
description: Generate PBI-style chart visuals as PowerPoint slides or PNG images from DAX query data. Use when the user asks to generate a chart, create a slide, render a visual, or produce a PowerPoint from CSV data. Also use when the user mentions chart generation, PPTX output, or visual rendering.
---

# Generate Charts from DAX Query Data

## What This Skill Does
Takes tabular CSV data (typically from executed DAX queries) and renders PBI-styled chart visuals. Supports two output formats: **PPTX** (default — single-slide PowerPoint with native editable chart or PNG fallback) and **PNG** (legacy plotly image). This is Step 3 of the pipeline — it consumes the data retrieved by executing the DAX queries from Step 2.

## When to Use
- User asks to "generate a chart", "create a slide", or "render a visual"
- User has CSV data from an executed DAX query and wants a visual
- User wants a PowerPoint slide for a specific visual
- This is Step 3 after extracting metadata (Step 1) and generating DAX queries (Step 2)

## How to Run

### Mode 1: Metadata-driven (preferred — reads visual spec from Skill 1 Excel):
```bash
python skills/chart_generator.py \
    --csv "output/revenue_data.csv" \
    --metadata "output/pbi_report_metadata.xlsx" \
    --visual "Pipeline by Stage" \
    --format pptx \
    --output "output/charts/"
```

### Mode 2: Manual CLI-driven (ad-hoc — user specifies visual type + fields):
```bash
python skills/chart_generator.py \
    --csv "output/revenue_data.csv" \
    --visual-type barChart \
    --field "Sales Stage:grouping" \
    --field "Revenue:measure" \
    --visual-name "Pipeline by Stage" \
    --format pptx \
    --output "output/charts/"
```

### As an import (preferred when chaining programmatically):
```python
import sys
sys.path.insert(0, "skills")
import pandas as pd
from chart_generator import (
    generate_chart_pptx, save_chart_pptx,
    generate_chart, save_chart,
    parse_visual_from_metadata, VisualSpec
)

df = pd.read_csv("output/revenue_data.csv", encoding="utf-8-sig")

# Mode 1: From metadata Excel
spec = parse_visual_from_metadata("output/pbi_report_metadata.xlsx", "Pipeline by Stage")
prs = generate_chart_pptx(df, spec=spec)
save_chart_pptx(prs, "output/charts/Pipeline_by_Stage.pptx")

# Mode 2: Manual spec
spec = VisualSpec(
    page_name="Overview",
    visual_name="Pipeline by Stage",
    visual_type="barChart",
    grouping_columns=["Sales Stage"],
    measure_columns=["Revenue"],
    y2_columns=[],
    dax_pattern="Pattern 3"
)
prs = generate_chart_pptx(df, spec=spec)
save_chart_pptx(prs, "output/charts/Pipeline_by_Stage.pptx")
```

## Required Inputs
1. **CSV data file** (`--csv`) — tabular data, typically from an executed DAX query
2. **Visual specification** — one of:
   - `--metadata` Excel (from Skill 1) + `--visual` name — auto-reads visual type, grouping/measure columns
   - `--visual-type` + `--field` entries in `name:role` format (roles: `measure`, `grouping`, `y2`)

## Optional Parameters
- `--format` — `pptx` (default) or `png`
- `--output` — output directory (default: `output/charts/`)
- `--report-name` — creates a subfolder per report (e.g., `output/charts/Revenue_Opportunities/`)
- `--visual-name` — display name for Mode 2 (defaults to visual type)
- `--width` / `--height` / `--scale` — PNG mode dimensions (default: 1100x500, scale=2)

## Output
- **PPTX mode (default):** Single-slide `.pptx` file with a native editable PowerPoint chart (bar, column, line, area, pie, donut, scatter) or an embedded PNG fallback for complex types
- **PNG mode:** Static plotly-rendered `.png` image

## Supported Chart Types

### Native PPTX charts (editable in PowerPoint):
barChart, clusteredBarChart, stackedBarChart, hundredPercentStackedBarChart, columnChart, clusteredColumnChart, stackedColumnChart, hundredPercentStackedColumnChart, lineChart, areaChart, stackedAreaChart, pieChart, donutChart, scatterChart

### PNG fallback (embedded image on slide):
lineClusteredColumnComboChart, lineStackedColumnComboChart, waterfallChart, funnelChart, treemap, gauge, card, multiRowCard, kpi, tableEx, pivotTable, ribbonChart

### Skipped (not meaningful as static charts):
slicer, advancedSlicerVisual, map, filledMap, shapeMap, decompositionTreeVisual, keyDriversVisual, qnaVisual

## Validation
- Check that the output file exists and is non-empty
- For PPTX: open in PowerPoint and verify the chart is editable (native types) or visible (PNG fallback)
- For PNG: verify the image renders correctly with proper PBI styling (colors, fonts, layout)
- Compare against the original PBI visual for data accuracy

## Known Limitations
- Charts are PBI-styled approximations, not pixel-perfect replicas of Power BI visuals
- Combo charts, waterfall, funnel, treemap, gauge, card, KPI, and tables use PNG fallback on slide (not natively editable)
- Slicers, maps, and AI visuals are skipped entirely
- Requires `plotly`, `kaleido`, `python-pptx`, `pandas`, `openpyxl` packages
- Empty DataFrames produce no output (skipped with a message)

## Critical Rules
- ALWAYS save output to the `output/` folder (never to `data/` or project root)
- NEVER modify the input CSV or metadata Excel files
- Field roles must match the visual type — grouping columns are categories/axis, measure columns are values
- For combo charts, `y2` fields go on the secondary axis
