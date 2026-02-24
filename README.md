# PBI DAX Query Generation Agent

This agent takes a Power BI report and produces the DAX `EVALUATE` queries that reproduce each visual's underlying data — so Lara's agent can run them against Fabric and get the actual numbers for PowerPoint slides. It can also **generate chart images** from the query results that visually resemble the original PBI visuals.

## How It Works

The pipeline is 3 steps:

1. **Extract metadata** — Parses the PBIP report files (JSON + TMDL) and pulls out every visual's fields, filters, measures, and bookmarks into a standardized Excel.
2. **Build DAX queries** — Reads that metadata and generates one DAX query per visual (e.g., `SUMMARIZECOLUMNS`, `ROW`, `VALUES`) plus bookmark-filtered variants if bookmarks exist.
3. **Generate charts** — Takes the tabular data from executing those DAX queries + the visual type from metadata, and produces PBI-styled chart images (PNG) using plotly.

## 3 Ways to Feed It Input

- **PBIP files** (most complete) — parses the report JSON + semantic model TMDL directly
- **CSV/Excel exports** — from right-click → Export data on a visual
- **Screenshots** — Claude reads the image and identifies the visual type/fields

All three produce the same metadata format, so the DAX builder and chart generator work identically regardless of input method.

## Chart Generation (Skill 5)

Once you have tabular data from executing DAX queries, Skill 5 generates chart images that resemble the original PBI visuals. Two input modes:

- **CSV + Metadata** — automated: reads visual type and field roles from the metadata Excel
- **CSV + Screenshot** — agent views a screenshot, identifies the chart type/fields, passes them as CLI args

Supports 16 chart types: bar, column, stacked bar/column, line, area, pie, donut, scatter, waterfall, combo (dual-axis), funnel, treemap, gauge, card, KPI, and table. All styled with PBI's default color palette and Segoe UI font.

## Important for Report Authors

Any change made in PBI Desktop — adding a column to a visual, creating a measure, adjusting filters — gets written into the PBIP's JSON and TMDL files on save. The agent reads those files directly, so it automatically picks up whatever the report currently contains. No manual mapping needed.

## Validation

3 sample reports validated (Revenue Opportunities, Store Sales, AI Sample), handles report/page/visual-level filters, bookmark filters (IN, NOT IN, date ranges), and recursive measure dependencies. Chart generation tested across bar, pie, line, waterfall, card, and combo chart types.
