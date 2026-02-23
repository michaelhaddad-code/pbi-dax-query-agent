# PBI DAX Query Generation Agent

This agent takes a Power BI report and produces the DAX `EVALUATE` queries that reproduce each visual's underlying data — so Lara's agent can run them against Fabric and get the actual numbers for PowerPoint slides.

## How It Works

The pipeline is 2 steps:

1. **Extract metadata** — Parses the PBIP report files (JSON + TMDL) and pulls out every visual's fields, filters, measures, and bookmarks into a standardized Excel.
2. **Build DAX queries** — Reads that metadata and generates one DAX query per visual (e.g., `SUMMARIZECOLUMNS`, `ROW`, `VALUES`) plus bookmark-filtered variants if bookmarks exist.

## 3 Ways to Feed It Input

- **PBIP files** (most complete) — parses the report JSON + semantic model TMDL directly
- **CSV/Excel exports** — from right-click → Export data on a visual
- **Screenshots** — Claude reads the image and identifies the visual type/fields

All three produce the same metadata format, so the DAX builder works identically regardless of input method.

## Important for Report Authors

Any change made in PBI Desktop — adding a column to a visual, creating a measure, adjusting filters — gets written into the PBIP's JSON and TMDL files on save. The agent reads those files directly, so it automatically picks up whatever the report currently contains. No manual mapping needed.

## Validation

3 sample reports validated (Revenue Opportunities, Store Sales, AI Sample), handles report/page/visual-level filters, bookmark filters (IN, NOT IN, date ranges), and recursive measure dependencies.
