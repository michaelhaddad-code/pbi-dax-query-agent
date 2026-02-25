---
name: run-dax-pipeline
description: Run the full DAX query generation pipeline end-to-end. Use when the user asks to run everything, process a full report into DAX queries, or chain extract-metadata and build-dax-queries together. Also use when the user says "run the pipeline" or "generate all DAX queries for this report".
---

# Run Full DAX Query Generation Pipeline

## What This Skill Does
Chains both skills in sequence to go from raw PBIP files to DAX queries in one run:
1. **extract_metadata.py** — parses report JSON + TMDL files into a metadata Excel
2. **dax_query_builder.py** — reads that metadata and generates DAX queries per visual

## When to Use
- User says "run the pipeline" or "process this report"
- User wants end-to-end DAX generation from PBIP files
- User provides PBIP folder paths and wants DAX queries out

## How to Run

### As Python imports (preferred):
```python
import sys
sys.path.insert(0, "skills")
from extract_metadata import extract_metadata, export_to_excel
from dax_query_builder import read_extractor_output, write_output

# --- Configuration ---
report_root = "data/<ReportName>.Report/definition"
model_root = "data/<ReportName>.SemanticModel/definition"
metadata_output = "output/pbi_report_metadata.xlsx"
dax_output = "output/dax_queries.xlsx"

# --- Step 1: Extract metadata ---
print("Step 1: Extracting metadata...")
df = extract_metadata(report_root, model_root)
export_to_excel(df, metadata_output)

# --- Step 2: Generate DAX queries ---
print("Step 2: Generating DAX queries...")
visuals, page_filters = read_extractor_output(metadata_output)
count = write_output(visuals, page_filters, dax_output)
print(f"Done. Generated DAX queries for {count} visuals.")
```

### From command line (two steps):
```bash
# Step 1: Extract metadata
python skills/extract_metadata.py \
    --report-root "data/<ReportName>.Report/definition" \
    --model-root "data/<ReportName>.SemanticModel/definition" \
    --output "output/pbi_report_metadata.xlsx"

# Step 2: Generate DAX queries
python skills/dax_query_builder.py "output/pbi_report_metadata.xlsx" "output/dax_queries.xlsx"
```

### Revenue Opportunities full test run:
```bash
python skills/extract_metadata.py \
    --report-root "data/Revenue Opportunities.Report/definition" \
    --model-root "data/Revenue Opportunities.SemanticModel/definition" \
    --output "output/pbi_report_metadata.xlsx"

python skills/dax_query_builder.py "output/pbi_report_metadata.xlsx" "output/dax_queries_revenue_opportunities.xlsx"
```

## Required Inputs
1. **PBIP Report folder** — `<ReportName>.Report/definition/` (with `pages/` and `report.json`)
2. **Semantic Model folder** — `<ReportName>.SemanticModel/definition/` (with `tables/` containing `.tmdl` files)

## Output (all in output/ folder)
- `pbi_report_metadata.xlsx` — full report metadata (intermediate, from Step 1)
- `dax_queries.xlsx` — DAX queries per visual (final output, from Step 2)

## Pre-Run Checklist
Before running, verify:
- [ ] PBIP report folder exists and contains `pages/` and `report.json`
- [ ] Semantic model folder exists and contains `tables/` with `.tmdl` files
- [ ] `output/` directory exists
- [ ] Input PBIP folders are placed under `data/`

## Post-Run Validation
After running, check:
- [ ] Compare `output/dax_queries*.xlsx` against `~/Desktop/dax_queries_by_visual.xlsx` (manual reference)
- [ ] All visuals with data fields have a DAX query (not "Unknown")
- [ ] Measure-only visuals (cards) use Pattern 1
- [ ] Column-only visuals (slicers) use Pattern 2
- [ ] Column + measure visuals use Pattern 3 (SUMMARIZECOLUMNS)

## Critical Rules
- NEVER modify input PBIP files — they are read-only
- ALWAYS save outputs to the `output/` folder
- Step 2 depends on Step 1 — always run extract_metadata first
