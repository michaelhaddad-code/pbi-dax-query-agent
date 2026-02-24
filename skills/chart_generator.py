# -*- coding: utf-8 -*-
"""
Skill 5: chart_generator.py
PBI AutoGov -- Chart Generator

Generates chart visuals from DAX query tabular data that visually resemble
the original Power BI visuals. Supports two output formats:
  - PPTX (default): single-slide .pptx with native editable chart or PNG fallback
  - PNG (legacy): plotly-rendered static image

Two input modes:
  Mode 1 (CSV + Metadata): reads metadata Excel to get visual type + field roles
  Mode 2 (CSV + Screenshot): agent passes visual type + fields as CLI args

Input:  CSV file with DAX query results + visual metadata (Excel or CLI args)
Output: .pptx file (one slide per visual) or .png image

Usage:
    # Mode 1: CSV + Metadata Excel
    python skills/chart_generator.py \
      --csv "output/revenue_data.csv" \
      --metadata "output/pbi_report_metadata.xlsx" \
      --visual "Pipeline by Stage" \
      --output "output/charts/"

    # Mode 2: CSV + Screenshot (agent-driven)
    python skills/chart_generator.py \
      --csv "output/revenue_data.csv" \
      --visual-type barChart \
      --visual-name "Pipeline by Stage" \
      --field "Sales Stage:grouping" \
      --field "Opportunity Count:measure" \
      --output "output/charts/"
"""

import argparse
import os
import re
import sys
import tempfile
from dataclasses import dataclass, field
from pathlib import Path

import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots

from pptx import Presentation as _PptxFactory
from pptx.presentation import Presentation as PptxPresentation  # actual class for isinstance
from pptx.chart.data import CategoryChartData, XyChartData
from pptx.util import Inches, Pt, Emu
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION, XL_LABEL_POSITION
from pptx.dml.color import RGBColor


# =============================================================================
# CONSTANTS
# =============================================================================

# Power BI default color palette (10 accent colors)
PBI_COLORS = [
    "#118DFF", "#12239E", "#E66C37", "#6B007B", "#E044A7",
    "#744EC2", "#D9B300", "#D64550", "#197278", "#1AAB40",
]

PBI_FONT = "Segoe UI"

# PBI colors as RGBColor tuples for python-pptx native charts
PBI_RGB_COLORS = [
    RGBColor(0x11, 0x8D, 0xFF),  # #118DFF
    RGBColor(0x12, 0x23, 0x9E),  # #12239E
    RGBColor(0xE6, 0x6C, 0x37),  # #E66C37
    RGBColor(0x6B, 0x00, 0x7B),  # #6B007B
    RGBColor(0xE0, 0x44, 0xA7),  # #E044A7
    RGBColor(0x74, 0x4E, 0xC2),  # #744EC2
    RGBColor(0xD9, 0xB3, 0x00),  # #D9B300
    RGBColor(0xD6, 0x45, 0x50),  # #D64550
    RGBColor(0x19, 0x72, 0x78),  # #197278
    RGBColor(0x1A, 0xAB, 0x40),  # #1AAB40
]

# Slide dimensions: 16:9 widescreen
SLIDE_WIDTH = Inches(13.333)
SLIDE_HEIGHT = Inches(7.5)

# Chart placement: centered with margins for title
CHART_LEFT = Inches(0.5)
CHART_TOP = Inches(0.8)
CHART_WIDTH = Inches(12.333)
CHART_HEIGHT = Inches(6.2)

# Maps PBI visual type -> XL_CHART_TYPE enum for native python-pptx rendering
NATIVE_CHART_MAP = {
    # Bar charts (horizontal)
    "barChart": XL_CHART_TYPE.BAR_CLUSTERED,
    "clusteredBarChart": XL_CHART_TYPE.BAR_CLUSTERED,
    "stackedBarChart": XL_CHART_TYPE.BAR_STACKED,
    "hundredPercentStackedBarChart": XL_CHART_TYPE.BAR_STACKED_100,
    # Column charts (vertical)
    "columnChart": XL_CHART_TYPE.COLUMN_CLUSTERED,
    "clusteredColumnChart": XL_CHART_TYPE.COLUMN_CLUSTERED,
    "stackedColumnChart": XL_CHART_TYPE.COLUMN_STACKED,
    "hundredPercentStackedColumnChart": XL_CHART_TYPE.COLUMN_STACKED_100,
    # Line
    "lineChart": XL_CHART_TYPE.LINE_MARKERS,
    # Area
    "areaChart": XL_CHART_TYPE.AREA,
    "stackedAreaChart": XL_CHART_TYPE.AREA_STACKED,
    # Pie and donut
    "pieChart": XL_CHART_TYPE.PIE,
    "donutChart": XL_CHART_TYPE.DOUGHNUT,
    # Scatter
    "scatterChart": XL_CHART_TYPE.XY_SCATTER,
}

# Visual types that require plotly PNG fallback (no native python-pptx support)
PNG_FALLBACK_TYPES = {
    "lineClusteredColumnComboChart", "lineStackedColumnComboChart",
    "waterfallChart", "funnelChart", "treemap",
    "gauge", "card", "cardVisual", "multiRowCard", "kpi",
    "tableEx", "pivotTable", "ribbonChart",
}


# =============================================================================
# DATA CLASS
# =============================================================================

@dataclass
class VisualSpec:
    """Metadata describing a single PBI visual for chart generation."""
    page_name: str
    visual_name: str
    visual_type: str                          # PBI camelCase identifier (e.g., "barChart")
    grouping_columns: list = field(default_factory=list)  # category/axis field names
    measure_columns: list = field(default_factory=list)   # value/measure field names
    y2_columns: list = field(default_factory=list)        # secondary-axis measures (Visual Y2)
    dax_pattern: str = ""                     # informational (e.g., "Pattern 3")


# =============================================================================
# PBI THEME
# =============================================================================

def get_pbi_plotly_layout():
    """Return a plotly layout dict matching PBI's default visual style.

    Applied to all chart types for consistent appearance:
    - PBI color palette as colorway
    - Segoe UI font
    - White background
    - Light gray horizontal gridlines only
    """
    return {
        "font": {"family": PBI_FONT, "size": 12, "color": "#333333"},
        "plot_bgcolor": "white",
        "paper_bgcolor": "white",
        "colorway": PBI_COLORS,
        "margin": {"l": 60, "r": 30, "t": 50, "b": 60},
        "xaxis": {
            "showgrid": False,
            "linecolor": "#E0E0E0",
            "tickfont": {"size": 10, "color": "#666666"},
        },
        "yaxis": {
            "showgrid": True,
            "gridcolor": "#E0E0E0",
            "linecolor": "#E0E0E0",
            "tickfont": {"size": 10, "color": "#666666"},
        },
    }


# =============================================================================
# DATA HELPERS
# =============================================================================

def classify_columns(df, spec):
    """Split DataFrame columns into categories (grouping) and values (measures).

    Uses the VisualSpec's field lists to match against actual DataFrame column
    names (case-insensitive). Falls back to dtype inference if no match.

    Returns:
        (categories: list[str], values: list[str]) — column names in df
    """
    df_cols_lower = {c.lower().strip(): c for c in df.columns}

    categories = []
    for gc in spec.grouping_columns:
        actual = df_cols_lower.get(gc.lower().strip())
        if actual and actual not in categories:
            categories.append(actual)

    values = []
    for mc in spec.measure_columns:
        actual = df_cols_lower.get(mc.lower().strip())
        if actual and actual not in values:
            values.append(actual)

    # Fallback: infer from data types if spec columns didn't match
    if not categories and not values:
        for col in df.columns:
            if pd.api.types.is_numeric_dtype(df[col]):
                values.append(col)
            else:
                categories.append(col)

    return categories, values


def _prepare_series_data(df, categories, values):
    """Prepare category labels and series data for bar/column/stacked charts.

    Handles the common pivot logic: when 2+ grouping columns exist, the second
    column is pivoted into legend series (like PBI's "Legend" well). Otherwise,
    each measure becomes its own series.

    Returns:
        (cat_labels: list[str], series: OrderedDict[str, list[float]], needs_legend: bool)
    """
    from collections import OrderedDict

    if len(categories) >= 2 and len(values) == 1:
        # Pivot second grouping column into legend series
        pivot_df = df.pivot_table(
            index=categories[0], columns=categories[1],
            values=values[0], aggfunc="sum"
        ).fillna(0)
        cat_labels = [str(c) for c in pivot_df.index]
        series = OrderedDict((str(col), pivot_df[col].tolist()) for col in pivot_df.columns)
        return cat_labels, series, True
    else:
        cat_labels = df[categories[0]].astype(str).tolist()
        series = OrderedDict((v, df[v].tolist()) for v in values)
        needs_legend = len(values) > 1
        return cat_labels, series, needs_legend


def parse_visual_from_metadata(metadata_excel, visual_name):
    """Read metadata Excel (from Skill 1) and build a VisualSpec for a named visual.

    Imports classify_field from dax_query_builder to determine field roles.
    Matches visual_name case-insensitively (partial match supported).

    Returns:
        VisualSpec or None if visual not found
    """
    # Import from sibling module (same directory)
    skills_dir = os.path.dirname(os.path.abspath(__file__))
    if skills_dir not in sys.path:
        sys.path.insert(0, skills_dir)
    from dax_query_builder import read_extractor_output, classify_field

    visuals, _, _, _ = read_extractor_output(metadata_excel)

    # Find the visual by name (case-insensitive partial match)
    target = visual_name.lower().strip()
    matched_key = None
    for key, data in visuals.items():
        if data["visual_name"].lower().strip() == target:
            matched_key = key
            break
    # Partial match fallback
    if not matched_key:
        for key, data in visuals.items():
            if target in data["visual_name"].lower():
                matched_key = key
                break

    if not matched_key:
        print(f"ERROR: Visual '{visual_name}' not found in metadata.")
        print(f"Available visuals:")
        for key, data in visuals.items():
            print(f"  - {data['visual_name']} ({data['visual_type']}) on page '{key[0]}'")
        return None

    data = visuals[matched_key]
    page_name = matched_key[0]

    grouping_cols = []
    measure_cols = []
    y2_cols = []
    for f in data["fields"]:
        role = classify_field(f["usage"])
        if role == "grouping":
            grouping_cols.append(f["ui_name"])
        elif role == "measure":
            if f["ui_name"] not in measure_cols:
                measure_cols.append(f["ui_name"])
                # Check if this measure is assigned to the secondary Y-axis
                if "y2" in f["usage"].lower():
                    y2_cols.append(f["ui_name"])

    return VisualSpec(
        page_name=page_name,
        visual_name=data["visual_name"],
        visual_type=data["visual_type"],
        grouping_columns=grouping_cols,
        measure_columns=measure_cols,
        y2_columns=y2_cols,
    )


# =============================================================================
# CHART RENDERERS — each returns a plotly Figure
# =============================================================================

# ---- Bar chart (horizontal) ----

def _render_bar(df, spec):
    """Horizontal bar chart. Categories on Y-axis, values on X-axis.

    Handles single measure (one series) or multiple measures (grouped bars).
    With 2+ grouping columns, pivots second column into legend series.
    """
    categories, values = classify_columns(df, spec)
    if not categories or not values:
        return _render_table(df, spec)

    # Sort by first measure descending for clean ranking (like PBI default)
    if len(values) == 1 and len(categories) == 1:
        df = df.sort_values(values[0], ascending=False)  # highest at top with reversed y-axis

    cat_labels, series, needs_legend = _prepare_series_data(df, categories, values)

    fig = go.Figure()
    for i, (name, data) in enumerate(series.items()):
        fig.add_trace(go.Bar(
            y=cat_labels, x=data, name=name, orientation="h",
            marker_color=PBI_COLORS[i % len(PBI_COLORS)],
        ))

    fig.update_layout(
        **get_pbi_plotly_layout(),
        title=spec.visual_name,
        barmode="group",
        showlegend=needs_legend,
        xaxis_title=values[0] if len(values) == 1 else None,
        yaxis_title=categories[0] if categories else None,
    )
    fig.update_yaxes(autorange="reversed")  # top-to-bottom ordering like PBI
    return fig


# ---- Column chart (vertical) ----

def _render_column(df, spec):
    """Vertical column chart. Categories on X-axis, values on Y-axis."""
    categories, values = classify_columns(df, spec)
    if not categories or not values:
        return _render_table(df, spec)

    # Sort by first measure descending
    if len(values) == 1 and len(categories) == 1:
        df = df.sort_values(values[0], ascending=False)

    cat_labels, series, needs_legend = _prepare_series_data(df, categories, values)

    fig = go.Figure()
    for i, (name, data) in enumerate(series.items()):
        fig.add_trace(go.Bar(
            x=cat_labels, y=data, name=name,
            marker_color=PBI_COLORS[i % len(PBI_COLORS)],
        ))

    fig.update_layout(
        **get_pbi_plotly_layout(),
        title=spec.visual_name,
        barmode="group",
        showlegend=needs_legend,
        xaxis_title=categories[0] if categories else None,
        yaxis_title=values[0] if len(values) == 1 else None,
    )
    return fig


# ---- Stacked bar (horizontal) ----

def _render_stacked_bar(df, spec):
    """Horizontal stacked bar chart. Uses barmode='stack' or barnorm='percent'."""
    categories, values = classify_columns(df, spec)
    if not categories or not values:
        return _render_table(df, spec)

    cat_labels, series, _ = _prepare_series_data(df, categories, values)

    fig = go.Figure()
    for i, (name, data) in enumerate(series.items()):
        fig.add_trace(go.Bar(
            y=cat_labels, x=data, name=name, orientation="h",
            marker_color=PBI_COLORS[i % len(PBI_COLORS)],
        ))

    layout_kwargs = {
        **get_pbi_plotly_layout(),
        "title": spec.visual_name,
        "barmode": "stack",
        "showlegend": True,
        "xaxis_title": values[0] if len(values) == 1 else None,
        "yaxis_title": categories[0] if categories else None,
    }
    if "hundredPercent" in spec.visual_type:
        layout_kwargs["barnorm"] = "percent"
    fig.update_layout(**layout_kwargs)
    fig.update_yaxes(autorange="reversed")  # top-to-bottom ordering like PBI
    return fig


# ---- Stacked column (vertical) ----

def _render_stacked_column(df, spec):
    """Vertical stacked column chart. Uses barmode='stack' or barnorm='percent'."""
    categories, values = classify_columns(df, spec)
    if not categories or not values:
        return _render_table(df, spec)

    cat_labels, series, _ = _prepare_series_data(df, categories, values)

    fig = go.Figure()
    for i, (name, data) in enumerate(series.items()):
        fig.add_trace(go.Bar(
            x=cat_labels, y=data, name=name,
            marker_color=PBI_COLORS[i % len(PBI_COLORS)],
        ))

    layout_kwargs = {
        **get_pbi_plotly_layout(),
        "title": spec.visual_name,
        "barmode": "stack",
        "showlegend": True,
        "xaxis_title": categories[0] if categories else None,
        "yaxis_title": values[0] if len(values) == 1 else None,
    }
    if "hundredPercent" in spec.visual_type:
        layout_kwargs["barnorm"] = "percent"
    fig.update_layout(**layout_kwargs)
    return fig


# ---- Line chart ----

def _render_line(df, spec):
    """Line chart. X-axis is first grouping column, one line per measure.

    Data is sorted by the category column for proper line ordering.
    With 2+ grouping columns, pivots second into legend series.
    """
    categories, values = classify_columns(df, spec)
    if not categories or not values:
        return _render_table(df, spec)

    df_sorted = df.sort_values(categories[0])

    if len(categories) >= 2 and len(values) == 1:
        pivot_df = df_sorted.pivot_table(
            index=categories[0], columns=categories[1],
            values=values[0], aggfunc="sum"
        ).fillna(0)
        fig = go.Figure()
        for i, col in enumerate(pivot_df.columns):
            fig.add_trace(go.Scatter(
                x=pivot_df.index.astype(str), y=pivot_df[col],
                name=str(col), mode="lines+markers",
                line={"color": PBI_COLORS[i % len(PBI_COLORS)]},
            ))
        needs_legend = True
    else:
        fig = go.Figure()
        cat_labels = df_sorted[categories[0]].astype(str).tolist()
        for i, v in enumerate(values):
            fig.add_trace(go.Scatter(
                x=cat_labels, y=df_sorted[v].tolist(),
                name=v, mode="lines+markers",
                line={"color": PBI_COLORS[i % len(PBI_COLORS)]},
            ))
        needs_legend = len(values) > 1

    fig.update_layout(
        **get_pbi_plotly_layout(),
        title=spec.visual_name,
        showlegend=needs_legend,
        xaxis_title=categories[0] if categories else None,
        yaxis_title=values[0] if len(values) == 1 else None,
    )
    return fig


# ---- Area chart ----

def _render_area(df, spec):
    """Area chart. Like line chart but with fill to zero.

    Stacked area uses stackgroup for overlapping fills.
    """
    categories, values = classify_columns(df, spec)
    if not categories or not values:
        return _render_table(df, spec)

    is_stacked = "stacked" in spec.visual_type.lower()
    df_sorted = df.sort_values(categories[0])

    fig = go.Figure()
    cat_labels = df_sorted[categories[0]].astype(str).tolist()
    for i, v in enumerate(values):
        trace_kwargs = {
            "x": cat_labels,
            "y": df_sorted[v].tolist(),
            "name": v,
            "mode": "lines",
            "line": {"color": PBI_COLORS[i % len(PBI_COLORS)]},
        }
        if is_stacked:
            trace_kwargs["stackgroup"] = "one"
        else:
            trace_kwargs["fill"] = "tozeroy"
        fig.add_trace(go.Scatter(**trace_kwargs))

    fig.update_layout(
        **get_pbi_plotly_layout(),
        title=spec.visual_name,
        showlegend=len(values) > 1,
        xaxis_title=categories[0] if categories else None,
        yaxis_title=values[0] if len(values) == 1 else None,
    )
    return fig


# ---- Pie chart ----

def _render_pie(df, spec):
    """Pie chart. Labels from first grouping column, values from first measure.

    Only uses first measure (pie charts show one value series). Warns if multiple.
    Measures-only case: each measure name becomes a slice label, each value a slice size.
    """
    categories, values = classify_columns(df, spec)

    # Measures-only: no grouping column, each measure is a slice
    if not categories and len(values) >= 2:
        row = df.iloc[0]
        slice_labels = values
        slice_values = [float(row[v]) if pd.notna(row[v]) else 0 for v in values]
        fig = go.Figure(go.Pie(
            labels=slice_labels,
            values=slice_values,
            marker={"colors": PBI_COLORS[:len(values)]},
            textinfo="percent+label",
            textfont={"size": 11, "family": PBI_FONT},
            hole=0,
        ))
        fig.update_layout(
            title=spec.visual_name,
            font={"family": PBI_FONT, "size": 12, "color": "#333333"},
            paper_bgcolor="white",
            showlegend=True,
            legend={"font": {"size": 10}},
        )
        return fig

    if not categories or not values:
        return _render_table(df, spec)

    if len(values) > 1:
        print(f"  WARNING: Pie chart '{spec.visual_name}' has {len(values)} measures "
              f"-- using only '{values[0]}'")

    fig = go.Figure(go.Pie(
        labels=df[categories[0]].astype(str).tolist(),
        values=df[values[0]].tolist(),
        marker={"colors": PBI_COLORS[:len(df)]},
        textinfo="percent+label",
        textfont={"size": 11, "family": PBI_FONT},
        hole=0,
    ))
    fig.update_layout(
        title=spec.visual_name,
        font={"family": PBI_FONT, "size": 12, "color": "#333333"},
        paper_bgcolor="white",
        showlegend=True,
        legend={"font": {"size": 10}},
    )
    return fig


# ---- Donut chart ----

def _render_donut(df, spec):
    """Donut chart. Same as pie but with a hole in the center (hole=0.4).

    Measures-only case: each measure name becomes a slice label, each value a slice size.
    """
    categories, values = classify_columns(df, spec)

    # Measures-only: no grouping column, each measure is a slice
    if not categories and len(values) >= 2:
        row = df.iloc[0]
        slice_labels = values
        slice_values = [float(row[v]) if pd.notna(row[v]) else 0 for v in values]
        fig = go.Figure(go.Pie(
            labels=slice_labels,
            values=slice_values,
            marker={"colors": PBI_COLORS[:len(values)]},
            textinfo="percent+label+value",
            textfont={"size": 11, "family": PBI_FONT},
            hole=0.4,
        ))
        fig.update_layout(
            title=spec.visual_name,
            font={"family": PBI_FONT, "size": 12, "color": "#333333"},
            paper_bgcolor="white",
            showlegend=True,
            legend={"font": {"size": 10}},
        )
        return fig

    if not categories or not values:
        return _render_table(df, spec)

    if len(values) > 1:
        print(f"  WARNING: Donut chart '{spec.visual_name}' has {len(values)} measures "
              f"-- using only '{values[0]}'")

    fig = go.Figure(go.Pie(
        labels=df[categories[0]].astype(str).tolist(),
        values=df[values[0]].tolist(),
        marker={"colors": PBI_COLORS[:len(df)]},
        textinfo="percent+label",
        textfont={"size": 11, "family": PBI_FONT},
        hole=0.4,
    ))
    fig.update_layout(
        title=spec.visual_name,
        font={"family": PBI_FONT, "size": 12, "color": "#333333"},
        paper_bgcolor="white",
        showlegend=True,
        legend={"font": {"size": 10}},
    )
    return fig


# ---- Scatter chart ----

def _render_scatter(df, spec):
    """Scatter plot. X = first measure, Y = second measure.

    If a grouping column exists, it becomes the point labels/color grouper.
    Third measure (if present) maps to marker size (bubble chart).
    """
    categories, values = classify_columns(df, spec)
    if len(values) < 2:
        return _render_table(df, spec)

    # Guard for bubble sizing: ensure positive max for sizeref calculation
    has_bubble = len(values) >= 3
    if has_bubble:
        max_bubble = df[values[2]].max()
        # Fallback to fixed size if all bubble values are zero or negative
        if pd.isna(max_bubble) or max_bubble <= 0:
            has_bubble = False

    fig = go.Figure()

    if categories:
        groups = df.groupby(categories[0])
        for i, (name, group) in enumerate(groups):
            trace_kwargs = {
                "x": group[values[0]].tolist(),
                "y": group[values[1]].tolist(),
                "name": str(name),
                "mode": "markers",
                "marker": {"color": PBI_COLORS[i % len(PBI_COLORS)], "size": 10},
            }
            if has_bubble:
                trace_kwargs["marker"]["size"] = group[values[2]].tolist()
                trace_kwargs["marker"]["sizemode"] = "area"
                trace_kwargs["marker"]["sizeref"] = 2.0 * max_bubble / (40.0 ** 2)
            fig.add_trace(go.Scatter(**trace_kwargs))
    else:
        trace_kwargs = {
            "x": df[values[0]].tolist(),
            "y": df[values[1]].tolist(),
            "mode": "markers",
            "marker": {"color": PBI_COLORS[0], "size": 10},
        }
        if has_bubble:
            trace_kwargs["marker"]["size"] = df[values[2]].tolist()
            trace_kwargs["marker"]["sizemode"] = "area"
            trace_kwargs["marker"]["sizeref"] = 2.0 * max_bubble / (40.0 ** 2)
        fig.add_trace(go.Scatter(**trace_kwargs))

    fig.update_layout(
        **get_pbi_plotly_layout(),
        title=spec.visual_name,
        xaxis_title=values[0],
        yaxis_title=values[1],
        showlegend=bool(categories),
    )
    return fig


# ---- Waterfall chart ----

def _render_waterfall(df, spec):
    """Waterfall chart using plotly's native go.Waterfall trace.

    Categories from first grouping column, values from first measure.
    All values treated as relative (incremental). PBI blue for positive,
    PBI red for negative.
    """
    categories, values = classify_columns(df, spec)
    if not categories or not values:
        return _render_table(df, spec)

    cat_labels = df[categories[0]].astype(str).tolist()
    val_data = df[values[0]].tolist()

    # All bars are relative (incremental values)
    measure_types = ["relative"] * len(val_data)

    fig = go.Figure(go.Waterfall(
        x=cat_labels,
        y=val_data,
        measure=measure_types,
        connector={"line": {"color": "#E0E0E0"}},
        increasing={"marker": {"color": PBI_COLORS[0]}},   # blue for positive
        decreasing={"marker": {"color": PBI_COLORS[7]}},    # red for negative
        totals={"marker": {"color": PBI_COLORS[1]}},        # dark blue for totals
        textposition="outside",
    ))
    fig.update_layout(
        **get_pbi_plotly_layout(),
        title=spec.visual_name,
        showlegend=False,
        xaxis_title=categories[0] if categories else None,
        yaxis_title=values[0] if len(values) == 1 else None,
    )
    return fig


# ---- Combo chart (dual axis: bars + line) ----

def _render_combo(df, spec):
    """Combo chart with bars on primary Y-axis and line on secondary Y-axis.

    Uses plotly's make_subplots with secondary_y for dual-axis support.
    If spec.y2_columns is populated (from metadata "Visual Y2" usage), those
    measures go on the secondary axis as lines. Otherwise falls back to
    putting all measures except the last on bars and the last on line.
    """
    categories, values = classify_columns(df, spec)
    if not categories or not values:
        return _render_table(df, spec)

    cat_labels = df[categories[0]].astype(str).tolist()

    # Resolve actual column names for y2 (case-insensitive match against df)
    df_cols_lower = {c.lower().strip(): c for c in df.columns}
    y2_actual = set()
    for y2 in spec.y2_columns:
        actual = df_cols_lower.get(y2.lower().strip())
        if actual:
            y2_actual.add(actual)

    # Split measures into bar vs line using y2 metadata when available
    if y2_actual:
        bar_measures = [v for v in values if v not in y2_actual]
        line_measures = [v for v in values if v in y2_actual]
    elif len(values) > 1:
        # Fallback: all except last → bars, last → line
        bar_measures = values[:-1]
        line_measures = [values[-1]]
    else:
        bar_measures = values
        line_measures = []

    fig = make_subplots(specs=[[{"secondary_y": bool(line_measures)}]])

    for i, m in enumerate(bar_measures):
        fig.add_trace(go.Bar(
            x=cat_labels, y=df[m].tolist(), name=m,
            marker_color=PBI_COLORS[i % len(PBI_COLORS)],
        ), secondary_y=False)

    for i, m in enumerate(line_measures):
        fig.add_trace(go.Scatter(
            x=cat_labels, y=df[m].tolist(), name=m,
            mode="lines+markers",
            line={"color": PBI_COLORS[(len(bar_measures) + i) % len(PBI_COLORS)], "width": 2},
        ), secondary_y=True)

    layout = get_pbi_plotly_layout()
    fig.update_layout(
        title=spec.visual_name,
        font=layout["font"],
        plot_bgcolor=layout["plot_bgcolor"],
        paper_bgcolor=layout["paper_bgcolor"],
        barmode="group",
        showlegend=True,
        margin=layout["margin"],
    )
    fig.update_xaxes(showgrid=False, linecolor="#E0E0E0",
                     title_text=categories[0] if categories else None)
    fig.update_yaxes(showgrid=True, gridcolor="#E0E0E0", linecolor="#E0E0E0")
    if bar_measures:
        fig.update_yaxes(title_text=bar_measures[0] if len(bar_measures) == 1 else None,
                         secondary_y=False)
    if line_measures:
        fig.update_yaxes(title_text=line_measures[0] if len(line_measures) == 1 else None,
                         secondary_y=True)
    return fig


# ---- Funnel chart ----

def _render_funnel(df, spec):
    """Funnel chart. Categories (stages) on Y-axis, values (counts) on X-axis."""
    categories, values = classify_columns(df, spec)
    if not categories or not values:
        return _render_table(df, spec)

    fig = go.Figure(go.Funnel(
        y=df[categories[0]].astype(str).tolist(),
        x=df[values[0]].tolist(),
        marker={"color": PBI_COLORS[:len(df)]},
        textinfo="value+percent initial",
        textfont={"family": PBI_FONT, "size": 11},
    ))
    fig.update_layout(
        title=spec.visual_name,
        font={"family": PBI_FONT, "size": 12, "color": "#333333"},
        paper_bgcolor="white",
        plot_bgcolor="white",
        showlegend=False,
        yaxis_title=categories[0] if categories else None,
        xaxis_title=values[0] if len(values) == 1 else None,
    )
    return fig


# ---- Treemap ----

def _render_treemap(df, spec):
    """Treemap chart. Labels from grouping column, tile sizes from measure.

    With 2+ grouping columns, creates nested hierarchy (parent -> child).
    """
    categories, values = classify_columns(df, spec)
    if not categories or not values:
        return _render_table(df, spec)

    if len(categories) >= 2:
        # Nested treemap: first grouping = parent, second = child
        labels = df[categories[1]].astype(str).tolist()
        parents = df[categories[0]].astype(str).tolist()
    else:
        labels = df[categories[0]].astype(str).tolist()
        parents = [""] * len(labels)

    fig = go.Figure(go.Treemap(
        labels=labels,
        parents=parents,
        values=df[values[0]].tolist(),
        marker={"colors": PBI_COLORS[:len(df)]},
        textinfo="label+value",
        textfont={"family": PBI_FONT, "size": 12},
    ))
    fig.update_layout(
        title=spec.visual_name,
        font={"family": PBI_FONT, "size": 12, "color": "#333333"},
        paper_bgcolor="white",
        margin={"l": 10, "r": 10, "t": 50, "b": 10},
    )
    return fig


# ---- Gauge ----

def _render_gauge(df, spec):
    """Gauge indicator. Value from first measure, optional target from second.

    Range auto-calculated as 0 to value * 1.2.
    """
    _, values = classify_columns(df, spec)
    if not values or df.empty:
        return _render_table(df, spec)

    value = float(df[values[0]].iloc[0])
    max_range = value * 1.2 if value > 0 else 100

    gauge_kwargs = {
        "axis": {"range": [0, max_range], "tickfont": {"size": 10}},
        "bar": {"color": PBI_COLORS[0]},
        "bgcolor": "white",
        "borderwidth": 0,
        "steps": [
            {"range": [0, max_range * 0.5], "color": "#F0F0F0"},
            {"range": [max_range * 0.5, max_range], "color": "#E0E0E0"},
        ],
    }

    if len(values) >= 2:
        target = float(df[values[1]].iloc[0])
        gauge_kwargs["threshold"] = {
            "line": {"color": PBI_COLORS[7], "width": 3},
            "thickness": 0.8,
            "value": target,
        }

    fig = go.Figure(go.Indicator(
        mode="gauge+number",
        value=value,
        title={"text": spec.visual_name, "font": {"size": 16, "family": PBI_FONT}},
        number={"font": {"size": 36, "family": PBI_FONT, "color": "#333333"}},
        gauge=gauge_kwargs,
    ))
    fig.update_layout(
        paper_bgcolor="white",
        font={"family": PBI_FONT},
        margin={"l": 30, "r": 30, "t": 60, "b": 30},
    )
    return fig


# ---- Card (single or multi-value) ----

def _render_card(df, spec):
    """Card visual. Displays measure values as large centered numbers.

    Single measure -> one big number. Multiple measures -> side-by-side indicators.
    Uses go.Indicator(mode="number") for clean big-number display.
    """
    _, values = classify_columns(df, spec)
    if not values or df.empty:
        return _render_table(df, spec)

    row = df.iloc[0]

    if len(values) == 1:
        val = row[values[0]]
        fig = go.Figure(go.Indicator(
            mode="number",
            value=float(val) if pd.notna(val) else 0,
            title={"text": values[0], "font": {"size": 16, "family": PBI_FONT, "color": "#666666"}},
            number={"font": {"size": 60, "family": PBI_FONT, "color": PBI_COLORS[0]},
                    "valueformat": ",.0f"},
        ))
        fig.update_layout(
            paper_bgcolor="white",
            margin={"l": 30, "r": 30, "t": 60, "b": 30},
        )
    else:
        fig = make_subplots(
            rows=1, cols=len(values),
            specs=[[{"type": "indicator"}] * len(values)],
        )
        for i, v in enumerate(values):
            val = row[v]
            fig.add_trace(go.Indicator(
                mode="number",
                value=float(val) if pd.notna(val) else 0,
                title={"text": v, "font": {"size": 13, "family": PBI_FONT, "color": "#666666"}},
                number={"font": {"size": 36, "family": PBI_FONT, "color": PBI_COLORS[i % len(PBI_COLORS)]},
                        "valueformat": ",.0f"},
            ), row=1, col=i + 1)
        fig.update_layout(
            paper_bgcolor="white",
            margin={"l": 20, "r": 20, "t": 60, "b": 20},
        )

    return fig


# ---- KPI ----

def _render_kpi(df, spec):
    """KPI visual. Shows value with delta (change from previous/target).

    First measure = value, second measure = reference (for delta calculation).
    """
    _, values = classify_columns(df, spec)
    if not values or df.empty:
        return _render_table(df, spec)

    row = df.iloc[0]
    value = float(row[values[0]]) if pd.notna(row[values[0]]) else 0

    indicator_kwargs = {
        "mode": "number+delta",
        "value": value,
        "title": {"text": spec.visual_name, "font": {"size": 16, "family": PBI_FONT}},
        "number": {"font": {"size": 48, "family": PBI_FONT, "color": "#333333"},
                   "valueformat": ",.0f"},
    }

    if len(values) >= 2:
        reference = float(row[values[1]]) if pd.notna(row[values[1]]) else 0
        indicator_kwargs["delta"] = {
            "reference": reference,
            "relative": True,
            "valueformat": ".1%",
            "increasing": {"color": "#1AAB40"},
            "decreasing": {"color": "#D64550"},
        }

    fig = go.Figure(go.Indicator(**indicator_kwargs))
    fig.update_layout(
        paper_bgcolor="white",
        font={"family": PBI_FONT},
        margin={"l": 30, "r": 30, "t": 60, "b": 30},
    )
    return fig


# ---- Ribbon chart ----

def _render_ribbon(df, spec):
    """Ribbon chart rendered as a stacked area chart (closest plotly equivalent)."""
    categories, values = classify_columns(df, spec)
    if not categories or not values:
        return _render_table(df, spec)

    df_sorted = df.sort_values(categories[0])
    fig = go.Figure()

    if len(categories) >= 2 and len(values) == 1:
        pivot_df = df_sorted.pivot_table(
            index=categories[0], columns=categories[1],
            values=values[0], aggfunc="sum"
        ).fillna(0)
        for i, col in enumerate(pivot_df.columns):
            fig.add_trace(go.Scatter(
                x=pivot_df.index.astype(str), y=pivot_df[col],
                name=str(col), mode="lines", stackgroup="one",
                line={"color": PBI_COLORS[i % len(PBI_COLORS)]},
            ))
    else:
        cat_labels = df_sorted[categories[0]].astype(str).tolist()
        for i, v in enumerate(values):
            fig.add_trace(go.Scatter(
                x=cat_labels, y=df_sorted[v].tolist(),
                name=v, mode="lines", stackgroup="one",
                line={"color": PBI_COLORS[i % len(PBI_COLORS)]},
            ))

    fig.update_layout(
        **get_pbi_plotly_layout(),
        title=spec.visual_name,
        showlegend=True,
    )
    return fig


# ---- Table (also used as fallback for unknown types) ----

def _render_table(df, spec):
    """Table visual using go.Table. Also the fallback for unknown visual types.

    Renders up to 50 rows with PBI-styled header and alternating row colors.
    """
    max_rows = 50
    display_df = df.head(max_rows)
    num_rows = len(display_df)

    header_values = list(display_df.columns)
    cell_values = [display_df[col].astype(str).tolist() for col in display_df.columns]

    # Build row color list matching exact row count
    row_colors = ["white" if i % 2 == 0 else "#F5F5F5" for i in range(num_rows)]

    fig = go.Figure(go.Table(
        header={
            "values": [f"<b>{h}</b>" for h in header_values],
            "fill_color": PBI_COLORS[1],        # dark blue header
            "font": {"color": "white", "size": 11, "family": PBI_FONT},
            "align": "left",
            "height": 30,
        },
        cells={
            "values": cell_values,
            "fill_color": [row_colors],
            "font": {"size": 10, "family": PBI_FONT, "color": "#333333"},
            "align": "left",
            "height": 25,
        },
    ))

    title = spec.visual_name
    if len(df) > max_rows:
        title += f" (showing {max_rows} of {len(df)} rows)"

    fig.update_layout(
        title=title,
        font={"family": PBI_FONT},
        paper_bgcolor="white",
        margin={"l": 10, "r": 10, "t": 50, "b": 10},
    )
    return fig


# =============================================================================
# CHART TYPE ROUTER
# =============================================================================

# Maps PBI visual type identifiers to renderer functions
CHART_TYPE_ROUTER = {
    # Bar charts (horizontal)
    "barChart": _render_bar,
    "clusteredBarChart": _render_bar,
    "stackedBarChart": _render_stacked_bar,
    "hundredPercentStackedBarChart": _render_stacked_bar,

    # Column charts (vertical)
    "columnChart": _render_column,
    "clusteredColumnChart": _render_column,
    "stackedColumnChart": _render_stacked_column,
    "hundredPercentStackedColumnChart": _render_stacked_column,

    # Line and area
    "lineChart": _render_line,
    "areaChart": _render_area,
    "stackedAreaChart": _render_area,

    # Pie and donut
    "pieChart": _render_pie,
    "donutChart": _render_donut,

    # Scatter
    "scatterChart": _render_scatter,

    # Waterfall
    "waterfallChart": _render_waterfall,

    # Combo (dual axis)
    "lineStackedColumnComboChart": _render_combo,
    "lineClusteredColumnComboChart": _render_combo,

    # Funnel
    "funnelChart": _render_funnel,

    # Treemap
    "treemap": _render_treemap,

    # Gauge
    "gauge": _render_gauge,

    # Cards and KPI
    "card": _render_card,
    "cardVisual": _render_card,
    "multiRowCard": _render_card,
    "kpi": _render_kpi,

    # Tables
    "tableEx": _render_table,
    "pivotTable": _render_table,

    # Ribbon (approximate as stacked area)
    "ribbonChart": _render_ribbon,
}

# Visual types to skip (not meaningful as static chart images)
SKIP_TYPES = {
    "slicer", "advancedSlicerVisual",
    "map", "filledMap", "shapeMap", "azureMap",
    "decompositionTreeVisual", "keyDriversVisual", "qnaVisual", "aiNarratives",
    "scriptVisual", "pythonVisual", "paginator", "referenceLabel",
}


# =============================================================================
# NATIVE PPTX CHART HELPERS
# =============================================================================

def _create_presentation():
    """Create a blank 16:9 widescreen Presentation."""
    prs = _PptxFactory()
    prs.slide_width = SLIDE_WIDTH
    prs.slide_height = SLIDE_HEIGHT
    return prs


def _build_category_chart_data(df, categories, values):
    """Build CategoryChartData from DataFrame for bar/column/line/area/pie charts.

    Handles the pivot logic: when 2+ grouping columns exist, the second
    column is pivoted into legend series (like PBI's "Legend" well).

    Returns:
        (CategoryChartData, num_series: int)
    """
    chart_data = CategoryChartData()

    if len(categories) >= 2 and len(values) == 1:
        # Pivot second grouping column into series
        pivot_df = df.pivot_table(
            index=categories[0], columns=categories[1],
            values=values[0], aggfunc="sum"
        ).fillna(0)
        chart_data.categories = [str(c) for c in pivot_df.index]
        for col in pivot_df.columns:
            chart_data.add_series(str(col), pivot_df[col].tolist())
        return chart_data, len(pivot_df.columns)
    else:
        chart_data.categories = df[categories[0]].astype(str).tolist()
        for v in values:
            chart_data.add_series(v, df[v].tolist())
        return chart_data, len(values)


def _build_xy_chart_data(df, spec):
    """Build XyChartData for scatter charts.

    Returns:
        (XyChartData, num_series: int) or (None, 0) if insufficient data
    """
    categories, values = classify_columns(df, spec)
    chart_data = XyChartData()

    if len(values) < 2:
        return None, 0

    num_series = 0
    if categories:
        groups = df.groupby(categories[0])
        for name, group in groups:
            series = chart_data.add_series(str(name))
            for _, row in group.iterrows():
                x_val = float(row[values[0]]) if pd.notna(row[values[0]]) else 0
                y_val = float(row[values[1]]) if pd.notna(row[values[1]]) else 0
                series.add_data_point(x_val, y_val)
            num_series += 1
    else:
        series = chart_data.add_series("Data")
        for _, row in df.iterrows():
            x_val = float(row[values[0]]) if pd.notna(row[values[0]]) else 0
            y_val = float(row[values[1]]) if pd.notna(row[values[1]]) else 0
            series.add_data_point(x_val, y_val)
        num_series = 1

    return chart_data, num_series


def _style_native_chart(chart, spec, num_series, categories=None, values=None):
    """Apply PBI styling to a native python-pptx chart.

    Sets title, legend, series colors, axis formatting, and axis labels
    to match PBI defaults.

    Args:
        chart: python-pptx Chart object
        spec: VisualSpec
        num_series: number of data series (for legend logic)
        categories: list of category column names (for axis labels)
        values: list of measure column names (for axis labels)
    """
    categories = categories or []
    values = values or []

    # Title
    chart.has_title = True
    title_para = chart.chart_title.text_frame.paragraphs[0]
    title_para.text = spec.visual_name
    title_para.font.size = Pt(18)
    title_para.font.name = PBI_FONT
    title_para.font.color.rgb = RGBColor(0x33, 0x33, 0x33)

    # Legend at bottom (show when multiple series)
    if num_series > 1:
        chart.has_legend = True
        chart.legend.position = XL_LEGEND_POSITION.BOTTOM
        chart.legend.include_in_layout = False
        chart.legend.font.size = Pt(10)
        chart.legend.font.name = PBI_FONT
    else:
        chart.has_legend = False

    # Apply PBI colors to each series
    plot = chart.plots[0]
    for i, series in enumerate(plot.series):
        fill = series.format.fill
        fill.solid()
        fill.fore_color.rgb = PBI_RGB_COLORS[i % len(PBI_RGB_COLORS)]

    # Axis font styling and labels
    try:
        cat_ax = chart.category_axis
        cat_ax.tick_labels.font.size = Pt(10)
        cat_ax.tick_labels.font.name = PBI_FONT
        cat_ax.tick_labels.font.color.rgb = RGBColor(0x66, 0x66, 0x66)
        # Category axis label (X-axis for column/line, Y-axis for bar)
        if categories:
            cat_ax.has_title = True
            cat_ax.axis_title.text_frame.paragraphs[0].text = categories[0]
            cat_ax.axis_title.text_frame.paragraphs[0].font.size = Pt(12)
            cat_ax.axis_title.text_frame.paragraphs[0].font.name = PBI_FONT
            cat_ax.axis_title.text_frame.paragraphs[0].font.color.rgb = RGBColor(0x66, 0x66, 0x66)
    except Exception:
        pass  # pie/donut charts have no category axis

    try:
        val_ax = chart.value_axis
        val_ax.tick_labels.font.size = Pt(10)
        val_ax.tick_labels.font.name = PBI_FONT
        val_ax.tick_labels.font.color.rgb = RGBColor(0x66, 0x66, 0x66)
        val_ax.has_major_gridlines = True
        val_ax.major_gridlines.format.line.color.rgb = RGBColor(0xE0, 0xE0, 0xE0)
        # Value axis label
        if len(values) == 1:
            val_ax.has_title = True
            val_ax.axis_title.text_frame.paragraphs[0].text = values[0]
            val_ax.axis_title.text_frame.paragraphs[0].font.size = Pt(12)
            val_ax.axis_title.text_frame.paragraphs[0].font.name = PBI_FONT
            val_ax.axis_title.text_frame.paragraphs[0].font.color.rgb = RGBColor(0x66, 0x66, 0x66)
    except Exception:
        pass  # pie/donut charts have no value axis


def _add_native_chart(slide, df, spec):
    """Add a native python-pptx chart to the slide.

    Routes to the correct chart type based on spec.visual_type.
    Handles bar, column, line, area, pie, donut, and scatter natively.

    Returns:
        True if chart was added successfully, False otherwise
    """
    chart_type_enum = NATIVE_CHART_MAP.get(spec.visual_type)
    if chart_type_enum is None:
        return False

    categories, values = classify_columns(df, spec)

    # --- Scatter/XY charts use XyChartData ---
    if spec.visual_type == "scatterChart":
        chart_data, num_series = _build_xy_chart_data(df, spec)
        if chart_data is None:
            return False
        chart_frame = slide.shapes.add_chart(
            chart_type_enum, CHART_LEFT, CHART_TOP, CHART_WIDTH, CHART_HEIGHT,
            chart_data
        )
        chart = chart_frame.chart
        # Scatter: X-axis = values[0], Y-axis = values[1]
        x_label = [values[0]] if values else []
        y_label = [values[1]] if len(values) >= 2 else []
        _style_native_chart(chart, spec, num_series,
                            categories=x_label, values=y_label)
        # Scatter markers: set size
        for series in chart.plots[0].series:
            series.marker.size = 10
        return True

    # --- Pie/donut: measures-only case (each measure is a slice) ---
    if spec.visual_type in ("pieChart", "donutChart"):
        if not categories and len(values) >= 2:
            chart_data = CategoryChartData()
            chart_data.categories = values
            row = df.iloc[0]
            chart_data.add_series(
                "Values",
                [float(row[v]) if pd.notna(row[v]) else 0 for v in values]
            )
            chart_frame = slide.shapes.add_chart(
                chart_type_enum, CHART_LEFT, CHART_TOP, CHART_WIDTH, CHART_HEIGHT,
                chart_data
            )
            chart = chart_frame.chart
            _style_native_chart(chart, spec, 1)
            # Color individual pie/donut slices
            plot = chart.plots[0]
            for i, point in enumerate(plot.series[0].points):
                point.format.fill.solid()
                point.format.fill.fore_color.rgb = PBI_RGB_COLORS[i % len(PBI_RGB_COLORS)]
            return True

    if not categories or not values:
        return False

    # Sort bar/column charts by first measure descending (matches PBI default)
    if spec.visual_type in ("barChart", "clusteredBarChart",
                            "columnChart", "clusteredColumnChart"):
        if len(values) == 1 and len(categories) == 1:
            df = df.sort_values(values[0], ascending=False)

    # --- Standard category charts (bar, column, line, area, pie, donut) ---
    chart_data, num_series = _build_category_chart_data(df, categories, values)
    chart_frame = slide.shapes.add_chart(
        chart_type_enum, CHART_LEFT, CHART_TOP, CHART_WIDTH, CHART_HEIGHT,
        chart_data
    )
    chart = chart_frame.chart
    _style_native_chart(chart, spec, num_series,
                        categories=categories, values=values)

    # Pie/donut: color individual slices instead of series
    if spec.visual_type in ("pieChart", "donutChart"):
        plot = chart.plots[0]
        for i, point in enumerate(plot.series[0].points):
            point.format.fill.solid()
            point.format.fill.fore_color.rgb = PBI_RGB_COLORS[i % len(PBI_RGB_COLORS)]

    return True


def _add_png_to_slide(slide, png_path):
    """Insert a plotly PNG image onto a slide, centered within the chart area."""
    slide.shapes.add_picture(
        str(png_path), CHART_LEFT, CHART_TOP, CHART_WIDTH, CHART_HEIGHT
    )


# =============================================================================
# CORE API
# =============================================================================

def _build_spec(spec, visual_type, visual_name, grouping_columns,
                measure_columns, y2_columns, page_name):
    """Build a VisualSpec from either an existing spec or individual params."""
    if spec is not None:
        return spec
    if not visual_type:
        raise ValueError("Either spec or visual_type must be provided")
    return VisualSpec(
        page_name=page_name or "",
        visual_name=visual_name or "",
        visual_type=visual_type,
        grouping_columns=grouping_columns or [],
        measure_columns=measure_columns or [],
        y2_columns=y2_columns or [],
    )


def generate_chart(df, spec=None, visual_type=None, visual_name=None,
                   grouping_columns=None, measure_columns=None,
                   y2_columns=None, page_name=""):
    """Generate a plotly Figure for a PBI visual from tabular data.

    This is the programmatic API for plotly/PNG output. Always returns a
    plotly Figure (or None). For PowerPoint output, use generate_chart_pptx().

    Args:
        df: DataFrame with DAX query results
        spec: VisualSpec (optional, takes precedence over individual params)
        visual_type: PBI visual type identifier (e.g., "barChart")
        visual_name: Chart title
        grouping_columns: list of category column names
        measure_columns: list of value/measure column names
        y2_columns: list of secondary-axis measure column names
        page_name: PBI page name (informational)

    Returns:
        plotly Figure object, or None if visual type should be skipped
    """
    spec = _build_spec(spec, visual_type, visual_name, grouping_columns,
                       measure_columns, y2_columns, page_name)

    if spec.visual_type in SKIP_TYPES:
        print(f"  Skipping: {spec.visual_name} ({spec.visual_type}) "
              f"-- not meaningful as static chart")
        return None

    if df is None or df.empty:
        print(f"  Skipping: {spec.visual_name} -- no data")
        return None

    renderer = CHART_TYPE_ROUTER.get(spec.visual_type)
    if renderer is None:
        print(f"  WARNING: Unknown visual type '{spec.visual_type}' for "
              f"'{spec.visual_name}' -- rendering as table fallback")
        renderer = _render_table
    try:
        fig = renderer(df, spec)
        print(f"  Generated: {spec.visual_name} ({spec.visual_type})")
        return fig
    except Exception as e:
        print(f"  ERROR generating chart for '{spec.visual_name}': {e}")
        return None


def generate_chart_pptx(df, spec=None, visual_type=None, visual_name=None,
                        grouping_columns=None, measure_columns=None,
                        y2_columns=None, page_name=""):
    """Generate a single-slide PowerPoint for a PBI visual from tabular data.

    Uses a native python-pptx chart when the visual type is supported (bar,
    column, line, area, pie, donut, scatter). Falls back to rendering a plotly
    PNG and embedding it as a picture on the slide for all other types.

    Args:
        df: DataFrame with DAX query results
        spec: VisualSpec (optional, takes precedence over individual params)
        visual_type: PBI visual type identifier (e.g., "barChart")
        visual_name: Chart title
        grouping_columns: list of category column names
        measure_columns: list of value/measure column names
        y2_columns: list of secondary-axis measure column names
        page_name: PBI page name (informational)

    Returns:
        pptx.presentation.Presentation (one slide), or None if skipped
    """
    spec = _build_spec(spec, visual_type, visual_name, grouping_columns,
                       measure_columns, y2_columns, page_name)

    if spec.visual_type in SKIP_TYPES:
        print(f"  Skipping: {spec.visual_name} ({spec.visual_type}) "
              f"-- not meaningful as static chart")
        return None

    if df is None or df.empty:
        print(f"  Skipping: {spec.visual_name} -- no data")
        return None

    try:
        prs = _create_presentation()
        blank_layout = prs.slide_layouts[6]  # blank slide layout
        slide = prs.slides.add_slide(blank_layout)

        # Try native chart first for supported types
        if spec.visual_type in NATIVE_CHART_MAP:
            success = _add_native_chart(slide, df, spec)
            if success:
                print(f"  Generated (native PPTX): {spec.visual_name} ({spec.visual_type})")
                return prs
            print(f"  Native chart failed for '{spec.visual_name}', using PNG fallback")

        # PNG fallback: render with plotly, insert image on slide
        renderer = CHART_TYPE_ROUTER.get(spec.visual_type)
        if renderer is None:
            print(f"  WARNING: Unknown visual type '{spec.visual_type}' for "
                  f"'{spec.visual_name}' -- rendering as table fallback")
            renderer = _render_table

        fig = renderer(df, spec)
        with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as tmp:
            tmp_path = tmp.name
        try:
            fig.write_image(tmp_path, format="png", width=1100, height=500, scale=2)
            _add_png_to_slide(slide, tmp_path)
        finally:
            os.unlink(tmp_path)

        print(f"  Generated (PNG fallback on PPTX): {spec.visual_name} ({spec.visual_type})")
        return prs

    except Exception as e:
        print(f"  ERROR generating chart for '{spec.visual_name}': {e}")
        return None


def save_chart(fig, output_path, width=1100, height=500, scale=2):
    """Save a plotly Figure as a PNG image.

    Args:
        fig: plotly Figure object
        output_path: file path for the PNG (directory is created if needed)
        width: image width in pixels (before scaling)
        height: image height in pixels (before scaling)
        scale: resolution multiplier (2 = 144 DPI, 3 = 216 DPI)
    """
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    fig.write_image(
        str(output_path),
        format="png",
        width=width,
        height=height,
        scale=scale,
    )
    print(f"  Saved: {output_path}")


def save_chart_pptx(prs, output_path):
    """Save a Presentation as a .pptx file.

    Args:
        prs: pptx.presentation.Presentation object
        output_path: file path for the .pptx (directory is created if needed)
    """
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(output_path))
    print(f"  Saved: {output_path}")


# =============================================================================
# CLI
# =============================================================================

def main():
    parser = argparse.ArgumentParser(
        description="PBI AutoGov -- Chart Generator (Skill 5): "
                    "Generate PBI-style chart images from DAX query results."
    )

    # Required: CSV data file
    parser.add_argument("--csv", required=True,
                        help="Path to CSV file with DAX query tabular data")

    # Mode 1: Metadata-driven
    parser.add_argument("--metadata",
                        help="Path to metadata Excel from Skill 1 (pbi_report_metadata.xlsx)")
    parser.add_argument("--visual",
                        help="Visual name to match in metadata (Mode 1)")

    # Mode 2: Manual/screenshot-driven (mirrors Skill 4's interface)
    parser.add_argument("--visual-type",
                        help="PBI visual type identifier (e.g., barChart, pieChart)")
    parser.add_argument("--visual-name",
                        help="Chart title (Mode 2)")
    parser.add_argument("--field", action="append", dest="fields",
                        help="Field in 'name:role' format, repeatable "
                             "(role = measure, grouping, or y2)")

    # Output options
    parser.add_argument("--format", choices=["pptx", "png"], default="pptx",
                        dest="output_format",
                        help="Output format: 'pptx' (default) for single-slide PowerPoint "
                             "with native chart or PNG fallback, 'png' for legacy plotly image")
    parser.add_argument("--report-name",
                        help="Report name for subfolder organization (e.g., 'Revenue Opportunities' "
                             "creates output/charts/Revenue_Opportunities/)")
    parser.add_argument("--output", default="output/charts/",
                        help="Output directory for chart files (default: output/charts/)")
    parser.add_argument("--width", type=int, default=1100,
                        help="Image width in pixels for PNG mode (default: 1100)")
    parser.add_argument("--height", type=int, default=500,
                        help="Image height in pixels for PNG mode (default: 500)")
    parser.add_argument("--scale", type=int, default=2,
                        help="Resolution scale factor for PNG mode (default: 2 for 144 DPI)")

    args = parser.parse_args()

    # Load CSV data
    csv_path = Path(args.csv)
    if not csv_path.exists():
        print(f"ERROR: CSV file not found: {csv_path}")
        sys.exit(1)
    df = pd.read_csv(csv_path, encoding="utf-8-sig")
    print(f"Loaded CSV: {csv_path} ({len(df)} rows, {len(df.columns)} columns)")

    # Determine mode and build VisualSpec
    if args.metadata and args.visual:
        # Mode 1: Metadata-driven
        print(f"Mode 1: Reading visual metadata from {args.metadata}")
        spec = parse_visual_from_metadata(args.metadata, args.visual)
        if spec is None:
            sys.exit(1)
    elif args.visual_type and args.fields:
        # Mode 2: Manual/screenshot-driven
        print(f"Mode 2: Using CLI-provided visual metadata")
        grouping_cols = []
        measure_cols = []
        y2_cols = []
        for field_str in args.fields:
            if ":" not in field_str:
                print(f"ERROR: Field '{field_str}' must be in 'name:role' format "
                      f"(e.g., 'Revenue:measure')")
                sys.exit(1)
            name, role = field_str.rsplit(":", 1)
            name = name.strip()
            role = role.strip().lower()
            if role == "measure":
                measure_cols.append(name)
            elif role == "grouping":
                grouping_cols.append(name)
            elif role == "y2":
                # y2 fields are measures on the secondary axis
                measure_cols.append(name)
                y2_cols.append(name)
            else:
                print(f"WARNING: Unknown role '{role}' for field '{name}' "
                      f"-- expected 'measure', 'grouping', or 'y2'")

        spec = VisualSpec(
            page_name="",
            visual_name=args.visual_name or args.visual_type,
            visual_type=args.visual_type,
            grouping_columns=grouping_cols,
            measure_columns=measure_cols,
            y2_columns=y2_cols,
        )
    else:
        print("ERROR: Provide either --metadata + --visual (Mode 1) "
              "or --visual-type + --field (Mode 2)")
        parser.print_help()
        sys.exit(1)

    print(f"\nVisual: {spec.visual_name}")
    print(f"  Type: {spec.visual_type}")
    print(f"  Grouping columns: {spec.grouping_columns}")
    print(f"  Measure columns: {spec.measure_columns}")
    if spec.y2_columns:
        print(f"  Y2 columns (secondary axis): {spec.y2_columns}")

    # Sanitize visual name for filename: replace non-alphanumeric chars with underscore
    safe_name = re.sub(r'[^\w\-]', '_', spec.visual_name).strip('_')
    output_dir = Path(args.output)
    # Create report subfolder if --report-name provided
    if args.report_name:
        safe_report = re.sub(r'[^\w\-]', '_', args.report_name).strip('_')
        output_dir = output_dir / safe_report

    # Generate and save chart
    if args.output_format == "pptx":
        prs = generate_chart_pptx(df, spec=spec)
        if prs is None:
            print("No chart generated.")
            sys.exit(0)
        output_path = output_dir / f"{safe_name}.pptx"
        save_chart_pptx(prs, output_path)
    else:
        fig = generate_chart(df, spec=spec)
        if fig is None:
            print("No chart generated.")
            sys.exit(0)
        output_path = output_dir / f"{safe_name}.png"
        save_chart(fig, output_path, width=args.width, height=args.height, scale=args.scale)

    print(f"\nDone! Chart saved to: {output_path}")


if __name__ == "__main__":
    main()
