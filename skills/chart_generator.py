# -*- coding: utf-8 -*-
"""
Skill 5: chart_generator.py
PBI AutoGov -- Chart Generator

Generates chart images (PNG) from DAX query tabular data that visually
resemble the original Power BI visuals. Uses plotly for all chart types
with a PBI-styled theme.

Two input modes:
  Mode 1 (CSV + Metadata): reads metadata Excel to get visual type + field roles
  Mode 2 (CSV + Screenshot): agent passes visual type + fields as CLI args

Input:  CSV file with DAX query results + visual metadata (Excel or CLI args)
Output: PNG chart image saved to output directory

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
from dataclasses import dataclass, field
from pathlib import Path

import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots


# =============================================================================
# CONSTANTS
# =============================================================================

# Power BI default color palette (10 accent colors)
PBI_COLORS = [
    "#118DFF", "#12239E", "#E66C37", "#6B007B", "#E044A7",
    "#744EC2", "#D9B300", "#D64550", "#197278", "#1AAB40",
]

PBI_FONT = "Segoe UI"


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
    )
    fig.update_yaxes(autorange="reversed")  # top-to-bottom ordering like PBI
    return fig


# ---- Column chart (vertical) ----

def _render_column(df, spec):
    """Vertical column chart. Categories on X-axis, values on Y-axis."""
    categories, values = classify_columns(df, spec)
    if not categories or not values:
        return _render_table(df, spec)

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
    )
    return fig


# ---- Pie chart ----

def _render_pie(df, spec):
    """Pie chart. Labels from first grouping column, values from first measure.

    Only uses first measure (pie charts show one value series). Warns if multiple.
    """
    categories, values = classify_columns(df, spec)
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
    """Donut chart. Same as pie but with a hole in the center (hole=0.4)."""
    categories, values = classify_columns(df, spec)
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
    fig.update_xaxes(showgrid=False, linecolor="#E0E0E0")
    fig.update_yaxes(showgrid=True, gridcolor="#E0E0E0", linecolor="#E0E0E0")
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
# CORE API
# =============================================================================

def generate_chart(df, spec=None, visual_type=None, visual_name=None,
                   grouping_columns=None, measure_columns=None,
                   y2_columns=None, page_name=""):
    """Generate a plotly Figure for a PBI visual from tabular data.

    This is the main programmatic API. Accepts either a VisualSpec or
    individual parameters. Routes to the appropriate renderer based on
    visual type.

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
    if spec is None:
        if not visual_type:
            raise ValueError("Either spec or visual_type must be provided")
        spec = VisualSpec(
            page_name=page_name or "",
            visual_name=visual_name or "",
            visual_type=visual_type,
            grouping_columns=grouping_columns or [],
            measure_columns=measure_columns or [],
            y2_columns=y2_columns or [],
        )

    if spec.visual_type in SKIP_TYPES:
        print(f"  Skipping: {spec.visual_name} ({spec.visual_type}) "
              f"-- not meaningful as static chart")
        return None

    renderer = CHART_TYPE_ROUTER.get(spec.visual_type)
    if renderer is None:
        print(f"  WARNING: Unknown visual type '{spec.visual_type}' for "
              f"'{spec.visual_name}' -- rendering as table fallback")
        renderer = _render_table

    if df is None or df.empty:
        print(f"  Skipping: {spec.visual_name} -- no data")
        return None

    try:
        fig = renderer(df, spec)
        print(f"  Generated: {spec.visual_name} ({spec.visual_type})")
        return fig
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
    parser.add_argument("--output", default="output/charts/",
                        help="Output directory for chart images (default: output/charts/)")
    parser.add_argument("--width", type=int, default=1100,
                        help="Image width in pixels (default: 1100)")
    parser.add_argument("--height", type=int, default=500,
                        help="Image height in pixels (default: 500)")
    parser.add_argument("--scale", type=int, default=2,
                        help="Resolution scale factor (default: 2 for 144 DPI)")

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

    # Generate chart
    fig = generate_chart(df, spec=spec)
    if fig is None:
        print("No chart generated.")
        sys.exit(0)

    # Save chart image
    # Sanitize visual name for filename: replace non-alphanumeric chars with underscore
    safe_name = re.sub(r'[^\w\-]', '_', spec.visual_name).strip('_')
    output_dir = Path(args.output)
    output_path = output_dir / f"{safe_name}.png"

    save_chart(fig, output_path, width=args.width, height=args.height, scale=args.scale)
    print(f"\nDone! Chart saved to: {output_path}")


if __name__ == "__main__":
    main()
