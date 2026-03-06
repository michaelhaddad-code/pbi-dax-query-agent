# -*- coding: utf-8 -*-
"""
Skill 5: Local DAX Executor
============================
Executes DAX queries locally using pandas against pre-extracted table data.

Reads structured metadata from Skill 1/2 Excel (field roles, measure formulas,
table/column references) and reassembles them into pandas operations — no DAX
string parsing required.

Data sources:
  - Pre-extracted table CSVs (from pbixray or manual export)
  - Direct .pbix extraction via pbixray PBIXRay.get_table()

Output:
  - Per-visual CSV files + csv_manifest.json (consumable by page_builder.py)

Usage:
    # From pre-extracted CSVs
    python skills/dax_executor.py \\
        --tables-dir "output/raw_data/" \\
        --metadata "output/Regional_Sales_Sample_metadata.xlsx" \\
        --model-root "data/Regional Sales Sample.SemanticModel/definition" \\
        --page "Sales Overview" \\
        --output "output/sales_overview_csv/"

    # From .pbix (extracts tables automatically)
    python skills/dax_executor.py \\
        --pbix "data/Regional Sales Sample.pbix" \\
        --metadata "output/Regional_Sales_Sample_metadata.xlsx" \\
        --page "Sales Overview" \\
        --output "output/sales_overview_csv/"
"""

import argparse
import json
import math
import numbers
import os
import re
import sys
import warnings
from collections import defaultdict, deque
from dataclasses import dataclass, field as dc_field
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple, Union

import numpy as np
import pandas as pd


def _is_numeric(val: Any) -> bool:
    """Check if a value is numeric (handles numpy scalars in numpy >=2.0)."""
    return isinstance(val, (int, float, numbers.Number))

# Add project root so sibling modules can be imported
_SKILL_DIR = Path(__file__).resolve().parent
_PROJECT_ROOT = _SKILL_DIR.parent
if str(_PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(_PROJECT_ROOT))

from skills.tmdl_parser import SemanticModel, TmdlRelationship, parse_semantic_model


# =============================================================================
# Data classes
# =============================================================================

@dataclass
class TableStore:
    """In-memory store of all table data plus semantic model metadata."""
    tables: Dict[str, pd.DataFrame] = dc_field(default_factory=dict)
    relationships: List[TmdlRelationship] = dc_field(default_factory=list)
    measures: Dict[Tuple[str, str], str] = dc_field(default_factory=dict)
    model: Optional[SemanticModel] = None

    def get_table(self, name: str) -> Optional[pd.DataFrame]:
        """Case-insensitive table lookup."""
        if name in self.tables:
            return self.tables[name]
        for k, v in self.tables.items():
            if k.lower() == name.lower():
                return v
        return None

    def table_name_normalized(self, name: str) -> str:
        """Return the canonical table name (preserving original casing)."""
        if name in self.tables:
            return name
        for k in self.tables:
            if k.lower() == name.lower():
                return k
        return name


@dataclass
class ExecResult:
    """Result of executing a single visual's query."""
    page_name: str
    visual_name: str
    visual_id: str
    visual_type: str
    dax_pattern: str
    df: Optional[pd.DataFrame] = None
    error: str = ""
    warnings: List[str] = dc_field(default_factory=list)


# =============================================================================
# Data loading
# =============================================================================

def load_tables_from_csv_dir(csv_dir: str) -> Dict[str, pd.DataFrame]:
    """Load per-table CSVs from a directory. Filename (without extension) = table name."""
    csv_dir = Path(csv_dir)
    tables = {}
    for csv_path in sorted(csv_dir.glob("*.csv")):
        # Skip manifest and internal files
        if csv_path.name.startswith("_"):
            continue
        table_name = csv_path.stem
        try:
            df = pd.read_csv(csv_path, encoding="utf-8-sig", low_memory=False)
            tables[table_name] = df
            print(f"  Loaded {table_name}: {len(df)} rows, {len(df.columns)} cols")
        except Exception as e:
            print(f"  WARNING: Failed to load {csv_path.name}: {e}")
    return tables


def load_tables_from_pbix(pbix_path: str) -> Dict[str, pd.DataFrame]:
    """Extract all tables from a .pbix file via pbixray."""
    try:
        from pbixray import PBIXRay
    except ImportError:
        print("ERROR: pbixray is required for .pbix table extraction.")
        print("Install with: pip install pbixray")
        sys.exit(1)

    pbix_path = Path(pbix_path)
    if not pbix_path.is_file():
        print(f"ERROR: .pbix file not found: {pbix_path}")
        sys.exit(1)

    print(f"Extracting tables from: {pbix_path.name}")
    model = PBIXRay(str(pbix_path))
    tables = {}

    for table_name in model.tables:
        try:
            raw_df = model.get_table(table_name)
            if raw_df is None or raw_df.empty:
                print(f"  Skipped {table_name}: empty")
                continue

            # Fix datetime columns: pbixray returns days-since-epoch as floats
            for col in raw_df.columns:
                sample = raw_df[col].dropna().head(5)
                if sample.empty:
                    continue
                # Detect columns that look like date offsets (large integers/floats)
                # and have "date" or "day" in name
                col_lower = col.lower()
                if any(kw in col_lower for kw in ["date", "day", "week", "month", "year month"]):
                    if pd.api.types.is_numeric_dtype(sample):
                        first_val = sample.iloc[0]
                        # Vertipaq stores dates as days since 1899-12-30
                        if isinstance(first_val, (int, float)) and 30000 < abs(first_val) < 100000:
                            try:
                                raw_df[col] = pd.Timestamp("1899-12-30") + pd.to_timedelta(
                                    raw_df[col], unit="D"
                                )
                            except Exception:
                                pass

            # Fix Decimal columns
            for col in raw_df.columns:
                sample = raw_df[col].dropna().head(1)
                if not sample.empty:
                    val = sample.iloc[0]
                    type_name = type(val).__name__
                    if type_name == "Decimal":
                        raw_df[col] = raw_df[col].apply(
                            lambda x: float(x) / 10000 if pd.notna(x) else x
                        )

            tables[table_name] = raw_df
            print(f"  Extracted {table_name}: {len(raw_df)} rows, {len(raw_df.columns)} cols")
        except Exception as e:
            print(f"  WARNING: Failed to extract {table_name}: {e}")

    return tables


def build_table_store(
    pbix_path: Optional[str] = None,
    tables_dir: Optional[str] = None,
    model_root: Optional[str] = None,
    metadata_measures: Optional[Dict[Tuple[str, str], str]] = None,
) -> TableStore:
    """Unified loader: build a TableStore from available sources.

    Args:
        pbix_path: Path to .pbix file (extracts tables via pbixray)
        tables_dir: Path to directory of per-table CSVs
        model_root: Path to SemanticModel/definition (for relationships + measures)
        metadata_measures: Measures extracted from metadata Excel (fallback)
    """
    # Load tables
    if tables_dir:
        print(f"Loading tables from CSV directory: {tables_dir}")
        tables = load_tables_from_csv_dir(tables_dir)
    elif pbix_path:
        tables = load_tables_from_pbix(pbix_path)
    else:
        print("ERROR: Must provide either --tables-dir or --pbix")
        sys.exit(1)

    print(f"Loaded {len(tables)} tables: {', '.join(sorted(tables.keys()))}")

    # Load semantic model for relationships and measures
    model = None
    relationships = []
    measures = {}

    if model_root and Path(model_root).is_dir():
        print(f"Parsing semantic model: {model_root}")
        model = parse_semantic_model(model_root)
        relationships = model.relationships
        measures = dict(model.measures)
        print(f"  {len(measures)} measures, {len(relationships)} relationships")
    elif metadata_measures:
        measures = dict(metadata_measures)
        print(f"  Using {len(measures)} measures from metadata Excel")

    # Ensure datetime columns are parsed in CSV-loaded tables
    for tname, df in tables.items():
        for col in df.columns:
            if df[col].dtype == object:
                sample = df[col].dropna().head(5)
                if not sample.empty:
                    # Check if values look like dates: "2021-08-17" pattern
                    first = str(sample.iloc[0])
                    if re.match(r"^\d{4}-\d{2}-\d{2}", first):
                        try:
                            tables[tname][col] = pd.to_datetime(df[col], errors="coerce")
                        except Exception:
                            pass

    return TableStore(
        tables=tables,
        relationships=relationships,
        measures=measures,
        model=model,
    )


# =============================================================================
# Star schema join
# =============================================================================

def _build_adjacency(relationships: List[TmdlRelationship]) -> Dict[str, List[dict]]:
    """Build an undirected adjacency graph from relationships.

    Each edge stores: neighbor, from_col (on current side), to_col (on neighbor side).
    """
    adj: Dict[str, List[dict]] = defaultdict(list)
    for rel in relationships:
        if not rel.is_active:
            continue
        # from_table (many side) -> to_table (one side)
        adj[rel.from_table].append({
            "neighbor": rel.to_table,
            "left_col": rel.from_column,
            "right_col": rel.to_column,
        })
        # Reverse edge
        adj[rel.to_table].append({
            "neighbor": rel.from_table,
            "left_col": rel.to_column,
            "right_col": rel.from_column,
        })
    return adj


def _detect_fact_table(relationships: List[TmdlRelationship], tables_needed: set) -> str:
    """Detect the fact table: the one that appears most often on the 'from' (many) side."""
    from_counts: Dict[str, int] = defaultdict(int)
    for rel in relationships:
        if not rel.is_active:
            continue
        if rel.from_table in tables_needed:
            from_counts[rel.from_table] += 1
        if rel.to_table in tables_needed:
            from_counts[rel.to_table] += 0  # ensure it's counted
    if not from_counts:
        return next(iter(tables_needed)) if tables_needed else ""
    return max(from_counts, key=from_counts.get)


def build_star_join(
    store: TableStore,
    tables_needed: set,
    fact_table: Optional[str] = None,
) -> pd.DataFrame:
    """Build a joined DataFrame by BFS from the fact table through relationships.

    Uses LEFT JOIN so dimension rows missing from fact are not lost.
    """
    if not tables_needed:
        return pd.DataFrame()

    # Normalize table names
    normalized = set()
    for t in tables_needed:
        normalized.add(store.table_name_normalized(t))
    tables_needed = normalized

    if not fact_table:
        fact_table = _detect_fact_table(store.relationships, tables_needed)
    fact_table = store.table_name_normalized(fact_table)

    fact_df = store.get_table(fact_table)
    if fact_df is None:
        raise ValueError(f"Fact table '{fact_table}' not found in store")

    if len(tables_needed) == 1:
        return fact_df.copy()

    adj = _build_adjacency(store.relationships)

    # BFS to find shortest paths from fact_table to each needed table
    visited = {fact_table}
    queue = deque([fact_table])
    # parent[table] = (parent_table, left_col_on_parent, right_col_on_table)
    parent: Dict[str, Tuple[str, str, str]] = {}

    while queue:
        current = queue.popleft()
        for edge in adj.get(current, []):
            neighbor = edge["neighbor"]
            if neighbor not in visited:
                visited.add(neighbor)
                parent[neighbor] = (current, edge["left_col"], edge["right_col"])
                queue.append(neighbor)

    # Build the join order: for each needed table (except fact), trace path back to fact
    result = fact_df.copy()
    joined = {fact_table}

    # Prefix columns to avoid ambiguity: "TableName.ColumnName"
    # Actually, keep original column names but handle duplicates with suffixes
    for target in tables_needed:
        if target in joined:
            continue
        if target not in parent:
            warnings.warn(f"No relationship path from '{fact_table}' to '{target}'")
            continue

        # Trace path from target back to fact
        path = []
        current = target
        while current in parent:
            par, left_col, right_col = parent[current]
            path.append((par, left_col, current, right_col))
            current = par
        path.reverse()

        # Execute joins along the path
        for par, left_col, child, right_col in path:
            if child in joined:
                continue
            child_df = store.get_table(child)
            if child_df is None:
                warnings.warn(f"Table '{child}' not found in store, skipping join")
                continue

            # Determine join columns: left_col should be in result (parent side),
            # right_col should be in child_df
            if left_col not in result.columns:
                # Try the reverse
                if right_col in result.columns and left_col in child_df.columns:
                    left_col, right_col = right_col, left_col
                else:
                    warnings.warn(
                        f"Join column '{left_col}' not found in joined data "
                        f"for {par} -> {child}"
                    )
                    continue

            if right_col not in child_df.columns:
                warnings.warn(
                    f"Join column '{right_col}' not found in table '{child}'"
                )
                continue

            # Coerce join columns to same dtype
            try:
                if result[left_col].dtype != child_df[right_col].dtype:
                    result[left_col] = result[left_col].astype(str)
                    child_df = child_df.copy()
                    child_df[right_col] = child_df[right_col].astype(str)
            except Exception:
                pass

            result = result.merge(
                child_df,
                left_on=left_col,
                right_on=right_col,
                how="left",
                suffixes=("", f"__{child}"),
            )
            joined.add(child)

    return result


# =============================================================================
# Measure evaluator
# =============================================================================

# Regex patterns for parsing DAX formulas
_RE_COLUMN_REF = re.compile(
    r"(?:'([^']+)'\[([^\]]+)\])"     # 'Table'[Column]
    r"|"
    r"(?:(\w[\w ]*)\[([^\]]+)\])"    # Table[Column] (no quotes)
)
_RE_MEASURE_REF = re.compile(
    r"(?<!')\[([^\]]+)\]"            # [MeasureName] not preceded by quote
)
_RE_CALCULATE = re.compile(
    r"\bCALCULATE\s*\(", re.IGNORECASE
)
_RE_SUMX = re.compile(
    r"\bSUMX\s*\(\s*(\w[\w ]*)\s*,\s*(.+?)\)\s*\)", re.IGNORECASE | re.DOTALL
)
_RE_COUNTAX = re.compile(
    r"\bCOUNTAX\s*\(", re.IGNORECASE
)
_RE_AVERAGEX = re.compile(
    r"\bAVERAGEX\s*\(", re.IGNORECASE
)
_RE_FILTER = re.compile(
    r"\bFILTER\s*\(", re.IGNORECASE
)
_RE_KEEPFILTERS = re.compile(
    r"\bKEEPFILTERS\s*\(", re.IGNORECASE
)
_RE_SELECTEDVALUE = re.compile(
    r"\bSELECTEDVALUE\s*\(", re.IGNORECASE
)
_RE_VAR_RETURN = re.compile(
    r"\bVAR\b", re.IGNORECASE
)
_RE_IF = re.compile(
    r"\bIF\s*\(", re.IGNORECASE
)
_RE_ISBLANK = re.compile(
    r"\bISBLANK\s*\(", re.IGNORECASE
)
_RE_MROUND = re.compile(
    r"\bMROUND\s*\(", re.IGNORECASE
)
_RE_LEFT_FUNC = re.compile(
    r"\bLEFT\s*\(", re.IGNORECASE
)
_RE_VALUE_FUNC = re.compile(
    r"\bVALUE\s*\(", re.IGNORECASE
)
_RE_CONCATENATE = re.compile(
    r"\bCONCATE(?:NATE)?\s*\(", re.IGNORECASE
)
_RE_SUM = re.compile(r"\bSUM\s*\(", re.IGNORECASE)
_RE_COUNT = re.compile(r"\bCOUNT\s*\(", re.IGNORECASE)
_RE_COUNTA = re.compile(r"\bCOUNTA\s*\(", re.IGNORECASE)
_RE_AVERAGE = re.compile(r"\bAVERAGE\s*\(", re.IGNORECASE)
_RE_MIN = re.compile(r"\bMIN\s*\(", re.IGNORECASE)
_RE_MAX = re.compile(r"\bMAX\s*\(", re.IGNORECASE)


def _find_matching_paren(text: str, start: int) -> int:
    """Find the index of the closing paren matching the opening one at 'start'."""
    depth = 0
    in_str = False
    str_char = None
    for i in range(start, len(text)):
        ch = text[i]
        if in_str:
            if ch == str_char:
                in_str = False
            continue
        if ch in ('"', "'"):
            in_str = True
            str_char = ch
        elif ch == '(':
            depth += 1
        elif ch == ')':
            depth -= 1
            if depth == 0:
                return i
    return len(text) - 1


def _extract_func_args(text: str, func_start: int) -> Tuple[str, int]:
    """Extract the full arguments string of a function call starting at func_start.

    func_start should point to the opening '(' character.
    Returns (args_string, end_index).
    """
    end = _find_matching_paren(text, func_start)
    inner = text[func_start + 1: end]
    return inner, end


def _split_top_level_args(text: str) -> List[str]:
    """Split a comma-separated argument list respecting nested parentheses and strings."""
    args = []
    depth = 0
    current = []
    in_str = False
    str_char = None
    for ch in text:
        if in_str:
            current.append(ch)
            if ch == str_char:
                in_str = False
            continue
        if ch in ('"', "'"):
            in_str = True
            str_char = ch
            current.append(ch)
        elif ch == '(':
            depth += 1
            current.append(ch)
        elif ch == ')':
            depth -= 1
            current.append(ch)
        elif ch == ',' and depth == 0:
            args.append(''.join(current).strip())
            current = []
        else:
            current.append(ch)
    if current:
        args.append(''.join(current).strip())
    return args


def _resolve_column_ref(ref_str: str) -> Tuple[Optional[str], Optional[str]]:
    """Parse a column reference like 'Table'[Column] or Table[Column].

    Returns (table_name, column_name) or (None, None).
    """
    m = _RE_COLUMN_REF.search(ref_str)
    if m:
        if m.group(1):
            return m.group(1), m.group(2)
        return m.group(3), m.group(4)
    return None, None


class MeasureEvaluator:
    """Evaluates DAX measure formulas using pandas DataFrames.

    Not a full DAX engine — covers patterns that appear in the target reports:
    SUM, SUMX, COUNT, COUNTA, COUNTAX, AVERAGE, AVERAGEX, MIN, MAX,
    CALCULATE, FILTER, KEEPFILTERS, VAR/RETURN, IF, ISBLANK,
    SELECTEDVALUE, MROUND, LEFT, VALUE, CONCATENATE, and arithmetic.
    """

    def __init__(self, store: TableStore):
        self.store = store
        self._eval_depth = 0
        self._max_depth = 20

    def lookup_formula(self, measure_name: str, table_hint: str = "") -> Optional[str]:
        """Look up the DAX formula for a measure by name."""
        # Exact match with table hint
        if table_hint:
            key = (table_hint, measure_name)
            if key in self.store.measures:
                return self.store.measures[key]

        # Case-insensitive search
        lower = measure_name.lower()
        for (t, m), formula in self.store.measures.items():
            if m.lower() == lower:
                if not table_hint or t.lower() == table_hint.lower():
                    return formula
        # Fallback: any table
        for (t, m), formula in self.store.measures.items():
            if m.lower() == lower:
                return formula
        return None

    def evaluate(
        self,
        measure_name: str,
        df: pd.DataFrame,
        table_name: str = "",
        context: Optional[Dict] = None,
    ) -> Union[float, int, str, None]:
        """Evaluate a measure in the context of the given DataFrame.

        Args:
            measure_name: The measure name (without brackets)
            df: The DataFrame representing the current filter context
            table_name: Hint for which table the measure belongs to
            context: Optional dict of additional context (e.g. current group values)

        Returns:
            Scalar result (float, int, str) or None on failure
        """
        self._eval_depth += 1
        if self._eval_depth > self._max_depth:
            self._eval_depth -= 1
            return None

        try:
            formula = self.lookup_formula(measure_name, table_name)
            if formula is None:
                return None
            return self._eval_formula(formula, df, table_name, context)
        finally:
            self._eval_depth -= 1

    def _eval_formula(
        self,
        formula: str,
        df: pd.DataFrame,
        table_name: str = "",
        context: Optional[Dict] = None,
    ) -> Union[float, int, str, None]:
        """Evaluate a DAX formula string against a DataFrame."""
        formula = formula.strip()

        # Strip DAX line comments (-- to end of line)
        formula = re.sub(r'--[^\n]*', '', formula).strip()

        # Handle VAR/RETURN blocks
        if _RE_VAR_RETURN.search(formula):
            return self._eval_var_return(formula, df, table_name, context)

        # Handle IF(...)
        if _RE_IF.search(formula) and not _RE_CALCULATE.search(formula[:formula.upper().find("IF")]):
            return self._eval_if(formula, df, table_name, context)

        # Handle CALCULATE(...)
        if _RE_CALCULATE.match(formula):
            return self._eval_calculate_expr(formula, df, table_name, context)

        # Handle SUMX(table, expr)
        m_sumx = re.match(
            r"SUMX\s*\(\s*(.+?)\s*,\s*(.+)\s*\)$",
            formula, re.IGNORECASE | re.DOTALL
        )
        if m_sumx:
            return self._eval_sumx(formula, df, table_name, context)

        # Handle COUNTAX(table_or_filter, expr)
        if _RE_COUNTAX.match(formula):
            return self._eval_countax(formula, df, table_name, context)

        # Handle AVERAGEX(table, expr)
        if _RE_AVERAGEX.match(formula):
            return self._eval_averagex(formula, df, table_name, context)

        # Handle SELECTEDVALUE(...)
        if _RE_SELECTEDVALUE.match(formula):
            return self._eval_selectedvalue(formula, df, table_name)

        # Handle simple aggregations: SUM('T'[C]), COUNT('T'[C]), etc.
        result = self._try_simple_agg(formula, df)
        if result is not None:
            return result

        # Handle MROUND(...)
        if _RE_MROUND.match(formula):
            return self._eval_mround(formula, df, table_name, context)

        # Handle CONCATENATE(...)
        if _RE_CONCATENATE.match(formula):
            return self._eval_concatenate(formula, df, table_name, context)

        # Handle arithmetic expressions with measure references
        if re.search(r'[\+\-\*/]', formula) and _RE_MEASURE_REF.search(formula):
            return self._eval_arithmetic(formula, df, table_name, context)

        # Handle parenthesized expression
        if formula.startswith('(') and formula.endswith(')'):
            return self._eval_formula(formula[1:-1].strip(), df, table_name, context)

        # Handle bare column or table-qualified measure reference
        table, col = _resolve_column_ref(formula)
        if table and col:
            resolved = self._resolve_column_or_measure(table, col, df)
            if resolved is not None:
                if isinstance(resolved, pd.Series):
                    return resolved.sum()  # Default: sum
                return resolved

        # Handle bare measure reference [MeasureName]
        m_ref = _RE_MEASURE_REF.match(formula)
        if m_ref:
            return self.evaluate(m_ref.group(1), df, table_name, context)

        # Handle numeric literal
        try:
            return float(formula)
        except (ValueError, TypeError):
            pass

        return None

    def _eval_var_return(
        self,
        formula: str,
        df: pd.DataFrame,
        table_name: str,
        context: Optional[Dict] = None,
    ) -> Union[float, int, str, None]:
        """Evaluate VAR ... RETURN ... blocks."""
        # Parse all VAR bindings and the RETURN expression
        var_bindings: Dict[str, Any] = {}
        remaining = formula.strip()

        while True:
            # Match VAR <name> = <expr> (stops at next VAR or RETURN)
            m = re.match(
                r"VAR\s+(\w+)\s*=\s*",
                remaining,
                re.IGNORECASE
            )
            if not m:
                break
            var_name = m.group(1)
            rest = remaining[m.end():]

            # Find the extent of this VAR's expression: up to next top-level VAR or RETURN
            expr_end = self._find_var_boundary(rest)
            var_expr = rest[:expr_end].strip()
            remaining = rest[expr_end:].strip()

            # Substitute earlier VAR bindings into this VAR's expression
            # (later VARs can reference earlier ones)
            if var_bindings:
                var_expr = self._substitute_vars(var_expr, var_bindings)

            # Evaluate the VAR expression
            val = self._eval_formula(var_expr, df, table_name, context)
            var_bindings[var_name.lower()] = val

        # Parse RETURN expression
        m_ret = re.match(r"RETURN\s+", remaining, re.IGNORECASE)
        if m_ret:
            return_expr = remaining[m_ret.end():].strip()
        else:
            return_expr = remaining.strip()

        # Substitute VAR references in RETURN expression and evaluate
        return self._eval_with_vars(return_expr, var_bindings, df, table_name, context)

    def _substitute_vars(self, expr: str, var_bindings: Dict[str, Any]) -> str:
        """Replace VAR name references in an expression with their numeric values."""
        result = expr
        for vname, vval in var_bindings.items():
            pattern = re.compile(r'\b' + re.escape(vname) + r'\b', re.IGNORECASE)
            if vval is None:
                result = pattern.sub("0", result)
            elif _is_numeric(vval):
                result = pattern.sub(str(float(vval)), result)
        return result

    def _find_var_boundary(self, text: str) -> int:
        """Find where the current VAR expression ends (next top-level VAR or RETURN)."""
        depth = 0
        in_str = False
        str_char = None
        i = 0
        while i < len(text):
            ch = text[i]
            if in_str:
                if ch == str_char:
                    in_str = False
                i += 1
                continue
            if ch in ('"', "'"):
                in_str = True
                str_char = ch
                i += 1
                continue
            if ch == '(':
                depth += 1
            elif ch == ')':
                depth -= 1

            if depth == 0:
                # Check for top-level VAR or RETURN keyword
                rest = text[i:]
                if re.match(r'\bVAR\b', rest, re.IGNORECASE):
                    return i
                if re.match(r'\bRETURN\b', rest, re.IGNORECASE):
                    return i
            i += 1
        return len(text)

    def _eval_with_vars(
        self,
        expr: str,
        var_bindings: Dict[str, Any],
        df: pd.DataFrame,
        table_name: str,
        context: Optional[Dict] = None,
    ) -> Union[float, int, str, None]:
        """Evaluate an expression that may reference VAR bindings."""
        # Replace VAR references (bare identifiers) with their values
        # First check if it's a simple arithmetic expression referencing vars
        simplified = expr.strip()

        # Try to evaluate as arithmetic with var substitution
        # Replace variable names with their numeric values
        result_expr = simplified
        for vname, vval in var_bindings.items():
            # Replace the variable name as a whole word (case-insensitive)
            pattern = re.compile(r'\b' + re.escape(vname) + r'\b', re.IGNORECASE)
            if vval is None:
                result_expr = pattern.sub("0", result_expr)
            elif _is_numeric(vval):
                result_expr = pattern.sub(str(float(vval)), result_expr)

        # Check for measure references
        measure_refs = _RE_MEASURE_REF.findall(result_expr)
        for mref in measure_refs:
            mval = self.evaluate(mref, df, table_name, context)
            if mval is None:
                mval = 0
            result_expr = result_expr.replace(f"[{mref}]", str(float(mval)))

        # Check for column/measure references ('Table'[Name])
        for m in _RE_COLUMN_REF.finditer(result_expr):
            tbl = m.group(1) or m.group(3)
            col_name = m.group(2) or m.group(4)
            resolved = self._resolve_column_or_measure(tbl, col_name, df)
            if resolved is not None:
                if isinstance(resolved, pd.Series):
                    val = resolved.iloc[0] if len(resolved) == 1 else resolved.sum()
                else:
                    val = resolved
                try:
                    result_expr = result_expr.replace(m.group(0), str(float(val)))
                except (ValueError, TypeError):
                    result_expr = result_expr.replace(m.group(0), "0")

        # Handle IF expression
        if _RE_IF.search(result_expr):
            return self._eval_if(result_expr, df, table_name, context)

        # Try safe arithmetic eval
        return self._safe_eval_arithmetic(result_expr)

    def _eval_if(
        self,
        formula: str,
        df: pd.DataFrame,
        table_name: str,
        context: Optional[Dict] = None,
    ) -> Union[float, int, str, None]:
        """Evaluate IF(condition, true_result, false_result)."""
        m = re.match(r"IF\s*\(", formula, re.IGNORECASE)
        if not m:
            return None

        paren_start = formula.index('(', m.start())
        args_str, _ = _extract_func_args(formula, paren_start)
        args = _split_top_level_args(args_str)
        if len(args) < 2:
            return None

        condition = args[0].strip()
        true_val = args[1].strip()
        false_val = args[2].strip() if len(args) > 2 else "0"

        # Evaluate condition
        cond_result = self._eval_condition(condition, df, table_name, context)

        if cond_result:
            return self._eval_formula(true_val, df, table_name, context)
        else:
            return self._eval_formula(false_val, df, table_name, context)

    def _eval_condition(
        self,
        condition: str,
        df: pd.DataFrame,
        table_name: str,
        context: Optional[Dict] = None,
    ) -> bool:
        """Evaluate a boolean condition (comparisons, ISBLANK, etc.)."""
        condition = condition.strip()

        # ISBLANK(expr)
        if _RE_ISBLANK.match(condition):
            paren_start = condition.index('(')
            inner, _ = _extract_func_args(condition, paren_start)
            val = self._eval_formula(inner.strip(), df, table_name, context)
            return val is None or (_is_numeric(val) and (math.isnan(float(val)) or val == 0))

        # Comparison operators: >, >=, <, <=, <>, =
        for op in [">=", "<=", "<>", ">", "<", "="]:
            parts = condition.split(op, 1)
            if len(parts) == 2:
                left_val = self._eval_formula(parts[0].strip(), df, table_name, context)
                right_val = self._eval_formula(parts[1].strip(), df, table_name, context)
                if left_val is None:
                    left_val = 0
                if right_val is None:
                    right_val = 0
                try:
                    left_val = float(left_val)
                    right_val = float(right_val)
                except (ValueError, TypeError):
                    return str(left_val) == str(right_val) if op == "=" else False
                if op == ">":
                    return left_val > right_val
                if op == ">=":
                    return left_val >= right_val
                if op == "<":
                    return left_val < right_val
                if op == "<=":
                    return left_val <= right_val
                if op == "<>":
                    return left_val != right_val
                if op == "=":
                    return left_val == right_val

        # Fallback: evaluate as expression and check truthiness
        val = self._eval_formula(condition, df, table_name, context)
        return bool(val)

    def _eval_calculate_expr(
        self,
        formula: str,
        df: pd.DataFrame,
        table_name: str,
        context: Optional[Dict] = None,
    ) -> Union[float, int, str, None]:
        """Evaluate CALCULATE(expr, filter1, filter2, ...)."""
        m = re.match(r"CALCULATE\s*\(", formula, re.IGNORECASE)
        if not m:
            return None

        paren_start = formula.index('(', m.start())
        args_str, _ = _extract_func_args(formula, paren_start)
        args = _split_top_level_args(args_str)

        if not args:
            return None

        inner_expr = args[0].strip()
        filter_args = [a.strip() for a in args[1:]]

        # Apply filters to df
        filtered_df = self._apply_filter_args(filter_args, df, table_name)

        # Evaluate inner expression on filtered df
        return self._eval_formula(inner_expr, filtered_df, table_name, context)

    def _apply_filter_args(
        self,
        filter_args: List[str],
        df: pd.DataFrame,
        table_name: str,
    ) -> pd.DataFrame:
        """Apply CALCULATE filter arguments to a DataFrame."""
        result = df.copy()
        for farg in filter_args:
            farg = farg.strip()
            if not farg:
                continue

            # FILTER(table, condition) or FILTER(KEEPFILTERS(table), condition)
            if re.match(r"FILTER\s*\(", farg, re.IGNORECASE):
                paren_start = farg.index('(')
                inner, _ = _extract_func_args(farg, paren_start)
                parts = _split_top_level_args(inner)
                if len(parts) >= 2:
                    condition = parts[1].strip()
                    mask = self._eval_row_condition(condition, result, table_name)
                    if mask is not None:
                        result = result[mask].copy()
            else:
                # Direct column filter: 'Table'[Col] = "value"
                mask = self._eval_row_condition(farg, result, table_name)
                if mask is not None:
                    result = result[mask].copy()

        return result

    def _eval_row_condition(
        self,
        condition: str,
        df: pd.DataFrame,
        table_name: str,
    ) -> Optional[pd.Series]:
        """Evaluate a row-level boolean condition, returning a boolean Series.

        Handles: col = "val", col <> "val", col >= val, && conjunctions,
                 VALUE(LEFT(col, n)) >= n patterns.
        """
        condition = condition.strip()
        # Strip DAX line comments
        condition = re.sub(r'--[^\n]*', '', condition).strip()

        # Handle && conjunction
        if "&&" in condition:
            parts = self._split_and_conditions(condition)
            mask = pd.Series(True, index=df.index)
            for part in parts:
                part_mask = self._eval_row_condition(part.strip(), df, table_name)
                if part_mask is not None:
                    mask = mask & part_mask
            return mask

        # Handle VALUE(LEFT(col, n)) >= n pattern
        m_val_left = re.match(
            r"VALUE\s*\(\s*LEFT\s*\(\s*(.+?)\s*,\s*(\d+)\s*\)\s*\)\s*([><=!]+)\s*(\d+)",
            condition, re.IGNORECASE
        )
        if m_val_left:
            col_ref = m_val_left.group(1)
            left_n = int(m_val_left.group(2))
            op = m_val_left.group(3)
            compare_val = int(m_val_left.group(4))

            table, col = _resolve_column_ref(col_ref)
            series = self._resolve_column(table, col, df) if table else None
            if series is None:
                # Try as bare column name
                for c in df.columns:
                    if col_ref.strip().strip("'\"").lower() == c.lower():
                        series = df[c]
                        break

            if series is not None:
                numeric_vals = series.astype(str).str[:left_n]
                try:
                    numeric_vals = pd.to_numeric(numeric_vals, errors="coerce")
                except Exception:
                    return None
                return self._compare_series(numeric_vals, op, compare_val)
            return None

        # Handle comparison: 'T'[C] op "value" / number
        for op in [">=", "<=", "<>", ">", "<", "="]:
            parts = condition.split(op, 1)
            if len(parts) != 2:
                continue

            left_str = parts[0].strip()
            right_str = parts[1].strip()

            table, col = _resolve_column_ref(left_str)
            if table and col:
                series = self._resolve_column(table, col, df)
                if series is None:
                    return None

                # Parse right-hand value
                rval = self._parse_literal(right_str)
                return self._compare_series(series, op, rval)

        return None

    def _split_and_conditions(self, condition: str) -> List[str]:
        """Split a condition on '&&' respecting nested parens and strings."""
        parts = []
        depth = 0
        in_str = False
        str_char = None
        current = []
        i = 0
        while i < len(condition):
            ch = condition[i]
            if in_str:
                current.append(ch)
                if ch == str_char:
                    in_str = False
                i += 1
                continue
            if ch in ('"', "'"):
                in_str = True
                str_char = ch
                current.append(ch)
                i += 1
                continue
            if ch == '(':
                depth += 1
            elif ch == ')':
                depth -= 1

            if depth == 0 and condition[i:i+2] == '&&':
                parts.append(''.join(current).strip())
                current = []
                i += 2
                continue
            current.append(ch)
            i += 1
        if current:
            parts.append(''.join(current).strip())
        return parts

    def _compare_series(
        self,
        series: pd.Series,
        op: str,
        value: Any,
    ) -> pd.Series:
        """Apply a comparison operator to a Series."""
        if isinstance(value, str):
            series_cmp = series.astype(str)
            if op == "=":
                return series_cmp == value
            elif op == "<>":
                return series_cmp != value
            else:
                return series_cmp == value  # String comparisons default to equality
        else:
            numeric_series = pd.to_numeric(series, errors="coerce")
            if op == "=":
                return numeric_series == value
            elif op == "<>":
                return numeric_series != value
            elif op == ">":
                return numeric_series > value
            elif op == ">=":
                return numeric_series >= value
            elif op == "<":
                return numeric_series < value
            elif op == "<=":
                return numeric_series <= value
        return pd.Series(True, index=series.index)

    def _parse_literal(self, text: str) -> Any:
        """Parse a DAX literal value (string, number, DATE)."""
        text = text.strip()
        # String literal: "value"
        if text.startswith('"') and text.endswith('"'):
            return text[1:-1]
        # DATE(y, m, d)
        m_date = re.match(r"DATE\s*\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*\)", text, re.IGNORECASE)
        if m_date:
            return pd.Timestamp(
                year=int(m_date.group(1)),
                month=int(m_date.group(2)),
                day=int(m_date.group(3)),
            )
        # TRUE/FALSE
        if text.upper() == "TRUE":
            return True
        if text.upper() == "FALSE":
            return False
        # Number
        try:
            if '.' in text:
                return float(text)
            return int(text)
        except (ValueError, TypeError):
            return text

    def _resolve_column(
        self,
        table_name: Optional[str],
        column_name: Optional[str],
        df: pd.DataFrame,
    ) -> Optional[pd.Series]:
        """Resolve a 'Table'[Column] reference to a Series in the DataFrame."""
        if not column_name:
            return None

        # Direct column name match in df
        if column_name in df.columns:
            return df[column_name]

        # Case-insensitive match
        lower = column_name.lower()
        for c in df.columns:
            if c.lower() == lower:
                return df[c]

        # Try with table suffix (from join: "Column__Table")
        if table_name:
            suffixed = f"{column_name}__{table_name}"
            if suffixed in df.columns:
                return df[suffixed]
            for c in df.columns:
                if c.lower() == suffixed.lower():
                    return df[c]

        return None

    def _resolve_column_or_measure(
        self,
        table_name: Optional[str],
        column_name: Optional[str],
        df: pd.DataFrame,
    ) -> Optional[Any]:
        """Resolve 'Table'[Name] — try as column first, then as measure.

        Returns a scalar (for measures) or pd.Series (for columns).
        """
        # Try as column first
        series = self._resolve_column(table_name, column_name, df)
        if series is not None:
            return series

        # Try as measure
        if column_name:
            formula = self.lookup_formula(column_name, table_name or "")
            if formula is not None:
                val = self._eval_formula(formula, df, table_name or "")
                return val

        return None

    def _try_simple_agg(
        self,
        formula: str,
        df: pd.DataFrame,
    ) -> Optional[float]:
        """Try to evaluate simple aggregation: SUM(col), COUNT(col), etc."""
        patterns = [
            (r"SUM\s*\((.+?)\)$", "sum"),
            (r"COUNT\s*\((.+?)\)$", "count"),
            (r"COUNTA\s*\((.+?)\)$", "count"),
            (r"AVERAGE\s*\((.+?)\)$", "mean"),
            (r"MIN\s*\((.+?)\)$", "min"),
            (r"MAX\s*\((.+?)\)$", "max"),
        ]
        for pat, agg in patterns:
            m = re.match(pat, formula.strip(), re.IGNORECASE)
            if m:
                col_ref = m.group(1).strip()
                table, col = _resolve_column_ref(col_ref)
                if table and col:
                    series = self._resolve_column(table, col, df)
                    if series is not None:
                        if agg == "count":
                            return series.count()
                        numeric = pd.to_numeric(series, errors="coerce")
                        return getattr(numeric, agg)()
        return None

    def _eval_sumx(
        self,
        formula: str,
        df: pd.DataFrame,
        table_name: str,
        context: Optional[Dict] = None,
    ) -> Optional[float]:
        """Evaluate SUMX(table, expr) — sum of per-row expression."""
        m = re.match(r"SUMX\s*\(", formula, re.IGNORECASE)
        if not m:
            return None
        paren_start = formula.index('(', m.start())
        args_str, _ = _extract_func_args(formula, paren_start)
        args = _split_top_level_args(args_str)
        if len(args) < 2:
            return None

        # First arg is the table expression (usually just a table name)
        table_expr = args[0].strip()
        row_expr = args[1].strip()

        # The table is the current df (already filtered by CALCULATE if applicable)
        # Evaluate row_expr — usually a column reference like T[Value]
        table, col = _resolve_column_ref(row_expr)
        if table and col:
            series = self._resolve_column(table, col, df)
            if series is not None:
                return pd.to_numeric(series, errors="coerce").sum()

        # Could be a measure ref applied per-row
        return None

    def _eval_countax(
        self,
        formula: str,
        df: pd.DataFrame,
        table_name: str,
        context: Optional[Dict] = None,
    ) -> Optional[int]:
        """Evaluate COUNTAX(table_or_filter_expr, count_expr)."""
        m = re.match(r"COUNTAX\s*\(", formula, re.IGNORECASE)
        if not m:
            return None
        paren_start = formula.index('(', m.start())
        args_str, _ = _extract_func_args(formula, paren_start)
        args = _split_top_level_args(args_str)
        if len(args) < 2:
            return None

        table_arg = args[0].strip()
        count_expr = args[1].strip()

        # Table arg could be FILTER(...) or FILTER(KEEPFILTERS(...), ...)
        working_df = df
        if re.match(r"FILTER\s*\(", table_arg, re.IGNORECASE):
            working_df = self._eval_filter_as_table(table_arg, df, table_name)

        # Count non-blank values
        if count_expr.upper() == "TRUE()":
            return len(working_df)

        # Column reference for counting
        table, col = _resolve_column_ref(count_expr)
        if table and col:
            series = self._resolve_column(table, col, working_df)
            if series is not None:
                return int(series.count())

        return len(working_df)

    def _eval_averagex(
        self,
        formula: str,
        df: pd.DataFrame,
        table_name: str,
        context: Optional[Dict] = None,
    ) -> Optional[float]:
        """Evaluate AVERAGEX(table, expr) — average of per-row expression.

        When expr is a measure reference, evaluates the measure for each row
        (each row as its own filter context).
        """
        m = re.match(r"AVERAGEX\s*\(", formula, re.IGNORECASE)
        if not m:
            return None
        paren_start = formula.index('(', m.start())
        args_str, _ = _extract_func_args(formula, paren_start)
        args = _split_top_level_args(args_str)
        if len(args) < 2:
            return None

        row_expr = args[1].strip()

        # Simple column reference
        table, col = _resolve_column_ref(row_expr)
        if table and col:
            series = self._resolve_column(table, col, df)
            if series is not None:
                return pd.to_numeric(series, errors="coerce").mean()

        # Measure reference: evaluate per row
        m_ref = _RE_MEASURE_REF.match(row_expr)
        if m_ref:
            measure_name = m_ref.group(1)
            # For AVERAGEX(Opportunities, [Revenue Won]), Revenue Won =
            # CALCULATE(SUMX(Opportunities, T[Value]), FILTER(T, Status="Won"))
            # Per-row: each row is individually filtered, so SUMX on a single row = Value
            # This means AVERAGEX = average of Value where Status = "Won"
            # Shortcut: evaluate the measure on the full table and return the value
            # (for AVERAGEX with CALCULATE+FILTER, per-row eval gives per-row Value)
            val = self.evaluate(measure_name, df, table_name, context)
            if val is not None:
                # Approximation: if the measure sums over filtered rows,
                # AVERAGEX = sum / count_of_all_rows, not sum / count_of_filtered_rows
                # For Revenue Won Avg Deal Size = AVERAGEX(Opportunities, [Revenue Won])
                # This means: for each opportunity, compute Revenue Won (which filters Status="Won"),
                # so non-Won rows contribute 0 and Won rows contribute their Value.
                # Average = total Revenue Won / total opportunity count
                if len(df) > 0:
                    return val / len(df)
            return val

        return None

    def _eval_filter_as_table(
        self,
        filter_expr: str,
        df: pd.DataFrame,
        table_name: str,
    ) -> pd.DataFrame:
        """Evaluate FILTER(table, condition) and return filtered DataFrame."""
        m = re.match(r"FILTER\s*\(", filter_expr, re.IGNORECASE)
        if not m:
            return df
        paren_start = filter_expr.index('(', m.start())
        args_str, _ = _extract_func_args(filter_expr, paren_start)
        args = _split_top_level_args(args_str)
        if len(args) < 2:
            return df

        # First arg: table or KEEPFILTERS(table)
        table_arg = args[0].strip()
        condition = args[1].strip()

        # Handle KEEPFILTERS wrapping (same behavior for our purposes)
        if re.match(r"KEEPFILTERS\s*\(", table_arg, re.IGNORECASE):
            pass  # KEEPFILTERS just preserves existing filter context — same df

        mask = self._eval_row_condition(condition, df, table_name)
        if mask is not None:
            return df[mask].copy()
        return df

    def _eval_selectedvalue(
        self,
        formula: str,
        df: pd.DataFrame,
        table_name: str,
    ) -> Union[float, int, str, None]:
        """Evaluate SELECTEDVALUE('Table'[Col], default)."""
        m = re.match(r"SELECTEDVALUE\s*\(", formula, re.IGNORECASE)
        if not m:
            return None
        paren_start = formula.index('(', m.start())
        args_str, _ = _extract_func_args(formula, paren_start)
        args = _split_top_level_args(args_str)
        if not args:
            return None

        col_ref = args[0].strip()
        default_val = args[1].strip() if len(args) > 1 else None

        table, col = _resolve_column_ref(col_ref)
        if table and col:
            series = self._resolve_column(table, col, df)
            if series is not None:
                unique = series.dropna().unique()
                if len(unique) == 1:
                    return unique[0]

        # Return default
        if default_val is not None:
            return self._parse_literal(default_val)
        return None

    def _eval_mround(
        self,
        formula: str,
        df: pd.DataFrame,
        table_name: str,
        context: Optional[Dict] = None,
    ) -> Optional[float]:
        """Evaluate MROUND(value, multiple)."""
        m = re.match(r"MROUND\s*\(", formula, re.IGNORECASE)
        if not m:
            return None
        paren_start = formula.index('(', m.start())
        args_str, _ = _extract_func_args(formula, paren_start)
        args = _split_top_level_args(args_str)
        if len(args) < 2:
            return None

        val = self._eval_formula(args[0].strip(), df, table_name, context)
        mult = self._eval_formula(args[1].strip(), df, table_name, context)
        if val is None or mult is None:
            return None
        try:
            val = float(val)
            mult = float(mult)
            if mult == 0:
                return 0.0
            return round(val / mult) * mult
        except (ValueError, TypeError):
            return None

    def _eval_concatenate(
        self,
        formula: str,
        df: pd.DataFrame,
        table_name: str,
        context: Optional[Dict] = None,
    ) -> Optional[str]:
        """Evaluate CONCATENATE(str1, str2)."""
        m = re.match(r"CONCATENATE?\s*\(", formula, re.IGNORECASE)
        if not m:
            return None
        paren_start = formula.index('(', m.start())
        args_str, _ = _extract_func_args(formula, paren_start)
        args = _split_top_level_args(args_str)
        if len(args) < 2:
            return None

        parts = []
        for arg in args:
            val = self._eval_formula(arg.strip(), df, table_name, context)
            parts.append(str(val) if val is not None else "")
        return "".join(parts)

    def _eval_arithmetic(
        self,
        formula: str,
        df: pd.DataFrame,
        table_name: str,
        context: Optional[Dict] = None,
    ) -> Optional[float]:
        """Evaluate arithmetic expressions with measure references like [A] + [B]."""
        expr = formula.strip()

        # Strip outer parens
        if expr.startswith('(') and expr.endswith(')'):
            # Check if these parens are matched
            depth = 0
            matched = True
            for i, ch in enumerate(expr):
                if ch == '(':
                    depth += 1
                elif ch == ')':
                    depth -= 1
                if depth == 0 and i < len(expr) - 1:
                    matched = False
                    break
            if matched:
                expr = expr[1:-1].strip()

        # Replace measure references with evaluated values
        result_expr = expr
        for m in _RE_MEASURE_REF.finditer(expr):
            mname = m.group(1)
            mval = self.evaluate(mname, df, table_name, context)
            if mval is None:
                mval = 0
            result_expr = result_expr.replace(m.group(0), str(float(mval)), 1)

        # Replace column/measure references
        for m in _RE_COLUMN_REF.finditer(result_expr):
            tbl = m.group(1) or m.group(3)
            col_name = m.group(2) or m.group(4)
            resolved = self._resolve_column_or_measure(tbl, col_name, df)
            if resolved is not None:
                if isinstance(resolved, pd.Series):
                    val = resolved.sum()
                else:
                    val = resolved
                try:
                    result_expr = result_expr.replace(m.group(0), str(float(val)), 1)
                except (ValueError, TypeError):
                    result_expr = result_expr.replace(m.group(0), "0", 1)

        return self._safe_eval_arithmetic(result_expr)

    def _safe_eval_arithmetic(self, expr: str) -> Optional[float]:
        """Safely evaluate a simple arithmetic expression (numbers and +-*/ only)."""
        expr = expr.strip()
        # Remove any remaining non-numeric characters except operators and parens
        if not re.match(r'^[\d\.\+\-\*/\(\)\s eE]+$', expr):
            return None
        try:
            # Use Python eval with no builtins for safety
            result = eval(expr, {"__builtins__": {}}, {})
            if isinstance(result, (int, float)):
                return float(result)
        except Exception:
            pass
        return None


# =============================================================================
# Filter application (pre-aggregation column filters from CALCULATETABLE)
# =============================================================================

def parse_filter_expr(dax_str: str) -> Optional[Dict]:
    """Parse a DAX filter expression into structured components.

    Handles:
      'Table'[Col] = "value"
      'Table'[Col] IN {"v1", "v2"}
      NOT 'Table'[Col] IN {"v1"}
      'Table'[Col] >= DATE(2020, 1, 1)
      cond1 && cond2

    Returns dict with keys: table, column, op, values, raw
    """
    dax_str = dax_str.strip()
    if not dax_str or dax_str.startswith("--"):
        return None

    # Handle && conjunction
    if "&&" in dax_str:
        return {"op": "AND", "parts": dax_str.split("&&"), "raw": dax_str}

    # NOT ... IN
    m = re.match(
        r"NOT\s+'([^']+)'\[([^\]]+)\]\s+IN\s*\{(.+?)\}",
        dax_str, re.IGNORECASE
    )
    if m:
        values = _parse_in_values(m.group(3))
        return {"table": m.group(1), "column": m.group(2), "op": "NOT IN", "values": values, "raw": dax_str}

    # IN
    m = re.match(
        r"'([^']+)'\[([^\]]+)\]\s+IN\s*\{(.+?)\}",
        dax_str, re.IGNORECASE
    )
    if m:
        values = _parse_in_values(m.group(3))
        return {"table": m.group(1), "column": m.group(2), "op": "IN", "values": values, "raw": dax_str}

    # Comparison: 'T'[C] op value
    m = re.match(
        r"'([^']+)'\[([^\]]+)\]\s*(>=|<=|<>|>|<|=)\s*(.+)$",
        dax_str, re.IGNORECASE
    )
    if m:
        table, col, op, val_str = m.group(1), m.group(2), m.group(3), m.group(4).strip()
        value = _parse_filter_value(val_str)
        return {"table": table, "column": col, "op": op, "values": [value], "raw": dax_str}

    return {"op": "UNKNOWN", "raw": dax_str}


def _parse_in_values(values_str: str) -> list:
    """Parse values from an IN clause: "val1", "val2" or 123, 456."""
    values = []
    for v in re.findall(r'"([^"]*)"', values_str):
        values.append(v)
    if not values:
        for v in re.findall(r'[\d.]+', values_str):
            try:
                values.append(float(v) if '.' in v else int(v))
            except ValueError:
                values.append(v)
    return values


def _parse_filter_value(val_str: str) -> Any:
    """Parse a single filter value: string, number, or DATE."""
    val_str = val_str.strip()
    if val_str.startswith('"') and val_str.endswith('"'):
        return val_str[1:-1]
    m = re.match(r"DATE\s*\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*\)", val_str, re.IGNORECASE)
    if m:
        return pd.Timestamp(year=int(m.group(1)), month=int(m.group(2)), day=int(m.group(3)))
    if val_str.upper() in ("TRUE", "FALSE"):
        return val_str.upper() == "TRUE"
    try:
        return float(val_str) if '.' in val_str else int(val_str)
    except (ValueError, TypeError):
        return val_str


def apply_filters(
    df: pd.DataFrame,
    filter_exprs: List[str],
    evaluator: Optional[MeasureEvaluator] = None,
) -> pd.DataFrame:
    """Apply a list of DAX filter expressions to a DataFrame (pre-aggregation)."""
    result = df.copy()
    for fexpr in filter_exprs:
        parsed = parse_filter_expr(fexpr)
        if parsed is None:
            continue

        if parsed.get("op") == "UNKNOWN":
            continue

        if parsed.get("op") == "AND":
            for part in parsed["parts"]:
                result = apply_filters(result, [part.strip()], evaluator)
            continue

        table = parsed.get("table", "")
        column = parsed.get("column", "")
        op = parsed.get("op", "")
        values = parsed.get("values", [])

        # Resolve column in DataFrame
        series = None
        if column in result.columns:
            series = result[column]
        else:
            lower = column.lower()
            for c in result.columns:
                if c.lower() == lower:
                    series = result[c]
                    column = c
                    break

        if series is None:
            continue

        if op == "IN":
            # Type coercion for IN comparison
            if values and isinstance(values[0], str):
                mask = series.astype(str).isin(values)
            else:
                mask = series.isin(values)
            result = result[mask].copy()
        elif op == "NOT IN":
            if values and isinstance(values[0], str):
                mask = ~series.astype(str).isin(values)
            else:
                mask = ~series.isin(values)
            result = result[mask].copy()
        elif op in ("=", "<>", ">", ">=", "<", "<="):
            val = values[0] if values else None
            if val is None:
                continue
            if isinstance(val, str):
                series_cmp = series.astype(str)
                if op == "=":
                    result = result[series_cmp == val].copy()
                elif op == "<>":
                    result = result[series_cmp != val].copy()
            elif isinstance(val, pd.Timestamp):
                dt_series = pd.to_datetime(series, errors="coerce")
                if op == "=":
                    result = result[dt_series == val].copy()
                elif op == ">=":
                    result = result[dt_series >= val].copy()
                elif op == "<=":
                    result = result[dt_series <= val].copy()
                elif op == ">":
                    result = result[dt_series > val].copy()
                elif op == "<":
                    result = result[dt_series < val].copy()
                elif op == "<>":
                    result = result[dt_series != val].copy()
            else:
                num_series = pd.to_numeric(series, errors="coerce")
                if op == "=":
                    result = result[num_series == val].copy()
                elif op == "<>":
                    result = result[num_series != val].copy()
                elif op == ">":
                    result = result[num_series > val].copy()
                elif op == ">=":
                    result = result[num_series >= val].copy()
                elif op == "<":
                    result = result[num_series < val].copy()
                elif op == "<=":
                    result = result[num_series <= val].copy()

    return result


# =============================================================================
# Pattern executors
# =============================================================================

# Map PBI aggregation function names to pandas agg functions
_PANDAS_AGG_MAP = {
    "Sum": "sum",
    "Avg": "mean",
    "Count": "count",
    "Min": "min",
    "Max": "max",
    "CountNonNull": "count",
    "Median": "median",
}


def _determine_tables_needed(fields: List[dict]) -> set:
    """Extract the set of table names referenced by a visual's fields."""
    tables = set()
    for f in fields:
        t = f.get("table_sm", "").strip()
        if t:
            tables.add(t)
    return tables


def exec_pattern1(
    visual: dict,
    store: TableStore,
    evaluator: MeasureEvaluator,
    grouping: List[dict],
    measures: List[dict],
    filter_exprs: List[str],
) -> ExecResult:
    """Pattern 1: Measures only (cards, KPIs) → single-row DataFrame.

    Evaluates each measure as a scalar and returns a 1-row DataFrame.
    """
    result = ExecResult(
        page_name=visual.get("_page_name", ""),
        visual_name=visual["visual_name"],
        visual_id=visual.get("visual_id", ""),
        visual_type=visual["visual_type"],
        dax_pattern="Pattern 1",
    )

    # Build joined DataFrame for evaluation context
    tables_needed = _determine_tables_needed(visual["fields"])
    if not tables_needed:
        tables_needed = {"Opportunities"}  # Default fact table

    try:
        joined_df = build_star_join(store, tables_needed)
    except Exception as e:
        result.error = f"Join failed: {e}"
        return result

    # Apply pre-filters
    if filter_exprs:
        joined_df = apply_filters(joined_df, filter_exprs, evaluator)

    # Evaluate each measure
    row = {}
    for m in measures:
        col_name = m["ui_name"] or m["col_sm"]
        formula = m.get("measure_formula", "")
        table_sm = m.get("table_sm", "")
        agg_func = m.get("agg_func", "")

        if formula:
            val = evaluator._eval_formula(formula, joined_df, table_sm)
        elif agg_func and agg_func in _PANDAS_AGG_MAP:
            # Implicit measure
            col_ref = m.get("col_sm", "")
            series = evaluator._resolve_column(table_sm, col_ref, joined_df)
            if series is not None:
                numeric = pd.to_numeric(series, errors="coerce")
                val = getattr(numeric, _PANDAS_AGG_MAP[agg_func])()
            else:
                val = None
                result.warnings.append(f"Column '{col_ref}' not found for implicit measure")
        else:
            # Try as named measure
            val = evaluator.evaluate(m["col_sm"], joined_df, table_sm)

        row[col_name] = val

    result.df = pd.DataFrame([row]) if row else pd.DataFrame()
    return result


def exec_pattern2(
    visual: dict,
    store: TableStore,
    grouping: List[dict],
    filter_exprs: List[str],
) -> ExecResult:
    """Pattern 2: Columns only (slicers, column lists) → distinct values."""
    result = ExecResult(
        page_name=visual.get("_page_name", ""),
        visual_name=visual["visual_name"],
        visual_id=visual.get("visual_id", ""),
        visual_type=visual["visual_type"],
        dax_pattern="Pattern 2",
    )

    tables_needed = _determine_tables_needed(visual["fields"])
    if not tables_needed:
        result.error = "No tables referenced"
        return result

    try:
        joined_df = build_star_join(store, tables_needed)
    except Exception as e:
        result.error = f"Join failed: {e}"
        return result

    if filter_exprs:
        joined_df = apply_filters(joined_df, filter_exprs)

    # Select the grouping columns and deduplicate
    cols = []
    col_names = []
    evaluator = MeasureEvaluator(store)
    for g in grouping:
        col = g.get("col_sm", "")
        table = g.get("table_sm", "")
        series = evaluator._resolve_column(table, col, joined_df)
        if series is not None:
            col_names.append(g["ui_name"] or col)
            cols.append(series.name)

    if cols:
        result.df = joined_df[cols].drop_duplicates().reset_index(drop=True)
        result.df.columns = col_names
    else:
        result.df = pd.DataFrame()
        result.warnings.append("No grouping columns found in joined data")

    return result


def exec_pattern3(
    visual: dict,
    store: TableStore,
    evaluator: MeasureEvaluator,
    grouping: List[dict],
    measures: List[dict],
    filter_exprs: List[str],
) -> ExecResult:
    """Pattern 3: Columns + Measures (most charts/tables) → grouped aggregation.

    Uses groupby for simple measures (SUM, COUNT on columns) and falls back
    to per-group evaluation for complex measures (CALCULATE, VAR, nested refs).
    """
    result = ExecResult(
        page_name=visual.get("_page_name", ""),
        visual_name=visual["visual_name"],
        visual_id=visual.get("visual_id", ""),
        visual_type=visual["visual_type"],
        dax_pattern="Pattern 3",
    )

    tables_needed = _determine_tables_needed(visual["fields"])
    if not tables_needed:
        tables_needed = {"Opportunities"}

    try:
        joined_df = build_star_join(store, tables_needed)
    except Exception as e:
        result.error = f"Join failed: {e}"
        return result

    if filter_exprs:
        joined_df = apply_filters(joined_df, filter_exprs, evaluator)

    # Resolve grouping columns
    group_cols = []  # actual column names in df
    group_labels = []  # display names
    for g in grouping:
        col = g.get("col_sm", "")
        table = g.get("table_sm", "")
        series = evaluator._resolve_column(table, col, joined_df)
        if series is not None:
            group_cols.append(series.name)
            group_labels.append(g["ui_name"] or col)
        else:
            result.warnings.append(f"Grouping column '{col}' not found")

    if not group_cols:
        result.warnings.append("No grouping columns resolved, falling back to Pattern 1")
        return exec_pattern1(visual, store, evaluator, grouping, measures, filter_exprs)

    # Classify measures as simple (vectorized agg) or complex (per-group eval)
    simple_aggs = {}  # col_in_df -> (display_name, pandas_agg)
    complex_measures = []  # (display_name, field_dict)

    for m in measures:
        formula = m.get("measure_formula", "")
        agg_func = m.get("agg_func", "")
        col_sm = m.get("col_sm", "")
        table_sm = m.get("table_sm", "")
        display = m["ui_name"] or col_sm

        if agg_func and agg_func in _PANDAS_AGG_MAP and not formula:
            # Implicit measure: direct column aggregation
            series = evaluator._resolve_column(table_sm, col_sm, joined_df)
            if series is not None:
                simple_aggs[series.name] = (display, _PANDAS_AGG_MAP[agg_func])
            else:
                result.warnings.append(f"Column '{col_sm}' not found for implicit measure")
        elif formula:
            # Check if it's a simple aggregation formula
            simple = _extract_simple_agg(formula)
            if simple:
                agg_type, ref_table, ref_col = simple
                series = evaluator._resolve_column(ref_table, ref_col, joined_df)
                if series is not None:
                    simple_aggs[series.name] = (display, agg_type)
                else:
                    complex_measures.append((display, m))
            else:
                complex_measures.append((display, m))
        else:
            # Named measure without formula in metadata — try to look up
            looked_up = evaluator.lookup_formula(col_sm, table_sm)
            if looked_up:
                simple = _extract_simple_agg(looked_up)
                if simple:
                    agg_type, ref_table, ref_col = simple
                    series = evaluator._resolve_column(ref_table, ref_col, joined_df)
                    if series is not None:
                        simple_aggs[series.name] = (display, agg_type)
                    else:
                        complex_measures.append((display, m))
                else:
                    complex_measures.append((display, m))
            else:
                complex_measures.append((display, m))

    # Execute simple aggregations via groupby
    if simple_aggs and not complex_measures:
        # Pure simple case: vectorized groupby
        agg_dict = {}
        rename_map = {}
        for col_in_df, (display_name, pandas_agg) in simple_aggs.items():
            agg_dict[col_in_df] = pandas_agg
            rename_map[col_in_df] = display_name

        # Ensure numeric types for sum/mean
        for col_in_df, pandas_agg in agg_dict.items():
            if pandas_agg in ("sum", "mean", "median"):
                joined_df[col_in_df] = pd.to_numeric(joined_df[col_in_df], errors="coerce")

        grouped = joined_df.groupby(group_cols, dropna=False).agg(agg_dict).reset_index()
        grouped.rename(columns=rename_map, inplace=True)

        # Rename group columns to display names
        col_rename = dict(zip(group_cols, group_labels))
        grouped.rename(columns=col_rename, inplace=True)
        result.df = grouped
        return result

    # Complex case: per-group evaluation
    rows = []
    grouped_obj = joined_df.groupby(group_cols, dropna=False)

    for group_key, group_df in grouped_obj:
        if not isinstance(group_key, tuple):
            group_key = (group_key,)

        row = dict(zip(group_labels, group_key))

        # Simple aggs on this group
        for col_in_df, (display_name, pandas_agg) in simple_aggs.items():
            series = group_df[col_in_df]
            if pandas_agg in ("sum", "mean", "median"):
                series = pd.to_numeric(series, errors="coerce")
            row[display_name] = getattr(series, pandas_agg)()

        # Complex measures on this group
        for display_name, m in complex_measures:
            formula = m.get("measure_formula", "")
            table_sm = m.get("table_sm", "")
            col_sm = m.get("col_sm", "")

            if formula:
                val = evaluator._eval_formula(formula, group_df, table_sm)
            else:
                val = evaluator.evaluate(col_sm, group_df, table_sm)

            row[display_name] = val

        rows.append(row)

    result.df = pd.DataFrame(rows) if rows else pd.DataFrame()
    return result


def _extract_simple_agg(formula: str) -> Optional[Tuple[str, str, str]]:
    """Check if a formula is a simple aggregation on a column.

    Returns (pandas_agg, table, column) or None.
    """
    formula = formula.strip()
    patterns = [
        (r"^SUM\s*\(\s*'?([^'\(\)]+)'?\[([^\]]+)\]\s*\)$", "sum"),
        (r"^COUNT\s*\(\s*'?([^'\(\)]+)'?\[([^\]]+)\]\s*\)$", "count"),
        (r"^COUNTA\s*\(\s*'?([^'\(\)]+)'?\[([^\]]+)\]\s*\)$", "count"),
        (r"^AVERAGE\s*\(\s*'?([^'\(\)]+)'?\[([^\]]+)\]\s*\)$", "mean"),
        (r"^MIN\s*\(\s*'?([^'\(\)]+)'?\[([^\]]+)\]\s*\)$", "min"),
        (r"^MAX\s*\(\s*'?([^'\(\)]+)'?\[([^\]]+)\]\s*\)$", "max"),
    ]
    for pat, agg in patterns:
        m = re.match(pat, formula, re.IGNORECASE)
        if m:
            return agg, m.group(1).strip(), m.group(2).strip()
    return None


# =============================================================================
# Visual type routing
# =============================================================================

# Visual types that should be skipped (not meaningful as data)
_SKIP_TYPES = {
    "slicer", "advancedSlicerVisual", "shape", "textbox", "image",
    "basicShape", "actionButton", "bookmarkNavigator",
}


def _should_skip_visual(visual_type: str) -> bool:
    """Check if a visual type should be skipped."""
    return visual_type.lower() in _SKIP_TYPES


# =============================================================================
# Main orchestrator
# =============================================================================

def _read_metadata_measures(visuals: dict) -> Dict[Tuple[str, str], str]:
    """Extract measure formulas from metadata for fallback when no model_root."""
    measures = {}
    for key, visual in visuals.items():
        for f in visual["fields"]:
            formula = f.get("measure_formula", "")
            table = f.get("table_sm", "")
            col = f.get("col_sm", "")
            if formula and table and col:
                measures[(table, col)] = formula
    return measures


def execute_visual(
    visual: dict,
    store: TableStore,
    evaluator: MeasureEvaluator,
    filter_exprs: List[str],
) -> ExecResult:
    """Execute a single visual's query and return the result."""
    # Import classify functions from dax_query_builder
    from skills.dax_query_builder import classify_visual_fields

    visual_type = visual["visual_type"]

    if _should_skip_visual(visual_type):
        return ExecResult(
            page_name=visual.get("_page_name", ""),
            visual_name=visual["visual_name"],
            visual_id=visual.get("visual_id", ""),
            visual_type=visual_type,
            dax_pattern="Skipped",
            warnings=[f"Visual type '{visual_type}' skipped"],
        )

    grouping, measures, vis_filters, slicer_fields, matrix_columns = classify_visual_fields(visual["fields"])
    # Merge matrix column-axis fields into grouping for flat execution
    if matrix_columns:
        grouping = list(grouping) + list(matrix_columns)

    # Combine filter expressions with visual-level filter fields
    all_filters = list(filter_exprs)

    # Determine pattern
    has_grouping = len(grouping) > 0
    has_measures = len(measures) > 0

    if has_measures and not has_grouping:
        return exec_pattern1(visual, store, evaluator, grouping, measures, all_filters)
    elif has_grouping and not has_measures:
        return exec_pattern2(visual, store, grouping, all_filters)
    elif has_grouping and has_measures:
        return exec_pattern3(visual, store, evaluator, grouping, measures, all_filters)
    else:
        return ExecResult(
            page_name=visual.get("_page_name", ""),
            visual_name=visual["visual_name"],
            visual_id=visual.get("visual_id", ""),
            visual_type=visual_type,
            dax_pattern="Unknown",
            error="No grouping columns or measures found",
        )


def _sanitize_filename(name: str) -> str:
    """Convert a visual name to a safe filename."""
    # Remove invalid characters
    safe = re.sub(r'[<>:"/\\|?*]', '', name)
    safe = re.sub(r'\s+', '_', safe.strip())
    return safe[:80]  # Truncate long names


def execute_all_visuals(
    metadata_path: str,
    store: TableStore,
    output_dir: str,
    page_filter: Optional[str] = None,
    visual_filter: Optional[str] = None,
) -> List[ExecResult]:
    """Execute DAX queries for all visuals (or filtered subset) and save CSVs.

    Args:
        metadata_path: Path to Skill 1 metadata Excel
        store: TableStore with loaded tables and model
        output_dir: Directory to write per-visual CSVs + manifest
        page_filter: Optional page name to filter (case-insensitive)
        visual_filter: Optional visual name to filter (case-insensitive)

    Returns:
        List of ExecResult for each processed visual
    """
    from skills.dax_query_builder import read_extractor_output, collect_filters_for_visual

    print(f"\nReading metadata: {metadata_path}")
    visuals, page_filters, bookmarks, filter_expr_data = read_extractor_output(metadata_path)
    print(f"  {len(visuals)} visuals across {len(set(k[0] for k in visuals))} pages")

    evaluator = MeasureEvaluator(store)
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    results = []
    manifest = {}
    visual_idx = 0

    for (page_name, visual_key), visual in visuals.items():
        # Page filter
        if page_filter and page_name.lower() != page_filter.lower():
            continue

        # Visual filter
        visual_name = visual["visual_name"]
        if visual_filter and visual_name.lower() != visual_filter.lower():
            continue

        visual["_page_name"] = page_name

        # Skip non-data visuals
        if _should_skip_visual(visual["visual_type"]):
            print(f"  Skipping: {visual_name} ({visual['visual_type']})")
            continue

        print(f"\n  Executing: {visual_name} ({visual['visual_type']})")

        # Collect applicable filters
        filter_exprs = collect_filters_for_visual(
            page_name, visual_name, visual.get("visual_id", ""), filter_expr_data
        )

        # Execute
        result = execute_visual(visual, store, evaluator, filter_exprs)
        results.append(result)

        if result.error:
            print(f"    ERROR: {result.error}")
            continue

        if result.df is not None and not result.df.empty:
            # Clean data: NaN/inf → safe values (PBI shows BLANK as 0 or "(Blank)")
            for col in result.df.columns:
                if pd.api.types.is_numeric_dtype(result.df[col]):
                    result.df[col] = result.df[col].replace([np.inf, -np.inf], 0).fillna(0)
                else:
                    result.df[col] = result.df[col].fillna("(Blank)")

            # Save CSV
            safe_name = _sanitize_filename(visual_name)
            csv_name = f"{visual_idx:02d}_{safe_name}.csv"
            csv_path = output_dir / csv_name
            result.df.to_csv(csv_path, index=False, encoding="utf-8-sig")
            print(f"    Saved: {csv_name} ({len(result.df)} rows, {len(result.df.columns)} cols)")

            # Manifest entry: use visual_id if available, else visual_name
            manifest_key = visual.get("visual_id", "") or visual_name
            manifest[manifest_key] = csv_name

            if result.warnings:
                for w in result.warnings:
                    print(f"    WARNING: {w}")
        else:
            print(f"    No data returned")
            if result.warnings:
                for w in result.warnings:
                    print(f"    WARNING: {w}")

        visual_idx += 1

    # Write manifest
    write_manifest(manifest, output_dir)

    # Summary
    success = sum(1 for r in results if r.df is not None and not r.df.empty)
    errors = sum(1 for r in results if r.error)
    print(f"\nDone: {success} visuals exported, {errors} errors, {len(results) - success - errors} empty")

    return results


def write_manifest(manifest: dict, output_dir: Union[str, Path]):
    """Write csv_manifest.json for page_builder.py compatibility."""
    output_dir = Path(output_dir)
    manifest_path = output_dir / "csv_manifest.json"
    manifest_path.write_text(
        json.dumps(manifest, indent=2, ensure_ascii=False),
        encoding="utf-8",
    )
    print(f"  Manifest: {manifest_path}")


# =============================================================================
# CLI
# =============================================================================

def main():
    parser = argparse.ArgumentParser(
        description="Skill 5: Local DAX Executor — execute DAX queries locally with pandas"
    )
    parser.add_argument(
        "--pbix", metavar="PATH",
        help="Path to .pbix file (extracts tables via pbixray)",
    )
    parser.add_argument(
        "--tables-dir", metavar="PATH",
        help="Path to directory of per-table CSVs",
    )
    parser.add_argument(
        "--metadata", metavar="PATH", required=True,
        help="Path to Skill 1 metadata Excel (pbi_report_metadata.xlsx)",
    )
    parser.add_argument(
        "--model-root", metavar="PATH",
        help="Path to SemanticModel/definition (for relationships + measure formulas)",
    )
    parser.add_argument(
        "--page", metavar="NAME",
        help="Filter to a specific page (case-insensitive)",
    )
    parser.add_argument(
        "--visual", metavar="NAME",
        help="Filter to a specific visual (case-insensitive)",
    )
    parser.add_argument(
        "--output", metavar="PATH", default="output/executor_csv/",
        help="Output directory for per-visual CSVs (default: output/executor_csv/)",
    )

    args = parser.parse_args()

    if not args.pbix and not args.tables_dir:
        parser.error("Must provide either --pbix or --tables-dir")

    if not Path(args.metadata).is_file():
        print(f"ERROR: Metadata file not found: {args.metadata}")
        sys.exit(1)

    # Pre-read metadata to extract measure formulas as fallback
    from skills.dax_query_builder import read_extractor_output
    visuals_tmp, _, _, _ = read_extractor_output(args.metadata)
    metadata_measures = _read_metadata_measures(visuals_tmp)

    # Build table store
    store = build_table_store(
        pbix_path=args.pbix,
        tables_dir=args.tables_dir,
        model_root=args.model_root,
        metadata_measures=metadata_measures,
    )

    # Execute
    results = execute_all_visuals(
        metadata_path=args.metadata,
        store=store,
        output_dir=args.output,
        page_filter=args.page,
        visual_filter=args.visual,
    )

    # Return for module use
    return results


if __name__ == "__main__":
    main()
