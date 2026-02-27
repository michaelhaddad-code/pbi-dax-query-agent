# -*- coding: utf-8 -*-
"""
Shared TMDL Parser Module
PBI AutoGov — Power BI Data Governance Automation Pipeline

Extracts measures AND columns from TMDL semantic model files into a structured
SemanticModel dataclass with case-insensitive lookup indexes.

Used by:
  - Skill 1 (extract_metadata.py) via parse_tmdl_files() legacy wrapper
  - Skill 2 (dax_query_builder.py) via parse_semantic_model() for formula lookup
"""

import json
import re
from dataclasses import dataclass, field
from pathlib import Path


# ============================================================
# Data classes
# ============================================================

@dataclass
class TmdlColumn:
    """A column parsed from a TMDL file."""
    table: str
    name: str
    data_type: str = ""
    is_hidden: bool = False


@dataclass
class TmdlRelationship:
    """A relationship between two tables in the semantic model."""
    from_table: str
    from_column: str
    to_table: str
    to_column: str
    is_active: bool = True
    cardinality: str = ""       # e.g. "oneToMany", "manyToOne"
    cross_filtering: str = ""   # e.g. "oneDirection", "bothDirections"


@dataclass
class SemanticModel:
    """Full semantic model parsed from TMDL files."""
    # (table, measure_name) -> DAX formula
    measures: dict = field(default_factory=dict)
    # (table, column_name) -> TmdlColumn
    columns: dict = field(default_factory=dict)
    # Flat indexes for case-insensitive lookup: lowercase name -> list of (table, name)
    _measure_index: dict = field(default_factory=dict)
    _column_index: dict = field(default_factory=dict)
    # Model provenance: "pbixray", "pbip", or "" (unknown)
    source: str = ""
    # Relationships between tables
    relationships: list = field(default_factory=list)

    @property
    def types_reliable(self) -> bool:
        """Whether column data types can be trusted.

        pbixray marks ALL columns as string regardless of actual type,
        so types are unreliable when source is "pbixray" or unknown.
        """
        return self.source not in ("pbixray", "")

    def build_indexes(self):
        """Build case-insensitive lookup indexes after parsing."""
        self._measure_index = {}
        for (table, mname) in self.measures:
            key = mname.lower()
            if key not in self._measure_index:
                self._measure_index[key] = []
            self._measure_index[key].append((table, mname))

        self._column_index = {}
        for (table, cname) in self.columns:
            key = cname.lower()
            if key not in self._column_index:
                self._column_index[key] = []
            self._column_index[key].append((table, cname))

    @property
    def measure_names(self) -> dict:
        """lowercase measure name -> list of (table, measure_name)"""
        return self._measure_index

    @property
    def column_names(self) -> dict:
        """lowercase column name -> list of (table, column_name)"""
        return self._column_index


# ============================================================
# TMDL file parser — measures (unchanged from extract_metadata.py)
# ============================================================

def _parse_measures(content: str, table_name: str) -> dict:
    """Extract measure definitions from TMDL content.
    Returns dict of (table_name, measure_name) -> dax_formula.
    """
    measures = {}

    # Regex to capture measure name and DAX formula body
    # Group 1: measure name, Group 2: DAX formula body
    measure_pattern = re.compile(
        r"^\tmeasure\s+'?([^'=\n]+?)'?\s*=\s*(.*?)(?=^\t(?:measure|column|hierarchy|partition|annotation)\s|\Z)",
        re.MULTILINE | re.DOTALL,
    )

    for m in measure_pattern.finditer(content):
        measure_name = m.group(1).strip().strip("'")
        raw_formula = m.group(2).strip()

        # Clean up: stop at TMDL metadata keywords
        formula_lines = []
        for line in raw_formula.split("\n"):
            stripped = line.strip()
            if re.match(r"^(formatString|lineageTag|annotation|extendedProperty|displayFolder|dataCategory)\s*[: ]", stripped):
                break
            formula_lines.append(line)

        formula = "\n".join(formula_lines).strip()
        # Remove fenced code block wrappers
        formula = re.sub(r"^```\s*\n?", "", formula)
        formula = re.sub(r"\n?\s*```\s*$", "", formula)
        # Clean up indentation
        formula = re.sub(r"\t{2,}", "    ", formula)
        formula = re.sub(r"\t", "    ", formula)
        # Collapse multiple blank lines
        formula = re.sub(r"\n\s*\n", "\n", formula)
        formula = formula.strip()

        measures[(table_name, measure_name)] = formula if formula else ""

    return measures


# ============================================================
# TMDL file parser — columns (NEW)
# ============================================================

def _parse_columns(content: str, table_name: str) -> dict:
    """Extract column definitions from TMDL content.
    Returns dict of (table_name, column_name) -> TmdlColumn.
    """
    columns = {}

    # Match column definitions: tab-indented "column" keyword followed by name
    # Columns can be: `column Name` or `column 'Name With Spaces'`
    # Some have `= <expression>` for calculated columns
    column_pattern = re.compile(
        r"^\tcolumn\s+'?([^'=\n]+?)'?\s*(?:=.*?)?$"
        r"(.*?)"
        r"(?=^\t(?:measure|column|hierarchy|partition|annotation|///)\s|\Z)",
        re.MULTILINE | re.DOTALL,
    )

    for m in column_pattern.finditer(content):
        col_name = m.group(1).strip().strip("'")
        body = m.group(2)

        # Extract data type
        dt_match = re.search(r"dataType:\s*(\S+)", body)
        data_type = dt_match.group(1) if dt_match else ""

        # Check if hidden
        is_hidden = bool(re.search(r"^\t\tisHidden", body, re.MULTILINE))

        columns[(table_name, col_name)] = TmdlColumn(
            table=table_name,
            name=col_name,
            data_type=data_type,
            is_hidden=is_hidden,
        )

    return columns


# ============================================================
# Single TMDL file parser
# ============================================================

def _parse_single_tmdl(filepath: Path) -> tuple[dict, dict]:
    """Parse a single TMDL file. Returns (measures_dict, columns_dict)."""
    content = filepath.read_text(encoding="utf-8-sig")

    # Extract table name from first line
    table_match = re.match(r"^table\s+(.+?)$", content, re.MULTILINE)
    if not table_match:
        return {}, {}
    table_name = table_match.group(1).strip().strip("'")

    measures = _parse_measures(content, table_name)
    columns = _parse_columns(content, table_name)

    return measures, columns


# ============================================================
# Public API
# ============================================================

def parse_semantic_model(model_root) -> SemanticModel:
    """Parse all TMDL files in a semantic model directory.

    Args:
        model_root: Path to semantic model definition root (contains tables/).

    Returns:
        SemanticModel with measures, columns, lookup indexes, source, and relationships.
    """
    model_root = Path(model_root)
    tables_dir = model_root / "tables"
    model = SemanticModel()

    # Read .source marker file if present (written by pbix_extractor)
    source_file = model_root / ".source"
    if source_file.is_file():
        model.source = source_file.read_text(encoding="utf-8").strip()
    else:
        # No marker → assume real PBIP export
        model.source = "pbip"

    # Read relationships.json if present (written by pbix_extractor)
    rel_file = model_root / "relationships.json"
    if rel_file.is_file():
        try:
            rel_data = json.loads(rel_file.read_text(encoding="utf-8"))
            for r in rel_data:
                model.relationships.append(TmdlRelationship(
                    from_table=r.get("fromTable", ""),
                    from_column=r.get("fromColumn", ""),
                    to_table=r.get("toTable", ""),
                    to_column=r.get("toColumn", ""),
                    is_active=r.get("isActive", True),
                    cardinality=r.get("cardinality", ""),
                    cross_filtering=r.get("crossFiltering", ""),
                ))
        except (json.JSONDecodeError, KeyError) as e:
            print(f"WARNING: Could not parse relationships.json: {e}")

    if not tables_dir.is_dir():
        print(f"WARNING: Tables directory not found: {tables_dir}")
        return model

    for tmdl_file in sorted(tables_dir.glob("**/*.tmdl")):
        measures, columns = _parse_single_tmdl(tmdl_file)
        model.measures.update(measures)
        model.columns.update(columns)

    model.build_indexes()
    return model


def parse_tmdl_files(tables_dir) -> dict:
    """Legacy wrapper for Skill 1 backward compatibility.
    Parse all TMDL files to extract measures and their DAX formulas.
    Returns dict of (table_name, measure_name) -> dax_formula.
    """
    tables_dir = Path(tables_dir)
    measures = {}
    if not tables_dir.is_dir():
        print(f"WARNING: Tables directory not found: {tables_dir}")
        return measures
    for tmdl_file in sorted(tables_dir.glob("**/*.tmdl")):
        file_measures, _ = _parse_single_tmdl(tmdl_file)
        measures.update(file_measures)
    return measures


# ============================================================
# Field-to-model matching
# ============================================================

def match_field_to_model(field_name: str, model: SemanticModel) -> dict | None:
    """Match a bare field name against the semantic model.

    Matching priority:
        1. Exact measure name (case-insensitive)
        2. Exact column name (case-insensitive)
        3. Fuzzy match (normalized: strip spaces/underscores, lowercase)
        4. None

    Returns:
        dict with keys: table, field_name, formula (str or ""), match_type
        or None if no match found.
    """
    key = field_name.lower()

    # 1. Exact measure match
    if key in model.measure_names:
        matches = model.measure_names[key]
        table, mname = matches[0]  # Take first match
        formula = model.measures.get((table, mname), "")
        return {
            "table": table,
            "field_name": mname,
            "formula": formula,
            "match_type": "measure",
        }

    # 2. Exact column match
    if key in model.column_names:
        matches = model.column_names[key]
        table, cname = matches[0]
        return {
            "table": table,
            "field_name": cname,
            "formula": "",
            "match_type": "column",
        }

    # 3. Fuzzy match — normalize by removing spaces, underscores, hyphens
    def _normalize(s):
        return re.sub(r"[\s_\-]+", "", s.lower())

    norm_key = _normalize(field_name)

    # Check measures first
    for mname_lower, matches in model.measure_names.items():
        if _normalize(mname_lower) == norm_key:
            table, mname = matches[0]
            formula = model.measures.get((table, mname), "")
            return {
                "table": table,
                "field_name": mname,
                "formula": formula,
                "match_type": "measure_fuzzy",
            }

    # Check columns
    for cname_lower, matches in model.column_names.items():
        if _normalize(cname_lower) == norm_key:
            table, cname = matches[0]
            return {
                "table": table,
                "field_name": cname,
                "formula": "",
                "match_type": "column_fuzzy",
            }

    return None
