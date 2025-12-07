"""
Full-AI Statement of Values (SOV) parsing engine — DUAL LLM + sheet_type aware.

- Loads Excel workbooks
- Sends sheet data to 2 LLMs (gpt-5.1 + gpt-4o-mini)
- Merges outputs with fuzzy-comparison logic
- Identifies general vs building sheets using the LLM ("sheet_type")
- Extracts metadata from general sheets
- Applies metadata to all building rows
"""

import os
import json
import re
from dataclasses import dataclass
from typing import Optional, List, Dict, Any

from openpyxl import load_workbook
from app.parser.mapping_agent_single import call_sov_sheet_agent

# =====================================================================
# Canonical Fields
# =====================================================================

BUILDING_FIELDS = [
    "building_number",
    "location_full_address",
    "location_address",
    "location_city",
    "location_state",
    "location_zip",
    "units_per_building",
    "replacement_cost_tiv",
    "num_units",
    "livable_sq_ft",
    "garage_sq_ft",
]

NUMERIC_BUILDING_FIELDS = {
    "replacement_cost_tiv",
    "num_units",
    "livable_sq_ft",
    "garage_sq_ft",
}

# =====================================================================
# Dataclass
# =====================================================================


@dataclass
class BuildingRecord:
    building_number: Optional[str] = None
    location_full_address: Optional[str] = None
    location_address: Optional[str] = None
    location_city: Optional[str] = None
    location_state: Optional[str] = None
    location_zip: Optional[str] = None
    units_per_building: Optional[float] = None
    replacement_cost_tiv: Optional[float] = None
    num_units: Optional[float] = None
    livable_sq_ft: Optional[float] = None
    garage_sq_ft: Optional[float] = None


# =====================================================================
# Helpers
# =====================================================================


def _rows(ws):
    return [list(r) for r in ws.iter_rows(values_only=True)]


def _coerce_float(v):
    if v is None:
        return None
    try:
        return float(str(v).replace(",", "").strip())
    except:
        return None


def normalize_address(s: str):
    if s is None:
        return None
    s = str(s)
    s = " ".join(s.split())  # collapse spaces
    s = re.sub(r"[–—−]", "-", s)  # unify dashes
    s = s.lower()
    s = s.rstrip(",.; ")
    return s


def _normalize_string(v):
    if v is None:
        return None
    s = str(v).lower().strip()
    s = " ".join(s.split())
    s = re.sub(r"[^\w\s\-]", "", s)
    return s


def _to_float_if_possible(v):
    try:
        return float(str(v).replace(",", "").strip())
    except:
        return None


def fuzzy_numeric_equal(a, b, tol=0.5):
    fa = _to_float_if_possible(a)
    fb = _to_float_if_possible(b)
    if fa is None or fb is None:
        return False
    if abs(fa) > 1000:
        fa = round(fa, 1)
        fb = round(fb, 1)
    return abs(fa - fb) <= tol


def _fuzzy_equal(a, b):
    # Case: both None
    if a is None and b is None:
        return True

    # One missing
    if a is None or b is None:
        return False

    # Numeric fuzzy
    if fuzzy_numeric_equal(a, b):
        return True

    # Address normalization
    na = normalize_address(a)
    nb = normalize_address(b)
    if na == nb:
        return True

    # Clean "null" pollution
    cleaned_na = " ".join(na.replace("null", "").replace(",", "").split())
    cleaned_nb = " ".join(nb.replace("null", "").replace(",", "").split())
    if cleaned_na == cleaned_nb:
        return True
    if cleaned_na in cleaned_nb or cleaned_nb in cleaned_na:
        return True

    # Generic string normalization
    sa = _normalize_string(a)
    sb = _normalize_string(b)
    if sa == sb:
        return True

    return False


def choose_richer_value(field: str, v1, v2):
    if v1 is None and v2 is None:
        return None

    # --- Address fields: pick more complete form ---
    if field in ["location_full_address", "location_address"]:
        return v1 if len(str(v1)) >= len(str(v2)) else v2

    # --- Units per building MUST be a range ---
    if field == "units_per_building":

        def is_range(v):
            if v is None:
                return False
            s = str(v)
            return (" to " in s) or ("-" in s) or ("," in s) or ("thru" in s.lower())

        # Prefer ranges over numeric counts
        if is_range(v1) and not is_range(v2):
            return v1
        if is_range(v2) and not is_range(v1):
            return v2

        # If both are ranges, use the more complete one
        return v1 if len(str(v1)) >= len(str(v2)) else v2

    # --- Default fallback ---
    return v1 if len(str(v1)) >= len(str(v2)) else v2


# =====================================================================
# DUAL LLM MERGE LOGIC — WITH sheet_type preserved
# =====================================================================


def compare_model_outputs(a: Dict[str, Any], b: Dict[str, Any]) -> Dict[str, Any]:
    """
    Dual-LLM merge with:
    - fuzzy matching
    - completeness scoring
    - enriched-address preference
    - range-over-count rule for units
    - robust missing-key handling
    """

    # -----------------------
    # Merge sheet_type first
    # -----------------------
    type_a = a.get("sheet_type")
    type_b = b.get("sheet_type")

    # If either says general → treat as general
    merged_type = (
        "general" if (type_a == "general" or type_b == "general") else "building"
    )

    rows_a = a.get("buildings", []) or []
    rows_b = b.get("buildings", []) or []

    max_len = max(len(rows_a), len(rows_b))
    merged_rows = []

    for idx in range(max_len):
        rec_a = rows_a[idx] if idx < len(rows_a) else {}
        rec_b = rows_b[idx] if idx < len(rows_b) else {}

        merged = {}
        mismatches = {}

        for key in BUILDING_FIELDS:
            v1 = rec_a.get(key)
            v2 = rec_b.get(key)

            # CASE 1 — both None
            if v1 is None and v2 is None:
                merged[key] = None
                continue

            # CASE 2 — fuzzy match results → keep v1
            if _fuzzy_equal(v1, v2):
                merged[key] = v1
                continue

            # CASE 3 — one missing → take non-missing value
            if v1 is None and v2 is not None:
                merged[key] = v2
                continue
            if v2 is None and v1 is not None:
                merged[key] = v1
                continue

            # CASE 4 — both exist but different → choose richer value
            chosen = choose_richer_value(key, v1, v2)
            if chosen is None:
                # fallback preserve v1
                chosen = v1

            merged[key] = chosen

            mismatches[key] = {"gpt_5_1": v1, "gpt_4o_mini": v2, "chosen": chosen}

        if mismatches:
            print(f"\n⚠️  LLM MISMATCH DETECTED at index {idx}:")
            print(json.dumps(mismatches, indent=4))

        merged_rows.append(merged)

    return {"sheet_type": merged_type, "buildings": merged_rows}


# =====================================================================
# Metadata extraction
# =====================================================================


def extract_property_metadata(record: Dict[str, Any]) -> Dict[str, Any]:
    return {
        k: record.get(k)
        for k in [
            "location_full_address",
            "location_address",
            "location_city",
            "location_state",
            "location_zip",
        ]
        if record.get(k)
    }


# =====================================================================
# Workbook Parsing — FULL Dual-LLM + sheet_type logic
# =====================================================================


def parse_workbook(path: str) -> List[BuildingRecord]:
    wb = load_workbook(path, data_only=True)

    all_buildings: List[BuildingRecord] = []
    property_meta: Dict[str, Any] = {}
    source_file = os.path.basename(path)

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        if getattr(ws, "sheet_state", "visible") != "visible":
            print(f"⚠️  Skipping hidden sheet: {sheet_name}")
            continue
        rows = _rows(ws)

        # --- Run dual models ---
        ai_a = call_sov_sheet_agent(source_file, sheet_name, rows, model="gpt-5.1")
        ai_b = call_sov_sheet_agent(source_file, sheet_name, rows, model="gpt-4o-mini")

        # --- Merge results (keeps fuzzy logic + sheet_type) ---
        merged = compare_model_outputs(ai_a, ai_b)
        sheet_type = merged.get("sheet_type", "building")
        returned_buildings = merged.get("buildings", [])

        # ---------------------------------------------------------
        # GENERAL SHEET
        # ---------------------------------------------------------
        if sheet_type == "general":
            if returned_buildings:
                meta = extract_property_metadata(returned_buildings[0])
                property_meta.update(meta)
            continue

        # ---------------------------------------------------------
        # BUILDING SHEET
        # ---------------------------------------------------------
        for rec in returned_buildings:
            safe_record = {}
            for field in BUILDING_FIELDS:
                raw = rec.get(field)
                if field in NUMERIC_BUILDING_FIELDS:
                    safe_record[field] = _coerce_float(raw)
                else:
                    safe_record[field] = raw if raw not in ("", "null") else None

            b = BuildingRecord(**safe_record)
            all_buildings.append(b)

    # ---------------------------------------------------------
    # Apply metadata inheritance
    # ---------------------------------------------------------
    for b in all_buildings:
        for field, value in property_meta.items():
            if getattr(b, field) in (None, "", "null"):
                setattr(b, field, value)

    return all_buildings


# =====================================================================
# Public Interface
# =====================================================================


class SOVParser:
    def parse_excel(self, path: str) -> Dict[str, Any]:
        buildings = parse_workbook(path)
        return {"buildings": [b.__dict__ for b in buildings]}
