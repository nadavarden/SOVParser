"""

Full-AI Statement of Values (SOV) parsing engine.

- Loads Excel workbooks
- Sends sheet data to an AI agent
- Receives structured JSON for:
    - Property-level summary
    - Building-level records
- Wraps results in SOVParser.parse_excel()
"""

import os
from dataclasses import dataclass
from typing import Optional, List, Dict, Any, Tuple

from openpyxl import load_workbook

from app.parser.mapping_agent import call_sov_sheet_agent

# =====================================================================
# 1. Canonical Fields
# =====================================================================

PROPERTY_FIELDS = [
    "source_file",
    "sheet_name",
    "number_of_buildings",
    "roof_type",
    "building_valuation_type",
    "building_replacement_cost",
    "blanket_outdoor_property",
    "business_personal_property",
    "total_insurable_value",
    "general_liability",
    "building_ordinance_a",
    "building_ordinance_b",
    "building_ordinance_c",
    "equipment_breakdown",
    "sewer_or_drain_backup",
    "business_income",
    "hired_and_non_owned_auto",
    "playgrounds_number",
    "streets_miles",
    "pools_number",
    "spas_number",
    "wader_pools_number",
    "restroom_building_sq_ft",
    "guardhouse_sq_ft",
    "clubhouse_sq_ft",
    "fitness_center_sq_ft",
    "tennis_courts_number",
    "basketball_courts_number",
    "other_sport_courts_number",
    "walking_biking_trails_miles",
    "lakes_or_ponds_number",
    "boat_docks_and_slips_number",
    "dog_parks_number",
    "elevators_number",
    "commercial_exposure_sq_ft",
]

BUILDING_FIELDS = [
    "source_file",
    "sheet_name",
    "row_index",
    "building_number",
    "location_full_address",
    "location_address",
    "location_city",
    "location_state",
    "location_zip",
    "lat",
    "long",
    "betterview_id",
    "betterview_building_number",
    "units_per_building",
    "replacement_cost_tiv",
    "num_units",
    "livable_sq_ft",
    "garage_sq_ft",
    "commercial_sq_ft",
    "building_class",
    "parking_type",
    "roof_type",
    "smoke_detectors",
    "sprinklered",
    "year_of_construction",
    "number_of_stories",
    "construction_type",
]

NUMERIC_PROPERTY_FIELDS = {
    "number_of_buildings",
    "building_replacement_cost",
    "blanket_outdoor_property",
    "business_personal_property",
    "total_insurable_value",
    "general_liability",
    "building_ordinance_a",
    "building_ordinance_b",
    "building_ordinance_c",
    "equipment_breakdown",
    "sewer_or_drain_backup",
    "business_income",
    "hired_and_non_owned_auto",
    "playgrounds_number",
    "streets_miles",
    "pools_number",
    "spas_number",
    "wader_pools_number",
    "restroom_building_sq_ft",
    "guardhouse_sq_ft",
    "clubhouse_sq_ft",
    "fitness_center_sq_ft",
    "tennis_courts_number",
    "basketball_courts_number",
    "other_sport_courts_number",
    "walking_biking_trails_miles",
    "lakes_or_ponds_number",
    "boat_docks_and_slips_number",
    "dog_parks_number",
    "elevators_number",
    "commercial_exposure_sq_ft",
}

NUMERIC_BUILDING_FIELDS = {
    "lat",
    "long",
    "units_per_building",
    "replacement_cost_tiv",
    "num_units",
    "livable_sq_ft",
    "garage_sq_ft",
    "commercial_sq_ft",
    "year_of_construction",
    "number_of_stories",
}

# =====================================================================
# 2. Dataclasses
# =====================================================================

@dataclass
class PropertyRecord:
    source_file: str
    sheet_name: str
    number_of_buildings: Optional[float] = None
    roof_type: Optional[str] = None
    building_valuation_type: Optional[str] = None
    building_replacement_cost: Optional[float] = None
    blanket_outdoor_property: Optional[float] = None
    business_personal_property: Optional[float] = None
    total_insurable_value: Optional[float] = None
    general_liability: Optional[float] = None
    building_ordinance_a: Optional[float] = None
    building_ordinance_b: Optional[float] = None
    building_ordinance_c: Optional[float] = None
    equipment_breakdown: Optional[float] = None
    sewer_or_drain_backup: Optional[float] = None
    business_income: Optional[float] = None
    hired_and_non_owned_auto: Optional[float] = None
    playgrounds_number: Optional[float] = None
    streets_miles: Optional[float] = None
    pools_number: Optional[float] = None
    spas_number: Optional[float] = None
    wader_pools_number: Optional[float] = None
    restroom_building_sq_ft: Optional[float] = None
    guardhouse_sq_ft: Optional[float] = None
    clubhouse_sq_ft: Optional[float] = None
    fitness_center_sq_ft: Optional[float] = None
    tennis_courts_number: Optional[float] = None
    basketball_courts_number: Optional[float] = None
    other_sport_courts_number: Optional[float] = None
    walking_biking_trails_miles: Optional[float] = None
    lakes_or_ponds_number: Optional[float] = None
    boat_docks_and_slips_number: Optional[float] = None
    dog_parks_number: Optional[float] = None
    elevators_number: Optional[float] = None
    commercial_exposure_sq_ft: Optional[float] = None

    def to_row(self):
        return [getattr(self, f) for f in PROPERTY_FIELDS]


@dataclass
class BuildingRecord:
    source_file: str
    sheet_name: str
    row_index: Optional[int] = None
    building_number: Optional[str] = None
    location_full_address: Optional[str] = None
    location_address: Optional[str] = None
    location_city: Optional[str] = None
    location_state: Optional[str] = None
    location_zip: Optional[str] = None
    lat: Optional[float] = None
    long: Optional[float] = None
    betterview_id: Optional[str] = None
    betterview_building_number: Optional[str] = None
    units_per_building: Optional[float] = None
    replacement_cost_tiv: Optional[float] = None
    num_units: Optional[float] = None
    livable_sq_ft: Optional[float] = None
    garage_sq_ft: Optional[float] = None
    commercial_sq_ft: Optional[float] = None
    building_class: Optional[str] = None
    parking_type: Optional[str] = None
    roof_type: Optional[str] = None
    smoke_detectors: Optional[str] = None
    sprinklered: Optional[str] = None
    year_of_construction: Optional[float] = None
    number_of_stories: Optional[float] = None
    construction_type: Optional[str] = None

    def to_row(self):
        return [getattr(self, f) for f in BUILDING_FIELDS]

# =====================================================================
# 3. Helpers
# =====================================================================

def _rows_from_sheet(ws) -> List[List[Any]]:
    return [list(r) for r in ws.iter_rows(values_only=True)]


def _coerce_float(v: Any) -> Optional[float]:
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return float(v)
    try:
        s = str(v).replace(",", "").strip()
        if not s:
            return None
        return float(s)
    except Exception:
        return None


def _property_from_ai_dict(d: Dict[str, Any]) -> PropertyRecord:
    # Ensure all canonical fields exist
    data: Dict[str, Any] = {f: None for f in PROPERTY_FIELDS}
    data.update({k: d.get(k) for k in d.keys() if k in PROPERTY_FIELDS})

    # Numeric coercion
    for f in NUMERIC_PROPERTY_FIELDS:
        if f in data:
            data[f] = _coerce_float(data[f])

    return PropertyRecord(
        **{
            k: data.get(k)
            for k in PROPERTY_FIELDS
        }
    )


def _building_from_ai_dict(d: Dict[str, Any]) -> BuildingRecord:
    data: Dict[str, Any] = {f: None for f in BUILDING_FIELDS}
    data.update({k: d.get(k) for k in d.keys() if k in BUILDING_FIELDS})

    for f in NUMERIC_BUILDING_FIELDS:
        if f in data:
            data[f] = _coerce_float(data[f])

    # row_index should be int or None
    if data.get("row_index") is not None:
        try:
            data["row_index"] = int(float(data["row_index"]))
        except Exception:
            data["row_index"] = None

    return BuildingRecord(
        **{
            k: data.get(k)
            for k in BUILDING_FIELDS
        }
    )

# =====================================================================
# 4. Workbook Parsing (Full-AI)
# =====================================================================

def parse_workbook(path: str) -> Tuple[List[PropertyRecord], List[BuildingRecord]]:
    wb = load_workbook(path, data_only=True)

    property_records: List[PropertyRecord] = []
    building_records: List[BuildingRecord] = []

    source_file = os.path.basename(path)

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        rows = _rows_from_sheet(ws)

        # Call AI for this sheet
        ai_result = call_sov_sheet_agent(
            source_file=source_file,
            sheet_name=sheet_name,
            rows=rows,
        )

        sheet_props = ai_result.get("properties", [])
        sheet_bldgs = ai_result.get("buildings", [])

        for p in sheet_props:
            # Force source_file/sheet_name if missing or wrong
            p.setdefault("source_file", source_file)
            p.setdefault("sheet_name", sheet_name)
            property_records.append(_property_from_ai_dict(p))

        for b in sheet_bldgs:
            b.setdefault("source_file", source_file)
            b.setdefault("sheet_name", sheet_name)
            building_records.append(_building_from_ai_dict(b))

    return property_records, building_records

# =====================================================================
# 5. Public Interface
# =====================================================================

class SOVParser:
    """
    Full-AI SOV parser interface.
    Use:
        parser = SOVParser()
        result = parser.parse_excel("path/to/file.xlsx")
    """

    def parse_excel(self, path: str) -> Dict[str, Any]:
        property_records, building_records = parse_workbook(path)
        return {
            "properties": [p.__dict__ for p in property_records],
            "buildings": [b.__dict__ for b in building_records],
        }