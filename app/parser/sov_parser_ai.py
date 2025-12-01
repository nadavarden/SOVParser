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

from app.parser.mapping_agent_single import call_sov_sheet_agent

# =====================================================================
# 1. Canonical Fields
# =====================================================================


BUILDING_FIELDS = [
    "source_file",
    "sheet_name",
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
    "units_per_building",
    "replacement_cost_tiv",
    "num_units",
    "livable_sq_ft",
    "garage_sq_ft",
}

# =====================================================================
# 2. Dataclasses
# =====================================================================


@dataclass
class BuildingRecord:
    source_file: str
    sheet_name: str
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

    def to_row(self):
        return [getattr(self, f) for f in BUILDING_FIELDS]


# =====================================================================
# 3. Helpers
# =====================================================================


def _rows_from_sheet(ws) -> List[List[Any]]:
    return [list(r) for r in ws.iter_rows(values_only=True)]


def _coerce_float(v: Any) -> Optional[float]:
    """
    Convert values like "1,234", "  123.45 ", or 123 to float.
    Treat empty strings, 'n/a', '-', etc. as None.
    """
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return float(v)
    try:
        s = str(v).strip().lower()
        if s in {"", "n/a", "na", "none", "null", "-", "--"}:
            return None
        s = s.replace(",", "")
        return float(s)
    except Exception:
        return None


def _building_from_ai_dict(d: Dict[str, Any]) -> BuildingRecord:
    data: Dict[str, Any] = {f: None for f in BUILDING_FIELDS}
    data.update({k: d.get(k) for k in d.keys() if k in BUILDING_FIELDS})

    for f in NUMERIC_BUILDING_FIELDS:
        if f in data:
            data[f] = _coerce_float(data[f])

    return BuildingRecord(**{k: data.get(k) for k in BUILDING_FIELDS})


# =====================================================================
# 4. Workbook Parsing (Full-AI)
# =====================================================================


def parse_workbook(path: str) -> Tuple[List[BuildingRecord]]:
    wb = load_workbook(path, data_only=True)

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
        sheet_bldgs = ai_result.get("buildings", [])

        for b in sheet_bldgs:
            b.setdefault("source_file", source_file)
            b.setdefault("sheet_name", sheet_name)
            building_records.append(_building_from_ai_dict(b))

    return building_records


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
        building_records = parse_workbook(path)
        return {
            "buildings": [b.__dict__ for b in building_records],
        }
