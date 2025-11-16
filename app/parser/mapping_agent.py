import os
import json
from typing import Any, Dict, List

from openai import OpenAI

from dotenv import load_dotenv
load_dotenv()

# Expects OPENAI_API_KEY in env
_client = OpenAI()

# You can change this to any suitable model
_DEFAULT_MODEL = os.getenv("SOV_PARSER_MODEL", "gpt-4.1-mini")


def _serialize_sheet_rows(rows: List[List[Any]], max_rows: int = 80, max_cols: int = 20) -> List[List[str]]:
    """
    Convert raw sheet rows (from openpyxl) into strings, truncating for prompt safety.
    """
    serialized: List[List[str]] = []
    for r_idx, row in enumerate(rows[:max_rows]):
        out_row: List[str] = []
        for c_idx, cell in enumerate(row[:max_cols]):
            if cell is None:
                out_row.append("")
            else:
                out_row.append(str(cell))
        serialized.append(out_row)
    return serialized


def call_sov_sheet_agent(
    source_file: str,
    sheet_name: str,
    rows: List[List[Any]],
    model: str = None,
) -> Dict[str, Any]:
    """
    Full AI mode:
    - Send the entire (truncated) sheet to the LLM
    - Ask it to produce a structured SOV JSON with:
        { "properties": [...], "buildings": [...] }

    Each property record uses canonical PROPERTY_FIELDS.
    Each building record uses canonical BUILDING_FIELDS.
    """
    model = model or _DEFAULT_MODEL
    sheet_data = _serialize_sheet_rows(rows)

    system_msg = """
You are an expert underwriter assistant that parses Statement of Values (SOV) spreadsheets
for insurance. You receive a 2D array representing a single worksheet from Excel (rows × columns).

Your job:
1. Identify the property-level / summary information (one record per sheet).
2. Identify building-level rows (one record per building).
3. Map them into a canonical JSON schema with specific field names.

You MUST return valid JSON ONLY, with this top-level structure:

{
  "properties": [
    {
      "source_file": "...",
      "sheet_name": "...",
      "number_of_buildings": float or null,
      "roof_type": string or null,
      "building_valuation_type": string or null,
      "building_replacement_cost": float or null,
      "blanket_outdoor_property": float or null,
      "business_personal_property": float or null,
      "total_insurable_value": float or null,
      "general_liability": float or null,
      "building_ordinance_a": float or null,
      "building_ordinance_b": float or null,
      "building_ordinance_c": float or null,
      "equipment_breakdown": float or null,
      "sewer_or_drain_backup": float or null,
      "business_income": float or null,
      "hired_and_non_owned_auto": float or null,
      "playgrounds_number": float or null,
      "streets_miles": float or null,
      "pools_number": float or null,
      "spas_number": float or null,
      "wader_pools_number": float or null,
      "restroom_building_sq_ft": float or null,
      "guardhouse_sq_ft": float or null,
      "clubhouse_sq_ft": float or null,
      "fitness_center_sq_ft": float or null,
      "tennis_courts_number": float or null,
      "basketball_courts_number": float or null,
      "other_sport_courts_number": float or null,
      "walking_biking_trails_miles": float or null,
      "lakes_or_ponds_number": float or null,
      "boat_docks_and_slips_number": float or null,
      "dog_parks_number": float or null,
      "elevators_number": float or null,
      "commercial_exposure_sq_ft": float or null
    }
  ],
  "buildings": [
    {
      "source_file": "...",
      "sheet_name": "...",
      "row_index": integer Excel row index or null,
      "building_number": string or null,
      "location_full_address": string or null,
      "location_address": string or null,
      "location_city": string or null,
      "location_state": string or null,
      "location_zip": string or null,
      "lat": float or null,
      "long": float or null,
      "betterview_id": string or null,
      "betterview_building_number": string or null,
      "units_per_building": float or null,
      "replacement_cost_tiv": float or null,
      "num_units": float or null,
      "livable_sq_ft": float or null,
      "garage_sq_ft": float or null,
      "commercial_sq_ft": float or null,
      "building_class": string or null,
      "parking_type": string or null,
      "roof_type": string or null,
      "smoke_detectors": string or null,
      "sprinklered": string or null,
      "year_of_construction": float or null,
      "number_of_stories": float or null,
      "construction_type": string or null
    }
  ]
}

Rules:
- Use null for any field you cannot confidently determine.
- Infer numeric fields even if the sheet has them formatted as text.
- Do NOT invent buildings or properties that aren't clearly present.
- Exclude rows that are amenities only (e.g., Lighting, Mailboxes, Signs/Monuments, Landscaping-only rows).
- Row index should be 1-based Excel row index if you can infer it; otherwise null.
- There should be at most ONE property record per worksheet (the summary for that sheet).
- Field names MUST exactly match the schema above.
"""

    user_msg = {
        "role": "user",
        "content": json.dumps(
            {
                "source_file": source_file,
                "sheet_name": sheet_name,
                "sheet_data": sheet_data,
            }
        ),
    }

    completion = _client.chat.completions.create(
        model=model,
        messages=[
            {"role": "system", "content": system_msg},
            user_msg,
        ],
        response_format={"type": "json_object"},
        temperature=0.1,
    )

    content = completion.choices[0].message.content
    try:
        result = json.loads(content)
    except Exception as e:
        raise RuntimeError(f"Failed to parse AI JSON response: {e}\nRaw: {content}")

    # Basic sanity
    if "properties" not in result:
        result["properties"] = []
    if "buildings" not in result:
        result["buildings"] = []

    return result