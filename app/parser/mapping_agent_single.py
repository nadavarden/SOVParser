import os
import json
import time
from typing import Any, Dict, List

from openai import OpenAI

# Load environment variables if python-dotenv is installed
try:
    from dotenv import load_dotenv

    load_dotenv()
except:
    pass

# Initialize client
_client = OpenAI()

# Choose a default model, overridable in .env
_DEFAULT_MODEL = os.getenv("SOV_PARSER_MODEL", "gpt-4.1-mini")

# ---------------------------------------------------------------------
# Utility: serialize Excel sheet into a prompt-friendly form
# ---------------------------------------------------------------------


def _serialize_sheet_rows(
    rows: List[List[Any]], max_rows: int = 80, max_cols: int = 25
) -> List[List[str]]:
    """
    Convert raw sheet rows (from openpyxl) into strings so they can be placed in a JSON prompt safely.
    """
    out = []
    for r in rows[:max_rows]:
        row_out = []
        for cell in r[:max_cols]:
            if cell is None:
                row_out.append("")
            else:
                try:
                    row_out.append(str(cell))
                except:
                    row_out.append("")
        out.append(row_out)
    return out


# ---------------------------------------------------------------------
# Core AI Function (with JSON fallback logic)
# ---------------------------------------------------------------------


def _call_openai_with_optional_json(
    model: str, system_msg: str, user_payload: Dict[str, Any]
) -> Dict[str, Any]:
    """
    Try structured JSON mode first.
    If the model rejects `response_format={'type': 'json_object'}`,
    fall back to plain text + json.loads().
    """

    messages = [
        {"role": "system", "content": system_msg},
        {"role": "user", "content": json.dumps(user_payload)},
    ]

    # --- 1. Attempt JSON-Mode ---
    try:
        completion = _client.chat.completions.create(
            model=model,
            messages=messages,
            temperature=0.1,
            response_format={"type": "json_object"},
        )

        content = completion.choices[0].message.content
        return json.loads(content)

    except Exception as e:
        # If the model does NOT support structured JSON:
        if "response_format" in str(e).lower():
            pass  # fall through to fallback mode
        else:
            raise  # real error

    # --- 2. Fallback: no structured JSON mode ---
    completion = _client.chat.completions.create(
        model=model,
        messages=messages,
        temperature=0.1,
    )

    raw = completion.choices[0].message.content

    # Attempt to parse JSON if present
    try:
        return json.loads(raw)
    except:
        raise RuntimeError("AI did not return valid JSON.\n" f"Raw output:\n{raw}\n")


# ---------------------------------------------------------------------
# Public Function: Call AI for a single sheet
# ---------------------------------------------------------------------


def call_sov_sheet_agent(
    source_file: str,
    sheet_name: str,
    rows: List[List[Any]],
    model: str = None,
) -> Dict[str, Any]:
    """
    Sends one Excel sheet to the AI engine and requests:

        {
            "properties": [...],
            "buildings": [...]
        }

    This is full-AI mode (Mode C) — the AI constructs the structure itself.
    """

    model = model or _DEFAULT_MODEL
    sheet_data = _serialize_sheet_rows(rows)

    system_msg = """
You are an expert in parsing insurance Statement of Values (SOV) spreadsheets.

You receive a JSON structure:
{
  "source_file": "...",
  "sheet_name": "...",
  "sheet_data": [[row1col1, row1col2, ...], [row2col1, ...], ...]
}

You must return **valid JSON only** with the structure:

{
  "properties": [
    { <ONE property record for this sheet> }
  ],
  "buildings": [
    { <ONE building record per building row> },
    ...
  ]
}

All property and building fields must match these canonical names exactly:

PROPERTY FIELDS:
- source_file
- sheet_name
- number_of_buildings
- roof_type
- building_valuation_type
- building_replacement_cost
- blanket_outdoor_property
- business_personal_property
- total_insurable_value
- general_liability
- building_ordinance_a
- building_ordinance_b
- building_ordinance_c
- equipment_breakdown
- sewer_or_drain_backup
- business_income
- hired_and_non_owned_auto
- playgrounds_number
- streets_miles
- pools_number
- spas_number
- wader_pools_number
- restroom_building_sq_ft
- guardhouse_sq_ft
- clubhouse_sq_ft
- fitness_center_sq_ft
- tennis_courts_number
- basketball_courts_number
- other_sport_courts_number
- walking_biking_trails_miles
- lakes_or_ponds_number
- boat_docks_and_slips_number
- dog_parks_number
- elevators_number
- commercial_exposure_sq_ft

BUILDING FIELDS:
- source_file
- sheet_name
- row_index          (1-based Excel index if possible)
- building_number
- location_full_address
- location_address
- location_city
- location_state
- location_zip
- lat
- long
- betterview_id
- betterview_building_number
- units_per_building
- replacement_cost_tiv
- num_units
- livable_sq_ft
- garage_sq_ft
- commercial_sq_ft
- building_class
- parking_type
- roof_type
- smoke_detectors
- sprinklered
- year_of_construction
- number_of_stories
- construction_type

RULES:
- ONE property record per worksheet.
- Use null where unknown.
- Exclude amenity rows (e.g., “Lighting”, “Mailboxes”, “Signs/Monuments”).
- Detect buildings using heuristics: numerical building number + SqFt/Units patterns.
- All numbers may be formatted as strings in the sheet; convert them to numbers if possible.
- If building field is missing → copy from property-level relevant field, if both are missing set to null
- For property fields you might need to consider summation of building fields if relevant
- Count only valid buildings (rows where address is present), for Number of buildings field
"""

    payload = {
        "source_file": source_file,
        "sheet_name": sheet_name,
        "sheet_data": sheet_data,
    }

    # Retry logic for robustness
    for attempt in range(3):
        try:
            return _call_openai_with_optional_json(model, system_msg, payload)
        except Exception as e:
            if attempt == 2:
                raise
            time.sleep(1.2)

    # should never hit here
    return {"properties": [], "buildings": []}
