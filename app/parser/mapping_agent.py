import os
import json
import time
from typing import Any, Dict, List

from openai import OpenAI

# Load environment variables if python-dotenv is installed
try:
    from dotenv import load_dotenv
    load_dotenv()
except Exception:
    pass

# ---------------------------------------------------------------------
# OpenAI client + model
# ---------------------------------------------------------------------

_client = OpenAI()

_DEFAULT_MODEL = os.getenv("SOV_PARSER_MODEL", "gpt-4.1-mini")

# ---------------------------------------------------------------------
# Canonical fields (must stay in sync with sov_parser_ai.py)
# ---------------------------------------------------------------------

PROPERTY_FIELDS: List[str] = [
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

BUILDING_FIELDS: List[str] = [
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

# ---------------------------------------------------------------------
# Sheet serialization
# ---------------------------------------------------------------------

def _serialize_sheet_rows(
    rows: List[List[Any]], max_rows: int = 150, max_cols: int = 40
) -> List[List[str]]:
    """
    Convert raw sheet rows (from openpyxl) into strings so they can be placed
    in a JSON prompt safely. Truncates to max_rows x max_cols.
    """
    out: List[List[str]] = []
    for r in rows[:max_rows]:
        row_out: List[str] = []
        for cell in r[:max_cols]:
            if cell is None:
                row_out.append("")
            else:
                try:
                    row_out.append(str(cell))
                except Exception:
                    row_out.append("")
        out.append(row_out)
    return out

# ---------------------------------------------------------------------
# Core AI helper (JSON mode + fallback)
# ---------------------------------------------------------------------

def _call_openai_with_optional_json(
    model: str,
    system_msg: str,
    user_payload: Dict[str, Any],
    temperature: float = 0.0,
) -> Dict[str, Any]:
    """
    Try JSON mode first. If the model doesn't cooperate or returns invalid JSON,
    fall back to plain-text + json.loads().
    """
    model = model or _DEFAULT_MODEL

    messages = [
        {"role": "system", "content": system_msg},
        {"role": "user", "content": json.dumps(user_payload)},
    ]

    # --- 1) Try JSON mode ---
    try:
        completion = _client.chat.completions.create(
            model=model,
            response_format={"type": "json_object"},
            messages=messages,
            temperature=temperature,
        )
        content = completion.choices[0].message.content or "{}"
        return json.loads(content)
    except Exception:
        # Fall through to text mode
        pass

    # --- 2) Plain text (still expecting JSON) ---
    completion = _client.chat.completions.create(
        model=model,
        messages=messages,
        temperature=temperature,
    )
    raw = completion.choices[0].message.content or "{}"

    try:
        return json.loads(raw)
    except Exception:
        raise RuntimeError(
            "AI did not return valid JSON.\n"
            f"Raw output:\n{raw}\n"
        )

# ---------------------------------------------------------------------
# System prompts (Extractor + Verifier)
# ---------------------------------------------------------------------

_EXTRACTOR_SYSTEM_MSG = f"""
You are an expert at parsing insurance Statement of Values (SOV) Excel sheets.

Your task:
- Read the provided sheet_data (array of rows; each row is an array of cell strings).
- Identify:
  - Exactly ONE property-level record for this sheet (if any).
  - Zero or more building-level records (for each insurable building row).

You MUST return a single JSON object with this shape:

{{
  "properties": [
    {{
      # one object per sheet (0 or 1)
      # keys MUST come from PROPERTY_FIELDS only
    }}
  ],
  "buildings": [
    {{
      # one object per building row
      # keys MUST come from BUILDING_FIELDS only
    }}
  ]
}}

Canonical PROPERTY_FIELDS (all lowercase, snake_case):
{PROPERTY_FIELDS}

Canonical BUILDING_FIELDS (all lowercase, snake_case):
{BUILDING_FIELDS}

STRICT RULES:
- ONLY use the canonical field names above. Never invent new keys.
- If a value is not explicitly present in the sheet, set it to null.
- Do NOT guess or infer values from outside the sheet (e.g., not from file name, not from sheet title alone).
- Use numbers (not strings) where the sheet clearly contains numeric values (e.g., replacement_cost_tiv, num_units, square footage).
- You may leave fields null if you're not sure.
- For building rows, always fill row_index with the 0-based row index within sheet_data.
- Exclude obvious amenity rows (e.g., 'Mailboxes', 'Signs / Monuments', 'Lighting') from buildings.
- If the sheet clearly has no insurable buildings, return "buildings": [].

IMPORTANT:
- The JSON you output MUST be syntactically valid.
- All keys MUST be one of the canonical fields above.
"""

_VERIFIER_SYSTEM_MSG = f"""
You are a STRICT VERIFIER for a Statement of Values (SOV) extraction.

You receive:
- sheet_data: the original table as an array of rows (each row is an array of cell strings).
- A draft JSON extraction with:
  - "properties": list of property records
  - "buildings": list of building records

Your job:
1. For EVERY field in every record (property + building):
   - Check if the value is clearly supported by sheet_data.
   - If you cannot find clear textual or numeric evidence, set the field to null.
2. Enforce canonical schema:
   - Property records may ONLY contain these keys:
     {PROPERTY_FIELDS}
   - Building records may ONLY contain these keys:
     {BUILDING_FIELDS}
   - Remove any extra or unknown keys.
3. Row index:
   - If "row_index" is present on buildings, ensure it's an integer (0-based).
   - If it's invalid or not a real row, set row_index to null.
4. Numeric sanity:
   - Obvious numeric fields (e.g., replacement_cost_tiv, num_units, square footage)
     must be numeric if present, otherwise null.
   - If the draft used text like "N/A" or "-", convert to null.
5. Buildings vs amenities:
   - If a "building" clearly refers to an amenity only (e.g., "MAILBOXES",
     "MONUMENT SIGN", "CARPORT LIGHTING"), drop that record entirely.
6. Do NOT add new records:
   - You may delete obviously wrong buildings.
   - You may delete obviously wrong property records.
   - But do NOT invent new properties or buildings.

Output:
- A single JSON object with:
{{
  "properties": [ ... cleaned canonical records ... ],
  "buildings":  [ ... cleaned canonical records ... ]
}}

You MUST return valid JSON. Do not include explanations, comments, or extra keys.
"""

# ---------------------------------------------------------------------
# Public function: Call dual-agent AI for a single sheet
# ---------------------------------------------------------------------

def call_sov_sheet_agent(
    source_file: str,
    sheet_name: str,
    rows: List[List[Any]],
    model: str = None,
) -> Dict[str, Any]:
    """
    Dual-agent SOV sheet parser.

    1) Agent A (Extractor) reads sheet_data and proposes:
       {
         "properties": [ ... ],
         "buildings": [ ... ]
       }

    2) Agent B (Verifier) receives:
       - original sheet_data
       - the draft JSON from Agent A
       and returns a cleaned, schema-corrected version.

    Returns a dict with keys "properties" (list) and "buildings" (list).
    """

    model = model or _DEFAULT_MODEL
    sheet_data = _serialize_sheet_rows(rows)

    # --------------------
    # 1) Extraction phase
    # --------------------
    extractor_payload = {
        "source_file": source_file,
        "sheet_name": sheet_name,
        "sheet_data": sheet_data,
        "property_fields": PROPERTY_FIELDS,
        "building_fields": BUILDING_FIELDS,
    }

    for attempt in range(3):
        try:
            draft = _call_openai_with_optional_json(
                model=model,
                system_msg=_EXTRACTOR_SYSTEM_MSG,
                user_payload=extractor_payload,
                temperature=0.0,
            )
            break
        except Exception:
            if attempt == 2:
                raise
            time.sleep(1.5)

    if not isinstance(draft, dict):
        draft = {}
    draft_props = draft.get("properties") or []
    draft_bldgs = draft.get("buildings") or []

    # Normalize shapes a bit before verification
    if not isinstance(draft_props, list):
        draft_props = []
    if not isinstance(draft_bldgs, list):
        draft_bldgs = []

    draft = {
        "properties": draft_props,
        "buildings": draft_bldgs,
    }

    # --------------------
    # 2) Verification phase
    # --------------------
    verifier_payload = {
        "source_file": source_file,
        "sheet_name": sheet_name,
        "sheet_data": sheet_data,
        "draft": draft,
        "property_fields": PROPERTY_FIELDS,
        "building_fields": BUILDING_FIELDS,
    }

    for attempt in range(3):
        try:
            cleaned = _call_openai_with_optional_json(
                model=model,
                system_msg=_VERIFIER_SYSTEM_MSG,
                user_payload=verifier_payload,
                temperature=0.0,
            )
            break
        except Exception:
            if attempt == 2:
                raise
            time.sleep(1.5)

    if not isinstance(cleaned, dict):
        cleaned = {}

    props = cleaned.get("properties") or []
    bldgs = cleaned.get("buildings") or []

    if not isinstance(props, list):
        props = []
    if not isinstance(bldgs, list):
        bldgs = []

    return {
        "properties": props,
        "buildings": bldgs,
    }