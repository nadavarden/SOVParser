import os
import json
import time
from typing import Any, Dict, List

from openai import OpenAI

try:
    from dotenv import load_dotenv

    load_dotenv()
except:
    pass

_client = OpenAI()

_DEFAULT_MODEL = os.getenv("SOV_PARSER_MODEL", "gpt-4.1-mini")


def _serialize_sheet_rows(rows, max_rows=80, max_cols=25):
    out = []
    for r in rows[:max_rows]:
        row = []
        for c in r[:max_cols]:
            row.append("" if c is None else str(c))
        out.append(row)
    return out


def _call_openai_with_optional_json(model, system_msg, user_payload):
    messages = [
        {"role": "system", "content": system_msg},
        {"role": "user", "content": json.dumps(user_payload)},
    ]

    try:
        completion = _client.chat.completions.create(
            model=model,
            messages=messages,
            temperature=0.1,
            response_format={"type": "json_object"},
        )
        return json.loads(completion.choices[0].message.content)

    except Exception as e:
        if "response_format" not in str(e).lower():
            raise

    completion = _client.chat.completions.create(
        model=model, messages=messages, temperature=0.1
    )
    raw = completion.choices[0].message.content

    try:
        return json.loads(raw)
    except:
        raise RuntimeError(f"AI did not return valid JSON:\n{raw}")


def call_sov_sheet_agent(source_file, sheet_name, rows, model=None):
    model = model or _DEFAULT_MODEL
    sheet_data = _serialize_sheet_rows(rows)

    system_msg = """
You are an expert in parsing insurance Statement of Values (SOV) spreadsheets.

You receive JSON:
{
  "source_file": "...",
  "sheet_name": "...",
  "sheet_data": [ [...], [...], ... ]
}

You must return **strict valid JSON** in the form:

{
  "sheet_type": "general" | "building",
  "buildings": [
      { ... building or metadata record ... }
  ]
}

====================================================================
MANDATORY OUTPUT BEHAVIOR
====================================================================

Every sheet MUST return:

► "sheet_type": "general"
    → when the sheet contains general/property-level information
    → you MUST output EXACTLY ONE metadata object in "buildings"
    → you MUST NOT output any real building rows

► "sheet_type": "building"
    → when the sheet contains building-level rows
    → you MUST output ONE record per real building row
    → never output null/empty rows

This field is REQUIRED so downstream logic can categorize sheets correctly.
Failure to obey this rule breaks the parser.

====================================================================
WHEN TO MARK A SHEET AS GENERAL ("sheet_type": "general")
====================================================================

A sheet IS GENERAL if ANY of the following are true:

- Contains labels like:
    "General Information", "Property Information", "Property Info",
    "Property Summary", "Location (State, City, Zip)",
    "Property Address", "Insured Location"

- Contains a single address/location block

- Contains totals, summaries, or metadata but NOT repeated building rows

- Does NOT contain building numbers, unit ranges, SqFt columns, replacement cost values, or multiple addresses

- Contains notes such as:
    LOCATION (STATE, CITY, ZIP): 13322 W Stonebrook Drive, Sun City West, AZ 85375

A GENERAL sheet MUST NOT produce any real buildings.

====================================================================
GENERAL SHEET OUTPUT FORMAT
====================================================================

When "sheet_type": "general":

Return EXACTLY ONE object inside "buildings":

{
  "building_number": null,
  "location_full_address": <best full address>,
  "location_address": <address without city/state>,
  "location_city": <city>,
  "location_state": <state>,
  "location_zip": <zip>,
  "units_per_building": null,
  "replacement_cost_tiv": null,
  "num_units": null,
  "livable_sq_ft": null,
  "garage_sq_ft": null
}

Never output more than one record.
Never output any building_number other than null.

====================================================================
WHEN TO MARK A SHEET AS BUILDING ("sheet_type": "building")
====================================================================

A sheet IS BUILDING if:

- It contains building numbers ("1", "2", ... OR "Bldg", "Bldg #")
- It contains multiple addresses
- It contains replacement cost, SqFt, units, or unit ranges
- Rows resemble typical SOV building rows

BUILDING RULES:
- Output one object per real building row
- Missing fields → null (Python will merge metadata later)

====================================================================
ADDRESS & CITY/STATE MERGING RULES
====================================================================

If one model produces a more complete address (more tokens),
such as including the state ("CA") or full ZIP,
you MUST output the more complete address.

Example:
- "720 Blue Oak Ave, Thousand Oaks 91320"
- "720 Blue Oak Ave, Thousand Oaks, CA 91320"

→ YOU MUST output:
  "720 Blue Oak Ave, Thousand Oaks, CA 91320"

====================================================================
UNIT RANGE RULE
====================================================================

If units_per_building can be expressed as a range, ALWAYS output the range.

Examples:
- "720 to 726"  ← ALWAYS preferred
- "720–726"
- "4 units" → convert into range ONLY IF the sheet clearly provides endpoint numbers.

NEVER replace a valid range with a numeric unit count.

====================================================================
CANONICAL FIELDS FOR EVERY RECORD
====================================================================

The following must ALWAYS appear in every record:

- building_number
- location_full_address
- location_address
- location_city
- location_state
- location_zip
- units_per_building
- replacement_cost_tiv
- num_units
- livable_sq_ft
- garage_sq_ft

Missing values → use null.

====================================================================
RULES FOR BUILDING DATA
====================================================================

1. Normalize all address fields (remove extra spaces/commas).
2. Normalize numeric values ("1,241,989" → 1241989).
3. Handle unit ranges ("1 thru 20").
4. Handle multi-number addresses ("13361, 59, 55").
5. Leave missing fields as null; Python will fill from general metadata.

FEW-SHOT EXAMPLE 1
====================================================================

### INPUT SHEET DATA (SIMPLIFIED)
[
  ["Bldg #", "Address #", "Street Name", "unit #s", "Building Replacement Cost", "# of units", "Livable Square Footage", "Garage Square Footage],
  ["1", "13361, 59,55,51,47", "W Aleppo", "", "$1,241,989", "5", "7,015", "1,500"],
  ["2", "13343, 39,35,31,27", "W Aleppo", "", "$1,241,989", "5", "7,015", "1,500"],
  ["3", "13323, 19, 15, 11, 07, 03", "W Aleppo", "", "$1,515,315", "6", "8,697", "1,800"],
  ["4", "13344, 42, 40, 38, 36", "Zephyr Court", "", $1,106,632", "5", "6,168", "1,500"]

]
somewhere in the sheet there are other rows/columns with notes, totals, etc. which shouldn't be ignored. such as:
[LOCATION (STATE, CITY, ZIP):		13322 W Stonebrook Drive, Sun City West, AZ 85375	]
### EXPECTED JSON OUTPUT
{
  "buildings": [
    {
      "building_number": "1",
      "location_full_address": "13361, 59,55,51,47 W Aleppo, Sun City West, AZ 85375",
      "location_address": "13361, 59,55,51,47 W Aleppo",
      "location_city": "Sun City West",
      "location_state": "AZ",
      "location_zip": "85375",
      "units_per_building": 5.0,
      "replacement_cost_tiv": 1241989.0,
      "num_units": 5.0,
      "livable_sq_ft": 7015.0,
      "garage_sq_ft": 1500.0
    }
    },
    {
      "building_number": "2",
      "location_full_address": ""13343, 39,35,31,27 W Aleppo, Sun City West, AZ 85375",
      "location_address": "13343, 39,35,31,27 W Aleppo",
      "location_city": "Sun City West",
      "location_state": "AZ",
      "location_zip": "85375",
      "units_per_building": 5.0,
      "replacement_cost_tiv": 1241989.0,
      "num_units": 5.0,
      "livable_sq_ft": 7015.0,
      "garage_sq_ft": 1500.0
    },
    {
        "building_number": 3,
        "location_full_address": "13323, 19, 15, 11, 07, 03 W Aleppo, Sun City West, AZ 85375",
        "location_address": "13323, 19, 15, 11, 07, 03 W Aleppo",
        "location_city": "Sun City West",
        "location_state": "AZ",
        "location_zip": "85375",
        "units_per_building": 6.0,
        "replacement_cost_tiv": 1515315.0,
        "num_units": 6.0,
        "livable_sq_ft": 8697.0,
        "garage_sq_ft": 1800.0
    },
    {
        "building_number": 4,
        "location_full_address": "13346, 48, 50,52,54 Zephyr Court, Sun City West, AZ 85375",
        "location_address": "13346, 48, 50,52,54 Zephyr Court",
        "location_city": "Sun City West",
        "location_state": "AZ",
        "location_zip": "85375",
        "units_per_building": 5.0,
        "replacement_cost_tiv": 11066320,
        "num_units": 5.0,
        "livable_sq_ft": 6168.0,
        "garage_sq_ft": 1500.0
    }
  ]
}

====================================================================
FEW-SHOT EXAMPLE 2
====================================================================

### INPUT SHEET DATA (SIMPLIFIED)
[
  ["Bldg", "Street #", "Street Name", "City", "State", "Zip","100% Replacement Cost", "# Of Units", "Unit #", "Building Square Footage", "Attached Garage Square Footage],
  ["1", "10555", "Montgomery Road", "Montgomery", "OH", "45242", "$3,769,090", "20", "1 thru 20", "19800"],
  ["2", "10555", "Montgomery Road", "Montgomery", "OH", "45242", "$1,118,054", "8", "21 thru 28", "5232"],
  ["3", "10555", "Montgomery Road", "Montgomery", "OH", "45242", "$3,112,306", "8", "29 thru 36", "12238"]
]

### EXPECTED JSON OUTPUT
{
  "buildings": [
    {
      "building_number": "1",
      "location_full_address": null,
      "location_address": null,
      "location_city": "Montgomery",
      "location_state": "OH",
      "location_zip": "45242",
      "units_per_building": "1 thru 20",
      "replacement_cost_tiv": 3769090.0,
      "num_units": 20.0,
      "livable_sq_ft": 19800.0,
      "garage_sq_ft": 0.0
    },
    {
      "building_number": "2",
      "location_full_address": null,
      "location_address": null,
      "location_city": "Montgomery",
      "location_state": "OH",
      "location_zip": "45242",
      "units_per_building": "21 thru 28",
      "replacement_cost_tiv": 1118054.0,
      "num_units": 8.0,
      "livable_sq_ft": 5232.0,
      "garage_sq_ft": 0.0
    },
    {
      "building_number": "3",
      "location_full_address": null,
      "location_address": null,
      "location_city": "Montgomery",
      "location_state": "OH",
      "location_zip": "45242",
      "units_per_building": "29 thru 36",
      "replacement_cost_tiv": 3112306.0,
      "num_units": 8.0,
      "livable_sq_ft": 12238.0,
      "garage_sq_ft": 0.0
    }
  ]
}
"""

    payload = {
        "source_file": source_file,
        "sheet_name": sheet_name,
        "sheet_data": sheet_data,
    }

    for _ in range(3):
        try:
            return _call_openai_with_optional_json(model, system_msg, payload)
        except:
            time.sleep(1.2)

    return {"sheet_type": "building", "buildings": []}
