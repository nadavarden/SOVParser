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
  "buildings": [
    { <ONE building record per building row> },
    ...
  ]
}

====================================================================
GLOBAL BEHAVIOR & MULTI-SHEET LOGIC
====================================================================

The workbook may contain multiple sheets. Some sheets contain building-level rows, while others contain property-level metadata.

**You MUST support:**
- Sheets with GENERAL PROPERTY INFORMATION (e.g., “General Information”, “Property Info”, “Summary”).
- Sheets with MANAGEMENT INFO that should NOT overwrite property info.
- Sheets with detailed BUILDING rows.
- Sheets with totals, notes, disclaimers—ignore those but read any useful metadata inside them.

If a sheet contains property-level information, extract:

- location_full_address  
- location_address  
- location_city  
- location_state  
- location_zip  

Store this information mentally as **property-level context**.

Then, for every building row in any sheet:

- If the building row provides its own values → use them.
- If the building row is missing a value → fill from property-level context.
- If both are missing → set null.

IMPORTANT RULE — IDENTIFYING AND HANDLING GENERAL / SUMMARY SHEETS
------------------------------------------------------------------

Some sheets contain GENERAL PROPERTY INFORMATION and **must not** be treated as building sheets.

A sheet is a GENERAL / SUMMARY sheet when it:

- Contains headers such as:
  “General Information”, “Property Information”, “Property Summary”, “Location (State, City, Zip)”
- Contains only one location
- Contains totals, summaries, management info, or global values
- Does NOT contain repeated building rows
- Does NOT contain building numbers
- Does NOT contain unit ranges, square footage values, or multiple addresses

When the sheet is a GENERAL sheet:

### YOU MUST:
- Extract property-level values:
  - location_address
  - location_city
  - location_state
  - location_zip
  - (optionally location_full_address)

### YOU MUST NOT:
- Produce any building objects from this sheet
- Include this sheet's information as a separate building

### For any building parsed on other sheets:
- If a building is missing location_city/location_state/location_zip/location_address → 
  fill in the values obtained from the GENERAL sheet.

====================================================================
IDENTIFYING PROPERTY-LEVEL SHEETS
====================================================================

A sheet should be treated as a PROPERTY-LEVEL info sheet if:

- It contains labels like:
    “General Information”, “Property Information”, “Location (State, City, Zip)”, “Property Address”
- It contains a single address or location block
- It contains no repeating building-like rows
- It does NOT contain unit numbers, ranges, or square footage patterns

A sheet should NOT be treated as property-level if it is:

- “Management Information”
- “Policy Information”
- “Insurance Carrier / Broker info”
- Pure notes or disclaimers

====================================================================
IDENTIFYING BUILDING ROWS
====================================================================

A row should be interpreted as a building record if:

- It contains a building number
OR
- It contains a valid address/cluster of addresses
OR
- It contains SqFt, unit counts, or replacement cost values

Use heuristics:

- Columns like “Bldg”, “Bldg #”, “Building Number”, “Street #”, “Address #”, “Unit #s”, “# of Units”
- Rows with numbers + addresses + cost/SqFt almost always represent a building.

A row MUST NOT be interpreted as a building if:

- It contains only one overall address for the property
- It has no building number, unit count, SqFt, or cost values
- It belongs to a GENERAL INFORMATION sheet

====================================================================
CANONICAL OUTPUT FIELDS (ALL REQUIRED IN EACH BUILDING OBJECT)
====================================================================

- source_file
- sheet_name
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

====================================================================
RULES FOR DATA EXTRACTION
====================================================================

1. **Normalization**
   - Strings may contain extra spaces, commas, punctuation, or ranges like “13361, 59, 55”.
   - Clean and standardize values when possible.
   - Convert numeric strings (“1,241,989”, “7015”) into numbers.

2. **Missing Data Handling**
   - If a building field is missing → copy from property-level info.
   - If BOTH are missing → set null.

3. **Ranges in Address or Units**
   - Handle multi-number addresses: “13344, 42, 40, 38”.
   - Handle unit ranges: “1 thru 20”, “29-36”.

4. **Irrelevant Sheets**
   - Ignore totals, summaries, disclaimers.
   - Ignore "Management Information" EXCEPT do NOT use its location for buildings.
5.- Note that some files may have several several sheets in them where one sheet is for total summary and the other is per building details
  - Somewhere in the sheet there are other rows/columns with notes, totals, etc. which shouldn't be ignored. such as:
  [LOCATION (STATE, CITY, ZIP):		13322 W Stonebrook Drive, Sun City West, AZ 85375	]

6. **Always output strictly valid JSON.**
====================================================================
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

    # Retry logic for robustness
    for attempt in range(3):
        try:
            return _call_openai_with_optional_json(model, system_msg, payload)
        except Exception as e:
            if attempt == 2:
                raise
            time.sleep(1.2)

    # should never hit here
    return {"buildings": []}
