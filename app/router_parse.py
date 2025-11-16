import os
import uuid
from fastapi import APIRouter, UploadFile, File, HTTPException
from app.parser.sov_parser import parse_workbook
from app.parser.mapping_agent import call_mapping_agent
from app.db import insert_property_records, insert_building_records, db_session
from app.utils import save_temp_file, collect_top_left_cells

router = APIRouter()

@router.post("/parse")
async def parse_sov_file(file: UploadFile = File(...)):
    if not file.filename.endswith(".xlsx"):
        raise HTTPException(status_code=400, detail="Only .xlsx files are accepted")

    # Save upload
    file_id = str(uuid.uuid4())
    temp_path = save_temp_file(file, file_id)

    # Parse file
    prop_records, building_records = parse_workbook(temp_path)

    # Persist to DB
    with db_session() as conn:
        insert_property_records(conn, prop_records)
        insert_building_records(conn, building_records)

    return {
        "file": file.filename,
        "file_id": file_id,
        "properties_extracted": len(prop_records),
        "buildings_extracted": len(building_records)
    }