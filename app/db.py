import sqlite3
from contextlib import contextmanager

DB_PATH = "sov.db"

def init_db():
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()

    # Property table
    c.execute("""
    CREATE TABLE IF NOT EXISTS sov_properties (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        source_file TEXT,
        sheet_name TEXT,
        number_of_buildings REAL,
        roof_type TEXT,
        building_valuation_type TEXT,
        building_replacement_cost REAL,
        blanket_outdoor_property REAL,
        business_personal_property REAL,
        total_insurable_value REAL,
        general_liability REAL,
        building_ordinance_a REAL,
        building_ordinance_b REAL,
        building_ordinance_c REAL,
        equipment_breakdown REAL,
        sewer_or_drain_backup REAL,
        business_income REAL,
        hired_and_non_owned_auto REAL,
        playgrounds_number REAL,
        streets_miles REAL,
        pools_number REAL,
        spas_number REAL,
        wader_pools_number REAL,
        restroom_building_sq_ft REAL,
        guardhouse_sq_ft REAL,
        clubhouse_sq_ft REAL,
        fitness_center_sq_ft REAL,
        tennis_courts_number REAL,
        basketball_courts_number REAL,
        other_sport_courts_number REAL,
        walking_biking_trails_miles REAL,
        lakes_or_ponds_number REAL,
        boat_docks_and_slips_number REAL,
        dog_parks_number REAL,
        elevators_number REAL,
        commercial_exposure_sq_ft REAL
    )
    """)

    # Building table
    c.execute("""
    CREATE TABLE IF NOT EXISTS sov_buildings (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        source_file TEXT,
        sheet_name TEXT,
        row_index INTEGER,
        building_number TEXT,
        location_full_address TEXT,
        location_address TEXT,
        location_city TEXT,
        location_state TEXT,
        location_zip TEXT,
        lat REAL,
        long REAL,
        betterview_id TEXT,
        betterview_building_number TEXT,
        units_per_building REAL,
        replacement_cost_tiv REAL,
        num_units REAL,
        livable_sq_ft REAL,
        garage_sq_ft REAL,
        commercial_sq_ft REAL,
        building_class TEXT,
        parking_type TEXT,
        roof_type TEXT,
        smoke_detectors TEXT,
        sprinklered TEXT,
        year_of_construction INTEGER,
        number_of_stories REAL,
        construction_type TEXT
    )
    """)

    conn.commit()
    conn.close()


@contextmanager
def db_session():
    conn = sqlite3.connect(DB_PATH)
    try:
        yield conn
    finally:
        conn.commit()
        conn.close()


def insert_property_records(conn, records):
    if not records:
        return
    cols = [f for f in records[0].__dict__.keys()]
    sql = f"INSERT INTO sov_properties ({','.join(cols)}) VALUES ({','.join(['?'] * len(cols))})"
    conn.executemany(sql, [[getattr(r, c) for c in cols] for r in records])


def insert_building_records(conn, records):
    if not records:
        return
    cols = [f for f in records[0].__dict__.keys()]
    sql = f"INSERT INTO sov_buildings ({','.join(cols)}) VALUES ({','.join(['?'] * len(cols))})"
    conn.executemany(sql, [[getattr(r, c) for c in cols] for r in records])