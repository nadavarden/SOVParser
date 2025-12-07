import pyodbc
import pandas as pd
import os
import sys
from app.parser.sov_parser_ai import SOVParser
import json
import zipfile

# Load environment variables if python-dotenv is installed
try:
    from dotenv import load_dotenv

    load_dotenv()
except:
    pass

OUTPUT_JSON_DIR = "tests/results"


def extract_embedded_excel(zip_path, output_dir="downloaded_sovs/extracted"):
    os.makedirs(output_dir, exist_ok=True)

    with zipfile.ZipFile(zip_path, "r") as z:
        for name in z.namelist():
            if name.lower().endswith((".xlsx", ".xls")):
                extracted_path = os.path.join(output_dir, os.path.basename(name))
                z.extract(name, output_dir)
                final_path = os.path.join(output_dir, name)
                print(f"📦 Extracted embedded Excel → {final_path}")
                return final_path

    raise Exception("❌ No Excel file found inside the ZIP wrapper")


def list_files_for_control_number(control_number: int, distinct: bool = True):
    # Get the connection string from the imported module
    connection_string = os.getenv(
        "CUSTOMCONNSTR_AZURE_SQL_ARDEN_PROD_CONNECTION_STRING"
    )
    # Connect to the database
    try:
        conn = pyodbc.connect(connection_string)
        print("Connection successful!")
        # Define the query
        distinct_sql = "DISTINCT" if distinct else ""
        query = f"""
            SELECT {distinct_sql}
                q.ControlNo,
                ds.DocumentStoreGUID,
                ds.FileName,
                ds.FileAssociation,
                ds.DateAdded,
                ds.FolderID
            FROM dbo.tblQuotes q
            INNER JOIN dbo.tblDocumentAssociations da
                ON da.ControlGuid = q.ControlGuid
            INNER JOIN dbo.tblDocumentStore ds
                ON ds.DocumentStoreGUID = da.DocumentStoreGUID
            WHERE q.ControlNo = ?
                AND LOWER(ISNULL(ds.FileName, '')) LIKE '%sov%'
            ORDER BY ds.DateAdded DESC;
            """
        # Read into a pandas DataFrame
        df = pd.read_sql(query, conn, params=[control_number])
        if df.empty:
            print(f"No files found for ControlNo={control_number}")
            return df

        print(f"Found {len(df)} file(s) for ControlNo={control_number}:\n")
        for name in df["FileName"].fillna("").tolist():
            print(name)

        return df
    finally:
        conn.close()


def download_sov_file(conn, document_guid, file_name, output_dir="downloaded_sovs"):
    """
    Fetches binary Excel file from tblDocumentStore and saves it locally.
    """
    os.makedirs(output_dir, exist_ok=True)

    cursor = conn.cursor()
    cursor.execute(
        """
        SELECT Document
        FROM dbo.tblDocumentStore
        WHERE DocumentStoreGUID = ?
    """,
        (document_guid,),
    )

    row = cursor.fetchone()
    if not row or not row[0]:
        print(f"❌ No file data found for GUID: {document_guid}")
        return None

    local_path = os.path.join(output_dir, file_name)

    with open(local_path, "wb") as f:
        f.write(row[0])

    print(f"📄 Saved Excel SOV file: {local_path}")
    return local_path


def process_sov(control_number: int):
    """
    Fetches ONE Excel SOV file for a control number, parses it, and writes output JSON.
    """
    connection_string = os.getenv(
        "CUSTOMCONNSTR_AZURE_SQL_ARDEN_PROD_CONNECTION_STRING"
    )
    conn = pyodbc.connect(connection_string)

    try:
        print("Connected to SQL!")

        # Fetch EXACTLY one SOV
        query = """
            SELECT TOP 1
                ds.DocumentStoreGUID,
                ds.FileName,
                ds.DateAdded
            FROM dbo.tblQuotes q
            INNER JOIN dbo.tblDocumentAssociations da
                ON da.ControlGuid = q.ControlGuid
            INNER JOIN dbo.tblDocumentStore ds
                ON ds.DocumentStoreGUID = da.DocumentStoreGUID
            WHERE q.ControlNo = ?
              AND LOWER(ISNULL(ds.FileName, '')) LIKE '%%sov%%'
              AND (ds.FileName LIKE '%.xls' OR ds.FileName LIKE '%.xlsx')
            ORDER BY ds.DateAdded DESC;
        """

        df = pd.read_sql(query, conn, params=[control_number])

        if df.empty:
            print(f"❌ No Excel SOV found for control number {control_number}")
            return

        # We expect exactly 1 row
        row = df.iloc[0]
        guid = row["DocumentStoreGUID"]
        file_name = row["FileName"]

        print("\n==============================")
        print(f"📄 Found SOV: {file_name}")
        print("==============================")

        # 1. Download the Excel SOV file
        local_zip = download_sov_file(conn, guid, file_name)
        # Extract real Excel from inside the wrapper ZIP
        real_excel = extract_embedded_excel(local_zip)

        if not local_zip:
            print("❌ Failed to download SOV file.")
            return

        # 2. Parse the Excel using your AI agent
        print("🤖 Running parser...")
        parser = SOVParser()

        parsed_json = parser.parse_excel(real_excel)

        # 3. Save JSON to output directory
        os.makedirs(OUTPUT_JSON_DIR, exist_ok=True)
        json_filename = os.path.splitext(file_name)[0] + ".json"
        json_path = os.path.join(OUTPUT_JSON_DIR, json_filename)

        with open(json_path, "w") as jf:
            json.dump(parsed_json, jf, indent=4)

        print(f"✅ SOV parsed successfully!")
        print(f"JSON saved → {json_path}")

    finally:
        conn.close()


if __name__ == "__main__":
    control_number = 53411
    process_sov(control_number)
