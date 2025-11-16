import os
from fastapi import UploadFile

def save_temp_file(upload: UploadFile, file_id: str):
    folder = "uploads"
    os.makedirs(folder, exist_ok=True)
    path = os.path.join(folder, f"{file_id}.xlsx")

    with open(path, "wb") as f:
        f.write(upload.file.read())

    return path


def collect_top_left_cells(ws, rows=50, cols=8):
    cells = []
    for r in range(1, rows):
        for c in range(1, cols):
            v = ws.cell(row=r, column=c).value
            if isinstance(v, str) and v.strip():
                cells.append({"cell": ws.cell(row=r, column=c).coordinate, "value": v})
    return cells