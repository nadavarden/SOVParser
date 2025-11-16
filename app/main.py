from fastapi import FastAPI
from app.router_parse import router as parse_router
from app.db import init_db

app = FastAPI(
    title="SOV Parsing API",
    version="1.0.0",
    description="API for ingesting, parsing, and storing Statement of Values (SOV) Excel files."
)

@app.on_event("startup")
def startup():
    init_db()

app.include_router(parse_router, prefix="/api")

@app.get("/health")
def health():
    return {"status": "ok"}