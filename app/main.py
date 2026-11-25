"""FastAPI приложение."""


from pathlib import Path


from fastapi import FastAPI, HTTPException
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel



BASE_DIR = Path(__file__).resolve().parent.parent
DATA_DIR = BASE_DIR / "data"
OUTPUT_DIR = BASE_DIR / "output"
TEMPLATE_PATH = DATA_DIR / "ПВД-УВД-ДВД-ГПА.xlsx"
REFERENCE_PATH = DATA_DIR / "Тех-Места-ГПА.xlsx"

from .services.reference_service import ReferenceService, ReferenceDataError
from .services.excel_service import ExcelService, ExcelGenerationError, WORK_TYPES



app = FastAPI(title="Формирование ведомости ГПА", version="0.1.0")
reference_service = ReferenceService(REFERENCE_PATH)
excel_service = ExcelService(TEMPLATE_PATH, OUTPUT_DIR, reference_service)

app.mount("/static", StaticFiles(directory=Path(__file__).parent / "static"), name="static")



class GenerateRequest(BaseModel):
    work_type: str
    station: str
    equipment: str



@app.get("/")
def home():
    return FileResponse(Path(__file__).parent / "templates" / "index.html")



@app.get("/api/work-types")
def work_types():
    return WORK_TYPES



@app.get("/api/stations")
def stations():
    try:
        return {"stations": reference_service.get_stations()}
    except ReferenceDataError as exc:
        raise HTTPException(status_code=500, detail=str(exc)) from exc



@app.get("/api/equipment")
def equipment(station: str):
    try:
        return {"equipment": reference_service.get_equipment(station)}
    except ReferenceDataError as exc:
        raise HTTPException(status_code=500, detail=str(exc)) from exc



@app.post("/api/generate")
def generate(request: GenerateRequest):
    try:
        output_path = excel_service.generate(request.work_type, request.station, request.equipment)
    except (ReferenceDataError, ExcelGenerationError) as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc

    return FileResponse(
        path=output_path,
        filename=output_path.name,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
