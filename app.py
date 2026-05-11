from fastapi import FastAPI, UploadFile, File
from fastapi.responses import HTMLResponse
from fastapi.staticfiles import StaticFiles
import uvicorn

app = FastAPI(title="ProPlan Scheduler Web")

# Servir archivos estáticos (HTML, JS, CSS)
app.mount("/static", StaticFiles(directory="static"), name="static")

@app.get("/", response_class=HTMLResponse)
async def read_root():
    with open("static/index.html", "r", encoding="utf-8") as f:
        return f.read()

@app.post("/api/schedule")
async def schedule_production(file: UploadFile = File(...)):
    # Aquí irá la lógica de cargar el archivo, procesar y retornar JSON
    # Para la fase inicial, solo leeremos y confirmaremos que llegó
    contents = await file.read()
    
    from data_loader import load_excel_data
    from scheduler_engine import run_scheduler
    
    # Cargar datos
    df_prog, machines, blocked_dates, color_map = load_excel_data(contents)
    
    if df_prog is None:
        return {"status": "error", "message": "Error al leer el archivo Excel."}
        
    # Ejecutar planificador
    scheduled_tasks, gantt_state = run_scheduler(df_prog, machines, blocked_dates)
    
    # Calcular métricas básicas
    total_ops = len(df_prog)
    machines_used = len([m for m, dates in gantt_state.items() if len(dates) > 0])
    
    return {
        "status": "success",
        "message": "Planificación calculada con éxito",
        "stats": {
            "total_ops": total_ops,
            "machines_used": machines_used
        },
        "tasks": scheduled_tasks,
        "color_map": color_map
    }

from fastapi.responses import Response
from pydantic import BaseModel
from typing import List, Dict, Any

class ExportRequest(BaseModel):
    tasks: List[Dict[str, Any]]
    color_map: Dict[str, str]

@app.post("/api/export_excel")
async def export_excel(req: ExportRequest):
    from gantt_exporter import generate_excel_bytes
    excel_bytes = generate_excel_bytes(req.tasks, req.color_map)
    
    return Response(
        content=excel_bytes,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": "attachment; filename=Planificacion_ProPlan.xlsx"}
    )

if __name__ == "__main__":
    uvicorn.run("app:app", host="0.0.0.0", port=8000, reload=True)
