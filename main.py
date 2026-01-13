from fastapi import FastAPI, Request
from fastapi.responses import HTMLResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from pathlib import Path
import io

# Import depuis les nouveaux modules
from app.models.schemas import ReportData
from app.generators import generate_report

# Chemins absolus pour production
BASE_DIR = Path(__file__).resolve().parent

app = FastAPI(title="Générateur de Rapport de Stage v3")

templates = Jinja2Templates(directory=str(BASE_DIR / "templates"))
app.mount("/static", StaticFiles(directory=str(BASE_DIR / "static")), name="static")


# ===== ROUTES =====

@app.get("/", response_class=HTMLResponse)
async def index(request: Request):
    return templates.TemplateResponse("index.html", {"request": request})


@app.post("/generate")
async def generate(data: ReportData):
    # Générer le document Word
    doc_buffer = generate_report(data)

    # Nom du fichier
    filename = f"rapport_stage_{data.nom or 'rapport'}.docx"

    return StreamingResponse(
        io.BytesIO(doc_buffer.getvalue()),
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers={"Content-Disposition": f"attachment; filename={filename}"}
    )


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)
