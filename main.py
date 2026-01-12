from fastapi import FastAPI, Request
from fastapi.responses import HTMLResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from pydantic import BaseModel
from typing import Optional
from pathlib import Path
import io

from generator import generate_report

# Chemins absolus pour production
BASE_DIR = Path(__file__).resolve().parent

app = FastAPI(title="Générateur de Rapport de Stage v3")

templates = Jinja2Templates(directory=str(BASE_DIR / "templates"))
app.mount("/static", StaticFiles(directory=str(BASE_DIR / "static")), name="static")


# ===== MODELS =====

class ChapterItem(BaseModel):
    id: int
    title: str
    level: int
    children: list["ChapterItem"] = []


class GlossaryItem(BaseModel):
    term: str
    definition: str


class FigureItem(BaseModel):
    name: str
    page: str


class GanttTask(BaseModel):
    task: str
    start: str
    end: str


class StyleConfig(BaseModel):
    font_family: str = "Times New Roman"
    font_size: int = 12
    line_spacing: float = 1.5
    title1_size: int = 16
    title1_bold: bool = True
    title1_color: str = "#1a365d"
    title2_size: int = 14
    title2_bold: bool = True
    title2_color: str = "#000000"
    title3_size: int = 12
    title3_italic: bool = True
    title3_color: str = "#333333"


class PageConfig(BaseModel):
    margin_top: float = 2.5
    margin_bottom: float = 2.5
    margin_left: float = 2.5
    margin_right: float = 2.5
    show_page_number: bool = True
    show_student_name: bool = True


class LogosConfig(BaseModel):
    logo_ecole: Optional[str] = None
    logo_entreprise: Optional[str] = None
    image_centrale: Optional[str] = None


class ReportData(BaseModel):
    # Modèle de page de garde
    cover_model: str = "classique"  # classique, moderne, corporate

    # Étudiant
    prenom: str = ""
    nom: str = ""
    formation: str = ""
    ecole: str = ""
    annee_scolaire: str = ""

    # Entreprise
    entreprise_nom: str = ""
    entreprise_secteur: str = ""
    entreprise_ville: str = ""
    tuteur_nom: str = ""
    tuteur_poste: str = ""
    tuteur_academique_nom: str = ""
    tuteur_academique_poste: str = ""

    # Stage
    date_debut: str = ""
    date_fin: str = ""
    sujet_stage: str = ""  # Titre/Sujet du stage (affiché en gros sur page de garde)
    poste: str = ""  # Intitulé du poste occupé

    # Structure
    chapters: list[ChapterItem] = []
    glossary: list[GlossaryItem] = []
    figures: list[FigureItem] = []
    ganttTasks: list[GanttTask] = []

    include_cover: bool = True
    include_thanks: bool = True
    include_toc: bool = True
    include_figures_list: bool = False
    include_abstract: bool = False
    include_glossary: bool = False
    include_gantt: bool = False
    include_annexes: bool = True

    # Mise en page
    style: StyleConfig = StyleConfig()
    page: PageConfig = PageConfig()
    logos: LogosConfig = LogosConfig()


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
