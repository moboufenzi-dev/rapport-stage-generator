"""
Pydantic models for the report generator.
"""
from pydantic import BaseModel
from typing import Optional


class ChapterItem(BaseModel):
    """Représente un chapitre ou sous-chapitre du rapport."""
    id: int
    title: str
    level: int
    children: list["ChapterItem"] = []


class GlossaryItem(BaseModel):
    """Entrée du glossaire."""
    term: str
    definition: str


class FigureItem(BaseModel):
    """Entrée de la liste des figures."""
    name: str
    page: str


class GanttTask(BaseModel):
    """Tâche pour le diagramme de Gantt."""
    task: str
    start: str
    end: str


class StyleConfig(BaseModel):
    """Configuration des styles typographiques."""
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
    """Configuration de la mise en page."""
    margin_top: float = 2.5
    margin_bottom: float = 2.5
    margin_left: float = 2.5
    margin_right: float = 2.5
    show_page_number: bool = True
    show_student_name: bool = True


class LogosConfig(BaseModel):
    """Configuration des logos et images."""
    logo_ecole: Optional[str] = None
    logo_entreprise: Optional[str] = None
    image_centrale: Optional[str] = None


class ReportData(BaseModel):
    """Données complètes pour la génération du rapport."""

    # Modèle de page de garde
    cover_model: str = "classique"

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
    sujet_stage: str = ""
    poste: str = ""

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
