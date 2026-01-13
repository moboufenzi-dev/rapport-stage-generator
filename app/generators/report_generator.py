"""
Générateur principal de rapport de stage.
"""
import io
from docx import Document
from docx.shared import Cm

from .utils import (
    format_date_fr,
    calculate_duration,
    setup_document_styles,
    setup_header_with_logos,
    setup_footer_with_page_number,
)
from .covers import get_cover_generator
from .sections import (
    generate_toc_section,
    generate_thanks_section,
    generate_abstract_section,
    generate_chapters,
    generate_annexes_section,
)


def generate_report(data) -> io.BytesIO:
    """Génère le rapport de stage complet."""
    doc = Document()

    # Configuration de la page (A4)
    section = doc.sections[0]
    section.page_width = Cm(21)
    section.page_height = Cm(29.7)
    section.top_margin = Cm(data.page.margin_top)
    section.bottom_margin = Cm(data.page.margin_bottom)
    section.left_margin = Cm(data.page.margin_left)
    section.right_margin = Cm(data.page.margin_right)

    # Première page différente (pas de header/footer sur page de garde)
    section.different_first_page_header_footer = True

    # Configurer les styles
    setup_document_styles(doc, data.style)

    # Variables
    duree = calculate_duration(data.date_debut, data.date_fin)
    date_debut_fr = format_date_fr(data.date_debut)
    date_fin_fr = format_date_fr(data.date_fin)

    # Page de garde
    if data.include_cover:
        cover_model = getattr(data, 'cover_model', 'classique')
        cover_generator = get_cover_generator(cover_model)
        cover_generator(doc, data, date_debut_fr, date_fin_fr, duree)
        doc.add_page_break()

    # Configurer header et footer pour les pages suivantes
    setup_header_with_logos(section, data)
    setup_footer_with_page_number(section, data)

    # Table des matières
    if data.include_toc:
        generate_toc_section(doc, data)

    # Remerciements
    if data.include_thanks:
        generate_thanks_section(doc, data)

    # Résumé/Abstract
    if data.include_abstract:
        generate_abstract_section(doc, data)

    # Chapitres
    generate_chapters(doc, data)

    # Annexes
    if data.include_annexes:
        generate_annexes_section(doc, data)

    # Sauvegarder
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer
