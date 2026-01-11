from docx import Document
from docx.shared import Inches, Pt, Cm, RGBColor, Twips
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT, WD_TAB_LEADER
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.section import WD_ORIENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import io
import base64
from datetime import datetime
from PIL import Image as PILImage


def hex_to_rgb(hex_color: str) -> RGBColor:
    hex_color = hex_color.lstrip('#')
    return RGBColor(int(hex_color[0:2], 16), int(hex_color[2:4], 16), int(hex_color[4:6], 16))


def decode_base64_image(base64_str: str) -> io.BytesIO:
    """Décode une image base64 et la convertit en PNG pour compatibilité python-docx."""
    if not base64_str:
        raise ValueError("Image base64 vide")
    if ',' in base64_str:
        base64_str = base64_str.split(',')[1]
    decoded = base64.b64decode(base64_str)

    # Convertir en PNG via Pillow pour assurer la compatibilité
    try:
        input_stream = io.BytesIO(decoded)
        pil_image = PILImage.open(input_stream)

        # Convertir en RGB si nécessaire (pour les images RGBA ou autres modes)
        if pil_image.mode in ('RGBA', 'LA', 'P'):
            # Créer un fond blanc pour les images avec transparence
            background = PILImage.new('RGB', pil_image.size, (255, 255, 255))
            if pil_image.mode == 'P':
                pil_image = pil_image.convert('RGBA')
            background.paste(pil_image, mask=pil_image.split()[-1] if pil_image.mode == 'RGBA' else None)
            pil_image = background
        elif pil_image.mode != 'RGB':
            pil_image = pil_image.convert('RGB')

        # Sauvegarder en PNG
        output_stream = io.BytesIO()
        pil_image.save(output_stream, format='PNG')
        output_stream.seek(0)
        return output_stream
    except Exception as e:
        print(f"[DEBUG] Conversion Pillow échouée: {e}, utilisation directe")
        return io.BytesIO(decoded)


def format_date_fr(date_str: str) -> str:
    if not date_str:
        return "[Date]"
    try:
        d = datetime.strptime(date_str, "%Y-%m-%d")
        mois = ["janvier", "février", "mars", "avril", "mai", "juin",
                "juillet", "août", "septembre", "octobre", "novembre", "décembre"]
        return f"{d.day} {mois[d.month-1]} {d.year}"
    except:
        return date_str


def calculate_duration(date_debut: str, date_fin: str) -> str:
    if not date_debut or not date_fin:
        return "[durée]"
    try:
        d1 = datetime.strptime(date_debut, "%Y-%m-%d")
        d2 = datetime.strptime(date_fin, "%Y-%m-%d")
        months = (d2.year - d1.year) * 12 + d2.month - d1.month
        return f"{months} mois" if months != 1 else "1 mois"
    except:
        return "[durée]"


def add_page_number_field(paragraph):
    """Ajoute un champ PAGE pour le numéro de page."""
    run = paragraph.add_run()
    fldChar1 = OxmlElement('w:fldChar')
    fldChar1.set(qn('w:fldCharType'), 'begin')

    instrText = OxmlElement('w:instrText')
    instrText.set(qn('xml:space'), 'preserve')
    instrText.text = " PAGE "

    fldChar2 = OxmlElement('w:fldChar')
    fldChar2.set(qn('w:fldCharType'), 'separate')

    # Placeholder text
    text_elem = OxmlElement('w:t')
    text_elem.text = "1"

    fldChar3 = OxmlElement('w:fldChar')
    fldChar3.set(qn('w:fldCharType'), 'end')

    run._r.append(fldChar1)
    run._r.append(instrText)
    run._r.append(fldChar2)
    run._r.append(text_elem)
    run._r.append(fldChar3)


def create_toc_entry(doc, text, level=1, page=""):
    """Crée une entrée de table des matières avec tabulation et points de suite."""
    paragraph = doc.add_paragraph()
    paragraph.paragraph_format.first_line_indent = Cm(0)
    paragraph.paragraph_format.space_after = Pt(4)

    # Indentation selon le niveau
    if level == 1:
        paragraph.paragraph_format.left_indent = Cm(0)
    elif level == 2:
        paragraph.paragraph_format.left_indent = Cm(0.75)
    else:
        paragraph.paragraph_format.left_indent = Cm(1.5)

    # Tabulation avec points de suite à 15cm (pour laisser de la place)
    tab_stops = paragraph.paragraph_format.tab_stops
    tab_stops.add_tab_stop(Cm(15), WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.DOTS)

    # Texte du titre
    run = paragraph.add_run(text)
    if level == 1:
        run.bold = True
        run.font.size = Pt(11)
    elif level == 2:
        run.font.size = Pt(10)
    else:
        run.font.size = Pt(10)
        run.italic = True

    # Tabulation + numéro de page
    paragraph.add_run("\t")
    page_run = paragraph.add_run(page)
    page_run.font.size = Pt(11 if level == 1 else 10)
    if level == 1:
        page_run.bold = True


def create_toc(doc, data):
    """Crée une table des matières avec numérotation automatique."""
    page_num = 3

    if data.include_thanks:
        create_toc_entry(doc, "REMERCIEMENTS", 1, str(page_num))
        page_num += 1

    if data.include_abstract:
        create_toc_entry(doc, "RÉSUMÉ", 1, str(page_num))
        page_num += 1

    for chapter_idx, chapter in enumerate(data.chapters, 1):
        # Numéroter les chapitres dans la TOC
        create_toc_entry(doc, f"{chapter_idx}. {chapter.title}", 1, str(page_num))
        for sub_idx, sub in enumerate(chapter.children, 1):
            # Numéroter les sous-chapitres
            create_toc_entry(doc, f"{chapter_idx}.{sub_idx}. {sub.title}", 2, str(page_num))
            # Sous-sous-chapitres si présents
            if hasattr(sub, 'children') and sub.children:
                for subsub_idx, subsub in enumerate(sub.children, 1):
                    create_toc_entry(doc, f"{chapter_idx}.{sub_idx}.{subsub_idx}. {subsub.title}", 3, str(page_num))
        page_num += 1

    if data.include_annexes:
        create_toc_entry(doc, "ANNEXES", 1, str(page_num))


def setup_document_styles(doc, style_config):
    """Configure les styles du document avec indentation professionnelle."""
    styles = doc.styles

    # Normal - texte avec indentation
    normal = styles['Normal']
    normal.font.name = style_config.font_family
    normal.font.size = Pt(style_config.font_size)
    normal.paragraph_format.line_spacing = style_config.line_spacing
    normal.paragraph_format.first_line_indent = Cm(1.25)  # Alinéa première ligne
    normal.paragraph_format.space_after = Pt(6)

    # Heading 1 - Chapitres principaux (pas d'indentation)
    h1 = styles['Heading 1']
    h1.font.name = style_config.font_family
    h1.font.size = Pt(style_config.title1_size)
    h1.font.bold = style_config.title1_bold
    h1.font.color.rgb = hex_to_rgb(style_config.title1_color)
    h1.paragraph_format.space_before = Pt(18)
    h1.paragraph_format.space_after = Pt(12)
    h1.paragraph_format.left_indent = Cm(0)
    h1.paragraph_format.first_line_indent = Cm(0)

    # Heading 2 - Sous-sections (indentation 1 niveau)
    h2 = styles['Heading 2']
    h2.font.name = style_config.font_family
    h2.font.size = Pt(style_config.title2_size)
    h2.font.bold = style_config.title2_bold
    h2.font.color.rgb = hex_to_rgb(style_config.title2_color)
    h2.paragraph_format.space_before = Pt(14)
    h2.paragraph_format.space_after = Pt(8)
    h2.paragraph_format.left_indent = Cm(0.75)
    h2.paragraph_format.first_line_indent = Cm(0)

    # Heading 3 - Sous-sous-sections (indentation 2 niveaux)
    h3 = styles['Heading 3']
    h3.font.name = style_config.font_family
    h3.font.size = Pt(style_config.title3_size)
    h3.font.italic = style_config.title3_italic
    h3.font.bold = False
    h3.font.color.rgb = hex_to_rgb(style_config.title3_color)
    h3.paragraph_format.space_before = Pt(10)
    h3.paragraph_format.space_after = Pt(6)
    h3.paragraph_format.left_indent = Cm(1.5)
    h3.paragraph_format.first_line_indent = Cm(0)


def setup_header_with_logos(section, data):
    """Configure l'en-tête avec logos - symétrique."""
    header = section.header

    # Supprimer le contenu existant
    for para in header.paragraphs:
        p = para._element
        p.getparent().remove(p)

    # Créer un tableau symétrique (largeurs fixes) - total 16cm
    tbl = header.add_table(rows=1, cols=3, width=Cm(16))
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl.autofit = False

    # Largeurs symétriques : gauche 5cm | centre 6cm | droite 5cm
    tbl.columns[0].width = Cm(5)
    tbl.columns[1].width = Cm(6)
    tbl.columns[2].width = Cm(5)

    remove_table_borders(tbl)

    row = tbl.rows[0]

    # Logo école (gauche)
    cell_left = row.cells[0]
    para_left = cell_left.paragraphs[0]
    para_left.alignment = WD_ALIGN_PARAGRAPH.LEFT
    para_left.paragraph_format.first_line_indent = Cm(0)
    if data.logos.logo_ecole and len(data.logos.logo_ecole) > 100:
        try:
            img = decode_base64_image(data.logos.logo_ecole)
            run = para_left.add_run()
            run.add_picture(img, height=Cm(1.2))
        except:
            pass

    # Centre (vide)
    cell_center = row.cells[1]
    para_center = cell_center.paragraphs[0]
    para_center.alignment = WD_ALIGN_PARAGRAPH.CENTER
    para_center.paragraph_format.first_line_indent = Cm(0)

    # Logo entreprise (droite)
    cell_right = row.cells[2]
    para_right = cell_right.paragraphs[0]
    para_right.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    para_right.paragraph_format.first_line_indent = Cm(0)
    if data.logos.logo_entreprise and len(data.logos.logo_entreprise) > 100:
        try:
            img = decode_base64_image(data.logos.logo_entreprise)
            run = para_right.add_run()
            run.add_picture(img, height=Cm(1.2))
        except:
            pass


def setup_footer_with_page_number(section, data):
    """Configure le pied de page : entreprise à gauche, numéro au centre, nom à droite."""
    footer = section.footer

    # Supprimer le contenu existant
    for para in footer.paragraphs:
        p = para._element
        p.getparent().remove(p)

    # Créer un tableau symétrique pour le footer - total 16cm
    tbl = footer.add_table(rows=1, cols=3, width=Cm(16))
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl.autofit = False

    # Largeurs symétriques
    tbl.columns[0].width = Cm(5)
    tbl.columns[1].width = Cm(6)
    tbl.columns[2].width = Cm(5)

    remove_table_borders(tbl)

    row = tbl.rows[0]

    # Entreprise (gauche)
    cell_left = row.cells[0]
    para_left = cell_left.paragraphs[0]
    para_left.alignment = WD_ALIGN_PARAGRAPH.LEFT
    para_left.paragraph_format.first_line_indent = Cm(0)
    if data.entreprise_nom:
        run = para_left.add_run(data.entreprise_nom)
        run.font.size = Pt(9)
        run.font.color.rgb = RGBColor(100, 100, 100)

    # Numéro de page (centre)
    cell_center = row.cells[1]
    para_center = cell_center.paragraphs[0]
    para_center.alignment = WD_ALIGN_PARAGRAPH.CENTER
    para_center.paragraph_format.first_line_indent = Cm(0)
    if data.page.show_page_number:
        run = para_center.add_run("- ")
        run.font.size = Pt(9)
        add_page_number_field(para_center)
        run = para_center.add_run(" -")
        run.font.size = Pt(9)

    # Nom étudiant (droite)
    cell_right = row.cells[2]
    para_right = cell_right.paragraphs[0]
    para_right.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    para_right.paragraph_format.first_line_indent = Cm(0)
    if data.page.show_student_name and data.nom:
        run = para_right.add_run(f"{data.prenom} {data.nom}")
        run.font.size = Pt(9)
        run.font.color.rgb = RGBColor(100, 100, 100)


def set_cell_shading(cell, color_hex):
    """Applique une couleur de fond à une cellule."""
    shading_elm = OxmlElement('w:shd')
    shading_elm.set(qn('w:fill'), color_hex.lstrip('#'))
    cell._tc.get_or_add_tcPr().append(shading_elm)


def remove_table_borders(table):
    """Supprime toutes les bordures d'un tableau."""
    tbl_element = table._tbl
    tbl_pr = tbl_element.find(qn('w:tblPr'))
    if tbl_pr is None:
        tbl_pr = OxmlElement('w:tblPr')
        tbl_element.insert(0, tbl_pr)
    tbl_borders = OxmlElement('w:tblBorders')
    for border_name in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
        border = OxmlElement(f'w:{border_name}')
        border.set(qn('w:val'), 'nil')
        tbl_borders.append(border)
    tbl_pr.append(tbl_borders)


def add_centered_paragraph(doc, space_after=0):
    """Crée un paragraphe centré sans indentation (pour page de garde)."""
    para = doc.add_paragraph()
    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    para.paragraph_format.first_line_indent = Cm(0)
    para.paragraph_format.left_indent = Cm(0)
    para.paragraph_format.space_after = Pt(space_after)
    para.paragraph_format.space_before = Pt(0)
    return para


# ════════════════════════════════════════════════════════════════
#                    MODÈLES DE PAGE DE GARDE
# ════════════════════════════════════════════════════════════════

def generate_cover_classique(doc, data, date_debut_fr, date_fin_fr, duree):
    """Page de garde style classique académique - alignement centré propre."""

    # Vérifier si les images existent vraiment
    has_logo_ecole = data.logos.logo_ecole and len(data.logos.logo_ecole) > 100
    has_logo_entreprise = data.logos.logo_entreprise and len(data.logos.logo_entreprise) > 100
    has_image_centrale = data.logos.image_centrale and len(data.logos.image_centrale) > 100

    # Logos en haut (tableau centré)
    if has_logo_ecole or has_logo_entreprise:
        logo_table = doc.add_table(rows=1, cols=2)
        logo_table.alignment = WD_TABLE_ALIGNMENT.CENTER
        logo_table.autofit = False
        logo_table.columns[0].width = Cm(8)
        logo_table.columns[1].width = Cm(8)
        remove_table_borders(logo_table)
        row = logo_table.rows[0]

        if has_logo_ecole:
            try:
                para = row.cells[0].paragraphs[0]
                para.alignment = WD_ALIGN_PARAGRAPH.LEFT
                para.paragraph_format.first_line_indent = Cm(0)
                img = decode_base64_image(data.logos.logo_ecole)
                para.add_run().add_picture(img, height=Cm(2))
            except Exception as e:
                print(f"Erreur logo école: {e}")

        if has_logo_entreprise:
            try:
                para = row.cells[1].paragraphs[0]
                para.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                para.paragraph_format.first_line_indent = Cm(0)
                img = decode_base64_image(data.logos.logo_entreprise)
                para.add_run().add_picture(img, height=Cm(2))
            except Exception as e:
                print(f"Erreur logo entreprise: {e}")

    # Espace avant titre
    add_centered_paragraph(doc, 30)

    # TITRE
    title = add_centered_paragraph(doc, 6)
    run = title.add_run("RAPPORT DE STAGE")
    run.bold = True
    run.font.size = Pt(28)
    run.font.color.rgb = hex_to_rgb(data.style.title1_color)

    # Formation
    formation = add_centered_paragraph(doc, 15)
    run = formation.add_run(data.formation or "[Formation]")
    run.font.size = Pt(13)
    run.italic = True

    # Ligne séparatrice
    line = add_centered_paragraph(doc, 15)
    run = line.add_run("━" * 40)
    run.font.color.rgb = hex_to_rgb(data.style.title1_color)

    # Image centrale optionnelle
    if has_image_centrale:
        try:
            img_para = add_centered_paragraph(doc, 15)
            img = decode_base64_image(data.logos.image_centrale)
            img_para.add_run().add_picture(img, height=Cm(4.5))
        except Exception as e:
            print(f"Erreur image centrale: {e}")

    # Étudiant
    student = add_centered_paragraph(doc, 2)
    run = student.add_run(f"{data.prenom or '[Prénom]'} {data.nom or '[NOM]'}")
    run.bold = True
    run.font.size = Pt(16)

    school = add_centered_paragraph(doc, 2)
    run = school.add_run(f"{data.ecole or '[Établissement]'}")
    run.font.size = Pt(12)

    year = add_centered_paragraph(doc, 15)
    run = year.add_run(f"Année {data.annee_scolaire or '[Année]'}")
    run.font.size = Pt(11)
    run.italic = True

    # Ligne séparatrice
    line2 = add_centered_paragraph(doc, 15)
    run = line2.add_run("━" * 40)
    run.font.color.rgb = hex_to_rgb(data.style.title1_color)

    # Entreprise
    ent_title = add_centered_paragraph(doc, 2)
    run = ent_title.add_run("Stage réalisé chez")
    run.font.size = Pt(10)
    run.font.color.rgb = RGBColor(120, 120, 120)

    ent_name = add_centered_paragraph(doc, 2)
    run = ent_name.add_run(data.entreprise_nom or "[Entreprise]")
    run.bold = True
    run.font.size = Pt(15)
    run.font.color.rgb = hex_to_rgb(data.style.title1_color)

    if data.entreprise_ville:
        ville = add_centered_paragraph(doc, 6)
        run = ville.add_run(data.entreprise_ville)
        run.font.size = Pt(11)

    # Dates
    dates = add_centered_paragraph(doc, 25)
    run = dates.add_run(f"Du {date_debut_fr} au {date_fin_fr} ({duree})")
    run.font.size = Pt(11)

    # Tuteurs (tableau centré)
    tuteur_table = doc.add_table(rows=1, cols=2)
    tuteur_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    tuteur_table.autofit = False
    tuteur_table.columns[0].width = Cm(8)
    tuteur_table.columns[1].width = Cm(8)
    remove_table_borders(tuteur_table)

    cell_left = tuteur_table.rows[0].cells[0]
    p1 = cell_left.paragraphs[0]
    p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p1.paragraph_format.first_line_indent = Cm(0)
    r1 = p1.add_run("Tuteur entreprise\n")
    r1.font.size = Pt(9)
    r1.font.color.rgb = RGBColor(120, 120, 120)
    run1 = p1.add_run(data.tuteur_nom or "[Nom]")
    run1.bold = True
    run1.font.size = Pt(10)

    cell_right = tuteur_table.rows[0].cells[1]
    p2 = cell_right.paragraphs[0]
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p2.paragraph_format.first_line_indent = Cm(0)
    r2 = p2.add_run("Tuteur académique\n")
    r2.font.size = Pt(9)
    r2.font.color.rgb = RGBColor(120, 120, 120)
    run2 = p2.add_run(data.tuteur_academique_nom or "[Nom]")
    run2.bold = True
    run2.font.size = Pt(10)


def generate_cover_moderne(doc, data, date_debut_fr, date_fin_fr, duree):
    """Page de garde style moderne - alignement centré propre."""

    # Vérifier les images
    has_logo_ecole = data.logos.logo_ecole and len(data.logos.logo_ecole) > 100
    has_logo_entreprise = data.logos.logo_entreprise and len(data.logos.logo_entreprise) > 100
    has_image_centrale = data.logos.image_centrale and len(data.logos.image_centrale) > 100

    # Logos en haut (tableau centré)
    if has_logo_ecole or has_logo_entreprise:
        logo_table = doc.add_table(rows=1, cols=2)
        logo_table.alignment = WD_TABLE_ALIGNMENT.CENTER
        logo_table.autofit = False
        logo_table.columns[0].width = Cm(8)
        logo_table.columns[1].width = Cm(8)
        remove_table_borders(logo_table)
        row = logo_table.rows[0]

        if has_logo_ecole:
            try:
                para = row.cells[0].paragraphs[0]
                para.alignment = WD_ALIGN_PARAGRAPH.LEFT
                para.paragraph_format.first_line_indent = Cm(0)
                img = decode_base64_image(data.logos.logo_ecole)
                para.add_run().add_picture(img, height=Cm(2))
            except Exception as e:
                print(f"Erreur logo école: {e}")

        if has_logo_entreprise:
            try:
                para = row.cells[1].paragraphs[0]
                para.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                para.paragraph_format.first_line_indent = Cm(0)
                img = decode_base64_image(data.logos.logo_entreprise)
                para.add_run().add_picture(img, height=Cm(2))
            except Exception as e:
                print(f"Erreur logo entreprise: {e}")

    # Espace
    space = add_centered_paragraph(doc, 10)

    # Bandeau image centrale
    if has_image_centrale:
        try:
            img_para = add_centered_paragraph(doc, 15)
            img = decode_base64_image(data.logos.image_centrale)
            img_para.add_run().add_picture(img, width=Cm(16), height=Cm(4.5))
        except Exception as e:
            print(f"Erreur image centrale: {e}")

    # TITRE
    title = add_centered_paragraph(doc, 6)
    run = title.add_run("RAPPORT DE STAGE")
    run.bold = True
    run.font.size = Pt(28)
    run.font.color.rgb = hex_to_rgb(data.style.title1_color)

    # Poste / Mission
    if data.poste:
        poste = add_centered_paragraph(doc, 6)
        run = poste.add_run(data.poste)
        run.font.size = Pt(13)
        run.italic = True

    # Ligne décorative
    line = add_centered_paragraph(doc, 15)
    run = line.add_run("━" * 40)
    run.font.color.rgb = hex_to_rgb(data.style.title1_color)

    # Étudiant
    student = add_centered_paragraph(doc, 2)
    run = student.add_run(f"{data.prenom or '[Prénom]'} {data.nom or '[NOM]'}")
    run.bold = True
    run.font.size = Pt(16)

    formation = add_centered_paragraph(doc, 2)
    run = formation.add_run(data.formation or "[Formation]")
    run.font.size = Pt(12)

    school = add_centered_paragraph(doc, 15)
    run = school.add_run(f"{data.ecole or '[Établissement]'} • {data.annee_scolaire or '[Année]'}")
    run.font.size = Pt(11)
    run.font.color.rgb = RGBColor(100, 100, 100)

    # Ligne
    line2 = add_centered_paragraph(doc, 15)
    run = line2.add_run("━" * 40)
    run.font.color.rgb = hex_to_rgb(data.style.title1_color)

    # Entreprise
    ent_label = add_centered_paragraph(doc, 2)
    run = ent_label.add_run("Stage réalisé chez")
    run.font.size = Pt(10)
    run.font.color.rgb = RGBColor(120, 120, 120)

    ent = add_centered_paragraph(doc, 2)
    run = ent.add_run(data.entreprise_nom or "[Entreprise]")
    run.bold = True
    run.font.size = Pt(15)
    run.font.color.rgb = hex_to_rgb(data.style.title1_color)

    if data.entreprise_ville:
        ville = add_centered_paragraph(doc, 6)
        run = ville.add_run(data.entreprise_ville)
        run.font.size = Pt(11)

    # Dates
    dates = add_centered_paragraph(doc, 20)
    run = dates.add_run(f"Du {date_debut_fr} au {date_fin_fr} ({duree})")
    run.font.size = Pt(11)

    # Tuteurs (tableau centré)
    tuteur_table = doc.add_table(rows=1, cols=2)
    tuteur_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    tuteur_table.autofit = False
    tuteur_table.columns[0].width = Cm(8)
    tuteur_table.columns[1].width = Cm(8)
    remove_table_borders(tuteur_table)

    cell_left = tuteur_table.rows[0].cells[0]
    p1 = cell_left.paragraphs[0]
    p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p1.paragraph_format.first_line_indent = Cm(0)
    r1 = p1.add_run("Tuteur entreprise\n")
    r1.font.size = Pt(9)
    r1.font.color.rgb = RGBColor(120, 120, 120)
    run1 = p1.add_run(data.tuteur_nom or "[Nom]")
    run1.bold = True
    run1.font.size = Pt(10)

    cell_right = tuteur_table.rows[0].cells[1]
    p2 = cell_right.paragraphs[0]
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p2.paragraph_format.first_line_indent = Cm(0)
    r2 = p2.add_run("Tuteur académique\n")
    r2.font.size = Pt(9)
    r2.font.color.rgb = RGBColor(120, 120, 120)
    run2 = p2.add_run(data.tuteur_academique_nom or "[Nom]")
    run2.bold = True
    run2.font.size = Pt(10)


def add_cell_paragraph(cell, space_after=0):
    """Crée un paragraphe sans indentation dans une cellule."""
    para = cell.add_paragraph()
    para.paragraph_format.first_line_indent = Cm(0)
    para.paragraph_format.left_indent = Cm(0)
    para.paragraph_format.space_after = Pt(space_after)
    return para


def generate_cover_corporate(doc, data, date_debut_fr, date_fin_fr, duree):
    """Page de garde style corporate - alignement propre."""

    # Vérifier si les images existent vraiment
    has_logo_ecole = data.logos.logo_ecole and len(data.logos.logo_ecole) > 100
    has_logo_entreprise = data.logos.logo_entreprise and len(data.logos.logo_entreprise) > 100
    has_image_centrale = data.logos.image_centrale and len(data.logos.image_centrale) > 100

    # Créer un tableau 2 colonnes (sidebar + contenu)
    main_table = doc.add_table(rows=1, cols=2)
    main_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    main_table.autofit = False
    remove_table_borders(main_table)

    # Définir les largeurs
    main_table.columns[0].width = Cm(4.5)
    main_table.columns[1].width = Cm(11.5)

    sidebar = main_table.rows[0].cells[0]
    content = main_table.rows[0].cells[1]

    # Couleur de la sidebar
    primary_color = data.style.title1_color.lstrip('#')
    set_cell_shading(sidebar, primary_color)

    # Contenu sidebar - aligné au centre
    sidebar_para = sidebar.paragraphs[0]
    sidebar_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sidebar_para.paragraph_format.first_line_indent = Cm(0)
    sidebar_para.paragraph_format.space_before = Pt(40)

    # Logo école dans sidebar
    if has_logo_ecole:
        try:
            logo_para = add_cell_paragraph(sidebar, 20)
            logo_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            img = decode_base64_image(data.logos.logo_ecole)
            logo_para.add_run().add_picture(img, height=Cm(2))
        except Exception as e:
            print(f"Erreur logo école: {e}")

    # Année dans sidebar
    year_para = add_cell_paragraph(sidebar, 20)
    year_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = year_para.add_run(data.annee_scolaire or "[Année]")
    run.font.size = Pt(12)
    run.font.color.rgb = RGBColor(255, 255, 255)
    run.bold = True

    # Logo entreprise dans sidebar
    if has_logo_entreprise:
        try:
            logo_para2 = add_cell_paragraph(sidebar, 0)
            logo_para2.alignment = WD_ALIGN_PARAGRAPH.CENTER
            img = decode_base64_image(data.logos.logo_entreprise)
            logo_para2.add_run().add_picture(img, height=Cm(2))
        except Exception as e:
            print(f"Erreur logo entreprise: {e}")

    # ═══ CONTENU PRINCIPAL ═══
    content_para = content.paragraphs[0]
    content_para.paragraph_format.first_line_indent = Cm(0)
    content_para.paragraph_format.space_before = Pt(20)

    # RAPPORT DE STAGE
    title1 = add_cell_paragraph(content, 2)
    run = title1.add_run("RAPPORT DE STAGE")
    run.bold = True
    run.font.size = Pt(24)
    run.font.color.rgb = hex_to_rgb(data.style.title1_color)

    # Formation
    formation = add_cell_paragraph(content, 8)
    run = formation.add_run(data.formation or "[Formation]")
    run.font.size = Pt(11)
    run.italic = True

    # Ligne
    line = add_cell_paragraph(content, 10)
    run = line.add_run("─" * 25)
    run.font.color.rgb = hex_to_rgb(data.style.title1_color)

    # Image centrale si présente
    if has_image_centrale:
        try:
            img_para = add_cell_paragraph(content, 10)
            img = decode_base64_image(data.logos.image_centrale)
            img_para.add_run().add_picture(img, width=Cm(9), height=Cm(3.5))
        except Exception as e:
            print(f"Erreur image centrale: {e}")

    # Étudiant
    student_label = add_cell_paragraph(content, 2)
    run = student_label.add_run("RÉALISÉ PAR")
    run.font.size = Pt(8)
    run.font.color.rgb = RGBColor(120, 120, 120)

    student = add_cell_paragraph(content, 2)
    run = student.add_run(f"{data.prenom or '[Prénom]'} {data.nom or '[NOM]'}")
    run.bold = True
    run.font.size = Pt(14)

    school = add_cell_paragraph(content, 12)
    run = school.add_run(data.ecole or "[Établissement]")
    run.font.size = Pt(10)

    # Entreprise
    ent_label = add_cell_paragraph(content, 2)
    run = ent_label.add_run("ENTREPRISE D'ACCUEIL")
    run.font.size = Pt(8)
    run.font.color.rgb = RGBColor(120, 120, 120)

    ent = add_cell_paragraph(content, 2)
    run = ent.add_run(data.entreprise_nom or "[Entreprise]")
    run.bold = True
    run.font.size = Pt(13)
    run.font.color.rgb = hex_to_rgb(data.style.title1_color)

    if data.entreprise_ville:
        ville = add_cell_paragraph(content, 12)
        run = ville.add_run(data.entreprise_ville)
        run.font.size = Pt(10)

    # Dates
    dates_label = add_cell_paragraph(content, 2)
    run = dates_label.add_run("PÉRIODE")
    run.font.size = Pt(8)
    run.font.color.rgb = RGBColor(120, 120, 120)

    dates = add_cell_paragraph(content, 12)
    run = dates.add_run(f"{date_debut_fr} → {date_fin_fr} ({duree})")
    run.font.size = Pt(10)

    # Tuteurs
    tut_label = add_cell_paragraph(content, 2)
    run = tut_label.add_run("ENCADREMENT")
    run.font.size = Pt(8)
    run.font.color.rgb = RGBColor(120, 120, 120)

    tut1 = add_cell_paragraph(content, 2)
    r1 = tut1.add_run("Entreprise : ")
    r1.font.size = Pt(9)
    run = tut1.add_run(data.tuteur_nom or "[Nom]")
    run.font.size = Pt(9)
    run.bold = True

    if data.tuteur_academique_nom:
        tut2 = add_cell_paragraph(content, 0)
        r2 = tut2.add_run("Académique : ")
        r2.font.size = Pt(9)
        run = tut2.add_run(data.tuteur_academique_nom)
        run.font.size = Pt(9)
        run.bold = True


def generate_report(data) -> io.BytesIO:
    """Génère le rapport de stage."""
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

    # ════════════════════════════════════════════════════════════════
    #                         PAGE DE GARDE
    # ════════════════════════════════════════════════════════════════
    if data.include_cover:
        # Sélection du modèle de page de garde
        cover_model = getattr(data, 'cover_model', 'classique')

        if cover_model == 'moderne':
            generate_cover_moderne(doc, data, date_debut_fr, date_fin_fr, duree)
        elif cover_model == 'corporate':
            generate_cover_corporate(doc, data, date_debut_fr, date_fin_fr, duree)
        else:
            generate_cover_classique(doc, data, date_debut_fr, date_fin_fr, duree)

        doc.add_page_break()

    # Configurer header et footer pour les pages suivantes
    setup_header_with_logos(section, data)
    setup_footer_with_page_number(section, data)

    # ════════════════════════════════════════════════════════════════
    #                      TABLE DES MATIÈRES
    # ════════════════════════════════════════════════════════════════
    if data.include_toc:
        toc_heading = doc.add_heading("TABLE DES MATIÈRES", level=1)
        toc_heading.alignment = WD_ALIGN_PARAGRAPH.CENTER

        doc.add_paragraph()
        create_toc(doc, data)

        doc.add_page_break()

    # ════════════════════════════════════════════════════════════════
    #                        REMERCIEMENTS
    # ════════════════════════════════════════════════════════════════
    if data.include_thanks:
        thanks_heading = doc.add_heading("REMERCIEMENTS", level=1)
        thanks_heading.alignment = WD_ALIGN_PARAGRAPH.CENTER

        p1 = doc.add_paragraph()
        p1.add_run(f"Je tiens à remercier {data.entreprise_nom or '[Entreprise]'} pour m'avoir accueilli durant ce stage.")

        if data.tuteur_nom:
            p2 = doc.add_paragraph()
            p2.add_run(f"Je remercie particulièrement {data.tuteur_nom}")
            if data.tuteur_poste:
                p2.add_run(f", {data.tuteur_poste},")
            p2.add_run(" pour son encadrement tout au long de ce stage.")

        if data.tuteur_academique_nom:
            p3 = doc.add_paragraph()
            p3.add_run(f"Je remercie également {data.tuteur_academique_nom}")
            if data.tuteur_academique_poste:
                p3.add_run(f", {data.tuteur_academique_poste},")
            p3.add_run(" pour son suivi académique.")

        p4 = doc.add_paragraph()
        run = p4.add_run("[Compléter les remerciements...]")
        run.italic = True
        run.font.color.rgb = RGBColor(128, 128, 128)

        doc.add_page_break()

    # ════════════════════════════════════════════════════════════════
    #                       RÉSUMÉ / ABSTRACT
    # ════════════════════════════════════════════════════════════════
    if data.include_abstract:
        resume_heading = doc.add_heading("RÉSUMÉ", level=1)
        resume_heading.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p = doc.add_paragraph()
        run = p.add_run("[Résumé du rapport en français...]")
        run.italic = True
        run.font.color.rgb = RGBColor(128, 128, 128)

        doc.add_paragraph()

        doc.add_heading("Abstract", level=2)
        p2 = doc.add_paragraph()
        run2 = p2.add_run("[English abstract...]")
        run2.italic = True
        run2.font.color.rgb = RGBColor(128, 128, 128)

        doc.add_page_break()

    # ════════════════════════════════════════════════════════════════
    #                          CHAPITRES
    # ════════════════════════════════════════════════════════════════
    for chapter_idx, chapter in enumerate(data.chapters, 1):
        # Numérotation automatique des chapitres
        chapter_num = chapter_idx
        doc.add_heading(f"{chapter_num}. {chapter.title}", level=1)

        p = doc.add_paragraph()
        run = p.add_run(get_chapter_hint(chapter.title))
        run.italic = True
        run.font.color.rgb = RGBColor(128, 128, 128)

        for sub_idx, sub in enumerate(chapter.children, 1):
            # Numérotation automatique des sous-chapitres
            doc.add_heading(f"{chapter_num}.{sub_idx}. {sub.title}", level=2)

            p_sub = doc.add_paragraph()
            run_sub = p_sub.add_run("[Contenu à rédiger...]")
            run_sub.italic = True
            run_sub.font.color.rgb = RGBColor(128, 128, 128)

            # Gérer les sous-sous-chapitres si présents
            if hasattr(sub, 'children') and sub.children:
                for subsub_idx, subsub in enumerate(sub.children, 1):
                    doc.add_heading(f"{chapter_num}.{sub_idx}.{subsub_idx}. {subsub.title}", level=3)

                    p_subsub = doc.add_paragraph()
                    run_subsub = p_subsub.add_run("[Contenu à rédiger...]")
                    run_subsub.italic = True
                    run_subsub.font.color.rgb = RGBColor(128, 128, 128)

        doc.add_page_break()

    # ════════════════════════════════════════════════════════════════
    #                           ANNEXES
    # ════════════════════════════════════════════════════════════════
    if data.include_annexes:
        annexes_heading = doc.add_heading("ANNEXES", level=1)
        annexes_heading.alignment = WD_ALIGN_PARAGRAPH.CENTER

        doc.add_heading("Annexe A - [Titre]", level=2)
        p = doc.add_paragraph()
        run = p.add_run("[Contenu de l'annexe...]")
        run.italic = True
        run.font.color.rgb = RGBColor(128, 128, 128)

    # Sauvegarder
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer


def get_chapter_hint(title: str) -> str:
    """Retourne une indication selon le type de chapitre."""
    t = title.lower()
    if "introduction" in t:
        return "[Présenter le contexte, les objectifs et le plan du rapport...]"
    elif "entreprise" in t or "présentation" in t:
        return "[Présenter l'entreprise, son histoire, ses activités, son organisation...]"
    elif "mission" in t:
        return "[Décrire les missions confiées et leurs objectifs...]"
    elif "travail" in t or "réalis" in t:
        return "[Détailler le travail effectué, les méthodes et outils utilisés...]"
    elif "bilan" in t:
        return "[Analyser les résultats, difficultés et compétences acquises...]"
    elif "conclusion" in t:
        return "[Synthétiser les apports du stage et les perspectives...]"
    return "[Contenu à rédiger...]"
