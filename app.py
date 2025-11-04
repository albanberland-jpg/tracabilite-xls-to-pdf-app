import streamlit as st
import pandas as pd
from io import BytesIO
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, PageBreak
from datetime import datetime
from xml.sax.saxutils import escape

st.set_page_config(page_title="Tracabilité XLS → PDF", layout="centered")
st.title("📘 Générateur de fiches d’évaluation")

uploaded_file = st.file_uploader("Importer un fichier Excel (.xlsx)", type=["xlsx"])

def nettoyer_texte_visible(texte):
    if pd.isna(texte):
        return ""
    return str(texte).replace("\xa0", " ").strip()

# Fonction pour convertir la valeur en texte coloré
def coloriser_texte(valeur):
    val = str(valeur).strip().lower()
    couleurs = {
        "fait": "#00B050",       # Vert clair
        "a": "#007A33",          # Vert foncé
        "en cours": "#FFD700",   # Jaune
        "eca": "#ED7D31",        # Orange
        "e.c.a": "#ED7D31",
        "ne": "#808080",         # Gris
        "na": "#C00000"          # Rouge
    }
    couleur = couleurs.get(val, "#000000")
    return f'<font color="{couleur}"><b>{escape(str(valeur))}</b></font>'

def add_section(elements, titre, colonnes, ligne, style_normal):
    titre_style = ParagraphStyle(
        name="TitreSection",
        parent=style_normal,
        fontSize=11,
        textColor=colors.HexColor("#1F4E79"),
        spaceBefore=10,
        spaceAfter=6
    )
    elements.append(Paragraph(f"<b>{escape(titre)}</b>", titre_style))
    
    ajoute = False
    for c in colonnes:
        valeur = ligne.get(c, "")
        if pd.notna(valeur) and str(valeur).strip():
            texte = f"{escape(c.replace('_', ' ').title())} : {coloriser_texte(valeur)}"
            elements.append(Paragraph(texte, style_normal))
            ajoute = True
    if not ajoute:
        elements.append(Paragraph("Aucun élément.", style_normal))
    elements.append(Spacer(1, 8))

if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)

    stagiaire_col = next((c for c in df.columns if "stagiaire" in c.lower()), None)
    prenom_col = next((c for c in df.columns if "prenom" in c.lower()), None)
    nom_col = next((c for c in df.columns if "nom" in c.lower()), None)
    date_col = next((c for c in df.columns if "date" in c.lower()), None)

    if not stagiaire_col:
        st.error("⚠️ Colonne 'stagiaire' introuvable.")
    else:
        buffer = BytesIO()
        doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=50, leftMargin=50, topMargin=40, bottomMargin=40)

        styles = getSampleStyleSheet()
        style_normal = ParagraphStyle("normal", parent=styles["Normal"], fontSize=10, leading=14)

        elements = []

        for stagiaire, groupe in df.groupby(stagiaire_col):
            ligne = groupe.iloc[0]

            titre_style = ParagraphStyle("titre", parent=styles["Heading1"], alignment=1, textColor=colors.HexColor("#007A33"), fontSize=14)
            elements.append(Paragraph("Fiche d’évaluation", titre_style))
            elements.append(Spacer(1, 12))

            date_eval = ligne.get(date_col, "")
            if isinstance(date_eval, datetime):
                date_eval = date_eval.strftime("%d/%m/%Y %H:%M")

            formateur = f"{ligne.get(prenom_col, '')} {ligne.get(nom_col, '')}".strip()

            elements.append(Paragraph(f"<b>Stagiaire :</b> {escape(str(stagiaire))}", style_normal))
            elements.append(Paragraph(f"<b>Date :</b> {escape(str(date_eval))}", style_normal))
            elements.append(Paragraph(f"<b>Formateur :</b> {escape(formateur)}", style_normal))
            elements.append(Spacer(1, 10))

            app_non_eval_cols = [c for c in df.columns if "non_evalue" in c.lower()]
            app_eval_cols = [c for c in df.columns if "evalue" in c.lower()]
            axes_cols = [c for c in df.columns if "axe" in c.lower()]
            ancrage_cols = [c for c in df.columns if "ancrage" in c.lower()]
            app_prop_cols = [c for c in df.columns if "propose" in c.lower()]

            if app_non_eval_cols:
                add_section(elements, "APP non soumis à évaluation", app_non_eval_cols, ligne, style_normal)
            if app_eval_cols:
                add_section(elements, "APP évalués", app_eval_cols, ligne, style_normal)
            if axes_cols:
                add_section(elements, "Axes de progression", axes_cols, ligne, style_normal)
            if ancrage_cols:
                add_section(elements, "Points d’ancrage", ancrage_cols, ligne, style_normal)
            if app_prop_cols:
                add_section(elements, "APP qui pourraient être proposés", app_prop_cols, ligne, style_normal)

            elements.append(PageBreak())

        doc.build(elements)

        st.success("✅ PDF généré avec succès !")
        st.download_button("📄 Télécharger le PDF", buffer.getvalue(), "fiches_stagiaires.pdf", mime="application/pdf")
