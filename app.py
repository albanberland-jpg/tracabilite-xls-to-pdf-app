import streamlit as st
import pandas as pd
from io import BytesIO
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib import colors
import re

st.title("📘 Générateur de fiches d'évaluation")

uploaded_file = st.file_uploader("Choisir le fichier Excel", type=["xlsx"])

# --- Fonction de coloration conditionnelle ---
def coloriser_valeur(val):
    if not isinstance(val, str):
        return str(val)
    val = val.strip().upper()
    couleurs = {
        "FAIT": "#007A33",       # vert foncé
        "A": "#00B050",          # vert clair
        "EN COURS": "#FFD700",   # jaune
        "ECA": "#ED7D31",        # orange
        "NE": "#808080",         # gris
        "NA": "#C00000",         # rouge
    }
    for mot, couleur in couleurs.items():
        if mot in val:
            return f"<font color='{couleur}'><b>{val}</b></font>"
    return val

# --- Fonction de nettoyage du texte ---
def nettoyer_texte(texte):
    if not isinstance(texte, str):
        return ""
    texte = re.sub(r"[■•_🚤🌊⛱🐟🔵\-]+", " ", texte)
    return texte.strip().capitalize()

if uploaded_file:
    df = pd.read_excel(uploaded_file)

    st.write("🔍 Colonnes importées :", list(df.columns))

    # Identifier les colonnes principales (en ignorant les majuscules/minuscules)
    def trouver_colonne(nom_recherche):
        for c in df.columns:
            if nom_recherche.lower() in c.lower():
                return c
        return None

    prenom_col = trouver_colonne("prenom")
    nom_col = trouver_colonne("nom")
    stagiaire_col = trouver_colonne("stagiaire_evalue")
    date_col = trouver_colonne("date")

    if prenom_col and nom_col:
        df["formateur"] = df[prenom_col].astype(str) + " " + df[nom_col].astype(str)
    else:
        st.warning("⚠️ Colonnes 'prenom' et/ou 'nom' introuvables — le champ 'formateur' sera laissé vide.")
        df["formateur"] = ""

    groupes_stagiaires = df.groupby(stagiaire_col)

    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, leftMargin=50, rightMargin=50, topMargin=40, bottomMargin=40)

    styles = getSampleStyleSheet()
    titre_style = ParagraphStyle("Titre", parent=styles["Heading1"], textColor=colors.HexColor("#003366"), spaceAfter=10)
    sous_titre_style = ParagraphStyle("SousTitre", parent=styles["Heading2"], textColor=colors.HexColor("#003366"), spaceAfter=8)
    texte_style = ParagraphStyle("Texte", parent=styles["Normal"], fontSize=11, leading=14, spaceAfter=6)

    elements = []
    elements.append(Paragraph("Test <font color='#FF0000'><b>rouge</b></font>", texte_style))
    elements.append(Spacer(1, 12))

    for stagiaire, data_stagiaire in groupes_stagiaires:
        elements.append(Paragraph("■ Fiche d’évaluation", titre_style))
        elements.append(Spacer(1, 10))
        elements.append(Paragraph(f"<b>Stagiaire évalué :</b> {stagiaire}", texte_style))
        if date_col in data_stagiaire:
            elements.append(Paragraph(f"<b>Évaluation du :</b> {data_stagiaire[date_col].iloc[0]}", texte_style))
        elements.append(Paragraph(f"<b>Formateur :</b> {data_stagiaire['formateur'].iloc[0]}", texte_style))
        elements.append(Spacer(1, 10))

        # --- Sections principales ---
        sections = {
            "APP non soumis à évaluation": "app_non_soumis_a_evaluation",
            "APP évalués": "app_evalues",
            "Axes de progression": "axes_de_progression",
            "Points d'ancrage": "point_d",
            "APP qui pourraient être proposés": "app_qui_pourrait",
        }

        for titre, mot_clef in sections.items():
            cols = [c for c in df.columns if mot_clef in c.lower()]
            if not cols:
                continue

            elements.append(Paragraph(f"■ {titre}", sous_titre_style))
            for col in cols:
                valeur = str(data_stagiaire[col].iloc[0]).strip()
                if valeur and valeur not in ["nan", ""]:
                    # Nettoyer l’intitulé
                    nom_app = nettoyer_texte(col.split("/")[-1])
                    # Coloriser la valeur
                    valeur_coloree = coloriser_valeur(valeur)
                    elements.append(Paragraph(f"• {nom_app} : {valeur_coloree}", texte_style))
            elements.append(Spacer(1, 10))

        elements.append(Spacer(1, 20))

    doc.build(elements)
    st.success("✅ PDF généré avec succès !")
    st.download_button("📄 Télécharger le PDF", data=buffer.getvalue(), file_name="fiches_evaluations.pdf")
