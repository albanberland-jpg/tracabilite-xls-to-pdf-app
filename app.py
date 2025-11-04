import streamlit as st
import pandas as pd
from io import BytesIO
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER
import re

st.title("📘 Générateur de fiches d’évaluation stagiaires")

uploaded_file = st.file_uploader("Choisissez le fichier Excel", type=["xlsx"])

# --- Fonction utilitaire pour nettoyer les intitulés ---
def nettoyer_texte(txt):
    if not isinstance(txt, str):
        return ""
    txt = re.sub(r"[^A-Za-zÀ-ÿ0-9'’() .:/-]", " ", txt)
    txt = re.sub(r"_+", " ", txt)
    txt = re.sub(r"\s+", " ", txt)
    return txt.strip().capitalize()

# --- Fonction pour coloriser les résultats d’évaluation ---
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
        "NA": "#C00000"          # rouge
    }
    couleur = couleurs.get(val, "#000000")
    return f"<font color='{couleur}'><b>{val}</b></font>"

if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)
    colonnes = [str(c).strip().lower().replace(" ", "_") for c in df.columns]
    df.columns = colonnes

    st.write("🔍 Colonnes importées :", colonnes)

    # --- Identifier les colonnes importantes ---
    prenom_col = next((c for c in df.columns if "prenom" in c), None)
    nom_col = next((c for c in df.columns if "nom" in c and "prenom" not in c), None)
    stagiaire_col = next((c for c in df.columns if "stagiaire" in c), None)
    date_col = next((c for c in df.columns if "date" in c), None)

    if not stagiaire_col:
        st.error("Colonne du stagiaire introuvable.")
    else:
        # Ajout d'une colonne "formateur"
        if prenom_col and nom_col:
            df["formateur"] = df[prenom_col].astype(str) + " " + df[nom_col].astype(str)
        else:
            df["formateur"] = ""

        # --- Définir les groupes de colonnes ---
        app_non_eval_cols = [c for c in df.columns if c.startswith("app_non_soumis_a_evaluation")]
        app_eval_cols = [c for c in df.columns if c.startswith("app_evalues")]
        progression_col = next((c for c in df.columns if "axes_de_progression" in c), None)
        ancrage_col = next((c for c in df.columns if "ancrage" in c), None)
        app_propose_col = next((c for c in df.columns if "app_qui_pourrait" in c), None)

        # --- Préparer le PDF ---
        buffer = BytesIO()
        pdf = SimpleDocTemplate(buffer, pagesize=A4)
        styles = getSampleStyleSheet()
        titre_style = ParagraphStyle(
            name="Titre",
            parent=styles["Heading1"],
            alignment=TA_CENTER,
            spaceAfter=10,
            textColor="#007A33"
        )
        section_style = ParagraphStyle(
            name="Section",
            parent=styles["Heading2"],
            textColor="#004C99",
            spaceBefore=8,
            spaceAfter=6
        )
        contenu_style = styles["Normal"]

        elements = []

        groupes = df.groupby(stagiaire_col)

        for stagiaire, data_stagiaire in groupes:
            ligne = data_stagiaire.iloc[0]

            elements.append(Paragraph(f"Fiche d’évaluation", titre_style))
            elements.append(Spacer(1, 10))
            elements.append(Paragraph(f"<b>Stagiaire :</b> {nettoyer_texte(stagiaire)}", contenu_style))
            elements.append(Paragraph(f"<b>Date :</b> {ligne.get(date_col, '')}", contenu_style))
            elements.append(Paragraph(f"<b>Formateur :</b> {nettoyer_texte(ligne.get('formateur', ''))}", contenu_style))
            elements.append(Spacer(1, 10))

            # Section : APP non soumis à évaluation
            if app_non_eval_cols:
                elements.append(Paragraph("🟦 APP non soumis à évaluation", section_style))
                for c in app_non_eval_cols:
                    nom_app = nettoyer_texte(c.split("/", 1)[-1])
                    val = coloriser_valeur(ligne[c])
                    elements.append(Paragraph(f"• {nom_app} : {val}", contenu_style))
                elements.append(Spacer(1, 10))

            # Section : APP évalués
            if app_eval_cols:
                elements.append(Paragraph("🟩 APP évalués", section_style))
                for c in app_eval_cols:
                    nom_app = nettoyer_texte(c.split("/", 1)[-1])
                    val = coloriser_valeur(ligne[c])
                    elements.append(Paragraph(f"• {nom_app} : {val}", contenu_style))
                elements.append(Spacer(1, 10))

            # Axes de progression
            if progression_col:
                val = nettoyer_texte(ligne.get(progression_col, ""))
                elements.append(Paragraph("📈 Axes de progression", section_style))
                elements.append(Paragraph(val, contenu_style))
                elements.append(Spacer(1, 10))

            # Points d’ancrage
            if ancrage_col:
                val = nettoyer_texte(ligne.get(ancrage_col, ""))
                elements.append(Paragraph("📌 Points d’ancrage", section_style))
                elements.append(Paragraph(val, contenu_style))
                elements.append(Spacer(1, 10))

            # APP proposés
            if app_propose_col:
                val = nettoyer_texte(ligne.get(app_propose_col, ""))
                elements.append(Paragraph("💡 APP qui pourraient être proposés", section_style))
                elements.append(Paragraph(val, contenu_style))

            elements.append(PageBreak())

        pdf.build(elements)
        buffer.seek(0)

        st.download_button(
            label="📄 Télécharger le PDF consolidé",
            data=buffer,
            file_name="fiches_stagiaires.pdf",
            mime="application/pdf"
        )
