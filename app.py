import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.styles import ParagraphStyle

st.set_page_config(page_title="Tracabilité XLS → PDF", layout="centered")

st.title("📘 Générateur de fiches d’évaluation")
st.write("Charge un fichier Excel pour créer un PDF clair et coloré, une fiche par stagiaire.")

uploaded_file = st.file_uploader("📂 Choisir le fichier Excel", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)
    st.success("✅ Fichier importé avec succès.")

    # 🔤 Normaliser les noms de colonnes
    def normaliser(n):
        return (
            str(n)
            .lower()
            .replace("é", "e")
            .replace("è", "e")
            .replace("ê", "e")
            .replace("à", "a")
            .replace("â", "a")
            .replace("ô", "o")
            .replace("ç", "c")
            .replace("ï", "i")
            .replace("î", "i")
            .replace(" ", "_")
        )

    df.columns = [normaliser(c) for c in df.columns]

    # 🔎 Colonnes détectées automatiquement
    stagiaire_col = next((c for c in df.columns if "stagiaire" in c), None)
    date_col = next((c for c in df.columns if "date" in c), None)
    formateur_col = next((c for c in df.columns if "formateur" in c), None)

    app_non_eval_cols = [c for c in df.columns if "app_non_soumis" in c]
    app_eval_cols = [c for c in df.columns if "app_evalue" in c]
    axes_cols = [c for c in df.columns if "axe" in c]
    ancrage_cols = [c for c in df.columns if "ancrage" in c]
    app_prop_cols = [c for c in df.columns if "app_qui_pourrait" in c]

    # 🖋 Styles
    titre_style = ParagraphStyle(
        "Titre",
        fontSize=16,
        leading=20,
        alignment=1,
        textColor=colors.HexColor("#008000"),
        spaceAfter=12,
    )
    section_style = ParagraphStyle(
        "Section",
        fontSize=12,
        textColor=colors.HexColor("#003366"),
        leading=14,
        spaceBefore=8,
        spaceAfter=4,
    )
    texte_style = ParagraphStyle(
        "Texte",
        fontSize=10,
        leading=12,
        textColor=colors.black,
        spaceBefore=2,
        allowHTML=True,
    )

    # 🎨 Couleurs d'évaluation
    def coloriser(val):
        if pd.isna(val): 
            return ""
        val = str(val).strip().upper().replace(".", "")
        couleurs = {
            "FAIT": colors.HexColor("#00B050"),
            "ECA": colors.HexColor("#ED7D31"),
            "A": colors.HexColor("#00B050"),
            "EN COURS": colors.HexColor("#FFD700"),
            "NE": colors.HexColor("#808080"),
            "NA": colors.HexColor("#C00000"),
        }
        c = couleurs.get(val)
        if c:
            return f'<font color="{c.hexval()}"><b>{val}</b></font>'
        return f"<b>{val}</b>"

    # 📄 Création du PDF
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4)
    elements = []

    for stagiaire, data_stagiaire in df.groupby(stagiaire_col):
        ligne = data_stagiaire.iloc[0]

        # --- En-tête ---
        elements.append(Paragraph("Fiche d’évaluation", titre_style))
        elements.append(Spacer(1, 8))
        elements.append(Paragraph(f"<b>Stagiaire :</b> {stagiaire}", texte_style))
        elements.append(Paragraph(f"<b>Date :</b> {ligne.get(date_col, '')}", texte_style))
        elements.append(Paragraph(f"<b>Formateur :</b> {ligne.get(formateur_col, '')}", texte_style))
        elements.append(Spacer(1, 10))

        # --- APP non soumis ---
        if app_non_eval_cols:
            elements.append(Paragraph("APP non soumis à évaluation", section_style))
            for c in app_non_eval_cols:
                nom = c.replace("app_non_soumis_a_evaluation_/_", "").replace("_", " ").capitalize()
                val = ligne.get(c, "")
                if pd.notna(val) and val != "":
                    elements.append(Paragraph(f"• {nom} : {coloriser(val)}", texte_style))
            elements.append(Spacer(1, 8))

        # --- APP évalués ---
        if app_eval_cols:
            elements.append(Paragraph("APP évalués", section_style))
            for c in app_eval_cols:
                nom = c.replace("app_evalues_/_", "").replace("_", " ").capitalize()
                val = ligne.get(c, "")
                if pd.notna(val) and val != "":
                    elements.append(Paragraph(f"• {nom} : {coloriser(val)}", texte_style))
            elements.append(Spacer(1, 8))

        # --- Axes de progression ---
        if axes_cols:
            elements.append(Paragraph("Axes de progression", section_style))
            for c in axes_cols:
                val = ligne.get(c, "")
                if pd.notna(val) and val != "":
                    elements.append(Paragraph(str(val), texte_style))
            elements.append(Spacer(1, 8))

        # --- Points d’ancrage ---
        if ancrage_cols:
            elements.append(Paragraph("Points d’ancrage", section_style))
            for c in ancrage_cols:
                val = ligne.get(c, "")
                if pd.notna(val) and val != "":
                    elements.append(Paragraph(str(val), texte_style))
            elements.append(Spacer(1, 8))

        # --- APP proposés ---
        if app_prop_cols:
            elements.append(Paragraph("APP qui pourraient être proposés", section_style))
            for c in app_prop_cols:
                val = ligne.get(c, "")
                if pd.notna(val) and val != "":
                    elements.append(Paragraph(str(val), texte_style))
            elements.append(Spacer(1, 20))

    # --- Génération du PDF ---
    doc.build(elements)
    buffer.seek(0)

    st.download_button(
        label="📄 Télécharger le PDF des fiches",
        data=buffer,
        file_name="fiches_stagiaires.pdf",
        mime="application/pdf",
    )
