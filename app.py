import streamlit as st
import pandas as pd
from io import BytesIO
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from datetime import datetime

st.set_page_config(page_title="Tracabilité XLS → PDF", layout="centered")

st.title("📘 Générateur de fiches d’évaluation")
st.write("Charge un fichier XLSX pour générer automatiquement les fiches stagiaires en PDF.")

uploaded_file = st.file_uploader("Choisir le fichier Excel", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)
    st.success("✅ Fichier importé avec succès !")

    # Normaliser les noms de colonnes (suppression accents et espaces)
    def normaliser_nom(n):
        return (
            str(n)
            .strip()
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

    df.columns = [normaliser_nom(c) for c in df.columns]

    # Détection automatique des colonnes principales
    prenom_col = next((c for c in df.columns if "prenom" in c), None)
    nom_col = next((c for c in df.columns if "nom" in c), None)
    stagiaire_col = next((c for c in df.columns if "stagiaire" in c), None)
    date_col = next((c for c in df.columns if "date" in c), None)

    if not stagiaire_col:
        st.error("❌ Colonne contenant les noms des stagiaires non trouvée.")
        st.stop()

    # Fusion prénom + nom → formateur
    if prenom_col and nom_col:
        df["formateur"] = df[prenom_col].astype(str) + " " + df[nom_col].astype(str)
    else:
        df["formateur"] = "Non spécifié"

    # --- Styles pour le PDF ---
    titre_style = ParagraphStyle(
        "Titre",
        fontSize=16,
        leading=20,
        alignment=1,
        textColor=colors.green,
        spaceAfter=12,
        spaceBefore=12,
    )

    section_style = ParagraphStyle(
        "Section",
        fontSize=12,
        textColor=colors.darkblue,
        leading=14,
        spaceBefore=10,
        spaceAfter=4,
    )

    item_style = ParagraphStyle(
        "Item",
        fontSize=10,
        leading=12,
        textColor=colors.black,
        spaceBefore=2,
        allowHTML=True,  # ⚠️ essentiel pour les couleurs
    )

    # --- Fonction de colorisation des valeurs ---
    def coloriser_valeur(val):
        if pd.isna(val):
            return ""
        s = str(val).strip().upper().replace(".", "").replace(" ", "")
        couleurs = {
            "FAIT": "#00B050",       # vert
            "ENCOURS": "#FFD700",    # jaune
            "NE": "#808080",         # gris
            "NA": "#C00000",         # rouge
            "ECA": "#ED7D31",        # orange
            "A": "#00B050",          # vert
        }
        couleur = couleurs.get(s)
        if couleur:
            return f"<font color='{couleur}'><b>{val}</b></font>"
        return f"<b>{val}</b>"

    # --- Détection des catégories ---
    app_non_eval_cols = [c for c in df.columns if "app_non_soumis" in c]
    app_eval_cols = [c for c in df.columns if "app_evalue" in c]
    axes_cols = [c for c in df.columns if "axe" in c]
    ancrage_cols = [c for c in df.columns if "ancrage" in c]
    app_prop_cols = [c for c in df.columns if "app_qui_pourrait" in c]

    # --- Génération PDF ---
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4)
    elements = []

    for stagiaire, data_stagiaire in df.groupby(stagiaire_col):
        ligne = data_stagiaire.iloc[0]

        date_eval = ligne.get(date_col, "")
        formateur = ligne.get("formateur", "")

        # --- Titre principal ---
        elements.append(Paragraph("Fiche d’évaluation", titre_style))
        elements.append(Spacer(1, 10))
        elements.append(Paragraph(f"<b>Stagiaire :</b> {stagiaire}", item_style))
        elements.append(Paragraph(f"<b>Date :</b> {date_eval}", item_style))
        elements.append(Paragraph(f"<b>Formateur :</b> {formateur}", item_style))
        elements.append(Spacer(1, 10))

        # --- Section : APP non soumis à évaluation ---
        if app_non_eval_cols:
            elements.append(Paragraph("APP non soumis à évaluation", section_style))
            for c in app_non_eval_cols:
                nom_app = c.replace("app_non_soumis_a_evaluation_/_", "").replace("_", " ").capitalize()
                v = ligne.get(c, "")
                if pd.notna(v) and str(v).strip() != "":
                    elements.append(Paragraph(f"• {nom_app} : {coloriser_valeur(v)}", item_style))
            elements.append(Spacer(1, 10))

        # --- Section : APP évalués ---
        if app_eval_cols:
            elements.append(Paragraph("APP évalués", section_style))
            for c in app_eval_cols:
                nom_app = c.replace("app_evalues_/_", "").replace("_", " ").capitalize()
                v = ligne.get(c, "")
                if pd.notna(v) and str(v).strip() != "":
                    elements.append(Paragraph(f"• {nom_app} : {coloriser_valeur(v)}", item_style))
            elements.append(Spacer(1, 10))

        # --- Section : Axes de progression ---
        if axes_cols:
            elements.append(Paragraph("Axes de progression", section_style))
            for c in axes_cols:
                v = ligne.get(c, "")
                if pd.notna(v) and str(v).strip() != "":
                    elements.append(Paragraph(str(v), item_style))
            elements.append(Spacer(1, 10))

        # --- Section : Points d’ancrage ---
        if ancrage_cols:
            elements.append(Paragraph("Points d’ancrage", section_style))
            for c in ancrage_cols:
                v = ligne.get(c, "")
                if pd.notna(v) and str(v).strip() != "":
                    elements.append(Paragraph(str(v), item_style))
            elements.append(Spacer(1, 10))

        # --- Section : APP proposés ---
        if app_prop_cols:
            elements.append(Paragraph("APP qui pourraient être proposés", section_style))
            for c in app_prop_cols:
                v = ligne.get(c, "")
                if pd.notna(v) and str(v).strip() != "":
                    elements.append(Paragraph(str(v), item_style))
            elements.append(Spacer(1, 20))

        # Saut de page entre stagiaires
        elements.append(Spacer(1, 40))
        elements.append(Paragraph("<br/><br/>", item_style))

    # --- Génération finale du PDF ---
    doc.build(elements)
    buffer.seek(0)

    st.download_button(
        label="📄 Télécharger le PDF des fiches",
        data=buffer,
        file_name="fiches_stagiaires.pdf",
        mime="application/pdf",
    )
