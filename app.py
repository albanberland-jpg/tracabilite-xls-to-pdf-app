import streamlit as st
import pandas as pd
from io import BytesIO
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, PageBreak
from reportlab.lib.units import cm
from xml.sax.saxutils import escape
from datetime import datetime

# --- Configuration de la page Streamlit ---
st.set_page_config(page_title="Tracabilité XLS → PDF", layout="centered")
st.title("📘 Générateur de fiches d’évaluation")

st.write("Chargez un fichier Excel pour générer un **PDF unique** contenant une fiche par stagiaire.")

uploaded_file = st.file_uploader("Importer un fichier Excel (.xlsx)", type=["xlsx"])

# --- Fonction de nettoyage texte ---
def nettoyer_texte_visible(texte):
    if pd.isna(texte):
        return ""
    return str(texte).replace("\xa0", " ").replace("\n", " ").strip()

# --- Fonction pour ajouter une section ---
def add_section(elements, title, cols, ligne, item_style):
    section_style = ParagraphStyle(
        name="Section",
        fontSize=11,
        leading=14,
        spaceBefore=10,
        textColor=colors.HexColor("#1F4E79"),
        parent=item_style
    )

    elements.append(Paragraph(f"<b>{escape(title)}</b>", section_style))
    added = False

    for c in cols:
        v = ligne.get(c, "")
        if pd.notna(v) and str(v).strip():
            nom_app = escape(c.replace("_", " ").title())
            val_text = str(v).strip()

            # --- Couleur selon la valeur ---
            couleur = colors.black
            val_lower = val_text.lower()
            if val_lower == "fait":
                couleur = colors.HexColor("#00B050")  # Vert clair
            elif val_lower == "a":
                couleur = colors.HexColor("#007A33")  # Vert foncé
            elif val_lower in ["en cours", "encours"]:
                couleur = colors.HexColor("#FFD700")  # Jaune
            elif val_lower in ["eca", "e.c.a", "e.c.a."]:
                couleur = colors.HexColor("#ED7D31")  # Orange
            elif val_lower == "ne":
                couleur = colors.HexColor("#808080")  # Gris
            elif val_lower == "na":
                couleur = colors.HexColor("#C00000")  # Rouge

            # --- Style de la valeur ---
            valeur_style = ParagraphStyle(
                name="valeur_style",
                parent=item_style,
                textColor=couleur
            )

            # --- Ajout du libellé + valeur ---
            elements.append(Paragraph(f"• {nom_app} :", item_style))
            elements.append(Paragraph(val_text, valeur_style))
            added = True

    if not added:
        elements.append(Paragraph("Aucun item", item_style))

    elements.append(Spacer(1, 6))

# --- Génération du PDF ---
if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)

    # Recherche automatique des colonnes
    stagiaire_col = next((c for c in df.columns if "stagiaire" in c.lower()), None)
    prenom_col = next((c for c in df.columns if "prenom" in c.lower()), None)
    nom_col = next((c for c in df.columns if "nom" in c.lower()), None)
    date_col = next((c for c in df.columns if "date" in c.lower()), None)

    if not stagiaire_col:
        st.error("⚠️ Colonne 'stagiaire' introuvable dans le fichier.")
    else:
        buffer = BytesIO()
        doc = SimpleDocTemplate(buffer, pagesize=A4,
                                rightMargin=2*cm, leftMargin=2*cm,
                                topMargin=2*cm, bottomMargin=2*cm)

        styles = getSampleStyleSheet()
        item_style = ParagraphStyle(
            name="item",
            parent=styles["Normal"],
            fontSize=10,
            leading=14
        )

        elements = []

        # --- Boucle par stagiaire ---
        groupes_stagiaires = df.groupby(stagiaire_col)
        for stagiaire, groupe in groupes_stagiaires:
            ligne = groupe.iloc[0]

            # En-tête
            titre_style = ParagraphStyle(
                name="Titre",
                parent=styles["Heading1"],
                alignment=1,
                textColor=colors.HexColor("#1F7A1F"),
                fontSize=14
            )

            elements.append(Paragraph("Fiche d’évaluation", titre_style))
            elements.append(Spacer(1, 6))

            date_eval = ligne.get(date_col, "")
            if isinstance(date_eval, datetime):
                date_eval = date_eval.strftime("%d/%m/%Y %H:%M")

            prenom = ligne.get(prenom_col, "")
            nom = ligne.get(nom_col, "")

            formateur = f"{prenom} {nom}".strip() or "—"

            elements.append(Paragraph(f"<b>Stagiaire :</b> {escape(str(stagiaire))}", item_style))
            elements.append(Paragraph(f"<b>Date :</b> {escape(str(date_eval))}", item_style))
            elements.append(Paragraph(f"<b>Formateur :</b> {escape(formateur)}", item_style))
            elements.append(Spacer(1, 12))

            # Détection des groupes de colonnes
            app_non_eval_cols = [c for c in df.columns if "non_evalue" in c.lower()]
            app_eval_cols = [c for c in df.columns if "evalue" in c.lower()]
            axes_cols = [c for c in df.columns if "axe" in c.lower()]
            ancrage_cols = [c for c in df.columns if "ancrage" in c.lower()]
            app_prop_cols = [c for c in df.columns if "propose" in c.lower()]

            # Ajout des sections
            if app_non_eval_cols:
                add_section(elements, "APP non soumis à évaluation", app_non_eval_cols, ligne, item_style)
            if app_eval_cols:
                add_section(elements, "APP évalués", app_eval_cols, ligne, item_style)
            if axes_cols:
                add_section(elements, "Axes de progression", axes_cols, ligne, item_style)
            if ancrage_cols:
                add_section(elements, "Points d’ancrage", ancrage_cols, ligne, item_style)
            if app_prop_cols:
                add_section(elements, "APP qui pourraient être proposés", app_prop_cols, ligne, item_style)

            elements.append(PageBreak())

        # --- Création du PDF ---
        doc.build(elements)

        st.success("✅ Fichier PDF généré avec succès !")
        st.download_button(
            label="📄 Télécharger le PDF",
            data=buffer.getvalue(),
            file_name="fiches_stagiaires.pdf",
            mime="application/pdf"
        )
