import streamlit as st
import pandas as pd
from io import BytesIO
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
import unicodedata, re
from xml.sax.saxutils import escape

st.set_page_config(page_title="Tracabilité XLS → PDF", layout="centered")
st.title("📘 Générateur de fiches d’évaluation")

uploaded_file = st.file_uploader("Choisir le fichier Excel (.xlsx)", type=["xlsx"])

# --- Nettoyage de texte ---
def nettoyer_texte_visible(txt):
    if pd.isna(txt):
        return ""
    s = str(txt)
    s = ''.join(ch for ch in s if ch.isprintable())
    s = re.sub(r"[_•■]+", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return escape(s)

# --- Nettoyage noms de colonnes ---
def normaliser_colname(name):
    s = str(name)
    s = ''.join(ch for ch in unicodedata.normalize("NFKD", s) if not unicodedata.combining(ch))
    s = s.lower().strip()
    s = re.sub(r"\s+", "_", s)
    return s

# --- Déterminer couleur selon valeur ---
def coloriser_valeur_html(val):
    if pd.isna(val):
        return ""
    s = str(val).strip().lower()
    if s in ["fait"]:
        color = "#00B050"
    elif s in ["a"]:
        color = "#007A33"
    elif s in ["en cours", "encours"]:
        color = "#FFD700"
    elif s in ["eca", "e.c.a.", "e.c.a"]:
        color = "#ED7D31"
    elif s in ["n.e."]:
        color = "#808080"
    elif s in ["n.a."]:
        color = "#C00000"
    else:
        color = "#000000"
    return f"<font color='{color}'><b>{escape(str(val))}</b></font>"

# --- Application principale ---
if uploaded_file:
    df = pd.read_excel(uploaded_file)
    df.columns = [normaliser_colname(c) for c in df.columns]

    # Détection automatique
    stagiaire_col = next((c for c in df.columns if "stagiaire" in c), None)
    date_col = next((c for c in df.columns if "date" in c), None)
    prenom_col = next((c for c in df.columns if "prenom" in c), None)
    nom_col = next((c for c in df.columns if "nom" in c and "prenom" not in c), None)

    if not stagiaire_col:
        st.error("⚠️ Colonne stagiaire introuvable.")
        st.stop()

    if prenom_col and nom_col:
        df["formateur"] = df[prenom_col].astype(str) + " " + df[nom_col].astype(str)
    else:
        df["formateur"] = ""

    # Groupes de colonnes
    app_non_eval_cols = [c for c in df.columns if "app_non" in c or "non_soumis" in c]
    app_eval_cols = [c for c in df.columns if "app_eval" in c or "app_evalue" in c]
    axes_cols = [c for c in df.columns if "axe" in c or "progression" in c]
    ancrage_cols = [c for c in df.columns if "ancrage" in c or "ancr" in c]
    app_prop_cols = [c for c in df.columns if "app_qui" in c or "propose" in c]

    # Styles
    styles = getSampleStyleSheet()
    titre_style = ParagraphStyle("Titre", parent=styles["Heading1"], alignment=1, textColor="#007A33")
    section_style = ParagraphStyle("Section", parent=styles["Heading3"], textColor="#003366")
    item_style = ParagraphStyle("Item", parent=styles["Normal"], fontSize=10, leading=13, spaceAfter=3, leftIndent=15)

    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, leftMargin=50, rightMargin=50, topMargin=50, bottomMargin=50)
    elements = []

    for stagiaire, group in df.groupby(stagiaire_col):
        ligne = group.iloc[0]

        # En-tête
        elements.append(Paragraph("Fiche d’évaluation", titre_style))
        elements.append(Spacer(1, 6))
        elements.append(Paragraph(f"<b>Stagiaire :</b> {nettoyer_texte_visible(stagiaire)}", item_style))
        if date_col:
            elements.append(Paragraph(f"<b>Date :</b> {nettoyer_texte_visible(ligne.get(date_col, ''))}", item_style))
        elements.append(Paragraph(f"<b>Formateur :</b> {nettoyer_texte_visible(ligne.get('formateur', ''))}", item_style))
        elements.append(Spacer(1, 10))

        # Fonction d’ajout de section
        def add_section(title, cols):
            elements.append(Paragraph(f"<b>{escape(title)}</b>", section_style))
            added = False
            for c in cols:
                v = ligne.get(c, "")
                if pd.notna(v) and str(v).strip():
                    nom_app = escape(c.replace("_", " ").title())
                    val_col = coloriser_valeur_html(v)
                    elements.append(Paragraph(f"• {nom_app} : {val_col}", item_style))
                    added = True
            if not added:
                elements.append(Paragraph("Aucun item", item_style))
            elements.append(Spacer(1, 6))

        # Sections
        add_section("APP non soumis à évaluation", app_non_eval_cols)
        add_section("APP évalués", app_eval_cols)
        add_section("Axes de progression", axes_cols)
        add_section("Points d’ancrage", ancrage_cols)
        add_section("APP qui pourraient être proposés", app_prop_cols)
        elements.append(PageBreak())

    doc.build(elements)
    buffer.seek(0)

    st.success("✅ PDF généré avec succès avec code couleur.")
    st.download_button("⬇️ Télécharger le PDF", data=buffer.getvalue(),
                       file_name="fiches_stagiaires.pdf", mime="application/pdf")
