import streamlit as st
import pandas as pd
from io import BytesIO
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
import re
import unicodedata

st.set_page_config(page_title="Tracabilité XLS → PDF", layout="centered")
st.title("📘 Générateur de fiches d’évaluation")

uploaded_file = st.file_uploader("Choisir le fichier Excel (.xlsx)", type=["xlsx"])

# --- Fonctions utilitaires ---
def normaliser_colname(name):
    s = str(name)
    s = ''.join(ch for ch in unicodedata.normalize("NFKD", s) if not unicodedata.combining(ch))
    s = s.lower().strip()
    s = re.sub(r"\s+", "_", s)
    s = re.sub(r"[^a-z0-9_/()'’.-]", "", s)
    return s

def nettoyer_texte_visible(txt):
    if pd.isna(txt):
        return ""
    s = str(txt)
    s = ''.join(ch for ch in unicodedata.normalize("NFKD", s) if not unicodedata.combining(ch))
    s = re.sub(r"[^A-Za-zÀ-ÖØ-öø-ÿ0-9 ,;:!\?'\(\)\[\]\-\/\.%&°%\"’]", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s

def valeur_cle(val):
    if pd.isna(val):
        return ""
    s = str(val).upper()
    s = s.replace(".", "").replace(" ", "")
    s = ''.join(ch for ch in unicodedata.normalize("NFKD", s) if not unicodedata.combining(ch))
    return s

def coloriser_valeur_html(val):
    key = valeur_cle(val)
    mapping = {
        "FAIT": "#00B050",      # vert clair
        "ENCOURS": "#FFD700",   # jaune
        "NE": "#808080",        # gris
        "NA": "#C00000",        # rouge
        "ECA": "#ED7D31",       # orange
        "A": "#007A33"          # vert foncé
    }
    color = mapping.get(key, "#000000")
    txt = nettoyer_texte_visible(val)
    return f"<font color='{color}'><b>{txt}</b></font>"

# --- Application principale ---
if uploaded_file:
    df = pd.read_excel(uploaded_file)
    df.columns = [normaliser_colname(c) for c in df.columns]

    st.write("🔍 Colonnes détectées :", list(df.columns))

    stagiaire_col = next((c for c in df.columns if "stagiaire" in c or "participant" in c or "eleve" in c), None)
    date_col = next((c for c in df.columns if "date" in c), None)
    prenom_col = next((c for c in df.columns if "prenom" in c), None)
    nom_col = next((c for c in df.columns if "nom" in c and "prenom" not in c), None)

    if not stagiaire_col:
        st.error("⚠️ Colonne stagiaire introuvable dans le fichier.")
        st.stop()

    if prenom_col and nom_col:
        df["formateur"] = df[prenom_col].astype(str).str.strip() + " " + df[nom_col].astype(str).str.strip()
    else:
        df["formateur"] = ""

    app_non_eval_cols = [c for c in df.columns if "app_non" in c or "non_soumis" in c]
    app_eval_cols = [c for c in df.columns if "app_evalue" in c or "app_eval" in c]
    axes_cols = [c for c in df.columns if "axe" in c or "progression" in c]
    ancrage_cols = [c for c in df.columns if "ancrage" in c or "ancr" in c]
    app_prop_cols = [c for c in df.columns if "app_qui" in c or "propose" in c]

    styles = getSampleStyleSheet()
    titre_style = ParagraphStyle(
        "Titre", parent=styles["Heading1"], alignment=1,
        textColor=colors.HexColor("#007A33"), spaceAfter=12
    )
    section_style = ParagraphStyle(
        "Section", parent=styles["Heading3"],
        textColor=colors.HexColor("#003366"), spaceBefore=8, spaceAfter=6
    )
    item_style = ParagraphStyle(
        "Item", parent=styles["Normal"],
        fontSize=10, leading=13, spaceAfter=3, leftIndent=15
    )

    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4,
                            leftMargin=50, rightMargin=50, topMargin=50, bottomMargin=50)
    elements = []

    for stagiaire, group in df.groupby(stagiaire_col):
        first_row = group.iloc[0]
        elements.append(Paragraph("Fiche d’évaluation", titre_style))
        elements.append(Paragraph(f"<b>Stagiaire :</b> {nettoyer_texte_visible(stagiaire)}", item_style))
        if date_col:
            elements.append(Paragraph(f"<b>Date :</b> {nettoyer_texte_visible(first_row.get(date_col,''))}", item_style))
        elements.append(Paragraph(f"<b>Formateur :</b> {nettoyer_texte_visible(first_row.get('formateur',''))}", item_style))
        elements.append(Spacer(1, 8))

        def add_section(title, cols):
            elements.append(Paragraph(f"<b>{title}</b>", section_style))
            any_item = False
            for c in cols:
                v = first_row.get(c, "")
                if pd.notna(v) and str(v).strip():
                    nom_app = nettoyer_texte_visible(c.replace("_", " "))
                    val_col = coloriser_valeur_html(v)
                    elements.append(Paragraph(f"- {nom_app} : {val_col}", item_style))
                    any_item = True
            if not any_item:
                elements.append(Paragraph("Aucun item", item_style))
            elements.append(Spacer(1, 6))

        add_section("APP non soumis à évaluation", app_non_eval_cols)
        add_section("APP évalués", app_eval_cols)
        add_section("Axes de progression", axes_cols)
        add_section("Points d’ancrage", ancrage_cols)
        add_section("APP qui pourraient être proposés", app_prop_cols)

        elements.append(PageBreak())

    doc.build(elements)
    buffer.seek(0)

    st.success("✅ PDF généré avec succès.")
    st.download_button("⬇️ Télécharger le PDF", data=buffer.getvalue(),
                       file_name="fiches_stagiaires.pdf", mime="application/pdf")
