import streamlit as st
import pandas as pd
from io import BytesIO
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT
import unicodedata
import re

st.set_page_config(page_title="Fiches d'évaluation", layout="centered")
st.title("📘 Générateur de fiches d’évaluation")

uploaded_file = st.file_uploader("Importer un fichier Excel (.xlsx)", type=["xlsx"])

# ---------- utilitaires ----------
def norm_colname(c):
    """Normalise un nom de colonne : retire accents, espaces -> underscore, minuscules."""
    c = str(c)
    c = ''.join(ch for ch in unicodedata.normalize("NFKD", c) if not unicodedata.combining(ch))
    c = c.strip().lower()
    c = re.sub(r"\s+", "_", c)
    c = re.sub(r"[^a-z0-9_/()'’.-]", "", c)  # conserve slash + underscores utiles
    return c

def nettoyer_intitule(texte):
    """Nettoie et rend lisible un intitulé de colonne / app."""
    if texte is None:
        return ""
    t = str(texte)
    # garder uniquement lettres, chiffres, espaces, apostrophes, parenthèses et / :
    t = re.sub(r"[_\-\s]+", " ", t)            # '_' et '-' -> espace
    t = re.sub(r"[^A-Za-zÀ-ÿ0-9()\s'’:/]", "", t)  # retire emojis et carrés
    t = t.strip()
    # souvent on a "app_evalues_/_🚤_mise_a_l_eau" -> garder après le "/"
    if "/" in t:
        t = t.split("/")[-1].strip()
    # mettre en casse lisible
    return t.capitalize()

def coloriser_valeur(val):
    """Retourne du HTML coloré selon la valeur, avec ordre de priorité correct."""
    if pd.isna(val):
        return ""
    s = str(val).strip().upper().replace(".", "").replace(" ", "")

    # ordre important : ECA avant A, sinon "A" match aussi dans "ECA"
    mapping = [
        ("FAIT", "#00B050"),   # vert
        ("ENCOURS", "#FFD700"),# jaune
        ("NE", "#808080"),     # gris
        ("NA", "#C00000"),     # rouge
        ("ECA", "#ED7D31"),    # orange
        ("A", "#00B050"),      # vert
    ]

    for key, color in mapping:
        if s == key:
            return f"<font color='{color}'><b>{key.title()}</b></font>"

    # fallback si la valeur n’est pas pile correspondante
    return f"<b>{val}</b>"

# ---------- traitement principal ----------
if uploaded_file:
    df = pd.read_excel(uploaded_file)

    # normaliser colonnes
    df.columns = [norm_colname(c) for c in df.columns]
    st.write("🔍 Colonnes importées :", list(df.columns))

    # recherche robustes de colonnes clés
    def find_col_by_keywords(keywords):
        for c in df.columns:
            for k in keywords:
                if k in c:
                    return c
        return None

    prenom_col = find_col_by_keywords(["prenom"])
    nom_col = find_col_by_keywords(["nom"])
    stagiaire_col = find_col_by_keywords(["stagiaire", "participant", "eleve"])
    date_col = find_col_by_keywords(["date", "evaluation_de_la_journee", "date_evaluation"])

    st.write(f"🧾 Détection → prenom: {prenom_col}, nom: {nom_col}, stagiaire: {stagiaire_col}, date: {date_col}")

    if stagiaire_col is None:
        st.error("Impossible de détecter la colonne 'stagiaire'. Vérifie l'en-tête du fichier.")
        st.stop()

    # constructeur colonne formateur robuste
    if prenom_col and nom_col and prenom_col in df.columns and nom_col in df.columns:
        df["formateur"] = df[prenom_col].astype(str).str.strip() + " " + df[nom_col].astype(str).str.strip()
    else:
        df["formateur"] = ""

    # détecter les groupes de colonnes par substring (tolérant)
    app_non_cols = [c for c in df.columns if "app_non" in c or "non_soumis" in c]
    app_eval_cols = [c for c in df.columns if "app_evalue" in c or "app_evalues" in c or "app_eval" in c]
    axes_cols = [c for c in df.columns if "axe" in c or "axes_de_progression" in c]
    ancrage_cols = [c for c in df.columns if "ancrage" in c or "point_d'ancrage" in c]
    app_prop_cols = [c for c in df.columns if "app_qui_pourrait" in c or "propose" in c]

    st.write("✅ Colonnes repérées :")
    st.write(" - app_non:", app_non_cols)
    st.write(" - app_eval:", app_eval_cols)
    st.write(" - axes:", axes_cols)
    st.write(" - ancrage:", ancrage_cols)
    st.write(" - app_propose:", app_prop_cols)

    # prepare PDF
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4,
                            leftMargin=50, rightMargin=50, topMargin=50, bottomMargin=50)
    styles = getSampleStyleSheet()
    titre_style = ParagraphStyle("Titre", parent=styles["Heading1"], alignment=TA_CENTER, textColor="#003366", spaceAfter=12)
    header_style = ParagraphStyle("Header", parent=styles["Normal"], spaceAfter=6)
    section_style = ParagraphStyle("Section", parent=styles["Heading3"], textColor="#004C99", spaceBefore=8, spaceAfter=6)
    item_style = ParagraphStyle("Item", parent=styles["Normal"], leftIndent=12, spaceAfter=4)

    elements = []

    # loop: pour chaque stagiaire -> pour chaque ligne (évaluation)
    groupes = df.groupby(stagiaire_col)
    for stagiaire, data_stagiaire in groupes:
        # pour chaque évaluation (si plusieurs lignes)
        for idx, ligne in data_stagiaire.iterrows():
            elements.append(Paragraph("Fiche d’évaluation", titre_style))

            # entêtes (stagiaire, date, formateur)
            elements.append(Paragraph(f"<b>Stagiaire évalué :</b> {nettoyer_intitule(stagiaire)}", header_style))
            date_val = ligne.get(date_col, "")
            elements.append(Paragraph(f"<b>Évaluation du :</b> {date_val}", header_style))
            elements.append(Paragraph(f"<b>Formateur :</b> {nettoyer_intitule(ligne.get('formateur',''))}", header_style))
            elements.append(Spacer(1, 8))

            # APP non soumis
            if app_non_cols:
                elements.append(Paragraph("APP non soumis à évaluation", section_style))
                any_item = False
                for c in app_non_cols:
                    v = ligne.get(c, "")
                    if pd.notna(v) and str(v).strip() not in ["", "nan"]:
                        nom = nettoyer_intitule(c)
                        val_col = coloriser_valeur(v)
                        elements.append(Paragraph(f"• {nom} : {val_col}", item_style))
                        any_item = True
                if not any_item:
                    elements.append(Paragraph("Aucun item", item_style))
                elements.append(Spacer(1, 6))

            # APP évalués
            if app_eval_cols:
                elements.append(Paragraph("APP évalués", section_style))
                any_item = False
                for c in app_eval_cols:
                    v = ligne.get(c, "")
                    if pd.notna(v) and str(v).strip() not in ["", "nan"]:
                        nom = nettoyer_intitule(c)
                        val_col = coloriser_valeur(v)
                        elements.append(Paragraph(f"• {nom} : {val_col}", item_style))
                        any_item = True
                if not any_item:
                    elements.append(Paragraph("Aucun item", item_style))
                elements.append(Spacer(1, 6))

            # Axes de progression (peut être texte long)
            if axes_cols:
                elements.append(Paragraph("Axes de progression", section_style))
                for c in axes_cols:
                    v = ligne.get(c, "")
                    if pd.notna(v) and str(v).strip() not in ["", "nan"]:
                        elements.append(Paragraph(str(v), item_style))
                elements.append(Spacer(1, 6))

            # Points d'ancrage
            if ancrage_cols:
                elements.append(Paragraph("Points d'ancrage", section_style))
                for c in ancrage_cols:
                    v = ligne.get(c, "")
                    if pd.notna(v) and str(v).strip() not in ["", "nan"]:
                        elements.append(Paragraph(str(v), item_style))
                elements.append(Spacer(1, 6))

            # APP proposés
            if app_prop_cols:
                elements.append(Paragraph("APP qui pourraient être proposés", section_style))
                for c in app_prop_cols:
                    v = ligne.get(c, "")
                    if pd.notna(v) and str(v).strip() not in ["", "nan"]:
                        elements.append(Paragraph(str(v), item_style))
                elements.append(Spacer(1, 6))

            elements.append(PageBreak())

    # build
    doc.build(elements)
    buffer.seek(0)

    st.success("✅ PDF généré (une fiche par évaluation, regroupées par stagiaire).")
    st.download_button("📥 Télécharger le PDF", buffer.getvalue(), file_name="fiches_evaluations.pdf", mime="application/pdf")
