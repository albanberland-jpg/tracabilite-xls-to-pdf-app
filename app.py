import streamlit as st
import pandas as pd
from io import BytesIO
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
import re
import unicodedata

st.set_page_config(page_title="Tracabilité XLS → PDF", layout="centered")
st.title("📘 Générateur de fiches d’évaluation")

uploaded_file = st.file_uploader("Choisir le fichier Excel (.xlsx)", type=["xlsx"])

# ---------------- utilities ----------------
def normaliser_colname(name):
    s = str(name)
    s = ''.join(ch for ch in unicodedata.normalize("NFKD", s) if not unicodedata.combining(ch))
    s = s.lower().strip()
    s = re.sub(r"\s+", "_", s)
    s = re.sub(r"[^a-z0-9_/()'’.-]", "", s)
    return s

def nettoyer_texte_visible(txt):
    """Garde seulement caractères lisibles (lettres accentuées, chiffres, ponctuation basique).
       Retire emojis, carrés noirs, symboles insolites."""
    if pd.isna(txt):
        return ""
    s = str(txt)
    # Normaliser accents
    s = ''.join(ch for ch in unicodedata.normalize("NFKD", s) if not unicodedata.combining(ch))
    # Remove emojis and other non-printable/rare symbols by keeping only a safe whitelist:
    s = re.sub(r"[^A-Za-zÀ-ÖØ-öø-ÿ0-9 ,;:!\?'\(\)\[\]\-\/\.%&°%\"’]", " ", s)
    # collapse spaces
    s = re.sub(r"\s+", " ", s).strip()
    return s

def valeur_cle(val):
    """Renvoie clé normalisée pour comparaison stricte (supprime points/espaces)."""
    if pd.isna(val):
        return ""
    s = str(val).upper()
    s = s.replace(".", "").replace(" ", "")
    # retirer accents pour sécurité
    s = ''.join(ch for ch in unicodedata.normalize("NFKD", s) if not unicodedata.combining(ch))
    return s

# ---------------- coloriser (ordre prioritaire) ----------------
def coloriser_valeur_html(val):
    """Retourne HTML <font> coloré si correspondance exacte après nettoyage."""
    key = valeur_cle(val)
    # ordre important : tester ECA avant A
    mapping = [
        ("FAIT", "#007A33"),     # vert foncé
        ("ENCOURS", "#FFD700"),  # jaune (normalized "ENCOURS")
        ("NE", "#808080"),       # gris
        ("NA", "#C00000"),       # rouge
        ("ECA", "#ED7D31"),      # orange
        ("A", "#00B050"),        # vert clair
    ]
    for k, hexc in mapping:
        if key == k:
            # affiche original (avec nettoyage visuel) but color
            display = nettoyer_texte_visible(val)
            return f"<font color='{hexc}'><b>{display}</b></font>"
    # fallback: just display cleaned text, not colored
    return f"{nettoyer_texte_visible(val)}"

# ---------------- main app ----------------
if uploaded_file:
    df = pd.read_excel(uploaded_file)
    # normalize column names
    df.columns = [normaliser_colname(c) for c in df.columns]
    st.write("🔍 Colonnes importées :", list(df.columns))

    # detect main columns robustly
    stagiaire_col = next((c for c in df.columns if "stagiaire" in c or "participant" in c or "eleve" in c), None)
    date_col = next((c for c in df.columns if "date" in c), None)
    prenom_col = next((c for c in df.columns if "prenom" in c), None)
    nom_col = next((c for c in df.columns if "nom" in c and "prenom" not in c), None)

    if stagiaire_col is None:
        st.error("Colonne stagiaire non trouvée — vérifie l'en-tête du fichier.")
        st.stop()

    # create formateur column
    if prenom_col and nom_col and prenom_col in df.columns and nom_col in df.columns:
        df["formateur"] = df[prenom_col].astype(str).str.strip() + " " + df[nom_col].astype(str).str.strip()
    else:
        df["formateur"] = ""

    # detect app columns (tolerant contains)
    app_non_eval_cols = [c for c in df.columns if "app_non" in c or "non_soumis" in c]
    app_eval_cols = [c for c in df.columns if "app_evalue" in c or "app_evalues" in c or "app_eval" in c]
    axes_cols = [c for c in df.columns if "axe" in c or "axes_de_progression" in c]
    ancrage_cols = [c for c in df.columns if "ancrage" in c or "point_d_ancrage" in c or "point_d'anc" in c]
    app_prop_cols = [c for c in df.columns if "app_qui_pourrait" in c or "app_qui_peut" in c or "propose" in c]

    st.write("✅ Colonnes repérées (exemples) :")
    st.write(" - app_non:", app_non_eval_cols)
    st.write(" - app_eval:", app_eval_cols)
    st.write(" - axes:", axes_cols)
    st.write(" - ancrage:", ancrage_cols)
    st.write(" - app_propose:", app_prop_cols)

    # PDF styles (allowHTML on item style)
    styles = getSampleStyleSheet()
    titre_style = ParagraphStyle("Titre", parent=styles["Heading1"], alignment=1, textColor=colors.HexColor("#003366"), spaceAfter=12)
    section_style = ParagraphStyle("Section", parent=styles["Heading3"], textColor=colors.HexColor("#003366"), spaceBefore=8, spaceAfter=6)
    item_style = ParagraphStyle("Item", parent=styles["Normal"], allowHTML=True, spaceAfter=4, leftIndent=8)

    # prepare PDF
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, leftMargin=50, rightMargin=50, topMargin=50, bottomMargin=50)
    elements = []

    # iterate per stagiaire and per row (keeps all evaluations)
    for stagiaire, group in df.groupby(stagiaire_col):
        for idx, row in group.iterrows():
            elements.append(Paragraph("Fiche d’évaluation", titre_style))
            elements.append(Paragraph(f"<b>Stagiaire :</b> {nettoyer_texte_visible(stagiaire)}", item_style))
            if date_col:
                elements.append(Paragraph(f"<b>Date :</b> {nettoyer_texte_visible(row.get(date_col,''))}", item_style))
            elements.append(Paragraph(f"<b>Formateur :</b> {nettoyer_texte_visible(row.get('formateur',''))}", item_style))
            elements.append(Spacer(1, 8))

            # APP non soumis
            if app_non_eval_cols:
                elements.append(Paragraph("APP non soumis à évaluation", section_style))
                any_item = False
                for c in app_non_eval_cols:
                    v = row.get(c, "")
                    if pd.notna(v) and str(v).strip() not in ["", "nan"]:
                        nom_app = nettoyer_texte_visible(c.split("/")[-1])
                        elements.append(Paragraph(f"• {nom_app} : {coloriser_valeur_html(v)}", item_style))
                        any_item = True
                if not any_item:
                    elements.append(Paragraph("Aucun item", item_style))
                elements.append(Spacer(1, 6))

            # APP évalués
            if app_eval_cols:
                elements.append(Paragraph("APP évalués", section_style))
                any_item = False
                for c in app_eval_cols:
                    v = row.get(c, "")
                    if pd.notna(v) and str(v).strip() not in ["", "nan"]:
                        nom_app = nettoyer_texte_visible(c.split("/")[-1])
                        elements.append(Paragraph(f"• {nom_app} : {coloriser_valeur_html(v)}", item_style))
                        any_item = True
                if not any_item:
                    elements.append(Paragraph("Aucun item", item_style))
                elements.append(Spacer(1, 6))

            # Axes de progression
            if axes_cols:
                elements.append(Paragraph("Axes de progression", section_style))
                for c in axes_cols:
                    v = row.get(c, "")
                    if pd.notna(v) and str(v).strip() not in ["", "nan"]:
                        elements.append(Paragraph(nettoyer_texte_visible(v), item_style))
                elements.append(Spacer(1, 6))

            # Points d'ancrage
            if ancrage_cols:
                elements.append(Paragraph("Points d'ancrage", section_style))
                for c in ancrage_cols:
                    v = row.get(c, "")
                    if pd.notna(v) and str(v).strip() not in ["", "nan"]:
                        elements.append(Paragraph(nettoyer_texte_visible(v), item_style))
                elements.append(Spacer(1, 6))

            # APP proposés
            if app_prop_cols:
                elements.append(Paragraph("APP qui pourraient être proposés", section_style))
                for c in app_prop_cols:
                    v = row.get(c, "")
                    if pd.notna(v) and str(v).strip() not in ["", "nan"]:
                        elements.append(Paragraph(nettoyer_texte_visible(v), item_style))
                elements.append(Spacer(1, 8))

    doc.build(elements)
    buffer.seek(0)

    st.success("PDF prêt.")
    st.download_button("Télécharger le PDF", data=buffer.getvalue(), file_name="fiches_evaluations.pdf", mime="application/pdf")
