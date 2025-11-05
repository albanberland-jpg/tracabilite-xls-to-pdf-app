# app.py
import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime
import unicodedata, re

# ReportLab
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape, portrait
from reportlab.lib.units import cm
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Table, TableStyle, Spacer, PageBreak, KeepTogether
)
from xml.sax.saxutils import escape

# ---------------- Streamlit UI ----------------
st.set_page_config(page_title="Synthèse évaluations → PDF", layout="centered")
st.title("🗂️ Synthèse hebdomadaire des évaluations - Stage SAV 2")
st.caption("Importe un .xlsx (export de ton application). Le PDF généré contient une page par stagiaire avec toutes ses évaluations regroupées par date.")

uploaded_file = st.file_uploader("Importer un fichier Excel (.xlsx)", type=["xlsx"])

# ---------------- Utilities ----------------
def normalise_colname(name: str) -> str:
    s = str(name)
    s = ''.join(ch for ch in unicodedata.normalize("NFKD", s) if not unicodedata.combining(ch))
    s = s.lower().strip()
    s = re.sub(r"\s+", " ", s)
    s = s.replace(" ", "_")
    s = re.sub(r"[^a-z0-9_/()'’.-]", "", s)
    return s

def clean_display_text(s) -> str:
    """Keep readable chars, remove emojis and weird control chars to avoid squares."""
    if pd.isna(s):
        return ""
    t = str(s)
    # normalize accents
    t = ''.join(ch for ch in unicodedata.normalize("NFKD", t) if not unicodedata.combining(ch))
    # remove non-printable
    t = ''.join(ch for ch in t if ch.isprintable())
    # remove common square/emoji remnants
    t = re.sub(r"[_\u25a0\uFFFD•■]", " ", t)
    t = re.sub(r"\s+", " ", t).strip()
    return t

def normalize_value_key(v) -> str:
    """Normalized key for color decisions: uppercase, remove dots/spaces."""
    if pd.isna(v):
        return ""
    s = str(v).upper()
    s = s.replace(".", "").replace(" ", "")
    s = ''.join(ch for ch in unicodedata.normalize("NFKD", s) if not unicodedata.combining(ch))
    s = re.sub(r"[^A-Z0-9]", "", s)
    return s

def color_for_value(v):
    k = normalize_value_key(v)
    if k in ("FAIT",):
        return colors.HexColor("#00B050")  # vert clair
    if k in ("A",):
        return colors.HexColor("#007A33")  # vert foncé
    if k in ("ENCOURS", "ENCOURS"):  # catch variants
        return colors.HexColor("#FFD700")  # jaune
    if k in ("ECA", "ECA"): 
        return colors.HexColor("#ED7D31")  # orange
    if k in ("NA", "N.A"):
        return colors.HexColor("#C00000")  # rouge
    if k in ("NE",):
        return colors.HexColor("#808080")  # gris
    return colors.black

def detect_eval_columns(df):
    """
    Return list of candidate evaluation columns (names as in df.columns).
    We'll include columns that look like APP / MSP / evaluation keywords,
    excluding known metadata columns.
    """
    meta_keywords = ["prenom", "nom", "e-mail", "email", "organisation", "departement",
                     "date", "temps", "taux", "score", "tentative", "reussite", "nbre",
                     "participan", "stagiaire_evalue", "evaluation_de_la_journee", "jcmsplugin"]
    eval_cols = []
    for c in df.columns:
        nc = normalise_colname(c)
        if any(m in nc for m in meta_keywords):
            continue
        # include columns that contain these substrings
        if ("app_" in nc) or ("evaluation" in nc) or ("msp" in nc) or ("test" in nc) or ("axe" in nc) or ("ancrage" in nc):
            eval_cols.append(c)
        else:
            # include if many rows have recognizable evaluation words (fait, en cours, eca, na, ne, a)
            sample_vals = df[c].dropna().astype(str).head(10).tolist()
            if any(re.search(r"\b(fait|en cours|e\.c\.a|eca|na|ne|a)\b", v, re.IGNORECASE) for v in sample_vals):
                eval_cols.append(c)
    # ensure unique and stable order as in original
    eval_cols = [c for c in df.columns if c in eval_cols]
    return eval_cols

# ---------------- PDF generation ----------------
def build_pdf_bytes(df, stagiaire_col_name, prenom_col, nom_col, date_col):
    # Normalize date column to datetime if possible
    if date_col and date_col in df.columns:
        try:
            df["_parsed_date"] = pd.to_datetime(df[date_col], dayfirst=True, errors='coerce')
        except Exception:
            df["_parsed_date"] = pd.to_datetime(df[date_col], errors='coerce')
    else:
        df["_parsed_date"] = pd.NaT

    # Sort stagiaires by name (normalize)
    df_sorted = df.sort_values(by=[stagiaire_col_name])

    eval_columns = detect_eval_columns(df_sorted)

    # Build styles
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle("Title", parent=styles["Heading1"], alignment=1, fontSize=14, textColor=colors.HexColor("#0B5394"))
    subtitle_style = ParagraphStyle("Sub", parent=styles["Normal"], alignment=1, fontSize=9, textColor=colors.grey)
    name_style = ParagraphStyle("Name", parent=styles["Heading2"], alignment=1, fontSize=12, textColor=colors.HexColor("#0B5394"))
    section_band_style = ParagraphStyle("Band", parent=styles["Normal"], alignment=0, fontSize=10, textColor=colors.white, leading=12)
    cell_style = ParagraphStyle("Cell", parent=styles["Normal"], fontSize=9, leading=11)
    legend_style = ParagraphStyle("Legend", parent=styles["Normal"], fontSize=8)

    # We will build pages per stagiaire
    buffer = BytesIO()

    # We will choose portrait or landscape per stagiaire depending on estimated width/density
    # For two-column tables this rarely requires landscape, but if many rows > threshold, choose landscape
    elements_all = []

    grouped = df_sorted.groupby(stagiaire_col_name, sort=True)

    export_dt = datetime.now().strftime("%d/%m/%Y %H:%M")

    for stagiaire, group in grouped:
        # header
        header_parts = []
        header_parts.append(Paragraph("Synthèse hebdomadaire des évaluations - Stage SAV 2", title_style))
        header_parts.append(Paragraph(f"Exporté le : {export_dt}", subtitle_style))
        header_parts.append(Spacer(1, 8))
        header_parts.append(Paragraph(clean_display_text(stagiaire), name_style))
        header_parts.append(Spacer(1, 8))
        # we'll collect content blocks per date
        # Determine unique dates (use parsed date date part if available, else string)
        group =
