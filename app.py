# app.py
import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime
import re
import unicodedata
from xml.sax.saxutils import escape

# ReportLab
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.units import cm
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import (
    SimpleDocTemplate,
    Paragraph,
    Table,
    TableStyle,
    Spacer,
    PageBreak,
)

# ---------- Streamlit UI ----------
st.set_page_config(page_title="Synthèse évaluations → PDF", layout="centered")
st.title("🗂️ Synthèse hebdomadaire des évaluations - Stage SAV 2")

uploaded_file = st.file_uploader("Importer un fichier Excel (.xlsx)", type=["xlsx"])

# ---------- Helpers ----------
def normalise_colname(name: str) -> str:
    """Nettoie / normalise les noms de colonnes pour comparaisons."""
    s = str(name)
    s = ''.join(ch for ch in unicodedata.normalize("NFKD", s) if not unicodedata.combining(ch))
    s = s.lower().strip()
    s = re.sub(r"\s+", " ", s)
    s = s.replace(" ", "_")
    s = re.sub(r"[^a-z0-9_/-]", "", s)
    return s

def clean_label(s: str) -> str:
    """Nettoie un libellé pour affichage lisible (suppr emojis / carrés / underscores)."""
    if pd.isna(s):
        return ""
    t = str(s)
    # Remove non printable and most emojis (keep accents)
    t = ''.join(ch for ch in t if ch.isprintable())
    t = t.replace("_", " ")
    t = re.sub(r"\s+", " ", t).strip()
    # remove stray unicode square characters if present
    t = t.replace("\u25a0", "").replace("\uFFFD", "")
    return t

def normalize_value_key(v: str) -> str:
    """Key used to decide color. returns lowercase without punctuation/spaces."""
    if pd.isna(v):
        return ""
    s = str(v).lower()
    s = s.replace(".", "").replace(" ", "")
    s = re.sub(r"[^a-z0-9]", "", s)
    return s

def color_for_value(v: str):
    """Return reportlab color for a given cell value."""
    k = normalize_value_key(v)
    # Map keys to colors (as hex)
    if k in ("fait",): 
        return colors.HexColor("#00B050")  # vert clair
    if k in ("a",): 
        return colors.HexColor("#007A33")  # vert foncé
    if k in ("encours","encours"): 
        return colors.HexColor("#FFD700")  # jaune
    if k in ("eca","eca"): 
        return colors.HexColor("#ED7D31")  # orange
    if k in ("na","n.a","na"): 
        return colors.HexColor("#C00000")  # rouge
    if k in ("ne",): 
        return colors.HexColor("#808080")  # gris
    # default
    return colors.black

# ---------- Main logic ----------
if uploaded_file is not None:
    try:
        df = pd.read_excel(uploaded_file, dtype=str)  # read everything as str to avoid surprises
    except Exception as e:
        st.error(f"Erreur lors de la lecture du fichier : {e}")
        st.stop()

    # Keep original columns mapping and also normalized names
    orig_columns = list(df.columns)
    norm_map = {normalise_colname(c): c for c in orig_columns}
    norm_cols = list(norm_map.keys())

    st.write("Colonnes importées :", orig_columns)

    # detect core columns
    stagiaire_col = None
    for k in norm_cols:
        if "stagiaire" in k or "participant" in k or "eleve" in k:
            stagiaire_col = norm_map[k]
            break

    date_col = None
    for k in norm_cols:
        if k.startswith("date") or "date" in k:
            date_col = norm_map[k]
            break

    prenom_col = None
    nom_col = None
    for k in norm_cols:
        if "prenom" == k or "prenom" in k:
            prenom_col = norm_map[k]
        if (k == "nom" or k.startswith("nom")) and "prenom" not in k:
            nom_col = norm_map[k]

    if stagiaire_col is None:
        st.error("Impossible de trouver la colonne 'Stagiaire évalué'. Assure-toi qu'elle existe.")
        st.stop()

    # find evaluation-type columns (we only include non-empty columns per trainee later)
    # categorize columns by types for easier processing
    app_non_eval_cols = [norm_map[k] for k in norm_cols if "app_non" in k or "non_soumis" in k or "non_evalue" in k]
    app_eval_cols = [norm_map[k] for k in norm_cols if "app_evalue" in k or "app_eval" in k or k.startswith("app_") and "eval" in k]
    msp_cols = [norm_map[k] for k in norm_cols if "evaluation_des_msp" in k or "evaluation_des_msp" in k]
    axes_cols = [norm_map[k] for k in norm_cols if "axe" in k or "progression" in k]
    ancrage_cols = [norm_map[k] for k in norm_cols if "ancrage" in k or "point_d'ancrage" in k or "point_d'anc" in k]
    app_prop_cols = [norm_map[k] for k in norm_cols if "app_qui" in k or "propose" in k or "qui_pourrait" in k]

    # fallback: if app_eval_cols empty, detect columns that start with "app_evalues_/_" patterns used in your sample
    if not app_eval_cols:
        for c in orig_columns:
            if re.search(r"app[_\s]evalue", c, re.IGNORECASE) or "app_evalues" in normalise_colname(c):
                app_eval_cols.append(c)

    # Prepare PDF generation button
    if st.button("📄 Générer la synthèse PDF (une page par stagiaire)"):
        buffer = BytesIO()
        doc = SimpleDocTemplate(buffer, pagesize=A4,
                                leftMargin=2*cm, rightMargin=2*cm,
                                topMargin=2*cm, bottomMargin=2*cm)

        styles = getSampleStyleSheet()
        style_title = ParagraphStyle("Title", parent=styles["Heading1"], alignment=1, fontSize=16, textColor=colors.HexColor("#0B5394"))
        style_sub = ParagraphStyle("Sub", parent=styles["Normal"], alignment=1, fontSize=9, textColor=colors.grey)
        style_name = ParagraphStyle("Name", parent=styles["Heading2"], alignment=1, fontSize=12, textColor=colors.HexColor("#0B5394"))
        style_table_header = ParagraphStyle("Th", parent=styles["Normal"], alignment=1, fontSize=9, textColor=colors.white, spaceBefore=2, spaceAfter=2)
        style_table_cell = ParagraphStyle("Cell", parent=styles["Normal"], alignment=0, fontSize=9)
        style_legend = ParagraphStyle("Legend", parent=styles["Normal"], fontSize=8, textColor=colors.black)

        elements = []

        # Title (fixed as requested)
        export_dt = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
        # We'll add title on each page top when building per-stagiaire.

        # For each stagiaire produce a page
        grouped = df.groupby(stagiaire_col, dropna=True)

        for stagiaire, group in grouped:
            # header
            elements.append(Paragraph("Synthèse hebdomadaire des évaluations - Stage SAV 2", style_title))
            elements.append(Paragraph(f"Exporté le : {export_dt}", style_sub))
            elements.append(Spacer(1, 12))
            elements.append(Paragraph(clean_label(str(stagiaire)).upper(), style_name))
            elements.append(Spacer(1, 12))

            # Build table rows: header then data rows of evaluations that are not empty
            # Table columns: Date | Séquence | Type | Résultat / Évaluations | Formateur
            table_header = ["Date", "Séquence", "Type", "Résultats / Évaluations", "Formateur"]
            table_data = [table_header]

            # We'll look into the evaluation columns (app_non_eval + app_eval + msp) and create a row for each non-empty cell
            eval_columns = []
            # prefer specific order for readability
            eval_columns.extend(app_non_eval_cols)
            eval_columns.extend(app_eval_cols)
            eval_columns.extend(msp_cols)
            # if still empty, fallback to any column that contains 'app' or 'evaluation' keywords
            if not eval_columns:
                for c in orig_columns:
                    nc = normalise_colname(c)
                    if "app" in nc or "eval" in nc or "msp" in nc:
                        eval_columns.append(c)

            # For each row (group may have multiple rows for same stagiaire), iterate rows
            for _, row in group.iterrows():
                date_value = ""
                if date_col and pd.notna(row.get(date_col, "")):
                    date_value = str(row.get(date_col, "")).strip()

                formateur = ""
                if prenom_col and nom_col:
                    p = row.get(prenom_col, "")
                    n = row.get(nom_col, "")
                    formateur = f"{p} {n}".strip()
                else:
                    formateur = ""

                # iterate columns: if cell not empty -> add row
                for col in eval_columns:
                    val = row.get(col, "")
                    if pd.isna(val) or str(val).strip() == "":
                        continue
                    # sequence label = last part after '/' if exists, else cleaned column name
                    seq = col
                    if "/" in col:
                        seq = col.split("/")[-1]
                    seq = clean_label(seq)
                    # type detection
                    nc = normalise_colname(col)
                    if "app" in nc:
                        typ = "APP"
                    elif "msp" in nc or "victime" in nc:
                        typ = "MSP"
                    else:
                        typ = "Autre"

                    # results cell will be a Paragraph with colored text
                    display_val = clean_label(str(val))
                    # color selection
                    color = color_for_value(display_val)

                    # create Paragraph for value with its color by creating a ParagraphStyle on the fly
                    val_style = ParagraphStyle(name="val_style", parent=styles["Normal"], fontSize=9, textColor=color)
                    seq_par = Paragraph(escape(seq), style_table_cell)
                    typ_par = Paragraph(escape(typ), style_table_cell)
                    val_par = Paragraph(escape(display_val), val_style)
                    form_par = Paragraph(escape(clean_label(formateur)), style_table_cell)

                    table_data.append([date_value, seq_par, typ_par, val_par, form_par])

            # If table has only header (no rows), show a message instead
            if len(table_data) == 1:
                elements.append(Paragraph("Aucune évaluation disponible pour ce stagiaire.", style_table_cell))
            else:
                # Create table; choose column widths
                col_widths = [3.0*cm, 6.0*cm, 2.0*cm, 6.0*cm, 3.0*cm]
                table = Table(table_data, colWidths=col_widths, hAlign='LEFT')

                # Table style: header row background blue, white text; grid lines; alternate row background
                tbl_style = TableStyle([
                    ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#9BC2E6")),  # header bg
                    ('TEXTCOLOR', (0,0), (-1,0), colors.white),
                    ('ALIGN', (0,0), (-1,0), 'CENTER'),
                    ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
                    ('INNERGRID', (0,0), (-1,-1), 0.25, colors.grey),
                    ('BOX', (0,0), (-1,-1), 0.5, colors.grey),
                    ('LEFTPADDING', (0,0), (-1,-1), 6),
                    ('RIGHTPADDING', (0,0), (-1,-1), 6),
                    ('TOPPADDING', (0,0), (-1,-1), 4),
                    ('BOTTOMPADDING', (0,0), (-1,-1), 4),
                ])

                # alternate row background for data rows (skip header)
                for i in range(1, len(table_data)):
                    if i % 2 == 0:
                        bg = colors.whitesmoke
                        tbl_style.add('BACKGROUND', (0,i), (-1,i), bg)

                table.setStyle(tbl_style)
                elements.append(table)

            elements.append(Spacer(1, 12))

            # Add sections below table: Axes de progression, Points d'ancrage, APP qui pourraient être proposés
            def add_text_section(title, cols):
                # find first non-empty in group (we keep first row contents)
                text_val = ""
                for c in cols:
                    v = group.iloc[0].get(c, "")
                    if pd.notna(v) and str(v).strip():
                        text_val = clean_label(v)
                        break
                if text_val:
                    elements.append(Paragraph(f"<b>{escape(title)}</b>", ParagraphStyle("sec", parent=styles["Heading4"], textColor=colors.HexColor("#0B5394"))))
                    elements.append(Paragraph(escape(text_val), style_table_cell))
                    elements.append(Spacer(1,6))

            add_text_section("Axes de progression", axes_cols)
            add_text_section("Points d'ancrage", ancrage_cols)
            add_text_section("APP qui pourraient être proposés", app_prop_cols)

            # Legend
            legend = ("Légende : NE = Non évalué | NA = Non acquis | ECA = En cours d'acquisition | "
                      "A = Acquis | Fait = Acquis")
            elements.append(Spacer(1, 12))
            elements.append(Paragraph(escape(legend), ParagraphStyle("legend", parent=styles["Normal"], fontSize=8, textColor=colors.black)))
            elements.append(PageBreak())

        # build and provide download
        doc.build(elements)
        buffer.seek(0)

        st.success("✅ PDF généré.")
        st.download_button("⬇️ Télécharger le PDF", data=buffer.getvalue(), file_name="synthese_evaluations.pdf", mime="application/pdf")
