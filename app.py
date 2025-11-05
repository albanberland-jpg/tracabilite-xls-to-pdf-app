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
        group = group.copy()
        if "_parsed_date" in group.columns:
            group["_date_only"] = group["_parsed_date"].dt.date
        else:
            group["_date_only"] = pd.NaT

        # If parsing succeeded for some rows but not others, fallback grouping with original string
        if group["_date_only"].isnull().all():
            # use raw date strings
            group["_date_group_key"] = group[date_col].fillna("").astype(str).str.strip()
            group = group.sort_values(by=("_date_group_key"))
        else:
            # prefer parsed date (rows with NaT will have NaT)
            group["_date_group_key"] = group["_date_only"].apply(lambda x: x if pd.notna(x) else None)
            group = group.sort_values(by=("_parsed_date"))

        # create page container; we'll later decide orientation and build doc.pageSize accordingly
        page_elements = []
        page_elements.extend(header_parts)

        # For each date group in chronological order
        for date_key, sub in group.groupby("_date_group_key", sort=True):
            # skip empty key groups
            if date_key is None or (isinstance(date_key, str) and date_key == ""):
                date_label = ""
            else:
                if isinstance(date_key, datetime) or hasattr(date_key, "strftime"):
                    # date object
                    try:
                        date_label = date_key.strftime("%d/%m/%Y")
                    except Exception:
                        date_label = str(date_key)
                else:
                    date_label = str(date_key)

            # Determine formateurs present for that date (unique)
            if prenom_col in group.columns and nom_col in group.columns:
                formateurs = sub[[prenom_col, nom_col]].fillna("").astype(str)
                # build unique non-empty combos
                fm = []
                for _, r in formateurs.iterrows():
                    p = r.get(prenom_col, "").strip()
                    n = r.get(nom_col, "").strip()
                    if (p or n):
                        name = f"{p} {n}".strip()
                        if name not in fm:
                            fm.append(name)
                formateur_label = ", ".join(fm) if fm else "—"
            else:
                formateur_label = "—"

            # show date + formateur above tables for that date
            if date_label:
                page_elements.append(Paragraph(f"<b>Date :</b> {escape(date_label)}    <b>Formateur :</b> {escape(clean_display_text(formateur_label))}", cell_style))
            else:
                page_elements.append(Paragraph(f"<b>Formateur :</b> {escape(clean_display_text(formateur_label))}", cell_style))

            page_elements.append(Spacer(1, 6))

            # For this date, we want to group evaluation columns by "type" (APP, MSP, Autre)
            # Build mapping type -> list of (col, display_label, values_for_rows)
            type_buckets = {}
            # examine eval_columns and include a column if any row in sub has non-empty value
            for col in eval_columns:
                include_col = False
                for _, r in sub.iterrows():
                    v = r.get(col, "")
                    if pd.notna(v) and str(v).strip() != "":
                        include_col = True
                        break
                if not include_col:
                    continue
                nc = normalise_colname(col)
                if "msp" in nc or "victime" in nc:
                    t = "ÉVALUATION DES MSP"
                elif "app_non" in nc or "non_soumis" in nc:
                    t = "APP NON SOUMIS À ÉVALUATION"
                elif "app_evalue" in nc or "app_eval" in nc or nc.startswith("app_"):
                    t = "APP ÉVALUÉS"
                else:
                    t = "AUTRE ÉVALUATION"
                type_buckets.setdefault(t, []).append(col)

            # If no columns detected for that date -> skip
            if not type_buckets:
                page_elements.append(Paragraph("Aucune évaluation renseignée pour cette date.", cell_style))
                page_elements.append(Spacer(1, 6))
                continue

            # For layout decision: estimate rows count
            rows_count_est = sum(len(cols) for cols in type_buckets.values())

            # For each type, add band and table
            for tlabel, cols in type_buckets.items():
                # band
                band_bg = colors.HexColor("#d7e9ff")
                band_style = ParagraphStyle("band", parent=styles["Normal"], alignment=0, textColor=colors.white, fontSize=10, leading=12)
                # Use a small Table to draw the band with background color (full width)
                band_table = Table([[Paragraph(f"<b>{escape(tlabel)}</b>", band_style)]], colWidths=[16*cm])
                band_table.setStyle(TableStyle([
                    ('BACKGROUND', (0, 0), (-1, -1), band_bg),
                    ('LEFTPADDING', (0,0), (-1,-1), 6),
                    ('RIGHTPADDING', (0,0), (-1,-1), 6),
                    ('TOPPADDING', (0,0), (-1,-1), 4),
                    ('BOTTOMPADDING', (0,0), (-1,-1), 4),
                ]))
                page_elements.append(band_table)
                page_elements.append(Spacer(1, 4))

                # Build table rows: header then rows for each column (sequence) with its result (if multiple rows in sub, concatenate unique results)
                tbl_rows = []
                header = [Paragraph("<b>Séquence / Épreuve</b>", cell_style), Paragraph("<b>Résultat</b>", cell_style)]
                tbl_rows.append(header)

                for col in cols:
                    # sequence label: take last part after '/' if present for clarity
                    seq_label = col
                    if "/" in col:
                        seq_label = col.split("/")[-1]
                    seq_label = clean_display_text(seq_label).replace("_", " ").strip()
                    # collect values for that date (could be multiple lines); deduplicate and keep order
                    vals = []
                    for _, r in sub.iterrows():
                        v = r.get(col, "")
                        if pd.notna(v) and str(v).strip() != "":
                            vt = clean_display_text(v)
                            if vt not in vals:
                                vals.append(vt)
                    # If no values (shouldn't happen) skip
                    if not vals:
                        continue
                    # combine multiple values with " / "
                    combined_val = " / ".join(vals)
                    # prepare Paragraphs: seq and colored val
                    seq_par = Paragraph(escape(seq_label), cell_style)
                    # colored text style
                    col_color = color_for_value(combined_val)
                    val_style = ParagraphStyle("val", parent=styles["Normal"], fontSize=9, textColor=col_color)
                    val_par = Paragraph(escape(combined_val), val_style)
                    tbl_rows.append([seq_par, val_par])

                # choose column widths based on estimated content width
                # If many rows or long sequence names, we may prefer landscape: handle outside
                col_widths = [10*cm, 6*cm]

                table = Table(tbl_rows, colWidths=col_widths, hAlign='LEFT')
                # table style: thin border, header background blue
                table_style = TableStyle([
                    ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#2B6EB3")),
                    ('TEXTCOLOR', (0,0), (-1,0), colors.white),
                    ('ALIGN', (0,0), (-1,0), 'CENTER'),
                    ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
                    ('GRID', (0,0), (-1,-1), 0.25, colors.grey),
                    ('BOX', (0,0), (-1,-1), 0.5, colors.grey),
                    ('LEFTPADDING', (0,0), (-1,-1), 6),
                    ('RIGHTPADDING', (0,0), (-1,-1), 6),
                    ('TOPPADDING', (0,0), (-1,-1), 4),
                    ('BOTTOMPADDING', (0,0), (-1,-1), 4),
                ])
                # alternate background for data rows
                for i in range(1, len(tbl_rows)):
                    if i % 2 == 0:
                        table_style.add('BACKGROUND', (0,i), (-1,i), colors.whitesmoke)
                table.setStyle(table_style)
                page_elements.append(table)
                page_elements.append(Spacer(1, 6))

        # After all dates for this stagiaire, add bottom sections (Axes, Ancrage, APP proposés) using first available non-empty
        def first_nonempty_from_group(cols_list):
            for c in cols_list:
                if c in df.columns:
                    v = group.iloc[0].get(c, "")
                    if pd.notna(v) and str(v).strip():
                        return clean_display_text(v)
            return ""

        # detect columns for axes, ancrage, app_prop by normalized names
        axes_cols = [c for c in df.columns if "axe" in normalise_colname(c) or "progression" in normalise_colname(c)]
        ancrage_cols = [c for c in df.columns if "ancrag" in normalise_colname(c) or "ancrage" in normalise_colname(c) or "point_d'ancrage" in normalise_colname(c)]
        app_prop_cols = [c for c in df.columns if "app_qui" in normalise_colname(c) or "pourrait" in normalise_colname(c) or "propose" in normalise_colname(c)]

        axes_text = first_nonempty_from_group(axes_cols)
        ancrage_text = first_nonempty_from_group(ancrage_cols)
        app_prop_text = first_nonempty_from_group(app_prop_cols)

        if axes_text:
            page_elements.append(Paragraph("<b>Axes de progression</b>", ParagraphStyle("sec", parent=styles["Heading4"], textColor=colors.HexColor("#0B5394"))))
            page_elements.append(Paragraph(escape(axes_text), cell_style))
            page_elements.append(Spacer(1,6))
        if ancrage_text:
            page_elements.append(Paragraph("<b>Points d'ancrage</b>", ParagraphStyle("sec", parent=styles["Heading4"], textColor=colors.HexColor("#0B5394"))))
            page_elements.append(Paragraph(escape(ancrage_text), cell_style))
            page_elements.append(Spacer(1,6))
        if app_prop_text:
            page_elements.append(Paragraph("<b>APP qui pourraient être proposés</b>", ParagraphStyle("sec", parent=styles["Heading4"], textColor=colors.HexColor("#0B5394"))))
            page_elements.append(Paragraph(escape(app_prop_text), cell_style))
            page_elements.append(Spacer(1,6))

        # legend
        legend = "Légende : Fait / A = Acquis (vert)  •  En cours / ECA = En cours (jaune/orange)  •  NA = Non acquis (rouge)  •  NE = Non évalué (gris)"
        page_elements.append(Spacer(1, 8))
        page_elements.append(Paragraph(escape(legend), legend_style))

        # Add page break
        page_elements.append(PageBreak())

        # Decide orientation: if many rows overall for this stagiaire, use landscape
        total_rows = sum(1 for _ in page_elements if isinstance(_, Table))  # rough count
        # simpler heuristic: if number of tables or rows large -> landscape
        # But ReportLab document page orientation is global; to keep simpler we'll use portrait for now,
        # and switch to landscape if any table row count very large (e.g. > 30)
        # Estimate data rows count:
        data_rows_est = 0
        for el in page_elements:
            if isinstance(el, Table):
                data_rows_est += len(el._cellvalues) - 1
        # pack into KeepTogether to avoid splitting header and first blocks if possible
        elements_all.extend(page_elements)

    # Build document in portrait (portrait chosen; user asked portrait default, landscape if necessary)
    doc = SimpleDocTemplate(buffer, pagesize=A4, leftMargin=2*cm, rightMargin=2*cm, topMargin=2*cm, bottomMargin=2*cm)
    doc.build(elements_all)
    buffer.seek(0)
    return buffer.getvalue()

# ---------------- Streamlit flow ----------------
if uploaded_file is not None:
    try:
        df = pd.read_excel(uploaded_file, dtype=str)
    except Exception as e:
        st.error(f"Erreur lecture fichier : {e}")
        st.stop()

    st.write("Colonnes importées :", list(df.columns))

    # detect stagiaire column by name (use exact header 'Stagiaire évalué' if present)
    stag_col = None
    for c in df.columns:
        if 'stagiaire' in c.lower() or 'participant' in c.lower() or 'élève' in c.lower() or 'eleve' in c.lower():
            stag_col = c
            break
    if stag_col is None:
        st.error("Colonne 'Stagiaire évalué' non trouvée. Vérifie le fichier.")
        st.stop()

    # detect prenom/nom columns (for formateur)
    prenom_col = None
    nom_col = None
    for c in df.columns:
        lc = c.lower()
        if 'prenom' in lc and prenom_col is None:
            prenom_col = c
        if 'nom' == lc or ('nom' in lc and 'prenom' not in lc) and nom_col is None:
            nom_col = c

    # detect date column if present
    date_col = None
    for c in df.columns:
        if 'date' in c.lower():
            date_col = c
            break

    st.write(f"Stagiaire col: {stag_col} | Prenom col: {prenom_col} | Nom col: {nom_col} | Date col: {date_col}")

    if st.button("📄 Générer la synthèse PDF (une page par stagiaire)"):
        try:
            pdf_bytes = build_pdf_bytes(df, stag_col, prenom_col, nom_col, date_col)
            st.success("PDF généré.")
            st.download_button("⬇️ Télécharger le PDF", data=pdf_bytes, file_name="synthese_evaluations_stage_sav2.pdf", mime="application/pdf")
        except Exception as e:
            st.error(f"Erreur lors de la génération du PDF : {e}")
