# app.py (copier-coller complet)
import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime
import unicodedata, re

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Table, TableStyle, Spacer, PageBreak
)
from xml.sax.saxutils import escape

st.set_page_config(page_title="Synthèse évaluations → PDF", layout="centered")
st.title("🗂️ Synthèse hebdomadaire des évaluations - Stage SAV 2")
st.caption("Importe un .xlsx (export de ton application). Le PDF final contient une fiche par stagiaire (toutes ses dates).")

# ---------- Helpers ----------
def clean_text(x):
    """Nettoyage lisible : supprime caractères non imprimables et trims."""
    if pd.isna(x):
        return ""
    s = str(x)
    s = ''.join(ch for ch in unicodedata.normalize("NFKD", s) if not unicodedata.combining(ch))
    s = ''.join(ch for ch in s if ch.isprintable())
    s = s.replace("\xa0", " ")
    s = re.sub(r"[_\u25a0\uFFFD•■]+", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s

def normalize_key(v):
    if pd.isna(v):
        return ""
    s = str(v).upper()
    s = s.replace(".", "").replace(" ", "")
    s = ''.join(ch for ch in unicodedata.normalize("NFKD", s) if not unicodedata.combining(ch))
    s = re.sub(r"[^A-Z0-9]", "", s)
    return s

def color_for(v):
    k = normalize_key(v)
    if k in ("FAIT",):
        return colors.HexColor("#00B050")
    if k in ("A",):
        return colors.HexColor("#007A33")
    if k in ("ENCOURS","ENCOURS"):
        return colors.HexColor("#FFD700")
    if k in ("ECA","ECA"):
        return colors.HexColor("#ED7D31")
    if k in ("NA","N A","N.A"):
        return colors.HexColor("#C00000")
    if k in ("NE",):
        return colors.HexColor("#808080")
    return colors.black

def detect_eval_columns(df):
    """Détecte automatiquement les colonnes d'évaluation à tester (tolérant)."""
    meta = {"prenom","nom","e-mail","email","organisation","departement","date","temps","taux","score","tentative","reussite","nbre","stagiaire","participant","jcmsplugin"}
    eval_cols = []
    for c in df.columns:
        nc = c.lower()
        if any(m in nc for m in meta):
            continue
        # heuristique : colonnes contenant 'app', 'evaluation', 'msp', 'test', 'axe', 'ancrage'
        if any(k in nc for k in ("app","evaluation","msp","test","axe","ancrag","ancrage","victime")):
            eval_cols.append(c)
        else:
            # sample check for evaluation keywords
            sample = df[c].dropna().astype(str).head(10).tolist()
            if any(re.search(r"\b(fait|en cours|e\.c\.a|eca|na|ne|a)\b", str(v), re.IGNORECASE) for v in sample):
                eval_cols.append(c)
    # preserve original order
    return [c for c in df.columns if c in eval_cols]

# ---------- PDF builder ----------
def build_pdf(df, stagiaire_col, prenom_col, nom_col, date_col):
    # parse date
    if date_col and date_col in df.columns:
        df["_parsed_date"] = pd.to_datetime(df[date_col], dayfirst=True, errors='coerce')
    else:
        df["_parsed_date"] = pd.NaT

    # tri
    df_sorted = df.sort_values(by=[stagiaire_col, "_parsed_date"], na_position='last')

    eval_cols = detect_eval_columns(df_sorted)

    styles = getSampleStyleSheet()
    title_style = ParagraphStyle("Title", parent=styles["Heading1"], alignment=1, fontSize=14, textColor=colors.HexColor("#0B5394"))
    small_center = ParagraphStyle("SmallCenter", parent=styles["Normal"], alignment=1, fontSize=9, textColor=colors.grey)
    name_style = ParagraphStyle("Name", parent=styles["Heading2"], alignment=1, fontSize=11, textColor=colors.HexColor("#0B5394"))
    date_form_style = ParagraphStyle("DateForm", parent=styles["Heading2"], alignment=1, fontSize=11, textColor=colors.black)  # same size as name
    band_text_style = ParagraphStyle("BandText", parent=styles["Normal"], alignment=0, fontSize=10, textColor=colors.white)
    cell_style = ParagraphStyle("Cell", parent=styles["Normal"], fontSize=9)
    legend_style = ParagraphStyle("Legend", parent=styles["Normal"], fontSize=8, textColor=colors.black)

    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, leftMargin=2*cm, rightMargin=2*cm, topMargin=2*cm, bottomMargin=2*cm)

    elements = []

    # regrouper par stagiaire (ordre alphabétique sur le libellé nettoyé)
    df_sorted["__stag_clean"] = df_sorted[stagiaire_col].apply(lambda x: clean_text(x).lower() if pd.notna(x) else "")
    for stagiaire_key, group in df_sorted.groupby("__stag_clean", sort=True):
        # use first non-null original label for display
        display_stag = ""
        for val in group[stagiaire_col].tolist():
            if pd.notna(val) and clean_text(val):
                display_stag = clean_text(val)
                break
        # header
        elements.append(Paragraph("Synthèse hebdomadaire des évaluations - Stage SAV 2", title_style))
        elements.append(Spacer(1, 4))
        elements.append(Paragraph(display_stag, name_style))
        elements.append(Spacer(1, 6))

        # grouper par date chronologique (use parsed date if available; else raw string)
        if "_parsed_date" in group.columns and group["_parsed_date"].notna().any():
            group = group.sort_values(by="_parsed_date", na_position='last')
            group["_date_key"] = group["_parsed_date"].dt.date
            date_groups = group.groupby("_date_key", sort=True)
        else:
            group["_date_key"] = group[date_col].fillna("").astype(str)
            date_groups = group.groupby("_date_key", sort=True)

        for date_key, sub in date_groups:
            # date label
            if pd.isna(date_key) or date_key == "":
                date_label = ""
            else:
                if isinstance(date_key, (pd.Timestamp, datetime)):
                    try:
                        date_label = date_key.strftime("%d/%m/%Y")
                    except Exception:
                        date_label = str(date_key)
                else:
                    date_label = str(date_key)

            # formateur(s) for that date: dedupe Prénom+Nom combos from rows in sub
            formateurs = []
            if prenom_col and nom_col and prenom_col in sub.columns and nom_col in sub.columns:
                for _, r in sub.iterrows():
                    p = clean_text(r.get(prenom_col, ""))
                    n = clean_text(r.get(nom_col, ""))
                    full = (p + " " + n).strip()
                    if full and full not in formateurs:
                        formateurs.append(full)
            formateur_label = ", ".join(formateurs) if formateurs else "—"

            # show date and formateur in same font size as name
            date_form_text = f"Date : {date_label}" if date_label else ""
            if formateur_label and formateur_label != "—":
                if date_form_text:
                    date_form_text = f"{date_form_text}  |  Formateur : {formateur_label}"
                else:
                    date_form_text = f"Formateur : {formateur_label}"
            elements.append(Paragraph(date_form_text, date_form_style))
            elements.append(Spacer(1, 6))

            # build type buckets (APP évalués / APP non soumis / MSP / Autre)
            type_buckets = {}
            for col in eval_cols:
                # include col if any row in sub has non-empty value
                col_has = False
                for _, r in sub.iterrows():
                    v = r.get(col, "")
                    if pd.notna(v) and str(v).strip() not in ("", "nan"):
                        col_has = True
                        break
                if not col_has:
                    continue
                nc = col.lower()
                if "evaluation_des_msp" in nc or "msp" in nc or "victime" in nc:
                    key = "ÉVALUATION DES MSP"
                elif "app_non" in nc or "non_soumis" in nc or "non_soumis_a_evaluation" in nc:
                    key = "APP NON SOUMIS À ÉVALUATION"
                elif "app_evalue" in nc or "app_evalues" in nc or ("app_" in nc and "evalue" in nc):
                    key = "APP ÉVALUÉS"
                else:
                    key = "AUTRES ÉVALUATIONS"
                type_buckets.setdefault(key, []).append(col)

            # for each type, display band and table
            for tlabel, cols in type_buckets.items():
                # band (dark blue bg, white text) - aligned full width
                band_table = Table([[Paragraph(f"<b>{escape(tlabel)}</b>", band_text_style)]], colWidths=[16*cm])
                band_table.setStyle(TableStyle([
                    ('BACKGROUND', (0,0), (-1,-1), colors.HexColor("#053b74")),  # bleu soutenu
                    ('TEXTCOLOR', (0,0), (-1,-1), colors.white),
                    ('LEFTPADDING', (0,0), (-1,-1), 6),
                    ('RIGHTPADDING', (0,0), (-1,-1), 6),
                    ('TOPPADDING', (0,0), (-1,-1), 4),
                    ('BOTTOMPADDING', (0,0), (-1,-1), 4),
                ]))
                elements.append(band_table)
                elements.append(Spacer(1, 4))

                # build table rows: header then each sequence (col) with combined values (if multiple rows)
                rows = []
                header = [Paragraph("<b>Séquence / Épreuve</b>", cell_style := ParagraphStyle("cell", parent=getSampleStyleSheet()["Normal"], fontSize=9)),
                          Paragraph("<b>Résultat</b>", ParagraphStyle("cellr", parent=getSampleStyleSheet()["Normal"], fontSize=9, alignment=0))]
                rows.append(header)

                for col in cols:
                    # sequence label: take part after '/' if present else column cleaned
                    seq = col
                    if "/" in col:
                        seq = col.split("/")[-1]
                    seq = clean_text(seq).replace("_", " ").strip()
                    # gather all values from rows in sub (dedupe preserving order)
                    vals = []
                    for _, r in sub.iterrows():
                        v = r.get(col, "")
                        if pd.notna(v) and str(v).strip() not in ("", "nan"):
                            vt = clean_text(v)
                            if vt not in vals:
                                vals.append(vt)
                    if not vals:
                        continue
                    combined = " / ".join(vals)
                    seq_par = Paragraph(escape(seq), cell_style)
                    # color the combined value by using a ParagraphStyle with textColor
                    col_color = color_for(combined)
                    val_style = ParagraphStyle("valstyle", parent=getSampleStyleSheet()["Normal"], fontSize=9, textColor=col_color)
                    val_par = Paragraph(escape(combined), val_style)
                    rows.append([seq_par, val_par])

                # table widths and header style: invert colors per your request:
                # header background = light blue, header text black
                # (type band already dark blue with white)
                col_widths = [10*cm, 6*cm]
                table = Table(rows, colWidths=col_widths, hAlign='LEFT')
                tbl_style = TableStyle([
                    ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#9BC2E6")),  # bleu clair
                    ('TEXTCOLOR', (0,0), (-1,0), colors.black),
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
                for i in range(1, len(rows)):
                    if i % 2 == 0:
                        tbl_style.add('BACKGROUND', (0,i), (-1,i), colors.whitesmoke)
                table.setStyle(tbl_style)
                elements.append(table)
                elements.append(Spacer(1, 8))

            # after types for this date, show paragraphs (axes, ancrage, app_prop) from first row for that date if present
            def first_nonempty(cols):
                for c in cols:
                    if c in sub.columns:
                        v = sub.iloc[0].get(c, "")
                        if pd.notna(v) and str(v).strip() not in ("", "nan"):
                            return clean_text(v)
                return ""

            axes_cols = [c for c in df.columns if "axe" in c.lower() or "progression" in c.lower()]
            ancrage_cols = [c for c in df.columns if "ancrag" in c.lower() or "ancrage" in c.lower() or "point_d'ancrage" in c.lower()]
            app_prop_cols = [c for c in df.columns if "app_qui" in c.lower() or "pourrait" in c.lower() or "propose" in c.lower()]

            axes_txt = first_nonempty(axes_cols)
            ancrage_txt = first_nonempty(ancrage_cols)
            app_prop_txt = first_nonempty(app_prop_cols)

            if axes_txt:
                elements.append(Paragraph("<b>Axes de progression</b>", ParagraphStyle("sec", parent=getSampleStyleSheet()["Heading4"], textColor=colors.HexColor("#0B5394"))))
                elements.append(Paragraph(escape(axes_txt), cell_style))
                elements.append(Spacer(1,6))
            if ancrage_txt:
                elements.append(Paragraph("<b>Points d'ancrage</b>", ParagraphStyle("sec", parent=getSampleStyleSheet()["Heading4"], textColor=colors.HexColor("#0B5394"))))
                elements.append(Paragraph(escape(ancrage_txt), cell_style))
                elements.append(Spacer(1,6))
            if app_prop_txt:
                elements.append(Paragraph("<b>APP qui pourraient être proposés</b>", ParagraphStyle("sec", parent=getSampleStyleSheet()["Heading4"], textColor=colors.HexColor("#0B5394"))))
                elements.append(Paragraph(escape(app_prop_txt), cell_style))
                elements.append(Spacer(1,6))

            elements.append(Spacer(1, 10))

        # legend per stagiaire
        legend = "Légende : Fait / A = Acquis (vert)  •  En cours / ECA = En cours (jaune/orange)  •  NA = Non acquis (rouge)  •  NE = Non évalué (gris)"
        elements.append(Paragraph(escape(legend), legend_style))
        elements.append(PageBreak())

    # build pdf
    doc.build(elements)
    buffer.seek(0)
    return buffer.getvalue()

# ---------------- Streamlit flow ----------------
uploaded_file = st.file_uploader("Importer un fichier Excel (.xlsx)", type=["xlsx"])
if uploaded_file is not None:
    try:
        df = pd.read_excel(uploaded_file, dtype=str)
    except Exception as e:
        st.error(f"Erreur lecture fichier : {e}")
        st.stop()

    st.write("Colonnes importées :", list(df.columns))

    # detect columns
    stag_col = None
    for c in df.columns:
        if 'stagiaire' in c.lower() or 'participant' in c.lower() or 'evalué' in c.lower() or 'evalue' in c.lower():
            stag_col = c
            break
    prenom_col = None
    nom_col = None
    for c in df.columns:
        lc = c.lower()
        if 'prenom' in lc and prenom_col is None:
            prenom_col = c
        if ('nom' in lc and 'prenom' not in lc) and nom_col is None:
            nom_col = c
    date_col = None
    for c in df.columns:
        if 'date' in c.lower():
            date_col = c
            break

    st.write(f"Détection colonnes → stagiaire: {stag_col}, prenom: {prenom_col}, nom: {nom_col}, date: {date_col}")

    if stag_col is None:
        st.error("Colonne 'Stagiaire' introuvable. Vérifie l'en-tête.")
    else:
        if st.button("📄 Générer la synthèse PDF (une page par stagiaire)"):
            try:
                pdf_bytes = build_pdf(df, stag_col, prenom_col, nom_col, date_col)
                st.success("PDF généré.")
                st.download_button("⬇️ Télécharger le PDF", data=pdf_bytes, file_name="synthese_evaluations_stage_sav2.pdf", mime="application/pdf")
            except Exception as e:
                st.error(f"Erreur lors de la génération : {e}")
