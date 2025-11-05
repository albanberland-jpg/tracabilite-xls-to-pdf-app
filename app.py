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

# --- NOUVEAU : Gère le fond et la couleur de police pour le tableau ---
def get_style_colors(v):
    k = normalize_value_key(v)
    
    # Couleurs de fond (similaire à la synthèse Excel que vous aviez)
    bg_color_map = {
        "FAIT": colors.HexColor("#A9D18E"),    # Vert clair
        "A": colors.HexColor("#70AD47"),       # Vert foncé
        "ENCOURS": colors.HexColor("#FFC000"), # Jaune/Orange
        "ECA": colors.HexColor("#ED7D31"),     # Orange
        "NA": colors.HexColor("#F8CBAD"),      # Rouge très clair
        "NE": colors.HexColor("#D9D9D9"),      # Gris clair
    }
    
    # Couleurs de police : blanc pour le vert foncé ('A'), noir pour le reste
    text_color_map = {
        "A": colors.white,
    }
    
    bg = bg_color_map.get(k, colors.white)
    text = text_color_map.get(k, colors.black)
    
    return bg, text

def color_for_value(v):
    """(Maintenu pour l'ancienne logique de texte simple si elle subsiste)"""
    k = normalize_value_key(v)
    if k in ("FAIT", "A"):
        return colors.HexColor("#007A33")
    if k in ("ENCOURS", "ECA"):
        return colors.HexColor("#ED7D31")
    if k in ("NA", "N.A"):
        return colors.HexColor("#C00000")
    if k in ("NE",):
        return colors.HexColor("#808080")
    return colors.black

def detect_eval_columns(df):
    """
    Return list of candidate evaluation columns (names as in df.columns).
    """
    meta_keywords = ["prenom", "nom", "e-mail", "email", "organisation", "departement",
                     "date", "temps", "taux", "score", "tentative", "reussite", "nbre",
                     "participan", "stagiaire_evalue", "evaluation_de_la_journee", "jcmsplugin"]
    eval_cols = []
    for c in df.columns:
        nc = normalise_colname(c)
        if any(m in nc for m in meta_keywords):
            continue
        if ("app_" in nc) or ("evaluation" in nc) or ("msp" in nc) or ("test" in nc) or ("axe" in nc) or ("ancrage" in nc):
            eval_cols.append(c)
        else:
            sample_vals = df[c].dropna().astype(str).head(10).tolist()
            if any(re.search(r"\b(fait|en cours|e\.c\.a|eca|na|ne|a)\b", v, re.IGNORECASE) for v in sample_vals):
                eval_cols.append(c)
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

    df_sorted = df.sort_values(by=[stagiaire_col_name])
    eval_columns = detect_eval_columns(df_sorted)

    # Build styles (AJUSTÉS)
    styles = getSampleStyleSheet()
    
    title_style = ParagraphStyle("Title", parent=styles["Heading1"], alignment=1, fontSize=16, textColor=colors.HexColor("#0B5394"))
    subtitle_style = ParagraphStyle("Sub", parent=styles["Normal"], alignment=1, fontSize=9, textColor=colors.grey)
    name_style = ParagraphStyle("Name", parent=styles["Heading2"], alignment=1, fontSize=14, textColor=colors.HexColor("#0B5394"), spaceAfter=6) 
    
    # Styles de base des cellules de texte
    cell_style = ParagraphStyle("Cell", parent=styles["Normal"], fontSize=9, leading=11, spaceAfter=2)
    legend_style = ParagraphStyle("Legend", parent=styles["Normal"], fontSize=8, spaceBefore=12) 
    
    # Style de la bande de section (Utilisé pour le titre de section)
    band_style = ParagraphStyle("Band", parent=styles["Normal"], alignment=0, textColor=colors.white, fontSize=10, leading=12, fontName='Helvetica-Bold')
    band_bg = colors.HexColor("#0B5394") # Bleu foncé pour le contraste
    
    buffer = BytesIO()

    elements_all = []
    grouped = df_sorted.groupby(stagiaire_col_name, sort=True)
    export_dt = datetime.now().strftime("%d/%m/%Y %H:%M")

    for stagiaire, group in grouped:
        header_parts = []
        header_parts.append(Paragraph("Synthèse hebdomadaire des évaluations - Stage SAV 2", title_style))
        header_parts.append(Paragraph(f"Exporté le : {export_dt}", subtitle_style))
        header_parts.append(Spacer(1, 8))
        header_parts.append(Paragraph(clean_display_text(stagiaire), name_style))
        header_parts.append(Spacer(1, 8))
        
        group = group.copy()
        if "_parsed_date" in group.columns:
            group["_date_only"] = group["_parsed_date"].dt.date
        else:
            group["_date_only"] = pd.NaT

        if group["_date_only"].isnull().all():
            group["_date_group_key"] = group[date_col].fillna("").astype(str).str.strip()
            group = group.sort_values(by=("_date_group_key"))
        else:
            group["_date_group_key"] = group["_date_only"].apply(lambda x: x if pd.notna(x) else None)
            group = group.sort_values(by=("_parsed_date"))

        page_elements = []
        page_elements.extend(header_parts)

        # For each date group in chronological order
        for date_key, sub in group.groupby("_date_group_key", sort=True):
            if date_key is None or (isinstance(date_key, str) and date_key == ""):
                date_label = ""
            else:
                if isinstance(date_key, datetime) or hasattr(date_key, "strftime"):
                    try:
                        date_label = date_key.strftime("%d/%m/%Y")
                    except Exception:
                        date_label = str(date_key)
                else:
                    date_label = str(date_key)

            # Determine formateurs present for that date
            if prenom_col in group.columns and nom_col in group.columns:
                formateurs = sub[[prenom_col, nom_col]].fillna("").astype(str)
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

            # Date + Formateur (plus espacé)
            if date_label:
                page_elements.append(Paragraph(f"<b>Date :</b> {escape(date_label)}    <b>Formateur :</b> {escape(clean_display_text(formateur_label))}", cell_style))
            else:
                page_elements.append(Paragraph(f"<b>Formateur :</b> {escape(clean_display_text(formateur_label))}", cell_style))

            page_elements.append(Spacer(1, 6))

            # Group evaluation columns
            type_buckets = {}
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

            if not type_buckets:
                page_elements.append(Paragraph("Aucune évaluation renseignée pour cette date.", cell_style))
                page_elements.append(Spacer(1, 6))
                continue

            # For each type, add band and table
            for tlabel, cols in type_buckets.items():
                # band
                band_table = Table([[Paragraph(f"<b>{escape(tlabel)}</b>", band_style)]], colWidths=[17*cm]) # Élargir un peu
                band_table.setStyle(TableStyle([
                    ('BACKGROUND', (0, 0), (-1, -1), band_bg),
                    ('LEFTPADDING', (0,0), (-1,-1), 8),
                    ('RIGHTPADDING', (0,0), (-1,-1), 8),
                    ('TOPPADDING', (0,0), (-1,-1), 4),
                    ('BOTTOMPADDING', (0,0), (-1,-1), 4),
                ]))
                page_elements.append(band_table)
                page_elements.append(Spacer(1, 4))

                # Build table rows
                tbl_rows = []
                header = [Paragraph("<b>Séquence / Épreuve</b>", cell_style), Paragraph("<b>Résultat</b>", cell_style)]
                tbl_rows.append(header)
                
                # TableStyle (sera rempli dans la boucle)
                table_style = TableStyle([
                    ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#2B6EB3")), # En-tête bleu
                    ('TEXTCOLOR', (0,0), (-1,0), colors.white),
                    ('ALIGN', (0,0), (-1,0), 'CENTER'),
                    ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
                    ('GRID', (0,0), (-1,-1), 0.25, colors.grey),
                    ('BOX', (0,0), (-1,-1), 0.5, colors.grey),
                    ('LEFTPADDING', (0,0), (-1,-1), 6),
                    ('RIGHTPADDING', (0,0), (-1,-1), 6),
                    ('TOPPADDING', (0,0), (-1,-1), 4),
                    ('BOTTOMPADDING', (0,0), (-1,-1), 4),
                    ('ALIGN', (1,1), (1,-1), 'CENTER'), # Centre les résultats de la colonne 1
                ])
                
                row_idx_in_table = 1 # Démarre après l'en-tête (index 0)

                for col in cols:
                    seq_label = col
                    if "/" in col:
                        seq_label = col.split("/")[-1]
                    seq_label = clean_display_text(seq_label).replace("_", " ").strip()
                    
                    vals = []
                    for _, r in sub.iterrows():
                        v = r.get(col, "")
                        if pd.notna(v) and str(v).strip() != "":
                            vt = clean_display_text(v)
                            if vt not in vals:
                                vals.append(vt)
                    
                    if not vals:
                        continue
                    combined_val = " / ".join(vals)
                    
                    # NOUVEAU : Récupérer les couleurs
                    bg_color, text_color = get_style_colors(combined_val)
                    
                    seq_par = Paragraph(escape(seq_label), cell_style)
                    
                    # Style de la valeur avec couleur de police adaptée et centré
                    val_style_custom = ParagraphStyle("val_custom", parent=cell_style, textColor=text_color, alignment=1)
                    val_par = Paragraph(escape(combined_val), val_style_custom)
                    
                    tbl_rows.append([seq_par, val_par])
                    
                    # AJOUT DU FOND DE CELLULE CONDITIONNEL
                    if bg_color != colors.white:
                        table_style.add('BACKGROUND', (0, row_idx_in_table), (-1, row_idx_in_table), bg_color)
                        
                    # Alternance de fond (maintenu, mais écrasé par la couleur conditionnelle)
                    if row_idx_in_table % 2 == 0 and bg_color == colors.white:
                        table_style.add('BACKGROUND', (0,row_idx_in_table), (-1,row_idx_in_table), colors.whitesmoke)
                        
                    row_idx_in_table += 1


                col_widths = [10*cm, 6*cm]
                table = Table(tbl_rows, colWidths=col_widths, hAlign='LEFT')
                table.setStyle(table_style)
                page_elements.append(table)
                page_elements.append(Spacer(1, 6))

        # --- Sections de texte libre (Axes, Ancrage, etc.) ---
        def first_nonempty_from_group(cols_list):
            for c in cols_list:
                if c in df.columns:
                    v = group.iloc[0].get(c, "")
                    if pd.notna(v) and str(v).strip():
                        return clean_display_text(v)
            return ""

        axes_cols = [c for c in df.columns if "axe" in normalise_colname(c) or "progression" in normalise_colname(c)]
        ancrage_cols = [c for c in df.columns if "ancrag" in normalise_colname(c) or "ancrage" in normalise_colname(c) or "point_d'ancrage" in normalise_colname(c)]
        app_prop_cols = [c for c in df.columns if "app_qui" in normalise_colname(c) or "pourrait" in normalise_colname(c) or "propose" in normalise_colname(c)]

        axes_text = first_nonempty_from_group(axes_cols)
        ancrage_text = first_nonempty_from_group(ancrage_cols)
        app_prop_text = first_nonempty_from_group(app_prop_cols)

        # Style pour les titres de section (H4)
        h4_style = ParagraphStyle("sec_h4", parent=styles["Heading4"], textColor=colors.HexColor("#0B5394"), spaceBefore=8)

        if axes_text:
            page_elements.append(Paragraph("<b>Axes de progression</b>", h4_style))
            page_elements.append(Paragraph(escape(axes_text), cell_style))
            page_elements.append(Spacer(1,6))
        if ancrage_text:
            page_elements.append(Paragraph("<b>Points d'ancrage</b>", h4_style))
            page_elements.append(Paragraph(escape(ancrage_text), cell_style))
            page_elements.append(Spacer(1,6))
        if app_prop_text:
            page_elements.append(Paragraph("<b>APP qui pourraient être proposés</b>", h4_style))
            page_elements.append(Paragraph(escape(app_prop_text), cell_style))
            page_elements.append(Spacer(1,6))

        # Légende
        legend = "Légende : Fait / A = Acquis (vert)  •  En cours / ECA = En cours (jaune/orange)  •  NA = Non acquis (rouge)  •  NE = Non évalué (gris)"
        page_elements.append(Spacer(1, 8))
        page_elements.append(Paragraph(escape(legend), legend_style))

        # Saut de page
        page_elements.append(PageBreak())

        elements_all.extend(page_elements)

    # Build document (Marges ajustées)
    doc = SimpleDocTemplate(buffer, pagesize=A4, leftMargin=1.5*cm, rightMargin=1.5*cm, topMargin=1.5*cm, bottomMargin=1.5*cm)
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

    stag_col = None
    for c in df.columns:
        if 'stagiaire' in c.lower() or 'participant' in c.lower() or 'élève' in c.lower() or 'eleve' in c.lower():
            stag_col = c
            break
    if stag_col is None:
        st.error("Colonne 'Stagiaire évalué' non trouvée. Vérifie le fichier.")
        st.stop()

    prenom_col = None
    nom_col = None
    for c in df.columns:
        lc = c.lower()
        if 'prenom' in lc and prenom_col is None:
            prenom_col = c
        if 'nom' == lc or ('nom' in lc and 'prenom' not in lc) and nom_col is None:
            nom_col = c

    date_col = None
    for c in df.columns:
        if 'date' in c.lower():
            date_col = c
            break

    st.write(f"Stagiaire col: **{stag_col}** | Prenom col: **{prenom_col}** | Nom col: **{nom_col}** | Date col: **{date_col}**")

    if st.button("📄 Générer la synthèse PDF (une page par stagiaire)"):
        try:
            pdf_bytes = build_pdf_bytes(df, stag_col, prenom_col, nom_col, date_col)
            st.success("PDF généré.")
            st.download_button("⬇️ Télécharger le PDF", data=pdf_bytes, file_name="synthese_evaluations_stage_sav2.pdf", mime="application/pdf")
        except Exception as e:
            st.error(f"Erreur lors de la génération du PDF : {e}")
