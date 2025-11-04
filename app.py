import streamlit as st
import pandas as pd
from io import BytesIO
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, PageBreak, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
import unicodedata, re

st.set_page_config(page_title="Tracabilité XLS → PDF", layout="centered")
st.title("📘 Générateur de fiches d’évaluation")
st.write("Ce script génère un PDF récapitulatif avec des tableaux colorés (couleur de fond de cellule et couleur de police adaptée).")

uploaded_file = st.file_uploader("Choisir le fichier Excel (.xlsx)", type=["xlsx"])

# --- COULEURS DE FOND DE CELLULE (ReportLab) ---
COULEURS_FOND = {
    "FAIT": colors.HexColor("#A9D18E"),    
    "A": colors.HexColor("#70AD47"),       
    "ENCOURS": colors.HexColor("#FFC000"), 
    "ECA": colors.HexColor("#ED7D31"),     
    "NE": colors.HexColor("#D9D9D9"),      
    "NA": colors.HexColor("#F8CBAD"),      
}

# COULEURS DE POLICE (ReportLab)
POLICE_BLANC = colors.white
POLICE_NOIR = colors.black

# --- Utilities ---
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
    s = re.sub(r"[_•■\u25a0\u200b\u2013\u2014]", " ", s) 
    s = unicodedata.normalize("NFKD", s).encode('ascii', 'ignore').decode('utf-8')
    s = re.sub(r"\s+", " ", s).strip()
    return s

def valeur_cle(val):
    if pd.isna(val):
        return ""
    s = str(val).upper()
    s = s.replace(".", "").replace(" ", "").strip()
    s = ''.join(ch for ch in unicodedata.normalize("NFKD", s) if not unicodedata.combining(ch))
    return s

# --- STYLES DE PARAGRAPHE PRÉ-COLORÉS (NOUVEAU) ---
styles = getSampleStyleSheet()
base_style = ParagraphStyle("BaseItem", parent=styles["Normal"], fontSize=10, leading=13, spaceAfter=0, leftIndent=0, alignment=1)

# Création des styles de police pour la colonne des notes
NOTE_STYLES = {
    "FAIT": ParagraphStyle("NoteFAIT", parent=base_style, textColor=POLICE_NOIR),
    "A": ParagraphStyle("NoteA", parent=base_style, textColor=POLICE_BLANC), # Blanc sur fond vert foncé
    "ENCOURS": ParagraphStyle("NoteEC", parent=base_style, textColor=POLICE_NOIR),
    "ECA": ParagraphStyle("NoteECA", parent=base_style, textColor=POLICE_NOIR),
    "NE": ParagraphStyle("NoteNE", parent=base_style, textColor=POLICE_NOIR),
    "NA": ParagraphStyle("NoteNA", parent=base_style, textColor=POLICE_NOIR),
    "DEFAULT": ParagraphStyle("NoteDEF", parent=base_style, textColor=POLICE_NOIR),
}

# --- Fonction de génération de tableau (Utilise les styles de note) ---
def generate_app_table(title, cols, row, item_style):
    data = []
    styles_table = [] # Renommé pour éviter la confusion avec styles=getSampleStyleSheet()
    
    # 1. En-tête du tableau
    # Style de l'en-tête (gras + centré)
    header_style = ParagraphStyle("Header", parent=item_style, fontName='Helvetica-Bold', alignment=1)
    header = [
        Paragraph("Séquence", header_style),
        Paragraph("Résultats / Évaluations", header_style),
    ]
    data.append(header)
    
    # 2. Remplissage des lignes et application des styles
    row_idx = 1
    
    for c in cols:
        v = row.get(c, "")
        if pd.notna(v) and str(v).strip():
            note_cle = valeur_cle(v)
            
            # Définition des éléments de la ligne
            nom_app_clean = c.split("/")[-1].replace("_", " ").strip().title()
            
            # --- PARAGRAPHES ---
            cell_nom = Paragraph(nettoyer_texte_visible(nom_app_clean), item_style)
            
            # CLÉ DE LA CORRECTION : Utiliser le style de paragraphe prédéfini (Note_STYLES)
            note_paragraph_style = NOTE_STYLES.get(note_cle, NOTE_STYLES["DEFAULT"])
            cell_valeur = Paragraph(nettoyer_texte_visible(v), note_paragraph_style)
            
            data.append([cell_nom, cell_valeur])
            
            # Application du style de fond de cellule
            if note_cle in COULEURS_FOND:
                styles_table.append(
                    ('BACKGROUND', (1, row_idx), (1, row_idx), COULEURS_FOND[note_cle])
                )
            
            # Pas besoin de la commande 'TEXTCOLOR' dans TableStyle, car le ParagraphStyle le gère.
            
            row_idx += 1

    if len(data) == 1: # Seulement l'en-tête
        return None
        
    # 3. Création du tableau et du style général
    table = Table(data, colWidths=[3.5 * inch, 1.5 * inch])
    
    general_style = [
        ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
        ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        # Style pour l'en-tête
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#D9E1F2")),
        ('ALIGN', (1, 0), (1, -1), 'CENTER'), # Aligne la colonne des notes au centre
    ]
    
    table.setStyle(TableStyle(general_style + styles_table))
    return table

# --- Application principale ---
if uploaded_file:
    df = pd.read_excel(uploaded_file)
    df.columns = [normaliser_colname(c) for c in df.columns]

    # Détection automatique des colonnes (inchangé)
    stagiaire_col = next((c for c in df.columns if "stagiaire" in c), None)
    date_col = next((c for c in df.columns if "date" in c), None)
    prenom_col = next((c for c in df.columns if "prenom" in c), None)
    nom_col = next((c for c in df.columns if "nom" in c and "prenom" not in c), None)

    if not stagiaire_col:
        st.error("⚠️ Colonne stagiaire introuvable dans le fichier.")
        st.stop()
    
    st.success(f"✅ Fichier importé. Fiches générées par la colonne : **{stagiaire_col}**.")


    # Définition de la colonne formateur (inchangé)
    formateur_col_auto = next((c for c in df.columns if "formateur" in c), None)
    if formateur_col_auto is not None:
         df["formateur_display"] = df[formateur_col_auto]
    elif prenom_col and nom_col:
         df["formateur_display"] = df[prenom_col].astype(str).str.strip() + " " + df[nom_col].astype(str).str.strip()
    else:
        df["formateur_display"] = "N/A"
    
    # Regroupement de colonnes par type (inchangé)
    app_non_eval_cols = [c for c in df.columns if "app_non" in c or "non_soumis" in c]
    app_eval_cols = [c for c in df.columns if "app_evalue" in c or "app_eval" in c]
    axes_cols = [c for c in df.columns if "axe" in c or "progression" in c]
    ancrage_cols = [c for c in df.columns if "ancrage" in c or "ancr" in c]
    app_prop_cols = [c for c in df.columns if "app_qui" in c or "propose" in c]

    # Styles PDF de texte
    item_style_normal = ParagraphStyle("ItemNormal", parent=styles["Normal"], fontSize=10, leading=13, spaceAfter=0, leftIndent=0)
    item_style_bold = ParagraphStyle("ItemBold", parent=styles["Normal"], fontSize=10, leading=13, spaceAfter=0, leftIndent=0, fontName='Helvetica-Bold')

    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, leftMargin=50, rightMargin=50, topMargin=50, bottomMargin=50)
    elements = []

    # --- Boucle de génération du PDF ---
    for stagiaire, group in df.groupby(stagiaire_col):
        first_row = group.iloc[0]

        if elements:
            elements.append(PageBreak())

        # En-tête de fiche
        elements.append(Paragraph("Fiche d’évaluation", titre_style))
        elements.append(Paragraph(f"<b>Stagiaire :</b> {nettoyer_texte_visible(stagiaire)}", item_style_bold))
        if date_col:
            elements.append(Paragraph(f"<b>Date :</b> {nettoyer_texte_visible(first_row.get(date_col, ''))}", item_style_bold))
        elements.append(Paragraph(f"<b>Formateur :</b> {nettoyer_texte_visible(first_row.get('formateur_display', ''))}", item_style_bold))
        elements.append(Spacer(1, 8))

        # --- GESTION DES SECTIONS APP EN TABLEAUX ---
        
        # 1. APP non soumis à évaluation
        elements.append(Paragraph(f"<b>APP non soumis à évaluation</b>", section_style))
        table_non_eval = generate_app_table("APP non soumis à évaluation", app_non_eval_cols, first_row, item_style_normal)
        if table_non_eval:
            elements.append(table_non_eval)
        else:
            elements.append(Paragraph("Aucun item", item_style_normal))
        elements.append(Spacer(1, 6))

        # 2. APP évalués
        elements.append(Paragraph(f"<b>APP évalués</b>", section_style))
        table_eval = generate_app_table("APP évalués", app_eval_cols, first_row, item_style_normal)
        if table_eval:
            elements.append(table_eval)
        else:
            elements.append(Paragraph("Aucun item", item_style_normal))
        elements.append(Spacer(1, 6))

        # --- GESTION DES SECTIONS TEXTE LIBRE ---
        
        # Fonction d'ajout de section (utilisée uniquement pour les blocs texte)
        def add_text_section(title, cols):
            elements.append(Paragraph(f"<b>{title}</b>", section_style))
            added = False
            for c in cols:
                v = first_row.get(c, "")
                if pd.notna(v) and str(v).strip():
                    elements.append(Paragraph(f"• {nettoyer_texte_visible(v)}", item_style_normal))
                    added = True
            
            if not added:
                elements.append(Paragraph("Aucun item", item_style_normal))
            elements.append(Spacer(1, 6))

        # Ajout des sections de texte
        add_text_section("Axes de progression", axes_cols)
        add_text_section("Points d’ancrage", ancrage_cols)
        add_text_section("APP qui pourraient être proposés", app_prop_cols)

    # --- Finalisation ---
    if elements:
        try:
            doc.build(elements)
            buffer.seek(0)

            st.success("✅ PDF généré avec succès.")
            st.download_button("⬇️ Télécharger le PDF (Format Tableau)", data=buffer.getvalue(),
                               file_name="fiches_stagiaires_tableau.pdf", mime="application/pdf")
        except Exception as e:
             st.error(f"Une erreur est survenue lors de la construction du PDF : {e}")
    else:
         st.warning("Aucune donnée n'a été trouvée pour générer les fiches.")
