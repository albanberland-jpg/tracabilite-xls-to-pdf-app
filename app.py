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

# --- COULEURS DE FOND DE CELLULE (Utilisation des couleurs ReportLab natives) ---
COULEURS_FOND = {
    "FAIT": colors.HexColor("#A9D18E"),    # Vert clair Excel
    "A": colors.HexColor("#70AD47"),       # Vert foncé Excel
    "ENCOURS": colors.HexColor("#FFC000"), # Jaune/Orange Excel
    "ECA": colors.HexColor("#ED7D31"),     # Orange Excel
    "NE": colors.HexColor("#D9D9D9"),      # Gris clair
    "NA": colors.HexColor("#F8CBAD"),      # Rouge très clair
}

# COULEURS DE POLICE ASSOCIÉES
# Blanc pour le fond foncé 'A' (Vert foncé), Noir pour les autres
COULEURS_POLICE = {
    "A": colors.white, 
}
DEFAULT_POLICE_COLOR = colors.black

# --- Utilities (Identiques pour la robustesse) ---
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

# --- Fonction de génération de tableau (AJOUT DE LA COULEUR DE POLICE) ---
def generate_app_table(title, cols, row, item_style, item_bold_style):
    data = []
    styles = []
    
    # 1. En-tête du tableau
    header = [
        Paragraph("<b>Séquence</b>", item_bold_style),
        Paragraph("<b>Résultats / Évaluations</b>", item_bold_style),
    ]
    data.append(header)
    
    # 2. Remplissage des lignes et application des styles
    row_idx = 1
    
    for c in cols:
        v = row.get(c, "")
        if pd.notna(v) and str(v).strip():
            note_cle = valeur_cle(v)
            
            # 2a. Définition des éléments de la ligne
            nom_app_clean = c.split("/")[-1].replace("_", " ").strip().title()
            
            # On utilise le style pour les items (texte noir)
            cell_nom = Paragraph(nettoyer_texte_visible(nom_app_clean), item_style)
            cell_valeur = Paragraph(nettoyer_texte_visible(v), item_style)
            
            data.append([cell_nom, cell_valeur])
            
            # 2b. Application du style de fond
            if note_cle in COULEURS_FOND:
                styles.append(
                    ('BACKGROUND', (1, row_idx), (1, row_idx), COULEURS_FOND[note_cle]) # Colonne des notes (index 1)
                )
                
            # 2c. APPLICATION DE LA COULEUR DE POLICE (C'est la correction !)
            text_color = COULEURS_POLICE.get(note_cle, DEFAULT_POLICE_COLOR)
            styles.append(
                ('TEXTCOLOR', (1, row_idx), (1, row_idx), text_color)
            )
            
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
    
    table.setStyle(TableStyle(general_style + styles))
    return table

# --- Application principale ---
if uploaded_file:
    df = pd.read_excel(uploaded_file)
    df.columns = [normaliser_colname(c) for c in df.columns]

    # Détection automatique des colonnes
    stagiaire_col = next((c for c in df.columns if "stagiaire" in c), None)
    date_col = next((c for c in df.columns if "date" in c), None)
    prenom_col = next((c for c in df.columns if "prenom" in c), None)
    nom_col = next((c for c in df.columns if "nom" in c and "prenom" not in c), None)

    if not stagiaire_col:
        st.error("⚠️ Colonne stagiaire introuvable dans le fichier.")
        st.stop()
    
    st.success(f"✅ Fichier importé. Fiches générées par la colonne : **{stagiaire_col}**.")


    # Définition de la colonne formateur
    formateur_col_auto = next((c for c in df.columns if "formateur" in c), None)
    if formateur_col_auto is not None:
         df["formateur_display"] = df[formateur_col_auto]
    elif prenom_col and nom_col:
         df["formateur_display"] = df[prenom_col].astype(str).str.strip() + " " + df[nom_col].astype(str).str.strip()
    else:
        df["formateur_display"] = "N/A"
    
    # Regroupement de colonnes par type
    app_non_eval_cols = [c for c in df.columns if "app_non" in c or "non_soumis" in c]
    app_eval_cols = [c for c in df.columns if "app_evalue" in c or "app_eval" in c]
    axes_cols = [c for c in df.columns if "axe" in c or "progression" in c]
    ancrage_cols = [c for c in df.columns if "ancrage" in c or "ancr" in c]
    app_prop_cols = [c for c in df.columns if "app_qui" in c or "propose" in c]

    # Styles PDF
    styles = getSampleStyleSheet()
    titre_style = ParagraphStyle("Titre", parent=styles["Heading1"], alignment=1, fontSize=18, textColor=colors.HexColor("#007A33"), spaceAfter=12)
    section_style = ParagraphStyle("Section", parent=styles["Heading3"], fontSize=12, textColor=colors.HexColor("#003366"), spaceBefore=8, spaceAfter=6)
    
    item_style = ParagraphStyle("Item", parent=styles["Normal"], fontSize=10, leading=13, spaceAfter=0, leftIndent=0, allowHTML=False)
    item_bold_style = ParagraphStyle("ItemBold", parent=styles["Normal"], fontSize=10, leading=13, spaceAfter=0, leftIndent=0, fontName='Helvetica-Bold')


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
        elements.append(Paragraph(f"<b>Stagiaire :</b> {nettoyer_texte_visible(stagiaire)}", item_bold_style))
        if date_col:
            elements.append(Paragraph(f"<b>Date :</b> {nettoyer_texte_visible(first_row.get(date_col, ''))}", item_bold_style))
        elements.append(Paragraph(f"<b>Formateur :</b> {nettoyer_texte_visible(first_row.get('formateur_display', ''))}", item_bold_style))
        elements.append(Spacer(1, 8))

        # --- GESTION DES SECTIONS APP EN TABLEAUX ---
        
        # 1. APP non soumis à évaluation
        elements.append(Paragraph(f"<b>APP non soumis à évaluation</b>", section_style))
        table_non_eval = generate_app_table("APP non soumis à évaluation", app_non_eval_cols, first_row, item_style, item_bold_style)
        if table_non_eval:
            elements.append(table_non_eval)
        else:
            elements.append(Paragraph("Aucun item", item_style))
        elements.append(Spacer(1, 6))

        # 2. APP évalués
        elements.append(Paragraph(f"<b>APP évalués</b>", section_style))
        table_eval = generate_app_table("APP évalués", app_eval_cols, first_row, item_style, item_bold_style)
        if table_eval:
            elements.append(table_eval)
        else:
            elements.append(Paragraph("Aucun item", item_style))
        elements.append(Spacer(1, 6))

        # --- GESTION DES SECTIONS TEXTE LIBRE ---
        
        # Fonction d'ajout de section (utilisée uniquement pour les blocs texte)
        def add_text_section(title, cols):
            elements.append(Paragraph(f"<b>{title}</b>", section_style))
            added = False
            for c in cols:
                v = first_row.get(c, "")
                if pd.notna(v) and str(v).strip():
                    elements.append(Paragraph(f"• {nettoyer_texte_visible(v)}", item_style))
                    added = True
            
            if not added:
                elements.append(Paragraph("Aucun item", item_style))
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
