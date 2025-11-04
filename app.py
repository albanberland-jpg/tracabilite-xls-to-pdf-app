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
st.write("Charge un fichier Excel pour créer un PDF clair et coloré, une fiche par stagiaire (une fiche par ligne du fichier).")

uploaded_file = st.file_uploader("Choisir le fichier Excel (.xlsx)", type=["xlsx"])

# ---------------- utilities ----------------
def normaliser_colname(name):
    """Normalise le nom de colonne : retire accents, espaces en tiret bas, retire certains caractères spéciaux."""
    s = str(name)
    s = ''.join(ch for ch in unicodedata.normalize("NFKD", s) if not unicodedata.combining(ch))
    s = s.lower().strip()
    s = re.sub(r"\s+", "_", s)
    s = re.sub(r"[^a-z0-9_/()'’.-]", "", s)
    return s

def nettoyer_texte_visible(txt):
    """Garde seulement caractères lisibles (lettres accentuées, chiffres, ponctuation basique)."""
    if pd.isna(txt):
        return ""
    s = str(txt)
    # Normaliser accents (déjà fait par unicodedata.normalize("NFKD", s) dans la fonction originale, mais on le sécurise)
    s = ''.join(ch for ch in unicodedata.normalize("NFKD", s) if not unicodedata.combining(ch))
    # Retire symboles non désirés par ReportLab/PDF et caractères invisibles
    s = re.sub(r"[^A-Za-zÀ-ÖØ-öø-ÿ0-9 ,;:!\?'\(\)\[\]\-\/\.%&°%\"’\n]", " ", s)
    # Collapse spaces
    s = re.sub(r"\s+", " ", s).strip()
    return s

def valeur_cle(val):
    """Renvoie clé normalisée pour comparaison stricte (supprime points/espaces/accents)."""
    if pd.isna(val):
        return ""
    s = str(val).upper()
    s = s.replace(".", "").replace(" ", "").strip()
    # Retirer accents pour la clé de mapping
    s = ''.join(ch for ch in unicodedata.normalize("NFKD", s) if not unicodedata.combining(ch))
    return s

# ---------------- coloriser (ordre prioritaire) ----------------
def coloriser_valeur_html(val):
    """Retourne HTML <font> coloré si correspondance exacte après nettoyage."""
    key = valeur_cle(val)
    
    # Définition des couleurs exactes demandées
    # IMPORTANT: L'ordre est maintenu (ECA avant A, car "A" pourrait être une sous-chaîne de "ECA" si la normalisation était différente)
    mapping = [
        ("FAIT", "#007A33"),      # Vert foncé pour FAIT
        ("ENCOURS", "#FFD700"),   # Jaune/Or pour EN COURS
        ("NE", "#808080"),        # Gris pour NE
        ("NA", "#C00000"),        # Rouge pour NA
        ("ECA", "#ED7D31"),       # Orange pour ECA
        ("A", "#00B050"),         # Vert clair pour A
    ]
    
    # Nettoyage visuel de la valeur
    display = nettoyer_texte_visible(val)
    
    for k, hexc in mapping:
        if key == k:
            # Retourne la valeur colorée et en gras
            return f"<font color='{hexc}'><b>{display}</b></font>"
            
    # Fallback: retourne le texte nettoyé sans couleur
    return f"<b>{display}</b>" # Garde le gras pour les notes non reconnues

# ---------------- main app ----------------
if uploaded_file:
    df = pd.read_excel(uploaded_file)
    # normalize column names
    df.columns = [normaliser_colname(c) for c in df.columns]

    # detect main columns robustly
    stagiaire_col = next((c for c in df.columns if "stagiaire" in c or "participant" in c or "eleve" in c), None)
    date_col = next((c for c in df.columns if "date" in c), None)
    # Tente de trouver le nom et prénom si pas de colonne "formateur" explicite
    prenom_col = next((c for c in df.columns if "prenom" in c), None)
    nom_col = next((c for c in df.columns if "nom" in c and "prenom" not in c), None)

    if stagiaire_col is None:
        st.error("❌ Colonne stagiaire non trouvée — vérifiez l'en-tête du fichier pour inclure 'stagiaire', 'participant' ou 'élève'.")
        st.stop()
    
    st.success("✅ Fichier importé avec succès. Colonne stagiaire détectée.")

    # Crée la colonne "formateur" si possible, sinon vide
    formateur_col = next((c for c in df.columns if "formateur" in c), None)
    if formateur_col is None and prenom_col and nom_col:
         df["formateur"] = df[prenom_col].astype(str).str.strip() + " " + df[nom_col].astype(str).str.strip()
         formateur_col = "formateur"
    elif formateur_col is None:
        df["formateur"] = ""
        formateur_col = "formateur" # Pour la référence dans le PDF

    # detect app columns (tolerant contains)
    app_non_eval_cols = [c for c in df.columns if "app_non" in c or "non_soumis" in c]
    app_eval_cols = [c for c in df.columns if "app_evalue" in c or "app_evalues" in c or "app_eval" in c]
    axes_cols = [c for c in df.columns if "axe" in c or "axes_de_progression" in c]
    ancrage_cols = [c for c in df.columns if "ancrage" in c or "point_d_ancrage" in c or "point_d'anc" in c]
    app_prop_cols = [c for c in df.columns if "app_qui_pourrait" in c or "app_qui_peut" in c or "propose" in c]

    # PDF styles
    styles = getSampleStyleSheet()
    titre_style = ParagraphStyle("Titre", parent=styles["Heading1"], alignment=1, fontSize=18, textColor=colors.HexColor("#003366"), spaceAfter=12)
    section_style = ParagraphStyle("Section", parent=styles["Heading3"], fontSize=12, textColor=colors.HexColor("#003366"), spaceBefore=10, spaceAfter=4)
    item_style = ParagraphStyle("Item", parent=styles["Normal"], fontSize=10, leading=14, allowHTML=True, spaceAfter=2, leftIndent=12)

    # prepare PDF
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, leftMargin=50, rightMargin=50, topMargin=50, bottomMargin=50)
    elements = []

    # --- Fonction utilitaire pour générer des sections d'évaluation ---
    def generate_evaluation_section(title, cols, elements, section_style, item_style, row):
        elements.append(Paragraph(title, section_style))
        any_item = False
        for c in cols:
            v = row.get(c, "")
            # Check for non-empty/non-NaN values
            if pd.notna(v) and str(v).strip() not in ["", "nan"]:
                # Nettoyage du nom de l'APP pour l'affichage (enlève le préfixe long et normalisé)
                nom_app_clean = c.split("/")[-1].replace("_", " ").title().strip()
                
                v_str = str(v).strip()
                
                # Hypothèse: si la valeur est courte (<= 15 chars), c'est une note -> coloriser
                if len(v_str) <= 15 and nom_app_clean != v_str.replace(" ", "").title():
                    # Colorise la note
                    formatted_value = coloriser_valeur_html(v)
                    elements.append(Paragraph(f"• {nom_app_clean} : {formatted_value}", item_style))
                else:
                    # Traite le contenu comme un commentaire long (texte noir)
                    elements.append(Paragraph(f"• {nom_app_clean} : {nettoyer_texte_visible(v)}", item_style))
                    
                any_item = True
        if not any_item:
            elements.append(Paragraph("(Aucun item trouvé)", item_style))
        elements.append(Spacer(1, 6))
        
    # --- Fonction utilitaire pour générer les sections de texte (axes, ancrages...) ---
    def generate_text_section(title, cols, elements, section_style, item_style, row):
        elements.append(Paragraph(title, section_style))
        any_item = False
        for c in cols:
            v = row.get(c, "")
            if pd.notna(v) and str(v).strip() not in ["", "nan"]:
                elements.append(Paragraph(f"– {nettoyer_texte_visible(v)}", item_style))
                any_item = True
        if not any_item:
            elements.append(Paragraph("(Aucun commentaire ou axe renseigné)", item_style))
        elements.append(Spacer(1, 6))
        
    # --- Itération et Génération des Fiches ---
    # Itère par stagiaire et par ligne (pour les cas où il y a plusieurs évaluations par stagiaire)
    for stagiaire, group in df.groupby(stagiaire_col):
        for idx, row in group.iterrows():
            # Saut de page pour la fiche suivante
            if elements:
                elements.append(Spacer(1, 10))
                elements.append(Paragraph("<pageBreak/>", styles['Normal']))

            # --- En-tête ---
            elements.append(Paragraph("Fiche d’évaluation", titre_style))
            elements.append(Spacer(1, 8))
            elements.append(Paragraph(f"<b>Stagiaire :</b> {nettoyer_texte_visible(stagiaire)}", item_style))
            
            date_info = nettoyer_texte_visible(row.get(date_col,'')) if date_col in row else ''
            formateur_info = nettoyer_texte_visible(row.get(formateur_col,'')) if formateur_col in row else ''

            if date_info:
                elements.append(Paragraph(f"<b>Date :</b> {date_info}", item_style))
            if formateur_info:
                elements.append(Paragraph(f"<b>Formateur :</b> {formateur_info}", item_style))
            elements.append(Spacer(1, 10))

            # --- Sections d'évaluation (avec couleurs) ---
            generate_evaluation_section("APP non soumis à évaluation", app_non_eval_cols, elements, section_style, item_style, row)
            generate_evaluation_section("APP évalués", app_eval_cols, elements, section_style, item_style, row)

            # --- Sections de texte libre (sans couleurs) ---
            generate_text_section("Axes de progression", axes_cols, elements, section_style, item_style, row)
            generate_text_section("Points d'ancrage", ancrage_cols, elements, section_style, item_style, row)
            generate_text_section("APP qui pourraient être proposés", app_prop_cols, elements, section_style, item_style, row)

    # --- Génération finale du PDF ---
    if elements:
        try:
            doc.build(elements)
            buffer.seek(0)
            st.success("PDF prêt.")
            st.download_button(
                "📄 Télécharger le PDF des fiches", 
                data=buffer.getvalue(), 
                file_name="fiches_evaluations.pdf", 
                mime="application/pdf"
            )
        except Exception as e:
             st.error(f"Une erreur est survenue lors de la création du PDF : {e}")
    else:
        st.warning("Le fichier ne contient aucune donnée pour générer les fiches.")
