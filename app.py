import streamlit as st
import pandas as pd
from io import BytesIO
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
import unicodedata, re

st.set_page_config(page_title="Tracabilité XLS → PDF", layout="centered")
st.title("📘 Générateur de fiches d’évaluation")

uploaded_file = st.file_uploader("Choisir le fichier Excel (.xlsx)", type=["xlsx"])

# --- Nettoyage des noms de colonnes ---
def normaliser_colname(name):
    s = str(name)
    s = ''.join(ch for ch in unicodedata.normalize("NFKD", s) if not unicodedata.combining(ch))
    s = s.lower().strip()
    s = re.sub(r"\s+", "_", s)
    s = re.sub(r"[^a-z0-9_/()'’.-]", "", s)
    return s

# --- Nettoyage du texte pour affichage (version robuste) ---
def nettoyer_texte_visible(txt):
    if pd.isna(txt):
        return ""
    s = str(txt)
    # Remplacer les symboles par des espaces
    s = re.sub(r"[_•■]", " ", s)
    # Normalisation Unicode pour un nettoyage plus large des caractères non standards
    s = ''.join(ch for ch in unicodedata.normalize("NFKC", s) if unicodedata.category(ch) not in ('Cc', 'Cf', 'Cs', 'Co', 'Cn'))
    # Remplace les blocs d'espaces par un seul espace
    s = re.sub(r"\s+", " ", s).strip()
    return s

# --- Conversion d'une valeur en clé standardisée ---
def valeur_cle(val):
    if pd.isna(val):
        return ""
    s = str(val).upper()
    s = s.replace(".", "").replace(" ", "").strip()
    s = ''.join(ch for ch in unicodedata.normalize("NFKD", s) if not unicodedata.combining(ch))
    return s

# --- Application du code couleur HTML ---
def coloriser_valeur_html(val):
    key = valeur_cle(val)
    # On utilise des objets colors.HexColor pour s'assurer que ReportLab les comprend
    mapping = {
        "FAIT": colors.HexColor("#00B050"),    # vert clair
        "A": colors.HexColor("#007A33"),       # vert foncé
        "ENCOURS": colors.HexColor("#FFD700"),  # jaune
        "ECA": colors.HexColor("#ED7D31"),     # orange
        "NE": colors.HexColor("#808080"),      # gris
        "NA": colors.HexColor("#C00000")       # rouge
    }
    # Fallback noir si non trouvé
    color = mapping.get(key, colors.HexColor("#000000")) 
    txt = nettoyer_texte_visible(val)
    # ReportLab supporte les codes Hex en minuscules dans le tag <font>
    return f"<font color='{color.hexval()}'><b>{txt}</b></font>"

# --- Application principale ---
if uploaded_file:
    df = pd.read_excel(uploaded_file)
    df.columns = [normaliser_colname(c) for c in df.columns]

    st.write("🔍 Colonnes détectées :", list(df.columns))

    # Détection automatique
    stagiaire_col = next((c for c in df.columns if "stagiaire" in c), None)
    date_col = next((c for c in df.columns if "date" in c), None)
    prenom_col = next((c for c in df.columns if "prenom" in c), None)
    nom_col = next((c for c in df.columns if "nom" in c and "prenom" not in c), None)

    if not stagiaire_col:
        st.error("⚠️ Colonne stagiaire introuvable dans le fichier.")
        st.stop()

    # Définition de la colonne formateur
    formateur_col_auto = next((c for c in df.columns if "formateur" in c), None)
    if formateur_col_auto is not None:
         df["formateur_display"] = df[formateur_col_auto]
    elif prenom_col and nom_col:
         df["formateur_display"] = df[prenom_col].astype(str).str.strip() + " " + df[nom_col].astype(str).str.strip()
    else:
        df["formateur_display"] = "N/A" # Valeur par défaut si non trouvé
    
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
    
    # CORRECTION : Ajout de allowHTML=True ici est essentiel
    item_style = ParagraphStyle("Item", parent=styles["Normal"], fontSize=10, leading=13, spaceAfter=3, leftIndent=15, allowHTML=True)

    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, leftMargin=50, rightMargin=50, topMargin=50, bottomMargin=50)
    elements = []

    # --- Boucle de génération du PDF ---
    for stagiaire, group in df.groupby(stagiaire_col):
        # Utiliser la première ligne pour les métadonnées (nom, date, formateur)
        first_row = group.iloc[0]

        # Saut de page pour les fiches suivantes
        if elements:
            elements.append(PageBreak())

        # En-tête de fiche
        elements.append(Paragraph("Fiche d’évaluation", titre_style))
        elements.append(Paragraph(f"<b>Stagiaire :</b> {nettoyer_texte_visible(stagiaire)}", item_style))
        if date_col:
            elements.append(Paragraph(f"<b>Date :</b> {nettoyer_texte_visible(first_row.get(date_col, ''))}", item_style))
        elements.append(Paragraph(f"<b>Formateur :</b> {nettoyer_texte_visible(first_row.get('formateur_display', ''))}", item_style))
        elements.append(Spacer(1, 8))

        # Fonction d'ajout de section (améliorée pour gérer les longs commentaires vs notes)
        def add_section(title, cols):
            elements.append(Paragraph(f"<b>{title}</b>", section_style))
            added = False
            for c in cols:
                v = first_row.get(c, "")
                if pd.notna(v) and str(v).strip():
                    nom_app = nettoyer_texte_visible(c.split("/")[-1].replace("_", " "))
                    v_str = str(v).strip()
                    
                    # Si c'est une note courte (moins de 20 caractères), on la colorise
                    if len(v_str) < 20 and valeur_cle(v) in ["FAIT", "A", "ENCOURS", "ECA", "NE", "NA"]:
                        val_display = coloriser_valeur_html(v)
                        elements.append(Paragraph(f"• {nom_app} : {val_display}", item_style))
                    else:
                        # Si c'est un long texte, on l'affiche simplement
                        val_display = nettoyer_texte_visible(v)
                        elements.append(Paragraph(f"• {nom_app} : {val_display}", item_style))
                    added = True
            
            if not added:
                elements.append(Paragraph("Aucun item", item_style))
            elements.append(Spacer(1, 6))

        # Ajout des sections
        add_section("APP non soumis à évaluation", app_non_eval_cols)
        add_section("APP évalués", app_eval_cols)
        add_section("Axes de progression", axes_cols)
        add_section("Points d’ancrage", ancrage_cols)
        add_section("APP qui pourraient être proposés", app_prop_cols)

    # --- Finalisation ---
    if elements:
        try:
            doc.build(elements)
            buffer.seek(0)

            st.success("✅ PDF généré avec succès.")
            st.download_button("⬇️ Télécharger le PDF", data=buffer.getvalue(),
                               file_name="fiches_stagiaires.pdf", mime="application/pdf")
        except Exception as e:
             st.error(f"Une erreur est survenue lors de la construction du PDF : {e}")
    else:
         st.warning("Aucune donnée n'a été trouvée pour générer les fiches.")
