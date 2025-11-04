import streamlit as st
import pandas as pd
from io import BytesIO
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, PageBreak, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.lib.colors import HexColor

st.set_page_config(page_title="Fiches d’évaluation", page_icon="📘")
st.title("📘 Générateur de fiches d’évaluation")

uploaded_file = st.file_uploader("Importer un fichier Excel (.xlsx)", type=["xlsx"])

# --- Fonction de nettoyage Unicode ---
def nettoyer_texte(texte):
    if not isinstance(texte, str):
        return texte
    return ''.join(ch for ch in texte if ord(ch) < 127 or ch in "éèàùçÉÈÀÙÇ ,.;:!?()/-")

if uploaded_file:
    df = pd.read_excel(uploaded_file)
    st.success("✅ Fichier importé avec succès !")
    st.dataframe(df.head())

    # --- Normalisation des noms de colonnes ---
    def normaliser(texte):
        return (
            str(texte)
            .strip()
            .lower()
            .replace("é", "e")
            .replace("è", "e")
            .replace("ê", "e")
            .replace("-", " ")
            .replace("_", " ")
        )

    df.columns = [normaliser(c) for c in df.columns]

    # --- Recherche colonnes principales ---
    prenom_col = next((c for c in df.columns if "prenom" in c), None)
    nom_col = next((c for c in df.columns if "nom" in c and "stagiaire" not in c and "prenom" not in c), None)
    stagiaire_col = next((c for c in df.columns if "stagiaire" in c or "participant" in c or "eleve" in c), None)
    date_col = next((c for c in df.columns if "date" in c), None)

    # --- Création colonne formateur ---
    df["formateur"] = ""
    if prenom_col and nom_col:
        df["formateur"] = df[prenom_col].fillna("") + " " + df[nom_col].fillna("")

    # --- Masquage des colonnes inutiles ---
    mots_cles_a_masquer = [
        "email", "e mail", "organisation", "departement", "jcmsplugin",
        "temps", "taux", "score", "tentative", "reussite", "question", "nom"
    ]
    colonnes_utiles = [c for c in df.columns if not any(m in c for m in mots_cles_a_masquer)]
    df = df[colonnes_utiles]

    # --- Détection intelligente des sections ---
    def contient_mot(c, *mots):
        c = c.lower()
        return any(m in c for m in mots)

    app_non_evalues_cols = [c for c in df.columns if contient_mot(c, "non soumis", "non evalue")]
    app_evalues_cols = [c for c in df.columns if contient_mot(c, "app evalue", "app evalué", "app evaluee")]
    axe_prog_cols = [c for c in df.columns if contient_mot(c, "axe", "progression", "amelioration")]
    points_ancrage_cols = [c for c in df.columns if contient_mot(c, "ancrage", "point fort", "reussi")]
    app_proposes_cols = [c for c in df.columns if contient_mot(c, "propose", "proposition", "a proposer")]

    # --- Tri ---
    if date_col:
        df = df.sort_values(by=[stagiaire_col, date_col])

    groupes_stagiaires = df.groupby(stagiaire_col)

def coloriser_valeur(val):
    if not isinstance(val, str):
        return str(val)

    val = val.strip().upper()
    if val == "FAIT":
        return f"<font color='#007A33'><b>{val}</b></font>"  # vert foncé
    elif val == "A":
        return f"<font color='#00B050'><b>{val}</b></font>"  # vert clair
    elif val == "EN COURS":
        return f"<font color='#FFD700'><b>{val}</b></font>"  # jaune
    elif val == "ECA":
        return f"<font color='#ED7D31'><b>{val}</b></font>"  # orange
    elif val == "NE":
        return f"<font color='#808080'><b>{val}</b></font>"  # gris
    elif val == "NA":
        return f"<font color='#C00000'><b>{val}</b></font>"  # rouge
    else:
        return val   
        
 # --- Génération du PDF ---
if st.button("📄 Générer les fiches PDF"):
    from io import BytesIO
    from reportlab.lib.pagesizes import A4
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, PageBreak
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.enums import TA_CENTER
    from reportlab.lib.colors import HexColor

    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4)
    styles = getSampleStyleSheet()

    # --- Styles personnalisés ---
    titre_style = ParagraphStyle(
        "TitrePrincipal",
        parent=styles["Title"],
        alignment=TA_CENTER,
        textColor=HexColor("#003366"),
    )
    sous_titre_style = ParagraphStyle(
        "SousTitre",
        parent=styles["Heading2"],
        textColor=HexColor("#006699"),
    )
    champ_style = ParagraphStyle("Champ", parent=styles["Normal"], spaceAfter=6)
    section_style = ParagraphStyle(
        "Section",
        parent=styles["Heading3"],
        textColor=HexColor("#004C99"),
        spaceBefore=12,
        spaceAfter=6,
    )
    contenu_style = ParagraphStyle(
        "Contenu",
        parent=styles["Normal"],
        leftIndent=12,
        spaceAfter=4,
        fontName="Helvetica",
        fontSize=10,
        textColor=HexColor("#000000"),  # par défaut noir
    )

    # --- Fonction coloration ---
    def coloriser_valeur(val):
        """Retourne le texte coloré selon la valeur d'évaluation."""
        if not isinstance(val, str):
            return str(val)
        val = val.strip().upper()
        couleurs = {
            "FAIT": "#007A33",     # vert foncé
            "A": "#00B050",        # vert clair
            "EN COURS": "#FFD700", # jaune
            "ECA": "#ED7D31",      # orange
            "NE": "#808080",       # gris
            "NA": "#C00000",       # rouge
        }
        couleur = couleurs.get(val)
        if couleur:
            return f'<b><font color="{couleur}">{val}</font></b>'
        return val

    elements = []
elements.append(Paragraph("Test <font color='#FF0000'><b>rouge</b></font>", contenu_style))
for stagiaire, data_stagiaire in groupes_stagiaires:
        elements.append(Paragraph("📘 Fiche d’évaluation", titre_style))
        elements.append(Spacer(1, 12))
        elements.append(Paragraph(f"<b>Stagiaire évalué :</b> {stagiaire}", sous_titre_style))
        elements.append(Spacer(1, 8))

        for _, ligne in data_stagiaire.iterrows():
            # --- Infos générales ---
            if date_col and pd.notna(ligne.get(date_col)):
                elements.append(Paragraph(f"<b>Date d’évaluation :</b> {ligne[date_col]}", champ_style))
            if ligne.get("formateur"):
                elements.append(Paragraph(f"<b>Formateur :</b> {ligne['formateur']}", champ_style))
            elements.append(Spacer(1, 10))

            # --- Section : APP non soumis à évaluation ---
            if app_non_evalues_cols:
                elements.append(Paragraph("🟡 APP non soumis à évaluation", section_style))
                for c in app_non_evalues_cols:
                    val = ligne.get(c)
                    if pd.notna(val):
                        texte_val = coloriser_valeur(val)
                        nom_app = c.split("/")[-1].strip().capitalize()
                        texte_html = f"• {nom_app} : {texte_val}"
                        elements.append(Paragraph(texte_html, contenu_style))
                elements.append(Spacer(1, 8))

            # --- Section : APP évalués ---
            if app_evalues_cols:
                elements.append(Paragraph("🟢 APP évalués", section_style))
                for c in app_evalues_cols:
                    val = ligne.get(c)
                    if pd.notna(val):
                        texte_val = coloriser_valeur(val)
                        nom_app = c.split("/")[-1].strip().capitalize()
                        texte_html = f"• {nom_app} : {texte_val}"
                        elements.append(Paragraph(texte_html, contenu_style))
                elements.append(Spacer(1, 8))

            # --- Section : Axes de progression ---
            if axe_prog_cols:
                elements.append(Paragraph("🔵 Axes de progression", section_style))
                for c in axe_prog_cols:
                    val = ligne.get(c)
                    if pd.notna(val):
                        elements.append(Paragraph(f"• {val}", contenu_style))
                elements.append(Spacer(1, 8))

            # --- Section : Points d’ancrage ---
            if points_ancrage_cols:
                elements.append(Paragraph("🟠 Points d’ancrage", section_style))
                for c in points_ancrage_cols:
                    val = ligne.get(c)
                    if pd.notna(val):
                        elements.append(Paragraph(f"• {val}", contenu_style))
                elements.append(Spacer(1, 8))

            # --- Section : APP qui pourraient être proposés ---
            if app_proposes_cols:
                elements.append(Paragraph("🟣 APP qui pourraient être proposés", section_style))
                for c in app_proposes_cols:
                    val = ligne.get(c)
                    if pd.notna(val):
                        elements.append(Paragraph(f"• {val}", contenu_style))
                elements.append(Spacer(1, 8))

            # --- Séparation entre évaluations ---
            elements.append(Spacer(1, 10))
            elements.append(Paragraph("<hr width='100%' color='#CCCCCC'/>", styles["Normal"]))
            elements.append(PageBreak())

    doc.build(elements)
    buffer.seek(0)

    st.download_button(
        label="⬇️ Télécharger les fiches PDF",
        data=buffer,
        file_name="fiches_evaluations.pdf",
        mime="application/pdf",
    )
