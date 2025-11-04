import streamlit as st
import pandas as pd
from io import BytesIO
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.lib.colors import HexColor

st.set_page_config(page_title="Fiches d’évaluation", page_icon="📘")
st.title("📘 Générateur de fiches d’évaluation")

uploaded_file = st.file_uploader("Importer un fichier Excel (.xlsx)", type=["xlsx"])

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

    # --- Recherche des colonnes principales ---
    prenom_col = next((c for c in df.columns if "prenom" in c), None)
    nom_col = next((c for c in df.columns if "nom" in c and "prenom" not in c and "stagiaire" not in c), None)
    stagiaire_col = next((c for c in df.columns if "stagiaire" in c or "participant" in c or "eleve" in c), None)
    date_col = next((c for c in df.columns if "date" in c), None)

    if not stagiaire_col:
        st.error("❌ Impossible de trouver la colonne du stagiaire évalué.")
        st.stop()

    # --- Création de la colonne 'formateur' ---
    df["formateur"] = ""
    if prenom_col and nom_col:
        df["formateur"] = df[prenom_col].fillna("") + " " + df[nom_col].fillna("")
    elif prenom_col:
        df["formateur"] = df[prenom_col].fillna("")
    elif nom_col:
        df["formateur"] = df[nom_col].fillna("")

    # --- Colonnes à masquer ---
    mots_cles_a_masquer = [
        "email", "e mail", "organisation", "departement", "jcmsplugin",
        "temps", "taux", "score", "tentative", "reussite", "question", "nombre de question", "nom"
    ]

    colonnes_utiles = [
        c for c in df.columns
        if not any(m in c for m in mots_cles_a_masquer)
    ]
    df = df[colonnes_utiles]

    # --- Groupes de colonnes par thématique ---
    app_non_evalues_cols = [c for c in df.columns if "non soumis" in c]
    app_evalues_cols = [c for c in df.columns if "app evalue" in c or "app évalué" in c]
    axe_prog_cols = [c for c in df.columns if "axes de progression" in c]
    points_ancrage_cols = [c for c in df.columns if "points d ancrage" in c or "ancrage" in c]
    app_proposes_cols = [c for c in df.columns if "app qui pourrait" in c or "propose" in c]

    # --- Tri des données ---
    if date_col:
        df = df.sort_values(by=[stagiaire_col, date_col])

    groupes_stagiaires = df.groupby(stagiaire_col)

    # --- Génération du PDF ---
    if st.button("📄 Générer les fiches PDF"):
        buffer = BytesIO()
        doc = SimpleDocTemplate(buffer, pagesize=A4)
        styles = getSampleStyleSheet()

        # --- Styles personnalisés ---
        titre_style = ParagraphStyle("TitrePrincipal", parent=styles["Title"], alignment=TA_CENTER, textColor=HexColor("#003366"))
        sous_titre_style = ParagraphStyle("SousTitre", parent=styles["Heading2"], textColor=HexColor("#006699"))
        champ_style = ParagraphStyle("Champ", parent=styles["Normal"], spaceAfter=6)
        section_style = ParagraphStyle("Section", parent=styles["Heading3"], textColor=HexColor("#004C99"), spaceBefore=12, spaceAfter=6, underlineWidth=0.5)
        contenu_style = ParagraphStyle("Contenu", parent=styles["Normal"], leftIndent=12, spaceAfter=4)

        elements = []

        for stagiaire, data_stagiaire in groupes_stagiaires:
            elements.append(Paragraph("📘 Fiche d’évaluation", titre_style))
            elements.append(Spacer(1, 12))
            elements.append(Paragraph(f"<b>Stagiaire évalué :</b> {stagiaire}", sous_titre_style))
            elements.append(Spacer(1, 8))

            for _, ligne in data_stagiaire.iterrows():
                # --- Informations générales ---
                if date_col and pd.notna(ligne.get(date_col)):
                    elements.append(Paragraph(f"<b>Évaluation du :</b> {ligne[date_col]}", champ_style))
                if ligne.get("formateur"):
                    elements.append(Paragraph(f"<b>Formateur :</b> {ligne['formateur']}", champ_style))
                elements.append(Spacer(1, 10))

                # --- Section : APP non soumis à évaluation ---
if app_non_evalues_cols:
    elements.append(Paragraph("🟡 APP non soumis à évaluation", section_style))
    for c in app_non_evalues_cols:
        val = ligne.get(c)
        if pd.notna(val):
            nom_app = c.split("/")[-1].strip().capitalize() if "/" in c else c.capitalize()
            elements.append(Paragraph(f"• {nom_app} : {val}", contenu_style))
    elements.append(Spacer(1, 8))

# --- Section : APP évalués ---
if app_evalues_cols:
    elements.append(Paragraph("🟢 APP évalués", section_style))
    for c in app_evalues_cols:
        val = ligne.get(c)
        if pd.notna(val):
            nom_app = c.split("/")[-1].strip().capitalize() if "/" in c else c.capitalize()
            elements.append(Paragraph(f"• {nom_app} : {val}", contenu_style))
    elements.append(Spacer(1, 8))

                # --- Section : Axe de progression ---
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

                # --- Séparation ---
                elements.append(Paragraph("<hr width='100%' color='#CCCCCC'/>", styles["Normal"]))
                elements.append(PageBreak())

        doc.build(elements)
        buffer.seek(0)

        st.download_button(
            label="⬇️ Télécharger les fiches PDF",
            data=buffer,
            file_name="fiches_evaluations.pdf",
            mime="application/pdf"
        )

else:
    st.info("📂 En attente du fichier Excel (.xlsx) à importer.")
