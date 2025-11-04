import streamlit as st
import pandas as pd
from io import BytesIO
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER

st.set_page_config(page_title="Fiches d’évaluation", page_icon="📘")
st.title("📘 Générateur de fiches d’évaluation")

uploaded_file = st.file_uploader("Importer un fichier Excel (.xlsx)", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)
    st.success("✅ Fichier importé avec succès !")
    st.dataframe(df.head())

    # --- Normalisation des noms de colonnes ---
    df.columns = [c.strip().lower().replace("é", "e").replace("è", "e").replace("ê", "e") for c in df.columns]

    # --- Colonnes à masquer du PDF ---
    colonnes_a_masquer = [
        "email", "organisation", "departement", "jcmsplugin", "temps", "taux",
        "score", "tentative", "reussite", "nombre de questions", "nom"  # "nom" masqué, utilisé ailleurs
    ]

    # --- Recherche intelligente des colonnes ---
    prenom_col = next((c for c in df.columns if "prenom" in c), None)
    nom_col = next((c for c in df.columns if "nom" in c and "stagiaire" not in c), None)
    stagiaire_col = next((c for c in df.columns if "stagiaire" in c or "participant" in c or "eleve" in c), None)
    date_col = next((c for c in df.columns if "date" in c), None)

    if not stagiaire_col:
        st.error("❌ Impossible de trouver la colonne du stagiaire évalué.")
        st.stop()

    # --- Nettoyage du DataFrame ---
    colonnes_utiles = [c for c in df.columns if all(x not in c for x in colonnes_a_masquer)]
    df = df[colonnes_utiles]

    # --- Suppression des lignes sans évaluation ---
    colonnes_eval = [c for c in df.columns if "eval" in c or "commentaire" in c or "observation" in c]
    if colonnes_eval:
        df = df.dropna(how="all", subset=colonnes_eval)

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
        titre_style = ParagraphStyle("TitrePrincipal", parent=styles["Title"], alignment=TA_CENTER, textColor="#003366")
        sous_titre_style = ParagraphStyle("SousTitre", parent=styles["Heading2"], textColor="#006699")
        champ_style = ParagraphStyle("Champ", parent=styles["Normal"], spaceAfter=6)

        elements = []

        for stagiaire, data_stagiaire in groupes_stagiaires:
            elements.append(Paragraph("📘 Fiche d’évaluation", titre_style))
            elements.append(Spacer(1, 12))
            elements.append(Paragraph(f"<b>Stagiaire évalué :</b> {stagiaire}", sous_titre_style))
            elements.append(Spacer(1, 8))

            for _, ligne in data_stagiaire.iterrows():
                # --- Date ---
                if date_col and date_col in ligne and pd.notna(ligne[date_col]):
                    elements.append(Paragraph(f"<b>Évaluation du :</b> {ligne[date_col]}", champ_style))

                # --- Formateur (nom + prénom) ---
                formateur = ""
                if prenom_col and prenom_col in ligne:
                    formateur += str(ligne[prenom_col]) + " "
                if nom_col and nom_col in ligne:
                    formateur += str(ligne[nom_col])
                if formateur.strip():
                    elements.append(Paragraph(f"<b>Formateur :</b> {formateur.strip()}", champ_style))

                elements.append(Spacer(1, 8))

                # --- Autres infos ---
                for col, val in ligne.items():
                    if pd.notna(val) and col not in [stagiaire_col, nom_col, prenom_col, date_col]:
                        col_affiche = col.capitalize().replace("_", " ")
                        elements.append(Paragraph(f"<b>{col_affiche} :</b> {val}", champ_style))

                elements.append(Spacer(1, 10))
                elements.append(Paragraph("<hr width='100%' color='#AAAAAA'/>", styles["Normal"]))
                elements.append(Spacer(1, 10))

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
