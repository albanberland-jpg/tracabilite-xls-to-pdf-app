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

    # Normalisation des noms de colonnes
    df.columns = [c.strip().lower() for c in df.columns]

    # Colonnes à masquer automatiquement
    colonnes_a_masquer = [
        "email", "e-mail", "organisation", "département", "jcmsplugin",
        "temps écoulé", "taux de réussite", "score", "tentative",
        "réussite", "nombre de questions", "nom"  # "nom" supprimé (mais utilisé pour formateur)
    ]

    # Recherche automatique des colonnes clés
    nom_cols = [c for c in df.columns if "nom" in c]
    prenom_cols = [c for c in df.columns if "prenom" in c or "prénom" in c]
    stagiaire_cols = [c for c in df.columns if "stagiaire" in c or "élève" in c or "participant" in c]
    date_cols = [c for c in df.columns if "date" in c]

    nom_col = nom_cols[0] if nom_cols else None
    prenom_col = prenom_cols[0] if prenom_cols else None
    stagiaire_col = stagiaire_cols[0] if stagiaire_cols else None
    date_col = date_cols[0] if date_cols else None

    if not stagiaire_col:
        st.error("❌ Impossible de trouver la colonne correspondant au stagiaire évalué.")
        st.stop()

    # Retirer les colonnes inutiles du DataFrame
    colonnes_utiles = [c for c in df.columns if all(x not in c for x in colonnes_a_masquer)]
    df = df[colonnes_utiles]

    # Supprimer les lignes sans évaluation
    colonnes_eval = [c for c in df.columns if "eval" in c or "commentaire" in c or "observation" in c]
    if colonnes_eval:
        df = df.dropna(how="all", subset=colonnes_eval)

    # Trier les données
    if date_col:
        df = df.sort_values(by=[stagiaire_col, date_col])

    groupes_stagiaires = df.groupby(stagiaire_col)

    # --- Génération du PDF ---
    if st.button("📄 Générer les fiches PDF"):
        buffer = BytesIO()
        doc = SimpleDocTemplate(buffer, pagesize=A4)
        styles = getSampleStyleSheet()

        # Styles personnalisés
        titre_style = ParagraphStyle(
            "TitrePrincipal",
            parent=styles["Title"],
            alignment=TA_CENTER,
            textColor="#003366"
        )
        sous_titre_style = ParagraphStyle(
            "SousTitre",
            parent=styles["Heading2"],
            textColor="#006699"
        )
        champ_style = ParagraphStyle(
            "Champ",
            parent=styles["Normal"],
            spaceAfter=6
        )

        elements = []

        for stagiaire, data_stagiaire in groupes_stagiaires:
            # --- Titre principal ---
            elements.append(Paragraph("📘 Fiche d’évaluation", titre_style))
            elements.append(Spacer(1, 12))

            # Stagiaire évalué
            elements.append(Paragraph(f"<b>Stagiaire évalué :</b> {stagiaire}", sous_titre_style))
            elements.append(Spacer(1, 8))

            for _, ligne in data_stagiaire.iterrows():
                # Date
                if date_col and pd.notna(ligne[date_col]):
                    elements.append(Paragraph(f"<b>Évaluation du :</b> {ligne[date_col]}", champ_style))

                # Formateur (composé de prénom + nom)
                if nom_col and prenom_col:
                    formateur = f"{ligne[prenom_col]} {ligne[nom_col]}"
                    elements.append(Paragraph(f"<b>Formateur :</b> {formateur}", champ_style))
                elif prenom_col:
                    elements.append(Paragraph(f"<b>Formateur :</b> {ligne[prenom_col]}", champ_style))

                elements.append(Spacer(1, 8))

                # Autres informations d'évaluation
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
