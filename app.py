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
    # évite que "prenom" soit aussi détecté comme nom
    nom_col = next((c for c in df.columns if "nom" in c and "prenom" not in c and "stagiaire" not in c), None)
    stagiaire_col = next((c for c in df.columns if "stagiaire" in c or "participant" in c or "eleve" in c), None)
    date_col = next((c for c in df.columns if "date" in c), None)

    if not stagiaire_col:
        st.error("❌ Impossible de trouver la colonne du stagiaire évalué.")
        st.stop()

    # --- Création de la colonne 'formateur' avant filtrage ---
    df["formateur"] = ""
    if prenom_col and nom_col:
        df["formateur"] = df[prenom_col].fillna("") + " " + df[nom_col].fillna("")
    elif prenom_col:  # fallback si le nom n’est pas trouvé
        df["formateur"] = df[prenom_col].fillna("")
    elif nom_col:
        df["formateur"] = df[nom_col].fillna("")

    # --- Colonnes à masquer du PDF ---
    mots_cles_a_masquer = [
        "email", "e mail",  # gère les deux versions
        "organisation", "departement", "jcmsplugin",
        "temps", "taux", "score", "tentative",
        "reussite", "question", "nombre de question",
        "nom",  # déjà inclus dans formateur
    ]

    colonnes_utiles = [
        c for c in df.columns
        if not any(m in c for m in mots_cles_a_masquer)
    ]
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

                # --- Formateur ---
                if "formateur" in ligne and isinstance(ligne["formateur"], str) and ligne["formateur"].strip():
                    elements.append(Paragraph(f"<b>Formateur :</b> {ligne['formateur'].strip()}", champ_style))

                elements.append(Spacer(1, 8))

                # --- Autres infos ---
                for col, val in ligne.items():
                    if pd.notna(val) and col not in [stagiaire_col, prenom_col, nom_col, date_col, "formateur"]:
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
