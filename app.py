import streamlit as st
import pandas as pd
from io import BytesIO
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, PageBreak
from reportlab.lib.styles import getSampleStyleSheet

st.set_page_config(page_title="Fiches d’évaluation", page_icon="📘")
st.title("📘 Générateur de fiches d’évaluation")

uploaded_file = st.file_uploader("Importer un fichier Excel (.xlsx)", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)
    st.success("✅ Fichier importé avec succès !")
    st.dataframe(df.head())

    # Normalisation des noms de colonnes (tout en minuscules)
    df.columns = [c.strip().lower() for c in df.columns]

    # Détection automatique du nom du stagiaire
    possible_nom_cols = [c for c in df.columns if "nom" in c]
    possible_prenom_cols = [c for c in df.columns if "prenom" in c or "prénom" in c]

    if possible_nom_cols:
        nom_col = possible_nom_cols[0]
    else:
        st.error("❌ Impossible de trouver une colonne 'Nom' dans ton fichier.")
        st.stop()

    prenom_col = possible_prenom_cols[0] if possible_prenom_cols else None

    # Création d’un identifiant complet du stagiaire
    if prenom_col:
        df["stagiaire"] = df[prenom_col].astype(str) + " " + df[nom_col].astype(str)
    else:
        df["stagiaire"] = df[nom_col].astype(str)

    # Tri et nettoyage
    if "date" in df.columns:
        df = df.sort_values(by=["stagiaire", "date"])
    else:
        df = df.sort_values(by=["stagiaire"])

    # Suppression des lignes vides (sans évaluation)
    colonnes_eval = [c for c in df.columns if "app" in c or "évalu" in c or "eval" in c]
    if colonnes_eval:
        df = df.dropna(how="all", subset=colonnes_eval)

    groupes_stagiaires = df.groupby("stagiaire")

    # --- Génération du PDF ---
    if st.button("📄 Générer les fiches PDF"):
        buffer = BytesIO()
        doc = SimpleDocTemplate(buffer, pagesize=A4)
        styles = getSampleStyleSheet()
        elements = []

        titre_global = Paragraph("📘 Fiches d’évaluation des stagiaires", styles["Title"])
        elements.append(titre_global)
        elements.append(Spacer(1, 12))

        for stagiaire, data_stagiaire in groupes_stagiaires:
            elements.append(Paragraph(f"<b>Stagiaire :</b> {stagiaire}", styles["Heading2"]))
            elements.append(Spacer(1, 8))

            for _, ligne in data_stagiaire.iterrows():
                for col, val in ligne.items():
                    if pd.notna(val) and col not in ["stagiaire"]:
                        elements.append(Paragraph(f"<b>{col.capitalize()} :</b> {val}", styles["Normal"]))
                elements.append(Spacer(1, 6))
                elements.append(Paragraph("──────────────────────────────", styles["Normal"]))

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
