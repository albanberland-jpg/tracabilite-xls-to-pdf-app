import streamlit as st
import pandas as pd
from io import BytesIO
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, PageBreak
from reportlab.lib.styles import getSampleStyleSheet

# --- Interface principale ---
st.set_page_config(page_title="Fiches d’évaluation", page_icon="📘")
st.title("📘 Générateur de fiches d’évaluation")
st.write("Importe ton fichier Excel (export de l’application) et génère automatiquement une fiche PDF par stagiaire.")

# --- Upload du fichier Excel ---
uploaded_file = st.file_uploader("Importer un fichier Excel (.xlsx)", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)

    st.success("✅ Fichier importé avec succès !")
    st.dataframe(df.head())  # aperçu des premières lignes

    # --- Nettoyage du dataframe ---
    colonnes_eval = [c for c in df.columns if "APP" in c or "Évaluation" in c or "Evaluation" in c]
    if not colonnes_eval:
        st.warning("⚠️ Aucune colonne d'évaluation détectée automatiquement. Vérifie les noms de colonnes.")
    else:
        df = df.dropna(how='all', subset=colonnes_eval)

    # On trie par stagiaire + date si disponible
    if "Date" in df.columns:
        df = df.sort_values(by=["Nom du stagiaire", "Date"])
    else:
        df = df.sort_values(by=["Nom du stagiaire"])

    groupes_stagiaires = df.groupby("Nom du stagiaire")

    # --- Génération du PDF ---
if st.button("📄 Générer les fiches PDF"):
        buffer = BytesIO()
        doc = SimpleDocTemplate(buffer, pagesize=A4)
        styles = getSampleStyleSheet()
        elements = []

        titre_global = Paragraph("📘 Fiches d’évaluation des stagiaires", styles["Title"])
        elements.append(titre_global)
        elements.append(Spacer(1, 12))

        for nom_stagiaire, data_stagiaire in groupes_stagiaires:
            elements.append(Paragraph(f"<b>Stagiaire :</b> {nom_stagiaire}", styles["Heading2"]))
            elements.append(Spacer(1, 8))

            for _, ligne in data_stagiaire.iterrows():
                for col, val in ligne.items():
                    if pd.notna(val) and col != "Nom du stagiaire":
                        elements.append(Paragraph(f"<b>{col} :</b> {val}", styles["Normal"]))
                elements.append(Spacer(1, 8))
                elements.append(Paragraph("──────────────────────────────", styles["Normal"]))
