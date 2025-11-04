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

    # --- Nettoyage
