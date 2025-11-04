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

    # Normalisation des colonnes
    df.columns = [c.strip().lower() for c in df.columns]

    # Colonnes inutiles à masquer
    colonnes_a_masquer = [
        "email", "e-mail", "organisation", "département",
        "jcmsplugin", "temps écoulé", "taux de réussite", "score",
        "tentative", "réussite", "nombre de questions"
    ]

    # Détection automatique des noms utiles
    nom_cols = [c for c in df.columns if "nom" in c]
    prenom_cols = [c for c in df.columns if "prenom" in c or "prénom" in c]
    stagiaire_cols = [c for c in df.columns if "stagia]()_
