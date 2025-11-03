import streamlit as st
import pandas as pd
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle

st.set_page_config(page_title="Convertisseur XLS → PDF", layout="centered")

st.title("📄 Convertisseur Excel → PDF")
uploaded_file = st.file_uploader("📂 Importer un fichier Excel (.xlsx)", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)
    st.success("Fichier importé avec succès !")
    st.dataframe(df.head())

    colonne_tri = st.selectbox("Trier par :", df.columns)
    df_sorted = df.sort_values(by=colonne_tri)

    if st.button("Générer le PDF"):
        pdf_path = "rapport.pdf"
        doc = SimpleDocTemplate(pdf_path, pagesize=A4)
        data = [df_sorted.columns.tolist()] + df_sorted.values.tolist()
        table = Table(data)
        table.setStyle(TableStyle([
            ("GRID", (0,0), (-1,-1), 0.5, colors.grey),
            ("BACKGROUND", (0,0), (-1,0), colors.lightgrey),
        ]))
        doc.build([table])

        with open(pdf_path, "rb") as f:
            st.download_button("⬇️ Télécharger le PDF", f, file_name="rapport.pdf")
