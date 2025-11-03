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
    from reportlab.lib.pagesizes import A4
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
    from reportlab.lib.styles import getSampleStyleSheet
    from reportlab.lib import colors

    pdf_path = "rapport.pdf"
    doc = SimpleDocTemplate(pdf_path, pagesize=A4)

    # Styles et titre
    styles = getSampleStyleSheet()
    titre = Paragraph("Rapport Excel → PDF", styles["Title"])

    # Conversion des données
    data = [df_sorted.columns.tolist()] + df_sorted.astype(str).values.tolist()

    # Table stylisée
    table = Table(data, repeatRows=1)
    table.setStyle(TableStyle([
        ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
        ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("BOTTOMPADDING", (0, 0), (-1, 0), 8),
    ]))

    # Construction du PDF
    elements = [titre, Spacer(1, 12), table]
    doc.build(elements)

    # Téléchargement
    with open(pdf_path, "rb") as f:
        st.download_button("⬇️ Télécharger le PDF", f, file_name="rapport.pdf")


    # Téléchargement
    with open(pdf_path, "rb") as f:
        st.download_button("⬇️ Télécharger le PDF", f, file_name="rapport.pdf")
