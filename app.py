import streamlit as st
import pandas as pd
from io import BytesIO
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, PageBreak
from reportlab.lib.enums import TA_CENTER, TA_LEFT
import re

st.set_page_config(page_title="Traçabilité XLS → PDF", layout="centered")

st.title("📄 Générateur de fiches d’évaluation (XLS → PDF)")

uploaded_file = st.file_uploader("📤 Importer un fichier Excel (.xlsx)", type=["xlsx"])

# ----------------------------------------------------------
# 🎨 Fonction pour coloriser les valeurs selon leur état
# ----------------------------------------------------------
def coloriser_valeur(val):
    if not isinstance(val, str):
        val = str(val)
    val = val.strip().upper()

    if val == "FAIT":
        return f"<font color='#007A33'><b>{val}</b></font>"  # Vert foncé
    elif val == "A":
        return f"<font color='#00B050'><b>{val}</b></font>"  # Vert clair
    elif val == "EN COURS":
        return f"<font color='#FFD700'><b>{val}</b></font>"  # Jaune
    elif val == "ECA":
        return f"<font color='#ED7D31'><b>{val}</b></font>"  # Orange
    elif val == "NE":
        return f"<font color='#808080'><b>{val}</b></font>"  # Gris
    elif val == "NA":
        return f"<font color='#C00000'><b>{val}</b></font>"  # Rouge
    else:
        return val

# ----------------------------------------------------------
# 🧹 Nettoyage des intitulés
# ----------------------------------------------------------
def nettoyer_intitule(texte):
    if not isinstance(texte, str):
        return texte
    texte = re.sub(r"[_\-]+", " ", texte)  # supprime _ et -
    texte = re.sub(r"\s+", " ", texte)  # supprime les doubles espaces
    texte = texte.strip().capitalize()
    return texte

# ----------------------------------------------------------
# 📄 Génération du PDF
# ----------------------------------------------------------
def generer_pdf(df):
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4,
                            rightMargin=40, leftMargin=40, topMargin=60, bottomMargin=40)

    styles = getSampleStyleSheet()
    style_titre = ParagraphStyle(
        "Titre",
        parent=styles["Heading1"],
        alignment=TA_CENTER,
        textColor="#003366",
        spaceAfter=20,
    )
    style_soustitre = ParagraphStyle(
        "SousTitre",
        parent=styles["Heading2"],
        textColor="#003366",
        spaceAfter=10,
    )
    style_normal = ParagraphStyle(
        "Normal",
        parent=styles["BodyText"],
        alignment=TA_LEFT,
        spaceAfter=6,
        leading=15,
    )

    elements = []

    # Vérif colonnes
    colonnes = [c.lower() for c in df.columns]
    st.write("🔍 Colonnes importées :", colonnes)

    # Recherche auto des colonnes clés
    prenom_col = next((c for c in df.columns if "prenom" in c.lower()), None)
    nom_col = next((c for c in df.columns if "nom" in c.lower()), None)
    stagiaire_col = next((c for c in df.columns if "stagiaire" in c.lower()), None)
    date_col = next((c for c in df.columns if "date" in c.lower()), None)

    st.write(f"🧾 Colonnes détectées → prenom: {prenom_col}, nom: {nom_col}, stagiaire: {stagiaire_col}, date: {date_col}")

    # Ajout colonne formateur
    if prenom_col and nom_col:
        df["formateur"] = df[prenom_col].astype(str) + " " + df[nom_col].astype(str)
    else:
        st.warning("⚠️ Colonnes 'prenom' et/ou 'nom' introuvables — le champ 'formateur' sera laissé vide.")
        df["formateur"] = ""

    # Groupement par stagiaire
    if stagiaire_col:
        groupes_stagiaires = df.groupby(stagiaire_col)
    else:
        st.error("❌ Colonne 'stagiaire' introuvable dans le fichier.")
        return None

    # Test couleur
    elements.append(Paragraph("Test <font color='#FF0000'><b>rouge</b></font>", style_normal))
    elements.append(Spacer(1, 10))

    for stagiaire, data_stagiaire in groupes_stagiaires:
        elements.append(Paragraph("■ Fiche d’évaluation", style_titre))
        elements.append(Spacer(1, 10))

        formateur = data_stagiaire["formateur"].iloc[0]
        date_eval = data_stagiaire[date_col].iloc[0] if date_col else ""

        elements.append(Paragraph(f"<b>Stagiaire évalué :</b> {stagiaire}", style_normal))
        elements.append(Paragraph(f"<b>Évaluation du :</b> {date_eval}", style_normal))
        elements.append(Paragraph(f"<b>Formateur :</b> {formateur}", style_normal))
        elements.append(Spacer(1, 15))

        # Section 1
        elements.append(Paragraph("■ APP non soumis à évaluation", style_soustitre))
        for col in [c for c in df.columns if "non_soumis" in c.lower()]:
            val = coloriser_valeur(data_stagiaire[col].iloc[0])
            titre = nettoyer_intitule(col.split("/")[-1])
            elements.append(Paragraph(f"{titre} : {val}", style_normal))
        elements.append(Spacer(1, 10))

        # Section 2
        elements.append(Paragraph("■ APP évalués", style_soustitre))
        for col in [c for c in df.columns if "app_evalue" in c.lower()]:
            val = coloriser_valeur(data_stagiaire[col].iloc[0])
            titre = nettoyer_intitule(col.split("/")[-1])
            elements.append(Paragraph(f"{titre} : {val}", style_normal))
        elements.append(Spacer(1, 10))

        # Section 3
        elements.append(Paragraph("■ Axes de progression", style_soustitre))
        for col in [c for c in df.columns if "axe" in c.lower()]:
            val = data_stagiaire[col].iloc[0]
            titre = nettoyer_intitule(col.split("/")[-1])
            elements.append(Paragraph(f"{titre} : {val}", style_normal))

        elements.append(PageBreak())

    doc.build(elements)
    buffer.seek(0)
    return buffer

# ----------------------------------------------------------
# 🎯 Interface Streamlit
# ----------------------------------------------------------
if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file)
        st.success("✅ Fichier importé avec succès !")

        if st.button("📘 Générer le PDF"):
            pdf = generer_pdf(df)
            if pdf:
                st.download_button(
                    label="💾 Télécharger le PDF",
                    data=pdf,
                    file_name="fiches_evaluations.pdf",
                    mime="application/pdf",
                )

    except Exception as e:
        st.error(f"Erreur lors de la lecture du fichier : {e}")
