import streamlit as st
import pandas as pd
from io import BytesIO
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.colors import HexColor

# --- Configuration de la page ---
st.set_page_config(page_title="Fiches d’évaluation", page_icon="📘")
st.title("📘 Générateur de fiches d’évaluation")

# --- Import du fichier Excel ---
uploaded_file = st.file_uploader("Importer un fichier Excel (.xlsx)", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)
    st.success("✅ Fichier importé avec succès !")
    st.dataframe(df.head())

   # --- Normalisation propre des colonnes ---
    def nettoyer_colonne(c):
        import unicodedata
        # Supprimer les accents et espaces parasites
        c = ''.join(
            ch for ch in unicodedata.normalize('NFKD', c)
            if not unicodedata.combining(ch)
        )
        return c.strip().lower().replace(" ", "_")

    df.columns = [nettoyer_colonne(c) for c in df.columns]

    st.write("🔍 Colonnes importées :", df.columns.tolist())

   # --- Recherche intelligente des colonnes ---
    prenom_col = next((c for c in df.columns if "prenom" in c and "stagiaire" not in c), None)
    nom_col = next((c for c in df.columns if "nom" in c and "stagiaire" not in c and "prenom" not in c), None)
    stagiaire_col = next((c for c in df.columns if any(x in c for x in ["stagiaire", "participant", "eleve"])), None)
    date_col = next((c for c in df.columns if "date" in c), None)

    st.write(f"🧾 Colonnes détectées → prenom: {prenom_col}, nom: {nom_col}, stagiaire: {stagiaire_col}, date: {date_col}")

    # --- Création du champ formateur ---
    if prenom_col and nom_col:
        df["formateur"] = df[prenom_col].astype(str).str.strip() + " " + df[nom_col].astype(str).str.strip()
    else:
        st.warning("⚠️ Colonnes 'prenom' et/ou 'nom' introuvables — le champ 'formateur' sera laissé vide.")
        df["formateur"] = ""
        
    # --- Colonnes à masquer ---
    colonnes_a_masquer = [
        "email", "organisation", "departement", "jcmsplugin",
        "temps", "taux", "score", "tentative", "reussite", "nombre de questions", "nom"
    ]

    prenom_col = next((c for c in df.columns if "prenom" in c), None)
    nom_col = next((c for c in df.columns if "nom" in c and "stagiaire" not in c), None)
    stagiaire_col = next((c for c in df.columns if "stagiaire" in c or "participant" in c), None)
    date_col = next((c for c in df.columns if "date" in c), None)

    if not stagiaire_col:
        st.error("❌ Impossible de trouver la colonne du stagiaire évalué.")
        st.stop()

    # --- Nettoyage du DataFrame ---
    colonnes_utiles = [c for c in df.columns if all(x not in c for x in colonnes_a_masquer)]
    df = df[colonnes_utiles]

    # --- Création colonne 'formateur' ---
    if prenom_col and nom_col and prenom_col in df.columns and nom_col in df.columns:
        df["formateur"] = df[prenom_col].astype(str) + " " + df[nom_col].astype(str)
    else:
        # Si on ne trouve pas les colonnes, on crée une colonne vide
        st.warning("⚠️ Colonnes 'prenom' et/ou 'nom' introuvables — le champ 'formateur' sera laissé vide.")
        df["formateur"] = ""
        
    # --- Tri des données ---
    if date_col:
        df = df.sort_values(by=[stagiaire_col, date_col])

    # --- Groupement par stagiaire ---
    groupes_stagiaires = df.groupby(stagiaire_col)

    # --- Bouton pour générer le PDF ---
    if st.button("📄 Générer les fiches PDF"):

        # --- Fonction coloration conditionnelle ---
        def coloriser_valeur(val):
            """Retourne le texte coloré selon la valeur d'évaluation."""
            if not isinstance(val, str):
                return str(val)
            val = val.strip().upper()
            couleurs = {
                "FAIT": "#007A33",
                "A": "#00B050",
                "EN COURS": "#FFD700",
                "ECA": "#ED7D31",
                "NE": "#808080",
                "NA": "#C00000",
            }
            couleur = couleurs.get(val)
            if couleur:
                return f'<b><font color="{couleur}">{val}</font></b>'
            return val

        # --- Initialisation du document ---
        buffer = BytesIO()
        doc = SimpleDocTemplate(buffer, pagesize=A4)
        styles = getSampleStyleSheet()

        # --- Styles personnalisés ---
        titre_style = ParagraphStyle("TitrePrincipal", parent=styles["Title"], alignment=TA_CENTER, textColor=HexColor("#003366"))
        sous_titre_style = ParagraphStyle("SousTitre", parent=styles["Heading2"], textColor=HexColor("#006699"))
        champ_style = ParagraphStyle("Champ", parent=styles["Normal"], spaceAfter=6)
        section_style = ParagraphStyle("Section", parent=styles["Heading3"], textColor=HexColor("#004C99"), spaceBefore=12, spaceAfter=6)
        contenu_style = ParagraphStyle("Contenu", parent=styles["Normal"], leftIndent=12, spaceAfter=4, fontName="Helvetica", fontSize=10)

        # --- Création des éléments PDF ---
        elements = []

        # 🔴 Test de rendu couleur (doit être rouge)
        elements.append(Paragraph("Test <font color='#FF0000'><b>rouge</b></font>", contenu_style))
        elements.append(Spacer(1, 12))

        for stagiaire, data_stagiaire in groupes_stagiaires:
            elements.append(Paragraph("📘 Fiche d’évaluation", titre_style))
            elements.append(Spacer(1, 12))
            elements.append(Paragraph(f"<b>Stagiaire évalué :</b> {stagiaire}", sous_titre_style))
            elements.append(Spacer(1, 8))

            for _, ligne in data_stagiaire.iterrows():
                if date_col and pd.notna(ligne.get(date_col)):
                    elements.append(Paragraph(f"<b>Évaluation du :</b> {ligne[date_col]}", champ_style))
                if ligne.get("formateur"):
                    elements.append(Paragraph(f"<b>Formateur :</b> {ligne['formateur']}", champ_style))
                elements.append(Spacer(1, 10))

                # Exemple d'une section
                elements.append(Paragraph("🟢 APP évalués", section_style))
                for col, val in ligne.items():
                    if pd.notna(val) and "app eval" in col:
                        texte_val = coloriser_valeur(str(val))
                        texte = f"• {col.split('/')[-1].capitalize()} : {texte_val}"
                        elements.append(Paragraph(texte, contenu_style))
                elements.append(Spacer(1, 8))

            elements.append(PageBreak())

        # --- Génération du PDF ---
        doc.build(elements)
        buffer.seek(0)

        # --- Téléchargement ---
        st.download_button(
            label="⬇️ Télécharger les fiches PDF",
            data=buffer,
            file_name="fiches_evaluations.pdf",
            mime="application/pdf"
        )

else:
    st.info("📂 En attente du fichier Excel (.xlsx) à importer.")
