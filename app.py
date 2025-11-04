import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.styles import ParagraphStyle

st.set_page_config(page_title="Tracabilité XLS → PDF", layout="centered")

st.title("📘 Générateur de fiches d’évaluation")
st.write("Charge un fichier Excel pour créer un PDF clair et coloré, une fiche par stagiaire.")

uploaded_file = st.file_uploader("📂 Choisir le fichier Excel", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)
    st.success("✅ Fichier importé avec succès.")

    # 🔤 Normaliser les noms de colonnes
    def normaliser(n):
        return (
            str(n)
            .lower()
            .replace("é", "e")
            .replace("è", "e")
            .replace("ê", "e")
            .replace("à", "a")
            .replace("â", "a")
            .replace("ô", "o")
            .replace("ç", "c")
            .replace("ï", "i")
            .replace("î", "i")
            .replace(" ", "_")
        )

    df.columns = [normaliser(c) for c in df.columns]

    # 🔎 Colonnes détectées automatiquement
    stagiaire_col = next((c for c in df.columns if "stagiaire" in c), None)
    date_col = next((c for c in df.columns if "date" in c), None)
    formateur_col = next((c for c in df.columns if "formateur" in c), None)

    app_non_eval_cols = [c for c in df.columns if "app_non_soumis" in c]
    app_eval_cols = [c for c in df.columns if "app_evalue" in c]
    axes_cols = [c for c in df.columns if "axe" in c]
    ancrage_cols = [c for c in df.columns if "ancrage" in c]
    app_prop_cols = [c for c in df.columns if "app_qui_pourrait" in c]

    # 🖋 Styles
    titre_style = ParagraphStyle(
        "Titre",
        fontSize=16,
        leading=20,
        alignment=1,
        textColor=colors.HexColor("#008000"),
        spaceAfter=12,
    )
    section_style = ParagraphStyle(
        "Section",
        fontSize=12,
        textColor=colors.HexColor("#003366"),
        leading=14,
        spaceBefore=8,
        spaceAfter=4,
    )
    texte_style = ParagraphStyle(
        "Texte",
        fontSize=10,
        leading=12,
        textColor=colors.black,
        spaceBefore=2,
        allowHTML=True,
    )

    # 🎨 Couleurs d'évaluation
    def coloriser(val):
        if pd.isna(val) or val == "":  
            return ""
        
        # Nettoyage et normalisation de la valeur
        # L'utilisation de str(val) est nécessaire au cas où la valeur n'est pas déjà une chaîne de caractères
        val_normalisee = str(val).strip().upper().replace(".", "")
        
        # Définition des couleurs exactes demandées
        couleurs = {
            # "fait" en vert foncé
            "FAIT": colors.HexColor("#00B050"), 
            # "A" en vert clair
            "A": colors.HexColor("#32CD32"), 
            # "en cours" en jaune
            "EN COURS": colors.HexColor("#FFD700"), 
            # "NE" en gris
            "NE": colors.HexColor("#808080"), 
            # "NA" en rouge
            "NA": colors.HexColor("#C00000"), 
            # "ECA" en orange
            "ECA": colors.HexColor("#FF8C00"), 
        }
        
        c = couleurs.get(val_normalisee)
        
        if c:
            # Retourne la valeur formatée en HTML avec la couleur et en gras
            # On utilise la valeur originale (val) pour l'affichage, 
            # mais on s'assure qu'elle est un str pour l'insertion
            return f'<font color="{c.hexval()}"><b>{str(val)}</b></font>'
        
        # Si aucune correspondance, retourne la valeur d'origine en gras (sans couleur)
        return f"<b>{str(val)}</b>"


    # 📄 Création du PDF
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4)
    elements = []

    # Vérification si la colonne 'stagiaire' est trouvée
    if stagiaire_col is None:
        st.error("❌ Colonne 'stagiaire' non trouvée. Veuillez vous assurer que le nom de la colonne contient 'stagiaire'.")
    else:
        # Groupement par stagiaire
        for stagiaire, data_stagiaire in df.groupby(stagiaire_col):
            # Utilise la première ligne du groupe pour les métadonnées
            ligne = data_stagiaire.iloc[0] 

            # --- En-tête ---
            elements.append(Paragraph("Fiche d’évaluation", titre_style))
            elements.append(Spacer(1, 8))
            elements.append(Paragraph(f"<b>Stagiaire :</b> {stagiaire}", texte_style))
            
            # Récupération sécurisée des données de métadonnées
            date_info = ligne.get(date_col, '') if date_col in ligne else ''
            formateur_info = ligne.get(formateur_col, '') if formateur_col in ligne else ''
            
            elements.append(Paragraph(f"<b>Date :</b> {date_info}", texte_style))
            elements.append(Paragraph(f"<b>Formateur :</b> {formateur_info}", texte_style))
            elements.append(Spacer(1, 10))

            # --- APP non soumis ---
            if app_non_eval_cols:
                elements.append(Paragraph("APP non soumis à évaluation", section_style))
                for c in app_non_eval_cols:
                    # Remplacement plus précis
                    nom = c.replace("app_non_soumis_a_evaluation_/_", "").replace("app_non_soumis_", "").replace("_", " ").capitalize()
                    val = ligne.get(c, "")
                    if pd.notna(val) and str(val).strip() != "":
                        elements.append(Paragraph(f"• {nom} : {coloriser(val)}", texte_style))
                elements.append(Spacer(1, 8))

            # --- APP évalués ---
            if app_eval_cols:
                elements.append(Paragraph("APP évalués", section_style))
                for c in app_eval_cols:
                    # Remplacement plus précis
                    nom = c.replace("app_evalues_/_", "").replace("app_evalue_", "").replace("_", " ").capitalize()
                    val = ligne.get(c, "")
                    if pd.notna(val) and str(val).strip() != "":
                        elements.append(Paragraph(f"• {nom} : {coloriser(val)}", texte_style))
                elements.append(Spacer(1, 8))

            # --- Axes de progression ---
            if axes_cols:
                elements.append(Paragraph("Axes de progression", section_style))
                for c in axes_cols:
                    val = ligne.get(c, "")
                    if pd.notna(val) and str(val).strip() != "":
                        elements.append(Paragraph(str(val), texte_style))
                elements.append(Spacer(1, 8))

            # --- Points d’ancrage ---
            if ancrage_cols:
                elements.append(Paragraph("Points d’ancrage", section_style))
                for c in ancrage_cols:
                    val = ligne.get(c, "")
                    if pd.notna(val) and str(val).strip() != "":
                        elements.append(Paragraph(str(val), texte_style))
                elements.append(Spacer(1, 8))

            # --- APP proposés ---
            if app_prop_cols:
                elements.append(Paragraph("APP qui pourraient être proposés", section_style))
                for c in app_prop_cols:
                    val = ligne.get(c, "")
                    if pd.notna(val) and str(val).strip() != "":
                        elements.append(Paragraph(str(val), texte_style))
                elements.append(Spacer(1, 20))

    # --- Génération du PDF ---
    if stagiaire_col is not None:
        doc.build(elements)
        buffer.seek(0)

        st.download_button(
            label="📄 Télécharger le PDF des fiches",
            data=buffer,
            file_name="fiches_stagiaires.pdf",
            mime="application/pdf",
        )
