import streamlit as st
import pandas as pd
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import cm
from reportlab.platypus import SimpleDocTemplate, Paragraph, Table, TableStyle, Spacer, PageBreak
from io import BytesIO
import unicodedata

# Fonction de nettoyage du texte
def nettoyer_texte_visible(texte):
    if pd.isna(texte):
        return ""
    texte = str(texte)
    texte = "".join(c for c in unicodedata.normalize("NFKD", texte) if not unicodedata.combining(c))
    texte = texte.replace("_", " ").replace("’", "'").replace("•", "-")
    return texte.strip()

# Fonction de détection de couleur selon le code d’évaluation
def coloriser_valeur(valeur):
    if not isinstance(valeur, str):
        return colors.black
    valeur = valeur.lower().strip()
    if "fait" in valeur or valeur == "a":
        return colors.green
    elif "en cours" in valeur:
        return colors.yellow
    elif "eca" in valeur:
        return colors.orange
    elif "ne" in valeur:
        return colors.grey
    elif "na" in valeur:
        return colors.red
    return colors.black

# Interface Streamlit
st.title("📊 Générateur de synthèse d'évaluations (PDF)")
st.write("Chargez un fichier `.xlsx` contenant les évaluations des stagiaires.")

uploaded_file = st.file_uploader("Importer le fichier Excel", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)

    # Normalisation des noms de colonnes
    df.columns = [c.strip().lower().replace(" ", "_") for c in df.columns]

    # Détection des colonnes clés
    prenom_col = next((c for c in df.columns if "prenom" in c), None)
    nom_col = next((c for c in df.columns if "nom" in c), None)
    stagiaire_col = next((c for c in df.columns if "stagiaire" in c), None)
    date_col = next((c for c in df.columns if "date" in c), None)

    if not all([prenom_col, nom_col, stagiaire_col, date_col]):
        st.error("Impossible de détecter les colonnes essentielles (Prénom, Nom, Stagiaire, Date).")
    else:
        st.success("✅ Colonnes détectées avec succès !")

        # Ajout du formateur
        df["formateur"] = df[prenom_col].astype(str) + " " + df[nom_col].astype(str)

        # Tri alphabétique stagiaires + chronologique
        df[date_col] = pd.to_datetime(df[date_col], errors='coerce')
        df = df.sort_values(by=[stagiaire_col, date_col])

        # Génération du PDF
        buffer = BytesIO()
        doc = SimpleDocTemplate(buffer, pagesize=A4, leftMargin=1.5*cm, rightMargin=1.5*cm, topMargin=1.5*cm)

        elements = []
        titre_style = ParagraphStyle(name="Titre", fontSize=16, leading=20, textColor=colors.darkblue, spaceAfter=10)
        sous_titre_style = ParagraphStyle(name="SousTitre", fontSize=12, leading=15, textColor=colors.black)
        texte_style = ParagraphStyle(name="Texte", fontSize=10, leading=14)

        stagiaires = df[stagiaire_col].dropna().unique()

        for stagiaire in stagiaires:
            sous_df = df[df[stagiaire_col] == stagiaire]
            dates = sous_df[date_col].dropna().unique()

            elements.append(Paragraph(f"<b>Stagiaire :</b> {nettoyer_texte_visible(stagiaire)}", titre_style))
            elements.append(Spacer(1, 6))

            for d in sorted(dates):
                evals = sous_df[sous_df[date_col] == d]

                formateur = evals["formateur"].iloc[0] if "formateur" in evals else ""
                elements.append(Paragraph(f"<b>Date :</b> {d.strftime('%d/%m/%Y')} — <b>Formateur :</b> {formateur}", sous_titre_style))
                elements.append(Spacer(1, 4))

                # Sélection des colonnes d’évaluation non vides
                eval_cols = [c for c in evals.columns if "app_" in c or "evaluation_des" in c or "test_de" in c]

                # Groupement par type de séquence
                groupes = {
                    "APP soumis à évaluation": [c for c in eval_cols if "app_evalue" in c],
                    "APP non soumis à évaluation": [c for c in eval_cols if "app_non" in c],
                    "Évaluation des MSP": [c for c in eval_cols if "evaluation_des" in c],
                    "Test de nage": [c for c in eval_cols if "test_de" in c],
                }

                for type_seq, cols in groupes.items():
                    if not any(col in evals.columns for col in cols):
                        continue

                    # Bande colorée type séquence
                    bande = Table([[type_seq]], colWidths=[16*cm])
                    bande.setStyle(TableStyle([
                        ('BACKGROUND', (0,0), (-1,-1), colors.HexColor("#003366")),
                        ('TEXTCOLOR', (0,0), (-1,-1), colors.white),
                        ('ALIGN', (0,0), (-1,-1), 'LEFT'),
                        ('FONTNAME', (0,0), (-1,-1), 'Helvetica-Bold'),
                        ('FONTSIZE', (0,0), (-1,-1), 11),
                        ('BOTTOMPADDING', (0,0), (-1,-1), 6),
                        ('TOPPADDING', (0,0), (-1,-1), 4),
                    ]))
                    elements.append(bande)
                    elements.append(Spacer(1, 4))

                    # Données du tableau
                    data = [["Élément évalué", "Résultat"]]
                    for col in cols:
                        val = nettoyer_texte_visible(evals[col].iloc[0])
                        if val:
                            color = coloriser_valeur(val)
                            data.append([nettoyer_texte_visible(col), Paragraph(f"<font color='{color.hexval}'>{val}</font>", texte_style)])

                    # Construction du tableau
                    t = Table(data, colWidths=[10*cm, 6*cm])
                    t.setStyle(TableStyle([
                        ('BACKGROUND', (0,0), (-1,0), colors.lightblue),
                        ('TEXTCOLOR', (0,0), (-1,0), colors.black),
                        ('GRID', (0,0), (-1,-1), 0.25, colors.grey),
                        ('ALIGN', (0,0), (-1,-1), 'LEFT'),
                        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
                        ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
                    ]))
                    elements.append(t)
                    elements.append(Spacer(1, 10))

                # Paragraphes d’analyse
                for champ in ["axes_de_progression", "point_d'ancrage_(_ce_qu'il_fait_bien_naturellement)", "app_qui_pourrait_etre_propose"]:
                    if champ in evals.columns and pd.notna(evals[champ].iloc[0]):
                        titre = champ.split("_")[0].capitalize().replace("app", "APP")
                        texte = nettoyer_texte_visible(evals[champ].iloc[0])
                        elements.append(Paragraph(f"<b>{titre} :</b> {texte}", texte_style))
                        elements.append(Spacer(1, 6))

            elements.append(PageBreak())

        doc.build(elements)
        st.success("✅ PDF généré avec succès !")
        st.download_button("📄 Télécharger le PDF", data=buffer.getvalue(), file_name="synthese_evaluations.pdf", mime="application/pdf")
