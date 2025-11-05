import streamlit as st
import pandas as pd
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape, portrait
from reportlab.lib.styles import ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Table, TableStyle, Spacer, PageBreak
from io import BytesIO

st.set_page_config(page_title="Synthèse des évaluations", layout="centered")

st.title("📘 Générateur de fiches de synthèse stagiaires")

uploaded_file = st.file_uploader("Téléverser le fichier XLSX", type=["xlsx"])

# --- Nettoyage du texte ---
def nettoyer_texte_visible(texte):
    if not isinstance(texte, str):
        texte = str(texte)
    texte = texte.replace("_", " ").replace("’", "'").replace("é", "e").replace("à", "a")
    texte = texte.replace("–", "-").replace("œ", "oe")
    texte = "".join(ch for ch in texte if ch.isprintable())
    return texte.strip()

# --- Couleur du texte selon la valeur ---
def coloriser_texte(valeur):
    v = str(valeur).strip().lower()
    couleurs = {
        "fait": "#008000",      # vert
        "a": "#008000",         # vert
        "eca": "#FFA500",       # orange
        "en cours": "#FFD700",  # jaune
        "ne": "#808080",        # gris
        "na": "#FF0000"         # rouge
    }
    for mot, couleur in couleurs.items():
        if mot in v:
            return f'<font color="{couleur}"><b>{valeur}</b></font>'
    return valeur

if uploaded_file:
    df = pd.read_excel(uploaded_file)
    df.columns = [str(c).strip().lower().replace(" ", "_") for c in df.columns]

    # Colonnes essentielles
    prenom_col = next((c for c in df.columns if "prenom" in c), None)
    nom_col = next((c for c in df.columns if "nom" in c), None)
    stagiaire_col = next((c for c in df.columns if "stagiaire" in c), None)
    date_col = next((c for c in df.columns if "date" in c), None)

    if not all([prenom_col, nom_col, stagiaire_col, date_col]):
        st.error("❌ Colonnes essentielles non détectées (Prénom, Nom, Stagiaire évalué, Date).")
    else:
        # Création du champ "formateur"
        df["formateur"] = df[prenom_col].astype(str) + " " + df[nom_col].astype(str)

        buffer = BytesIO()
        doc = SimpleDocTemplate(buffer, pagesize=A4,
                                rightMargin=30, leftMargin=30, topMargin=40, bottomMargin=30)

        elements = []
        normal_style = ParagraphStyle("Normal", fontSize=10, spaceAfter=6)
        title_style = ParagraphStyle("Titre", fontSize=14, spaceAfter=12, textColor="#053b74")
        bande_style = ParagraphStyle("Bande", fontSize=11, textColor="white", alignment=1, backColor="#053b74")

        # Groupement stagiaire > date
        for stagiaire, data_stagiaire in sorted(df.groupby(stagiaire_col)):
            elements.append(Paragraph(f"<b>Stagiaire :</b> {nettoyer_texte_visible(stagiaire)}", title_style))

            for date_eval, data_date in sorted(data_stagiaire.groupby(date_col)):
                formateur = data_date["formateur"].iloc[0]

                # Bandeau date + formateur
                elements.append(Paragraph(f"<b>Date :</b> {pd.to_datetime(date_eval).strftime('%d/%m/%Y')} &nbsp;&nbsp;&nbsp; "
                                          f"<b>Formateur :</b> {formateur}", normal_style))
                elements.append(Spacer(1, 6))

                # Regrouper les colonnes par section
                sections = {
                    "APP non soumis à évaluation": [c for c in df.columns if "app_non" in c],
                    "APP soumis à évaluation": [c for c in df.columns if "app_evalue" in c],
                    "MSP": [c for c in df.columns if "evaluation_des_msp" in c]
                }

                for titre_section, colonnes in sections.items():
                    if not colonnes:
                        continue

                    # Titre de section (bande bleue foncée)
                    elements.append(Paragraph(titre_section.upper(), bande_style))

                    # Tableau des résultats
                    data_table = [["Compétence", "Évaluation"]]
                    for col in colonnes:
                        valeur = data_date[col].iloc[0]
                        if pd.notna(valeur) and str(valeur).strip():
                            data_table.append([
                                nettoyer_texte_visible(col.split("/")[-1]).capitalize(),
                                coloriser_texte(valeur)
                            ])

                    if len(data_table) > 1:
                        t = Table(data_table, colWidths=[270, 200])
                        t.setStyle(TableStyle([
                            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#b7d3f2")),
                            ('TEXTCOLOR', (0, 0), (-1, 0), colors.black),
                            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                            ('BOX', (0, 0), (-1, -1), 0.25, colors.black),
                            ('GRID', (0, 0), (-1, -1), 0.25, colors.grey)
                        ]))
                        elements.append(t)
                        elements.append(Spacer(1, 10))

                # Paragraphes de synthèse
                for champ, titre in [
                    ("axes_de_progression", "Axes de progression"),
                    ("point_d'ancrage_(_ce_qu'il_fait_bien_naturellement)", "Points d’ancrage"),
                    ("app_qui_pourrait_etre_propose", "APP qui pourraient être proposés")
                ]:
                    if champ in data_date.columns and pd.notna(data_date[champ].iloc[0]):
                        texte = nettoyer_texte_visible(data_date[champ].iloc[0])
                        elements.append(Paragraph(f"<b>{titre} :</b> {texte}", normal_style))
                        elements.append(Spacer(1, 6))

                elements.append(Spacer(1, 20))

            elements.append(PageBreak())

        doc.build(elements)
        st.success("✅ PDF généré avec succès !")
        st.download_button("📄 Télécharger le PDF", data=buffer.getvalue(),
                           file_name="synthese_stagiaires.pdf", mime="application/pdf")
