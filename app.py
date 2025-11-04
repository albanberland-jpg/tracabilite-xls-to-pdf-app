import streamlit as st
import pandas as pd
from io import BytesIO
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.colors import HexColor
import unicodedata

# --- Configuration de la page Streamlit ---
st.set_page_config(page_title="Fiches d’évaluation", page_icon="📘")
st.title("📘 Générateur de fiches d’évaluation")

# --- Import du fichier Excel (.xlsx) ---
uploaded_file = st.file_uploader("Importer un fichier Excel (.xlsx)", type=["xlsx"])
if uploaded_file:
    df = pd.read_excel(uploaded_file)
    st.success("✅ Fichier importé avec succès !")
    st.dataframe(df.head())

    # --- Normalisation des noms de colonnes ---
    def nettoyer_colonne(c):
        c = str(c)
        c = ''.join(ch for ch in unicodedata.normalize('NFKD', c) if not unicodedata.combining(ch))
        return c.strip().lower().replace(" ", "_")
    df.columns = [nettoyer_colonne(c) for c in df.columns]

    st.write("🔍 Colonnes importées :", df.columns.tolist())

    # --- Recherche intelligente des colonnes ---
    def find_column(keyword, exclude=None):
        exclude = exclude or []
        for c in df.columns:
            cname = str(c).lower().strip()
            if keyword in cname and not any(exc in cname for exc in exclude):
                return c
        return None

    prenom_col = find_column("prenom", ["stagiaire"])
    nom_col     = find_column("nom",    ["stagiaire", "prenom"])
    stagiaire_col = find_column("stagiaire")
    date_col      = find_column("date")

    st.write(f"🧾 Colonnes détectées → prenom: {prenom_col}, nom: {nom_col}, stagiaire: {stagiaire_col}, date: {date_col}")

    # --- Création du champ « formateur » ---
    if prenom_col is not None and nom_col is not None:
        try:
            df["formateur"] = df[prenom_col].astype(str).str.strip() + " " + df[nom_col].astype(str).str.strip()
            st.success("✅ Champ 'formateur' créé avec succès.")
        except Exception as e:
            st.error(f"❌ Erreur lors de la création du champ formateur : {e}")
            df["formateur"] = ""
    else:
        st.warning("⚠️ Colonnes 'prenom' et/ou 'nom' introuvables — le champ 'formateur' sera laissé vide.")
        df["formateur"] = ""

    # --- Colonnes à masquer dans le PDF ---
    colonnes_a_masquer = [
        "e-mail", "organisation_-_departement", "jcmsplugin.liveform.export.title",
        "temps_ecoule", "taux_de_reussite", "score", "tentative_n°", "reussite",
        "nbre_de_questions", # également le champ « nom » (mais nom_col utilisé ailleurs)
    ]

    # --- Sélection des colonnes utiles ---
    colonnes_utiles = [c for c in df.columns if all(x not in c for x in colonnes_a_masquer)]
    df = df[colonnes_utiles]

    # --- Suppression des lignes sans évaluation (selon colonnes « eval » ou « commentaire ») ---
    colonnes_eval = [c for c in df.columns if "eval" in c or "commentaire" in c or "observation" in c]
    if colonnes_eval:
        df = df.dropna(how="all", subset=colonnes_eval)

    # --- Tri des données ---
    if stagiaire_col and date_col:
        df = df.sort_values(by=[stagiaire_col, date_col])

    # --- Groupement par stagiaire ---
    groupes_stagiaires = df.groupby(stagiaire_col)

    # --- Bouton pour générer les fiches PDF ---
    if st.button("📄 Générer les fiches PDF"):

        # --- Fonction de coloration conditionnelle ---
        def coloriser_valeur(val):
            if not isinstance(val, str):
                return str(val)
            val_up = val.strip().upper()
            couleurs = {
                "FAIT":      "#007A33",  # vert foncé
                "A":         "#00B050",  # vert clair
                "EN COURS":  "#FFD700",  # jaune
                "ECA":       "#ED7D31",  # orange
                "NE":        "#808080",  # gris
                "NA":        "#C00000",  # rouge
            }
            couleur = couleurs.get(val_up)
            if couleur:
                return f'<font color="{couleur}"><b>{val}</b></font>'
            return val

        # --- Initialisation du document PDF ---
        buffer = BytesIO()
        doc = SimpleDocTemplate(buffer, pagesize=A4)
        styles = getSampleStyleSheet()

        # --- Styles personnalisés ---
        titre_style       = ParagraphStyle("TitrePrincipal", parent=styles["Title"],       alignment=TA_CENTER, textColor=HexColor("#003366"))
        sous_titre_style  = ParagraphStyle("SousTitre",       parent=styles["Heading2"],  textColor=HexColor("#006699"))
        champ_style       = ParagraphStyle("Champ",           parent=styles["Normal"],    spaceAfter=6)
        section_style     = ParagraphStyle("Section",         parent=styles["Heading3"],  textColor=HexColor("#004C99"), spaceBefore=12, spaceAfter=6)
        contenu_style     = ParagraphStyle("Contenu",         parent=styles["Normal"],    leftIndent=12, spaceAfter=4, fontName="Helvetica", fontSize=10)

        # --- Construction des éléments PDF ---
        elements = []
        elements.append(Paragraph("Test <font color='#FF0000'><b>rouge</b></font>", contenu_style))
        elements.append(Spacer(1, 12))

        for stagiaire, data_stagiaire in groupes_stagiaires:
            elements.append(Paragraph("📘 Fiche d’évaluation", titre_style))
            elements.append(Spacer(1, 12))
            elements.append(Paragraph(f"<b>Stagiaire évalué :</b> {stagiaire}", sous_titre_style))
            elements.append(Spacer(1, 8))

            for _, ligne in data_stagiaire.iterrows():
                # Informations générales
                if date_col and pd.notna(ligne.get(date_col)):
                    elements.append(Paragraph(f"<b>Évaluation du :</b> {ligne[date_col]}", champ_style))
                if ligne.get("formateur"):
                    elements.append(Paragraph(f"<b>Formateur :</b> {ligne['formateur']}", champ_style))
                elements.append(Spacer(1, 10))

                # Section : APP non soumis à évaluation
                # (identifie les colonnes qui contiennent « app_non_soumis »)
                cols_non = [c for c in df.columns if "app_non_soumis" in c]
                if cols_non:
                    elements.append(Paragraph("🟡 APP non soumis à évaluation", section_style))
                    for c in cols_non:
                        val = ligne.get(c)
                        if pd.notna(val):
                            nom_app = c.split("/")[-1].strip().capitalize()
                            texte_val = coloriser_valeur(str(val))
                            elements.append(Paragraph(f"• {nom_app} : {texte_val}", contenu_style))
                    elements.append(Spacer(1, 8))

                # Section : APP évalués
                cols_eval = [c for c in df.columns if "app_evalues" in c]
                if cols_eval:
                    elements.append(Paragraph("🟢 APP évalués", section_style))
                    for c in cols_eval:
                        val = ligne.get(c)
                        if pd.notna(val):
                            nom_app = c.split("/")[-1].strip().capitalize()
                            texte_val = coloriser_valeur(str(val))
                            elements.append(Paragraph(f"• {nom_app} : {texte_val}", contenu_style))
                    elements.append(Spacer(1, 8))

                # Section : Axes de progression
                cols_prog = [c for c in df.columns if "axes_de_progression" == c]
                if cols_prog:
                    elements.append(Paragraph("🔵 Axes de progression", section_style))
                    for c in cols_prog:
                        val = ligne.get(c)
                        if pd.notna(val):
                            elements.append(Paragraph(f"• {val}", contenu_style))
                    elements.append(Spacer(1, 8))

                # Section : Points d’ancrage
                cols_ancr = [c for c in df.columns if "point_d'ancrage" in c]
                if cols_ancr:
                    elements.append(Paragraph("🟠 Points d’ancrage", section_style))
                    for c in cols_ancr:
                        val = ligne.get(c)
                        if pd.notna(val):
                            elements.append(Paragraph(f"• {val}", contenu_style))
                    elements.append(Spacer(1, 8))

                # Section : APP qui pourraient être proposés
                cols_prop = [c for c in df.columns if "app_qui_pourrait_etre_propose" in c]
                if cols_prop:
                    elements.append(Paragraph("🟣 APP qui pourraient être proposés", section_style))
                    for c in cols_prop:
                        val = ligne.get(c)
                        if pd.notna(val):
                            elements.append(Paragraph(f"• {val}", contenu_style))
                    elements.append(Spacer(1, 8))

                # Séparation entre évaluations
                elements.append(Spacer(1, 10))
                elements.append(Paragraph("<hr width='100%' color='#CCCCCC'/>", styles["Normal"]))
                elements.append(PageBreak())

        # --- Générer le PDF et proposer téléchargement ---
        doc.build(elements)
        buffer.seek(0)
        st.download_button(
            label="⬇️ Télécharger les fiches PDF",
            data=buffer,
            file_name="fiches_evaluations.pdf",
            mime="application/pdf"
        )

else:
    st.info("📂 En attente du fichier Excel (.xlsx) à importer.")
