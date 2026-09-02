import streamlit as st
from docx import Document
import io
import pypdf
import google.generativeai as genai

# Configuration de la page
st.set_page_config(
    page_title="Analyseur DCE & Revue d'Offre - Eiffage",
    page_icon="🏗️",
    layout="wide",
)

st.title("🏗️ Assistant IA (Gemini 3.5) - Analyse de DCE & Revue d'Offre")
st.markdown("Dépose tes pièces de DCE. Gemini 3.5 analyse les documents, traite les 67 points de contrôle et te génère le fichier Word entièrement complété.")

st.divider()

# Saisie de la clé API Gemini dans la barre latérale
api_key = st.sidebar.text_input("Clé API Gemini", type="password", help="Entre ta clé API Google AI Studio.")

# Zone de dépôt des pièces du DCE
dce_files = st.file_uploader(
    "📁 Dépose tes pièces du DCE (PDF, CCTP, RC, etc.)", 
    type=["pdf", "docx", "txt"], 
    accept_multiple_files=True,
    help="Glisse l'ensemble des pièces de l'appel d'offres ici."
)

st.divider()

if st.button("🚀 Lancer l'analyse du DCE et générer la Fiche Word", type="primary", use_container_width=True):
    if not dce_files:
        st.warning("⚠️ Veuillez déposer au moins une pièce du DCE.")
    elif not api_key:
        st.error("⚠️ Veuillez renseigner votre clé API Gemini dans la barre latérale.")
    else:
        with st.spinner("🔄 Étape 1/3 : Extraction du texte des pièces du DCE..."):
            try:
                texte_dce_total = ""
                for uploaded_file in dce_files:
                    if uploaded_file.name.endswith(".pdf"):
                        reader = pypdf.PdfReader(uploaded_file)
                        for i, page in enumerate(reader.pages):
                            texte_dce_total += f"\n--- Document: {uploaded_file.name} (Page {i+1}) ---\n"
                            texte_dce_total += page.extract_text() or ""
                    else:
                        texte_dce_total += f"\n--- Document: {uploaded_file.name} ---\n"
                        texte_dce_total += uploaded_file.read().decode("utf-8", errors="ignore")

            except Exception as e:
                st.error(f"Erreur lors de la lecture des fichiers : {e}")
                st.stop()

        with st.spinner("🤖 Étape 2/3 : Analyse approfondie par Gemini 3.5 des 67 points de contrôle..."):
            try:
                # Configuration de l'API avec la clé utilisateur
                genai.configure(api_key=api_key)
                
                prompt_systeme = (
                    "Tu es un expert en réponse aux appels d'offres dans le domaine de la maintenance et de la performance énergétique "
                    "pour Eiffage Énergie Systèmes. Analyse les documents du DCE fournis et réponds précisément aux 67 points de la grille "
                    "de revue d'offre, en indiquant pour chaque point l'information trouvée ainsi que le document et la page de référence."
                )

                # Utilisation du modèle gemini-3.5-flash demandé
                model = genai.GenerativeModel('gemini-3.5-flash')
                
                response = model.generate_content([
                    prompt_systeme,
                    f"Voici le contenu complet du DCE :\n{texte_dce_total}"
                ])
                
                analyse_resultat = response.text

            except Exception as e:
                st.error(f"Erreur lors de l'appel à l'API Gemini : {e}")
                st.stop()

        with st.spinner("📝 Étape 3/3 : Création et remplissage du document Word..."):
            try:
                # Création du document Word
                doc = Document()
                doc.add_heading("FICHE DE REVUE D’OFFRE - EXPLOITATION MAINTENANCE", level=1)
                doc.add_paragraph("Analyse automatisée des pièces du DCE par Gemini 3.5 - Eiffage Énergie Systèmes.")
                
                doc.add_heading("Synthèse de l'analyse des 67 points", level=2)
                doc.add_paragraph(analyse_resultat)

                # Sauvegarde dans un flux mémoire
                output_io = io.BytesIO()
                doc.save(output_io)
                output_io.seek(0)

                st.success("✅ Analyse terminée et Fiche Word complétée avec succès !")
                
                st.download_button(
                    label="📥 Télécharger la Fiche de Revue d'Offre Word complétée (.docx)",
                    data=output_io,
                    file_name="Fiche_Revue_Offre_Gemini_3.5.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )

            except Exception as e:
                st.error(f"Erreur lors de la génération du fichier Word : {e}")
