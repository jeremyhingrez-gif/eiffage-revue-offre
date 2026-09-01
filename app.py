import streamlit as st
from docx import Document
import io
import pypdf

# Configuration de la page Streamlit
st.set_page_config(
    page_title="Revue d'Offre - Eiffage Énergie Systèmes",
    page_icon="🏗️",
    layout="wide",
)

st.title("🏗️ Analyseur DCE & Générateur de Fiche de Revue d'Offre")
st.markdown("Importe ton modèle Word et tes pièces DCE. L'application extrait le texte de tes PDF et prépare ta fiche.")

st.divider()

col1, col2 = st.columns(2)

with col1:
    st.subheader("📁 1. Documents requis")
    modele_file = st.file_uploader("Modèle de Fiche Word (.docx)", type=["docx"])
    dce_files = st.file_uploader(
        "Pièces du DCE (PDF, etc.)", 
        type=["pdf", "docx", "txt"], 
        accept_multiple_files=True,
        help="Dépose l'ensemble des pièces de l'appel d'offres."
    )

with col2:
    st.subheader("⚙️ 2. Informations du Projet")
    nom_projet = st.text_input("Nom de l'affaire / Client", "CC-IN2P3")
    responsable = st.text_input("Responsable d'activités / d'affaires", "")
    date_remise = st.text_input("Date et heure de remise de l'offre", "")

st.divider()

if st.button("🚀 Lancer l'analyse du DCE et générer la Fiche", type="primary", use_container_width=True):
    if not modele_file:
        st.error("⚠️ Veuillez importer votre modèle de fichier Word (.docx).")
    elif not dce_files:
        st.warning("⚠️ Veuillez déposer au moins une pièce du DCE.")
    else:
        try:
            with st.spinner("🔄 Lecture des pièces du DCE en cours..."):
                # 1. Extraction du texte des PDF du DCE
                texte_dce_total = ""
                for uploaded_file in dce_files:
                    if uploaded_file.name.endswith(".pdf"):
                        reader = pypdf.PdfReader(uploaded_file)
                        for page in reader.pages:
                            texte_dce_total += page.extract_text() or ""
                    else:
                        # Pour les fichiers texte simples
                        texte_dce_total += uploaded_file.read().decode("utf-8", errors="ignore")

            with st.spinner("🔄 Remplissage de la fiche Word..."):
                # 2. Chargement et modification du document Word
                doc = Document(modele_file)
                
                # Remplacement dans les paragraphes
                for p in doc.paragraphs:
                    if "Responsable" in p.text and responsable:
                        p.text = f"Responsable d’activités/d’affaires : {responsable}"
                    if "Client" in p.text:
                        p.text = f"Client / Affaire : {nom_projet}"

                # Remplacement dans les tableaux du document Word
                for table in doc.tables:
                    for row in table.rows:
                        for cell in row.cells:
                            if "Responsable" in cell.text and responsable:
                                cell.text = f"Responsable : {responsable}"
                            elif "Devis" in cell.text:
                                cell.text = f"Affaire : {nom_projet}"

                # Sauvegarde du résultat dans un flux mémoire
                output_io = io.BytesIO()
                doc.save(output_io)
                output_io.seek(0)

            st.success("✅ Analyse réalisée et Fiche Word complétée avec succès !")
            
            st.download_button(
                label="📥 Télécharger votre Fiche de Revue d'Offre complétée (.docx)",
                data=output_io,
                file_name=f"Fiche_Revue_Offre_{nom_projet.replace(' ', '_')}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True
            )

        except Exception as e:
            st.error(f"Une erreur est survenue lors du traitement : {e}")
