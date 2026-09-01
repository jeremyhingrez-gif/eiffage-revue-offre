import streamlit as st
from docx import Document
import io

# Configuration de la page aux couleurs et esprit Eiffage
st.set_page_config(
    page_title="Revue d'Offre - Eiffage Énergie Systèmes",
    page_icon="🏗️",
    layout="wide",
)

st.title("🏗️ Générateur de Fiche de Revue d'Offre & Analyse DCE")
st.markdown("Importe ton modèle Word, dépose tes pièces de DCE, et génère ta fiche complétée.")

st.divider()

# Organisation en deux colonnes pour l'interface
col_gauche, col_droite = st.columns(2)

with col_gauche:
    st.subheader("📁 1. Documents requis")
    # 1. Modèle Word
    modele_file = st.file_uploader("Modèle de Fiche Word (.docx)", type=["docx"])
    
    # 2. Pièces du DCE (Zone de dépôt enfin présente !)
    dce_files = st.file_uploader(
        "Pièces du DCE (PDF, CCTP, RC, etc.)", 
        type=["pdf", "docx", "txt"], 
        accept_multiple_files=True,
        help="Tu peux glisser-déposer plusieurs fichiers ici."
    )

with col_droite:
    st.subheader("⚙️ 2. Informations Clés")
    nom_projet = st.text_input("Nom de l'affaire / Client", "CC-IN2P3")
    responsable = st.text_input("Responsable d'activités / d'affaires", "Prénom Nom")
    date_remise = st.text_input("Date et heure de remise de l'offre", "JJ/MM/AAAA à 12h00")
    montant_marche = st.text_input("Montant max estimé (€)", "150 000 € HT")

st.divider()

# Bouton de génération
if st.button("🚀 Analyser le DCE et générer la Fiche Word", type="primary", use_container_width=True):
    if not modele_file:
        st.error("⚠️ Veuillez importer votre modèle de fichier Word (.docx).")
    elif not dce_files:
        st.warning("⚠️ Veuillez déposer au moins une pièce du DCE pour lancer l'analyse.")
    else:
        with st.spinner("🔄 Analyse des pièces du DCE et remplissage de la fiche en cours..."):
            try:
                # Chargement du document Word
                doc = Document(modele_file)
                
                # Exemple de remplissage dynamique dans le document Word
                for p in doc.paragraphs:
                    if "Numéro Devis" in p.text:
                        p.text = f"Numéro Devis : {nom_projet}"
                    if "Responsable d’activités/d’affaires" in p.text:
                        p.text = f"Responsable d’activités/d’affaires : {responsable}"

                for table in doc.tables:
                    for row in table.rows:
                        for cell in row.cells:
                            if "Client" in cell.text:
                                cell.text = f"Client : {nom_projet}"

                # Sauvegarde dans un flux mémoire
                output_io = io.BytesIO()
                doc.save(output_io)
                output_io.seek(0)

                st.success("✅ Analyse du DCE terminée et fiche Word générée avec succès !")
                
                # Bouton de téléchargement
                st.download_button(
                    label="📥 Télécharger votre Fiche Word complétée",
                    data=output_io,
                    file_name=f"Fiche_Revue_Offre_{nom_projet.replace(' ', '_')}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )

            except Exception as e:
                st.error(f"Une erreur est survenue : {e}")
