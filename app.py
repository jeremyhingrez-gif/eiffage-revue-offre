import streamlit as st
from docx import Document
import io

# Configuration de la page aux couleurs et esprit Eiffage
st.set_page_config(
    page_title="Revue d'Offre - Eiffage Énergie Systèmes",
    page_icon="🏗️",
    layout="centered",
)

st.title("🏗️ Générateur de Fiche de Revue d'Offre")
st.markdown("Importe ton modèle Word vierge, renseigne les informations clés de l'affaire, et télécharge ta fiche complétée.")

st.divider()

# 1. Importation du modèle Word
st.subheader("1. Modèle de Fiche Word")
modele_file = st.file_uploader("Sélectionne ton fichier modèle (.docx)", type=["docx"])

# 2. Saisie des informations principales
st.subheader("2. Informations Clés de l'Affaire")
col1, col2 = st.columns(2)

with col1:
    nom_projet = st.text_input("Nom de l'affaire / Client", "CC-IN2P3")
    responsable = st.text_input("Responsable d'activités / d'affaires", "Prénom Nom")
    date_remise = st.text_input("Date et heure de remise de l'offre", "JJ/MM/AAAA à 12h00")

with col2:
    montant_marche = st.text_input("Montant max estimé (€)", "150 000 € HT")
    contact_envoi = st.text_input("Lien / Contact pour envoi offre", "Plateforme e-marchepub")

st.divider()

# 3. Bouton de génération
if st.button("🚀 Générer la Fiche Word Complétée", type="primary", use_container_width=True):
    if not modele_file:
        st.error("⚠️ Veuillez d'abord importer votre modèle de fichier Word (.docx).")
    else:
        try:
            # Chargement du document Word depuis le fichier uploadé
            doc = Document(modele_file)
            
            # Parcours et remplacement des informations dans les paragraphes et tableaux du document Word
            # On recherche les champs clés pour les remplacer dynamiquement
            for p in doc.paragraphs:
                if "Numéro Devis" in p.text:
                    p.text = f"Numéro Devis : {nom_projet}"
                if "Responsable d’activités/d’affaires" in p.text:
                    p.text = f"Responsable d’activités/d’affaires : {responsable}"

            # Parcours des tableaux du document Word (là où se trouvent les cases à cocher / tableaux de la fiche)
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        if "Client" in cell.text:
                            cell.text = f"Client : {nom_projet}"
                        elif "Date et heure de remise" in cell.text:
                            cell.text = f"Remise : {date_remise}"

            # Sauvegarde dans un flux mémoire pour téléchargement direct
            output_io = io.BytesIO()
            doc.save(output_io)
            output_io.seek(0)

            st.success("✅ Fiche de revue d'offre générée avec succès !")
            
            # Bouton de téléchargement du vrai fichier Word
            st.download_button(
                label="📥 Télécharger votre Fiche Word complétée",
                data=output_io,
                file_name=f"Fiche_Revue_Offre_{nom_projet.replace(' ', '_')}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True
            )

        except Exception as e:
            st.error(f"Une erreur est survenue lors du traitement du fichier Word : {e}")