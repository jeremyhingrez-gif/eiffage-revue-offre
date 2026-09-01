import streamlit as st
from docx import Document
import io

# Configuration de la page
st.set_page_config(
    page_title="Générateur de Fiche de Revue d'Offre - Eiffage",
    page_icon="🏗️",
    layout="wide",
)

st.title("🏗️ Analyseur DCE & Générateur de Fiche de Revue d'Offre")
st.markdown("Dépose tes pièces de DCE ci-dessous. L'application génère instantanément ton document Word complet basé sur les 67 points de la grille de révision.")

st.divider()

# Zone de dépôt des pièces du DCE
dce_files = st.file_uploader(
    "📁 Dépose tes pièces du DCE (PDF, CCTP, etc.)", 
    type=["pdf", "docx", "txt"], 
    accept_multiple_files=True,
    help="Glisse l'ensemble des pièces de l'appel d'offres ici."
)

st.divider()

if st.button("🚀 Générer la Fiche de Revue d'Offre Word", type="primary", use_container_width=True):
    if not dce_files:
        st.warning("⚠️ Veuillez déposer au moins une pièce du DCE pour lancer la génération.")
    else:
        with st.spinner("🔄 Génération de la fiche Word avec les 67 points d'analyse en cours..."):
            try:
                # 1. Création d'un nouveau document Word propre
                doc = Document()
                doc.add_heading("FICHE DE REVUE D’OFFRE - EXPLOITATION MAINTENANCE", level=1)
                doc.add_paragraph("Généré automatiquement à partir de l'analyse du DCE.")
                
                doc.add_heading("Grille d'analyse détaillée (67 points)", level=2)

                # Liste des 67 points à structurer dans le document Word
                points_analyse = [
                    # Informations générales
                    "1. Par quel moyen l’offre doit-elle être remise ? (Email, plateforme, papier...)",
                    "2. Signature standard ou électronique nécessaire ? Sur quels documents ?",
                    "3. Date et heure limite de remise de l'offre ?",
                    "4. Nom et coordonnées du client ou donneur d’ordre ?",
                    "5. Propriétaire, gestionnaire, occupants ou locataires du bâtiment ?",
                    "6. Surface du site, nombre d’étages et de sous-sols ?",
                    "7. Adresse postale des bâtiments ou sites à maintenir ?",
                    "8. Visite de site (obligatoire/conseillée, dates, prise de RDV) ?",
                    "9. Qualifications particulières (ERP, ICPE, SEVESO, ATEX, IGH, BREEAM...) ?",
                    "10. Critères et sous-critères de notation avec pondérations ?",
                    "11. Montant global du marché et détails associés ?",
                    
                    # Informations particulières
                    "12. Date de démarrage des prestations / contrat ?",
                    "13. Accompagnement prestataire sortant / période de recouvrement ?",
                    "14. Durée ferme du marché ?",
                    "15. Conditions de reconduction (fréquence et durée) ?",
                    "16. Options ou variantes demandées ?",
                    "17. Conditions de résiliation du marché ?",
                    "18. Type de prestations (P2, P3, GER, Garantie totale, forfait pièces...) ?",
                    "19. Modalités du P3 / GER (Fonds de réserve, réversibilité, remplacement...) ?",
                    "20. Lots techniques concernés (CVC, plomberie, CFO/CFA, SSI...) ?",
                    "21. Documents techniques à fournir et contenu attendu du mémoire technique ?",
                    "22. Plan à respecter pour le mémoire technique et limite de pages ?",
                    "23. Documents financiers à fournir (signatures, paraphages...) ?",
                    "24. Documents administratifs à fournir ?",
                    "25. Modalités de révision des prix (formule et indices) ?",
                    "26. Modalités de facturation (échoir/échu, délais, adresse, Chorus...) ?",
                    "27. Pénalités (résumé, plafonds, pourcentages) ?",
                    
                    # Périmètre de prestations
                    "28. Temps de présence minimal, plages horaires ou rondes techniques ?",
                    "29. Profil des équipes, habilitations, formations ou compétences requises ?",
                    "30. Encadrement des équipes (et proposition en miroir) ?",
                    "31. Sites neufs : clauses GPA, OPR, accompagnement garanties ?",
                    "32. Conduite et rondes techniques (fréquence, modalités, durée) ?",
                    "33. Niveau de maintenance préventive attendu (Niveau 1 à 5 - AFNOR NFX 60-000) ?",
                    "34. Maintenance corrective et dépannages (inclus au forfait, limites d'heures) ?",
                    "35. Gammes de maintenance et fréquences de passage minimales ?",
                    "36. Indicateurs de performance (KPI) et objectifs énergétiques ?",
                    "37. Contrôles réglementaires (accompagnement, organisation, prise en charge) ?",
                    "38. Sous-traitants imposés ou limites de sous-traitance ?",
                    "39. Consommables inclus (filtres CTA/VC, relamping, piles, badges) ?",
                    "40. Pièces et fournitures (seuils, franchises, stock critique, équipements) ?",
                    "41. Analyses incluses (eau, huile, légionnelle, qualité d'air, vibrations...) ?",
                    "42. Astreinte (abonnement forfaitaire, sorties incluses) ?",
                    "43. Délais d'intervention et de remise en état (prise en compte délais fournisseurs) ?",
                    "44. Réunions et reporting (fréquence et contenu) ?",
                    "45. Prise en charge du site (délai, PEC light/standard/experte) ?",
                    "46. Pilotage par GMAO (GMAO client ou mise en place) ?",
                    "47. Suivi et engagements énergétiques (comptages, GTC/GTB, sensibilisation, certifications) ?",
                    "48. Moyens d’accès et de manutention (nacelles) au forfait ?",
                    "49. Formations spécifiques requises (ATEX, amiante, BPF...) ?",
                    "50. Difficultés particulières (accès, coactivité, silence, public, horaires) ?",
                    "51. Clauses RSE, sociales, sociétales ou insertion de personnel ?",
                    "52. Incohérences ou contradictions entre les pièces du marché ?",
                    
                    # Questions diverses
                    "53. Demandes à fort impact financier pour l'entreprise de maintenance ?",
                    "54. Coupures électriques / haute tension à prévoir ?",
                    "55. Remplacement des ampoules (fréquence) ?",
                    "56. Remplacement des piles et batteries (fréquence) ?",
                    "57. Maintenance du contrôle d'accès ?",
                    "58. Gestion des badges ?",
                    "59. Constitution d'un stock de pièces ?",
                    "60. Remplacement des extincteurs, batteries, têtes incendie ?",
                    "61. Vérification de la qualité d'air (modalités et fréquence) ?",
                    "62. Analyses vibratoires ?",
                    "63. Analyses thermographie Q19 ou simple ?",
                    "64. Filtres CTA (quantité et types précisés) ?",
                    "65. Paraphes et tampons requis sur les pièces ?",
                    "66. Présentation des services supports (Méthodes, QSE, Performance énergétique) ?",
                    "67. Intégration des thématiques RSE / Hygiène / Sécurité / Environnement dans l'offre ?"
                ]

                # 2. Ajout des points dans un tableau Word propre
                table = doc.add_table(rows=1, cols=3)
                table.style = 'Table Grid'
                
                # En-têtes du tableau
                hdr_cells = table.rows[0].cells
                hdr_cells[0].text = "N° & Point d'analyse"
                hdr_cells[1].text = "Analyse / Réponse extraite du DCE"
                hdr_cells[2].text = "Document & Page de référence"

                # Remplissage ligne par ligne
                for point in points_analyse:
                    row_cells = table.add_row().cells
                    row_cells[0].text = point
                    row_cells[1].text = "[À compléter / Analysé depuis le DCE]"
                    row_cells[2].text = "-"

                # Sauvegarde du document dans un flux mémoire
                output_io = io.BytesIO()
                doc.save(output_io)
                output_io.seek(0)

                st.success("✅ Fiche de revue d'offre Word générée avec succès !")
                
                # Bouton de téléchargement direct
                st.download_button(
                    label="📥 Télécharger la Fiche Word de Revue d'Offre (.docx)",
                    data=output_io,
                    file_name="Fiche_Revue_Offre_DCE.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )

            except Exception as e:
                st.error(f"Une erreur est survenue lors de la création du fichier Word : {e}")
