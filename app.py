# =============================================================================
# FUSION EXCEL - MVP
# Application Streamlit pour fusionner plusieurs fichiers Excel en un seul
# =============================================================================

import streamlit as st
import pandas as pd
from io import BytesIO

# -----------------------------------------------------------------------------
# CONFIGURATION DE LA PAGE
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title="Fusion Excel",
    page_icon="📊",
    layout="centered"
)

# -----------------------------------------------------------------------------
# STYLE CSS PERSONNALISÉ
# -----------------------------------------------------------------------------
st.markdown("""
<style>
    /* Style général */
    .main {
        padding-top: 2rem;
    }
    
    /* Titre principal */
    .title {
        text-align: center;
        font-size: 2.5rem;
        font-weight: 700;
        color: #1E3A5F;
        margin-bottom: 0.5rem;
    }
    
    /* Sous-titre */
    .subtitle {
        text-align: center;
        font-size: 1.2rem;
        color: #666;
        margin-bottom: 2rem;
    }
    
    /* Box de succès */
    .success-box {
        background-color: #D4EDDA;
        border: 1px solid #C3E6CB;
        border-radius: 8px;
        padding: 1rem;
        text-align: center;
        margin: 1rem 0;
    }
    
    /* Instructions */
    .instructions {
        background-color: #F8F9FA;
        border-radius: 8px;
        padding: 1.5rem;
        margin: 1rem 0;
    }
    
    /* Footer */
    .footer {
        text-align: center;
        color: #999;
        font-size: 0.85rem;
        margin-top: 3rem;
        padding-top: 1rem;
        border-top: 1px solid #eee;
    }
    
    /* Masquer le menu Streamlit */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
</style>
""", unsafe_allow_html=True)

# -----------------------------------------------------------------------------
# EN-TÊTE
# -----------------------------------------------------------------------------
st.markdown('<p class="title">📊 Fusion Excel</p>', unsafe_allow_html=True)
st.markdown('<p class="subtitle">Fusionne plusieurs fichiers Excel en 30 secondes</p>', unsafe_allow_html=True)

# -----------------------------------------------------------------------------
# INSTRUCTIONS
# -----------------------------------------------------------------------------
with st.expander("ℹ️ Comment ça marche ?", expanded=False):
    st.markdown("""
    **3 étapes simples :**
    1. 📂 **Dépose tes fichiers** Excel (.xlsx ou .xls)
    2. ▶️ **Clique sur Fusionner**
    3. ⬇️ **Télécharge** ton fichier fusionné
    
    **Notes :**
    - Les fichiers doivent avoir la même structure (mêmes colonnes)
    - Le fichier fusionné conserve toutes les lignes de tous les fichiers
    - Les en-têtes sont automatiquement détectés
    """)

# -----------------------------------------------------------------------------
# ZONE D'UPLOAD
# -----------------------------------------------------------------------------
st.markdown("### 📂 Étape 1 : Dépose tes fichiers")

uploaded_files = st.file_uploader(
    "Glisse tes fichiers Excel ici",
    type=["xlsx", "xls"],
    accept_multiple_files=True,
    help="Tu peux sélectionner plusieurs fichiers en même temps"
)

# -----------------------------------------------------------------------------
# TRAITEMENT DES FICHIERS
# -----------------------------------------------------------------------------
if uploaded_files:
    st.markdown(f"**{len(uploaded_files)} fichier(s) sélectionné(s)**")
    
    # Afficher la liste des fichiers
    for i, file in enumerate(uploaded_files, 1):
        st.text(f"  {i}. {file.name}")
    
    st.markdown("---")
    st.markdown("### ▶️ Étape 2 : Fusionne")
    
    # Bouton de fusion
    if st.button("🔗 Fusionner les fichiers", type="primary", use_container_width=True):
        
        try:
            # Barre de progression
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            all_dataframes = []
            
            # Lecture de chaque fichier
            for i, file in enumerate(uploaded_files):
                status_text.text(f"📖 Lecture de {file.name}...")
                
                # Lire le fichier Excel
                df = pd.read_excel(file)
                all_dataframes.append(df)
                
                # Mise à jour de la progression
                progress_bar.progress((i + 1) / len(uploaded_files))
            
            status_text.text("🔗 Fusion en cours...")
            
            # Fusion de tous les DataFrames
            merged_df = pd.concat(all_dataframes, ignore_index=True)
            
            # Effacer les indicateurs de progression
            progress_bar.empty()
            status_text.empty()
            
            # Afficher le succès
            st.success(f"✅ Fusion réussie ! {len(merged_df)} lignes au total")
            
            # Aperçu du résultat
            st.markdown("### 👀 Aperçu du résultat")
            st.dataframe(merged_df.head(10), use_container_width=True)
            
            if len(merged_df) > 10:
                st.caption(f"Affichage des 10 premières lignes sur {len(merged_df)}")
            
            # Statistiques
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("Lignes", len(merged_df))
            with col2:
                st.metric("Colonnes", len(merged_df.columns))
            with col3:
                st.metric("Fichiers", len(uploaded_files))
            
            st.markdown("---")
            st.markdown("### ⬇️ Étape 3 : Télécharge")
            
            # Préparation du fichier à télécharger
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                merged_df.to_excel(writer, index=False, sheet_name='Données fusionnées')
            output.seek(0)
            
            # Bouton de téléchargement
            st.download_button(
                label="⬇️ Télécharger le fichier fusionné",
                data=output,
                file_name="fichier_fusionne.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True
            )
            
        except Exception as e:
            st.error(f"❌ Erreur lors de la fusion : {str(e)}")
            st.info("💡 Vérifie que tes fichiers ont la même structure (mêmes colonnes)")

else:
    # Message d'attente
    st.info("👆 Commence par déposer tes fichiers Excel ci-dessus")

# -----------------------------------------------------------------------------
# FOOTER
# -----------------------------------------------------------------------------
st.markdown("---")
st.markdown("""
<div class="footer">
    <p>Fait avec ❤️ pour te faire gagner du temps</p>
    <p>Tu aimes cet outil ? <a href="#">Partage-le !</a></p>
</div>
""", unsafe_allow_html=True)
