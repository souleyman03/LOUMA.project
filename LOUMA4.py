import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
import tempfile

# Uploader du fichier Excel brut
uploaded_file = st.file_uploader("📁 Importer le fichier Excel brut (hebdomadaire)", type=["xlsx", "csv"])

if uploaded_file: 

    if uploaded_file.name.endswith('.csv'):
        df = pd.read_csv(uploaded_file, encoding='utf-8', sep=';')


    else:
        # Charger toutes les feuilles sans les lire entièrement
        xls = pd.ExcelFile(uploaded_file)
            
        # Afficher les noms de feuilles disponibles
        sheet_names = xls.sheet_names
        selected_sheet = st.selectbox("🗂️ Choisir la feuille à exploiter :", options=sheet_names)
            
        # Lire uniquement la feuille sélectionnée
        df = pd.read_excel(uploaded_file, sheet_name=selected_sheet)

    
    # 2. Ajouter les colonnes TOTAL, OBJ, TR (%)
    df['TOT'] = df[["M1", "M2", "M3", "M4"]].sum(axis=1)
    
    

    
    df['TR (%)'] = (df['TOT'] / df['OBJ'] * 100).round(1).astype(str) + '%'
    



    
    # Export Excel
    buffer_paiement = BytesIO()
    with pd.ExcelWriter(buffer_paiement, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='REALISATION BILAN 3MOIS', index=False)
        
    buffer_paiement.seek(0)

    st.download_button(
        label="📥 Télécharger le fichier de Paiement Mensuel",
        data=buffer_paiement,
        file_name="BILAN 3MOIS.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )