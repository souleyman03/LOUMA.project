import streamlit as st
import pandas as pd
from io import BytesIO

st.title("📊 Comparaison des Résultats Hebdomadaires (Deux Feuilles)")

# 1. Uploader le fichier contenant les deux feuilles
uploaded_file = st.file_uploader("📁 Importer le fichier Excel (avec les 2 feuilles)", type=["xlsx"])

if uploaded_file:
    xls = pd.ExcelFile(uploaded_file)
    sheet_names = xls.sheet_names

    st.info("✅ Fichier chargé. Feuilles disponibles : " + ", ".join(sheet_names))

    # 2. Sélectionner la feuille de l'app et la feuille manuelle
    feuille_app = st.selectbox("📄 Sélectionner la feuille 'App LOUMA'", options=sheet_names, key="app")
    feuille_manu = st.selectbox("📄 Sélectionner la feuille 'Résultats Manuels'", options=sheet_names, key="manuel")

    if feuille_app and feuille_manu and feuille_app != feuille_manu:
        df_app = pd.read_excel(uploaded_file, sheet_name=feuille_app)
        df_manu = pd.read_excel(uploaded_file, sheet_name=feuille_manu)

        # Renommer les colonnes
        df_app = df_app.rename(columns={'TOTAL_SIM': 'TOTAL_SIM_APP'})
        df_manu = df_manu.rename(columns={'TOTAL_SIM': 'TOTAL_SIM_MANUEL'})

        # Nettoyage des identifiants
        for df in [df_app, df_manu]:
            for col in ['DRV', 'PVT', 'PRENOM_VENDEUR', 'NOM_VENDEUR']:
                df[col] = df[col].astype(str).str.strip().str.upper()

        # Fusion des deux feuilles
        df_merged = pd.merge(
            df_app,
            df_manu,
            on=['DRV', 'PVT', 'PRENOM_VENDEUR', 'NOM_VENDEUR'],
            how='outer'
        )

        # Calcul de la différence
        df_merged['DIFF'] = df_merged['TOTAL_SIM_APP'].fillna(0) - df_merged['TOTAL_SIM_MANUEL'].fillna(0)

        # Affichage du tableau comparatif
        st.success("✅ Comparaison effectuée avec succès !")
        st.dataframe(df_merged)

        # Export Excel
        buffer = BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            df_merged.to_excel(writer, sheet_name="Comparaison", index=False)
        buffer.seek(0)

        st.download_button(
            label="📥 Télécharger le fichier comparatif",
            data=buffer,
            file_name="comparaison_reporting.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
