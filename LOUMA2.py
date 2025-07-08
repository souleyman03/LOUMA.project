import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
import tempfile

# Titre de l'application
st.title("📦 Générateur de Reporting Ventes SIM")

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

    logins_concernes = ["pvt_mwadk0290", "pvt_mwadk194", "pvt_mwadk181", "pvt_mwadk236",
        "pvt_sosy134", "pvt_sosy0290", "pvt_sosy150", "pvt_sosy165",
        "pvt_dfallf0271", "pvt_dfallf0182", "pvt_dfallf0272", "pvt_dfallf0220",
        "Pvt_mbpling114", "Pvt_mbpling009", "Pvt_mbpling0230", "Pvt_mbpling173",
        "pvt_smmc301", "pvt_smmc2695", "pvt_smmc303", "pvt_smmc653",
        "pvt_tcg_0260", "pvt_tcg_0331", "pvt_tcg_0124", "pvt_tcg_0035"]

        # Nettoyage / Préparation
    df = df.rename(columns={'SOMME_SIM_VENDUES': 'TOTAL_SIM'})
    df = df.rename(columns={'ACCUEIL_VENDEUR': 'PVT'})
    df = df.rename(columns={'LOGIN_VENDEUR': 'LOGIN'})
    df = df.rename(columns={'AGENCE_VENDEUR': 'DRV'})
        

    def clean_cols(df):
        df['DRV'] = df['DRV'].astype(str).str.strip().str.upper()
        #df['PVT'] = df['PVT'].astype(str).str.strip().str.upper()
        df['NOM_VENDEUR'] = df['NOM_VENDEUR'].astype(str).str.strip().str.upper()
        df['PRENOM_VENDEUR'] = df['PRENOM_VENDEUR'].astype(str).str.strip().str.upper()
        return df

    df = clean_cols(df)

    # 🔎 Filtrer les ventes LOUMA
    df_filtre = df[df['LOGIN'].astype(str).isin(logins_concernes)]
    st.write("📊 Ventes LOUMA mensuelles :", df_filtre.shape[0], "lignes")

    st.success(f"✅ Feuille chargée avec succès !")
    st.dataframe(df.head())

    # 1. Dictionnaire de correspondance DRV ➡️ PVT
    correspondance_pvt = {
        "DR2": "PVT TEKNICOM CONSULTING GROUPE",
        "DR CENTRE": "PVT SERIGNE MBACKE MADINA CISSE",
        "DR EST": "PVT DEMBA FALL",
        "DR NORD": "PVT MADINA BUSINESS PRO",
        "SUD EST": "PVT SEYDINA OUSMANE SY",
        "DR SUD": "PVT MOR WADE"
    }

    df_filtre["DRV"] = df_filtre["DRV"].replace({ 
    "DV-DRV2_DIRECTION REGIONALE DES VENTES DAKAR 2": "DR2",
    "DV-DRVS_DIRECTION REGIONALE DES VENTES SUD": "DR SUD",
    "DV-DRVSE_DIRECTION REGIONALE DES VENTES SUD-EST": "SUD EST",
    "DV-DRVN_DIRECTION REGIONALE DES VENTES NORD": "DR NORD",
    "DV-DRVC_DIRECTION REGIONALE DES VENTES CENTRE": "DR CENTRE",
    "DV-DRVE_DIRECTION REGIONALE DES VENTES EST": "DR EST"
        })

    # 2. Nettoyer la colonne DRV (au cas où)
    df_filtre['DRV'] = df_filtre['DRV'].astype(str).str.strip().str.upper()

    # 3. Ajouter la colonne PVT à partir du mapping
    df_filtre['PVT'] = df_filtre['DRV'].map(correspondance_pvt)

    #définir les colonnes pour les paiements
    df_filtre['OBJECTIF'] = 240
    df_filtre["TAUX D'ATTEINTE"] = (df_filtre['TOTAL_SIM'] / df_filtre['OBJECTIF']).apply(lambda x: f"{round(x*100)}%")
    df_filtre['SI 100% ATTEINT'] = 100000
    df_filtre['PAIEMENT'] = df_filtre['TOTAL_SIM'].apply(lambda x: 100000 if x >= 240 else round((x/240)*100000))
    df_filtre['PAIEMENT CHAUFFEUR'] = 150000
    df_filtre['PAIEMENT CHAUFFEUR'] = df_filtre['PAIEMENT CHAUFFEUR'].mask(df_filtre['DRV'].duplicated())
    df_filtre['TOTAL SIM+CHAUFFEUR'] = None

    # 👉 Ajouter les lignes de total après chaque DRV
    df_with_totals = pd.DataFrame(columns=df_filtre.columns)

    for drv, group in df_filtre.groupby('DRV'):
        df_with_totals = pd.concat([df_with_totals, group], ignore_index=True)

        total_paiement = group['PAIEMENT'].sum()
        total_general = group['PAIEMENT'].sum() + group['PAIEMENT CHAUFFEUR'].sum()
        row_total = {
            'DRV': f"{drv}",
            'PVT': "TOTAL PVT",
            'PAIEMENT': total_paiement ,
            'TOTAL SIM+CHAUFFEUR': total_general
                }
        df_with_totals = pd.concat([df_with_totals, pd.DataFrame([row_total])], ignore_index=True)



    # === Générer tableau Paiement par PVT ===

    # 1. Grouper par DRV et PVT pour obtenir le total des paiements
    df_par_pvt = df_filtre.groupby(['DRV', 'PVT']).agg({'PAIEMENT': 'sum'}).reset_index()
    df_par_pvt = df_par_pvt.rename(columns={'PAIEMENT': 'MONTANT'})

    df_par_pvt['MONTANT'] = df_par_pvt['MONTANT'] + 150000

    # 2. Ajouter GAIN PVT (5%) et TOTAL GENERAL
    df_par_pvt['GAIN PVT (5%)'] = df_par_pvt['MONTANT'] * 0.05
    df_par_pvt['TOTAL GENERAL'] = df_par_pvt['MONTANT'] + df_par_pvt['GAIN PVT (5%)']




    # Affichage du tableau simplifié
    #cols_affichage = ['DRV', 'PVT', 'PRENOM_VENDEUR', 'NOM_VENDEUR', 'TOTAL_SIM']
    cols_affichage = ['DRV', 'PVT', 'PRENOM_VENDEUR', 'NOM_VENDEUR', 'TOTAL_SIM', 'OBJECTIF', "TAUX D'ATTEINTE", 'SI 100% ATTEINT', 'PAIEMENT', 'PAIEMENT CHAUFFEUR', 'TOTAL SIM+CHAUFFEUR']
    st.dataframe(df_with_totals[cols_affichage])

    # Export Excel
    buffer_paiement = BytesIO()
    with pd.ExcelWriter(buffer_paiement, engine='openpyxl') as writer:
        df_with_totals[cols_affichage].to_excel(writer, sheet_name='DETAILS PAIEMENT JUIN VTO', index=False)
        #df_filtre[cols_affichage].to_excel(writer, sheet_name='PAIEMENT PAR PVT', index=False)
        df_par_pvt.to_excel(writer, sheet_name='PAIEMENT PAR PVT', index=False)
    buffer_paiement.seek(0)

    st.download_button(
        label="📥 Télécharger le fichier de Paiement Mensuel",
        data=buffer_paiement,
        file_name="paiement_mensuel_vto.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )