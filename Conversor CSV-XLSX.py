import streamlit as st
import pandas as pd
import os
from io import BytesIO

def process_csv(df, basename):
    basename = basename.lower()

    if "patient" in basename:
        if 'Type' not in df.columns:
            df.insert(0, 'Type', 'PATIENT')
        else:
            df.insert(0, 'Type', df.pop('Type'))

        if 'ImportType' not in df.columns:
            df.insert(1, 'ImportType', 'Person')

        if 'Patient ID' in df.columns:
            df.rename(columns={'Patient ID': 'PatientId'}, inplace=True)

    elif "appointment" in basename:
        if 'fromTime' in df.columns:
            df.insert(0, 'fromTime', df.pop('fromTime'))

        if 'ImportType' not in df.columns:
            df.insert(1, 'ImportType', 'Appointment')

    elif "bookentry" in basename:
        if 'PostDate' in df.columns:
            df.insert(0, 'PostDate', df.pop('PostDate'))

        if 'ImportType' not in df.columns:
            df.insert(1, 'ImportType', 'BookEntry')

    elif "dentist" in basename:
        if 'Type' not in df.columns:
            df.insert(0, 'Type', 'DENTIST')
        else:
            df.insert(0, 'Type', df.pop('Type'))

        if 'ImportType' not in df.columns:
            df.insert(1, 'ImportType', 'Person')

    elif "financialclinics" in basename:
        if 'Account' not in df.columns:
            df.insert(0, 'Account', 'Caixa Geral')
        else:
            df.insert(0, 'Account', df.pop('Account'))

        if 'ImportType' not in df.columns:
            df.insert(1, 'ImportType', 'FinancialClinics')

    elif "openbudget" in basename:
        if 'TableName' not in df.columns:
            df.insert(0, 'TableName', 'Importação')
        else:
            df.insert(0, 'TableName', df.pop('TableName'))

        if 'ImportType' not in df.columns:
            df.insert(1, 'ImportType', 'Budgets')

    elif "treatmentoperation" in basename:
        if 'ProcedureDescription' in df.columns:
            df.insert(0, 'ProcedureDescription', df.pop('ProcedureDescription'))
            df['ProcedureDescription'].fillna('Consulta', inplace=True)
        else:
            st.warning("Coluna ProcedureDescription não encontrada")

        if 'ImportType' not in df.columns:
            df.insert(1, 'ImportType', 'TreatmentOperation')

    return df

st.title("Processamento e Visualização de CSV")

uploaded_files = st.file_uploader("Escolha arquivos CSV", type="csv", accept_multiple_files=True)

if uploaded_files:
    processed_dataframes = {}

    for uploaded_file in uploaded_files:
        df = pd.read_csv(uploaded_file, encoding='latin1')  # ou encoding='ISO-8859-1'
        st.write(f"Processando arquivo: {uploaded_file.name}")
        df_processed = process_csv(df, uploaded_file.name)
        processed_dataframes[uploaded_file.name] = df_processed

        st.subheader(f"Arquivo: {uploaded_file.name}")
        st.dataframe(df_processed)

    if st.button("Salvar todos em Excel"):
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            for name, df in processed_dataframes.items():
                df.to_excel(writer, sheet_name=name[:31], index=False)
        
        st.download_button(
            label="Baixar arquivo Excel",
            data=output.getvalue(),
            file_name="processed_files.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
