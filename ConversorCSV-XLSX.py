import streamlit as st
import pandas as pd
import os
from io import BytesIO
import zipfile
# Conversor de CSV para XLSX com Streamlit

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
            df.rename(columns={'Patient ID': 'ImportedId'}, inplace=True)
            
        if 'OtherDocumentId' in df.columns:
            df['OtherDocumentId'] = df['OtherDocumentId'].apply(
                lambda x: str(x).zfill(11) if pd.notnull(x) and str(x).strip() != '' else x
            )

        if 'CivilStatus' in df.columns:
            df['CivilStatus'] = df['CivilStatus'].replace({
                'Casado (MARRIED)': 'MARRIED',
                'Casado': 'MARRIED',
                'Solteiro (SINGLE)': 'SINGLE',
                'Solteiro': 'SINGLE',
                'Divorciado (DIVORCED)': 'DIVORCED',
                'Divorciado': 'DIVORCED',
                'Viúvo (WIDOWED)': 'WIDOWED',
                'Viúvo': 'WIDOWED'
            })

    elif "appointment" in basename:
        if 'FromTime' in df.columns:
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

st.title("Conversão de arquivo .CSV para .XLSX")
st.markdown("""
    Os arquivos são identificados por nome, e deve devem ter os nomes exatos:
    - patient.csv
    - appointment.csv
    - bookentry.csv
    - dentist.csv
    - financialclinics.csv
    - openbudget.csv
    - treatmentoperation.csv
            
    Este aplicativo permite carregar arquivos CSV, processá-los e convertê-los para XLSX.
    Os arquivos convertidos serão disponibilizados para download em um arquivo ZIP.
    
    OBS: adiciona as colunas ImportType e move as colunas-chave para a primeira posição na tabela.
""")

uploaded_files = st.file_uploader("Escolha arquivos CSV", type="csv", accept_multiple_files=True)

# Iniciar a conversão dos arquivos carregados e salvar cada um deles separadamente com o nome original mas em formato XLSX, depois zipar e fazer o download do zip

if uploaded_files:
    xlsx_files = []
    for uploaded_file in uploaded_files:
        df = pd.read_csv(uploaded_file, encoding='latin1')  # ou encoding='ISO-8859-1'
        df = process_csv(df, uploaded_file.name)
        # Salva em memória como XLSX
        xlsx_buffer = BytesIO()
        df.to_excel(xlsx_buffer, index=False)
        xlsx_buffer.seek(0)
        # Guarda o nome original, trocando .csv por .xlsx
        xlsx_name = os.path.splitext(uploaded_file.name)[0] + ".xlsx"
        xlsx_files.append((xlsx_name, xlsx_buffer.read()))

    # Cria o ZIP em memória
    zip_buffer = BytesIO()
    with zipfile.ZipFile(zip_buffer, "w") as zip_file:
        for xlsx_name, xlsx_bytes in xlsx_files:
            zip_file.writestr(xlsx_name, xlsx_bytes)
    zip_buffer.seek(0)

    st.success("Conversão concluída! Baixe o arquivo ZIP abaixo.")
    st.download_button(
        label="Baixar todos em ZIP",
        data=zip_buffer,
        file_name="planilhas_convertidas.zip",
        mime="application/zip"
    )

