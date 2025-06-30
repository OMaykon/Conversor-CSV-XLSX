import streamlit as st
import pandas as pd
import os
from io import BytesIO
import zipfile

st.cache_data.clear()

# ============================================== Utilitário para conversão de datas ==============================================
def convert_date_columns(df, columns):
    for col in columns:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors='coerce')
    return df

# ============================================= Regras de negócio ===============================================

# ====================================== Processamento de arquivos XLSX =========================================
def process_xlsx(file):
    df = pd.read_excel(file, engine='openpyxl')

    if 'patient' in file.name.lower() or 'Patient' in file.name:
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

        df = convert_date_columns(df, ['BirthDate'])

    elif 'dentist' in file.name.lower() or 'Dentist' in file.name:
        if 'Type' not in df.columns:
            df.insert(0, 'Type', 'DENTIST')
        else:
            df.insert(0, 'Type', df.pop('Type'))

        if 'ImportType' not in df.columns:
            df.insert(1, 'ImportType', 'Person')

    elif 'appointment' in file.name.lower() or 'Appointment' in file.name:
        if 'FromTime' in df.columns:
            df.insert(0, 'FromTime', df.pop('FromTime'))
            df.rename(columns={'FromTime': 'fromTime'}, inplace=True)
            # Muda o formato de hora para 24h
            df['fromTime'] = pd.to_datetime(df['fromTime'], format='%H:%M:%S', errors='coerce').dt.strftime('%H:%M:%S')

        if 'ToTime' in df.columns:
            df.rename(columns={'ToTime': 'toTime'}, inplace=True)
            # Muda o formato de hora para 24h
            df['toTime'] = pd.to_datetime(df['toTime'], format='%H:%M:%S', errors='coerce').dt.strftime('%H:%M:%S')

        if 'Date' in df.columns:
            df.rename(columns={'Date': 'date'}, inplace=True)

        if 'ImportType' not in df.columns:
            df.insert(1, 'ImportType', 'Appointment')

        if 'Status' in df.columns:
            df['Status'] = df['Status'].replace({
                'Faltou': 'MISSED',
                'Atendido': 'CHECKOUT',
                'Agendado': 'CONFIRMED',
                'Confirmado': 'CONFIRMED',
                'Cancelado Dentist': 'CANCELED_DENTIST',
                'Cancelado Patient': 'CANCELED_PATIENT',
                'Atrasado': 'LATE',
                'Em espera': 'ARRIVED'
            })

        df = convert_date_columns(df, ['date'])

    elif 'bookentry' in file.name.lower() or 'BookEntry' in file.name:
        if 'PostDate' in df.columns:
            df.insert(0, 'PostDate', df.pop('PostDate'))

        if 'ImportType' not in df.columns:
            df.insert(1, 'ImportType', 'BookEntry')

        df = convert_date_columns(df, ['PostDate', 'DueDate', 'ConfirmedDate', 'ReceivedDate'])

    elif 'financialclinics' in file.name.lower() or 'FinancialClinics' in file.name:
        if 'Account' not in df.columns:
            df.insert(0, 'Account', 'Caixa Geral')
        else:
            df.insert(0, 'Account', df.pop('Account'))

        if 'ImportType' not in df.columns:
            df.insert(1, 'ImportType', 'FinancialClinics')
        
        if 'ReleaseDate' in df.columns:
            # Converte ReleaseDate para datetime
            df['ReleaseDate'] = pd.to_datetime(df['ReleaseDate'], errors='coerce')

    elif 'openbudget' in file.name.lower() or 'budget' in file.name.lower() or 'OpenBudget' in file.name:
        if 'TableName' not in df.columns:
            df.insert(0, 'TableName', 'Importação')
        else:
            df.insert(0, 'TableName', df.pop('TableName'))

        if 'ImportType' not in df.columns:
            df.insert(1, 'ImportType', 'Budgets')

        if 'SpecialtyDescription' not in df.columns:
            df.insert(len(df.columns), 'SpecialtyDescription', 'Clínica Geral')

        df = convert_date_columns(df, ['BudgetsCreateDate'])

    elif 'treatmentoperation' in file.name.lower() or 'TreatmentOperation' in file.name:
        if 'ProcedureDescription' in df.columns or 'Procedure' in df.columns or 'ProcedureName' in df.columns:
            df.insert(0, 'ProcedureDescription', df.pop('ProcedureDescription'))
            df['ProcedureDescription'].fillna('Consulta', inplace=True)
        else:
            st.warning("Coluna ProcedureDescription não encontrada")
            df.insert(0, 'ProcedureDescription', 'Consulta')
            df['ProcedureDescription'].fillna('Consulta', inplace=True)

        

        if 'ImportType' not in df.columns:
            df.insert(1, 'ImportType', 'TreatmentOperation')

        df = convert_date_columns(df, ['CreateDate', 'ExecutedDate'])

    return df

# ============================================== Processamento de CSV ==============================================

def process_csv(df, basename):
    basename = basename.lower()

    if "patient" in basename or "Patient" in basename:
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

        df = convert_date_columns(df, ['BirthDate'])

    elif "appointment" in basename or 'Appointment' in basename:
        if 'FromTime' in df.columns:
            df.insert(0, 'FromTime', df.pop('FromTime'))
            df.rename(columns={'FromTime': 'fromTime'}, inplace=True)
            # Muda o formato de hora para 24h
            df['fromTime'] = pd.to_datetime(df['fromTime'], format='%H:%M:%S', errors='coerce').dt.strftime('%H:%M:%S')

        if 'ToTime' in df.columns:
            df.rename(columns={'ToTime': 'toTime'}, inplace=True)
            # Muda o formato de hora para 24h
            df['toTime'] = pd.to_datetime(df['toTime'], format='%H:%M:%S', errors='coerce').dt.strftime('%H:%M:%S')

        if 'Date' in df.columns:
            df.rename(columns={'Date': 'date'}, inplace=True)

        if 'ImportType' not in df.columns:
            df.insert(1, 'ImportType', 'Appointment')

        if 'Status' in df.columns:
            df['Status'] = df['Status'].replace({
                'Faltou': 'MISSED',
                'Atendido': 'CHECKOUT',
                'Agendado': 'CONFIRMED',
                'Confirmado': 'CONFIRMED',
                'Cancelado Dentist': 'CANCELED_DENTIST',
                'Cancelado Patient': 'CANCELED_PATIENT',
                'Atrasado': 'LATE',
                'Em espera': 'ARRIVED'
            })

        df = convert_date_columns(df, ['date'])

    elif "bookentry" in basename or 'BookEntry' in basename:
        if 'PostDate' in df.columns:
            df.insert(0, 'PostDate', df.pop('PostDate'))

        if 'ImportType' not in df.columns:
            df.insert(1, 'ImportType', 'BookEntry')

        df = convert_date_columns(df, ['PostDate', 'DueDate', 'ConfirmedDate', 'ReceivedDate'])

    elif "dentist" in basename or 'Dentist' in basename:
        if 'Type' not in df.columns:
            df.insert(0, 'Type', 'DENTIST')
        else:
            df.insert(0, 'Type', df.pop('Type'))

        if 'ImportType' not in df.columns:
            df.insert(1, 'ImportType', 'Person')

    elif "financialclinics" in basename or "FinancialClinics" in basename:
        if 'Account' not in df.columns:
            df.insert(0, 'Account', 'Caixa Geral')
        else:
            df.insert(0, 'Account', df.pop('Account'))

        if 'ImportType' not in df.columns:
            df.insert(1, 'ImportType', 'FinancialClinics')

        if 'ReleaseDate' in df.columns:
            # Converte ReleaseDate para datetime
            df['ReleaseDate'] = pd.to_datetime(df['ReleaseDate'], errors='coerce')

    elif "openbudget" in basename or "budget" in basename or "OpenBudget" in basename:
        if 'TableName' not in df.columns:
            df.insert(0, 'TableName', 'Importação')
        else:
            df.insert(0, 'TableName', df.pop('TableName'))

        if 'ImportType' not in df.columns:
            df.insert(1, 'ImportType', 'Budgets')

        if 'SpecialtyDescription' not in df.columns:
            df.insert(len(df.columns), 'SpecialtyDescription', 'Clínica Geral')

        df = convert_date_columns(df, ['BudgetsCreateDate'])

    elif "treatmentoperation" in basename or "TreatmentOperation" in basename:
        if 'ProcedureDescription' in df.columns:
            df.insert(0, 'ProcedureDescription', df.pop('ProcedureDescription'))
            df['ProcedureDescription'].fillna('Consulta', inplace=True)
        else:
            st.warning("Coluna ProcedureDescription não encontrada")
            df.insert(0, 'ProcedureDescription', 'Consulta')
            df['ProcedureDescription'].fillna('Consulta', inplace=True)

        if 'ImportType' not in df.columns:
            df.insert(1, 'ImportType', 'TreatmentOperation')

        df = convert_date_columns(df, ['CreateDate', 'ExecutedDate'])

    return df

# ============================================== UI ==============================================

st.title("Conversão/Modelagem de arquivo .CSV para .XLSX")
st.markdown("""
    Os arquivos são identificados por nome, e devem ter os nomes exatos:
    - patient
    - appointment
    - bookentry
    - dentist
    - financialclinics
    - openbudget
    - treatmentoperation

    - Permite carregar arquivos CSV, processá-los e convertê-los para XLSX.
    - Permite carregar arquivos XLSX, processá-los e modelar removendo erros comuns.
    Os arquivos convertidos serão disponibilizados para download em um arquivo ZIP.

    OBS: adiciona as colunas ImportType e move as colunas-chave para a primeira posição na tabela.
    

    | Version: 0.25.06-1024.
""")

uploaded_files = st.file_uploader(
    "Carregue seus arquivos CSV ou XLSX",
    type=["csv", "xlsx"],
    accept_multiple_files=True,
    help="Carregue um ou mais arquivos CSV ou XLSX para conversão."
)

if not uploaded_files:
    st.warning("Por favor, carregue um ou mais arquivos CSV ou XLSX para conversão.")
    st.stop()

xlsx_files = []

for uploaded_file in uploaded_files:
    if uploaded_file.name.lower().endswith('.xlsx'):
        df = process_xlsx(uploaded_file)
    else:
        df = pd.read_csv(uploaded_file, encoding='latin1')
        df = process_csv(df, uploaded_file.name)

    xlsx_buffer = BytesIO()
    df.to_excel(xlsx_buffer, index=False, engine='openpyxl')
    xlsx_buffer.seek(0)

    xlsx_name = os.path.splitext(uploaded_file.name)[0] + ".xlsx"
    xlsx_files.append((xlsx_name, xlsx_buffer.read()))

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
