import streamlit as st
import pandas as pd
import os
from io import BytesIO
import zipfile
# Conversor de CSV para XLSX com Streamlit
st.cache_data.clear()

# Função para processar o xlsx
def process_xlsx(file):
    df = pd.read_excel(file, engine='openpyxl')
    
    # Processa o DataFrame conforme necessário
    # verifica se tem o nome 'patient' ou 'Patient' no nome do arquivo
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
    elif 'dentist' in file.name.lower() or 'Dentist' in file.name:
        if 'Type' not in df.columns:
            df.insert(0, 'Type', 'DENTIST')
        else:
            df.insert(0, 'Type', df.pop('Type'))

        if 'ImportType' not in df.columns:
            df.insert(1, 'ImportType', 'Person')

    elif 'appointment' in file.name.lower() or 'Appointment' in file.name:
        # Verifica se as colunas FromTime, ToTime e Date existem e renomeia ou insere conforme necessário
    
        if 'FromTime' in df.columns:
            df.insert(0, 'FromTime', df.pop('FromTime'))
            df.rename(columns={'FromTime': 'fromTime'}, inplace=True)

        if 'ToTime' in df.columns:
            df.rename(columns={'ToTime': 'toTime'}, inplace=True)

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

    elif 'bookentry' in file.name.lower() or 'BookEntry' in file.name:
        if 'PostDate' in df.columns:
            df.insert(0, 'PostDate', df.pop('PostDate'))

        if 'ImportType' not in df.columns:
            df.insert(1, 'ImportType', 'BookEntry')

    elif 'financialclinics' in file.name.lower() or 'FinancialClinics' in file.name:
        if 'Account' not in df.columns:
            df.insert(0, 'Account', 'Caixa Geral')
        else:
            df.insert(0, 'Account', df.pop('Account'))
        if 'ImportType' not in df.columns:
            df.insert(1, 'ImportType', 'FinancialClinics')

    elif 'openbudget' in file.name.lower() or 'OpenBudget' in file.name:
        if 'TableName' not in df.columns:
            df.insert(0, 'TableName', 'Importação')
        else:
            df.insert(0, 'TableName', df.pop('TableName'))
        if 'ImportType' not in df.columns:
            df.insert(1, 'ImportType', 'Budget')

        if 'SpecialtyDescription' not in df.columns:
            df.insert('SpecialtyDescription', 'Clínica Geral')

    elif 'treatmentoperation' in file.name.lower() or 'TreatmentOperation' in file.name:
        if 'ProcedureDescription' in df.columns:
            df.insert(0, 'ProcedureDescription', df.pop('ProcedureDescription'))
            df['ProcedureDescription'].fillna('Consulta', inplace=True)
        else:
            st.warning("Coluna ProcedureDescription não encontrada")
        if 'ImportType' not in df.columns:
            df.insert(1, 'ImportType', 'TreatmentOperation')
    return df

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
            df.insert(0, 'FromTime', df.pop('FromTime'))
            df.rename(columns={'FromTime': 'fromTime'}, inplace=True)

        if 'ToTime' in df.columns:
            df.rename(columns={'ToTime': 'toTime'}, inplace=True)

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
            df.insert(1, 'ImportType', 'Person')

        if 'SpecialtyDescription' not in df.columns:
            df.insert('SpecialtyDescription', 'Clínica Geral')
        
        # Procura a coluna SpecialtyDescription, se existir e o valor for vazio, preenche com 'Clínica Geral'
        if 'SpecialtyDescription' in df.columns:
            df['SpecialtyDescription'].fillna('Clínica Geral', inplace=True)

    elif "treatmentoperation" in basename:
        if 'ProcedureDescription' in df.columns:
            df.insert(0, 'ProcedureDescription', df.pop('ProcedureDescription'))
            df['ProcedureDescription'].fillna('Consulta', inplace=True)
        else:
            st.warning("Coluna ProcedureDescription não encontrada")

        if 'ImportType' not in df.columns:
            df.insert(1, 'ImportType', 'TreatmentOperation')
    return df

st.title("Conversão/Modelagem de arquivo .CSV para .XLSX")
st.markdown("""
    Os arquivos são identificados por nome, e deve devem ter os nomes exatos:
    - patient
    - appointment
    - bookentry
    - dentist
    - financialclinics
    - openbudget
    - treatmentoperation
    
            
    - Permite carregar arquivos CSV, processá-los e convertê-los para XLSX.
    - Permite carregar arquivos XLSX, processá-los e modelar removendo alguns erros comuns.
    Os arquivos convertidos serão disponibilizados para download em um arquivo ZIP.
    
    OBS: adiciona as colunas ImportType e move as colunas-chave para a primeira posição na tabela.
""")

# Uploader de arquivos
uploaded_files = st.file_uploader(
    "Carregue seus arquivos CSV ou XLSX",
    type=["csv", "xlsx"],
    accept_multiple_files=True,
    help="Carregue um ou mais arquivos CSV ou XLSX para conversão."
)

# Iniciar a conversão dos arquivos carregados e salvar cada um deles separadamente com o nome original mas em formato XLSX, depois zipar e fazer o download do zip
# identiifica o tipo de arquivo pelo nome do arquivo, e processa de acordo com o tipo
if not uploaded_files:
    st.warning("Por favor, carregue um ou mais arquivos CSV ou XLSX para conversão.")
    st.stop()
# Processa os arquivos carregados
xlsx_files = []

for uploaded_file in uploaded_files:
    if uploaded_file.name.lower().endswith('.xlsx'):
        df = process_xlsx(uploaded_file)
    else:
        df = pd.read_csv(uploaded_file, encoding='latin1')  # ou encoding='ISO-8859-1'
        df = process_csv(df, uploaded_file.name)

    # Salva em memória como XLSX
    xlsx_buffer = BytesIO()
    df.to_excel(xlsx_buffer, index=False, engine='openpyxl')
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
# Exibe mensagem de sucesso e botão para download do ZIP
st.success("Conversão concluída! Baixe o arquivo ZIP abaixo.")
st.download_button(
    label="Baixar todos em ZIP",
    data=zip_buffer,
    file_name="planilhas_convertidas.zip",
    mime="application/zip"
)


