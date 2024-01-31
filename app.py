import streamlit as st
import pandas as pd
import gspread
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
import os

# Escopos e ID da planilha
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
SAMPLE_SPREADSHEET_ID = '1FPBeAXQBKy8noJ3bTF52p8JL_Eg-ptuSP6djDTsRfKE'
SAMPLE_RANGE_NAME = 'Página1!A1:X500'

def authenticate_google_sheets():
    creds = None
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('creden.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.json', 'w') as token:
            token.write(creds.to_json())
    client = gspread.authorize(creds)
    return client

def main():
    st.title("Upload de Arquivo Excel para Google Sheets")

    uploaded_file = st.file_uploader("Escolha um arquivo Excel", type=['xlsx', 'xls'])
    sheet = None

    if uploaded_file is not None:
        # Lê e exibe os dados da planilha Excel
        data = pd.read_excel(uploaded_file, header=0)
        data.columns = data.columns.map(str)
        st.write("Dados lidos do arquivo Excel:")
        st.dataframe(data)

        # Botão para conectar ao Google Sheets
        if st.button("Conectar ao Google Sheets"):
            try:
                client = authenticate_google_sheets()
                # Acessa a planilha e o intervalo específico
                sheet = client.open_by_key(SAMPLE_SPREADSHEET_ID).worksheet_by_title(SAMPLE_RANGE_NAME)
                st.success("Conectado com sucesso ao Google Sheets!")
            except Exception as e:
                st.error(f"Falha ao conectar: {e}")

    # Botão para enviar dados para o Google Sheets
    if st.button("Enviar para Google Sheets"):
        if sheet and not data.empty:
            try:
                data_list = data.values.tolist()
                sheet.update(SAMPLE_RANGE_NAME, data_list, value_input_option='USER_ENTERED')
                st.success("Dados enviados com sucesso para o Google Sheets.")
            except Exception as e:
                st.error(f"Falha ao enviar dados: {e}")
        else:
            st.warning("Primeiro conecte-se ao Google Sheets e/ou faça upload de um arquivo.")

if __name__ == '__main__':
    main()
