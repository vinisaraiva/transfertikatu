import streamlit as st
import pandas as pd
import gspread
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from oauth2client.service_account import ServiceAccountCredentials
import os

def authenticate_google_sheets():
    SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
    # Caminho para o arquivo de credenciais do Google Cloud Platform
    creds = ServiceAccountCredentials.from_json_keyfile_name('creden.json', SCOPES)
    client = gspread.authorize(creds)
    return client

def upload_data_to_sheet(client, data_list, sheet_url):
    try:
        sheet = client.open_by_url(sheet_url)
        worksheet = sheet.get_worksheet(0)  # Primeira aba
        worksheet.append_rows(data_list, value_input_option='USER_ENTERED')
        return True
    except Exception as e:
        st.error(f"Falha ao enviar dados: {e}")
        return False

def main():
    st.title("Upload de Arquivo Excel para Google Sheets")

    uploaded_file = st.file_uploader("Escolha um arquivo Excel", type=['xlsx', 'xls'])

    if uploaded_file is not None:
        # Lê e exibe os dados da planilha Excel
        data = pd.read_excel(uploaded_file, header=0)
        st.write("Dados lidos do arquivo Excel:")
        st.dataframe(data)

        # Botão para conectar ao Google Sheets
        if st.button("Conectar ao Google Sheets"):
            try:
                client = authenticate_google_sheets()
                st.success("Conectado com sucesso ao Google Sheets!")
            except Exception as e:
                st.error(f"Falha ao conectar: {e}")

        # Botão para enviar dados para o Google Sheets
        if st.button("Enviar para Google Sheets"):
            if not data.empty:
                data_list = data.values.tolist()
                if upload_data_to_sheet(client, data_list, "https://docs.google.com/spreadsheets/d/1FPBeAXQBKy8noJ3bTF52p8JL_Eg-ptuSP6djDTsRfKE/edit#gid=0"):
                    st.success("Dados enviados com sucesso para o Google Sheets.")
            else:
                st.warning("Nenhum dado para enviar. Por favor, faça upload de um arquivo Excel primeiro.")

if __name__ == '__main__':
    main()
