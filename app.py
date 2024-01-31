import streamlit as st
import pandas as pd
import gspread
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
import os

def authenticate_google_sheets():
    SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
    creds = None

    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'creden.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    client = gspread.authorize(creds)
    return client, creds

def upload_data_to_sheet(client, data_list, sheet_url):
    try:
        sheet = client.open_by_url(sheet_url)
        worksheet = sheet.get_worksheet(0)
        worksheet.append_rows(data_list, value_input_option='USER_ENTERED')
        return True
    except Exception as e:
        st.error(f"Falha ao enviar dados: {e}")
        return False

def main():
    st.title("Upload de Arquivo Excel para Google Sheets")

    uploaded_file = st.file_uploader("Escolha um arquivo Excel", type=['xlsx', 'xls'])
    sheet_url = "https://docs.google.com/spreadsheets/d/1FPBeAXQBKy8noJ3bTF52p8JL_Eg-ptuSP6djDTsRfKE/edit#gid=0"
    data = None

    if uploaded_file is not None:
        # Lê e exibe os dados da planilha Excel
        data = pd.read_excel(uploaded_file, header=1)
        st.write("Dados lidos do arquivo Excel:")
        st.dataframe(data)

    # Botão para conectar ao Google Sheets
    if st.button("Conectar ao Google Sheets"):
        try:
            client, creds = authenticate_google_sheets()
            if creds.valid:
                st.success("Conectado com sucesso ao Google Sheets!")
            else:
                st.error("Falha na conexão com o Google Sheets.")
        except Exception as e:
            st.error(f"Falha ao conectar: {e}")

    # Botão para enviar dados para o Google Sheets
    if st.button("Enviar para Google Sheets"):
        if data is not None:
            data_list = data.values.tolist()
            if upload_data_to_sheet(client, data_list, sheet_url):
                st.success("Dados enviados com sucesso para o Google Sheets.")
        else:
            st.warning("Nenhum dado para enviar. Por favor, faça upload de um arquivo Excel primeiro.")

if __name__ == '__main__':
    main()
