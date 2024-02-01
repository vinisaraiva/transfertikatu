import streamlit as st
import pandas as pd
import numpy as np
import gspread
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from oauth2client.service_account import ServiceAccountCredentials
import os

def authenticate_google_sheets():
    SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
    creds = ServiceAccountCredentials.from_json_keyfile_name('cacesso.json', SCOPES)
    client = gspread.authorize(creds)
    return client

def upload_data_to_sheet(client, data_list, sheet_url):
    try:
        sheet = client.open_by_url(sheet_url)
        worksheet = sheet.sheet1  # Acessando a primeira aba diretamente
        for data_row in data_list:
            worksheet.append_row(data_row, value_input_option='USER_ENTERED')
        return True
    except Exception as e:
        st.error(f"Falha ao enviar dados: {e}")
        return False

def main():
    st.title("Upload de Arquivo Excel para Google Sheets")
    client = None

    uploaded_file = st.file_uploader("Escolha um arquivo Excel", type=['xlsx', 'xls'])
    data = None
    if uploaded_file is not None:
        data = pd.read_excel(uploaded_file, header=0)
        st.write("Dados lidos do arquivo Excel:")
        st.dataframe(data)

    sheet_url = "https://docs.google.com/spreadsheets/d/1FPBeAXQBKy8noJ3bTF52p8JL_Eg-ptuSP6djDTsRfKE/edit#gid=0"

    if st.button("Conectar ao Google Sheets"):
        client = authenticate_google_sheets()
        if client:
            st.success("Conectado com sucesso ao Google Sheets!")
        else:
            st.error("Falha ao conectar ao Google Sheets.")

    if st.button("Enviar para Google Sheets") and data is not None and client is not None:
        data_list = data.values.tolist()
        if upload_data_to_sheet(client, data_list, sheet_url):
            st.success("Dados enviados com sucesso para o Google Sheets.")
        else:
            st.error("Falha ao enviar dados para o Google Sheets.")

if __name__ == '__main__':
    main()
