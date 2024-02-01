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
    creds = ServiceAccountCredentials.from_json_keyfile_name('creden.json', SCOPES)
    client = gspread.authorize(creds)
    return client

def get_last_rows(sheet_url, nrows=3):
    client = authenticate_google_sheets()
    sheet = client.open_by_url(sheet_url)
    worksheet = sheet.sheet1
    # Obtém todas as linhas da planilha
    rows = worksheet.get_all_values()
    # Converte para DataFrame
    df = pd.DataFrame.from_records(rows)
    # Exibe as últimas n linhas
    return df.tail(nrows)

def main():
    st.title("Conexão com Google Sheets e Visualização de Dados")

    sheet_url = "https://docs.google.com/spreadsheets/d/1FPBeAXQBKy8noJ3bTF52p8JL_Eg-ptuSP6djDTsRfKE/edit#gid=0"

    if st.button("Conectar e Exibir Últimas Linhas do Google Sheets"):
        try:
            df_last_rows = get_last_rows(sheet_url)
            st.write("Últimas linhas da planilha Google Sheets:")
            st.dataframe(df_last_rows)
        except Exception as e:
            st.error(f"Falha ao conectar ou ler dados: {e}")

if __name__ == '__main__':
    main()
