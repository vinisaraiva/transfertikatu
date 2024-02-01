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

def insert_data_to_sheet(client, df, sheet_url):
    try:
        sheet = client.open_by_url(sheet_url)
        worksheet = sheet.get_worksheet(0)
        # Encontra a primeira linha vazia
        next_row = len(worksheet.get_all_values()) + 1
        # Prepara os dados para inserção
        data_list = df.applymap(lambda x: str(x) if pd.notnull(x) else '').values.tolist()
        # Insere os dados na planilha
        worksheet.insert_rows(data_list, next_row)
        st.success("Dados inseridos com sucesso no Google Sheets.")
    except Exception as e:
        st.error(f"Falha ao inserir dados: {e}")

def main():
    st.title("Upload e Inserção de Arquivo Excel no Google Sheets")
    uploaded_file = st.file_uploader("Escolha um arquivo Excel", type=['xlsx', 'xls'])

    if uploaded_file is not None:
        data = pd.read_excel(uploaded_file, header=0)
        st.write("Dados lidos do arquivo Excel:")
        st.dataframe(data)

    sheet_url = "https://docs.google.com/spreadsheets/d/1FPBeAXQBKy8noJ3bTF52p8JL_Eg-ptuSP6djDTsRfKE/edit#gid=0"

    if st.button("Conectar ao Google Sheets"):
        client = authenticate_google_sheets()
        st.success("Conectado com sucesso ao Google Sheets.")

    if st.button("Enviar para Google Sheets") and uploaded_file is not None:
        insert_data_to_sheet(client, data, sheet_url)

if __name__ == '__main__':
    main()


