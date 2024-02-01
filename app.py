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
    return gspread.authorize(creds)

def insert_data_to_sheet(client, df, sheet_url):
    try:
        sheet = client.open_by_url(sheet_url)
        worksheet = sheet.get_worksheet(0)
        # Converte o DataFrame para uma lista de listas
        data_list = df.astype(str).values.tolist()
        # Encontra a primeira linha vazia
        empty_row = len(worksheet.get_all_values()) + 1
        # Insere os dados a partir da primeira linha vazia
        worksheet.append_rows(data_list, table_range=f"A{empty_row}")
        st.success("Dados inseridos com sucesso no Google Sheets.")
    except Exception as e:
        st.error(f"Falha ao inserir dados: {str(e)}")

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
        insert_data_to_sheet(client, data, sheet_url)

if __name__ == '__main__':
    main()
